#!/usr/bin/env python3
"""
Pull tax-assessed values from Wisconsin's statewide parcel layer (V12, 2026), with
county selection and sorting. Python standard library only.

Source: Wisconsin Statewide Parcel Map Database, State Cartographer's Office /
Wisconsin Land Information Program (DOA), served as an ArcGIS FeatureServer.

Examples:
  python3 wi_assessments.py --list-counties
  python3 wi_assessments.py --county Sheboygan
  python3 wi_assessments.py --county Sheboygan,Manitowoc,Fond du Lac --sort county
  python3 wi_assessments.py --county Sheboygan --sort total --desc --min-total 300000
  python3 wi_assessments.py --county Sheboygan --address "MAIN ST"
  python3 wi_assessments.py --summary                 # per-county parcel count and avg value
  python3 wi_assessments.py --all -o wisconsin.csv    # entire state (~3.5M rows, slow, big)

Run with no arguments in a terminal and it will ask which counties you want.

Notes:
  * Values are what each county reported to the state; the layer's load date and tax roll
    year are included in the CSV. For the freshest numbers, a county's own service
    (e.g. Sheboygan County's) may be newer than the statewide layer.
  * Owner names are not requested.
  * Check the data's terms before commercial use or redistribution.
"""
import argparse
import csv
import json
import os
import sys
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request

LAYER = ("https://services3.arcgis.com/n6uYoouQZW75n5WI/ArcGIS/rest/services/"
         "Wisconsin_Statewide_Parcels_DB/FeatureServer/0")
PAGE_SIZE = 2000          # layer's MaxRecordCount
DELAY_SECONDS = 0.5       # polite pause between requests
UA = "wi-assessments-script/1.0 (personal research)"

FIELDS = ["CONAME", "PARCELID", "TAXPARCELID", "SITEADRESS", "PLACENAME", "ZIPCODE",
          "TAXROLLYEAR", "PROPCLASS", "GISACRES", "LNDVALUE", "IMPVALUE",
          "CNTASSDVALUE", "ESTFMKVALUE", "LOADDATE"]
HEADERS = {
    "CONAME": "county", "PARCELID": "parcel_id", "TAXPARCELID": "tax_parcel_id",
    "SITEADRESS": "site_address", "PLACENAME": "place_name", "ZIPCODE": "zip",
    "TAXROLLYEAR": "tax_roll_year", "PROPCLASS": "property_class", "GISACRES": "gis_acres",
    "LNDVALUE": "assessed_land", "IMPVALUE": "assessed_improvements",
    "CNTASSDVALUE": "assessed_total", "ESTFMKVALUE": "estimated_fair_market_value",
    "LOADDATE": "load_date",
}
SORT_FIELDS = {
    "county": "CONAME", "address": "SITEADRESS", "parcel": "PARCELID",
    "total": "CNTASSDVALUE", "land": "LNDVALUE", "improvements": "IMPVALUE",
    "fmv": "ESTFMKVALUE", "acres": "GISACRES",
}


def sql_text(s):
    return s.replace("'", "''").upper()


def like_text(s):
    """Escape LIKE wildcards so user input matches literally (paired with ESCAPE '\\')."""
    s = s.upper().replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
    return s.replace("'", "''")


def parse_counties(text):
    return [c.strip() for c in text.split(",") if c.strip()]


def county_base(name):
    base = name.strip().upper()
    if base.endswith(" COUNTY"):
        base = base[: -len(" COUNTY")]
    return base.strip()


def county_clause(name):
    """Match 'Sheboygan' or 'Sheboygan County' case-insensitively; avoids Green/Green Lake mixups."""
    base = county_base(name).replace("'", "''")
    return f"UPPER(CONAME) IN ('{base}','{base} COUNTY')"


def build_where(args, county=None):
    """Filters shared by every query; `county` (one name) adds the county restriction."""
    clauses = []
    if county:
        clauses.append(county_clause(county))
    if args.address:
        clauses.append(f"UPPER(SITEADRESS) LIKE '%{like_text(args.address)}%' ESCAPE '\\'")
    if args.parcel:
        clauses.append(f"UPPER(PARCELID) LIKE '%{like_text(args.parcel)}%' ESCAPE '\\'")
    if args.min_total is not None:
        clauses.append(f"CNTASSDVALUE >= {args.min_total}")
    if args.max_total is not None:
        clauses.append(f"CNTASSDVALUE <= {args.max_total}")
    if args.require_value:
        clauses.append("CNTASSDVALUE > 0")
    return " AND ".join(clauses) if clauses else "1=1"


def build_order(args):
    """Order within one county (rows are always grouped by county by the caller)."""
    order = []
    if args.sort and args.sort != "county":
        order.append(SORT_FIELDS[args.sort] + (" DESC" if args.desc else ""))
    order.append("OBJECTID")  # unique tiebreaker so paging is stable
    return ",".join(order)


class TransientError(Exception):
    pass


def get_json(params, retries=4):
    """POST a query. Retries network errors, HTTP 429/5xx and transient in-body service errors."""
    body = urllib.parse.urlencode(params).encode()
    req = urllib.request.Request(LAYER + "/query", data=body, headers={"User-Agent": UA})
    for attempt in range(retries):
        try:
            try:
                with urllib.request.urlopen(req, timeout=90) as resp:
                    data = json.load(resp)
            except urllib.error.HTTPError as exc:
                if exc.code == 429 or exc.code >= 500:
                    raise TransientError(f"HTTP {exc.code}")
                raise RuntimeError(f"HTTP {exc.code}: {exc.reason}")
            if "error" in data:
                err = data["error"]
                code = err.get("code") if isinstance(err, dict) else None
                msg = str(err.get("message", "")) if isinstance(err, dict) else ""
                # ArcGIS Online sometimes returns a spurious 400 "Cannot perform query" under load.
                if isinstance(code, int) and (code == 429 or code >= 500
                                              or "Cannot perform query" in msg):
                    raise TransientError(f"service error {err}")
                raise RuntimeError(f"Service error: {err}")
            return data
        except (TransientError, urllib.error.URLError, TimeoutError, OSError, json.JSONDecodeError) as exc:
            if attempt == retries - 1:
                raise RuntimeError(f"Request failed after {retries} tries: {exc}")
            time.sleep(2 ** attempt)


def list_counties(echo=True):
    data = get_json({
        "where": "1=1", "outFields": "CONAME", "returnGeometry": "false",
        "returnDistinctValues": "true", "orderByFields": "CONAME", "f": "json",
    })
    names = [f["attributes"].get("CONAME") for f in data.get("features", [])]
    names = [n for n in names if n]
    if echo:
        for n in names:
            print(n)
    return names


def summary(where):
    stats = [
        {"statisticType": "count", "onStatisticField": "OBJECTID", "outStatisticFieldName": "parcels"},
        {"statisticType": "avg", "onStatisticField": "CNTASSDVALUE", "outStatisticFieldName": "avg_total"},
        {"statisticType": "sum", "onStatisticField": "CNTASSDVALUE", "outStatisticFieldName": "sum_total"},
    ]
    data = get_json({
        "where": where, "groupByFieldsForStatistics": "CONAME",
        "outStatistics": json.dumps(stats), "orderByFields": "CONAME",
        "returnGeometry": "false", "f": "json",
    })
    w = csv.writer(sys.stdout)
    w.writerow(["county", "parcels", "avg_assessed_total", "sum_assessed_total"])
    for f in data.get("features", []):
        a = f["attributes"]
        avg = a.get("avg_total")
        w.writerow([a.get("CONAME"), a.get("parcels"),
                    round(avg) if avg is not None else "", a.get("sum_total")])


def fetch_county(where, order, keyset, fields=None, progress=None):
    """Yield attribute dicts for one county's matches.

    keyset=True pages by OBJECTID > last (fast and stable at any depth); otherwise uses
    resultOffset with the requested sort order.
    """
    count = get_json({"where": where, "returnCountOnly": "true", "f": "json"}).get("count", 0)
    print(f"  {count} parcels match.", file=sys.stderr)
    if progress:
        progress(0, count)
    got, last_oid = 0, -1
    while got < count:
        w = f"({where}) AND OBJECTID > {last_oid}" if keyset else where
        params = {"where": w, "outFields": "OBJECTID," + ",".join(fields or FIELDS),
                  "returnGeometry": "false", "orderByFields": "OBJECTID" if keyset else order,
                  "resultRecordCount": PAGE_SIZE, "f": "json"}
        if not keyset:
            params["resultOffset"] = got
        feats = get_json(params).get("features", [])
        if not feats:
            break
        for f in feats:
            yield f["attributes"]
        got += len(feats)
        if keyset:
            last_oid = max(f["attributes"]["OBJECTID"] for f in feats)
        print(f"  fetched {got}/{count}", file=sys.stderr)
        if progress:
            progress(got, count)
        time.sleep(DELAY_SECONDS)
    if got < count:
        print(f"  WARNING: expected {count} rows but received {got}.", file=sys.stderr)


def fetch_all(args, counties):
    """Yield rows county by county (alphabetical, reversed for --sort county --desc)."""
    if counties:
        names = sorted(set(county_base(c) for c in counties), reverse=(args.sort == "county" and args.desc))
    else:  # --all
        names = [None]
    keyset = not (args.sort and args.sort != "county")
    order = build_order(args)
    for name in names:
        print(f"County: {name or 'ALL'}", file=sys.stderr)
        if name is None:
            # whole state: iterate counties so memory/paging stay bounded and output is grouped
            for c in list_counties(echo=False):
                print(f"County: {c}", file=sys.stderr)
                yield from fetch_county(build_where(args, c), order, keyset)
        else:
            yield from fetch_county(build_where(args, name), order, keyset)


def prompt_for_counties():
    print("Which counties? Names separated by commas, or 'list' to see them all.")
    while True:
        try:
            text = input("County: ").strip()
        except EOFError:
            sys.exit("\nNo county given.")
        if text.lower() == "list":
            list_counties()
            continue
        if text:
            return parse_counties(text)


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("-o", "--output", default="wi_assessments.csv")
    p.add_argument("--county", help="county name(s), comma-separated (e.g. Sheboygan,Manitowoc)")
    p.add_argument("--all", action="store_true", help="whole state (millions of rows)")
    p.add_argument("--list-counties", action="store_true", help="print county names and exit")
    p.add_argument("--summary", action="store_true",
                   help="print per-county parcel count and average assessed value, then exit")
    p.add_argument("--sort", choices=sorted(SORT_FIELDS), help="sort field; rows stay grouped by county, sorted within each (default: county)")
    p.add_argument("--desc", action="store_true", help="descending sort")
    p.add_argument("--address", help="site address contains text")
    p.add_argument("--parcel", help="parcel ID contains text")
    p.add_argument("--min-total", type=int, help="minimum total assessed value")
    p.add_argument("--max-total", type=int, help="maximum total assessed value")
    p.add_argument("--require-value", action="store_true", help="skip parcels with no assessed value")
    args = p.parse_args()

    if args.list_counties:
        list_counties()
        return

    counties = parse_counties(args.county) if args.county else []

    if args.summary:
        summary(build_where(args))
        return

    if not counties and not args.all:
        if sys.stdin.isatty():
            counties = prompt_for_counties()
        else:
            sys.exit("Specify --county NAME (or --all for the whole state). "
                     "Use --list-counties to see valid names.")

    if counties:
        known = {c.upper() for c in list_counties(echo=False)}
        for c in counties:
            if county_base(c) not in known and county_base(c) + " COUNTY" not in known:
                sys.exit(f"Unknown county '{c}'. Use --list-counties to see valid names.")

    # Write to a temp file and rename on success so a failed run never clobbers an existing CSV.
    out_dir = os.path.dirname(os.path.abspath(args.output))
    fd, tmp = tempfile.mkstemp(dir=out_dir, prefix=".wi_assessments_", suffix=".tmp")
    n = 0
    try:
        with os.fdopen(fd, "w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow([HEADERS[f] for f in FIELDS])
            for row in fetch_all(args, counties):
                w.writerow([row.get(f) for f in FIELDS])
                n += 1
        os.replace(tmp, args.output)
    except BaseException:
        if os.path.exists(tmp):
            os.unlink(tmp)
        raise
    print(f"Wrote {n} rows to {args.output}", file=sys.stderr)


if __name__ == "__main__":
    main()
