"""
Act 10 repeal property-tax estimator: data + calculation + report generation.
Standard library only. Re-implements the method described in WILL's Feb 2025 write-up
("Analysis: Act 10 Repeal - What It Will Cost You"):

  added cost  = max(0, 2009 comp per FTE * inflation - current comp per FTE) * FTEs
  new mills   = added cost / district equalized value (TIF-out)
  home impact = home value * added cost / district equalized value

Inputs:
  * Parcels: Wisconsin Statewide Parcel layer (via wi_assessments.get_json), by SCHOOLDISTNO
  * District equalized value: DPI longitudinal tax-values / rates workbooks (data/*.xlsx)
  * Cost inputs: data/district_costs.csv (filled in by you; see README text in the app)
"""
import csv
import datetime
import html
import http.client
import json
import os
import re
import statistics
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile

import wi_assessments as wa

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, "data")
REPORT_DIR = os.path.join(BASE, "reports")
COSTS_CSV = os.path.join(DATA_DIR, "district_costs.csv")
BASELINE_CSV = os.path.join(DATA_DIR, "teacher_comp_2009_10.csv")   # DPI PI-1202 teacher salary + fringe, 2009-10
STAFF_CSV = os.path.join(DATA_DIR, "derived_staff_current.csv")     # built by refresh_staff_data()
SETTINGS_JSON = os.path.join(DATA_DIR, "settings.json")
TAXVALUES_XLSX = os.path.join(DATA_DIR, "dpi_taxvalues.xlsx")
RATES_XLSX = os.path.join(DATA_DIR, "dpi_rates.xlsx")

DPI_SOURCES = {
    TAXVALUES_XLSX: "https://dpi.wi.gov/media/56145/download?inline",
    RATES_XLSX: "https://dpi.wi.gov/media/28056/download?inline",
}
COST_FIELDS = ["district_code", "district_name", "ftes", "comp_2009_per_fte",
               "comp_now_per_fte", "added_cost_override", "source"]
DEFAULT_SETTINGS = {
    # CPI-U annual avg 2009 (214.5) -> 2024 (313.7). WILL doesn't name its index; edit as needed.
    "inflation_factor": 1.462,
    "home_value": 300000,
    "staff_year": 2025,            # DPI contract year value: 2025 = the 2024-25 school year
    "fte_basis": "all_contracted",  # "teachers" or "all_contracted" (extends the teacher pay gap to all salaried staff)
}
WILL_STATEWIDE_COST = 1_788_000_000
WILL_AVG_HOME_INCREASE = 624
PARCEL_FIELDS = ["SCHOOLDIST", "SCHOOLDISTNO", "CONAME", "PARCELID", "SITEADRESS", "PLACENAME",
                 "PROPCLASS", "ESTFMKVALUE", "CNTASSDVALUE", "LNDVALUE", "IMPVALUE", "TAXROLLYEAR"]


# ---------------------------------------------------------------- xlsx (stdlib reader)
_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _col(ref):
    n = 0
    for ch in re.match(r"[A-Z]+", ref).group(0):
        n = n * 26 + ord(ch) - 64
    return n - 1


def read_sheet(path, sheet_name):
    """Return rows (lists of strings) for the named sheet of an .xlsx file."""
    z = zipfile.ZipFile(path)
    strings = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall(_NS + "si"):
            strings.append("".join(t.text or "" for t in si.iter(_NS + "t")))
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rid = next(s.get(_RNS + "id") for s in wb.iter(_NS + "sheet") if s.get("name") == sheet_name)
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    target = next(r.get("Target") for r in rels if r.get("Id") == rid)
    target = target.lstrip("/")
    sheet = ET.fromstring(z.read(target if target.startswith("xl/") else "xl/" + target))
    rows = []
    for row in sheet.iter(_NS + "row"):
        vals = {}
        for c in row.findall(_NS + "c"):
            v = c.find(_NS + "v")
            if v is None or v.text is None:
                continue
            vals[_col(c.get("r"))] = strings[int(v.text)] if c.get("t") == "s" else v.text
        if vals:
            r = [""] * (max(vals) + 1)
            for k, v in vals.items():
                r[k] = v
            rows.append(r)
    return rows


def _full_year(yy):
    return 1900 + yy if yy >= 80 else 2000 + yy


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------- settings / downloads
def ensure_dirs():
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(REPORT_DIR, exist_ok=True)


def load_settings():
    ensure_dirs()
    s = dict(DEFAULT_SETTINGS)
    try:
        with open(SETTINGS_JSON) as fh:
            s.update(json.load(fh))
    except (OSError, ValueError):
        pass
    return s


def save_settings(updates):
    s = load_settings()
    s.update(updates)
    with open(SETTINGS_JSON, "w") as fh:
        json.dump(s, fh, indent=2)
    return s


def refresh_dpi_data():
    """Download DPI's district value workbooks. Returns list of messages."""
    ensure_dirs()
    msgs = []
    for path, url in DPI_SOURCES.items():
        req = urllib.request.Request(url, headers={"User-Agent": wa.UA})
        with urllib.request.urlopen(req, timeout=120) as resp:
            body = resp.read()
        if not body.startswith(b"PK"):
            raise RuntimeError(f"Unexpected response from {url} (not an Excel file)")
        with open(path, "wb") as fh:
            fh.write(body)
        msgs.append(f"Downloaded {os.path.basename(path)} ({len(body) // 1024} KB)")
    _DPI_CACHE.clear()
    _write_cost_template()
    return msgs


# ---------------------------------------------------------------- DPI district data
_DPI_CACHE = {}


def load_dpi():
    """{'districts': {code: {...}}, 'state': {...}} from the DPI workbooks, cached in memory."""
    if _DPI_CACHE:
        return _DPI_CACHE
    if not os.path.exists(TAXVALUES_XLSX):
        raise FileNotFoundError("DPI data not downloaded yet. Click 'Update DPI data'.")
    rows = read_sheet(TAXVALUES_XLSX, "Data")
    hdr = rows[0]
    yy = max((int(m.group(1)) for h in hdr for m in [re.fullmatch(r"tifout(\d\d)", h)] if m), key=_full_year)
    ci = hdr.index(f"tifin{yy:02d}"), hdr.index(f"tifout{yy:02d}")
    districts, state = {}, {}
    for r in rows[1:]:
        r = r + [""] * (len(hdr) - len(r))
        name = r[1]
        vin, vout = _num(r[ci[0]]), _num(r[ci[1]])
        if name == "State Totals":
            state = {"tif_in": vin, "tif_out": vout}
        elif r[0].strip().isdigit():
            districts[int(r[0])] = {"code": int(r[0]), "name": name, "tif_in": vin, "tif_out": vout}
    year = _full_year(yy)
    if os.path.exists(RATES_XLSX):
        rr = read_sheet(RATES_XLSX, "Data")
        h = rr[0]
        ym = max((int(m.group(1)) for x in h for m in [re.fullmatch(r"fall(\d\d)_ave", x)] if m), key=_full_year)
        ix = {k: h.index(f"fall{ym:02d}_{k}") for k in ("levy", "value", "ave")}
        seen = set()
        for r in rr[1:]:
            r = r + [""] * (len(h) - len(r))
            if r[0].strip().isdigit() and int(r[0]) in districts and int(r[0]) not in seen:
                seen.add(int(r[0]))
                d = districts[int(r[0])]
                d["rate_year"] = _full_year(ym)
                d["levy_prior"] = _num(r[ix["levy"]])
                val = _num(r[ix["value"]])
                # the file's "ave" column is a statewide average, so derive the district's own rate
                d["rate_prior"] = (d["levy_prior"] / val * 1000) if (d["levy_prior"] and val) else None
    _DPI_CACHE.update({"districts": districts, "state": state, "year": year})
    return _DPI_CACHE


def district_list():
    dpi = load_dpi()
    return sorted(dpi["districts"].values(), key=lambda d: d["name"])


# ---------------------------------------------------------------- cost table
def _write_cost_template():
    if os.path.exists(COSTS_CSV):
        return
    with open(COSTS_CSV, "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=COST_FIELDS)
        w.writeheader()
        for d in district_list():
            w.writerow({"district_code": d["code"], "district_name": d["name"]})


def load_costs():
    ensure_dirs()
    out = {}
    if not os.path.exists(COSTS_CSV):
        return out
    with open(COSTS_CSV, newline="") as fh:
        for r in csv.DictReader(fh):
            try:
                out[int(r["district_code"])] = r
            except (ValueError, KeyError, TypeError):
                pass
    return out


def save_cost(code, values):
    _write_cost_template()
    costs = load_costs()
    names = {d["code"]: d["name"] for d in district_list()}
    row = costs.get(code) or {"district_code": code, "district_name": names.get(code, "")}
    changed = False
    for k in ("ftes", "comp_2009_per_fte", "comp_now_per_fte", "added_cost_override"):
        if k in values:
            new = str(values[k]).replace(",", "").replace("$", "").strip()
            old_v = (row.get(k) or "").strip()
            if _num(new) != _num(old_v) or (new == "") != (old_v == ""):
                changed = True
            row[k] = new
    if changed:
        row["source"] = "manual"
    costs[code] = row
    _write_costs(costs)


def _write_costs(costs):
    with open(COSTS_CSV, "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=COST_FIELDS)
        w.writeheader()
        for c in sorted(costs):
            w.writerow({k: costs[c].get(k, "") for k in COST_FIELDS})


def added_cost(cost_row, inflation):
    """Return (added_cost or None, explanation)."""
    if not cost_row:
        return None, "no cost data"
    ov = _num(cost_row.get("added_cost_override"))
    if ov is not None:
        return ov, "entered directly"
    f, c09, cnow = (_num(cost_row.get(k)) for k in ("ftes", "comp_2009_per_fte", "comp_now_per_fte"))
    if None in (f, c09, cnow):
        return None, "no cost data"
    per = max(0.0, c09 * inflation - cnow)
    return per * f, f"(${c09:,.0f} x {inflation} - ${cnow:,.0f}) x {f:,.1f} FTEs"


def statewide_check():
    dpi = load_dpi()
    s = load_settings()
    costs = load_costs()
    total, n = 0.0, 0
    for code in dpi["districts"]:
        a, _ = added_cost(costs.get(code), s["inflation_factor"])
        if a is not None:
            total += a
            n += 1
    ev = dpi["state"].get("tif_out")
    return {"districts_with_data": n, "districts_total": len(dpi["districts"]),
            "total_added_cost": total, "state_equalized_value": ev,
            "implied_300k_increase": (300000 * total / ev) if ev else None,
            "will_total": WILL_STATEWIDE_COST, "will_home_increase": WILL_AVG_HOME_INCREASE}


# ---------------------------------------------------------------- DPI staff data -> cost table
PUB = "https://publicstaffreports.dpi.wi.gov/PubStaffReport/Public/PublicReport/"


def _download(url_base, params, dest, progress=None, attempts=5):
    """Stream a file to disk; verify the byte count against Content-Length and retry on short
    reads or dropped connections (DPI's server sometimes cuts large downloads off)."""
    url = url_base + "?" + urllib.parse.urlencode(params)
    for attempt in range(attempts):
        problem = None
        try:
            req = urllib.request.Request(url, headers={"User-Agent": wa.UA})
            done = 0
            with urllib.request.urlopen(req, timeout=300) as resp, open(dest, "wb") as fh:
                expected = int(resp.headers.get("Content-Length") or 0)
                while True:
                    chunk = resp.read(1 << 20)
                    if not chunk:
                        break
                    fh.write(chunk)
                    done += len(chunk)
                    if progress:
                        progress(f"Downloaded {done // (1 << 20)} MB")
            if not expected or done == expected:
                return
            problem = f"cut short ({done} of {expected} bytes)"
        except (OSError, http.client.HTTPException) as exc:
            problem = f"{type(exc).__name__}: {exc}"
        if attempt < attempts - 1:
            wait = 15 * (attempt + 1)
            if progress:
                progress(f"Download problem ({problem}); retrying in {wait}s...")
            time.sleep(wait)
    raise RuntimeError(f"DPI download kept failing ({problem}). Try again in a few minutes.")


def _code(text):
    m = re.match(r"\s*(\d{1,4})\s*-", text or "")
    return int(m.group(1)) if m else None


def _money(x):
    try:
        return float(re.sub(r"[^\d.]", "", x))
    except ValueError:
        return None


def refresh_staff_data(progress=None):
    """Download DPI's current-year Teacher Salary Report and All Staff file and aggregate them
    per district. The raw All Staff file (it contains names) is deleted after aggregation."""
    ensure_dirs()
    say = progress or (lambda m: None)
    year = load_settings()["staff_year"]
    tpath, apath = os.path.join(DATA_DIR, "_teacher_raw.csv"), os.path.join(DATA_DIR, "_allstaff_raw.csv")
    try:
        say("Downloading teacher salary report...")
        _download(PUB + "TeacherSalaryReportDownload",
                  {"selectedYear": year, "selectedAgencies": "", "criteria": "Filter Criteria: all"}, tpath)
        say("Downloading All Staff file (about 90 MB; a few minutes)...")
        _download(PUB + "AllStaffReportDownload",
                  {"selectedYear": year, "hiringLea": "", "workingLea": "", "position": "",
                   "assignmentArea": "", "licenseType": "", "firstNameContains": "",
                   "lastNameContains": "", "criteria": "Filter Criteria: all"}, apath, say)

        teach = {}
        with open(tpath, newline="", encoding="utf-8", errors="replace") as fh:
            rd = csv.reader(fh)
            next(rd, None)
            for r in rd:
                # data rows hold: year, agency, county, cesa, low, high, avg salary, low, high, avg fringe, ...
                code = _code(r[1]) if len(r) > 9 else None
                sal, fr = (_money(r[6]), _money(r[9])) if code is not None else (None, None)
                if sal and fr is not None and 15000 < sal < 200000:
                    teach[code] = (sal, fr)

        say("Aggregating staff by district...")
        persons = {}     # (research id, contract agency) -> [salary+fringe, fte]
        tfte, comp_fte = {}, {}
        with open(apath, newline="", encoding="utf-8", errors="replace") as fh:
            fh.readline()    # "Filter Criteria" banner line
            for r in csv.DictReader(fh):
                if not r.get("Assignment Hire Agency Type", "").startswith("03") or None in r.values():
                    continue
                code = _code(r["Hire Agency"])
                if code is None:
                    continue
                fte = _num(r["Assignment FTE"]) or 0.0
                if r["Assignment Position"].startswith("53") and r["Assignment Subcontracted"] != "Y":
                    tfte[code] = tfte.get(code, 0.0) + fte
                sal, fr = _money(r["Total Salary"]), _money(r["Total Fringe"])
                if sal:
                    key = (r["Research Id"], code)
                    p = persons.setdefault(key, [sal + (fr or 0.0), 0.0])
                    p[1] += fte
        all_comp, all_fte = {}, {}
        for (_, code), (comp, fte) in persons.items():
            all_comp[code] = all_comp.get(code, 0.0) + comp
            all_fte[code] = all_fte.get(code, 0.0) + fte
        with open(STAFF_CSV, "w", newline="") as fh:
            w = csv.writer(fh)
            w.writerow(["district_code", "teacher_avg_salary", "teacher_avg_fringe", "teacher_fte",
                        "contracted_fte", "contracted_comp_per_fte", "staff_year"])
            for code in sorted(set(teach) | set(tfte)):
                sal, fr = teach.get(code, ("", ""))
                w.writerow([code, sal, fr, round(tfte.get(code, 0.0), 2), round(all_fte.get(code, 0.0), 2),
                            round(all_comp[code] / all_fte[code], 2) if all_fte.get(code) else "", year])
    finally:
        for p in (tpath, apath):
            if os.path.exists(p):
                os.unlink(p)
    return autofill_costs()


def autofill_costs():
    """Fill the cost table from the 2009-10 baseline and current staff data. Never overwrites
    rows you entered by hand (source = manual)."""
    if not (os.path.exists(BASELINE_CSV) and os.path.exists(STAFF_CSV)):
        raise FileNotFoundError("Staff data not downloaded yet.")
    _write_cost_template()
    basis = load_settings()["fte_basis"]
    with open(BASELINE_CSV, newline="") as fh:
        base = {int(r["district_code"]): float(r["total_comp"]) for r in csv.DictReader(fh)}
    with open(STAFF_CSV, newline="") as fh:
        cur = {int(r["district_code"]): r for r in csv.DictReader(fh)}
    costs = load_costs()
    names = {d["code"]: d["name"] for d in district_list()}
    filled = skipped_manual = no_match = 0
    for code, name in names.items():
        row = costs.get(code) or {"district_code": code, "district_name": name}
        if (row.get("source") or "") == "manual":
            skipped_manual += 1
            costs[code] = row
            continue
        c = cur.get(code)
        sal, fr = _num(c["teacher_avg_salary"]) if c else None, _num(c["teacher_avg_fringe"]) if c else None
        fte = _num(c["teacher_fte"] if basis == "teachers" else c["contracted_fte"]) if c else None
        if code not in base or sal is None or fr is None or not fte:
            no_match += 1
            costs[code] = row
            continue
        row.update(ftes=f"{fte:.2f}", comp_2009_per_fte=f"{base[code]:.0f}", comp_now_per_fte=f"{sal + fr:.0f}",
                   added_cost_override="", source="auto")
        costs[code] = row
        filled += 1
    _write_costs(costs)
    return {"filled": filled, "kept_manual": skipped_manual, "no_match": no_match, "basis": basis}


# ---------------------------------------------------------------- parcels
def fetch_district_parcels(code, progress=None):
    where = f"SCHOOLDISTNO='{int(code):04d}'"
    return list(wa.fetch_county(where, "OBJECTID", True, fields=PARCEL_FIELDS, progress=progress))


def parcel_value(p):
    """Market value basis: estimated fair market value, else assessed total."""
    for k in ("ESTFMKVALUE", "CNTASSDVALUE"):
        v = p.get(k)
        if v and v > 0:
            return float(v)
    return 0.0


def is_home(p):
    return (p.get("PROPCLASS") or "").strip() == "1" and (p.get("IMPVALUE") or 0) > 0 and parcel_value(p) > 0


# ---------------------------------------------------------------- calculation
def compute(code, parcels, home_value, added_override=None):
    dpi = load_dpi()
    d = dpi["districts"][int(code)]
    s = load_settings()
    costs = load_costs()
    cost, how = added_cost(costs.get(int(code)), s["inflation_factor"])
    if added_override is not None:
        cost, how = added_override, "entered for this run"
    ev = d["tif_out"]
    share = (cost / ev) if (cost is not None and ev) else None   # fraction of value added to levy
    per_1m = 1_000_000 / ev if ev else None

    homes = [parcel_value(p) for p in parcels if is_home(p)]
    pv = [parcel_value(p) for p in parcels]
    med = statistics.median(homes) if homes else None
    summary = {
        "district": d["name"], "district_code": d["code"], "dpi_value_year": dpi["year"],
        "parcels": len(parcels), "residential_parcels": len(homes),
        "parcel_total_value": sum(pv), "district_equalized_value_tif_out": ev,
        "district_equalized_value_tif_in": d["tif_in"],
        "parcel_value_vs_equalized": (sum(pv) / ev) if ev else None,
        "median_home_value": med, "average_home_value": (sum(homes) / len(homes)) if homes else None,
        "prior_school_rate_per_1000": d.get("rate_prior"), "rate_year": d.get("rate_year"),
        "inflation_factor": s["inflation_factor"], "added_cost": cost, "added_cost_basis": how,
        "added_mills": (share * 1000) if share is not None else None,
        "increase_on_chosen_home": (home_value * share) if share is not None else None,
        "chosen_home_value": home_value,
        "increase_on_median_home": (med * share) if (share is not None and med) else None,
        "per_1m_added_cost_on_median_home": (med * per_1m) if (per_1m and med) else None,
        "per_1m_added_cost_on_chosen_home": (home_value * per_1m) if per_1m else None,
    }

    home_table = []
    for v in sorted({100000, 150000, 200000, 250000, 300000, 400000, 500000, int(home_value)}):
        home_table.append({"home_value": v,
                           "annual_increase": (v * share) if share is not None else None,
                           "per_1m_added_cost": (v * per_1m) if per_1m else None})

    by_mun = {}
    for p in parcels:
        m = by_mun.setdefault(p.get("PLACENAME") or "(unknown)", {"parcels": 0, "homes": [], "total": 0.0})
        m["parcels"] += 1
        m["total"] += parcel_value(p)
        if is_home(p):
            m["homes"].append(parcel_value(p))
    munis = []
    for name, m in sorted(by_mun.items()):
        med_m = statistics.median(m["homes"]) if m["homes"] else None
        munis.append({
            "municipality": name, "parcels": m["parcels"], "residential_parcels": len(m["homes"]),
            "total_value": m["total"], "median_home_value": med_m,
            "median_home_increase": (med_m * share) if (share is not None and med_m) else None,
            "total_increase": (m["total"] * share) if share is not None else None,
        })

    rows = []
    for p in parcels:
        v = parcel_value(p)
        rows.append({
            "county": p.get("CONAME"), "municipality": p.get("PLACENAME"), "parcel_id": p.get("PARCELID"),
            "site_address": p.get("SITEADRESS"), "property_class": p.get("PROPCLASS"),
            "value_used": v or None, "assessed_land": p.get("LNDVALUE"),
            "assessed_improvements": p.get("IMPVALUE"), "tax_roll_year": p.get("TAXROLLYEAR"),
            "estimated_annual_increase": round(v * share, 2) if (share is not None and v) else None})
    return summary, home_table, munis, rows


# ---------------------------------------------------------------- output
def _write_csv(path, rows, fields=None):
    fields = fields or (list(rows[0].keys()) if rows else [])
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)


def _m(x, d=0):
    return "n/a" if x is None else f"${x:,.{d}f}"


def _n(x, d=0):
    return "n/a" if x is None else f"{x:,.{d}f}"


def _table(headers, rows):
    h = "".join(f"<th>{html.escape(x)}</th>" for x in headers)
    b = "".join("<tr>" + "".join(f"<td>{c}</td>" for c in r) + "</tr>" for r in rows)
    return f"<table><thead><tr>{h}</tr></thead><tbody>{b}</tbody></table>"


def render_report(summary, home_table, munis):
    s = summary
    has = s["added_cost"] is not None
    e = html.escape
    if has:
        head = (f"<p class='big'>Estimated increase on a {_m(s['chosen_home_value'])} home: "
                f"<b>{_m(s['increase_on_chosen_home'])}/yr</b>"
                + (f" &nbsp;|&nbsp; on the district's median home ({_m(s['median_home_value'])}): "
                   f"<b>{_m(s['increase_on_median_home'])}/yr</b>" if s["increase_on_median_home"] else "")
                + "</p>")
    else:
        head = ("<p class='warn'><b>No cost data for this district yet.</b> Property-value figures below are real; "
                "the tax increase can't be computed until you enter the district's staffing costs "
                "(or an added-cost figure). The sensitivity table shows the effect of any amount.</p>")
    kv = [
        ("School district", f"{e(s['district'])} (DPI code {s['district_code']})"),
        ("Parcels in district / residential", f"{_n(s['parcels'])} / {_n(s['residential_parcels'])}"),
        (f"Equalized value, TIF-out (DPI {s['dpi_value_year']})", _m(s["district_equalized_value_tif_out"])),
        ("Equalized value, TIF-in", _m(s["district_equalized_value_tif_in"])),
        ("Parcel-layer value (market-value basis)", _m(s["parcel_total_value"])),
        ("Parcel value / equalized value", _n((s["parcel_value_vs_equalized"] or 0) * 100, 1) + "%"
         if s["parcel_value_vs_equalized"] else "n/a"),
        ("Median / average home", f"{_m(s['median_home_value'])} / {_m(s['average_home_value'])}"),
        (f"Prior school tax rate (per $1,000, fall {s['rate_year']})", _n(s["prior_school_rate_per_1000"], 2)),
        ("Added annual cost from repeal", f"{_m(s['added_cost'])} &mdash; {e(s['added_cost_basis'])}"),
        ("Added mills (per $1,000 of value)", _n(s["added_mills"], 3)),
    ]
    hrows = [[_m(r["home_value"]), _m(r["annual_increase"]), _m(r["per_1m_added_cost"], 2)] for r in home_table]
    mrows = [[e(m["municipality"]), _n(m["parcels"]), _n(m["residential_parcels"]), _m(m["total_value"]),
              _m(m["median_home_value"]), _m(m["median_home_increase"]), _m(m["total_increase"])] for m in munis]
    return f"""<!doctype html><html><head><meta charset="utf-8"><title>{e(s['district'])} - Act 10 estimate</title>
<style>body{{font:15px/1.5 -apple-system,Segoe UI,Arial,sans-serif;max-width:980px;margin:24px auto;padding:0 16px;color:#1c2330}}
h1{{font-size:24px;margin:0}}h2{{margin-top:28px;font-size:18px;border-bottom:1px solid #d6dbe3;padding-bottom:4px}}
table{{border-collapse:collapse;width:100%;margin:8px 0}}th,td{{border:1px solid #d6dbe3;padding:5px 8px;text-align:right}}
th:first-child,td:first-child{{text-align:left}}th{{background:#f1f4f8}}.big{{font-size:18px;background:#eef6ff;padding:12px;border-radius:8px}}
.warn{{background:#fff4e0;padding:12px;border-radius:8px}}.small{{color:#5b6573;font-size:13px}}</style></head><body>
<h1>{e(s['district'])}: estimated property-tax impact of an Act 10 repeal</h1>
<p class="small">Generated {datetime.date.today():%B %d, %Y}. Estimate only; see assumptions at the bottom.</p>
{head}
<h2>District summary</h2>{_table(["Item", "Value"], [[a, b] for a, b in kv])}
<h2>Impact by home value</h2>{_table(["Home value", "Estimated increase per year", "Per $1M of added district cost"], hrows)}
<h2>By municipality (within this school district)</h2>
{_table(["Municipality", "Parcels", "Residential", "Total value", "Median home", "Median-home increase/yr", "Total increase/yr"], mrows)}
<h2>Method and assumptions</h2><ul class="small">
<li>Method re-implemented from WILL's Feb 2025 write-up: added cost = (2009 salary+fringe per FTE &times; {s['inflation_factor']} inflation &minus; current salary+fringe per FTE) &times; FTEs, never below zero (districts already paying more than the inflation-adjusted 2009 level add nothing); increase on a property = its value &times; added cost &divide; district equalized value (TIF-out).</li>
<li>Assumes the full added cost lands on the property-tax levy. Wisconsin revenue limits, state aid changes or spending cuts could change that; WILL itself says full pass-through is unlikely.</li>
<li>Property value uses the parcel layer's estimated fair market value (assessed total when missing). "Residential" means property class 1 with improvements.</li>
<li>Sources: Wisconsin Statewide Parcel Map Database (V12); DPI longitudinal tax values and levy/rate workbooks; cost inputs from data/district_costs.csv.</li></ul>
</body></html>"""


def slug(s):
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")


def write_reports(summary, home_table, munis, rows):
    ensure_dirs()
    folder = f"{slug(summary['district'])}_{datetime.datetime.now():%Y%m%d_%H%M%S}"
    out = os.path.join(REPORT_DIR, folder)
    os.makedirs(out)
    _write_csv(os.path.join(out, "district_summary.csv"),
               [{"item": k, "value": v} for k, v in summary.items()], ["item", "value"])
    _write_csv(os.path.join(out, "home_values.csv"), home_table)
    _write_csv(os.path.join(out, "municipalities.csv"), munis)
    _write_csv(os.path.join(out, "parcels.csv"), rows)
    with open(os.path.join(out, "report.html"), "w", encoding="utf-8") as fh:
        fh.write(render_report(summary, home_table, munis))
    return folder


def run_district(code, home_value, added_override=None, progress=None):
    parcels = fetch_district_parcels(code, progress)
    summary, home_table, munis, rows = compute(code, parcels, home_value, added_override)
    folder = write_reports(summary, home_table, munis, rows)
    return folder, summary
