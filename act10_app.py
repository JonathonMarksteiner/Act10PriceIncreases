#!/usr/bin/env python3
"""
Act 10 Property Tax Estimator: a small local web app (standard library only).

Run:   python3 act10_app.py        (opens your browser; Ctrl-C to quit)
Or double-click "Start Act10 Tool.command" on a Mac.
"""
import json
import mimetypes
import os
import socket
import threading
import traceback
import uuid
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, unquote, urlparse

import act10_core as core

JOBS = {}
LOCK = threading.Lock()

PAGE = r"""<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Act 10 Property Tax Estimator</title><style>
:root{--bg:#f6f7f9;--card:#fff;--ink:#1c2330;--mut:#5b6573;--line:#d6dbe3;--acc:#1f5fbf;--warn:#fff4e0}
@media(prefers-color-scheme:dark){:root{--bg:#14171c;--card:#1d222a;--ink:#e8ecf2;--mut:#9aa5b4;--line:#333b47;--acc:#6aa5ff;--warn:#3a2f17}}
body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 -apple-system,Segoe UI,Arial,sans-serif}
main{max-width:1000px;margin:0 auto;padding:20px 16px 60px}h1{font-size:24px;margin:0 0 4px}.mut{color:var(--mut)}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:16px;margin:14px 0}
label{display:block;font-weight:600;margin:10px 0 3px}input,select{font:inherit;padding:8px;border:1px solid var(--line);border-radius:6px;background:var(--bg);color:var(--ink);width:100%;box-sizing:border-box}
.row{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px}
button{font:inherit;font-weight:600;padding:10px 18px;border:0;border-radius:8px;background:var(--acc);color:#fff;cursor:pointer}
button.sec{background:transparent;color:var(--acc);border:1px solid var(--acc)}button:disabled{opacity:.5}
.warn{background:var(--warn);padding:10px 12px;border-radius:8px}a{color:var(--acc)}
iframe{width:100%;height:900px;border:1px solid var(--line);border-radius:8px;background:#fff}
.bar{height:10px;background:var(--line);border-radius:5px;overflow:hidden}.bar>div{height:100%;background:var(--acc);width:0}
details summary{cursor:pointer;font-weight:600}.links a{margin-right:14px}
</style></head><body><main>
<h1>Act 10 Property Tax Estimator</h1>
<div class="mut">Estimate the property-tax effect of an Act 10 repeal for one Wisconsin school district, using WILL's published method.</div>

<div class="card" id="setup" style="display:none"><b>First-time setup:</b> the app needs DPI's district property-value files.
<p><button id="dl">Download DPI data</button> <span id="dlmsg" class="mut"></span></p></div>

<div class="card" id="main" style="display:none">
<label>1. School district</label>
<input id="dist" list="dlist" placeholder="Start typing, e.g. Sheboygan Falls" autocomplete="off"><datalist id="dlist"></datalist>
<div class="row"><div><label>2. Home value to illustrate ($)</label><input id="home" value="300000" inputmode="numeric"></div>
<div><label>Inflation factor, 2009 to now</label><input id="infl" value="1.462"></div>
<div><label>Staff counted</label><select id="basis"><option value="teachers">Teachers only (conservative)</option><option value="all_contracted">All salaried staff (teacher pay gap applied to all)</option></select></div></div>

<details style="margin-top:12px" id="costbox"><summary>3. Cost inputs for this district (staff costs)</summary>
<p class="mut">Enter salary + fringe per FTE for 2009 and today and the number of FTEs, <b>or</b> type the added annual cost directly. Saved for next time.</p>
<div class="row"><div><label>FTEs</label><input id="ftes"></div><div><label>2009 salary+fringe per FTE ($)</label><input id="c09"></div>
<div><label>Current salary+fringe per FTE ($)</label><input id="cnow"></div><div><label>OR added annual cost ($)</label><input id="over"></div></div>
<p class="mut" id="costnote"></p></details>

<p><button id="run">Generate report</button> <button class="sec" id="sw">Statewide check</button></p>
<div id="prog" style="display:none"><div class="bar"><div id="pb"></div></div><div class="mut" id="pt"></div></div>
<div id="err" class="warn" style="display:none"></div></div>

<div class="card" id="res" style="display:none"><div class="links" id="links"></div><p></p><iframe id="frame"></iframe></div>
<div class="card" id="swres" style="display:none"></div>
<div class="card"><details><summary>Maintenance</summary><p><button class="sec" id="upd">Update DPI data</button> <span id="updmsg" class="mut"></span></p>
<p><button class="sec" id="stf">Update staff pay data (about 90 MB, a few minutes)</button> <span id="stfmsg" class="mut"></span></p>
<p class="mut">Cost inputs live in <code>data/district_costs.csv</code> (you can also edit that file in Excel). Reports are saved in the <code>reports</code> folder.</p></details></div>
</main><script>
const $=id=>document.getElementById(id);let districts=[],byName={};
const j=async(u,o)=>{const r=await fetch(u,o);const t=await r.json();if(!r.ok)throw new Error(t.error||r.statusText);return t};
const post=(u,b)=>j(u,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(b)});
const money=x=>x==null?'n/a':'$'+Math.round(x).toLocaleString();
async function init(){const s=await j('/api/status');
 if(!s.ready){$('setup').style.display='';return}
 $('setup').style.display='none';$('main').style.display='';
 districts=await j('/api/districts');$('infl').value=s.settings.inflation_factor;$('basis').value=s.settings.fte_basis;$('home').value=s.settings.home_value;
 $('dlist').innerHTML='';districts.forEach(d=>{const l=d.name+' ('+d.code+')';byName[l]=d;const o=document.createElement('option');o.value=l;$('dlist').appendChild(o)})}
function cur(){return byName[$('dist').value]||districts.find(d=>d.name.toLowerCase()===$('dist').value.trim().toLowerCase())}
$('dist').addEventListener('change',()=>{const d=cur();if(!d)return;const c=d.cost||{};$('ftes').value=c.ftes||'';$('c09').value=c.comp_2009_per_fte||'';$('cnow').value=c.comp_now_per_fte||'';$('over').value=c.added_cost_override||'';
 $('costnote').textContent=d.has_cost?'Cost data found for this district.':'No cost data saved for this district yet.';$('costbox').open=!d.has_cost});
async function dl(btn,msg){btn.disabled=true;msg.textContent='Downloading...';try{const r=await post('/api/refresh',{});msg.textContent=r.messages.join('; ');await init()}catch(e){msg.textContent='Failed: '+e.message}btn.disabled=false}
$('basis').onchange=async()=>{await post('/api/settings',{fte_basis:$('basis').value});await init();if(cur())$('dist').dispatchEvent(new Event('change'))};
$('stf').onclick=async()=>{$('stf').disabled=true;try{const {id}=await post('/api/staff_refresh',{});for(;;){await new Promise(r=>setTimeout(r,1500));const s=await j('/api/job?id='+id);$('stfmsg').textContent=s.msg||'';if(s.error)throw new Error(s.error);if(s.done){const r=s.result;$('stfmsg').textContent='Done: filled '+r.filled+' districts ('+r.no_match+' had no match, '+r.kept_manual+' manual entries kept).';await init();break}}}catch(e){$('stfmsg').textContent='Failed: '+e.message}$('stf').disabled=false};
$('dl').onclick=()=>dl($('dl'),$('dlmsg'));$('upd').onclick=()=>dl($('upd'),$('updmsg'));
$('run').onclick=async()=>{const d=cur();$('err').style.display='none';if(!d){$('err').textContent='Pick a district from the list first.';$('err').style.display='';return}
 $('run').disabled=true;$('res').style.display='none';$('prog').style.display='';$('pb').style.width='0';$('pt').textContent='Starting...';
 try{await post('/api/settings',{inflation_factor:parseFloat($('infl').value),home_value:parseFloat($('home').value.replace(/,/g,''))});
  await post('/api/cost',{code:d.code,ftes:$('ftes').value,comp_2009_per_fte:$('c09').value,comp_now_per_fte:$('cnow').value,added_cost_override:$('over').value});
  const {id}=await post('/api/run',{code:d.code,home_value:parseFloat($('home').value.replace(/,/g,''))});
  for(;;){await new Promise(r=>setTimeout(r,1000));const s=await j('/api/job?id='+id);
   if(s.total){$('pb').style.width=Math.round(100*s.got/s.total)+'%';$('pt').textContent='Fetched '+s.got.toLocaleString()+' of '+s.total.toLocaleString()+' parcels'}
   if(s.error)throw new Error(s.error);
   if(s.done){const f='/reports/'+s.folder+'/';$('links').innerHTML=['report.html','district_summary.csv','municipalities.csv','home_values.csv','parcels.csv'].map(n=>'<a href="'+f+n+(n.endsWith('csv')?'?dl=1':'')+'" '+(n==='report.html'?'target="_blank"':'')+'>'+n+'</a>').join('')+'<span class="mut">Saved in reports/'+s.folder+'</span>';
    $('frame').src=f+'report.html';$('res').style.display='';$('prog').style.display='none';break}}
  await init()}catch(e){$('err').textContent=e.message;$('err').style.display='';$('prog').style.display='none'}$('run').disabled=false};
$('sw').onclick=async()=>{const s=await j('/api/statewide');const c=$('swres');c.style.display='';
 c.innerHTML='<b>Statewide check (vs. WILL)</b><p>Districts with cost data: '+s.districts_with_data+' of '+s.districts_total+'<br>Total added cost: '+money(s.total_added_cost)+' (WILL: '+money(s.will_total)+')<br>Implied increase on a $300,000 home, value-weighted: '+money(s.implied_300k_increase)+' (WILL average: $'+s.will_home_increase+')</p><p class="mut">Only meaningful once most districts have cost data. WILL\'s $624 is an average across districts, so it needn\'t equal the value-weighted number.</p>'};
init().catch(e=>{$('err').textContent=e.message;$('err').style.display=''});
</script></body></html>"""


def status():
    ready = os.path.exists(core.TAXVALUES_XLSX)
    return {"ready": ready, "settings": core.load_settings()}


def run_job(job_id, code, home_value):
    def progress(got, total):
        with LOCK:
            JOBS[job_id].update(got=got, total=total)
    try:
        folder, _ = core.run_district(code, home_value, progress=progress)
        with LOCK:
            JOBS[job_id].update(done=True, folder=folder)
    except Exception as exc:  # surfaced to the UI
        traceback.print_exc()
        with LOCK:
            JOBS[job_id].update(error=str(exc))


def staff_job(job_id):
    def say(msg):
        with LOCK:
            JOBS[job_id]["msg"] = msg
    try:
        result = core.refresh_staff_data(say)
        with LOCK:
            JOBS[job_id].update(done=True, result=result, msg="Done.")
    except Exception as exc:
        traceback.print_exc()
        with LOCK:
            JOBS[job_id].update(error=str(exc))


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, body, ctype="application/json", extra=None):
        data = body if isinstance(body, bytes) else (json.dumps(body) if ctype == "application/json" else body).encode()
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def _body(self):
        n = int(self.headers.get("Content-Length") or 0)
        return json.loads(self.rfile.read(n) or b"{}")

    def do_GET(self):
        u = urlparse(self.path)
        try:
            if u.path == "/":
                return self._send(200, PAGE, "text/html; charset=utf-8")
            if u.path == "/api/status":
                return self._send(200, status())
            if u.path == "/api/districts":
                costs = core.load_costs()
                s = core.load_settings()
                out = []
                for d in core.district_list():
                    row = costs.get(d["code"]) or {}
                    out.append({"code": d["code"], "name": d["name"], "cost": row,
                                "has_cost": core.added_cost(row, s["inflation_factor"])[0] is not None})
                return self._send(200, out)
            if u.path == "/api/job":
                with LOCK:
                    return self._send(200, dict(JOBS.get(parse_qs(u.query)["id"][0], {"error": "unknown job"})))
            if u.path == "/api/statewide":
                return self._send(200, core.statewide_check())
            if u.path.startswith("/reports/"):
                rel = unquote(u.path[len("/reports/"):])
                full = os.path.realpath(os.path.join(core.REPORT_DIR, rel))
                if not full.startswith(os.path.realpath(core.REPORT_DIR) + os.sep) or not os.path.isfile(full):
                    return self._send(404, {"error": "not found"})
                with open(full, "rb") as fh:
                    data = fh.read()
                extra = {"Content-Disposition": f'attachment; filename="{os.path.basename(full)}"'} \
                    if "dl=1" in u.query else None
                return self._send(200, data, mimetypes.guess_type(full)[0] or "application/octet-stream", extra)
            self._send(404, {"error": "not found"})
        except Exception as exc:
            traceback.print_exc()
            self._send(500, {"error": str(exc)})

    def do_POST(self):
        u = urlparse(self.path)
        try:
            b = self._body()
            if u.path == "/api/refresh":
                return self._send(200, {"messages": core.refresh_dpi_data()})
            if u.path == "/api/settings":
                keep = {k: b[k] for k in ("inflation_factor", "home_value") if isinstance(b.get(k), (int, float)) and b[k] > 0}
                if b.get("fte_basis") in ("teachers", "all_contracted"):
                    changed = b["fte_basis"] != core.load_settings()["fte_basis"]
                    keep["fte_basis"] = b["fte_basis"]
                    out = core.save_settings(keep)
                    if changed and os.path.exists(core.STAFF_CSV):
                        core.autofill_costs()
                    return self._send(200, out)
                return self._send(200, core.save_settings(keep))
            if u.path == "/api/staff_refresh":
                job_id = uuid.uuid4().hex
                with LOCK:
                    JOBS[job_id] = {"msg": "Starting..."}
                threading.Thread(target=staff_job, args=(job_id,), daemon=True).start()
                return self._send(200, {"id": job_id})
            if u.path == "/api/cost":
                core.save_cost(int(b["code"]), {k: b.get(k, "") for k in
                               ("ftes", "comp_2009_per_fte", "comp_now_per_fte", "added_cost_override")})
                return self._send(200, {"ok": True})
            if u.path == "/api/run":
                job_id = uuid.uuid4().hex
                with LOCK:
                    JOBS[job_id] = {"got": 0, "total": 0}
                threading.Thread(target=run_job, args=(job_id, int(b["code"]), float(b["home_value"])),
                                 daemon=True).start()
                return self._send(200, {"id": job_id})
            self._send(404, {"error": "not found"})
        except Exception as exc:
            traceback.print_exc()
            self._send(500, {"error": str(exc)})


def free_port(start=8765):
    for p in range(start, start + 20):
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", p)) != 0:
                return p
    raise SystemExit("No free port found.")


def main():
    core.ensure_dirs()
    port = free_port()
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{port}/"
    print(f"Act 10 Property Tax Estimator running at {url}\nPress Ctrl-C to quit.", flush=True)
    threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nBye.")


if __name__ == "__main__":
    main()
