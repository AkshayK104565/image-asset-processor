"""
core/engine.py  –  Pattern Image Asset Processor engine (no auth).
"""
import os, re, shutil, subprocess, tempfile, threading, time, urllib.parse, zipfile
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

import openpyxl
import requests
from PIL import Image

MAX_WORKERS  = 10
HTTP_TIMEOUT = 30
MAX_RETRIES  = 2

SHEET_IMAGE = "Image Downloader"
SHEET_PXM   = "PXM advanced search"
SHEET_INST  = "Instructions"

IMAGE_EXTS = {".jpg",".jpeg",".png",".bmp",".gif",".tif",".tiff",".webp",".jfif"}
CT_EXT = {
    "application/pdf":".pdf","image/jpeg":".jpg","image/png":".png",
    "image/webp":".webp","image/gif":".gif","image/bmp":".bmp",
    "image/tiff":".tif","image/jfif":".jpg",
}

_cnt=0; _cnt_lock=threading.Lock()
_sess=None; _sess_lock=threading.Lock()

def _uid():
    global _cnt
    with _cnt_lock:
        _cnt+=1
        return f"img_{int(time.time()*1000)}_{_cnt}"

def _get_sess():
    global _sess
    if _sess is None:
        with _sess_lock:
            if _sess is None:
                s=requests.Session(); s.headers.update({"User-Agent":"Mozilla/5.0"}); _sess=s
    return _sess

def _trim(v): return str(v).strip() if v is not None else ""
def _clean(s):
    for c in r'\/:*?"<>|': s=s.replace(c,"")
    return s or "UNKNOWN"
def _ext(p): return os.path.splitext(p)[1].lower()
def _normext(e): e=e.strip().lower(); return ("."+e) if e and not e.startswith(".") else e
def _rm(p):
    try:
        if p and os.path.exists(p): os.remove(p)
    except: pass

def parse_dim(txt):
    txt=txt.strip().lower().replace(" ","")
    if "x" in txt:
        parts=txt.split("x")
        if len(parts)!=2: return 0
        try: w,h=int(parts[0]),int(parts[1]); return max(w,h) if w>0 and h>0 else 0
        except: return 0
    try: v=int(txt); return v if v>0 else 0
    except: return 0

def is_valid_master_id(s):
    s=s.strip(); return len(s)==8 and all(c.isalnum() for c in s)

def find_magick():
    return shutil.which("magick") or shutil.which("convert") or ""

def _run_magick(magick, args):
    try:
        r=subprocess.run([magick]+args, capture_output=True, timeout=120)
        return r.returncode==0
    except: return False

def _dims(p):
    try:
        with Image.open(p) as img: return img.size
    except: return (0,0)

def _to_square_jpg(magick, src, dst, mn, mx):
    if os.path.exists(dst): os.remove(dst)
    w,h=_dims(src)
    if w<=0 or h<=0: return False,""
    if w<mn or h<mn:
        t,action=mn,"UPSIZED"
    elif w>mx or h>mx:
        t,action=mx,"DOWNSIZED"
    else:
        action="ORIGINAL"
        args=[src,"-auto-orient","-background","white","-alpha","remove","-alpha","off",
              "-gravity","center","-extent","%[fx:max(w,h)]x%[fx:max(w,h)]",
              "-colorspace","sRGB","-strip","-quality","92",dst]
        ok=_run_magick(magick,args) and os.path.isfile(dst)
        return ok,action
    args=[src,"-auto-orient","-background","white","-alpha","remove","-alpha","off",
          "-resize",f"{t}x{t}","-gravity","center","-extent",f"{t}x{t}",
          "-colorspace","sRGB","-strip","-quality","92",dst]
    ok=_run_magick(magick,args) and os.path.isfile(dst)
    return ok,action

def _to_jpg(magick, src, dst):
    if os.path.exists(dst): os.remove(dst)
    return _run_magick(magick,[src,"-auto-orient","-background","white",
                               "-alpha","remove","-alpha","off",
                               "-colorspace","sRGB","-strip","-quality","92",dst]) and os.path.isfile(dst)

def _to_pdf(magick, src, dst):
    if os.path.exists(dst): os.remove(dst)
    return _run_magick(magick,[src,dst]) and os.path.isfile(dst)

def _gdrive(url):
    url=url.strip(); fid=""
    if "drive.usercontent.google.com" in url:
        fid=_qp(url,"id")
    else:
        m=re.search(r"/file/d/([^/]+)",url)
        if m: fid=m.group(1)
        elif any(x in url for x in ("drive.google.com/open","drive.google.com/uc","docs.google.com/uc")):
            fid=_qp(url,"id")
    return f"https://drive.usercontent.google.com/u/0/uc?id={fid}&export=download" if fid else url

def _qp(url,p):
    return (urllib.parse.parse_qs(urllib.parse.urlparse(url).query).get(p) or [""])[0]

def _det_ext(orig,ct,src):
    e=_ext(orig)
    if e: return _normext(e)
    ct=ct.lower().split(";")[0].strip()
    for k,v in CT_EXT.items():
        if k in ct: return v
    u=src.split("?")[0].split("#")[0]; e=_ext(u)
    return _normext(e) if e and len(e)<=10 else ""

def _sniff_pdf(p):
    try:
        with open(p,"rb") as f: return f.read(4)==b"%PDF"
    except: return False

def _download(url, tmp):
    final=_gdrive(url); last=None
    for attempt in range(MAX_RETRIES+1):
        try:
            r=_get_sess().get(final,timeout=HTTP_TIMEOUT,allow_redirects=True,stream=True)
            r.raise_for_status()
            ct=r.headers.get("Content-Type","").lower()
            if "text/html" in ct or "application/xhtml" in ct:
                raise ValueError(f"URL returned HTML: {url}")
            orig=""
            cd=r.headers.get("Content-Disposition","")
            if cd:
                m=re.search(r'filename="?([^";]+)"?',cd)
                if m: orig=m.group(1).strip()
            e=_det_ext(orig,ct,final) or ".bin"
            lp=os.path.join(tmp,_uid()+e)
            with open(lp,"wb") as f:
                for chunk in r.iter_content(chunk_size=131072):
                    if chunk: f.write(chunk)
            return lp,ct,orig
        except Exception as exc:
            last=exc
            if attempt<MAX_RETRIES: time.sleep(1.5*(attempt+1))
    raise last

def _process_one(task):
    url=task["url"]; fname=task["final_name"]; atype=task["asset_type"]
    out=task["out_dir"]; magick=task["magick"]
    mn=task["min_size"]; mx=task["max_size"]; tmp=task["tmp_dir"]
    res={"final_name":fname,"ok":False,"is_non_jpg":False,
         "resize_action":"","ext":".jpg" if atype!="sds" else ".pdf",
         "row":task["row"],"error":""}
    tp=""
    try:
        tp,ct,_=_download(url,tmp); se=_ext(tp)
        if atype in ("image","label"):
            if se not in (".jpg",".jpeg"): res["is_non_jpg"]=True
            if se in IMAGE_EXTS or "image/" in ct:
                ok,action=_to_square_jpg(magick,tp,os.path.join(out,fname),mn,mx)
                res["ok"]=ok; res["resize_action"]=action
            else: res["error"]=f"unrecognised image ({se})"
        elif atype=="sds":
            dst=os.path.join(out,fname)
            if se==".pdf" or "application/pdf" in ct or _sniff_pdf(tp):
                res["ok"]=bool(shutil.copy2(tp,dst)) or os.path.isfile(dst)
                res["ext"]=".pdf"
            elif se in IMAGE_EXTS or "image/" in ct:
                tj=os.path.join(tmp,_uid()+".jpg")
                if _to_jpg(magick,tp,tj): res["ok"]=_to_pdf(magick,tj,dst); _rm(tj)
                res["ext"]=".pdf"
            else: res["error"]=f"SDS not PDF/image ({se})"
    except Exception as e: res["error"]=str(e)
    finally: _rm(tp)
    return res

# ── Excel helpers ─────────────────────────────────────────
def _fc(ws,name):
    for cell in ws[1]:
        if _trim(cell.value)==name: return cell.column
    return 0

def _fcp(ws,prefix):
    up=prefix.upper()
    for cell in ws[1]:
        if _trim(cell.value).upper().startswith(up): return cell.column
    return 0

def _lcp(ws,start,prefix):
    up=prefix.upper(); end=start
    for c in range(start+1,ws.max_column+1):
        if _trim(ws.cell(1,c).value).upper().startswith(up): end=c
        else: break
    return end

def _mlr(ws,cols):
    mx=0
    for c in cols:
        if c and c>0:
            for r in range(ws.max_row,1,-1):
                if ws.cell(r,c).value not in (None,""):
                    if r>mx: mx=r
                    break
    return max(mx,1)

def _rha(ws,row,cis,cie,cls_,cle,csds):
    for s,e in [(cis,cie),(cls_,cle)]:
        if s and e and e>=s:
            for c in range(s,e+1):
                if _trim(ws.cell(row,c).value): return True
    return bool(csds and _trim(ws.cell(row,csds).value))

# ── Validation ────────────────────────────────────────────
def validate_workbook(wb, upload_choice):
    errors=[]
    needed=[SHEET_IMAGE,SHEET_INST]
    if upload_choice=="1": needed.append(SHEET_PXM)
    for s in needed:
        if s not in wb.sheetnames: errors.append(f"Sheet '{s}' not found.")
    if errors: return errors
    ws=wb[SHEET_IMAGE]; wi=wb[SHEET_INST]
    cm=_fc(ws,"Master ID"); cmpn=_fc(ws,"MPN")
    cc=_fc(ws,"Country Code"); ca=_fc(ws,"ASIN")
    cis=_fcp(ws,"Image URL"); cie=_lcp(ws,cis,"Image URL") if cis else 0
    cls_=_fcp(ws,"Label Image URL"); cle=_lcp(ws,cls_,"Label Image URL") if cls_ else 0
    csds=_fc(ws,"SDS URL")
    cols=[c for c in [cm,cmpn,cc,ca,cis,cie,cls_,cle,csds] if c]
    lr=_mlr(ws,cols)
    if lr<2: errors.append("No data rows found in 'Image Downloader'."); return errors
    valid_cc=set()
    for r in range(2,wi.max_row+1):
        code=_trim(wi.cell(r,2).value)
        if code: valid_cc.add(code.upper())
    bad=[]
    for r in range(2,lr+1):
        if not _rha(ws,r,cis,cie,cls_,cle,csds): continue
        if upload_choice=="1":
            for col,name in [(cm,"Master ID"),(cmpn,"MPN"),(cc,"Country Code")]:
                if not col or not _trim(ws.cell(r,col).value): bad.append(f"Row {r}: missing {name}")
            mid=_trim(ws.cell(r,cm).value) if cm else ""
            if mid and not is_valid_master_id(mid): bad.append(f"Row {r}: Master ID '{mid}' must be 8 alphanumeric chars")
        else:
            for col,name in [(cc,"Country Code"),(ca,"ASIN")]:
                if not col or not _trim(ws.cell(r,col).value): bad.append(f"Row {r}: missing {name}")
        code=_trim(ws.cell(r,cc).value) if cc else ""
        if code and valid_cc and code.upper() not in valid_cc:
            bad.append(f"Row {r}: Country Code '{code}' not in Instructions sheet")
    errors.extend(bad[:20])
    if len(bad)>20: errors.append(f"…and {len(bad)-20} more row errors.")
    return errors

# ── Main job ──────────────────────────────────────────────
def run_job(excel_path, upload_choice, min_size, max_size, progress_cb=None):
    magick=find_magick()
    if not magick: return {"error":"ImageMagick not found on server."}
    wb=openpyxl.load_workbook(excel_path,data_only=True)
    ws=wb[SHEET_IMAGE]; wi=wb[SHEET_INST]
    cm=_fc(ws,"Master ID"); cmpn=_fc(ws,"MPN")
    cc=_fc(ws,"Country Code"); ca=_fc(ws,"ASIN")
    cis=_fcp(ws,"Image URL"); cie=_lcp(ws,cis,"Image URL") if cis else 0
    cls_=_fcp(ws,"Label Image URL"); cle=_lcp(ws,cls_,"Label Image URL") if cls_ else 0
    csds=_fc(ws,"SDS URL")
    cols=[c for c in [cm,cmpn,cc,ca,cis,cie,cls_,cle,csds] if c]
    lr=_mlr(ws,cols)
    jid=_uid()
    tmp_dir=tempfile.mkdtemp(prefix=f"tmp_{jid}_")
    out_dir=tempfile.mkdtemp(prefix=f"out_{jid}_")
    tasks=[]; seen=set(); dupes=0
    for row in range(2,lr+1):
        if not _rha(ws,row,cis,cie,cls_,cle,csds): continue
        mid=_trim(ws.cell(row,cm).value) if cm else ""
        mpn=_clean(_trim(ws.cell(row,cmpn).value) if cmpn else "")
        country=_trim(ws.cell(row,cc).value) if cc else ""
        asin=_trim(ws.cell(row,ca).value) if ca else ""
        row_out=out_dir
        if upload_choice=="2" and country:
            row_out=os.path.join(out_dir,country); os.makedirs(row_out,exist_ok=True)
        base={"magick":magick,"min_size":min_size,"max_size":max_size,
              "tmp_dir":tmp_dir,"out_dir":row_out,"row":row}
        def _add(url,fname,atype):
            nonlocal dupes
            if not url: return
            if fname.upper() in seen: dupes+=1; return
            seen.add(fname.upper()); tasks.append({**base,"url":url,"final_name":fname,"asset_type":atype})
        if cis and cie:
            for c in range(cis,cie+1):
                url=_trim(ws.cell(row,c).value); idx=c-cis+1
                fname=(f"{mid}_{mpn}_{country}_ISP{idx:02d}.jpg" if upload_choice=="1"
                       else (f"{asin}.MAIN.jpg" if idx==1 else f"{asin}.PT{idx-1:02d}.jpg"))
                _add(url,fname,"image")
        if cls_ and cle:
            for c in range(cls_,cle+1):
                url=_trim(ws.cell(row,c).value); idx=c-cls_+1
                fname=(f"{mid}_{mpn}_{country}_LBL{idx:02d}.jpg" if upload_choice=="1"
                       else f"{asin}.PS{idx:02d}.jpg")
                _add(url,fname,"label")
        if csds:
            url=_trim(ws.cell(row,csds).value)
            fname=(f"{mid}_{mpn}_{country}_SDS.pdf" if upload_choice=="1" else f"{asin}_SDS.pdf")
            _add(url,fname,"sds")
    total=len(tasks); downloaded=0; failed=0
    upsized=0; downsized=0; kept=0; ftypes=defaultdict(int); errs=[]
    pxm_lock=threading.Lock(); pxm_entries=[]; pxm_seen=set(); non_jpg=set()
    def rec_pxm(fname,row,is_nj):
        k=fname.upper()
        with pxm_lock:
            if is_nj: non_jpg.add(row)
            if k not in pxm_seen: pxm_seen.add(k); pxm_entries.append((row,fname))
    done=0
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futs={ex.submit(_process_one,t):t for t in tasks}
        for fut in as_completed(futs):
            r=fut.result(); done+=1
            if r["ok"]:
                downloaded+=1; ftypes[r["ext"]]+=1
                a=r.get("resize_action","")
                if a=="UPSIZED": upsized+=1
                elif a=="DOWNSIZED": downsized+=1
                elif a=="ORIGINAL": kept+=1
                if upload_choice=="1": rec_pxm(r["final_name"],r["row"],r.get("is_non_jpg",False))
            else:
                failed+=1
                if r.get("error"): errs.append(f"Row {r['row']} – {r['final_name']}: {r['error']}")
            if progress_cb: progress_cb(done,total,r["final_name"])
    zip_path=os.path.join(tmp_dir,f"images_{jid}.zip")
    with zipfile.ZipFile(zip_path,"w",zipfile.ZIP_DEFLATED) as zf:
        for root,_,files in os.walk(out_dir):
            for fn in files:
                fp=os.path.join(root,fn); zf.write(fp,os.path.relpath(fp,out_dir))
    shutil.rmtree(out_dir,ignore_errors=True)
    prio=[f for (rn,f) in pxm_entries if rn in non_jpg]
    norm=[f for (rn,f) in pxm_entries if rn not in non_jpg]
    return {"zip_path":zip_path,"downloaded":downloaded,"failed":failed,"duplicates":dupes,
            "total":total,"upsized":upsized,"downsized":downsized,"kept_orig":kept,
            "min_size":min_size,"max_size":max_size,"file_types":dict(ftypes),
            "pxm_list":prio+norm,"errors_log":errs,"error":None}
