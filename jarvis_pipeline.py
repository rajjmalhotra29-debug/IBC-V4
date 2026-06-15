#!/usr/bin/env python3
"""
Jarvis live pipeline (production)
---------------------------------
Runs daily (e.g. GitHub Actions). No AI / no API key.
1. Scrapes the IBBI public-announcement register (CIRP + liquidation, newest first).
2. Drops voluntary liquidations (solvent wind-downs).
3. Classifies each target's sector from its name; rule-based match vs the client master
   (29A-ineligible excluded, KCM-reachable only).
4. REMEMBERS past listings (seen.json) so only genuinely NEW targets are flagged/alerted.
5. Attaches Form-G / process tracker + financial-availability status; writes opportunities.json.
6. Alerts NEW high-fit matches to a Power Automate HTTP trigger (-> Teams/Outlook), staying in M365.
"""
import os, re, json
from datetime import date, timedelta
import requests
from bs4 import BeautifulSoup
import pandas as pd

IBBI_URL      = "https://www.ibbi.gov.in/en/public-announcement"
CLIENT_MASTER = os.environ.get("CLIENT_MASTER", "KCM_IBC_Client_Master.xlsx")
CLIENT_SHEET  = os.environ.get("CLIENT_SHEET", "Working Base - All Clients")
OUTPUT        = os.environ.get("OUTPUT", "opportunities.json")
SEEN_FILE     = os.environ.get("SEEN_FILE", "seen.json")   # remembers past listings -> only NEW ones alert
ALERT_URL     = os.environ.get("ALERT_URL", "")            # Power Automate HTTP-trigger URL (optional)
HIGH_FIT      = 25
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/124 Safari/537.36"}
TODAY = date.today()

# sector classification from company name (the free, no-AI step) — broadened
SECTOR = {
 "vitrified":["ceramic","vitrified","tile","sanitaryware","porcelain"], "ceramic":["ceramic","tile","sanitaryware","porcelain"],
 "paper":["paper","pulp","board","kraft","packaging"], "ispat":["steel","ispat","iron","tmt","alloy","casting","forging"],
 "steel":["steel","ispat","iron","tmt","alloy","casting","forging","fabrication"], "metal":["steel","iron","alloy","metal","casting"],
 "bio":["pharma","bio","drug","api","formulation","intermediate"], "pharma":["pharma","bio","drug","api","formulation","healthcare"],
 "genics":["pharma","bio","drug","api","formulation"], "life":["pharma","bio","drug","life science","healthcare"],
 "restaurant":["food","restaurant","fmcg","beverage","hospitality"], "food":["food","restaurant","fmcg","beverage"],
 "dairy":["dairy","milk","food","beverage"], "agro":["agro","agri","seed","crop","agrochemical","fertiliser"],
 "agri":["agro","agri","seed","crop","agrochemical"], "textile":["textile","denim","yarn","spinning","cotton","fabric","garment"],
 "chemical":["chemical","intermediate","specialty","dye","pigment"], "organics":["chemical","intermediate","specialty","dye"],
 "cement":["cement","concrete","clinker","rmc"], "plastic":["plastic","polymer","masterbatch","pvc"],
 "polymer":["plastic","polymer","masterbatch","pvc"], "sugar":["sugar","distillery","ethanol"],
 "power":["power","electric","energy","solar","renewable"], "electric":["electrical","transformer","cable","switchgear","motor"],
 "transformer":["transformer","lamination","core","electrical"], "cable":["cable","wire","conductor","electrical"],
 "auto":["auto","component","vehicle","rubber","tyre"], "motor":["motor","auto","component","vehicle"],
 "rubber":["rubber","tyre","tire","reclaim"], "glass":["glass","float glass","container glass"],
 "leather":["leather","footwear","tannery"], "footwear":["footwear","leather","shoe"],
 "infra":["infrastructure","construction","epc","engineering"], "construct":["construction","epc","cement","infrastructure"],
 "realty":["real estate","realty","developer","housing"], "infratech":["infrastructure","construction","engineering"],
 "mining":["mining","mineral","ore","coal"], "mineral":["mineral","mining","ore"],
 "hotel":["hotel","hospitality","resort"], "hospitality":["hotel","hospitality","resort","restaurant"],
 "logistic":["logistics","warehouse","transport","supply chain"], "edible":["edible oil","oil","fats","food"],
 "engineering":["engineering","machinery","capital goods","fabrication"], "pump":["pump","valve","machinery"],
 "packaging":["packaging","carton","corrugated","film"], "ispatglobal":["steel","iron"],
}
def classify(name):
    n=name.lower(); kws=set()
    for tok, exp in SECTOR.items():
        if tok in n: kws.update(exp)
    return sorted(kws)

INEL = re.compile(r"under cirp|liquidation|struck.?off|wilful|disqualif|sebi.?(bar|debar)|undischarged|ineligible|insolvent", re.I)
TEXT = ["Industry","Services","Products / Services","Business (Layer 2)","End-markets","Value-chain position","Synergy directions"]
WT   = {"Products / Services":3,"Industry":3,"Business (Layer 2)":2,"Synergy directions":2,"End-markets":1,"Value-chain position":1,"Services":1}
def reachable(loc):
    l=str(loc).lower()
    return ("reachable" in l) and ("not reachable" not in l) and ("unclear" not in l)
def conf_class(c): return {"HIGH":"g-ok","MEDIUM":"g-warn","LOW":"g-bad","UNVERIFIED":"g-bad"}.get(c,"g-warn")

def match(df, kws, topn=3):
    res=[]
    for _,row in df.iterrows():
        if INEL.search(" ".join(str(row[c]) for c in ["29A flags","Business (Layer 2)","Acquirer-fit"])): continue
        if not reachable(row["Decision locus"]): continue
        sc=0; hit=set()
        for c in TEXT:
            t=str(row[c]).lower()
            for k in kws:
                if k in t: sc+=WT.get(c,1); hit.add(k)
        if sc<=0: continue
        conf=(re.findall(r"HIGH|MEDIUM|LOW|UNVERIFIED",str(row["Confidence"]).upper()) or [""])[0]
        res.append({"client":str(row["Company"]),"score":sc,"matched":sorted(hit),
                    "reach":str(row["Decision locus"]),"na":str(row["29A flags"])[:60] or "Verify",
                    "conf":conf or "—","confClass":conf_class(conf),
                    "naClass":"g-ok" if "clean" in str(row["29A flags"]).lower() else "g-warn",
                    "syn":str(row["Synergy directions"])[:240]})
    res.sort(key=lambda x:-x["score"])
    return res[:topn]

def parse(d):
    try: dd,mm,yy=d.split("-"); return date(int(yy),int(mm),int(dd))
    except: return None

def fetch_ibbi():
    import time
    rows=[]
    for pg in ("","?page=1","?page=2"):
        try:
            r=requests.get(IBBI_URL+pg, headers=HEADERS, timeout=40)
            if r.status_code!=200: print(f"  (skipped '{pg}': HTTP {r.status_code})"); time.sleep(2); continue
        except Exception as ex:
            print(f"  (skipped '{pg}': {ex})"); continue
        t=BeautifulSoup(r.text,"html.parser").find("table")
        if t:
            for tr in t.find_all("tr")[1:]:
                c=[x.get_text(" ",strip=True) for x in tr.find_all("td")]
                if len(c)>=6: rows.append({"type":c[0],"date":c[1],"last":c[2],"company":c[3],"applicant":c[4],"rp":c[5]})
        time.sleep(1.5)
    seen=set(); out=[]
    for e in rows:
        if e["company"] in seen: continue
        seen.add(e["company"]); out.append(e)
    return out

def build():
    df = pd.read_excel(CLIENT_MASTER, sheet_name=CLIENT_SHEET, header=0).fillna("")
    try: known=set(json.load(open(SEEN_FILE)))
    except Exception: known=set()
    raw = [e for e in fetch_ibbi() if "voluntary" not in e["type"].lower()]
    targets=[]
    for e in raw:
        kws=classify(e["company"])
        if not kws: continue
        ms=match(df,kws)
        if not ms: continue
        e["matches"]=ms; targets.append(e)
    maxsc=max((m["score"] for t in targets for m in t["matches"]), default=40)
    data=[]
    for e in targets:
        isLiq="liquidation" in e["type"].lower()
        admit=parse(e["date"]); ld=parse(e["last"])
        for m in e["matches"]: m["barW"]=min(100,round(m["score"]/maxsc*100))
        data.append({"company":e["company"],"isLiq":isLiq,
            "stageLabel":e["type"].replace("Public Announcement of ",""),
            "stageClass":"liq" if isLiq else "cirp","admit":e["date"],"claimsBy":e["last"],
            "claimsDays":(ld-TODAY).days if ld else None,
            "formgBy":(admit+timedelta(days=75)).strftime("%d %b %Y") if (admit and not isLiq) else None,
            "applicant":e["applicant"],"rp":e["rp"],"isNew":(e["company"] not in known),"matches":e["matches"]})
    new_n=sum(1 for d in data if d["isNew"])
    json.dump(sorted(known | {e["company"] for e in raw}), open(SEEN_FILE,"w"))
    payload={"generated":TODAY.strftime("%d %b %Y"),"count":len(data),"new":new_n,"opportunities":data}
    json.dump(payload, open(OUTPUT,"w"), indent=1)
    print(f"Wrote {OUTPUT}: {len(data)} matched targets ({new_n} new) from {len(raw)} distressed listings.")
    alert(data)
    return data

def alert(data):
    if not ALERT_URL: return
    fresh=[d for d in data if d["isNew"] and any(m["score"]>=HIGH_FIT for m in d["matches"])]
    if not fresh: return
    items=[f"{d['company']} -> {d['matches'][0]['client']} (fit {d['matches'][0]['score']})" for d in fresh[:10]]
    try:
        requests.post(ALERT_URL, json={"title":f"{len(fresh)} new IBC opportunities","items":items}, timeout=20); print("Alert posted.")
    except Exception as ex: print("Alert failed:", ex)

if __name__ == "__main__":
    build()
