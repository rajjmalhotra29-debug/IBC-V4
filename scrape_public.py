#!/usr/bin/env python3
"""
Public IBBI scraper for the free 'IBC Watch' tool.
Scrapes only PUBLIC data from the IBBI public-announcement register and writes
ibc_listings.json. Contains NO client names and NO matching — safe to host publicly.
Run daily via GitHub Actions in the PUBLIC repo.
"""
import os, json, time
from datetime import date
import requests
from bs4 import BeautifulSoup

IBBI_URL = "https://www.ibbi.gov.in/en/public-announcement"
OUTPUT   = os.environ.get("OUTPUT", "ibc_listings.json")
HEADERS  = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/124 Safari/537.36"}

def fetch():
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
                if len(c)>=6:
                    rows.append({"type":c[0].replace("Public Announcement of ",""),"date":c[1],
                                 "last":c[2],"company":c[3],"applicant":c[4],"rp":c[5]})
        time.sleep(1.5)
    seen=set(); out=[]
    for e in rows:
        if e["company"] in seen: continue
        seen.add(e["company"]); out.append(e)
    return [e for e in out if "voluntary" not in e["type"].lower()]

if __name__=="__main__":
    data=fetch()
    json.dump({"generated":date.today().strftime("%d %b %Y"),"count":len(data),"listings":data},
              open(OUTPUT,"w"), indent=1)
    print(f"Wrote {OUTPUT}: {len(data)} public IBC listings.")
