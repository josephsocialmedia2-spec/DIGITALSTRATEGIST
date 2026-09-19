#!/usr/bin/env python3
import json, re, time
from datetime import datetime, timezone
from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup

BASE = "https://lasacraimmobiliare.it"
COLLECTION = BASE + "/collections/in-vendita"
OUT = "campagna-email-test/data/annunci.json"
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; F1CampaignSync/1.0; +https://f1immobiliare.com/)"}

def clean(s):
    return re.sub(r"\s+", " ", s or "").strip()

def abs_url(u):
    if not u:
        return ""
    if u.startswith("//"):
        return "https:" + u
    return urljoin(BASE, u)

def extract_product(url):
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    text = soup.get_text("\n", strip=True)

    title = clean((soup.find("h1") or soup.title).get_text(" ", strip=True))
    canonical = soup.find("link", rel="canonical")
    canonical_url = canonical.get("href") if canonical else url

    price = ""
    price_meta = soup.find("meta", attrs={"property":"product:price:amount"})
    if price_meta and price_meta.get("content"):
        try:
            price = "€" + format(float(price_meta["content"]), ",.2f").replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            price = "€" + price_meta["content"]
    if not price:
        m = re.search(r"€\s*([0-9\.]+,[0-9]{2})", text)
        if m: price = "€" + m.group(1)

    description = ""
    for sel in [".product__description", ".product-description", ".rte"]:
        el = soup.select_one(sel)
        if el:
            candidate = clean(el.get_text(" ", strip=True))
            if len(candidate) > 80:
                description = candidate
                break

    def field(pattern):
        m = re.search(pattern, text, re.I)
        return clean(m.group(1)) if m else ""

    mq = field(r"\bMq\s*([0-9\.,]+)")
    locali = field(r"\bLocali\s*([0-9]+)")
    bagni = field(r"\bBagno(?:i)?\s*([0-9]+)")
    classe = field(r"Classe energetica\s*([A-G])")
    anno = field(r"Anno di costruzione:?\s*([0-9\.]+)").replace(".","")
    impianto = field(r"Tipo impianto:?\s*([^\n]+)")
    alimentazione = field(r"Tipo alimentazione:?\s*([^\n]+)")
    riscaldamento = ""
    if re.search(r"riscaldamento\s+autonom", text, re.I):
        riscaldamento = "Autonomo"
    elif re.search(r"riscaldamento\s+centralizz", text, re.I):
        riscaldamento = "Centralizzato"

    video = ""
    for a in soup.find_all("a", href=True):
        href = a.get("href","")
        if "youtu.be/" in href or "youtube.com/" in href:
            video = href
            break

    images = []
    for img in soup.find_all("img"):
        src = img.get("src") or img.get("data-src") or ""
        src = abs_url(src)
        if src and "cdn/shop" in src and src not in images:
            images.append(src)
    images = images[:20]

    handle = canonical_url.rstrip("/").split("/")[-1].split("?")[0]
    return {
        "id": handle,
        "title": title,
        "price": price,
        "url": canonical_url,
        "description": description,
        "mq": mq,
        "locali": locali,
        "bagni": bagni,
        "classe": classe,
        "anno": anno,
        "impianto": impianto,
        "alimentazione": alimentazione,
        "riscaldamento": riscaldamento,
        "video": video,
        "images": images
    }

def main():
    r = requests.get(COLLECTION, headers=HEADERS, timeout=30)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    urls = []
    for a in soup.find_all("a", href=True):
        href = a["href"].split("#")[0]
        if href.startswith("/products/"):
            u = abs_url(href)
            if u not in urls:
                urls.append(u)

    items = []
    for i, url in enumerate(urls, 1):
        try:
            items.append(extract_product(url))
            print(f"[{i}/{len(urls)}] {items[-1]['title']}")
        except Exception as exc:
            print(f"[WARN] {url}: {exc}")
        time.sleep(0.15)

    payload = {
        "source": COLLECTION,
        "synced_at": datetime.now(timezone.utc).isoformat(),
        "count": len(items),
        "items": items
    }
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    main()
