#!/usr/bin/env python3
import json, re, time
from datetime import datetime, timezone
from urllib.parse import urljoin, urlsplit, urlunsplit
import requests
from bs4 import BeautifulSoup

BASE = "https://lasacraimmobiliare.it"
COLLECTION = BASE + "/collections/in-vendita"
OUT = "campagna-email-test/data/annunci.json"
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; CampaignSync/1.1)"}

def clean(s):
    return re.sub(r"\s+", " ", s or "").strip()

def abs_url(u):
    if not u:
        return ""
    if u.startswith("//"):
        return "https:" + u
    return urljoin(BASE, u)

def strip_query(u):
    p = urlsplit(u)
    return urlunsplit((p.scheme,p.netloc,p.path,"",""))

def empty_stub(url, title="", price=""):
    handle = strip_query(url).rstrip("/").split("/")[-1]
    return {
        "id": handle, "title": title or handle.replace("-"," ").upper(),
        "price": price, "url": strip_query(url), "description": "",
        "mq": "", "locali": "", "bagni": "", "classe": "", "anno": "",
        "impianto": "", "alimentazione": "", "riscaldamento": "",
        "video": "", "images": [], "partial": True
    }

def extract_product(url, fallback_title="", fallback_price=""):
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    text = soup.get_text("\n", strip=True)

    h1 = soup.find("h1")
    title = clean(h1.get_text(" ", strip=True)) if h1 else fallback_title
    canonical = soup.find("link", rel="canonical")
    canonical_url = strip_query(canonical.get("href") if canonical else url)

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
    if not price:
        price = fallback_price

    # Description: prefer structured Product JSON-LD. Never use the first generic .rte,
    # because Shopify themes can wrap gallery/accessibility text inside it.
    description = ""
    def find_product_description(obj):
        if isinstance(obj, dict):
            typ = obj.get("@type")
            types = typ if isinstance(typ, list) else [typ]
            if "Product" in types and obj.get("description"):
                return obj.get("description")
            for value in obj.values():
                found = find_product_description(value)
                if found:
                    return found
        elif isinstance(obj, list):
            for value in obj:
                found = find_product_description(value)
                if found:
                    return found
        return ""

    for script in soup.find_all("script", attrs={"type":"application/ld+json"}):
        raw = script.string or script.get_text("", strip=True)
        if not raw:
            continue
        try:
            obj = json.loads(raw)
            raw_desc = find_product_description(obj)
            if raw_desc:
                description = clean(BeautifulSoup(str(raw_desc), "html.parser").get_text(" ", strip=True))
                if len(description) >= 40:
                    break
        except Exception:
            pass

    # Fallback tailored to La Sacra product pages:
    # take only the real editorial block after ARREDATO and stop before
    # points/CTA/energy/mortgage sections.
    if not description or "Passa alle informazioni sul prodotto" in description:
        lines = [clean(x) for x in text.splitlines() if clean(x)]
        start = -1
        for idx, line in enumerate(lines):
            if re.match(r"^ARREDATO(?:\\s|$)", line, re.I):
                start = idx
        if start < 0:
            for label in ("TERRAZZO", "ASCENSORE", "RISCALDAMENTO", "GARAGE", "GIARDINO", "BAGNO", "LOCALI"):
                for idx, line in enumerate(lines):
                    if re.match(r"^" + label + r"(?:\\s|$)", line, re.I):
                        start = max(start, idx)

        stop_markers = (
            "Punti di forza", "CTA WhatsApp", "Video disponibile",
            "Classe energetica", "Aggiornamento mutui", "Calcola il mutuo"
        )
        end = len(lines)
        if start >= 0:
            for idx in range(start + 1, len(lines)):
                low = lines[idx].lower()
                if any(marker.lower() in low for marker in stop_markers):
                    end = idx
                    break
            block = lines[start + 1:end]
            # Remove non-editorial artifacts that can sit between specs and copy.
            block = [
                x for x in block
                if x not in ("Spedizione", "Produzione")
                and not re.fullmatch(r"\\d{1,2}\\s+[a-zà-ÿ]+\\s+20\\d{2}", x, re.I)
                and "partner di produzione" not in x.lower()
            ]
            candidate = clean(" ".join(block))
            if len(candidate) >= 40:
                description = candidate

    # Final safety cleaning if a theme artifact slipped through.
    description = re.sub(r"^.*?\\bARREDATO\\b\\s*(?:SI|NO)?\\s*", "", description, flags=re.I)
    description = re.split(
        r"\\b(?:Punti di forza|CTA WhatsApp|Video disponibile|Classe energetica|Aggiornamento mutui|Calcola il mutuo)\\b",
        description, maxsplit=1, flags=re.I
    )[0]
    description = clean(description)[:2600]

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
    rm = re.search(r"\bRiscaldamento\s*(Autonomo|Centralizzato)", text, re.I)
    if rm:
        riscaldamento = rm.group(1).capitalize()
    elif re.search(r"riscaldamento\s+autonom", text, re.I):
        riscaldamento = "Autonomo"
    elif re.search(r"riscaldamento\s+centralizz", text, re.I):
        riscaldamento = "Centralizzato"

    video = ""
    for tag in soup.find_all(True):
        candidates = [tag.get("href",""), tag.get("src",""), tag.get("data-src",""), tag.get("data-url","")]
        for href in candidates:
            if href and ("youtu.be/" in href or "youtube.com/" in href):
                video = abs_url(href)
                break
        if video:
            break

    images = []
    seen_base = set()
    for img in soup.find_all("img"):
        src = abs_url(img.get("src") or img.get("data-src") or "")
        if not src or "cdn/shop" not in src:
            continue
        base_img = src.split("?")[0]
        fname = base_img.rsplit("/",1)[-1].lower()
        if fname.startswith("icons8-") or "untitled_design_14" in fname or "logo" in fname:
            continue
        if base_img in seen_base:
            continue
        seen_base.add(base_img)
        images.append(src)
    images = images[:20]

    handle = canonical_url.rstrip("/").split("/")[-1]
    return {
        "id": handle, "title": title or fallback_title, "price": price,
        "url": canonical_url, "description": description, "mq": mq,
        "locali": locali, "bagni": bagni, "classe": classe, "anno": anno,
        "impianto": impianto, "alimentazione": alimentazione,
        "riscaldamento": riscaldamento, "video": video, "images": images,
        "partial": False
    }

def main():
    r = requests.get(COLLECTION, headers=HEADERS, timeout=30)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    products = {}
    for a in soup.find_all("a", href=True):
        href = a["href"].split("#")[0]
        if not href.startswith("/products/"):
            continue
        url = strip_query(abs_url(href))
        if url in products:
            continue
        title = clean(a.get_text(" ", strip=True))
        parent_text = clean(a.parent.get_text(" ", strip=True)) if a.parent else ""
        pm = re.search(r"€\s*([0-9\.]+,[0-9]{2})", parent_text)
        price = "€"+pm.group(1) if pm else ""
        products[url] = {"title": title, "price": price}

    items = []
    failures = []
    urls = list(products.keys())
    for i, url in enumerate(urls, 1):
        meta = products[url]
        try:
            item = extract_product(url, meta["title"], meta["price"])
            items.append(item)
            print(f"[{i}/{len(urls)}] {item['title']}")
        except Exception as exc:
            failures.append({"url":url,"error":str(exc)})
            items.append(empty_stub(url, meta["title"], meta["price"]))
            print(f"[WARN] {url}: {exc}")
        time.sleep(0.2)

    payload = {
        "source": COLLECTION,
        "synced_at": datetime.now(timezone.utc).isoformat(),
        "count": len(items),
        "partial_count": sum(1 for x in items if x.get("partial")),
        "failures": failures,
        "items": items
    }
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    main()
