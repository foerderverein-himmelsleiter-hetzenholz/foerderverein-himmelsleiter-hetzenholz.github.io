#!/usr/bin/env python3
"""Neuigkeiten und letzte Projekte aus SharePoint in die Webseite übernehmen.

Quelle ist die Liste "Projektübersicht" auf der SharePoint-Website "Projekte":
  - Neuigkeiten:          "Webseite: Neuigkeit" ist Ankündigung, Helfer gesucht oder Zwischeninfo,
                          das Projekt ist nicht abgeschlossen und "anzeigen bis" ist nicht vorbei.
  - Unsere letzten Projekte: Projektstatus "Abgeschlossen" und Veröffentlichung "Öffentlich"
                          (die neuesten 5).

Das Skript schreibt nur die markierten Bereiche in index.html neu und legt die Plakate
in img/auto/ ab. Veröffentlicht wird erst, wenn jemand den Pull Request auf GitHub übernimmt.

Aufruf in GitHub (Zugangsdaten als Secrets):
    python scripts/neuigkeiten.py
Lokaler Test ohne SharePoint:
    python scripts/neuigkeiten.py --beispiel scripts/beispiel.json
"""
import argparse
import datetime as dt
import hashlib
import html
import io
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from zoneinfo import ZoneInfo

SITE_HOST = "himmelsleiterhetzenholz.sharepoint.com"
SITE_PATH = "/sites/Projekte"
LIST_ID = "a55e3ef5-87ea-4a9b-b88f-3f481f32ee09"  # Projektübersicht
MAX_NEUIGKEITEN = 6
MAX_PROJEKTE = 5
ARTEN = ("Ankündigung", "Helfer gesucht", "Zwischeninfo")
FELDER = ("Title,Projektstatus,Veroeffentlichung,Webseite,WebseiteBis,Webtext,Webtermine,"
          "Webplakat,Enddatum,Projektergebnis,Kurzbeschreibung")

WURZEL = Path(__file__).resolve().parent.parent
SEITE = WURZEL / "index.html"
BILDER = WURZEL / "img" / "auto"
BERLIN = ZoneInfo("Europe/Berlin")


# ---------- SharePoint (Microsoft Graph) ----------

def _anfrage(url, token=None, daten=None, roh=False):
    kopf = {"Accept": "application/json"}
    if token:
        kopf["Authorization"] = "Bearer " + token
    if daten is not None:
        daten = urllib.parse.urlencode(daten).encode()
    try:
        with urllib.request.urlopen(urllib.request.Request(url, data=daten, headers=kopf), timeout=60) as r:
            inhalt = r.read()
    except urllib.error.HTTPError as fehler:
        try:
            info = json.loads(fehler.read())
            grund = info.get("error_description") or info.get("error", {}).get("message") or info
        except Exception:
            grund = ""
        ziel = url.split("?")[0]
        sys.exit(f"Fehler {fehler.code} bei {ziel}: {str(grund).splitlines()[0] if grund else ''}")
    return inhalt if roh else json.loads(inhalt)


class SharePoint:
    def __init__(self):
        fehlt = [n for n in ("TENANT_ID", "CLIENT_ID", "CLIENT_SECRET") if not os.environ.get(n)]
        if fehlt:
            sys.exit("Fehlende Zugangsdaten: " + ", ".join(fehlt))
        self.token = _anfrage(
            f"https://login.microsoftonline.com/{os.environ['TENANT_ID']}/oauth2/v2.0/token",
            daten={"grant_type": "client_credentials", "client_id": os.environ["CLIENT_ID"],
                   "client_secret": os.environ["CLIENT_SECRET"],
                   "scope": "https://graph.microsoft.com/.default"})["access_token"]
        self.site = self.graph(f"/sites/{SITE_HOST}:{SITE_PATH}")["id"]
        self._laufwerke = None

    def graph(self, pfad, roh=False):
        url = pfad if pfad.startswith("https://") else "https://graph.microsoft.com/v1.0" + pfad
        return _anfrage(url, self.token, roh=roh)

    def eintraege(self):
        url = (f"/sites/{self.site}/lists/{LIST_ID}/items"
               f"?$top=200&$expand=fields($select={FELDER})")
        while url:
            antwort = self.graph(url)
            for e in antwort["value"]:
                felder = e["fields"]
                felder["ID"] = int(e["id"])
                felder.setdefault("Modified", e.get("lastModifiedDateTime"))
                yield felder
            url = antwort.get("@odata.nextLink")

    def bild(self, wert):
        """Lädt das Bild einer Bildspalte. wert ist JSON-Text oder Objekt."""
        info = json.loads(wert) if isinstance(wert, str) else wert
        adresse = info.get("serverUrl", "https://" + SITE_HOST) + info["serverRelativeUrl"]
        if self._laufwerke is None:
            self._laufwerke = self.graph(f"/sites/{self.site}/drives?$select=id,webUrl")["value"]
        for lw in sorted(self._laufwerke, key=lambda l: -len(l["webUrl"])):
            basis = urllib.parse.unquote(lw["webUrl"]).rstrip("/") + "/"
            if adresse.startswith(basis):
                rest = urllib.parse.quote(adresse[len(basis):])
                return self.graph(f"/drives/{lw['id']}/root:/{rest}:/content", roh=True)
        kennung = "u!" + __import__("base64").urlsafe_b64encode(adresse.encode()).decode().rstrip("=")
        return self.graph(f"/shares/{kennung}/driveItem/content", roh=True)


class Beispiel:
    """Lokaler Test: Einträge aus einer JSON-Datei, Bilder aus dem Repository."""
    def __init__(self, datei):
        self.daten = json.loads(Path(datei).read_text(encoding="utf-8"))

    def eintraege(self):
        return iter(self.daten)

    def bild(self, wert):
        return (WURZEL / wert).read_bytes()


# ---------- Auswahl ----------

def datum(wert):
    if not wert:
        return None
    return dt.datetime.fromisoformat(wert.replace("Z", "+00:00")).astimezone(BERLIN).date()


def auswahl(eintraege, heute):
    neu, fertig = [], []
    for e in eintraege:
        status = e.get("Projektstatus") or ""
        if status == "Abgeschlossen" and e.get("Veroeffentlichung") == "Öffentlich":
            fertig.append(e)
        elif status not in ("Abgeschlossen", "Archiviert") and e.get("Webseite") in ARTEN:
            bis = datum(e.get("WebseiteBis"))
            if bis is None or bis >= heute:
                neu.append(e)
    neu.sort(key=lambda e: (datum(e.get("WebseiteBis")) or dt.date.max, e["ID"]))
    fertig.sort(key=lambda e: (e.get("Enddatum") or e.get("Modified") or "", e["ID"]), reverse=True)
    return neu[:MAX_NEUIGKEITEN], fertig[:MAX_PROJEKTE]


# ---------- Bilder ----------

def plakat_speichern(quelle, e, inhalt):
    """Speichert Plakat groß (max. 1600 px) und klein (640 px breit). Name enthält eine Prüfsumme,
    damit unveränderte Plakate nicht bei jedem Lauf neu geschrieben werden."""
    from PIL import Image, ImageOps

    pruef = hashlib.sha256(inhalt).hexdigest()[:10]
    name = f"{e['ID']}-{pruef}"
    gross, klein = BILDER / f"{name}.jpg", BILDER / f"{name}-klein.jpg"
    if not (gross.exists() and klein.exists()):
        BILDER.mkdir(parents=True, exist_ok=True)
        bild = ImageOps.exif_transpose(Image.open(io.BytesIO(inhalt)))
        if bild.mode != "RGB":
            grund = Image.new("RGB", bild.size, "white")
            grund.paste(bild, mask=bild.convert("RGBA").split()[-1])
            bild = grund
        g = bild.copy()
        g.thumbnail((1600, 1600))
        g.save(gross, "JPEG", quality=85, optimize=True, progressive=True)
        k = bild.copy()
        k.thumbnail((640, 10000))
        k.save(klein, "JPEG", quality=82, optimize=True, progressive=True)
    with Image.open(klein) as k:
        breite, hoehe = k.size
    return {"gross": f"img/auto/{gross.name}", "klein": f"img/auto/{klein.name}",
            "breite": breite, "hoehe": hoehe}


def plakat(quelle, e):
    if not e.get("Webplakat"):
        return None
    try:
        return plakat_speichern(quelle, e, quelle.bild(e["Webplakat"]))
    except Exception as fehler:  # ein kaputtes Bild soll nicht die ganze Seite blockieren
        print(f"Plakat für {e.get('Title')} (ID {e['ID']}) nicht geladen: {fehler}", file=sys.stderr)
        return None


# ---------- HTML ----------

def esc(t):
    return html.escape(t or "", quote=True)


def ohne_html(t):
    t = re.sub(r"<(br|/p|/li|/div)[^>]*>", "\n", t or "", flags=re.I)
    return html.unescape(re.sub(r"<[^>]+>", "", t)).strip()


def absaetze(text, einrueck):
    teile = [t.strip() for t in re.split(r"\n\s*\n", (text or "").replace("\r", "")) if t.strip()]
    return "".join(f"\n{einrueck}<p>{esc(t).replace(chr(10), '<br>')}</p>" for t in teile)


def termine(text, einrueck):
    zeilen = [z.strip() for z in (text or "").splitlines() if z.strip()]
    if not zeilen:
        return ""
    teile = []
    for z in zeilen[:4]:
        fett, _, rest = z.partition("|")
        teile.append(f'<div class="termin"><b>{esc(fett.strip())}</b>{esc(rest.strip())}</div>')
    innen = "".join(f"\n{einrueck}  {t}" for t in teile)
    return f'\n{einrueck}<div class="termine">{innen}\n{einrueck}</div>'


def plakat_html(p, titel, einrueck, lazy):
    if not p:
        return ""
    laden = ' loading="lazy"' if lazy else ""
    return (f'\n{einrueck}<a class="plakat" href="{p["gross"]}" target="_blank" rel="noopener">'
            f'<img src="{p["klein"]}" alt="Plakat: {esc(titel)}" width="{p["breite"]}" '
            f'height="{p["hoehe"]}"{laden}></a>')


def neuigkeiten_html(quelle, liste):
    e6 = " " * 6
    if not liste:
        return f'\n{e6}<p class="leise">Gerade stehen keine Termine an. Schau bald wieder vorbei.</p>\n'
    teile = []
    for e in liste:
        p = plakat(quelle, e)
        klasse = "aktuell" if p else "aktuell ohne-plakat"
        teile.append(
            f'\n{e6}<article class="{klasse}">'
            f'{plakat_html(p, e.get("Title"), e6 + "  ", False)}'
            f'\n{e6}  <div class="aktuell-text">'
            f'\n{e6}    <p class="art">{esc(e["Webseite"])}</p>'
            f'\n{e6}    <h3>{esc(e.get("Title"))}</h3>'
            f'{termine(e.get("Webtermine"), e6 + "    ")}'
            f'{absaetze(e.get("Webtext"), e6 + "    ")}'
            f'\n{e6}  </div>'
            f'\n{e6}</article>')
    return "".join(teile) + "\n"


def kurz(text, laenge=280):
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) <= laenge:
        return text
    return text[:laenge].rsplit(" ", 1)[0].rstrip(",;:-") + " …"


def projekte_html(quelle, liste):
    if not liste:
        return "\n"
    e4 = " " * 4
    karten = []
    for e in liste:
        p = plakat(quelle, e)
        text = e.get("Webtext") or ohne_html(e.get("Projektergebnis")) or ohne_html(e.get("Kurzbeschreibung"))
        jahr = (datum(e.get("Enddatum")) or datum(e.get("Modified")) or dt.date.today()).year
        karten.append(
            f'\n{e4}  <li class="rueck-karte">'
            f'{plakat_html(p, e.get("Title"), e4 + "    ", True)}'
            f'\n{e4}    <p class="art">Abgeschlossen {jahr}</p>'
            f'\n{e4}    <h4>{esc(e.get("Title"))}</h4>'
            f'\n{e4}    <p>{esc(kurz(text))}</p>'
            f'\n{e4}  </li>')
    return (f'\n{e4}<h3 class="unterkopf">Unsere letzten Projekte</h3>'
            f'\n{e4}<ul class="rueckblick">{"".join(karten)}\n{e4}</ul>\n')


def einsetzen(seite, name, inhalt):
    muster = re.compile(rf"(<!-- AUTO:{name} START[^>]*-->).*?(\s*<!-- AUTO:{name} ENDE -->)", re.S)
    if not muster.search(seite):
        sys.exit(f"Markierung AUTO:{name} fehlt in index.html")
    return muster.sub(lambda m: m.group(1) + inhalt.rstrip("\n") + m.group(2), seite, count=1)


def main():
    teil = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    teil.add_argument("--beispiel", help="JSON-Datei statt SharePoint (lokaler Test)")
    teil.add_argument("--heute", help="Datum für den Test, z. B. 2026-11-09")
    arg = teil.parse_args()

    quelle = Beispiel(arg.beispiel) if arg.beispiel else SharePoint()
    heute = dt.date.fromisoformat(arg.heute) if arg.heute else dt.datetime.now(BERLIN).date()
    neu, fertig = auswahl(list(quelle.eintraege()), heute)

    vorher = set(BILDER.glob("*.jpg")) if BILDER.exists() else set()
    seite = SEITE.read_text(encoding="utf-8")
    seite = einsetzen(seite, "NEUIGKEITEN", neuigkeiten_html(quelle, neu))
    seite = einsetzen(seite, "PROJEKTE", projekte_html(quelle, fertig))
    SEITE.write_text(seite, encoding="utf-8")

    genutzt = set(re.findall(r'img/auto/([^"]+\.jpg)', seite))
    for datei in vorher:
        if datei.name not in genutzt:
            datei.unlink()
    print(f"{len(neu)} Neuigkeiten, {len(fertig)} abgeschlossene Projekte eingetragen.")


if __name__ == "__main__":
    main()
