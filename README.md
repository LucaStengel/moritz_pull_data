# Rohstoffpreis-Tracker

Dieses Skript ruft täglich die aktuellen Marktpreise für **Latex**, **Baumwolle** und **Nitril/Butadien** von offiziellen Quellen ab, rechnet sie automatisch in **EUR/Tonne** um und speichert alles übersichtlich in einer Excel-Datei. Optional können Preisverlaufsgraphen angezeigt werden.

---

## Was macht das Skript genau?

Bei jedem Aufruf passiert Folgendes:

1. **Wechselkurse abrufen** – aktuelle EUR/USD- und EUR/CNY-Kurse werden live von einer kostenlosen API geladen
2. **Latex-Preis** – wird direkt von der offiziellen API des Malaysian Rubber Board (LGM) abgerufen
3. **Baumwoll-Future** – wird von Onvista.de (ICE Futures) abgerufen
4. **Nitril/Butadien-Preis** – Benchmark-Preis wird von der chinesischen Rohstoffplattform 100ppi.com abgerufen
5. **Alles wird in `prices.xlsx` gespeichert** – neue Zeile wird angehängt, bestehende Daten bleiben erhalten

**Ausgabe in der Konsole (Beispiel):**
```
Rufe aktuelle Preise und Wechselkurse ab …
  ✓ Wechselkurse: 1 EUR = 1.16 USD | 1 EUR = 7.90 CNY
  ✓ Latex:     758.0 USD/t  →  653.45 EUR/t
  ✓ Baumwolle: 77.37 USc/lb  →  1470.44 EUR/t
  ✓ Nitril:    12100.0 CNY/t  →  1531.65 EUR/t
```

---

## Aufbau der Excel-Datei (`prices.xlsx`)

| Spalte | Inhalt | Erklärung |
|--------|--------|-----------|
| A – Datum | `2026-05-27 14:32:00` | Zeitpunkt des Abrufs |
| B – Latex (USD/t) | `758.0` | Originalpreis in US-Dollar pro Tonne |
| C – Latex (EUR/t) | `653.45` | Umgerechnet in Euro pro Tonne |
| D – Baumwolle (USc/lb) | `77.37` | Originalpreis in US-Cent pro Pfund (Börsenkurs) |
| E – Baumwolle (EUR/t) | `1470.44` | Umgerechnet in Euro pro Tonne |
| F – Nitril (CNY/t) | `12100.0` | Originalpreis in Chinesischen Yuan pro Tonne |
| G – Nitril (EUR/t) | `1531.65` | Umgerechnet in Euro pro Tonne |
| H – EUR/USD Kurs | `1.16` | Kurs des Tages: 1 EUR entsprach X USD |
| I – EUR/CNY Kurs | `7.90` | Kurs des Tages: 1 EUR entsprach X CNY |

**Umrechnungsformeln:**
- Latex: `USD/t ÷ EUR/USD-Kurs`
- Baumwolle: `USc/lb × 22.046 (→ USD/t) ÷ EUR/USD-Kurs`
- Nitril: `CNY/t ÷ EUR/CNY-Kurs`

---

## Einmalige Installation

Die folgende Anleitung muss nur **ein einziges Mal** durchgeführt werden.

### Schritt 1 – uv installieren

`uv` ist das einzige Programm, das du installieren musst. Es kümmert sich selbstständig um Python und alle weiteren Pakete.

**Windows – Eingabeaufforderung öffnen** (Windows-Taste → `cmd` eintippen → Enter drücken) und folgendes eingeben:
```
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```
Danach die Eingabeaufforderung **schließen und neu öffnen**.

**Mac/Linux – Terminal öffnen** und eingeben:
```
curl -LsSf https://astral.sh/uv | sh
```
Danach das Terminal **schließen und neu öffnen**.

**Überprüfen** – in der neu geöffneten Eingabeaufforderung eingeben:
```
uv --version
```
Es sollte `uv x.x.x` erscheinen. Dann hat die Installation geklappt.

---

### Schritt 2 – Projektdateien herunterladen

Du bekommst die Dateien `main.py` und `pyproject.toml` – lege beide in einem Ordner ab, z. B.:

```
C:\Users\DeinName\Dokumente\rohstoffpreise\
    main.py
    pyproject.toml
```

---

### Schritt 3 – Abhängigkeiten installieren

Öffne die Eingabeaufforderung **in dem Ordner**, in dem du die Dateien abgelegt hast.

**Tipp für Windows:** Im Windows-Explorer den Ordner öffnen, dann in der Adressleiste oben `cmd` eintippen und Enter drücken – die Eingabeaufforderung öffnet sich direkt im richtigen Ordner.

Dann eingeben:
```
uv sync
```

`uv` lädt jetzt automatisch Python und alle benötigten Pakete herunter (das kann beim ersten Mal 1–2 Minuten dauern). Am Ende sollte stehen:
```
Installed XX packages ...
```

**Die Installation ist abgeschlossen.** Du musst Schritt 1–3 nie wieder durchführen.

---

## Tägliche Verwendung

### Preise abrufen und speichern

Eingabeaufforderung im Projektordner öffnen (wie oben beschrieben) und eingeben:

```
uv run main.py
```

Das Ergebnis erscheint direkt in der Konsole, und die Datei `prices.xlsx` wird im selben Ordner angelegt bzw. um eine neue Zeile ergänzt.

---

### Preisverlauf als Graph anzeigen

```
uv run main.py -graph
```
Zeigt Graphen für die letzten **30 Tage** (Standard).

```
uv run main.py -graph -7d
```
Zeigt Graphen für die letzten **7 Tage**.

```
uv run main.py -graph -90d
```
Zeigt Graphen für die letzten **90 Tage**.

Es öffnen sich nacheinander drei Fenster. Das nächste öffnet sich jeweils erst, wenn das aktuelle geschlossen wird.

1. **Rohstoffpreise in EUR/t** – alle drei Rohstoffe in einem Diagramm direkt vergleichbar
2. **Wechselkurs EUR/USD** – Kursverlauf des Euro gegenüber dem US-Dollar
3. **Wechselkurs EUR/CNY** – Kursverlauf des Euro gegenüber dem Chinesischen Yuan

> **Hinweis:** Für aussagekräftige Graphen sind mehrere gespeicherte Datenpunkte nötig – also das Skript idealerweise jeden Werktag einmal ausführen.

---

## Automatisch ausführen (Cronjob / Aufgabenplanung)

Damit das Skript jeden Werktag automatisch läuft, ohne dass man daran denken muss, kann es als geplante Aufgabe eingerichtet werden.

---

### Windows – Aufgabenplanung

1. Windows-Taste drücken, `Aufgabenplanung` eintippen und öffnen
2. Rechts auf **„Einfache Aufgabe erstellen…"** klicken
3. **Name** vergeben, z. B. `Rohstoffpreise`, dann **Weiter**
4. Trigger: **„Wöchentlich"** wählen → **Weiter**
5. Uhrzeit festlegen (z. B. `09:00 Uhr`), Wochentage **Mo–Fr** ankreuzen → **Weiter**
6. Aktion: **„Programm starten"** → **Weiter**
7. Felder ausfüllen:
   - **Programm/Skript:**
     ```
     uv
     ```
   - **Argumente hinzufügen:**
     ```
     run main.py
     ```
   - **Starten in** (den vollständigen Pfad zum Projektordner eintragen, z. B.):
     ```
     C:\Users\DeinName\Dokumente\rohstoffpreise
     ```
8. **Weiter** → **Fertig stellen**

> **Wichtig:** Der Computer muss zum eingestellten Zeitpunkt eingeschaltet und mit dem Internet verbunden sein. Falls er gerade ausgeschaltet war, holt die Aufgabenplanung den verpassten Lauf **nicht** automatisch nach.

**Verpassten Lauf nachholen aktivieren (optional):**
Nach dem Erstellen die Aufgabe in der Liste doppelklicken → Reiter **„Einstellungen"** → Haken setzen bei **„Aufgabe so bald wie möglich nach einem verpassten Start ausführen"**.

---

### Mac – Automatisch per cron

1. Terminal öffnen (Spotlight → `Terminal`)
2. Zuerst den vollständigen Pfad zu `uv` herausfinden:
   ```
   which uv
   ```
   Die Ausgabe sieht z. B. so aus: `/Users/DeinName/.local/bin/uv`
3. Cron-Editor öffnen:
   ```
   crontab -e
   ```
4. Mit den Pfeiltasten ans Ende der Datei navigieren, dann diese Zeile einfügen (Pfade anpassen):
   ```
   0 9 * * 1-5 cd /Users/DeinName/Dokumente/rohstoffpreise && /Users/DeinName/.local/bin/uv run main.py
   ```
   Das bedeutet: **jeden Werktag (Mo–Fr) um 9:00 Uhr** ausführen.
5. Speichern und beenden: `Escape` drücken, dann `:wq` tippen, Enter drücken

**Cron-Zeile erklären:**
```
0   9   *   *   1-5   <Befehl>
│   │   │   │   └── Wochentage: 1=Mo, 5=Fr
│   │   │   └────── Monat: * = jeden Monat
│   │   └────────── Tag: * = jeden Tag
│   └────────────── Stunde: 9
└────────────────── Minute: 0
```

---

### Linux – Automatisch per cron

Gleich wie Mac, Terminal öffnen und:
```
crontab -e
```
Dann dieselbe Zeile wie oben eintragen (Pfade entsprechend anpassen).

---

## Quellen

| Rohstoff | Quelle | Originaleinheit |
|----------|--------|-----------------|
| Latex | [Malaysian Rubber Board (LGM)](https://www.lgm.gov.my) – Grade: Latex in Bulk | USD/t |
| Baumwolle | [Onvista.de](https://www.onvista.de/rohstoffe/Baumwolle-Future-21474937) – ICE Futures U.S. | USc/lb |
| Nitril/Butadien | [100ppi.com](https://www.100ppi.com/vane/detail-886.html) – 生意社 Benchmark | CNY/t |
| Wechselkurse | [exchangerate-api.com](https://www.exchangerate-api.com) – kostenlose API | – |

---

## Fehlerbehebung

**„uv: command not found" oder „uv wird nicht erkannt"**
→ Die Eingabeaufforderung nach der uv-Installation nicht neu geöffnet. Schließen, neu öffnen, erneut versuchen.

**„✗ Latex: ..."** oder andere Fehlermeldungen beim Abrufen
→ Kurz warten und nochmal versuchen. Die Preise werden nur an Handelstagen aktualisiert (Montag–Freitag). An Wochenenden und Feiertagen kann es sein, dass kein neuer Preis verfügbar ist – der letzte bekannte Wert bleibt in der Excel erhalten.

**Die Excel-Datei lässt sich nicht öffnen während das Skript läuft**
→ Excel schließen, Skript ausführen, danach Excel wieder öffnen.

**Wechselkurs-Warnung in der Ausgabe**
→ Falls die Wechselkurs-API nicht erreichbar ist, werden automatisch Näherungswerte verwendet (ca. 1 EUR = 1,08 USD / 7,80 CNY). Die Preise werden trotzdem gespeichert.
