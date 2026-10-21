# Apollo Import GUI Prototype

Dieser Prototyp zeigt eine moegliche Desktop-GUI fuer die Erfassung neuer Artikel und erzeugt daraus die Importdateien als `.xlsx`.

Das Programm nutzt jetzt ein eigenes App-Logo aus `assets/apollo_import_logo.png` bzw. `assets/apollo_import_logo.ico`, das sowohl in der laufenden GUI als auch im Windows-`exe`-Build verwendet wird.

## Start

```powershell
python apollo_import_gui.py
```

Alternativ unter Windows:

```powershell
.\start_gui.bat
```

## Was die GUI kann

- Artikelnummer erfassen
- Zwei getrennte `Text Modul ID`s automatisch pro Artikel vergeben
  - eine fuer `Kurzbezeichnung`
  - eine fuer `Text`
- Kurztexte pro Sprache pflegen
  - Umlaute in Kurzbezeichnungen werden automatisch ersetzt, z. B. `ä -> ae`, `ö -> oe`, `ü -> ue`, `ß -> ss`
- Langtexte pro Sprache pflegen
- Attribute pro Artikel mit `TecDoc Kriterien ID`, Format-Hinweis und Wert pflegen
- Attribute automatisch finden und fuellen (hybrid: Regeln + Claude-KI)
  - erkennt `Label: Wert`-Zeilen und technische Daten (inkl. Bereichen und `LxBxH`-Abmessungen)
  - optional liest eine Claude-KI zusaetzlich Fliesstext und PDF-Dokumente aus
  - sichere, gegen die Stammdaten validierte Treffer werden automatisch uebernommen; unsichere landen im Bestaetigungsdialog
- Suchwoerter pro Artikel pflegen
  - werden beim Export als Referenz vom Typ `Gebrauchsnummer` (TecDoc ID `2`) in die `Gebrauchsnummern.xlsx` geschrieben, eine Zeile pro Suchwort
- OE-Nummern als freie Referenzen pro Artikel pflegen
- Vergleichsnummern mit Mitbewerber-ID pro Artikel pflegen
- Mitbewerber aus `KHer.csv` nachschlagen und direkt in Vergleichsnummern uebernehmen
- Fahrzeugverknuepfungen pro Artikel ueber Motorcodes pflegen
- Gefuehrter Modus (Wizard): fuehrt Schritt fuer Schritt durch alle Eingaben bis zum Export
- Tabs nach Befuellung gruppiert: erst die automatisch befuellten (Kunzer-Abruf, Praefix ⚙), dann die manuell zu pflegenden (Praefix ✎)
  - ueber einen oder mehrere Motorcodes werden die passenden KTyp-Nummern (TopMotive und TecDoc) automatisch aus den KTyp-Stammdaten (`KTyp.xlsx`) ergaenzt
  - waehrend der Eingabe werden Motorcode-Vorschlaege inklusive `meintest du`-Naehe angezeigt
  - der Fahrzeugtyp ist per Dropdown waehlbar (Standard `PKW = 2`)
- Bilder, Dokumente, Videos und Web Links erfassen
  - mit Bild-Thumbnail und PDF-Vorschau in den Bilder- und Dokumente-Tabs
  - vorhandene Zeilen koennen direkt im Formular nachbearbeitet und aktualisiert werden
  - Aenderungen werden direkt in die Output-Excel-Dateien geschrieben
- Kurzbezeichnungen und Texte per DeepL aus dem deutschen Feld in `EN`, `CZ`, `FR`, `IT` und `NL` uebersetzen
- GenArt manuell ueber die Suche auswaehlen
  - mit direkter Suche ueber das GenArt-Feld oder den Button `Suchen...`
- Produktdaten direkt von `kunzer.de` laden
  - per Artikelnummer
  - inklusive Titel, Web-Link, Bildern und Dokumenten
  - optional mit automatischer DeepL-Uebersetzung
  - mit Hintergrundladen, damit die GUI waehrenddessen fluessiger bleibt
- Eine Produktliste als CSV oder XLSX importieren und mehrere Produkte in Serie ueber Kunzer anlegen
  - mit Auswahl der zu scrapenden Daten
  - mit automatischer DeepL-Uebersetzung fuer Texte
  - mit direktem Schreiben in den festen Output-Pfad
- Exportdateien wahlweise in einen Zeitstempel-Unterordner oder direkt in einen festen Importpfad schreiben
- Vorhandene Artikel aus den Output-Dateien in einem Artikelverzeichnis erneut aufrufen
- Bild-, PDF- und Web-Vorschauen werden zwischengespeichert, damit wiederholtes Oeffnen spuerbar schneller ist
- Die GUI merkt sich beim Beenden ihren letzten Zustand, inklusive Pfaden, Optionen, aktuellem Artikel, Texten, Medien und geoeffnetem Tab
- Ueber `Nach Updates suchen` kann die App auf GitHub nach einem neueren Release schauen und in der gebauten EXE die neue Version direkt herunterladen und starten
- Die Projektansicht passt sich auf kleineren Displays automatisch an; grosse Bereiche und Vorschauen koennen eingeklappt werden

## Kunzer Import

Die GUI kann Produktdaten direkt von `https://www.kunzer.de` laden. Dabei nutzt sie die oeffentliche Produkt-Sitemap und rendert die Produktseite in einem Headless-Browser, damit auch clientseitig geladene Inhalte erfasst werden.

Aktuell werden automatisch befuellt:

- `Artikelnummer`
- `Kurzbezeichnung` auf Deutsch
- `Text` auf Deutsch
- `Bilder` als direkte URL von `kunzer.de`
- `Dokumente` als direkte URL von `kunzer.de`
- `Web Link`
- produktbezogene Video-Links, falls auf der Produktseite vorhanden

Bild- und Dokumentdaten aus Kunzer werden bewusst als URL in der GUI und im Export hinterlegt.
Bei Dokumenten wird die `Art` automatisch aus der Bezeichnung bzw. dem Dateinamen abgeleitet, zum Beispiel:

- `Bedienungsanleitung` -> `14`
- `Produktinfo` -> `17`
- `Zubehoer` -> `17`

Video-Links werden automatisch auf ein einbettbares YouTube-Format im Stil `https://www.youtube.com/embed/...` normalisiert. Das gilt fuer gescrapte Kunzer-Videos und auch fuer manuell eingetragene `embed`-, `watch`-, `shorts`- oder `youtu.be`-Links.

Wenn `DEEPL_API_KEY` gesetzt ist oder der Key in der GUI eingetragen wurde, koennen die geladenen deutschen Texte direkt weiter uebersetzt werden.

Die GenArt wird nach dem Laden bewusst nicht automatisch gesetzt, sondern manuell ueber das GenArt-Suchfeld oder den Button `Suchen...` ausgewaehlt.

## Artikelverzeichnis

Die bisherige reine Text-Vorschau wurde durch ein `Artikelverzeichnis` ersetzt. Dort werden nur Artikel aus den aktuellen Output-Dateien tabellarisch gelistet.

- Suchfeld ueber der Tabelle: mehrere Begriffe mit Leerzeichen trennen, alle muessen vorkommen (Artikelnummer, Kurz-/Text-ID, GenArt, Kurzbezeichnungen aller Sprachen, deutscher Text, OE-/Vergleichsnummern samt Hersteller, Suchwoerter); Treffer-Zaehler, `Enter` waehlt einen eindeutigen Treffer, `Esc` oder `Leeren` setzt zurueck
- Artikel koennen per Button oder Doppelklick wieder in die Maske geladen werden; der Wechsel blockiert die Oberflaeche nicht mehr (Bild-, PDF- und Video-Vorschauen laden im Hintergrund, die Liste wird nur bei geaenderten Dateien neu aufgebaut)
- fuer jeden Artikel werden Quelle, IDs und Medienanzahlen angezeigt
- rechts werden alle aktuell in den Export-Excel-Dateien vorhandenen Daten des markierten Artikels angezeigt
- Bilder zeigen ein Thumbnail, Dokumente bei PDFs die erste Seite und sonst eine kompakte Dateivorschau
- Web Links zeigen Host-Badge und Seitentitel; der Webseiten-Screenshot (Playwright) wird nur noch auf Knopfdruck `Webseiten-Vorschau laden` erzeugt und dann zwischengespeichert

## Leistung, Live-Speichern und Caches

- Live-Speichern laeuft in einem Hintergrund-Thread mit Warteschlange: Aenderungen in der Maske blockieren die Oberflaeche nicht mehr (Statuszeile zeigt `Speichere im Hintergrund ...` und danach `Live gespeichert: <Artikel>`). Geschrieben wird nur, was sich seit dem Laden bzw. letzten Speichern wirklich geaendert hat (Fingerabdruck je Artikel und Bereich); ein Fokusverlust ohne Aenderung schreibt nichts.
- Scheitert ein Bereich (z. B. Datei in Excel geoeffnet), werden die uebrigen Bereiche trotzdem geschrieben; der gescheiterte wird beim naechsten Anlass nachgeholt und beim Beenden gibt es eine Warnung, falls noch etwas offen ist. Export, Loeschen und Massenimport warten vorher auf die Warteschlange.
- Excel-Dateien werden immer erst als Temp-Datei geschrieben und dann eingetauscht - eine gesperrte Zieldatei bleibt vollstaendig erhalten.
- Lesen und Schreiben der Output-Dateien nutzt einen eigenen XLSX-Leser/-Schreiber (Inline-Strings, gleiches Dateiformat wie bisher) plus Zeilen-Cache im Speicher; dadurch dauert z. B. das Schreiben der `Fahrzeugverknuepfungen.xlsx` (16.000 Zeilen) rund 0,1 s statt 2 s.
- Stammdaten (KTyp, Attribute, Schluesselwerte, GenArten) und die geparsten Output-Dateien werden unter `%LOCALAPPDATA%\Apollo Import GUI\cache` zwischengespeichert; der Cache erkennt geaenderte Quelldateien und neue Programmversionen selbst. Der Ordner darf jederzeit geloescht werden (naechster Start parst dann einmal neu, KTyp ca. 10 s im Hintergrund).
- Schwere Pakete (Playwright, PyMuPDF, Anthropic) werden erst beim ersten Gebrauch geladen; der Start ist dadurch deutlich schneller.
- Leistungsprotokoll: UI-Blockaden ueber 250 ms landen mit Name und Dauer in `apollo_import_perf.log` im Arbeitsverzeichnis (neben `apollo_import_gui_state.json`), max. 2 MB mit Rotation - hilfreich, wenn etwas langsam wirkt.

## Listenimport CSV/XLSX

Ueber `Produktliste CSV/XLSX` kann eine Liste von Produkten importiert werden. Jeder Eintrag wird ueber `kunzer.de` geladen und direkt in den festen Output-Pfad geschrieben.

Im Bereich `Listenimport` kann ausgewaehlt werden, welche Daten geholt werden sollen:

- `Kurzbezeichnung`
- `Text`
- `Bilder`
- `Dokumente`
- `Videos`
- `Web Links`
- `Attribute` (automatische Attribut-Findung, siehe Abschnitt "Attribute automatisch finden")

Wenn `Kurzbezeichnung` oder `Text` aktiviert sind, wird die Uebersetzung im Listenimport automatisch ueber DeepL ausgefuehrt. Dafuer muss ein gueltiger API Key eingetragen sein.

Ist `Attribute` aktiviert, laeuft pro Artikel die automatische Attribut-Findung:

- sichere, validierte Treffer werden direkt in die `Attribute.xlsx` geschrieben - dabei werden nur Kriterien-IDs ergaenzt, die der Artikel noch nicht hat; bestehende (auch manuell gepflegte) Zeilen bleiben unveraendert
- unsichere Treffer und nicht zugeordnete Angaben landen in der Datei `Attribute_Pruefliste.xlsx` im Output-Ordner
- faellt die KI aus (z. B. fehlender oder ungueltiger API Key), laeuft der Batch regelbasiert weiter; der Hinweis erscheint in der Warnungsliste
- **Testlauf** (standardmaessig aktiv): mit der Checkbox `Testlauf: nur Bericht, nichts schreiben` wird NICHTS in die Attribute.xlsx geschrieben - stattdessen entsteht der vollstaendige `Attribut-Testbericht.xlsx` (je Angabe: Attribut, Wert, Status, Quelle Regel/KI). Ideal, um einen grossen Lauf erst zu pruefen. Der Bericht wird auch im scharfen Lauf als Protokoll geschrieben
- ab 4 Artikeln laufen die KI-Aufrufe gesammelt ueber die Anthropic **Batch API** (50 % Rabatt bei gleicher Qualitaet)
- der Abschlussdialog zeigt den KI-Verbrauch (Token und ungefaehre Kosten in $)

Unterstuetzte Spalten sind flexibel. Mindestens eine der beiden Gruppen muss vorhanden sein:

- Artikelnummer: `Artikelnummer`, `Produktnummer`, `ArtNr`, `SKU`
- Produkt-URL: `Kunzer Produkt-URL`, `Kunzer URL`, `Produkt-URL`, `URL`, `Link`

CSV-Dateien duerfen mit `;`, `,`, Tab oder `|` getrennt sein. XLSX-Dateien werden aus dem ersten Tabellenblatt gelesen. Neue Artikel werden im festen Importpfad angehaengt; existiert die Artikelnummer bereits, werden nur deren alte Zeilen ersetzt.

## Gefuehrter Modus (Wizard)

Der Button `Gefuehrter Modus` oben rechts startet eine Schrittleiste, die nacheinander durch alle Tabs fuehrt: Artikelnummer, Kurzbezeichnung, Produkttext, GenArt, Attribute, Suchwoerter, OE-Nummern, Vergleichsnummern, Fahrzeuge, Bilder, Dokumente und Links. Zum Schluss zeigt eine Zusammenfassung alle Zaehler und `Exportieren & Fertigstellen` schreibt die Importdateien.

- Pflichtschritte (Artikelnummer, deutsche Kurzbezeichnung, deutscher Produkttext, mindestens eine GenArt, Hersteller je OE-Nummer) lassen erst weiter, wenn sie gueltig sind; die Fehlermeldung erscheint direkt in der Leiste
- Optionale Schritte koennen mit `Weiter` uebersprungen werden
- `Zurueck` und `Beenden` sind jederzeit moeglich; alle Eingaben landen in den normalen Tabs und bleiben erhalten

Damit koennen auch neue Kolleginnen und Kollegen Artikel vollstaendig erfassen, ohne die Reihenfolge oder Pflichtfelder zu kennen.

## Fahrzeugverknuepfungen

Im Tab `Fahrzeuge` koennen einem Artikel Fahrzeuge (KTyp-Nummern) zugeordnet werden. Grundlage ist die KTyp-Stammdatendatei (`KTyp.xlsx`), die im Projekt-Tab unter `Datenstaemme` gewaehlt und jederzeit ueber `Neu laden` aktualisiert werden kann. Standardpfad ist `G:\Apollo\KTyp.xlsx`.

Ablauf:

- einen oder mehrere Motorcodes eingeben (mehrere mit `;` trennen)
- waehrend der Eingabe erscheinen Vorschlaege fuer den gemeinten Motorcode inklusive `meintest du`-Naehe (Tippfehler-tolerant)
- mit `Fahrzeuge hinzufuegen` bzw. `Enter` werden **alle** passenden Fahrzeuge aus den Stammdaten uebernommen (Hersteller, Modell, Bezeichnung, Bauzeit, Leistung, KTyp TopMotive und KTyp TecDoc)
- der Fahrzeugtyp ist per Dropdown waehlbar; Standard ist `PKW (2)`, weitere sind `NKW (16)`, `Motor (14)` und `Achse (19)`

Die Suche ueber Motorcodes ist robust gegen Gross-/Kleinschreibung, Leerzeichen und Klammerzusaetze (z. B. `108C (XV8)` findet auch `108C`).

Der Export erzeugt `Fahrzeugverknuepfungen.xlsx` mit den Spalten:

- `Artikelnummer`
- `Fahrzeugtyp` (fuer den Apollo-Import auf `TecDoc Verknuepfungstyp ID` mappen, `PKW = 2`)
- `KTypNr` (fuer den Apollo-Import auf `TecDoc Verknuepfungs ID` mappen)
- `KTyp-System` (Info, ob die `KTypNr` aus dem `Topmotive`- oder `TecDoc`-Nummernkreis stammt)
- `GenArt ID` und `GenArt Bezeichnung` (die GenArt des Artikels aus dem GenArt-Tab)

Pro Fahrzeug werden zwei Zeilen geschrieben: eine mit der TopMotive-Nummer und eine mit der TecDoc-Nummer. So kann im Apollo-Import genau die eine passende Spalte auf `TecDoc Verknuepfungs ID` gemappt werden. Ueber die Spalte `KTyp-System` lassen sich die Zeilen bei Bedarf vorher filtern. Jede Zeile traegt zusaetzlich die GenArt des Artikels; sind einem Artikel mehrere GenArts zugeordnet, wird je GenArt und KTyp-Nummer eine eigene Zeile geschrieben. Bestehende `Fahrzeugverknuepfungen.xlsx` ohne GenArt-Spalten werden beim naechsten Schreiben automatisch auf das neue Format migriert (GenArt-Spalten bleiben bei alten Zeilen leer, bis der Artikel neu exportiert wird).

## Attribute automatisch finden

Die Attribut-Findung arbeitet hybrid in zwei Stufen und wird ueber den Button `Attribute automatisch ausfuellen` im Attribute-Tab oder automatisch nach dem Kunzer-Laden gestartet.

**Stufe 1 - Regeln (immer aktiv, kostenlos):** Erkannt werden `Label: Wert`- und Tab-getrennte Zeilen aus Technische-Daten-Bloecken, z. B.:

- `Gewicht: 612 g` -> `Gewicht [kg]` = `0,612` (Einheiten werden automatisch umgerechnet)
- `Hubhoehe: 150 - 530 mm` -> Bereich als `Wert` + `Wert bis`
- `Abmessungen (LxBxH): 155 x 17 x 13,5 mm` -> aufgeteilt in `Laenge`, `Breite`, `Hoehe`
- Schluesselwert-Attribute werden nur uebernommen, wenn der Wert exakt in der Werteliste steht

Grundlage ist die pflegbare Zuordnungsdatei `Attribut_Zuordnung.xlsx` (Standard `G:\Apollo\Attribut_Zuordnung.xlsx`, Spalten `Text-Label` und `TecDoc Kriterien ID`), die im Projekt-Tab unter `Datenstaemme` konfiguriert wird.

**Stufe 2 - Claude-KI (optional, braucht Anthropic API Key):** Alles, was die Regeln nicht zuordnen konnten (Fliesstext, Synonyme, unbekannte Labels), geht zusammen mit dem Text der verlinkten PDF-Dokumente an die Claude API. Ueber den zweiten Button `Automatisch ermitteln (Internet)` recherchiert die KI zusaetzlich per Websuche auf Herstellerseiten und in Datenblaettern (ca. 10-30 ct pro Artikel; nur im Einzelartikel, nicht im Batch). Auch Recherche-Treffer zaehlen nur, wenn sie eindeutig zum Artikel gehoeren und die lokale Validierung bestehen. Die KI bekommt dabei nur eine lokal vorgefilterte Kandidatenliste (max. 300 plausible Attribute statt aller ~5.000) - das haelt die Kosten klein. Der API Key und das Modell werden im Projekt-Tab unter `APIs` eingestellt (auch per Umgebungsvariable `ANTHROPIC_API_KEY`); Standardmodell ist `claude-opus-4-8`, guenstigere Modelle sind waehlbar.

**Ohne Text in der Maske:** Ist kein deutscher Text erfasst (z. B. weil die Texte bereits in Apollo gepflegt sind), holt sich die Attribut-Findung den Produkttext selbst von kunzer.de - nur zum Auswerten. Gespeichert wird er nirgends; Texte, IDs und alle anderen Daten bleiben unangetastet.

**Gestufte Uebernahme:** Jeder Treffer - egal ob aus Regeln oder KI - wird lokal streng gegen die Attribut-Stammdaten validiert (Zahlformat, Einheiten-Umrechnung, Schluesselwert exakt in Werteliste, maximale Laenge). Nur validierte Treffer werden automatisch uebernommen, und nur fuer Kriterien-IDs, die der Artikel noch nicht hat (bestehende Zeilen werden nie ueberschrieben). Unsichere Treffer erscheinen wie bisher im Bestaetigungsdialog; dort koennen nicht erkannte Angaben manuell zugewiesen und mit `Zuordnung merken` dauerhaft gelernt werden.

**Zusatzinfo-Slots als Auffangbecken:** Wichtige technische Angaben, die zu keinem Attribut passen, liefert die KI als kurze Stichworte (max. 20 Zeichen, z. B. `IP54`, `inkl. Koffer`). Da Apollo pro Kriterien-ID nur eine Zeile akzeptiert, verteilt die App die wichtigsten Infos auf drei Slots: `9202 Zusatzinfo`, `9213 Zusatz LI/SI 2`, `9220 Information`. Weitere Infos landen im Pruefdialog bzw. in der Pruefliste.

**Label-Mining (Selbst-Training):** Der Button `Labels per KI zuordnen` bei den Datenstaemmen analysiert alle Texte im Output-Ordner, sammelt unbekannte Text-Labels samt Beispielwerten und laesst die KI sie EINMALIG zuordnen. Uebernommen wird nur, was an echten Beispielwerten lokal validiert wurde (Hinweis `KI-Miner` in der Zuordnungsdatei); der Rest landet in der `Label-Zuordnung-Pruefliste.xlsx`. Danach loest die kostenlose Regel-Stufe diese Labels ohne KI-Aufruf - der Grossteil aller Spezifikationszeilen wird damit gratis.

**Lernschleife:** Ordnet die KI ein neues Text-Label erfolgreich zu (z. B. `Farbton` -> `Farbe`), wird die Zuordnung mit Hinweis `KI` in die `Attribut_Zuordnung.xlsx` geschrieben. Beim naechsten Artikel greift dann schon die Regel-Stufe - ohne API-Aufruf. KI-Eintraege sind in der Datei ueber die Hinweis-Spalte auffindbar und koennen dort jederzeit korrigiert oder geloescht werden.

Die KI erhaelt den Produkttext als reinen Inhalt und darf nur Attribute aus der mitgeschickten Kandidatenliste waehlen; alles andere wird verworfen. Ohne API Key (oder bei API-Fehlern) laeuft die Findung vollstaendig regelbasiert weiter.

## Marke (Attribut 8988)

Die Marke eines Artikels wird in Apollo als Attribut `8988 Marke` (Format `Alphanumerisch`, max. 20 Zeichen) gepflegt. Bei jedem Kunzer-Abruf - im Einzelartikel (`Artikel anlegen`) wie im Massenimport (`Liste importieren`, `Artikelliste aktualisieren`) - setzt die App diese Zeile automatisch, rein lokal und ohne KI-Aufruf. Die Marke steht auf kunzer.de nicht im Seitentext, sondern wird aus den Signalen der Produktseite erkannt, in dieser Reihenfolge (der erste Treffer gewinnt):

1. **Medienordner**: erstes Pfadsegment der `restapi.kunzer.de`-Links von Bildern und Dokumenten, z. B. `airman`, `tdl`, `weber hydraulik`, `moto meter`, `norma`, `oetiker`, `suprabeam`, `handschuhe` (Tiger Grip) oder `kunzer werkzeuge`
2. **Kategoriepfad** der Produktseite (Breadcrumb), z. B. `Produkte > Werkzeuge > Rad und Reifen > airMAN` oder `Produkte > Hebetechnik > Lufthydraulische Heber T D L`
3. **Shop-Suchwoerter** des Artikels (Feld `SUCHWOERTER` der Produktdaten, z. B. `Dichtmilch;Pannenpilot;Airman`)
4. **Titel, Kurzbezeichnung, Text** - Treffer nur an Wortgrenzen (`Norma` trifft nicht `normal`); `Kunzer` zaehlt in Texten nicht als Treffer, weil es ohnehin der Standard ist
5. **Link-Labels** der Produktseite (z. B. `AIRMAN - 71-051-011`)
6. ohne Treffer: `Kunzer` (Kunzer-Eigenartikel tragen zudem EAN mit GS1-Praefix `4260174`)

Erkannt werden nur Marken aus der **Markenliste** (Einstellungen, Tab `API & Datenstaemme`, Block `Marke (Attribut 8988)`). Jeder Eintrag nennt den kanonischen Namen, der in die `Attribute.xlsx` geschrieben wird, und optional Schreibweisen (Aliasse): Eintraege mit `;` trennen, Aliasse nach `=` mit `,` - Gross-/Kleinschreibung spielt keine Rolle, Leerzeichen und Bindestriche in Aliassen sind beliebig (`T D L` trifft auch `TDL`, `weber hydraulik` auch `Weber-Hydraulik`). Standard:

```
Kunzer; Airman=airman,airMAN,air man; TDL=T D L,tdl; Weber Hydraulik=weber hydraulik,weber-hydraulik; MotoMeter=moto meter,motometer,moto-meter; Norma; Oetiker; Suprabeam; Gefo; Tiger Grip=handschuhe,tigergrip,tiger-grip; Argus
```

Die Liste wird mit dem Sitzungszustand gespeichert; `Standard` stellt die eingebaute Liste wieder her. Kanonische Namen muessen in 20 Zeichen passen (laengere werden abgeschnitten).

Verhalten:

- **Manuelle Pflege gewinnt:** Hat der Artikel bereits eine `8988`-Zeile mit Wert (in der `Attribute.xlsx` bzw. im Artikelverzeichnis), bleibt sie unveraendert - die erkannte Marke wird nie darueber geschrieben. Nur eine leere `8988`-Zeile wird ersetzt. Pro Artikel gibt es genau eine `8988`-Zeile; die automatische Attribut-Findung legt keine zweite an.
- **Einzelabruf:** Bis v0.1.35 wurden beim `Artikel anlegen` alle Attribute der Maske geleert und dieser Leerstand live geschrieben - vorhandene Attributzeilen des Artikels gingen damit verloren. Jetzt bleiben die gespeicherten Attributzeilen des Artikels erhalten (Grundlage ist der gespeicherte Stand des Artikels, nicht die Maske: zeigt die Maske noch einen zuvor geladenen anderen Artikel, gehen dessen Zeilen nicht auf den neuen Artikel ueber), die Marke wird ergaenzt und zusammen mit den uebrigen Daten live gespeichert - auch wenn nur Medien im Abrufumfang sind.
- **Massenimport:** Die Marke wird fuer jeden Artikel ergaenzt (add-only in die `Attribute.xlsx`), unabhaengig von der Option `Attribute (automatisch, Regeln + KI)` und auch im Testlauf. Der Abschlussdialog nennt, wie viele Artikel eine Marke bekommen haben und wie viele schon eine hatten; der `Attribut-Testbericht.xlsx` enthaelt je Artikel die erkannte Marke mit Signal (Status `automatisch` bzw. `vorhanden`, Quelle `Marke`).
- Die Werkzeuge mit Kunzer-Artikelnummer im Medienordner `kunzer werkzeuge` (z. B. die Oetiker-Zangen 14100392/14100501, bei denen nur der Titel `Oetiker` nennt) bekommen ueber den Medienordner `Kunzer`; soll der Titel Vorrang haben, muss die Zeile manuell gepflegt werden.

## Suchwoerter

Suchwoerter sind in Apollo Referenzen vom Typ `Gebrauchsnummer` (Reiter `Referenzen`, Typ mit TecDoc ID `2`); das Suchwort steht als Freitext in der Spalte `Nummer`. Im Tab `Suchwoerter` koennen deshalb pro Artikel freie Suchbegriffe gepflegt werden, die beim Export in eine eigene Referenzdatei geschrieben werden:

- jedes Suchwort wird eine eigene Zeile in `Gebrauchsnummern.xlsx`
- Spalten: `Artikelnummer`, `TecDoc ID` = `2`, `Gebrauchsnummer` = Suchwort (gleicher Aufbau wie `OE-Nummern.xlsx`, nur ohne Hersteller-ID)
- maximal 50 Zeichen pro Suchwort; laengere Eingaben werden abgewiesen
- mehrere Suchwoerter koennen mit `;` getrennt auf einmal eingegeben werden
- Duplikate werden automatisch ignoriert

Bis v0.1.31 wurden Suchwoerter faelschlich als Attribut `Zusatzbezeichnung` (Kriterien ID `9595`) in die `Attribute.xlsx` geschrieben. Solche Altzeilen werden beim Laden aus den Output-Dateien weiterhin als Suchwoerter erkannt und beim naechsten Speichern des Artikels nach `Gebrauchsnummern.xlsx` ueberfuehrt; aus der `Attribute.xlsx` verschwinden sie dabei.

### Import aus Excel/CSV

Ueber `Aus Datei importieren…` im Tab `Suchwoerter` lassen sich Suchwoerter fuer viele Artikel auf einmal aus einer XLSX- (Blatt 1) oder CSV-Datei (Trennzeichen automatisch, utf-8 oder Windows-1252) uebernehmen. Nach der Dateiauswahl zeigt ein Dialog die Vorschau (Anzahl Artikel und Suchwoerter, uebersprungene Eintraege, Artikel, die noch nicht im Verzeichnis stehen) und fragt den Modus ab:

- `Ergaenzen`: vorhandene Suchwoerter der betroffenen Artikel bleiben, neue kommen dazu (Duplikate ohne Gross-/Kleinschreibung werden ignoriert)
- `Ersetzen`: die Suchwoerter der in der Datei genannten Artikel werden komplett durch die Datei ersetzt
- Artikel, die nicht in der Datei stehen, bleiben in beiden Faellen unberuehrt; Artikel, die noch nicht im Verzeichnis sind, werden nur mit ihren Suchwoertern angelegt

Akzeptierte Spalten (Kopfzeile Pflicht, Gross-/Kleinschreibung egal):

- Artikelspalte: `Artikelnummer`, `ArtikelNr`, `ArtNr`, `Produktnummer`, `SKU` oder `Article`; numerische Zellen (z. B. `14100501`) werden als Text uebernommen
- Suchwortspalte(n): `Suchwort`, `Suchwoerter`, `Suchbegriff(e)`, `Gebrauchsnummer(n)`, `Keyword(s)` oder `Nummer` - auch mehrere Spalten wie `Suchwort 1`, `Suchwort 2`, `Keyword 3`, die zusammengefuehrt werden; eine exportierte `Gebrauchsnummern.xlsx` laesst sich damit direkt wieder einlesen
- eine Zelle darf mehrere Suchwoerter enthalten, getrennt durch `;`, `|`, Zeilenumbruch oder Komma - ein Komma zwischen zwei Ziffern trennt nicht, `10,5 mm` oder `1,5 Zoll` bleiben ein Suchwort
- bei CSV-Dateien bestimmt die Kopfzeile das Trennzeichen (`;`, `,`, Tab oder `|`); ueberzaehlige Zellen einer Datenzeile (z. B. `14100501;Bremse;Scheibe` ohne Anfuehrungszeichen) zaehlen zur letzten Suchwortspalte, wenn diese die letzte Spalte der Kopfzeile ist - sonst werden sie als "ohne Spaltenkopf" uebersprungen und in Vorschau und Zusammenfassung genannt
- fehlt die Artikelspalte, gelten alle Suchwoerter fuer den Artikel, der gerade in der Maske steht (ohne Artikelnummer in der Maske bricht der Import mit Hinweis ab); fehlt die Suchwortspalte bei genau zwei Spalten, gilt die Spalte neben der Artikelnummer als Suchwortspalte
- leere Eintraege und Suchwoerter ueber 50 Zeichen werden uebersprungen und in der Zusammenfassung aufgefuehrt

Der Import schreibt die `Gebrauchsnummern.xlsx` in einem Zug (Zeitstempel unveraenderter Zeilen bleiben erhalten) und aktualisiert Artikelverzeichnis und Maske direkt, ohne alle Output-Dateien neu zu laden; vorher gleicht er das Verzeichnis einmal mit dem Ausgabeordner ab, damit die Vorschau die noch nicht vorhandenen Artikel richtig zaehlt. Alt-Suchwoerter (Kriterien-ID `9595` in der `Attribute.xlsx`) liest der Import direkt aus der Datei: beim Ergaenzen wandern sie mit in die `Gebrauchsnummern.xlsx`, beim Ersetzen entfallen sie, ihre Altzeilen verschwinden in beiden Faellen aus der `Attribute.xlsx`. Ist die `Attribute.xlsx` dabei gesperrt (z. B. in Excel geoeffnet), bleiben die Suchwoerter trotzdem vollstaendig importiert; die Zusammenfassung weist darauf hin, und ein erneuter Import derselben Datei raeumt die Altzeilen nach.

### Automatisch erzeugen

Suchwoerter entstehen auch ohne Tippen - regelbasiert immer, per KI optional:

- **Beim Kunzer-Laden** (`Artikel anlegen`): die Maske zeigt die bereits gespeicherten Suchwoerter des Artikels plus die regelbasierten Kandidaten; nichts wird geloescht. Quellen in dieser Reihenfolge: die Marke (Wert des Attributs `8988 Marke`), die Shop-Suchwoerter von kunzer.de (Feld `SUCHWOERTER` der Produktseite, z. B. `Dichtmilch;Pannenhilfe;Airman`), die GenArt-Bezeichnungen des gespeicherten Standes (Komma-Listen wie `Dichtmittel-Set, Reifen-Dichtungs-Set` je Variante einzeln) und die Woerter der Kurzbezeichnung (ohne Fuellwoerter, Artikelnummer, Zahlen/Masse/Codes; Bindestrich-Komposita wie `Motor-Einstellwerkzeugsatz` zusaetzlich als Ganzes und in ihren Teilen). Das Live-Speichern schreibt die Suchwoerter mit dem Artikel in die `Gebrauchsnummern.xlsx`.
- **Button `Suchwoerter automatisch`** im Tab `Suchwoerter`: zeigt alle neuen Kandidaten (Spalte `Quelle`: Marke, kunzer.de, GenArt, Titel, KI) mit Haekchen zur Auswahl; `Uebernehmen` haengt die angehakten Suchwoerter an die vorhandenen an und speichert. Laeuft gerade ein Hintergrundvorgang (z. B. die Attribut-Findung), erscheint der Dialog sofort nur mit den Regel-Kandidaten und einem Hinweis - spaeter erneut druecken holt die KI-Vorschlaege nach.
- **KI-Vorschlaege** (Einstellungen > `API` > `Suchwoerter per KI vorschlagen`, Standard an, wirksam nur mit Claude API Key): bis zu 10 Begriffe, wie Werkstaetten und Haendler im Teilekatalog suchen - Synonyme, Umgangssprache, Produktgattung, Einsatzzweck; keine Marken- oder Artikelnummern, keine Saetze, je maximal 50 Zeichen. Nach dem Kunzer-Laden laufen sie als Nachlauf hinter der automatischen Attribut-Findung und werden ohne Rueckfrage angehaengt und gespeichert (Statuszeile nennt die Anzahl); beim Button landen sie mit im Auswahldialog. Die KI bekommt den Produkttext als reinen Inhalt; ihre Vorschlaege werden lokal nachgeprueft (Laenge, keine Artikelnummer/Marke, keine Dubletten).
- **Dubletten**: Vergleich ohne Gross-/Kleinschreibung, Umlaute, Bindestriche und Leerzeichen (`T D L` = `TDL`, `Reifen-Dichtmittel` = `Reifendichtmittel`); Terme ueber 50 Zeichen oder nur aus der Artikelnummer entfallen.
- **Massenimport**: je Artikel werden die Regel-Kandidaten gesammelt und am Ende des Laufs in EINEM Schreibvorgang (Modus `Ergaenzen`, wie der Datei-Import) in die `Gebrauchsnummern.xlsx` geschrieben; vorhandene Suchwoerter bleiben. KI-Vorschlaege gibt es im Massenimport nur mit `Suchwoerter per KI auch im Massenimport` (Standard aus); vor dem Start zeigt die App die geschaetzten Kosten und fragt nach. Das Abschlussfenster nennt die Anzahl der ergaenzten Suchwoerter; KI-Fehler je Artikel stehen in der Warnungsliste.

## TecDoc Anhangsformattyp ID

Die Exportdateien fuer Bilder, Dokumente, Videos und Web Links enthalten zusaetzlich die Spalte `TecDoc Anhangsformattyp ID`.

Aktuelle Zuordnung:

- `JPG` und `JPEG` -> `3`
- `PNG` -> `6`
- `PDF` -> `2`
- `GIF` -> `7`
- jede `http/https` URL -> `4`

Das gilt jetzt bewusst auch dann, wenn eine URL auf `.png`, `.jpg`, `.jpeg` oder `.pdf` endet. Nur lokale/dateibasierte Pfade werden noch ueber die Dateiendung typisiert.

Wenn fuer einen Anhang kein unterstuetztes Format erkannt wird, bricht der Export mit einer klaren Fehlermeldung ab, damit keine fehlerhaften Importdateien entstehen.

## DeepL

Der API Key wird bewusst nicht fest im Quellcode gespeichert. Du kannst ihn in der GUI eintragen oder vor dem Start als Umgebungsvariable setzen:

```powershell
$env:DEEPL_API_KEY="dein-key"
python apollo_import_gui.py
```

Standardmaessig nutzt die GUI `https://api.deepl.com`. Fuer DeepL Free kannst du in der GUI oder per `DEEPL_API_BASE_URL` auf `https://api-free.deepl.com` wechseln.

## Exportierte Dateien

Die GUI erzeugt aktuell diese Dateien:

- `Kurzbezeichnung-NEU.xlsx`
- `Kurzbezeichnung_zu_ID.xlsx`
- `Text-NEU.xlsx`
- `Attribute.xlsx`
- `OE-Nummern.xlsx`
- `Vergleichsnummern.xlsx`
- `Gebrauchsnummern.xlsx`
- `Fahrzeugverknuepfungen.xlsx`
- `Bilder.xlsx`
- `Dokumente.xlsx`
- `Videos.xlsx`
- `Web Link.xlsx`

Alle Output-Dateien enthalten zusaetzlich die Spalte `Zuletzt geschrieben am`. Darin steht pro Zeile, wann diese Daten zuletzt aus der GUI geschrieben wurden.

## Hinweis

Die GUI kann auf zwei Arten exportieren:

- mit Zeitstempel-Unterordner, damit alte Exporte erhalten bleiben
- direkt in einen festen Ausgabeordner, damit immer derselbe Importpfad genutzt werden kann

Im festen Ausgabeordner bleiben die zehn Exportdateien bestehen. Neue Artikel werden angehaengt, und bei erneutem Export derselben Artikelnummer werden nur die alten Zeilen dieses Artikels ersetzt.
Bei einer neuen Artikelnummer werden die `Text Modul ID`s automatisch im bestehenden 6-stelligen Format erzeugt und gegen vorhandene IDs geprueft. Bestehende Artikel behalten ihre bereits gespeicherten IDs.

Die Output-Dateien dienen dabei als Live-Datenbank: Aenderungen an Texten, Bildern, Dokumenten, Videos und Links werden aus der GUI direkt in diese Excel-Dateien zurueckgeschrieben.

## Kompaktlayout

Die GUI ist jetzt auf kleinere Displays wie 14-Zoll-Laptops ausgelegt:

- im Projekt-Tab werden `Artikel`, `Export`, `APIs` und `Artikelverzeichnis` bei schmalerer Breite automatisch untereinander angeordnet
- `Artikel`, `APIs` und `Artikelverzeichnis` koennen direkt eingeklappt und wieder eingeblendet werden
- die Vorschauen in `Bilder`, `Dokumente`, `Videos` und `Web Links` lassen sich ausblenden und werden bei wenig Breite automatisch unter die Tabelle verschoben
- diese Sichtbarkeits-Einstellungen werden zusammen mit dem restlichen GUI-Zustand gespeichert

## Updates

Im Projekt-Tab gibt es einen Bereich `App-Update`.

- `Nach Updates suchen` fragt das neueste GitHub-Release des Repositories ab
- wenn ein neueres Release gefunden wird, laedt die GUI bevorzugt ein `Setup`- oder `onefile`-Paket herunter
- laeuft die Anwendung bereits als gebaute Windows-EXE, kann sie sich selbst schliessen, die EXE ersetzen und danach automatisch neu starten
- wenn die GUI aus dem Python-Quellcode gestartet wurde, wird die Release-Datei nur heruntergeladen und der Ordner geoeffnet
