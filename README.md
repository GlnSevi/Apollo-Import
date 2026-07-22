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
  - werden beim Export als Attribut `Zusatzbezeichnung` (TecDoc Kriterien ID `9595`) in die `Attribute.xlsx` geschrieben, eine Zeile pro Suchwort
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

- Artikel koennen per Button oder Doppelklick wieder in die Maske geladen werden
- fuer jeden Artikel werden Quelle, IDs und Medienanzahlen angezeigt
- rechts werden alle aktuell in den Export-Excel-Dateien vorhandenen Daten des markierten Artikels angezeigt
- Bilder zeigen ein Thumbnail, Dokumente bei PDFs die erste Seite und sonst eine kompakte Dateivorschau

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

**Stufe 2 - Claude-KI (optional, braucht Anthropic API Key):** Alles, was die Regeln nicht zuordnen konnten (Fliesstext, Synonyme, unbekannte Labels), geht zusammen mit dem Text der verlinkten PDF-Dokumente an die Claude API. Die KI bekommt dabei nur eine lokal vorgefilterte Kandidatenliste (max. 300 plausible Attribute statt aller ~5.000) - das haelt die Kosten klein. Der API Key und das Modell werden im Projekt-Tab unter `APIs` eingestellt (auch per Umgebungsvariable `ANTHROPIC_API_KEY`); Standardmodell ist `claude-opus-4-8`, guenstigere Modelle sind waehlbar.

**Gestufte Uebernahme:** Jeder Treffer - egal ob aus Regeln oder KI - wird lokal streng gegen die Attribut-Stammdaten validiert (Zahlformat, Einheiten-Umrechnung, Schluesselwert exakt in Werteliste, maximale Laenge). Nur validierte Treffer werden automatisch uebernommen, und nur fuer Kriterien-IDs, die der Artikel noch nicht hat (bestehende Zeilen werden nie ueberschrieben). Unsichere Treffer erscheinen wie bisher im Bestaetigungsdialog; dort koennen nicht erkannte Angaben manuell zugewiesen und mit `Zuordnung merken` dauerhaft gelernt werden.

**Lernschleife:** Ordnet die KI ein neues Text-Label erfolgreich zu (z. B. `Farbton` -> `Farbe`), wird die Zuordnung mit Hinweis `KI` in die `Attribut_Zuordnung.xlsx` geschrieben. Beim naechsten Artikel greift dann schon die Regel-Stufe - ohne API-Aufruf. KI-Eintraege sind in der Datei ueber die Hinweis-Spalte auffindbar und koennen dort jederzeit korrigiert oder geloescht werden.

Die KI erhaelt den Produkttext als reinen Inhalt und darf nur Attribute aus der mitgeschickten Kandidatenliste waehlen; alles andere wird verworfen. Ohne API Key (oder bei API-Fehlern) laeuft die Findung vollstaendig regelbasiert weiter.

## Suchwoerter

Apollo kennt keine nativen Keywords oder Suchbegriffe. Im Tab `Suchwoerter` koennen deshalb pro Artikel freie Suchbegriffe gepflegt werden, die beim Export als Attribut mitgeschrieben werden:

- jedes Suchwort wird eine eigene Zeile in `Attribute.xlsx`
- `TecDoc Kriterien ID` = `9595`, `Attribut Bezeichnung` = `Zusatzbezeichnung`, `Format` = `Alphanumerisch`
- maximal 20 Zeichen pro Suchwort (Vorgabe aus den Attribut-Stammdaten); laengere Eingaben werden abgewiesen
- mehrere Suchwoerter koennen mit `;` getrennt auf einmal eingegeben werden
- Duplikate werden automatisch ignoriert, auch gegenueber manuell gepflegten `9595`-Zeilen im Attribute-Tab

Beim erneuten Laden eines Artikels aus den Output-Dateien werden `9595`-Zeilen automatisch wieder dem Tab `Suchwoerter` zugeordnet (nicht dem Attribute-Tab).

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
