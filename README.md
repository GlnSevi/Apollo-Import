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
- Bilder, Dokumente, Videos und Web Links erfassen
  - mit Bild-Thumbnail und PDF-Vorschau in den Bilder- und Dokumente-Tabs
  - vorhandene Zeilen koennen direkt im Formular nachbearbeitet und aktualisiert werden
  - Aenderungen werden direkt in die Output-Excel-Dateien geschrieben
- Kurzbezeichnungen und Texte per DeepL aus dem deutschen Feld in `EN`, `CZ`, `FR`, `IT` und `NL` uebersetzen
- GenArt lokal und optional ueber inoffizielle Google-Lens-Treffer zum Produktbild verfeinern
  - mit direkter Suche ueber das GenArt-Feld oder den Button `Suchen...`
- Produktdaten direkt von `kunzer.de` laden
  - per Artikelnummer oder Produkt-URL
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

Wenn `Google Lens ohne API-Key fuer GenArt verwenden (inoffiziell)` aktiv ist, kann die GUI GenArt-Vorschlaege ueber sichtbare Treffer aus Google Lens verfeinern. Dabei werden Treffertexte und erkannte Seiten aus dem Lens-Ergebnis als zusaetzliches Signal fuer die GenArt-Suche genutzt.

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

Wenn `Kurzbezeichnung` oder `Text` aktiviert sind, wird die Uebersetzung im Listenimport automatisch ueber DeepL ausgefuehrt. Dafuer muss ein gueltiger API Key eingetragen sein.

Unterstuetzte Spalten sind flexibel. Mindestens eine der beiden Gruppen muss vorhanden sein:

- Artikelnummer: `Artikelnummer`, `Produktnummer`, `ArtNr`, `SKU`
- Produkt-URL: `Kunzer Produkt-URL`, `Kunzer URL`, `Produkt-URL`, `URL`, `Link`

CSV-Dateien duerfen mit `;`, `,`, Tab oder `|` getrennt sein. XLSX-Dateien werden aus dem ersten Tabellenblatt gelesen. Neue Artikel werden im festen Importpfad angehaengt; existiert die Artikelnummer bereits, werden nur deren alte Zeilen ersetzt.

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

## Google Lens

Die GenArt-Vorschlaege koennen optional ueber einen inoffiziellen Google-Lens-Abruf ohne API Key verfeinert werden.

- es ist kein zusaetzlicher API Key noetig
- die GUI nutzt dafuer einen Headless-Browser ueber `Playwright`
- ausgewertet werden sichtbare Treffertexte und erkannte Zielseiten aus dem Lens-Ergebnis
- der Weg ist bewusst experimentell und kann sich durch Aenderungen bei Google jederzeit veraendern

Die lokale Such- und Bildaehnlichkeitslogik bleibt als Fallback erhalten, falls Lens keine brauchbaren Treffer liefert.

## Exportierte Dateien

Die GUI erzeugt aktuell diese Dateien:

- `Kurzbezeichnung-NEU.xlsx`
- `Kurzbezeichnung_zu_ID.xlsx`
- `Text-NEU.xlsx`
- `Bilder.xlsx`
- `Dokumente.xlsx`
- `Videos.xlsx`
- `Web Link.xlsx`

Alle Output-Dateien enthalten zusaetzlich die Spalte `Zuletzt geschrieben am`. Darin steht pro Zeile, wann diese Daten zuletzt aus der GUI geschrieben wurden.

## Hinweis

Die GUI kann auf zwei Arten exportieren:

- mit Zeitstempel-Unterordner, damit alte Exporte erhalten bleiben
- direkt in einen festen Ausgabeordner, damit immer derselbe Importpfad genutzt werden kann

Im festen Ausgabeordner bleiben die sieben Exportdateien bestehen. Neue Artikel werden angehaengt, und bei erneutem Export derselben Artikelnummer werden nur die alten Zeilen dieses Artikels ersetzt.
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
