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
- Stueckliste (Lieferumfang mit Menge) und Zubehoerliste (optionales Zubehoer) pro Artikel pflegen wie in IMS, Reiter `Artikelliste(n)`
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
- Apollo-Importstand je Ausgabeordner: seit dem letzten Apollo-Import geaenderte Artikel werden im Verzeichnis markiert und koennen als Loeschliste (`757` / Artikelnummer, plus Blatt `Aenderungen`) gespeichert werden
- Reiter `Uebersicht`: alle Artikel wie in Excel auf einen Blick, je Reiter Anzahl und Datum der letzten Aenderung, sortierbar, durchsuchbar, als Excel speicherbar
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

Abrufergebnisse von kunzer.de werden innerhalb einer Programmsitzung zwischengespeichert (spart den zweiten Seitenaufruf, z. B. fuer die Attribut-Findung). Ein erneutes `Artikel anlegen` fuer denselben Artikel holt ab v0.1.45 trotzdem immer den aktuellen Shop-Stand (vorher kam bis zum Neustart das alte Ergebnis zurueck - ein im Shop nachtraeglich ergaenztes Bild wurde so nicht gefunden), und jeder Massenimport startet mit leerem Zwischenspeicher.

## Artikelverzeichnis

Die bisherige reine Text-Vorschau wurde durch ein `Artikelverzeichnis` ersetzt. Dort werden nur Artikel aus den aktuellen Output-Dateien tabellarisch gelistet.

- Suchfeld ueber der Tabelle: mehrere Begriffe mit Leerzeichen trennen, alle muessen vorkommen (Artikelnummer, Kurz-/Text-ID, GenArt, Kurzbezeichnungen aller Sprachen, deutscher Text, OE-/Vergleichsnummern samt Hersteller, Suchwoerter); Treffer-Zaehler, `Enter` waehlt einen eindeutigen Treffer, `Esc` oder `Leeren` setzt zurueck
- Artikel koennen per Button oder Doppelklick wieder in die Maske geladen werden; der Wechsel blockiert die Oberflaeche nicht mehr (Bild-, PDF- und Video-Vorschauen laden im Hintergrund, die Liste wird nur bei geaenderten Dateien neu aufgebaut)
- fuer jeden Artikel werden Quelle, IDs und Medienanzahlen angezeigt
- Spalte `Geaendert seit Import` (ab v0.1.52): Zeitpunkt der letzten Aenderung, wenn der Artikel seit dem Apollo-Importstand bearbeitet wurde; solche Zeilen sind hellorange hinterlegt (siehe `Apollo-Importstand und Loeschliste`). Die Markierung wird nach jedem Live-Speichern, Massenimport und `Apollo-Import erledigt` nachgezogen
- ueber der Tabelle steht die Zeile `Letzter Apollo-Import: ... - seitdem n Artikel geaendert, m geloescht` (gezaehlt wie die Loeschliste seit Apollo-Import) mit den Buttons `Apollo-Import erledigt`, `Importstand setzen...` und `Loeschliste seit Apollo-Import`
- rechts werden alle aktuell in den Export-Excel-Dateien vorhandenen Daten des markierten Artikels angezeigt
- Bilder zeigen ein Thumbnail, Dokumente bei PDFs die erste Seite und sonst eine kompakte Dateivorschau
- Web Links zeigen Host-Badge und Seitentitel; der Webseiten-Screenshot (Playwright) wird nur noch auf Knopfdruck `Webseiten-Vorschau laden` erzeugt und dann zwischengespeichert

## Apollo-Importstand und Loeschliste

Apollo (IMS) kann vorhandene Artikel beim Import nicht aktualisieren: Artikel, die seit dem letzten Apollo-Import in der GUI geaendert wurden, muessen vor dem naechsten Import ueber eine Loeschliste aus Apollo entfernt werden. Ab v0.1.52 merkt sich die GUI dafuer einen Importstand.

- Der Importstand liegt als `Apollo-Importstand.json` im Ausgabeordner (`last_import_at`, `set_at`, `set_by`, `note`, `changed_since`, `delete_list_at`, `delete_list_base`, `delete_list_size`, `delete_list_count`, `delete_list_checked_at`; `deleted_since` aus einem Vorabstand wird nur noch gelesen und bei `Apollo-Import erledigt` mit dem neuen Stand als Zeitpunkt ins Journal uebernommen) und gilt damit fuer alle Instanzen, die denselben Ordner nutzen. `changed_since` ist das Aenderungsjournal (Artikel -> Zeitpunkt der letzten Aenderung, siehe unten); es wird auch gefuehrt, solange noch kein Stand gesetzt ist - die Datei hat dann ein leeres `last_import_at`. Eine fehlende oder nur das Journal enthaltende Datei bedeutet `noch kein Importstand gesetzt` - dann wird nichts markiert. Ein UTF-8-BOM (Datei von Hand gespeichert) ist erlaubt. Ist die Datei nicht lesbar (kaputt, gesperrt), zeigt die Zeile `Apollo-Importstand nicht lesbar: <Fehler>` (bei kurzer Sperre bleibt der zuletzt gelesene Stand mit Markierung stehen, mit Hinweis); `Loeschliste seit Apollo-Import` und `Apollo-Import erledigt` lesen dann erneut und melden den Fehler, statt `noch kein Importstand gesetzt` - eine nicht lesbare Datei wird nie ueberschrieben.
- Massgeblich ist die Spalte `Zuletzt geschrieben am` der Output-Dateien: je Artikel und Bereich (Kurzbezeichnung, Text, GenArt, Attribute, Suchwoerter, OE-/Vergleichsnummern, Fahrzeuge, Listen, Bilder, Dokumente, Videos, Web Links) zaehlt der juengste Zeilen-Zeitstempel. Liegt er nach dem Importstand, gilt der Artikel als `geaendert`; sind alle Zeilen juenger als der Stand, als `neu`. Beim Live-Speichern bekommen die geaenderten Bereiche den Schreibzeitpunkt sofort im Verzeichnis, ohne die Dateien neu zu lesen.
- Unveraenderte Zeilen behalten ihren Zeitstempel - beim Live-Speichern, im Massenimport und beim Export in den festen Pfad. Ein Speichern ohne inhaltliche Aenderung (Fokus- oder Artikelwechsel, Massenimport mit unveraendertem Shop-Stand, Bild-Link, der wieder dieselbe lokale Datei ergibt, `Medien neu herunterladen` mit unveraendertem Dateiinhalt) markiert deshalb nichts und schreibt auch die `Apollo-Importstand.json` nicht. Ob ein Bereich der Maske geaendert ist, wird gegen das verglichen, was die Maske beim Laden (bzw. letzten Speichern) gezeigt hat - solange der Artikel in der Maske ist, auch wenn der Verzeichnis-Eintrag danach neu eingelesen wird, weil ein Katalog erst spaeter fertig geladen ist (z. B. KTyp-Stammdaten beim ersten Start nach einer Aenderung der `KTyp.xlsx`) oder eine zweite Instanz den Artikel geaendert hat. Laden, Feld- oder Reiterwechsel, Artikelwechsel und Speichern ohne Eingabe schreiben nichts; erst eine echte Eingabe schreibt den Bereich (dann mit dem Modell der GUI, bei Fahrzeugen z. B. je KTyp-Nummer und GenArt eine Zeile). Hat eine zweite Instanz den geladenen Artikel geaendert, zeigt die Maske bis zum erneuten Laden weiter den alten Stand: eine Eingabe im selben Bereich ueberschreibt die Fremdaenderung dort, die anderen Bereiche bleiben unberuehrt. Hat ein vorhandener Artikel keine Text- bzw. Kurzbezeichnungs-Zeile und in diesem Bereich auch keinen Text, wird beim Speichern keine leere Zeile mit neu erzeugter Modul-ID mehr angelegt (ausser der Bereich ist zuvor beim Speichern gescheitert - dann wird er immer nachgeholt).
- `Apollo-Import erledigt`: nach dem Import in Apollo klicken (mit Rueckfrage). Der neue Stand ist der Zeitpunkt der ERSTEN `Loeschliste seit Apollo-Import`, die seit dem bisherigen Stand gegen ihn oder einen frueheren Stand gespeichert wurde (`delete_list_at` mit `delete_list_base` <= Stand < `delete_list_at`) - sonst jetzt; die Rueckfrage nennt den Zeitpunkt mit Sekunden, die Artikelzahl der Liste und den Grund. Gab es seit dem Stand nur die Pruefung `Eine Loeschliste ist nicht noetig` (keine Datei), gilt deren Zeitpunkt (die Rueckfrage sagt das). Folgte auf eine solche Pruefung (nach dem Stand) noch eine echte Liste, weiss nur der Nutzer, was in Apollo ging: `erledigt` fragt dann ausdruecklich `Wurde die Loeschliste vom <Zeitpunkt mit Sekunden> (<n> Artikel) in Apollo importiert?` (Warnsymbol, Vorauswahl `Nein`). `Ja` = Stand auf die Liste; `Nein` = der Apollo-Import lief schon nach der Pruefung (ohne Liste), Stand auf die Pruefung - die Artikel der Liste bleiben markiert und kommen noch einmal auf die naechste Liste (sicher); die Liste bleibt dabei vermerkt: das naechste `Apollo-Import erledigt` nimmt wieder ihren Zeitpunkt (sicher - was danach geaendert wurde, bleibt markiert); `Abbrechen` aendert nichts (z. B. wenn der Apollo-Import noch gar nicht gelaufen ist). Die Ueberschrift dieser Rueckfrage ist die Frage selbst (`Wurde die Loeschliste vom ... in Apollo importiert?`). Ohne diese Frage gaelte bei einer nur zum Ansehen gespeicherten Liste alles, was zwischen dem Apollo-Import und der Liste geaendert wurde, still als importiert. Eine Pruefung, die nicht nach dem Stand liegt, loest die Frage nicht aus. Wurden seit dem Stand mehrere echte Listen gespeichert (`delete_list_count`, z. B. eine zweite nur zum Ansehen oder aus einer anderen Instanz), zaehlt die erste; die Rueckfrage sagt das - wurde in Apollo eine spaetere Liste importiert, danach `Importstand setzen...` auf deren Zeitpunkt stellen (sonst stehen die Artikel dazwischen nur noch einmal auf der naechsten Liste, doppelt loeschen und neu importieren schadet nicht). Wurde der Stand nach dem Speichern der Liste mit `Importstand setzen...` auf einen FRUEHEREN Zeitpunkt gestellt als den, gegen den die Liste erstellt wurde (`delete_list_base`), fehlen ihr die Artikel dazwischen: dann gilt jetzt, und die Rueckfrage warnt (Vorauswahl `Nein`), die Liste neu zu erstellen (unterscheiden sich beide Zeitpunkte nur in den Sekunden, zeigt sie sie mit Sekunden). Das Journal bleibt vollstaendig erhalten: Eintraege bis zum neuen Stand zaehlen nur nicht mehr - ein spaeter mit `Importstand setzen...` frueher gesetzter Stand (z. B. nach einem versehentlichen `erledigt`) erfasst sie wieder, auch Loeschungen und komplett geleerte Bereiche. Juengere Eintraege bleiben markiert: was nach dem Speichern der Loeschliste geaendert wurde (z. B. waehrend TecAlliance importiert, auch in einer anderen Instanz oder waehrend der Rueckfrage), stand auf keiner Liste und bleibt markiert bis zur naechsten. Offene Live-Auftraege werden vorher geschrieben, ebenso noch nicht vermerkte Artikel frueherer Massenimporte (scheitert das wieder, werden sie spaeter mit einem Zeitpunkt nach dem neuen Stand vermerkt und bleiben markiert).
- `Importstand setzen...`: Datum und Uhrzeit frei eintragen (`TT.MM.JJJJ HH:MM`, Sekunden optional `HH:MM:SS`, vorbelegt mit jetzt, optional eine Notiz), z. B. wenn der Import schon gestern lief. Das Journal bleibt erhalten und zaehlt mit allen Eintraegen nach dem neuen Stand - auch mit Aenderungen und Loeschungen aus der Zeit, bevor ueberhaupt ein Stand gesetzt war (erste Nutzung: Stand nachtraeglich auf den echten letzten Import setzen) oder bevor ein frueheres `erledigt` den Stand weitergesetzt hat. Der Vermerk der Loeschliste bleibt stehen; ob er zum neuen Stand passt, pruefen `erledigt` und die naechste Liste nach derselben Regel (`delete_list_base` <= Stand < `delete_list_at`) - passt er nicht mehr, ersetzt ihn die naechste Liste. Nur der Zeitpunkt einer Pruefung ohne Liste (`delete_list_checked_at`) entfaellt, wenn der neue Stand nicht davor liegt (die Pruefung ist damit erledigt, `erledigt` fragt nicht mehr nach ihr). Die Statuszeile zeigt den Stand mit Sekunden, wenn sie eingegeben wurden.
- Waehrend eines Massenimports, einer IMS-Uebernahme oder eines anderen Hintergrundvorgangs sind die drei Buttons gesperrt bzw. melden nur `Es laeuft bereits ein Hintergrundvorgang` - ein Stand mitten im Lauf liesse die bis dahin geschriebenen Artikel dauerhaft als importiert gelten.
- `Loeschliste seit Apollo-Import`: speichert alle seit dem Stand geaenderten, neuen und geloeschten Artikel als XLSX (oder CSV). Blatt `Loeschliste` hat den Aufbau der Apollo-Loeschlisten (Spalte A = `757`, Spalte B = Artikelnummer), Blatt `Aenderungen` zeigt je Artikel `letzte Aenderung`, `geaenderte Bereiche` und `Status` (`geaendert`, `neu`, `geloescht`). Der Zeitpunkt der Liste (nach dem Schreiben aller offenen Live-Auftraege, vor dem Einlesen) wird als `delete_list_at` im Importstand vermerkt, dazu der Stand, gegen den sie erstellt wurde (`delete_list_base`), und ihre Artikelzahl (`delete_list_size`); Bestaetigung und Statuszeile nennen ihn mit Sekunden (`Loeschliste vom 29.10.2026 08:54:50 ...`), so laesst er sich bei Bedarf mit `Importstand setzen...` eintragen. Gibt es seit dem Stand keine Aenderung, kommt nur der Hinweis `Eine Loeschliste ist nicht noetig (geprueft am ...)`; diese Pruefung wird mit `delete_list_size` 0 vermerkt, zaehlt aber nicht als Liste: die erste echte Liste danach ersetzt sie (der Zeitpunkt der Pruefung bleibt als `delete_list_checked_at` fuer die Rueckfrage in `erledigt`), weitere Pruefungen aendern nichts. Eine weitere echte Liste, solange der Vermerk zum Stand passt, aendert ihn nicht, sie zaehlt nur `delete_list_count` hoch (Pruefungen ohne Liste zaehlen nicht). Vermerke ohne `delete_list_size` (aus einem Entwicklungsstand) gelten als echte Liste. Laesst sich der Zeitpunkt nicht vermerken, warnt eine Meldung und nennt, was `erledigt` dann nimmt (eine frueher vermerkte Liste seit dem Stand, eine Pruefung ohne Liste - dann bleiben die Artikel der neuen Liste markiert - oder die Klickzeit). Laufen noch Live-Auftraege, wird keine Liste erstellt. Im Massenimport vorgemerkte Artikel, deren Vermerk im Journal gescheitert war (z. B. ein dort geleerter Bereich, der nur im Journal steht), werden vorher vermerkt; scheitert das erneut, wird keine (unvollstaendige) Liste erstellt, die Meldung nennt die Artikel. Die bisherige `Apollo-Loeschliste speichern` im Massenimport (nur der letzte Lauf) bleibt erhalten.
- Im Verzeichnis geloeschte Artikel (`Zeilen loeschen`) stehen in keiner Output-Datei mehr; sie werden mit dem Loeschzeitpunkt im Journal vermerkt und landen so trotzdem auf der Loeschliste (Status `geloescht`) - auch wenn derselbe Artikel spaeter neu angelegt wird, bis zum naechsten `Apollo-Import erledigt`. Scheitert das Loeschen mittendrin (z. B. `Bilder.xlsx` in Excel geoeffnet), werden die markierten Artikel trotzdem vermerkt - ein Teil ihrer Zeilen ist dann schon entfernt.
- Aenderungsjournal (`changed_since`): Jeder Schreibvorgang mit echter Aenderung vermerkt den Artikel mit dem Zeitpunkt seiner LETZTEN Aenderung (ein spaeterer Vermerk ersetzt einen frueheren) - Live-Speichern, Export in den festen Pfad (auch `Exportieren & Fertigstellen` im gefuehrten Modus; Export in einen Zeitstempel-Unterordner nicht), Massenimport (jeder Artikel mit geaendertem Bereich, auch wenn der Artikel spaeter im Lauf scheitert oder der Lauf abgebrochen wird; geaenderte Bereiche aus einem gescheiterten 1. Versuch zaehlen mit, ebenso nur umsortierte Attribute einer Attribut-Vorlage; vermerkt wird gesammelt waehrend des Laufs alle 30 Sekunden bzw. 25 Artikel, am Laufende und beim Schliessen des Programms mitten im Lauf - auch waehrend auf den KI-Batch gewartet wird -, jeweils mit dem Zeitpunkt des Vermerks; scheitert der Vermerk, bleiben die Artikel vorgemerkt und werden beim naechsten Vermerk nachgeholt - auch erst im naechsten Lauf, vor `Loeschliste seit Apollo-Import` bzw. `Apollo-Import erledigt` oder beim Schliessen - und immer im Ausgabeordner ihres Laufs, nie in dem eines spaeteren Laufs in einen anderen Ordner), Suchwort-Import (nur Artikel, deren Suchwoerter sich wirklich aendern) und Loeschen im Verzeichnis. Die IMS-Uebernahme vermerkt nichts (ihre Zeilen sind der Stand, den Apollo schon hat). Das Journal ist noetig, weil die Zeilen-Zeitstempel nicht jeden Fall abbilden: wird ein Bereich komplett geleert oder ein Artikel geloescht, bleibt in den Dateien kein juengerer Zeitstempel. Markierung, Zaehler und Loeschliste beruecksichtigen das Journal - auch nach einem Neuladen von der Platte oder einem Neustart. Das Journal wird nicht gekuerzt (ein Eintrag je Artikel mit dem juengsten Zeitpunkt); Eintraege bis zum Stand zaehlen einfach nicht.
- Entfernte oder umsortierte Zeilen: Aendern sich die Zeilen eines Bereichs, ohne dass eine neue Zeile dazukommt (z. B. ein Attribut, eine OE-Nummer, ein Bild oder ein Suchwort entfernt, Bilder bzw. Dokumente nur in anderer Reihenfolge), bekommen alle Zeilen des Artikels in diesem Bereich den Schreibzeitpunkt. Laedt `Medien neu herunterladen` (Massenimport) unter derselben URL eine Datei mit anderem Inhalt, wird sie unter demselben Pfad ersetzt und der Bereich gilt ebenfalls als geaendert (Zeilen gestempelt, Journal) - auch wenn danach im selben Versuch etwas scheitert (der 2. Versuch findet gleiche Bytes vor und stempelt die Zeilen trotzdem; scheitert der Artikel ganz, steht er im Journal); bei gleichem Inhalt bleibt die Datei unberuehrt. Nativ gespeicherte Dateien (JPG, PNG, GIF, PDF) werden Byte fuer Byte verglichen - so zaehlt auch eine nur geaenderte EXIF-Ausrichtung, ein anderes ICC-Profil, eine andere Transparenz oder ein geaenderter spaeterer GIF-Frame. Nur nach PNG umgewandelte Bilder (webp/tif/bmp) gelten bei anderen Bytes, aber gleichen Bilddaten (jeder Frame mit Groesse, Modus, Palette, Transparenz und Pixeln, dazu das ICC-Profil) als gleich - die Umwandlung kann mit einer anderen Pillow-Version andere Bytes ergeben. Die Zeitstempel der `Kurzbezeichnung_zu_ID.xlsx` zaehlen zum Bereich Kurzbezeichnung.
- Zeilen ohne `Zuletzt geschrieben am` (z. B. von Hand in Excel eingefuegt) zaehlen in allen Output-Dateien mit der Aenderungszeit der Datei - wie beim Schreiben, wo diese Zeit dann auch in die Zelle uebernommen wird.
- Grenzen: Nach einer IMS-Uebernahme (`IMS-Export uebernehmen...`) tragen die ergaenzten Zeilen das Uebernahmedatum - danach den Importstand neu setzen, sonst stehen diese Artikel auf der naechsten Loeschliste. Wegen des Rueckfalls auf die Aenderungszeit der Datei gelten Zeilen ohne Zeitstempel als geaendert, sobald die Datei nach dem Stand gespeichert wurde (auch von Hand in Excel). Das Journal wird von mehreren Instanzen ohne Sperre geschrieben; schreiben zwei Instanzen im selben Moment, kann ein Vermerk verloren gehen. Zeitstempel haben Sekundenaufloesung: was eine andere Instanz in derselben Sekunde schreibt, in der die Loeschliste erstellt wird, kann nach `erledigt` als importiert gelten. Als erste Liste zaehlt die zuerst im Importstand vermerkte: der Vermerk folgt erst nach dem Speichern-Dialog, speichert eine andere Instanz in dieser Zeit eine Liste, gilt deren (spaeterer) Zeitpunkt. Ein im Massenimport vorgemerkter Artikel, dessen Vermerk auch beim Schliessen des Programms noch scheitert, geht verloren (die Meldung vom Laufende nennt ihn). Ist die `Apollo-Importstand.json` kurz gesperrt oder nicht lesbar (z. B. liest eine zweite Instanz sie gerade, ein Virenscanner prueft sie oder das Netzlaufwerk setzt aus), werden Lesen und Schreiben einige Male wiederholt; scheitert ein Vermerk trotzdem, erscheint eine Meldung mit den Artikeln, die dann von Hand auf die Apollo-Loeschliste muessen (je Artikel einmal, bis ein Vermerk wieder gelingt oder `erledigt`/`Importstand setzen` folgt; Loeschungen immer), und die Statuszeile behaelt den Hinweis `Importstand nicht vermerkt`. `erledigt`, `Importstand setzen` und `Loeschliste seit Apollo-Import` aendern eine nicht lesbare Datei nicht (Meldung `konnte nicht gelesen werden`).

## Uebersicht

Der Reiter `Uebersicht` (direkt hinter `Projekt`, ab v0.1.52) zeigt alle Artikel des Ausgabeordners wie in Excel auf einen Blick - eine Zeile je Artikel, dazu je Reiter, was gepflegt ist.

- Spalten: Artikelnummer, Kurzbezeichnung (DE, gekuerzt), Marke (Attribut 8988), GenArten (IDs), Letzte Aenderung, Seit Import (`ja`, wenn seit dem Apollo-Importstand geaendert), danach je Bereich `Anzahl (TT.MM.)` = Zeilen in der Output-Datei und Datum der letzten Aenderung (Kurzbezeichnung und Text als `Text-Modul-ID · n Spr. (TT.MM.)`), `-` wenn leer
- Suchfeld wie im Artikelverzeichnis (alle Begriffe muessen vorkommen; zusaetzlich zaehlen die Zellen der Zeile, also auch Marke und Datum), Treffer-Zaehler `n Artikel (m geaendert seit Import)`
- Klick auf einen Spaltenkopf sortiert, ein zweiter Klick dreht die Richtung um (Anzahlen und Daten werden als Zahl bzw. Datum sortiert, nicht als Text)
- Doppelklick laedt den Artikel wie aus dem Verzeichnis in die Maske; geaenderte Artikel sind wie im Verzeichnis hellorange hinterlegt
- `Als Excel speichern...` schreibt die Tabelle (aktuelle Suche und Sortierung) als `Uebersicht_<Datum>.xlsx`, Blatt `Uebersicht`
- Die Zeilen werden je Artikel nur neu gebaut, wenn sich sein Eintrag oder der Importstand geaendert hat; gezeichnet wird nur, solange der Reiter sichtbar ist

## IMS-Export uebernehmen (TAF 2.7)

Ueber `IMS-Export uebernehmen...` im Artikelverzeichnis (ab v0.1.47) liest die App einen kompletten Apollo-IMS-Export (TAF 2.7: Exportordner mit Unterordner `0757`, `0757_PIC.zip` und `0757_DOC.zip`) und ergaenzt damit die Output-Dateien. Ziel: der Datenstamm der GUI ist danach mindestens so vollstaendig wie Apollo.

- Regel `GUI gewinnt je Artikel und Datei`: Zeilen, die in der GUI schon vorhanden sind, bleiben unveraendert; nur in der GUI fehlende Artikel bzw. Bereiche kommen aus dem Export dazu (Kurzbezeichnungen, Texte, GenArten, Attribute, Suchwoerter, OE-/Vergleichsnummern, Fahrzeugverknuepfungen, Stueck-/Zubehoerlisten, Bilder, Dokumente, Videos, Web Links). Abweichungen zwischen GUI und IMS werden nicht angefasst, sondern im Bericht aufgelistet. Einzige Aenderung an vorhandenen GUI-Zeilen: leere Text-Modul-IDs werden gefuellt (im Bericht unter Hinweise).
- Vorher werden alle Excel-Dateien des Ausgabeordners als `<Name>_Sicherung_<Zeitstempel>.xlsx` in den Sicherungsordner neben dem Ausgabeordner kopiert (`G:\Apollo\GUI` -> `G:\Apollo\GUI-Sicherungen`); dort entstehen auch der Bericht `TAF-Sync-Bericht_<Zeitstempel>.xlsx` (Uebersicht, Konflikte mit beiden Staenden, Hinweise; bei einem Abbruch der bis dahin gesammelte Teilbericht) und das Protokoll `TAF-Sync-Protokoll_<Zeitstempel>.txt`. Ist eine Output-Datei in Excel geoeffnet, startet die Uebernahme nicht.
- Benutzt werden die in den Einstellungen hinterlegten Kataloge (GenArten, Attribute, KHer, KTyp). Fehlt einer, startet die Uebernahme nicht - ohne `KHer.csv` wuerden z. B. alle Referenznummern zu OE-Nummern, und das liesse sich wegen der Regel mit einem zweiten Lauf nicht mehr reparieren.
- Medien aus den Zips werden nur fuer Artikel kopiert, die in der GUI noch keine Bilder bzw. Dokumente haben (Ablage wie beim Kunzer-Abruf unter `Medien\<Artikel>`); das Herstellerlogo landet unter `Medien\_logo` (mit `--no-media` entfaellt auch das Logo).
- Hersteller aus Tabelle 203 werden ueber `KHer.csv` (Spalte VGL) als Vergleichs- oder OE-Nummern eingeordnet; unbekannte Herstellernummern kommen als OE-Nummern mit Hinweis im Bericht. EAN (209), Status (212) und GUID (299) werden nur gezaehlt.
- Waehrend der Uebernahme laeuft die Oberflaeche weiter (Schreibsperre wie beim Massenimport); das Laden aus dem Verzeichnis ist solange gesperrt und Schliessen fragt nach. Danach werden Artikelverzeichnis und Maske aus den Dateien neu geladen; die Maske muss dafuer vorher gespeichert sein (bei nicht gespeicherten Bereichen, z. B. einer OE-Zeile ohne Hersteller, startet die Uebernahme nicht).
- Ohne Oberflaeche: `python apollo_import_gui.py --ims-sync <Exportordner> <Ausgabeordner> [--no-media]` (auch mit der EXE; die Zusammenfassung steht dann nur im Protokoll, ein Fehler in `TAF-Sync-Fehler_<Zeitstempel>.txt` im Sicherungsordner). Die Kommandozeile nimmt die Standard-Katalogpfade und hat keine Sperre gegen eine gleichzeitig laufende GUI - die GUI vorher schliessen. Das Modul `taf_sync.py` laesst sich auch direkt aufrufen (`python taf_sync.py <Exportordner> <Ausgabeordner> --report <xlsx>`), dann ohne Sicherung und Katalogpruefung.

## Leistung, Live-Speichern und Caches

- Live-Speichern laeuft in einem Hintergrund-Thread mit Warteschlange: Aenderungen in der Maske blockieren die Oberflaeche nicht mehr (Statuszeile zeigt `Speichere im Hintergrund ...` und danach `Live gespeichert: <Artikel>`). Geschrieben wird nur, was sich seit dem Laden bzw. letzten Speichern wirklich geaendert hat (Fingerabdruck je Artikel und Bereich); ein Fokusverlust ohne Aenderung schreibt nichts. Eingereihte Auftraege verwirft der Hintergrund-Schreiber nur, solange ein Massenimport die Output-Dateien besitzt (eigenes Flag, gesetzt vom Start bis zum Ende des Laufs) - nie, weil die Maske gerade gefuellt wird. Bis v0.1.52 pruefte er dafuer die Sperre, die auch jedes Laden und jeder Artikelwechsel setzt: eine Eingabe, deren Auftrag genau waehrend eines Wechsels begann (z. B. Suchwort in 77M505, sofort 77M810 getippt), ging still verloren (15 von 20 Versuchen); ebenso konnte der Folgeauftrag ausbleiben, der nach einer GenArt-Aenderung die Fahrzeugzeilen nachzieht. Trifft dieser Folgeauftrag auf einen schon wartenden Auftrag desselben Artikels, wird er mit ihm zusammengefasst und behaelt dessen juengeren Stand (bisher ersetzte er ihn durch den Stand des GenArt-Auftrags, und z. B. ein direkt danach getipptes Suchwort wurde nicht geschrieben); zwei wartende Folgeauftraege behalten den juengeren. Wurde nach dem GenArt-Auftrag schon ein Auftrag mit Fahrzeugzeilen fuer den Artikel eingereicht (er traegt die neue GenArt und die aktuellen Fahrzeuge), entfaellt der Folgeauftrag ganz (Reihenfolge der Einreichung je Auftrag): bis v0.1.52 lief er, wenn dieser Auftrag beim Abholen des GenArt-Ergebnisses schon begonnen hatte, danach und schrieb die alte Fahrzeugliste zurueck - ein gerade entferntes Fahrzeug kam wieder, auch das Schliessen reparierte das nicht.
- Scheitert ein Bereich (z. B. Datei in Excel geoeffnet), werden die uebrigen Bereiche trotzdem geschrieben; der gescheiterte wird beim naechsten Anlass nachgeholt (mit dem Maskenstand) und beim Beenden gibt es eine Warnung, falls noch etwas offen ist. Export, Loeschen und Massenimport warten vorher auf die Warteschlange. Ab v0.1.52 geht die Eingabe dabei auch beim Artikelwechsel nicht mehr verloren:
  - Verlassen eines Artikels mit gescheitertem Bereich (getippte Nummer, Nummer geleert, Doppelklick im Verzeichnis oder in der Uebersicht): der Nachholer wird geschrieben und (begrenzt, max. 30 s) abgewartet; scheitert er wieder, wird der Wechsel abgebrochen wie bei einem fachlichen Fehler, und die Meldung nennt die Datei (`OE-Nummern.xlsx konnte nicht geschrieben werden (z. B. in Excel geoeffnet) - bitte die Datei schliessen und erneut versuchen`). Bisher galt der eingereihte Nachholer als gesichert, scheiterte nach dem Wechsel erneut, und zurueck im Artikel schrieb der naechste Nachholer den Dateistand - die Eingabe war still weg, die Warnung beim Beenden ebenfalls.
  - Scheitert ein Auftrag erst, nachdem der Artikel verlassen wurde, merkt sich die App seine Eingabe (Bundle des juengsten gescheiterten Auftrags samt Ausgabeordner). Wird der Artikel wieder geladen (getippt, Doppelklick, Uebersicht, leere Maske + Nummer), schreibt sie diese Eingabe zuerst erneut (begrenzt gewartet): gelingt es, liest sie den Artikel aus den Dateien neu und laedt diesen Stand; scheitert es wieder, laedt sie den Dateistand und setzt die gescheiterten Bereiche aus der gemerkten Eingabe in die Maske (Statuszeile: `... noch nicht in den Dateien (<Datei> gesperrt?): Eingabe wiederhergestellt ...`). Der Marker bleibt, der naechste Speichervorgang schreibt also die Eingabe - nie den Dateistand. Ein noch nie geschriebener Artikel wird dabei aus der gemerkten Eingabe geladen statt als neu geleert.
  - Von der gemerkten Eingabe gelten nur ihre gescheiterten Bereiche: Der Nachholer uebernimmt sein (altes) Bundle nie in den Verzeichnis-Eintrag, sondern das Verzeichnis wird nach dem Schreiben neu gelesen (Zeilen-Cache: nur die geschriebenen Dateien werden neu geparst; mit dem Echtdatenstand von 3.800 Artikeln rund 1 s, nur in diesem seltenen Fall) und der gemerkte Stand des Artikels daraus abgeleitet. Was eine zweite Instanz oder der eigene Massenimport inzwischen in anderen Bereichen geaendert hat (z. B. Web-Link, Langtext, GenArt), bleibt so in Maske, Verzeichnis und Dateien. Uebernaehme der Nachholer sein Bundle (bei mehreren Bereichen den ganzen Eintrag, sonst GenArt und Modul-IDs), verschwaenden die Fremdzeilen aus der Maske, gaelten als gespeichert, und die naechste eigene Eingabe loeschte sie aus den Dateien. Zieht der Nachholer nach einer gemerkten GenArt die Fahrzeugzeilen nach, nimmt er die Fahrzeuge aus den Dateien. Laeuft der Nachholer nach der Wartezeit noch, zeigt die Maske den Dateistand und die gescheiterten Bereiche aus der gemerkten Eingabe (nicht das ganze alte Bundle). Die gescheiterten Bereiche selbst ueberschreiben eine neuere Fremdaenderung im selben Bereich (Maske gewinnt je Bereich).
  - Beim Beenden werden auch die gemerkten Eingaben verlassener Artikel noch einmal geschrieben; was weiter scheitert, nennt die bisherige Warnung (`Trotzdem beenden?`).
  - Waehrend eines Massenimports (auch beim Warten auf den KI-Batch, bis zu 2 h) wird nicht gespeichert. Wird dann ein Artikel mit ungespeicherten oder gescheiterten Bereichen verlassen (getippte Nummer, Nummer geleert, Doppelklick im Verzeichnis oder in der Uebersicht), fragt die App `Waehrend des Massenimports wird nicht gespeichert. Eingaben in <A> (<Bereiche>) verwerfen und <B> laden?` (Warnsymbol, Vorauswahl `Nein`; Bereiche, die nur gemerkt werden, nennt die Frage als `merken`, nicht als `verwerfen`). Solange ein Wechsel bzw. das Laufende des Massenimports auf das Schreiben wartet, nimmt die Maske keinen weiteren Wechsel an (Statuszeile `Bitte warten - das Speichern des bisherigen Artikels laeuft noch.`). `Nein` bleibt beim Artikel; `Ja` laedt den anderen Artikel: die Eingabe in Bereichen, deren Schreiben schon vor dem Lauf gescheitert war, wird als gemerkte Eingabe geparkt (nachgeholt bzw. wiederhergestellt beim naechsten Laden oder beim Schliessen, verworfen, wenn der Lauf den Bereich neu schreibt), alle anderen Eingaben werden wie bestaetigt verworfen (Statuszeile nennt beides). `Bitte zuerst korrigieren` waere waehrend des Laufs nicht moeglich (bis v0.1.52 lud der Doppelklick ohne Pruefung, die Eingaben waren still verloren). Der Hinweis `<Datei> konnte nicht geschrieben werden` erscheint nur, wenn beim Verlassen wirklich ein Schreibversuch erneut scheiterte - ein Marker allein (z. B. waehrend des Laufs, der nichts schreibt) sagt nichts ueber die Datei.
  - Am Ende eines Massenimports werden ungespeicherte Eingaben in der Maske zuerst geschrieben, wenn der Lauf diesen Artikel nicht neu geschrieben hat (bisher lud der Abschluss den letzten Artikel des Laufs und verwarf sie still); laesst sich eine Eingabe auch dann nicht speichern (fachlicher Fehler), bleibt die Maske bei diesem Artikel. Hat der Lauf den Artikel neu geschrieben, gewinnt wie bisher der Lauf; die Statuszeile nennt die verworfenen Eingaben.
  - Die gemerkte Eingabe entfaellt mit dem Marker: sobald die Bereiche geschrieben sind (Live-Speichern, Export in den festen Pfad), ein Massenimport die Bereiche neu geschrieben hat (nur ergaenzte Marken-Attribute zaehlen nicht), der Suchwort-Import die Suchwoerter geschrieben hat, der Artikel im Verzeichnis geloescht wird oder eine neue, leere Maske unter der Nummer angelegt wird. Gemerkt wird nur fuer den aktuellen Ausgabeordner; nach einem Wechsel des Ausgabeordners wird nichts nachgeholt oder wiederhergestellt.
- Massgeblich sind immer die Output-Dateien: Laden aus dem Artikelverzeichnis liest geaenderte Dateien vorher neu ein (ab v0.1.41), und beim Programmstart fuellt die App die Maske eines gespeicherten Artikels aus den Dateien statt aus der Sitzungsdatei (ab v0.1.43). Vorher konnte der beim letzten Schliessen gemerkte Maskenstand nach einer Aenderung von aussen (Massenimport, Bereinigung, zweite Instanz) beim naechsten Fokuswechsel zurueckgeschrieben werden. Nur ein noch nicht gespeicherter Entwurf kommt weiter aus der Sitzungsdatei.
- Artikelwechsel in der Maske (ab v0.1.48): Wird die Artikelnummer in einer gefuellten Maske auf eine andere Nummer geaendert, gilt das als Wechsel, nicht als Umbenennen. Offene Aenderungen werden noch unter der alten Nummer gespeichert; danach wird die neue Nummer aus dem Verzeichnis geladen, wenn es sie gibt, sonst die Maske fuer den neuen Artikel geleert (Statuszeile sagt, was passiert ist). Vorher wurde der komplette Maskeninhalt unter der neuen Nummer gespeichert und blieb beim folgenden `Artikel anlegen` als gepflegter Stand stehen - so bekam z. B. 800WSG01.1.6 alle Attribute von 800WSG01.1.5; eine vorhandene Nummer wurde dabei mit dem Maskeninhalt ueberschrieben. Ohne geladenen Artikel (frischer Start, leere Maske): eine vorhandene Nummer wird in die leere Maske geladen; steht ein Entwurf in der Maske, fragt die App vor dem Ueberschreiben nach (`Nein` ueberschreibt den Artikel vollstaendig mit dem Entwurf, verglichen gegen den aktuellen Dateistand - auch wenn der Artikel frueher in der Sitzung geladen und seitdem von aussen oder vom eigenen Massenimport geaendert wurde; bis v0.1.52 blieben Bereiche, die der Entwurf wie beim frueheren Laden zeigte, mit den neuen Fremdzeilen stehen); eine neue Nummer wird erst geschrieben, sobald etwas eingegeben ist (erst Daten, dann Nummer tippen bleibt moeglich). Wird die Nummer geleert, gilt der Artikel als verlassen (offene Aenderungen werden noch gesichert, die Maske geleert). Kann ein Wechsel nicht sichern (z. B. OE-Zeile ohne Hersteller, Web-Link ohne `http(s)://`, gesperrte Datei), bleibt die Maske beim alten Artikel und der naechste Klick auf `Artikel anlegen` zeigt den Grund (waehrend eines Massenimports fragt die App stattdessen, siehe oben). Dieselbe Pruefung gilt ab v0.1.52 auch fuer den Doppelklick im Artikelverzeichnis bzw. in der Uebersicht und fuer `Ausgewaehlten Artikel laden`: zeigt die Maske einen anderen Artikel, werden dessen offene Aenderungen erst gesichert; was sich nicht sichern laesst, bricht das Laden mit Meldung ab (bis dahin lud der Doppelklick ohne Pruefung, und die Eingabe war still verloren). Derselbe Artikel laesst sich weiter per Doppelklick neu laden (das zeigt den Dateistand: eine nicht speicherbare Eingabe wie ein Web-Link ohne Schema entfaellt dabei, eine wegen gesperrter Datei gescheiterte Eingabe wird wiederhergestellt). Der beim Laden gemerkte Stand (siehe `Apollo-Importstand und Loeschliste`) gilt nur, solange der Artikel in der Maske ist: fuer einen verlassenen Artikel, dessen Verzeichnis-Eintrag danach ersetzt wird, zaehlt wieder der Dateistand.
- Excel-Dateien werden immer erst als Temp-Datei geschrieben und dann eingetauscht - eine gesperrte Zieldatei bleibt vollstaendig erhalten.
- Lesen und Schreiben der Output-Dateien nutzt einen eigenen XLSX-Leser/-Schreiber (Inline-Strings, gleiches Dateiformat wie bisher) plus Zeilen-Cache im Speicher; dadurch dauert z. B. das Schreiben der `Fahrzeugverknuepfungen.xlsx` (16.000 Zeilen) rund 0,1 s statt 2 s.
- Stammdaten (KTyp, Attribute, Schluesselwerte, GenArten) und die geparsten Output-Dateien werden unter `%LOCALAPPDATA%\Apollo Import GUI\cache` zwischengespeichert; der Cache erkennt geaenderte Quelldateien und neue Programmversionen selbst. Der Ordner darf jederzeit geloescht werden (naechster Start parst dann einmal neu, KTyp ca. 10 s im Hintergrund).
- Schwere Pakete (Playwright, PyMuPDF, Anthropic) werden erst beim ersten Gebrauch geladen; der Start ist dadurch deutlich schneller.
- PDF-Vorschau (ein Hintergrund-Thread je Vorschau) und PDF-Text fuer die Attribut-Findung nutzen PyMuPDF, das nicht threadsicher ist: jede Verwendung (Oeffnen, Rendern bzw. Text lesen, Schliessen) laeuft unter einer gemeinsamen Sperre (`PYMUPDF_LOCK`), Download und Oberflaeche nicht. Bis v0.1.52 konnte schnelles Durchklicken von Artikeln mit PDF-Dokumenten das Programm abstuerzen lassen (`access violation` beim Rendern, im Test nach einigen hundert Artikeln).
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
- **KI auch bei kurzem Text** (Haken im Abrufumfang, standardmaessig an): Ohne diese Option fragt die Attribut-Findung die KI erst ab 200 Zeichen Text (oder bei unerkannten `Label: Wert`-Zeilen). Viele Shop-Texte sind aber knapp - z. B. MotoMeter `KPS fuer Dieselmotor 10-40 bar mit Druckschlauch, Schnellkupplung` (64 Zeichen): ohne Option kein einziges Attribut, mit Option `Druck von/bis 10/40 bar` plus Zusatzinfo. Kostet ein paar Cent je Artikel; gilt fuer Einzel- und Massenimport.
- **Internet-Recherche auch im Massenimport** (Haken, standardmaessig aus): wie der Button `Automatisch ermitteln (Internet)` - die KI recherchiert je Artikel auf Herstellerseiten und in Datenblaettern (ca. 10-30 ct pro Artikel mit Opus, deutlich laengere Laufzeit; laeuft je Artikel direkt statt ueber die Batch API). Vor dem Start nennt eine Rueckfrage die geschaetzten Kosten; `Nein` startet den Lauf ohne Internet-Recherche.
- **Attribute nach Referenzartikel** (Haken + Feld `Referenzartikel`, standardmaessig aus): Fuer Listen gleichartiger Produkte. Alle Artikel des Massenimports bekommen genau die Attribute des Referenzartikels (ausser Marke) - z. B. mit `0012,8-706R 10` fuer stufenlose Oetiker-1-Ohr-Klemmen: Klemmschellenausfuehrung 3507, Spannbereich 727 (von-bis), Anzahl Teile 9339. Seit v0.1.50 gilt die Vorlage genauso fuer den Einzelartikel: bei `Artikel anlegen` (mit `Attribute` im Abrufumfang) und ueber die Buttons `Attribute automatisch ausfuellen` / `Automatisch ermitteln (Internet)` im Attribute-Tab - vorher wurde der Haken dort nicht beachtet. Treffer ausserhalb der Vorlage erscheinen im Einzelartikel im Bestaetigungsdialog mit dem Hinweis `nicht in der Vorlage (Referenzartikel)` statt in der Maske; die Statuszeile nennt die Vorlage (`Vorlage 0012,8-706R 10: 3 Attribute`) und danach, welche Vorlage-Attribute offen geblieben sind (`Vorlage 0012,8-706R 10: 1 von 3 Attributen nicht gefunden (3507)`). Die Werte kommen weiterhin je Artikel aus Text, KI bzw. Internet-Recherche; die Werte des Referenzartikels bekommt die KI nur als Formatbeispiel (Schluesselwerte mit Bezeichnung, z. B. `1 | Einohrklemme`). Regel- oder KI-Treffer fuer andere Attribute (z. B. Klemmbereich von/bis 3128/3129 oder Zusatzinfo) werden nicht geschrieben, sondern landen in der `Attribute_Pruefliste.xlsx`; aus einem Vorlage-Lauf wird nichts in die `Attribut_Zuordnung.xlsx` gelernt. Solange Vorlage-Attribute fehlen, fragt die App die KI unabhaengig von der Textlaenge. `Aus Maske` uebernimmt die Artikelnummer der Maske. Der Referenzartikel muss im Output-Ordner Attribute haben, sonst startet der Lauf nicht (Einzelartikel: der Button zeigt denselben Hinweis, nach `Artikel anlegen` steht er nur in der Statuszeile und die Findung entfaellt). Ohne Haken entscheidet die KI je Artikel selbst - bei fachlich gleichwertigen Attributen (Spannbereich vs. Klemmbereich von/bis) kann das zwischen gleichartigen Artikeln wechseln. **1:1 aus dem Referenzartikel (seit v0.1.51):** Das Feld `davon 1:1 uebernehmen (Kriterien-IDs)` unter dem Referenzartikel (Standard `9202, 9213, 9220, 9798` = Zusatzinfo, Zusatz LI/SI 2, Information, Entspricht Norm; Komma, Semikolon oder Leerzeichen als Trenner, unbekannte IDs werden mit Hinweis ignoriert) nennt die Attribute, die nicht je Artikel gesucht, sondern mit allen Zeilen des Referenzartikels uebernommen werden - der Referenzartikel wird von Hand gepflegt und ist die Wahrheit, auch bei mehreren Zeilen derselben ID (z. B. vier Zusatzinfo-Zeilen). Diese IDs bekommt weder die Regel-Stufe noch die KI als Kandidaten; vorhandene Zeilen dieser IDs beim Artikel werden durch die Zeilen des Referenzartikels ersetzt - aber nur fuer IDs, zu denen der Referenzartikel wirklich Zeilen hat. Eine Kopier-ID ohne Zeile im Referenzartikel (z. B. `9220`, wenn der Referenzartikel nur Zusatzinfo-Zeilen hat) wird weder kopiert noch gesucht, und vorhandene Zeilen des Artikels dazu bleiben stehen (nichts wird still geloescht); Startmeldung, Statuszeile und Abschlussdialog nennen sie als `ohne Zeile im Referenzartikel: 9220`. Soll eine solche Zeile trotzdem ersetzt werden, bekommt der Referenzartikel eine eigene Zeile dieser ID. Die uebrigen Vorlage-Attribute werden wie bisher gesucht und add-only ergaenzt, und die Attribute des Artikels werden wie beim Referenzartikel sortiert: Marke (8988) zuerst, dann in der Reihenfolge des Referenzartikels, uebrige Zeilen dahinter - auch nach `Uebernehmen` im Bestaetigungsdialog. Ein erneuter Lauf auf demselben Artikel aendert nichts mehr (keine Doppelten). Die Zeilen des Referenzartikels kommen aus der Maske, wenn sie ihn gerade zeigt (eine dort eben gepflegte Zeile gilt sofort), sonst aus dem noch wartenden Live-Schreibauftrag bzw. dem Artikelverzeichnis; der Massenimport loest die Vorlage erst nach Abschluss der Live-Speicherungen auf. Im Einzelartikel erscheinen die 1:1-Zeilen im Bestaetigungsdialog als angehakte Zeilen mit dem Hinweis `aus Referenzartikel <Nr> uebernommen`, die Statuszeile meldet `Vorlage <Nr>: 5 Zeilen 1:1 uebernommen, 3 von 3 gesuchten Attributen gefunden` und, falls Zeilen des Artikels ersetzt wurden, `2 vorhandene Zeilen ersetzt (9202, 9213)`; im Massenimport ersetzt die App die Zeilen in der `Attribute.xlsx`, der Abschlussdialog nennt `Vorlage: n Zeilen 1:1 uebernommen, k vorhandene Zeilen ersetzt`, und der `Attribut-Testbericht.xlsx` fuehrt die 1:1-Zeilen mit Status `Referenz 1:1` sowie jede ersetzte Zeile mit ihrem alten Wert und Status `ersetzt` (Quelle `Referenz`). Der Testlauf berichtet beides (`wuerde ersetzt`), ohne zu schreiben. Die Marke bleibt davon unberuehrt. Feld leeren = nichts 1:1, alles wird gesucht wie bis v0.1.50.
- **Auslaufartikel ueberspringen** (Haken, standardmaessig an): Zeigt kunzer.de statt des Produkts nur `Artikel ist nicht mehr im Sortiment.`, wird nichts uebernommen - im Massenimport stehen diese Artikel in der Warnliste und im Abschlussdialog, nicht in der Loeschliste; im Einzelabruf erscheint ein Hinweis. Ohne die Option landete der Hinweis bisher als Kurzbezeichnung und Text in den Output-Dateien und die Marke wurde `KUNZER`.
- **Testlauf** (standardmaessig aktiv): mit der Checkbox `Testlauf: nur Bericht, nichts schreiben` wird NICHTS in die Attribute.xlsx geschrieben - stattdessen entsteht der vollstaendige `Attribut-Testbericht.xlsx` (je Angabe: Attribut, Wert, Status, Quelle Regel/KI). Ideal, um einen grossen Lauf erst zu pruefen. Der Bericht wird auch im scharfen Lauf als Protokoll geschrieben
- ab 4 Artikeln laufen die KI-Aufrufe gesammelt ueber die Anthropic **Batch API** (50 % Rabatt bei gleicher Qualitaet). Die Artikel werden erst nach dem Batch finalisiert (Attribute schreiben, Pruefliste, Testbericht); waehrend des Wartens zeigt der Fortschritt die Wartezeit und den Stand bei Anthropic, z. B. `KI-Batch 1/1: 0/56 verarbeitet - wartet seit 12 min (Anthropic: in_progress) ...` (Aktualisierung jede Minute)
- **Offener KI-Batch (Sicherung `apollo_import_ki_batch.pkl`)**: Vor dem Einreichen sichert die App die gesammelten KI-Anfragen samt Artikeldaten und danach die Batch-ID in der Datei `apollo_import_ki_batch.pkl` neben dem Programm (wie `apollo_import_gui_state.json`). Ist der Batch nach **2 Stunden** noch nicht fertig, wird er NICHT abgebrochen (er laeuft bei Anthropic weiter und ist bereits bezahlt): der Lauf endet mit dem Hinweis `KI-Batch <ID> nach 2 h noch nicht fertig - die Attribute von N Artikeln koennen spaeter abgeholt werden`, fuer diese Artikel wird nichts geschrieben, die Sicherung bleibt liegen. Dasselbe gilt bei Absturz oder Schliessen waehrend des Wartens. Abholen: Button `KI-Batch abholen` im Massenimport (nur aktiv, solange die Sicherung liegt) oder die Rueckfrage beim naechsten Programmstart (`Ein KI-Batch vom ... mit N Artikeln ist noch offen. Ergebnisse jetzt abholen?`; `Nein` laesst die Sicherung fuer den Button liegen). Das Abholen laeuft wie ein normaler Massenimport (ohne erneuten Kunzer-Abruf), schreibt die Attribute add-only, ergaenzt Pruefliste und Testbericht des urspruenglichen Laufs und loescht die Sicherung; weicht der eingestellte Ausgabeordner von dem des Laufs ab, startet es nicht. `Abbrechen` waehrend des Wartens bricht den Batch bei Anthropic ab, finalisiert die Artikel regelbasiert und loescht die Sicherung. Ein neuer Lauf mit KI-Batch wuerde eine noch liegende Sicherung ueberschreiben - die App fragt vorher nach.
- **Abholen ohne Verlust**: Scheitert die Abfrage eines bereits eingereichten Batches am Claude API Key oder an der Verbindung (Key gewechselt, offline/VPN), bleibt die Sicherung liegen und der Lauf meldet `KI-Batch <ID> konnte nicht abgerufen werden: ... - Sicherung bleibt liegen`; nach dem Beheben einfach erneut `KI-Batch abholen`. Bricht die Finalisierung mittendrin ab (typisch: `Attribute.xlsx` in Excel geoeffnet), bleibt die Sicherung ebenfalls liegen und merkt sich die bereits geschriebenen Artikel - ein erneutes Abholen schreibt nur die restlichen. Artikel, die zwischen Einreichen und Abholen im Verzeichnis geloescht wurden, werden beim Abholen uebersprungen (Hinweis in der Warnliste) statt in der `Attribute.xlsx` wieder angelegt. Zeigt die Maske gerade einen Artikel des Batches, wird sie nach dem Lauf aus den Dateien neu geladen (wie der letzte Artikel eines normalen Massenimports), damit das naechste Live-Speichern die abgeholten Attribute nicht ueberschreibt; hat die Maske ungespeicherte Aenderungen, bleibt sie unveraendert und der Abschlussdialog bittet darum, den Artikel neu zu laden, bevor weiter gespeichert wird.
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

## Stueckliste und Zubehoerliste

Reiter `✎ Stückliste/Zubehör` - entspricht in IMS dem Reiter `Artikelliste(n)` mit den Listentypen `Stückliste` und `Zubehörliste`:

- **Stückliste**: was im Lieferumfang enthalten ist, mit Menge (z. B. Heber `AJ-45T3GFL`: 1x `02041`, 2x `02042`).
- **Zubehörliste**: optionales Zubehoer (Menge in der Regel 1).

Je Position: Typ, `Lieferantenartikelnummer` (Listenartikel), `Menge`, `TecDoc GenArt ID` und `Produktgruppe` **des Listenartikels** (nicht des Hauptartikels) und `Sort Nr`. Beim Eingeben der Artikelnummer schlaegt die App Artikel aus dem Artikelverzeichnis vor und belegt GenArt und Produktgruppe aus der gespeicherten GenArt des Listenartikels vor (bei mehreren GenArts zur Auswahl). Steht der Listenartikel nicht im Output, GenArt-ID von Hand eingeben - die Produktgruppe kommt dann aus dem GenArt-Katalog. Die `Sort Nr` zaehlt je Artikel und Typ ab 1 in der Reihenfolge der Tabelle; mit ▲/▼ innerhalb eines Typs umsortieren.

Geschrieben wird live in `Artikellisten.xlsx` (Blatt `Artikellisten`): `Artikelnummer`, `Typ`, `Lieferantenartikelnummer`, `Menge`, `TecDoc GenArt ID`, `Produktgruppe`, `Sort Nr`, `Zuletzt geschrieben am`. Wie GenArt, OE-Nummern und Attribute zaehlen die Listen zu den gepflegten Daten: Massenimport und Kunzer-Abruf lassen sie unveraendert, Loeschen im Artikelverzeichnis entfernt die Listen des Artikels (nicht seine Verwendung in Listen anderer Artikel). Abgelehnt werden: Menge keine ganze Zahl ab 1, GenArt fehlt bzw. keine Zahl, der Artikel selbst als Listenartikel, dieselbe Position zweimal im selben Typ. Das Artikelverzeichnis zeigt die Anzahl in der Spalte `Listen`; die Suche findet einen Artikel auch ueber die Nummern seiner Listenartikel.

## Attribute automatisch finden

Die Attribut-Findung arbeitet hybrid in zwei Stufen und wird ueber den Button `Attribute automatisch ausfuellen` im Attribute-Tab oder automatisch nach dem Kunzer-Laden gestartet - Letzteres nur, wenn `Attribute` im Abrufumfang angehakt ist (bis v0.1.45 lief sie nach jedem `Artikel anlegen`, auch bei reinem Bild-Umfang, und verursachte dabei KI-Kosten). Der KI-Nachlauf fuer Suchwoerter laeuft davon unabhaengig.

Ist im Abrufumfang `Attribute nach Referenzartikel` angehakt, gilt die Attribut-Vorlage des Referenzartikels seit v0.1.50 auch fuer den Einzelartikel (Button und `Artikel anlegen`), nicht nur im Massenimport: gefunden und uebernommen werden nur die Attribute des Referenzartikels, Treffer fuer andere Attribute gehen mit dem Hinweis `nicht in der Vorlage (Referenzartikel)` in den Bestaetigungsdialog, die KI wird nur gefragt, solange Vorlage-Attribute fehlen, und die Statuszeile nennt die nicht gefundenen Vorlage-Attribute (Details im Abschnitt "Listenimport CSV/XLSX"). Seit v0.1.51 werden die Kriterien-IDs aus `davon 1:1 uebernehmen` (Standard Zusatzinfo 9202, Zusatz LI/SI 2 9213, Information 9220, Entspricht Norm 9798) nicht gesucht, sondern mit allen Zeilen des Referenzartikels uebernommen: vorhandene Zeilen dieser IDs werden ersetzt (nur fuer IDs, zu denen der Referenzartikel Zeilen hat; ersetzte Zeilen nennt die Statuszeile), die Attribute wie beim Referenzartikel sortiert (Marke zuerst, auch nach `Uebernehmen` im Bestaetigungsdialog), und die KI bekommt diese IDs nicht als Kandidaten - Zusatzinfos werden dann nicht mehr je Artikel auf die Slots 9202/9213/9220 verteilt. Zeigt die Maske gerade den Referenzartikel, zaehlen ihre Zeilen (eine eben gepflegte Zeile gilt sofort).

PDF-Dokumente des Artikels fliessen als Text mit ein. Ab v0.1.49 werden bei mehrseitigen PDFs nur die Seiten verwendet, auf denen die eigene Artikelnummer steht; steht sie auf keiner Seite, wird das PDF ignoriert. Vorher uebernahm die Findung aus einer Broschuere, die an allen Artikeln einer Serie haengt (WES-Werkstattmodule), die Werte fremder Produkte - 54 WES-Artikel bekamen so die Kabellaenge des LED-Adapters und die Laenge von Modul 1. Einseitige PDFs (Datenblaetter) gelten immer als artikelbezogen.

**Stufe 1 - Regeln (immer aktiv, kostenlos):** Erkannt werden `Label: Wert`- und Tab-getrennte Zeilen aus Technische-Daten-Bloecken, z. B.:

- `Gewicht: 612 g` -> `Gewicht [kg]` = `0,612` (Einheiten werden automatisch umgerechnet)
- `Hubhoehe: 150 - 530 mm` -> Bereich als `Wert` + `Wert bis`
- `Abmessungen (LxBxH): 155 x 17 x 13,5 mm` -> aufgeteilt in `Laenge`, `Breite`, `Hoehe`
- Schluesselwert-Attribute werden nur uebernommen, wenn der Wert exakt in der Werteliste steht
- Normbezeichnungen mit Ausgabejahr wie `EN ISO 374-1:2016/TYPE B` (DIN, EN, ISO, IEC, ANSI, ASTM, VDE, UL) gelten seit v0.1.51 nicht als `Label: Wert`-Zeile - der Doppelpunkt gehoert zur Norm, die Zeile bleibt Freitext fuer die KI. Vorher wurde ueber die gelernte Zuordnung `EN ISO 374-1 -> 9798` der Wert `2016/TYPE B` als `Entspricht Norm` uebernommen. `Norm: EN ISO 374-1:2016` (Label `Norm`) funktioniert weiterhin.

Grundlage ist die pflegbare Zuordnungsdatei `Attribut_Zuordnung.xlsx` (Standard `G:\Apollo\Attribut_Zuordnung.xlsx`, Spalten `Text-Label` und `TecDoc Kriterien ID`), die im Projekt-Tab unter `Datenstaemme` konfiguriert wird.

**Stufe 2 - Claude-KI (optional, braucht Anthropic API Key):** Alles, was die Regeln nicht zuordnen konnten (Fliesstext, Synonyme, unbekannte Labels), geht zusammen mit dem Text der verlinkten PDF-Dokumente an die Claude API. Ueber den zweiten Button `Automatisch ermitteln (Internet)` recherchiert die KI zusaetzlich per Websuche auf Herstellerseiten und in Datenblaettern (ca. 10-30 ct pro Artikel; im Massenimport ueber die Option `Internet-Recherche auch im Massenimport`, siehe unten). Auch Recherche-Treffer zaehlen nur, wenn sie eindeutig zum Artikel gehoeren und die lokale Validierung bestehen. Die KI bekommt dabei nur eine lokal vorgefilterte Kandidatenliste (max. 300 plausible Attribute statt aller ~5.000) - das haelt die Kosten klein. Der API Key und das Modell werden im Projekt-Tab unter `APIs` eingestellt (auch per Umgebungsvariable `ANTHROPIC_API_KEY`); Standardmodell ist `claude-opus-4-8`, guenstigere Modelle sind waehlbar.

**Ohne Text in der Maske:** Ist kein deutscher Text erfasst (z. B. weil die Texte bereits in Apollo gepflegt sind), holt sich die Attribut-Findung den Produkttext selbst von kunzer.de - nur zum Auswerten. Gespeichert wird er nirgends; Texte, IDs und alle anderen Daten bleiben unangetastet.

**Gestufte Uebernahme:** Jeder Treffer - egal ob aus Regeln oder KI - wird lokal streng gegen die Attribut-Stammdaten validiert (Zahlformat, Einheiten-Umrechnung, Schluesselwert exakt in Werteliste, maximale Laenge). Nur validierte Treffer werden automatisch uebernommen, und nur fuer Kriterien-IDs, die der Artikel noch nicht hat (bestehende Zeilen werden nie ueberschrieben). Unsichere Treffer erscheinen wie bisher im Bestaetigungsdialog; dort koennen nicht erkannte Angaben manuell zugewiesen und mit `Zuordnung merken` dauerhaft gelernt werden.

**Zusatzinfo-Slots als Auffangbecken:** Wichtige technische Angaben, die zu keinem Attribut passen, liefert die KI als kurze Stichworte (max. 20 Zeichen, z. B. `IP54`, `inkl. Koffer`). Da Apollo pro Kriterien-ID nur eine Zeile akzeptiert, verteilt die App die wichtigsten Infos auf drei Slots: `9202 Zusatzinfo`, `9213 Zusatz LI/SI 2`, `9220 Information`. Weitere Infos landen im Pruefdialog bzw. in der Pruefliste.

**Label-Mining (Selbst-Training):** Der Button `Labels per KI zuordnen` bei den Datenstaemmen analysiert alle Texte im Output-Ordner, sammelt unbekannte Text-Labels samt Beispielwerten und laesst die KI sie EINMALIG zuordnen. Uebernommen wird nur, was an echten Beispielwerten lokal validiert wurde (Hinweis `KI-Miner` in der Zuordnungsdatei); der Rest landet in der `Label-Zuordnung-Pruefliste.xlsx`. Danach loest die kostenlose Regel-Stufe diese Labels ohne KI-Aufruf - der Grossteil aller Spezifikationszeilen wird damit gratis.

**Lernschleife:** Ordnet die KI ein neues Text-Label erfolgreich zu (z. B. `Farbton` -> `Farbe`), wird die Zuordnung mit Hinweis `KI` in die `Attribut_Zuordnung.xlsx` geschrieben. Beim naechsten Artikel greift dann schon die Regel-Stufe - ohne API-Aufruf. KI-Eintraege sind in der Datei ueber die Hinweis-Spalte auffindbar und koennen dort jederzeit korrigiert oder geloescht werden.

Die KI erhaelt den Produkttext als reinen Inhalt und darf nur Attribute aus der mitgeschickten Kandidatenliste waehlen; alles andere wird verworfen. Ohne API Key (oder bei API-Fehlern) laeuft die Findung vollstaendig regelbasiert weiter.

## Marke (Attribut 8988)

Die Marke eines Artikels wird in Apollo als Attribut `8988 Marke` (Format `Alphanumerisch`, max. 20 Zeichen) gepflegt. Bei jedem Kunzer-Abruf - im Einzelartikel (`Artikel anlegen`) wie im Massenimport (`Liste importieren`, `Artikelliste aktualisieren`) - setzt die App diese Zeile automatisch, rein lokal und ohne KI-Aufruf. Die Marke steht auf kunzer.de nicht im Seitentext, sondern wird aus den Signalen der Produktseite erkannt, in dieser Reihenfolge (der erste Treffer gewinnt):

1. **Medienordner**: erstes Pfadsegment der `restapi.kunzer.de`-Links von Bildern und Dokumenten, z. B. `airman`, `tdl`, `weber hydraulik`, `moto meter`, `norma`, `oetiker`, `argus - batterietester` oder `kunzer werkzeuge` (Ordner ohne Marke aus der Liste, z. B. `suprabeam` oder `handschuhe`, zaehlen nicht). Nennt der Ordner `KUNZER`, gewinnt trotzdem eine Fremdmarke aus dem Kategoriepfad oder aus den ersten Woertern von Titel/Kurzbezeichnung (z. B. `Oetiker Montagezange ...` im Ordner `kunzer werkzeuge` -> `OETIKER`), aus einem Link-Label, das mit der Marke beginnt (`Oetiker - Lochbandklemmen 259 - DualHook`; `Anleitung Oetiker-Zange` zaehlt nicht), oder ein einzelnes Bild/Dokument im Ordner einer Fremdmarke neben der Kunzer-Mehrheit (DualHook `18500260`-`18500267`: drei Bilder in `kunzer werkzeuge`, eins in `oetiker` -> `OETIKER`). Shop-Suchwoerter und Langtext zaehlen in diesem Fall nicht - dort stehen Gattungsnamen wie `Oetiker-Zange` auch bei Kunzer-Produkten
2. **Kategoriepfad** der Produktseite (Breadcrumb), z. B. `Produkte > Werkzeuge > Rad und Reifen > airMAN` oder `Produkte > Hebetechnik > Lufthydraulische Heber T D L`
3. **Shop-Suchwoerter** des Artikels (Feld `SUCHWOERTER` der Produktdaten, z. B. `Dichtmilch;Pannenpilot;Airman`)
4. **Titel, Kurzbezeichnung, Text** - Treffer nur an Wortgrenzen (`Norma` trifft nicht `normal`); `KUNZER` zaehlt in Texten nicht als Treffer, weil es ohnehin der Standard ist
5. **Link-Labels** der Produktseite (z. B. `AIRMAN - 71-051-011`)
6. ohne Treffer: `KUNZER` (Kunzer-Eigenartikel tragen zudem EAN mit GS1-Praefix `4260174`)

Erkannt werden nur Marken aus der **Markenliste** (Einstellungen, Tab `API & Datenstaemme`, Block `Marke (Attribut 8988)`). Jeder Eintrag nennt den kanonischen Namen, der in die `Attribute.xlsx` geschrieben wird - das ist die Schreibweise der Apollo-Liste `Product Brand Name` - und optional Schreibweisen (Aliasse): Eintraege mit `;` trennen, Aliasse nach `=` mit `,` - Gross-/Kleinschreibung spielt keine Rolle, Leerzeichen und Bindestriche in Aliassen sind beliebig (`T D L` trifft auch `TDL`, `weber hydraulik` auch `Weber-Hydraulik`). Standard:

```
KUNZER; OETIKER; T D L=tdl; JOKON; Airman=air man; MOTO METER=motometer,moto-meter; WEBER=weber hydraulik,weber-hydraulik; ARGUS; NORMA
```

Die Liste wird mit dem Sitzungszustand gespeichert; `Standard` stellt die eingebaute Liste wieder her (eine gespeicherte unveraenderte Standardliste aus v0.1.36 wird automatisch ersetzt). Kanonische Namen muessen in 20 Zeichen passen (laengere werden abgeschnitten).

**Schreibweise angleichen:** Steht in einer vorhandenen `8988`-Zeile eine Marke der Liste in anderer Schreibweise (z. B. `MotoMeter`, `TDL`, `Kunzer`, `Weber Hydraulik`), bringt der naechste Kunzer-Abruf bzw. Massenimport des Artikels sie auf die Schreibweise der Liste (`MOTO METER`, `T D L`, `KUNZER`, `WEBER`). Werte, die keiner Marke der Liste entsprechen, bleiben unveraendert.

Verhalten:

- **Manuelle Pflege gewinnt:** Hat der Artikel bereits eine `8988`-Zeile mit Wert (in der `Attribute.xlsx` bzw. im Artikelverzeichnis), bleibt sie unveraendert - die erkannte Marke wird nie darueber geschrieben. Nur eine leere `8988`-Zeile wird ersetzt. Pro Artikel gibt es genau eine `8988`-Zeile; die automatische Attribut-Findung legt keine zweite an.
- **Einzelabruf:** Bis v0.1.35 wurden beim `Artikel anlegen` Attribute, GenArt, OE-/Vergleichsnummern, Fahrzeuge und alle nicht abgerufenen Texte/Medien der Maske geleert und dieser Leerstand live geschrieben - vorhandene Daten des Artikels gingen damit verloren. Jetzt kommt alles, was nicht neu von kunzer.de abgerufen wird, aus dem gespeicherten Stand des Artikels (inkl. noch nicht geschriebener Aenderungen); die gespeicherten Attributzeilen bleiben erhalten (Grundlage ist der gespeicherte Stand des Artikels, nicht die Maske: zeigt die Maske noch einen zuvor geladenen anderen Artikel, gehen dessen Zeilen nicht auf den neuen Artikel ueber), die Marke wird ergaenzt und zusammen mit den uebrigen Daten live gespeichert - auch wenn nur Medien im Abrufumfang sind.
- **Massenimport:** Die Marke wird fuer jeden Artikel ergaenzt (add-only in die `Attribute.xlsx`), unabhaengig von der Option `Attribute (automatisch, Regeln + KI)` und auch im Testlauf. Der Abschlussdialog nennt, wie viele Artikel eine Marke bekommen haben und wie viele schon eine hatten; der `Attribut-Testbericht.xlsx` enthaelt je Artikel die erkannte Marke mit Signal (Status `automatisch` bzw. `vorhanden`, Quelle `Marke`).
- **Ungueltige vorhandene Attributzeile:** Bricht eine gespeicherte Zeile die Formatpruefung (z. B. Bereich mit vertauschtem von/bis), kann der Artikel keine Attribute speichern - auch die Marke nicht. Die App meldet das nach dem Abruf in einem Hinweis; nach der Korrektur im Tab `Attribute` wird automatisch gespeichert. Suchwoerter werden davon unabhaengig geschrieben.
- **Massenimport bei gesperrter `Attribute.xlsx`** (z. B. in Excel geoeffnet): Der Marken-Schritt setzt nach dem ersten Fehler fuer den Rest des Laufs aus, der Abschlussdialog nennt die Zahl der Artikel ohne Marke.

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
- `Artikellisten.xlsx`
- `Gebrauchsnummern.xlsx`
- `Fahrzeugverknuepfungen.xlsx`
- `Bilder.xlsx`
- `Dokumente.xlsx`
- `Videos.xlsx`
- `Web Link.xlsx`

Alle Output-Dateien enthalten zusaetzlich die Spalte `Zuletzt geschrieben am`. Darin steht pro Zeile, wann diese Daten zuletzt aus der GUI geschrieben wurden; unveraenderte Zeilen behalten beim erneuten Speichern ihren Zeitstempel. Daneben liegt ab v0.1.52 die `Apollo-Importstand.json` (siehe `Apollo-Importstand und Loeschliste`); sie ist keine Importdatei fuer Apollo.

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
