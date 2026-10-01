# Arbeitszeit-App iPhone v2

Mobile-first PWA zur lokalen Auswertung der eigenen Arbeitszeit-Excel-Dateien.

## Neu in v2
- Mehrere Zeiträume am selben Kalendertag werden zusammengeführt.
- Excel-Zeilen ohne Datum übernehmen das zuletzt genannte Datum.
- Wochentage werden ausschließlich aus dem echten Datum berechnet.
- Dienste über Mitternacht werden korrekt als Zeitspanne behandelt.
- 00:00–00:00 wird als 24-Stunden-Dienst interpretiert.
- Normale 30-Minuten-Pause wird nur bei einem zusammenhängenden Mo–Do-Block ab 07:00 bis mindestens 16:00 abgezogen.
- Abendpause wird je Zeitblock bei Ende nach 16:30 berücksichtigt.
- Urlaub schreibt die jeweilige Tages-Sollzeit gut.
- „Ausgleich Mehrarbeit Grundbetrieb“ ist dienstfrei und reduziert das Überstundenkonto um die normale Tages-Sollzeit.
- Feiertage und andere explizit dienstfreie Abwesenheiten erzeugen kein Tages-Soll.
- Beim erneuten Import eines Monats werden die betroffenen Tage ersetzt, damit alte Fehlinterpretationen nicht erhalten bleiben.
- Export enthält Zeiträume, Wochentag, Rohzeit, Arbeitszeit, Sollzeit und Saldo.

## Installation
Wie bisher über GitHub Pages veröffentlichen und auf dem iPhone in Safari öffnen → Teilen → Zum Home-Bildschirm.

Die importierten Arbeitszeitdaten werden lokal im Browser gespeichert. Keine Excel-Dateien ins GitHub-Repository hochladen.
