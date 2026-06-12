# Bewerbungsmanager

Kleines, lokales Tool zur Verwaltung von Bewerbungen. Alles läuft im Browser, Daten werden in `localStorage` abgelegt.

## 🎯 Features

- **📋 Bewerbungen verwalten**: Erfasse Firma, Position, Datum, Status und Priorität
- **📊 Dashboard mit Analytics**: 
  - Response-Rate
  - Durchschnittliche Antwortzeit
  - Erfolgsquote
  - Wöchentlicher Trend
  - Bewerbungs-Funnel
- **📝 Intelligente Notizen**: 
  - Interview-Vorbereitung mit Templates
  - Feedback-Erfassung (Absagen, Erfolge, Learnings)
  - Nachfass-Notizen
- **🏷️ Filterung & Suche**: Nach Status, Priorität, Tags
- **⏰ Erinnerungen**: Termine für Follow-ups
- **📎 Anhänge**: PDF, Anschreiben als Base64 im Browser speichern
- **📈 Diagramme**: Bewerbungen pro Monat, Status-Verteilung

## 🚀 Nutzung

1. Öffne die [Live-Version](https://[dein-github-username].github.io/bewerbungsmanager/)
2. Neue Bewerbung mit "Neue Bewerbung" hinzufügen
3. Dashboard ansehen für Statistiken
4. Intelligente Notizen für Interviews und Feedback nutzen

**Hinweise:**
- Alle Daten sind lokal im Browser (localStorage)
- Funktioniert vollständig offline
- Keine Synchronisierung zwischen Geräten
- Löschen des Browserverlaufs = Daten weg!

## 💻 Technologie

- HTML5
- CSS3 (Flexbox)
- Vanilla JavaScript (kein Framework)
- Chart.js für Diagramme
- localStorage für Datenspeicherung

## 📄 Dateien

- `index.html` — Hauptseite & UI
- `style.css` — Styling
- `app.js` — Logik & Speicherung

