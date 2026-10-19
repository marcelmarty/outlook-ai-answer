# KI-Antwort – Outlook Add-in

Antwortvorschläge für E-Mails mit Claude (Anthropic API) oder einem OpenAI-kompatiblen Endpoint.

## Funktionen
- **Lesemodus**: Ribbon-Button «Antwortvorschlag mit KI erstellen» → öffnet direkt ein Antwortfenster mit dem Entwurf.
- **«KI-Antwort (Optionen)»**: Aufgabenbereich mit Tonalität pro Antwort, Stichworten («Was soll die Antwort aussagen?»), Live-Vorschau, Bearbeiten, Neu generieren.
- **Verfassen-Modus**: Vorschlag wird direkt in die offene E-Mail eingefügt.
- **Einstellungen** (Outlook Roaming Settings, geräteübergreifend): Anbieter, API-Key, Modell (inkl. «Modelle laden»), Tonalität (an E-Mail anpassen / formell / freundlich / locker / kurz / Custom), Länge, Sprache, Schweizer Rechtschreibung, Über mich, Signatur, Zusatzanweisungen.

## Installation
1. Outlook im Web öffnen → https://aka.ms/olksideload
2. «Meine Add-Ins» → «Benutzerdefinierte Add-Ins» → «Benutzerdefiniertes Add-In hinzufügen» → «Aus Datei hinzufügen…» → `manifest.xml`
3. Das Add-in erscheint danach auch in New Outlook / Outlook Desktop (gleiches Konto).

## API-Key
Claude-Abos (Pro/Max) enthalten keinen API-Zugang. Key erstellen unter https://console.anthropic.com → API Keys (Abrechnung nach Verbrauch).

## Deployment
Gehostet auf GitHub Pages: https://marcelmarty.github.io/outlook-ai-answer/ – `manifest.xml` wird aus `manifest.template.xml` erzeugt (`{{BASE}}` = Hosting-URL).
