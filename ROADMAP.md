# Roadmap

Planungsstand: **27.10.2026**. Hier stehen ausschließlich offene Arbeiten, geplante Meilensteine und deren Abnahmekriterien. Den aktuellen Funktionsumfang beschreibt die [README](README.md); abgeschlossene Änderungen und historische Prüfergebnisse stehen im [CHANGELOG](CHANGELOG.md).

## VBA – Ausstehende Abnahme des veröffentlichten Quellstands

Diese Aufgaben betreffen den in der README beschriebenen integrierten Entwicklungsstand. Eine Quellveröffentlichung ist keine fachliche Abnahme; erforderlich bleiben die folgenden Prüfungen.

| Offene Arbeit | Abnahmekriterium |
|---|---|
| Outlook-Import und vollständige Kompilierung | Passender Entwicklungsstand importiert, *Debuggen > Kompilieren* ohne Fehler; Sicherung des VBA-Projekts vorhanden |
| M6: Matcher und Wiederanlauf | Eindeutige offene Treffer, Mehrdeutigkeit, ausgeschlossene/geschlossene/nicht lesbare Tickets, Probelauf und Wiederanlauf im isolierten Test-PST geprüft |
| Dauerhafte Konversationszuordnung entscheiden | Festlegen, ob Zuordnungen Outlook-Neustarts überstehen müssen; falls ja, Speicherung, Wiederherstellung und Tests umsetzen |
| M7: Audit und manuelle Zuordnungen | Nachrichtenschlüssel bleibt nach Verschieben stabil, Audit-Schreibfehler stoppen keine Verarbeitung, Bericht anhand echter freigegebener Entscheidungen geprüft |
| M7: erste Regelrunde | Ausreichende Bestätigungen/Ablehnungen auswerten, Kandidaten prüfen und ausschließlich ausdrücklich freigegebene Regeln aktivieren |
| M8: Dialoge in Outlook | Formular und MsgBox-Fallback prüfen: bestätigen, ablehnen, überspringen/abbrechen, wiederholter Aufruf, Hinweisbündelung und Audit-Einträge |

Die [Importanleitung](IMPORT.md) gilt ausschließlich für den dort bezeichneten Entwicklungsstand. Tests mit Outlook-Ordnern und Mails erfolgen nur in der lokalen Nicht-Exchange-PST `NinjaTest`, niemals im produktiven Postfach.
## M9 – Migration von VBA auf ein .NET-VSTO-Add-in (begonnen)

Ziel ist ein C#-VSTO-Add-in für **Outlook Classic unter Windows**, das die fachlichen Funktionen der VBA-Lösung übernimmt und regulär gebaut, installiert und aktualisiert werden kann. VSTO verwendet **.NET Framework**; die Arbeitsentscheidung .NET Framework 4.8 und der lokale Umgebungsbefund stehen in der [Architektur](docs/m9/architecture.md). Reguläre Entwicklungswerkzeuge und die Hostvalidierung bleiben Voraussetzungen für die weiteren Schritte. Neues Outlook unterstützt COM/VSTO nicht und benötigt eine eigene Architekturentscheidung. Quellen: [Microsoft: VSTO entwickeln](https://learn.microsoft.com/en-us/visualstudio/vsto/create-vsto-add-ins-for-office-by-using-visual-studio?view=visualstudio), [Microsoft: Outlook-Add-in-Plattformen](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/outlook-add-ins-overview), geprüft am 27.10.2026.

### M9.1 – Zielumgebung und Referenzverhalten festlegen

Architektur und Referenzmatrix sind angelegt ([Architektur](docs/m9/architecture.md), [Referenzfälle](docs/m9/reference-cases.md)). Offen bleiben die reguläre Visual-Studio-/VSTO-Toolchain mit net48-Targeting-Pack, deren Buildnachweis und die fachliche Bestätigung des Referenzumfangs.

- Reguläre Visual-Studio-VSTO-Toolchain und net48-Targeting-Pack einrichten, die in der Architektur erfasste Zielmatrix am VSTO-Host verifizieren und den regulären Projektbuild nachweisen.
- Den Funktionsumfang der ersten Add-in-Version anhand der vorhandenen VBA-Referenzfälle festlegen. Unabgenommene VBA-Funktionen benötigen mindestens eine dokumentierte Verhaltensspezifikation und Testfälle vor der Portierung; sie gelten nicht allein durch einen Merge als abgenommen.
- Referenzfälle und Testdaten für Sortieren, Archivieren, Offline-Fallback, Matching, Review, Audit und Regeln festlegen. Offene Punkte wie persistente Konversationszuordnung ausdrücklich entscheiden.
- Abnahme: dokumentierte Zielmatrix, Funktionsumfang und zugehörige Referenzfälle; keine Installation erforderlich.

### M9.2 – Fachlogik und API in C# portieren

Der initiale Kern ist in der [README](README.md) beschrieben. Noch zu portieren sind vollständiger Matcher und Ticketindex, API/Token/Credential-Adapter, UTC-Konversion, Vorschlags-/Review-Logik, Audit, Regeln und Datenmigration. Die Runtime-Diagnose ersetzt keinen Build gegen die net48-Referenzassemblies.

- Vom Office-Host getrennte Bibliothek für Parsing, Archiventscheidungen, Matching, Stufen, Vorschläge, Audit und Regeln; vorhandene anonymisierte Fixtures und Schutzfälle übernehmen.
- NinjaOne-API und Client-Credentials-Flow lesend erhalten, mit Token-Cache, Timeouts, Abbruch und Credential-Manager-Zugriff. Keine Secrets im Paket oder in Logs.
- Datenformate für Audit, Regeldatei und Ablehnungen versionieren; Wiederverwendung oder Migration mit Sicherung und Rückweg testen.
- Abnahme: automatisierter Build, Unit-Tests und Negativfälle; Abgleich mit den vereinbarten VBA-Referenzfällen. Dies ist noch keine Outlook-Abnahme.

### M9.3 – Outlook-Host, Ribbon und Review-Oberfläche

- Reguläres Outlook-VSTO-Projekt mit generiertem Host, Startup/Shutdown, Ribbon-Schaltflächen für Sortieren, Archivieren, Vorschläge und Berichte sowie einer geeigneten Review-Oberfläche.
- Mailereignisse, Sweep, Ticketordner und Archiv hinter Outlook-Adaptern kapseln. COM-Zugriffe im vorgesehenen Host-/STA-Kontext halten; HTTP/Datenarbeit so ausführen, dass Outlook bedienbar bleibt. Wiedereintritt, doppelte Events, Abbruch und Fehlerfälle behandeln.
- Abnahme: echter Add-in-/Ribbonstart sowie fachliche Tests ausschließlich im Test-PST; warmen/kalten Start, mehrere Fenster und Shutdown prüfen.

### M9.4 – Signierung, Installation, Updates und Rückweg

- Installationsverfahren (z. B. ClickOnce oder verwalteter Installer), Benutzer-/Maschinenumfang, feste Add-in-Identität und Versionsschema auswählen; Code-/Manifestsignierung und Herausgebervertrauen mit der IT abstimmen.
- Voraussetzungen, Updatequelle und Deinstallation dokumentieren; kein Entwickler-/Forge-Schreibtoken im Add-in. Updatefehler, Offlinebetrieb und Rollback einschließlich Registrierung und Konfiguration prüfen.
- VBA-Projekt sichern. Vor Aktivierung im Pilot die überlappenden VBA-Watcher, Regeln und automatischen Aufrufe deaktivieren, damit VBA und Add-in keine Mail doppelt verarbeiten. Beim Rückweg Add-in deaktivieren und den gesicherten VBA-Stand kontrolliert wieder aktivieren.
- Abnahme: Installation, Update und Deinstallation auf einer repräsentativen Zielmaschine, echter Hoststart und nachgewiesener Rückweg; Paket-/Manifestprüfung allein reicht nicht.

### M9.5 – Pilot und Ablösung

- Erst Protokollbetrieb, danach bestätigte Zuordnungen und erst nach fachlicher Freigabe Automatik je Regel. Bestehende Ordner, Daten und Zugänge erhalten.
- Sortieren, Archivmonat, Offlineverhalten, sichere Zuordnung, Bericht und Audit gemeinsam mit dem Nutzer prüfen; Startdauer und Outlook-Bedienbarkeit erfassen.
- VBA erst nach erfolgreicher fachlicher Abnahme ablösen. Ergebnis, Paketversion, Zielumgebung und verbleibende Grenzen dokumentieren.

## M10 – Lesbare Vorschlagsdetails nach dem VSTO-Rebase

**Abhängigkeit:** nach Abschluss der VSTO-Migration (M9). Die Darstellungsverbesserung wird in der VSTO-Oberfläche umgesetzt.

- Die Vorschlagsliste bleibt in der Standardansicht kurz. Eine eigene schreibgeschützte Textbox zeigt die Details des aktuell fokussierten Eintrags: vollständigen Ticketbetreff, Mailbetreff, Absenderdomain und Zuordnungsregel.
- Automatischen Zeilenumbruch und vertikales Scrollen im Detailbereich vorsehen, damit lange Texte vollständig lesbar sind.
- Den Mailinhalt bei Bedarf als einblendbare Klartextvorschau ergänzen. Einen begrenzten Ausschnitt ausdrücklich als solchen kennzeichnen; die aktuelle VBA-Schnittstelle liefert höchstens 4.000 Zeichen.
- Fokussierten Vorschlag und angekreuzte Vorschläge getrennt behandeln: Das Anzeigen der Details verändert weder die Häkchen noch eine Zuordnungsentscheidung.
- Abnahme: lange Betreffe und mehrzeiliger Inhalt sind ohne seitliches Abschneiden lesbar; Wechsel per Maus und Tastatur aktualisiert die passenden Details. Übernehmen, Ablehnen und Schließen behalten ihr bisheriges Verhalten.
## Reihenfolge und Entscheidungen

Zunächst den Umfang der ersten VSTO-Version in M9.1 festlegen. Wenn M6–M8 als VBA-Erweiterung veröffentlicht werden sollen, zuerst deren offene Abnahmen abschließen. Für eine direkte Portierung mindestens die Referenzfälle und das Sollverhalten freigeben; ein vorheriger VBA-Release ist keine zwingende Voraussetzung.

Reihenfolge: M9.1 → M9.2 → M9.3 → M9.4 → M9.5 → M10 (Vorschlagsdetails). Eine Kalenderzusage entsteht daraus nicht. Offene Entscheidungen sind Zielumgebung und Toolchain, Umfang der ersten Add-in-Version, dauerhafte Konversationszuordnung, Datenmigration sowie Deployment- und Updatequelle.

## Übergreifende Abnahmekriterien

- Sortierung, Archivstruktur und manuelle Zuordnungen erhalten; keine schreibenden NinjaOne-Ticketoperationen oder Ticketerstellung ergänzen.
- Referenzfälle für API-Zustände, Schließdatum, lokalisierte Statusnamen, Offline-Fallback und sichere Zuordnung prüfen. Aktuelle Tenantrechte und Erreichbarkeit bei der jeweiligen Live-Abnahme nachweisen.
- Für geänderte Logik automatisierte Tests und Negativfälle ausführen. Ergebnisse an Commit, Umgebung und Datum binden. Excel-/Parser-Erfolg ersetzt keine Outlook-Abnahme.
- Vor VBA-Veröffentlichung Signatur und Vertrauen prüfen; bei VSTO zusätzlich Paket, Installation, Host-/Ribbonstart und Rückweg nachweisen.
- Zugangsdaten aus Repo und Logs fernhalten. Bei Weitergabe von Audit und Berichten berücksichtigen, dass gekürzte Betreffe personenbezogene Angaben enthalten können.
