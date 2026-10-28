# M9: Referenzverhalten und Portierungsnachweise

Die VBA-Spezifikation ergibt sich aus dem vorhandenen Code und den Tests derselben Revision. Die Fälle sind synthetisch; produktive Mailtexte und Credentials werden nicht übernommen. Aktuelle Durchführungsergebnisse gehören ins [Changelog](../../CHANGELOG.md), offene Arbeiten in die [Roadmap](../../ROADMAP.md).

## Prüfumfang

VBA-Unit-Tests und C#-Referenztests verwenden synthetische Daten. Die separat registrierten Outlook-Integrationstests laufen ausschließlich im isolierten Test-PST. Testergebnisse dieses Quellstands stehen in [PUBLICATION.md](../../PUBLICATION.md); fachliche Zuordnungsqualität und tatsächliche Outlook-Bedienung sind eigene Nachweise.
## Referenzmatrix

| Bereich | Verbindliches Referenzverhalten | Quelle in VBA | C#-Portierung |
|---|---|---|---|
| Archiventscheidung | Closed/NotFound per API archivieren; Open bleibt unabhängig von alter Statusmail; Unavailable nur mit aktiviertem Fallback und geschlossener Statusmail | `Tests/TicketDecisionTests.bas` | Initial implementiert, vollständige 16-Kombinationen-Matrix plus Bedarf der Mailprüfung |
| Sicherheitsstufen | Fünf Regeln, Default 1, globale Obergrenze 3; Minimum und Begrenzung auf 0–3; keine globale Zustandsübernahme zwischen Instanzen | `Tests/MatchLevelsTests.bas` | Initial implementiert, alle Regeln mit Konfigurationswerten und Caps von -1 bis 4 geprüft |
| Ticket-/Ordner-Parsing | Tag am Betreffanfang, Groß-/Kleinschreibung, Ausgabe-Reset bei Fehlschlag, Präfixverkürzung, kurze IDs nicht mit längeren verwechseln | `Tests/TicketParsingTests.bas` | Initiale Referenzfälle implementiert |
| Archivdatum | Unbekanntes Datum nutzt vorgegebenen heutigen Tag; Monatsname folgt lokaler Kultur | `Tests/TicketParsingTests.bas`, `Tests/TimeUtilTests.bas` | Lokale Datums-/Kulturformatierung implementiert; UTC-Konversion noch offen |
| API und Token | Zustände, Schließdatum, gelöschte/fehlende Tickets, 401-Retry, Timeouts, Token-Cache, Credential-Fehler ohne Secrets | `Tests/NinjaApiTests.bas`, `Tests/NinjaTicketLogicTests.bas`, `Tests/NinjaAuthTests.bas`, `Tests/NinjaTokenLogicTests.bas` | Offen |
| Ticketindex und Matcher | Paging, Cache, Priorität der Regeln, Ausschlüsse, Mehrdeutigkeit, geschlossen/nicht lesbar niemals verschieben | `Tests/NinjaTicketIndexTests.bas`, `Tests/TicketMatcherTests.bas`, `Tests/TicketMatchServiceTests.bas` | Offen |
| Vorschläge und Entscheidungen | Dry-Run nie verschiebend/vormerkend; Bestätigen prüft Schutzregeln erneut; Ablehnungen bleiben erhalten; Abbruch entscheidet nichts | `Tests/MatchReviewTests.bas`, `Tests/MatchLevelsTests.bas` | Offen |
| Audit und Regeln | Stabile Nachrichtenschlüssel, Monatsdateien/Aufbewahrung, Schreibfehler stören nicht, Freigabe vor Regelaktivierung | `Tests/AuditTests.bas`, `Tests/AuditReportTests.bas`, `Tests/ManualScanTests.bas`, `Tests/LearnedRulesTests.bas` | Offen |
| Outlook und Oberfläche | Ereignisse, Stop/Shutdown, Mail-/Ordneradapter, Archiv-Restore, Formular/Fallback, keine Änderungen im realen Postfach durch Tests | `Tests/OutlookAdapterTests.bas`, `Tests/ReviewFormTests.bas` | Regulärer VSTO-Host und eigene Abnahme noch offen |

## Ausführung

Regulärer Projektbuild und C#-Referenztests nach Einrichtung der Toolchain:

```powershell
.\tools\Test-DotNetCore.ps1
```

Begrenzte lokale Diagnose ohne Visual Studio/Targeting-Pack:

```powershell
.\tools\Test-DotNetCore.ps1 -RuntimeOnly
```

Der zweite Weg kompiliert ausschließlich die Fachlogik und ihre Tests gegen installierte Framework-Runtime-Assemblies. Er validiert weder net48-Referenz-APIs noch MSBuild-Projekt, VSTO-Host, Manifest, Installation oder Outlook-Verhalten. Die Tests starten kein Office, rufen keine API auf und verwenden keine echten Benutzerdateien. Diagnosedateien liegen unter dem ignorierten Verzeichnis `artifacts/core-runtime`.
