# M9: Architektur und Zielumgebung

Zielarchitektur des begonnenen M9-Schritts. Die VBA-Implementierung bleibt die funktionale Referenz; ein Outlook-VSTO-Host, Installer und Deployment sind noch nicht vorhanden.

## Ziel

- C# und **.NET Framework 4.8** als zunächst festgelegte gemeinsame Zielbasis für Fachlogik und späteren VSTO-Host. Die installierte Runtime allein ersetzt kein Targeting-Pack.
- Outlook Classic unter Windows, erster Pilot mit **64-Bit-Office**. Fachlogik und Referenztests sind AnyCPU; Host, Interop, VSTO-Runtime und spätere native Abhängigkeiten müssen zur Zielumgebung geprüft werden. 32-Bit-Office ist noch nicht abgenommen.
- Reguläres Outlook-VSTO-Projekt aus der Visual-Studio-Vorlage, mit generiertem Host-/Designer-Code, regulärem Build und Publish-Werkzeugen. Kein handgeschriebener Ersatzhost und keine kopierten Office-Assemblies aus lokalen Pilotprojekten.
- M0–M8 als angestrebte funktionale Ausgangsbasis: Sortierung, Archivierung, Matcher, Stufen, Vorschläge, Audit und Regelrunde. Teilportierungen aktivieren keine produktiven Aktionen. M10 mit zusätzlicher Detail-Textbox bleibt ausdrücklich nach der Migration.

## Voraussetzungen und Nachweisgrenzen

Die erste Zielumgebung ist Windows mit Outlook Classic in 64 Bit. Für den regulären Build werden Visual Studio mit Office/VSTO-Werkzeugen und .NET-Framework-4.8-Targeting-Pack benötigt. Eine installierte Framework- oder VSTO-Runtime allein ersetzt diese Werkzeuge nicht.

Der eingeschränkte Runtime-Testweg des C#-Kerns ist unabhängig vom Office-Host. Er bestätigt weder einen regulären net48-Build noch Installation oder Start eines Outlook-Add-ins. Aktuelle Checks des veröffentlichten Snapshots stehen in [PUBLICATION.md](../../PUBLICATION.md).
## Struktur des gestarteten Codes

```text
dotnet/NinjaRmm.Outlook.sln
  NinjaRmm.Outlook.Core          reine Fachlogik, keine Office-/Netzwerkzugriffe
  NinjaRmm.Outlook.Core.Tests    ausführbare Referenztests ohne NuGet/Office/Zugangsdaten

noch ausstehend:
  Infrastruktur                API, Token, Credential Manager, Dateien und Migration
  Outlook-VSTO-Host             Outlook-Adapter, Ribbon, Ereignisse, Review-Dialog
```

Die bisherigen VBA-Dateien bleiben unverändert. Der initiale C#-Kern implementiert `ArchivePolicy`, `MatchLevels` und `TicketParsing`. API, vollständiger Matcher, Audit, Speicherung, Vorschläge und Dialoge sind noch nicht portiert.

API- und Dateioperationen werden über getrennte Schnittstellen eingebunden. Office-Objekte verbleiben im Host-/STA-Kontext; HTTP und reine Datenarbeit benötigen Zeitlimits, Abbruch und eine kontrollierte Rückkehr zum UI-Kontext. Startup darf keine doppelte VBA-/Add-in-Verarbeitung starten.

## Bewusste Typ- und Fehlerentscheidungen

- Unbekanntes Schließdatum wird als `DateTime? = null` abgebildet statt als VBA-Datumswert `0`. Archivfunktionen erhalten lokales Datum, heutigen Tag und Kultur explizit; UTC-Konversion bleibt eine gesonderte Aufgabe.
- Ungültige numerische Ticketzustände werfen eine Ausnahme. Der VBA-`Case Else` für den Fallback wird nicht auf ungültige .NET-Enumwerte ausgedehnt; der spätere API-Adapter muss Fehler ausdrücklich auf `Unavailable` abbilden.
- Konfiguration ist instanzgebunden statt globaler Testüberschreibungen. Die fünf Regelstufen beginnen wie in VBA mit `LogOnly`, globale Obergrenze `3`.
- Regex-Zeitlimits begrenzen die Verarbeitung; ein Timeout wird nicht als Treffer behandelt. Der spätere Aufrufer muss ihn sicher protokollieren. ASCII-Ziffern erhalten die Ticketnummernsemantik; Ordnerpräfixe werden ordinal verglichen.
- Audit, Regeln und Ablehnungen werden vor dem ersten Schreibzugriff auf .NET-Seite versioniert und mit Sicherung/Rückweg migriert. Dieser Startschritt liest oder ändert keine vorhandenen Benutzerdateien oder Credentials.
- Persistente Konversationszuordnung über Neustarts bleibt eine offene Entscheidung; der bisherige In-Memory-Stand gilt nicht als bereits vorhandene Persistenz.

## Erforderliche reguläre Entwicklungswerkzeuge

Benötigt wird eine für das Unternehmen lizenzierte Visual-Studio-Ausgabe mit Office/VSTO-Werkzeugen, z. B. Visual Studio 2022 mit:

- Workload `Microsoft.VisualStudio.Workload.Office` (Office/SharePoint development).
- Komponente `Microsoft.VisualStudio.Component.TeamOffice` (VSTO), ClickOnce-Werkzeugen und MSBuild.
- .NET Framework 4.8 SDK und Targeting-Pack/Developer Tools.
- Passenden Outlook-Interop-/Office-Referenzen aus der regulären Entwicklungsumgebung.

Die Installation ist in diesem Schritt nicht erfolgt. Nach Einrichtung zuerst `tools/Test-DotNetCore.ps1` ohne Diagnoseschalter ausführen, dann ein reguläres Outlook-VSTO-Add-in-Projekt in diese Solution aufnehmen und den ersten Host-/Ribbonstart gesondert prüfen.

Quellen, am 27.10.2026 geprüft: [Microsoft: Office-Entwicklungsumgebung](https://learn.microsoft.com/en-us/visualstudio/vsto/how-to-configure-a-computer-to-develop-office-solutions?view=visualstudio), [Microsoft: VS-2022-Komponenten](https://github.com/MicrosoftDocs/visualstudio-docs/blob/main/docs/install/includes/vs-2022/workload-component-id-vs-community.md). Die Nennung der Komponentenliste legt keine Community-Lizenz für das Unternehmen fest.
