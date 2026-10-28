# CLAUDE.md

## Projekt

VBA-Projekt für **Outlook Classic** (Outlook 2019+), das Ticket-Mails von NinjaRMM (lokal konfigurierte Instanz) automatisch in Ticket-Ordner einsortiert und geschlossene Tickets per NinjaRMM-API ins Archiv verschiebt. Ticketerstellung wurde mit M2 entfernt und bleibt außerhalb des Funktionsumfangs; siehe ROADMAP. Kommentare und Meldungen im Altbestand sind überwiegend deutsch, neuer Code ist englisch. Weitere Hintergründe: `README.md`.

Für VBA gibt es **keinen Build und keine CI**. Der gestartete .NET-Kern hat eine eigene Solution und `tools/Test-DotNetCore.ps1`; die lokale Runtime-Diagnose ist kein regulärer VSTO-Build. Der Code ist exportierter VBA-Quelltext (`.bas`, `.cls`, `.frm`/`.frx`), der manuell in den VBA-Editor von Outlook (`Alt+F11`) importiert wird. Unit-Tests für Logik ohne Outlook-Typen laufen headless (siehe „Arbeitsweise“), alles andere wird im VBA-Editor von Outlook getestet.

## Stand und geplante Migration

Dieses Repository enthält einen eigenständigen Quellstand mit VBA-Automation und initialem .NET-Kern. Aktueller Funktionsumfang: `README.md`; veröffentlichte Änderungen: `CHANGELOG.md`; offene Arbeiten: `ROADMAP.md`; Import: `IMPORT.md`.

M9 plant die Migration zu einem C#/.NET-Framework-VSTO-Add-in für Outlook Classic. Ein erster .NET-Framework-Kern mit Referenztests liegt unter `dotnet/`; VSTO-Host und Installer existieren noch nicht. Architektur und Toolchain stehen in `docs/m9/architecture.md`. Zielumgebung, Funktionsumfang und Deployment werden zuerst festgelegt; die Beschreibung ist keine implementierte Funktion. Ein Word-VSTO-Beispiel ersetzt keine Outlook-Abnahme. Neues Outlook benötigt eine eigene Architekturentscheidung.

## Basisstruktur (M0–M5; Erweiterungen M6–M8 in der README)

- `outlook/` – eigener Code
  - `Ticketautomatismus.bas` – **nur Outlook-Einstiegspunkte** (seit M3, ASCII/englisch, Regelpfade ohne Dialoge, explizites Fortschrittsfenster, Fehler in `TicketLog`): `RunEmailRule`, `RunEmailRuleScript(mail)` (Regel „Skript ausführen“), `RunArchiveRule`, `RunArchiveRuleWithProgress` (nicht-modales `ProgressForm` über `IProgressSink`/`ProgressSinkForm`), `StartMailWatcher`/`StopMailWatcher` (ereignisgesteuert über `TicketMailWatcher`, `Items.ItemAdd`; aus `Application_Startup` zu starten), `ProcessEmail` (Kompatibilität). Die Logik liegt in `TicketService.bas` (`ProcessMail`, `ArchiveClosedTickets`), reine Helfer in `TicketParsing`, `TicketDecision` (`DecideArchive`), `TicketConfig` (Konstanten inkl. `ENABLE_OFFLINE_FALLBACK = True`), `TimeUtil` (UTC→lokal), `TicketLog`. Outlook-Zugriff nur über die Schnittstellen `ITicketStore`/`ITicketFolder`/`IIncomingMail` und die dünnen Adapter `OutlookTicketStore`/`OutlookTicketFolder`/`OutlookIncomingMail` (diese Adapter sowie Einstiegspunkte und Watcher benötigen zusätzlich Outlook-Tests).
  - **API (read-only, seit M2):** `NinjaApi.bas` (`GetTicketState` mit `NinjaTicketOpen/Closed/NotFound/Unavailable`, `IsTicketClosed` = geschlossen, gelöscht oder 404, `GetTicketSubject`, `GetTicketClosedDate` = neuester Log-Eintrag mit Statuswechsel auf „Geschlossen“/„Closed“ oder `DELETE`, Zeit in UTC, 0 = unbekannt, `LastError`), `NinjaTicketLogic.bas` (reine Antwortlogik), `INinjaTransport.cls` + `NinjaTransport.cls` (HTTP-Naht mit Bearer-Token, UTF-8, ein Retry nach 401), `JsonLite.bas` (JSON-Parser ohne Verweise; `Parse`, `TryParse`, `GetString`/`TryGetNumber`/`GetBoolean`/`GetNode` per Punktpfad). Base-URL `NinjaAuth.NINJA_API_BASE_URL`. Ticketerstellung wurde entfernt.
  - **Authentifizierung (OAuth2 Client Credentials, seit M1):** `NinjaAuth.bas` (`GetAccessToken`, Token-Cache im Speicher, Fehlernummern `ERR_NO_CREDENTIAL`/`ERR_TOKEN_REQUEST`/`ERR_TOKEN_RESPONSE`, `IsServiceUnavailable` für den Offline-Fallback), `NinjaTokenLogic.bas` (reine Logik: Request-Body, Token-Antwort, Ablauf), `NinjaCredentialStore.bas` (Windows Credential Manager per `CredReadW`), `ITokenFetcher.cls` + `HttpTokenFetcher.cls` (HTTP-Naht, `MSXML2.ServerXMLHTTP.6.0`).
  - `ProgressHelper.bas` + `forms/ProgressForm.frm` – Fortschrittsanzeige (`ProgressStart/Update/Step/End`).
- `Tests/` – Unit-Test-Framework und Tests: `TestAssert.bas` (Assertions), `TestRunner.bas` (**generiert** mit `tools/New-TestRunnerModule.ps1`, nicht von Hand ändern; dasselbe Skript erzeugt `IntegrationRunner.bas` aus `Outlook*Tests.bas`; `RunAllTests`, `RunFailureProbe`, `RunLiveTests`; schreibt `%TEMP%\ninja_test_results.txt`), `SelfTests.bas`, `JsonLiteTests`, `NinjaTokenLogicTests`, `NinjaCredentialStoreTests`, `NinjaAuthTests` (mit `FakeTokenFetcher.cls`), `NinjaTicketLogicTests`, `NinjaApiTests` (mit `FakeNinjaTransport.cls`), `TicketParsingTests`, `TicketDecisionTests`, `TimeUtilTests`, `TicketServiceTests` (mit `FakeTicketStore.cls`, `FakeTicketFolder.cls`, `FakeMail.cls`); **Outlook-typisierte Integrationstests** `OutlookAdapterTests.bas` + `IntegrationRunner.bas` (generiert, `RunIntegrationTests`, F5, nur im Test-PST `NinjaTest`, geschützt durch `outlook/OutlookTestGuard.bas`; Ergebnis `%TEMP%\ninja_integration_results.txt`; nicht in `unit-test-modules.txt`), `LiveTests.bas` (opt-in, echter Token-Abruf), `Fixtures.bas` (**generiert** aus `Tests/fixtures/*.json` mit `tools/New-FixturesModule.ps1`, nicht von Hand ändern). Neue Tests: `Public Sub Test_Name()` in einem `*Tests.bas`-Modul anlegen und `tools\New-TestRunnerModule.ps1` ausführen (Outlook hat kein `Application.Run`, daher werden die Aufrufe generiert). Alt: manuelles Makro `ProgressFormTests`.
- `tools/` – PowerShell-Hilfen: `Invoke-VbaTests.ps1` (headless Unit-Tests, `-Probe` prüft den Runner, `-Live` macht einen echten Token-Abruf), `unit-test-modules.txt`, `ninja-smoketest.ps1` (lesender Live-Test der API mit dem Credential-Manager-Eintrag), `New-FixturesModule.ps1`, `New-TestRunnerModule.ps1`, `Test-VbaFiles.ps1` (prüft CRLF und ASCII-only; vor jedem Commit ausführen).

## Dokumentation nach Zeitbezug

- `CHANGELOG.md`: abgeschlossene Änderungen und historische Nachweise, mit Datum und Release-/Branch-/Abnahmestatus.
- `README.md`: aktueller Funktionsumfang, Voraussetzungen, Installation, Konfiguration und Bedienung; Branch-Geltungsbereich ausdrücklich nennen.
- `ROADMAP.md`: ausschließlich offene Aufgaben, geplante Änderungen, Abhängigkeiten und Abnahmekriterien. Erledigte Punkte in das Changelog überführen und den aktuellen Zustand in der README nachziehen.
- Historische Testergebnisse weder als aktuelle Verifikation noch lokale Implementierung als veröffentlichten Release darstellen. In den drei Dateien Inhalte verlinken statt denselben Verlauf mehrfach fortzuschreiben.

## Arbeitsweise

- **Human-in-the-loop:** Ausführbaren VBA-Code erst nach Outlook-Test und Freigabe veröffentlichen. Ein ausdrücklich beauftragter reiner Dokumentations-PR darf ohne Outlook-Import gepusht werden. Branch pro Meilenstein, kleine Commits; kein automatischer Merge.
- **Unit-Tests, wo möglich und sinnvoll:** neue oder geänderte Logik wird mit Tests geliefert. Reine Logik (Parsing, Datum, JSON-Auswertung) hat Vorrang. HTTP wird hinter einer Schnittstelle gefakt, Outlook-Ordnerlogik wird nur gegen ein Test-PST getestet, nie gegen das produktive Postfach. Neue Funktionen, die getestet werden sollen, werden `Public` angelegt.
- **Nicht im Scope:** Ticketerstellung in NinjaRMM. Dafür keinen neuen Code schreiben.
- **Sprache:** Neuer und geänderter Code, Kommentare, Bezeichner und Meldungen auf **Englisch**. Neue Dateien sind **ASCII-only** (keine Umlaute), weil der VBA-Editor beim Import ANSI liest. Bestehende deutsche Kommentare werden bei Gelegenheit (M5) übersetzt.
- **Tests ausführen:** Unit-Tests ohne Outlook-Typen laufen headless mit `tools\Invoke-VbaTests.ps1` (wegwerfbare Excel-Instanz, braucht `AccessVBOM=1` für Excel; Exit 0 = grün, 1 = Testfehler, 2 = Umgebungsproblem). Vor jedem Commit: `tools\Invoke-VbaTests.ps1` und `tools\Test-VbaFiles.ps1`. Module für diesen Lauf stehen in `tools\unit-test-modules.txt`; Logik ohne Outlook-Typen deshalb in eigene Module legen. Outlook-abhängige Tests starten vorerst manuell (`RunAllTests` im VBA-Editor), Ergebnis in `%TEMP%\ninja_test_results.txt`. Sicherheitseinstellungen (`AccessVBOM`, Makro-Level, Zertifikate) ändert Claude nicht, der Nutzer setzt sie selbst.
- **Integrationstests nur im Test-PST:** Der Store `NinjaTest` (lokale PST-Datei mit Anzeigenamen `NinjaTest`, kein Konto) ist die einzige Umgebung, in der Tests Outlook-Ordner und -Mails anlegen/löschen. Immer über `OutlookTestGuard.FindTestTicketsFolder` zugreifen, nie über den Standardposteingang.
- Das produktive Outlook-VBA-Projekt (`%APPDATA%\Microsoft\Outlook\VbaProject.OTM`) nie ohne vorherige Sicherung und Rückfrage verändern. Nach jedem Push-Kandidaten testet der Nutzer in Outlook.
- Offline-Fallback beim Archivieren (seit M3): `ENABLE_OFFLINE_FALLBACK` ist standardmäßig `True`; Entscheidung in `TicketDecision.DecideArchive`, getestet in `TicketDecisionTests`/`TicketServiceTests`.
- Python ist auf diesem Rechner nur ein Store-Stub. Skripte in PowerShell oder Node schreiben.
- Plan und Zielarchitektur: `ROADMAP.md`. Die Anmeldung läuft seit M1 über Client Credentials (Secret im Windows Credential Manager, Zielname aus `CredentialTarget` in lokalen Einstellungen, Benutzername = Client-ID, Kennwort = Secret).
- Nach dem Bearbeiten von Dateien (auch mit `sed -i`) immer `tools\Test-VbaFiles.ps1` ausführen: Zeilenenden können sich unbemerkt ändern. Markdown explizit als UTF-8 lesen und schreiben; VBA-Quellen bytegetreu als ASCII mit CRLF erhalten.
- Secrets nie ausgeben, loggen oder speichern. Aufrufe gegen die produktive Ninja-API nur lesend.

## Konventionen

- `Option Explicit` in jedem Modul; `PtrSafe`-Declares (64-bit Outlook).
- VBA-Dateien sind **ASCII mit CRLF**; Sonderzeichen über `ChrW` bilden. Markdown ist UTF-8. Beim Bearbeiten die jeweilige Kodierung und Zeilenenden erhalten. Die `Attribute VB_Name`-Zeilen und Klassen-Header (`VERSION 1.0 CLASS …`) nicht entfernen.
- Versionskommentar im Modulkopf (`Ticketautomatismus v2.0.0`, `NinjaApi v0.2.0`) bei inhaltlichen Änderungen mitpflegen.
- Neue Outlook-Objekte immer mit `Nothing`-Prüfungen und Fehlerbehandlung verwenden. Einstiegsmakros laufen aus Outlook-Regeln heraus und dürfen keine unbehandelten Fehler oder Dialoge werfen.
- Konfiguration steht in `TicketConfig.bas` (`FOLDER_TICKETS`, `FOLDER_ARCHIVE`, `AllowedSender()`, `REGEX_*`, `ENABLE_OFFLINE_FALLBACK`). Ticketnummern kommen aus dem Betreff (`[example] (#1234) …`) und aus Ordnernamen (`#1234 (Betreff)`).

## Sicherheit

- Installationsdaten nur über `LocalSettings.bas` aus `%LOCALAPPDATA%\NinjaRMM-Outlook-Classic\settings.ini` lesen; Vorlage `config/settings.example.ini`. `.local/` nie veröffentlichen. Copyright-/Lizenzhinweise und die Git-Historie nicht als beiläufige Textbereinigung verändern. Siehe `PUBLICATION.md`.

- Client-ID und Secret liegen **nur** im Windows Credential Manager, nie im Code, in Outlook-Items, Logs, Fehlermeldungen oder Testausgaben. Keine Client-ID, Tokens oder Secrets committen.
- Das Verzeichnis `outlook/` enthält keine Zugangsdaten; so belassen. Tests legen nur eindeutige, temporäre Credential-Einträge (`NinjaRMM-Outlook-Classic/UnitTest/...`) an und löschen sie wieder.

## Voraussetzungen zur Laufzeit

- Keine zusätzlichen VBA-Verweise nötig (Regex, HTTP, JSON, Dictionary sind spät gebunden); alle Module müssen importiert sein, da sie sich gegenseitig referenzieren.
- Zugriff auf die NinjaRMM-API: M2M-Anwendung („API Services“, Grant „Client Credentials“, Scope `monitoring`) und der Credential-Manager-Eintrag (generische Anmeldeinformation) wie oben beschrieben. Kein Browser nötig.
- Outlook-Regel „Skript ausführen“ → `RunEmailRuleScript`, als letzte Regel der Liste.
