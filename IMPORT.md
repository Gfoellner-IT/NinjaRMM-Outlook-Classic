# Import M6 bis M8 in Outlook und gemeinsamer Test

**Geltungsbereich: veröffentlichter VBA-Quellstand dieser Revision.** Diese Anleitung beschreibt das Upgrade eines älteren Basisprojekts um Matcher, Audit, Sicherheitsstufen und Dialoge. Die lokale Konfiguration muss zur Installation passen. Vollständige Outlook-Kompilierung und fachliche Abnahme sind vor Produktivbetrieb erforderlich. Die VSTO-Migration ist eine separate Aufgabe der [Roadmap](ROADMAP.md).

Diese Anleitung beschreibt den integrierten Entwicklungsstand. Sie geht vom Stand aus, den du zuletzt in Outlook hattest (M0 bis M5), und enthält alles, was seitdem dazugekommen ist: Matcher und Ticketindex (M6), Audit-Log, Bericht, manuelle Zuordnungen und gelernte Regeln (M7), Sicherheitsstufen und Bestätigungsdialog (M8).

**Der Fremdmail-Matcher steht zunächst auf nur Logging:** Alle fünf Regelstufen sind 1. Die bestehende Sortierung von Ticket-Benachrichtigungen und die Archivierung können weiterhin Mails beziehungsweise Ordner verschieben. Stufe 2 oder 3 aktiviert Aktionen ausschließlich für den Matcher.

Aktuelle Testergebnisse stehen im [Changelog](CHANGELOG.md). Der Headless-Lauf prüft die dort genannten Logik- und Formularfälle in Excel, aber weder Outlook-Typen noch den tatsächlichen Outlook-Dialogbetrieb. Deshalb ist Schritt 3 (Kompilieren) erforderlich; bei Fehlern Modul, Zeile und Meldung notieren.

## 0. Vorbereitung

**Nach der Veröffentlichungsbereinigung:** Vor dem Import die [lokale Konfiguration](docs/configuration.md) anlegen. `outlook\LocalSettings.bas` zusätzlich importieren und `NinjaAuth`, `TicketConfig`, `TicketMatcher` sowie die aktualisierten Testmodule ersetzen. Neutrale Defaults ersetzen die frühere Betriebs-Konfiguration; die vorhandene Outlook-Installation wird vom Repository nicht automatisch aktualisiert.

1. Outlook schließen.
2. Das VBA-Projekt sichern:
   ```powershell
   Copy-Item "$env:APPDATA\Microsoft\Outlook\VbaProject.OTM" "$env:APPDATA\Microsoft\Outlook\VbaProject.OTM.bak-$(Get-Date -Format yyyyMMdd)"
   ```
3. Outlook starten, `Alt+F11` (VBA-Editor). Die Quelldateien liegen im selbst gewählten Repository-Checkout.

## 1. Alte Module entfernen

Im Projekt-Explorer je Modul: Rechtsklick, *Entfernen…*, die Frage „Exportieren?“ mit **Nein** beantworten. `ThisOutlookSession` und alle nicht genannten Module bleiben.

**Aus `outlook\` (12 geänderte):** `IIncomingMail`, `INinjaTransport`, `ITicketFolder`, `NinjaTicketLogic`, `NinjaTransport`, `OutlookIncomingMail`, `OutlookTicketFolder`, `TicketConfig`, `TicketLog`, `TicketMailWatcher`, `TicketService`, `Ticketautomatismus`.

**Aus `Tests\` (3 Fakes, Pflicht, falls die Tests im Projekt sind):** `FakeNinjaTransport`, `FakeMail`, `FakeTicketFolder`. Sie implementieren Schnittstellen, die sich geändert haben. Bleiben die alten Versionen, meldet der Compiler z. B. „In diesem Objektmodul muß 'HttpPostJson' für die Schnittstelle 'INinjaTransport' implementiert sein“.

Wird ein Modul nicht vorher entfernt, legt der Import eine Kopie `Name1` an und nichts kompiliert.

## 2. Importieren

*Datei > Datei importieren…* (`Strg+M`), eine Datei nach der anderen. Reihenfolge egal.

**Aus `outlook` (36 Dateien):**

| Neu (24) | Geändert (12) |
|---|---|
| `AuditLogic.bas`, `AuditReport.bas`, `ClipboardHelper.bas`, `FileAuditSink.cls`, `FormMatchPrompt.cls`, `FileRejectionStore.cls`, `GermanText.bas`, `IAuditSink.cls`, `IMatchPrompt.cls`, `IRejectionStore.cls`, `ITicketIndex.cls`, `LearnedRules.bas`, `ManualScan.bas`, `MatchLevels.bas`, `MatchPromptText.bas`, `MatchProposal.cls`, `MatchProposals.bas`, `MatchRejections.bas`, `MatchReview.bas`, `MatchTicket.cls`, `MsgBoxMatchPrompt.cls`, `NinjaTicketIndex.cls`, `TicketAudit.bas`, `TicketMatcher.bas` | `IIncomingMail.cls`, `INinjaTransport.cls`, `ITicketFolder.cls`, `NinjaTicketLogic.bas`, `NinjaTransport.cls`, `OutlookIncomingMail.cls`, `OutlookTicketFolder.cls`, `TicketConfig.bas`, `TicketLog.bas`, `TicketMailWatcher.cls`, `TicketService.bas`, `Ticketautomatismus.bas` |

**Aus `Tests` (3 Dateien):** `FakeNinjaTransport.cls`, `FakeMail.cls`, `FakeTicketFolder.cls`.

**Aus `forms` (1 Formular, empfohlen):** `ReviewForm.frm` (die Datei `ReviewForm.frx` muss im selben Ordner liegen). Das ist der Bestätigungsdialog mit Liste (Dialog B). Fehlt das Formular im Projekt oder lässt es sich nicht importieren, zeigt Outlook die Vorschläge stattdessen mit einzelnen Fenstern (Dialog A), und es kommt ein Hinweis ins Log.

## 3. Kompilieren

*Debuggen > Kompilieren von Project1*. Kein Fehler: weiter. Sonst bitte Modul, Zeile und Meldung notieren oder Screenshot.

## 4. Erster Lauf (nur lesen, nichts wird verschoben)

Im Direktfenster (`Strg+G`), jeweils mit Enter:

```
Ticketautomatismus.RunMatchDryRun
```

Ergebnis in `%TEMP%\ninja_ticket_log.txt` und `%LOCALAPPDATA%\NinjaRMM-Outlook-Classic\audit\match_audit_<Jahr>-<Monat>.jsonl`. Seit dem Datenschutz-Fix steht bei Mails ohne Treffer nur ein Hash des Betreffs im Log (`subject_hash=...`), der Betreff selbst nur bei Treffer oder Mehrdeutigkeit. Alte Audit-Dateien von früheren Läufen enthalten noch Betreffe: Du kannst sie löschen.

## 5. Bestätigungsdialog ausprobieren (Stufe 2)

1. In `TicketConfig` eine Regel auf Stufe 2 setzen, zum Beispiel `MATCH_LEVEL_SUBJECT = 2` (Stufen: 0 aus, 1 nur Logging, 2 Bestätigen, 3 automatisch; die globale Obergrenze `MATCH_LEVEL_MAX` steht auf 3). Speichern.
2. Im Direktfenster `Ticketautomatismus.RunMatchReview` eingeben. Es werden die Mails der letzten 14 Tage geprüft. Mit importiertem Formular (Standard, `REVIEW_DIALOG = 2`) erscheint eine **Liste mit Ankreuzfeldern**: Mails ankreuzen, dann *Ausgewählte übernehmen* (verschiebt in den Ticketordner) oder *Ausgewählte ablehnen* (wird nicht wieder vorgeschlagen); *Alle auswählen* markiert alles; was beim *Schließen* übrig ist, bleibt ein Vorschlag. Ohne Formular (oder mit `REVIEW_DIALOG = 1`) erscheint je Vorschlag ein Fenster: **Ja** = übernehmen, **Nein** = ablehnen, **Abbrechen** = Prüfung beenden (die restlichen bleiben).
3. Prüfe: Landet eine bestätigte Mail im richtigen Ticketordner (`#<Nummer> (<Betreff>)`)? Taucht eine abgelehnte Mail bei einem zweiten Aufruf nicht wieder auf?
4. Die Hinweisfrage „Es gibt N neue Vorschläge für die Ticketzuordnung. Diese jetzt prüfen?“ kommt, wenn der Watcher nach einer neuen Mail Vorschläge hat (höchstens alle 15 Minuten). Beim Start von Outlook fragt der Durchlauf nie, die Frage kommt dann mit der nächsten neuen Mail oder du startest `RunMatchReview`.

Für den Test im Test-PST (`NinjaTest`) gilt weiter: Die Integrationstests laufen nur dort, nie im produktiven Postfach.

## 6. Menüleisten-Knopf einrichten

*Datei > Optionen > Menüband anpassen*:
1. Rechts eine neue Registerkarte oder Gruppe anlegen (z. B. „Tickets“).
2. Links bei „Befehle auswählen“ den Eintrag **Makros** wählen.
3. `Project1.Ticketautomatismus.RunMatchReview` markieren, *Hinzufügen*, dann *Umbenennen* (z. B. „Vorschläge prüfen“) und ein Symbol wählen.
4. Auf dieselbe Art kannst du hinzufügen: `RunMatchRulesReview` („Regeln prüfen“), `RunArchiveRuleWithProgress` und `RunEmailRule`.

Nur Makros **ohne Parameter** erscheinen in der Liste. `RunMatchDryRun` (optionaler Parameter) und `RunEmailRuleScript` (Regelaktion) gehören nicht dazu.

## 7. Bericht und Regelrunde mit Claude

`Ticketautomatismus.RunMatchRulesReview` (oder der Knopf) sucht zuerst nach von Hand verschobenen Mails, schreibt dann den Bericht nach `%LOCALAPPDATA%\NinjaRMM-Outlook-Classic\reports\match_review_<Datum>.md` und legt einen Einleitungssatz in die Zwischenablage. In einer Claude-Sitzung im Repo-Ordner den Satz einfügen oder `/review-rules` tippen. Der Bericht enthält Zahlen, Domains, Ticketnummern und Betreffstämme von Mails mit Ticketbezug. Er übernimmt keine separaten Absendernamen oder Mailtexte; Betreffstämme können jedoch personenbezogene Angaben enthalten. Bericht vor Weitergabe prüfen.

## 8. Gelernte Regeln (später)

Regeln stehen in `%LOCALAPPDATA%\NinjaRMM-Outlook-Classic\audit\match_rules.json` (die Datei gibt es nicht, solange du sie nicht anlegst). Der Bericht schlägt Regeln als JSON-Zeilen vor, zunächst mit `"enabled":false`. Zum Aktivieren `enabled` auf `true` setzen. Gelernte Regeln laufen als letzte Regel, auf der Stufe `MATCH_LEVEL_LEARNED` (Standard 1: nur Logging) und beschreiben eine Suche (Domain und Betreffstamm führen zum einzigen offenen Ticket mit diesem Stamm im Titel), keine Ticketnummern. Die Datei wird bei jedem Durchlauf und bei jedem Probelauf neu gelesen.

## 9. Gut zu wissen

- **Wiederholte Einträge:** `RunEmailRule` und der Durchlauf beim Start prüfen die Fremdmails der letzten 14 Tage erneut und schreiben dabei jedes Mal Audit-Zeilen. Dieselbe Mail kann also mehrfach im Log stehen. Der Bericht zählt dieselbe automatische Entscheidung zu derselben Mail nur einmal (Schlüssel `msg`), die Erkennung manueller Zuordnungen legt jede Mail nur einmal an. Wer das vermeiden will, setzt `MATCH_SWEEP_DAYS = 0` in `TicketConfig`.
- **Manuelle Zuordnungen:** Beim Start (und vor dem Bericht) werden Mails, die du von Hand in einen Ticketordner gezogen hast (letzte 30 Tage), einmal als `manual_assignment` im Audit-Log festgehalten, mit dem Ergebnis, was der Matcher getan hätte (`agree`, `differ`, `missed`, `excluded`, `unchecked`).
- **Dauer beim Start:** Der Start-Durchlauf lädt einmal die offenen Tickets aus NinjaOne (etwa 3 Seiten) und sucht in allen Ticketordnern nach neuen, von Hand verschobenen Mails. Mit vielen Ticketordnern kann das einige Sekunden dauern. Ist es zu langsam, sag Bescheid; Stellschrauben sind `MANUAL_SCAN_DAYS`, `MATCH_SWEEP_DAYS` und `INDEX_CACHE_MINUTES` in `TicketConfig`.
- **Audit-Log abschalten:** `ENABLE_AUDIT_LOG = False`. **Alle Regeln aus:** `MATCH_LEVEL_MAX = 0`.
- **Signatur:** Nach dem Import ist das VBA-Projekt nicht mehr signiert. Bei Bedarf neu signieren (*Extras > Digitale Signatur*, eigenes freigegebenes Code-Signing-Zertifikat).
- **Zurück zum alten Stand:** Outlook schließen, `VbaProject.OTM.bak-<Datum>` über `VbaProject.OTM` kopieren.

## 10. Alle Tests vollständig in Outlook importieren

Für den vollständigen Self-Test müssen die Testmodule zum importierten Produktionscode passen. Nur die drei Fakes aus Schritt 2 zu ersetzen aktualisiert den TestRunner und die übrigen Tests noch nicht.

1. VBA-Projekt wie in Schritt 0 sichern. Im Projekt-Explorer alle bereits vorhandenen Module aus der folgenden Liste entfernen; bei Bedarf vorher exportieren. Produktionsmodule und `ThisOutlookSession` bleiben erhalten.
2. Jede der folgenden **39 Dateien aus `Tests\`** über *Datei > Datei importieren…* (`Strg+M`) importieren. Der Ordner `Tests\fixtures` enthält JSON-Quelldaten und wird nicht im VBA-Editor importiert; `Fixtures.bas` enthält die vorbereiteten Daten.
3. Zusätzlich `outlook\OutlookTestGuard.bas` importieren. Falls das Modul bereits vorhanden ist, zuerst entfernen, damit keine Kopie `OutlookTestGuard1` entsteht.
4. `ReviewForm` aus Schritt 2 muss vorhanden sein (oder gemäß Schritt 12 manuell angelegt werden), weil `ReviewFormTests` das Formular instanziiert. Der MsgBox-Fallback für den normalen Betrieb ersetzt diese Testvoraussetzung nicht.
5. *Debuggen > Kompilieren von Project1* ausführen. Erst wenn das ohne Fehler durchläuft, mit dem Self-Test fortfahren. Bei Namen wie `TestRunner1`, `FakeMail1` oder anderen Importduplikaten zuerst die doppelten Module bereinigen.

| Datei in `Tests\` | Modulname im Projekt-Explorer |
|---|---|
| `AuditReportTests.bas` | `AuditReportTests` |
| `AuditTests.bas` | `AuditTests` |
| `FakeAuditSink.cls` | `FakeAuditSink` |
| `FakeMail.cls` | `FakeMail` |
| `FakeMatchPrompt.cls` | `FakeMatchPrompt` |
| `FakeNinjaTransport.cls` | `FakeNinjaTransport` |
| `FakeProgressSink.cls` | `FakeProgressSink` |
| `FakeRejectionStore.cls` | `FakeRejectionStore` |
| `FakeTicketFolder.cls` | `FakeTicketFolder` |
| `FakeTicketIndex.cls` | `FakeTicketIndex` |
| `FakeTicketStore.cls` | `FakeTicketStore` |
| `FakeTokenFetcher.cls` | `FakeTokenFetcher` |
| `Fixtures.bas` | `Fixtures` |
| `IntegrationRunner.bas` | `IntegrationRunner` |
| `JsonLiteTests.bas` | `JsonLiteTests` |
| `LearnedRulesTests.bas` | `LearnedRulesTests` |
| `LiveTests.bas` | `LiveTests` |
| `LocalSettingsTests.bas` | `LocalSettingsTests` |
| `ManualScanTests.bas` | `ManualScanTests` |
| `MatchLevelsTests.bas` | `MatchLevelsTests` |
| `MatchPromptTextTests.bas` | `MatchPromptTextTests` |
| `MatchReviewTests.bas` | `MatchReviewTests` |
| `NinjaApiTests.bas` | `NinjaApiTests` |
| `NinjaAuthTests.bas` | `NinjaAuthTests` |
| `NinjaCredentialStoreTests.bas` | `NinjaCredentialStoreTests` |
| `NinjaTicketIndexTests.bas` | `NinjaTicketIndexTests` |
| `NinjaTicketLogicTests.bas` | `NinjaTicketLogicTests` |
| `NinjaTokenLogicTests.bas` | `NinjaTokenLogicTests` |
| `OutlookAdapterTests.bas` | `OutlookAdapterTests` |
| `ReviewFormTests.bas` | `ReviewFormTests` |
| `SelfTests.bas` | `SelfTests` |
| `TestAssert.bas` | `TestAssert` |
| `TestRunner.bas` | `TestRunner` |
| `TicketDecisionTests.bas` | `TicketDecisionTests` |
| `TicketMatcherTests.bas` | `TicketMatcherTests` |
| `TicketMatchServiceTests.bas` | `TicketMatchServiceTests` |
| `TicketParsingTests.bas` | `TicketParsingTests` |
| `TicketServiceTests.bas` | `TicketServiceTests` |
| `TimeUtilTests.bas` | `TimeUtilTests` |

**Self-Test:** Im Modul `TestRunner` den Cursor in `RunAllTests` setzen und `F5` drücken. Erwartet werden **303 bestandene Tests, 0 Fehler**. Ergebnis im Direktfenster (`Strg+G`) und in `%TEMP%\ninja_test_results.txt`. Weiterhin nur 117 Tests deutet auf einen alten oder unvollständig aktualisierten Runner hin. Die Tests arbeiten mit Fakes; Protokollmeldungen über simulierte Fehler gehören zu den Negativfällen.

**Separate Outlook-Integrationstests:** `IntegrationRunner.RunIntegrationTests` ausschließlich mit der lokalen Test-PST `NinjaTest` ausführen. Es sind **21 Tests** registriert; dies ist ein eigener Lauf und nicht Teil der 298 Unit-Tests. Ergebnis in `%TEMP%\ninja_integration_results.txt`.

`TestRunner.RunLiveTests` ist ein separater Opt-in-Aufruf mit echter NinjaOne-Verbindung und gehört nicht zum Self-Test.
## 11. Was bewusst noch fehlt

- Noch keine echten Entscheidungen im Audit-Log: Der Bericht wird erst durch Bestätigen, Ablehnen und manuelles Zuordnen aussagekräftig.
- Die Dialoge selbst (Formular und Fenster) und alles, was Outlook-Typen benutzt, wurden hier nicht in Outlook ausgeführt. Die Logik dahinter ist headless getestet, das Formular wurde in Excel erzeugt und mit simuliertem Anklicken geprüft.

## 12. Falls sich ReviewForm.frm nicht importieren lässt

Das Formular zu Fuß anlegen (dann bleibt der Rest unverändert, die Logik steht im Code):
1. Im VBA-Editor *Einfügen > UserForm*, im Eigenschaftenfenster `(Name)` = `ReviewForm` setzen.
2. Sechs Steuerelemente anlegen und genau so benennen: `lblInfo` (Bezeichnung), `lstProposals` (Listenfeld, `ListStyle` = 1 *fmListStyleOption*, `MultiSelect` = 1 *fmMultiSelectMulti*), `btnSelectAll`, `btnAccept`, `btnReject`, `btnClose` (Schaltflächen). Die Beschriftungen setzt der Code beim Öffnen.
3. Doppelklick auf das Formular (Codefenster), den gesamten Inhalt durch den Text aus `forms\ReviewForm.code.txt` ersetzen.
4. Neu erzeugen lässt sich die Datei jederzeit mit `tools\New-ReviewForm.ps1`.

Wer lieber ganz ohne Formular arbeitet: `REVIEW_DIALOG = 1` in `TicketConfig` (einzelne Fenster).
