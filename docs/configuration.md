# Lokale Konfiguration

Der Repository-Stand enthält ausschließlich neutrale Beispiele. Für eine bestehende Installation müssen vor dem Import die passenden lokalen Werte eingerichtet werden; die Beispieldomain `.invalid` ist kein produktiver API-Endpunkt.

## Ablage und Format

1. `config/settings.example.ini` als Vorlage nehmen.
2. Die angepasste Datei unter `%LOCALAPPDATA%\NinjaRMM-Outlook-Classic\settings.ini` ablegen. Der neue VBA-Baustein `outlook/LocalSettings.bas` liest diese Datei ohne Änderung der Registry oder Outlook-Sicherheitseinstellungen.
3. Bei nicht-ASCII-Zeichen, etwa im Absendernamen, als **UTF-16 LE mit BOM** speichern. Das Beispiel selbst ist ASCII. Die Windows-INI-API erkennt Unicode anhand des BOM; UTF-8 mit Umlauten ist für diese Datei ungeeignet.
4. **Keine Zugangsdaten eintragen.** Client-ID und Secret bleiben ausschließlich im Windows Credential Manager. `CredentialTarget` ist nur der Name des dortigen generischen Eintrags.

| Schlüssel | Bedeutung |
|---|---|
| `ApiBaseUrl` | Eigene NinjaOne-API-Basis mit HTTPS, abschließendem Slash, ohne Benutzerinformationen, Query oder Fragment |
| `CredentialTarget` | Name des vorhandenen generischen Credential-Manager-Eintrags; der Standard ist `NinjaRMM-Outlook-Classic/API` |
| `AllowedSender` | Vollständiger Anzeigename der Ticket-Benachrichtigungen; kein SMTP-Adressfilter |
| `TicketSubjectPattern` | Regulärer Ausdruck mit zwei Gruppen: Ticketnummer und restlicher Betreff |
| `TicketPrefixPattern` | Zu verkürzendes `TICKET (...) / `-Präfix; Gruppen enthalten Abteilung und Trenner |
| `OrganizationPrefixPattern` | Optional zu entfernendes Organisationspräfix in freien Betreffzeilen |
| `ExcludedDomains` | Durch Komma getrennte Ausschlussdomains |
| `AllTicketsBoard` | Positive numerische ID des Boards für Ticketabfragen |

Die fünf Matching-Stufen und der Offline-Fallback bleiben in `TicketConfig.bas` konfiguriert. Der Token-Endpunkt ist weiterhin der öffentliche EU-Endpunkt in `NinjaAuth.bas`; für andere Regionen muss er passend zur NinjaOne-Anwendung ausgewählt werden. Ein `TokenUrl`-INI-Schlüssel wird nicht ausgewertet.

## Bestehende Installation übernehmen

Vor dem Import des bereinigten Codes die bisherigen nicht geheimen Einstellungen lokal sichern. Eigene Sicherungsdateien sind installationsspezifisch, gehören nicht in Releases oder ZIPs und dürfen nicht erzwungen zu Git hinzugefügt werden. Die Vorlage muss bewusst an den oben genannten Laufzeitpfad kopiert, angepasst und vor einem Produktivimport geprüft werden. Dieses Repository liefert keine privaten Einstellungen mit.

Danach `LocalSettings.bas` zusätzlich importieren und `NinjaAuth`, `TicketConfig`, `TicketMatcher` sowie alle geänderten Testmodule aus derselben Revision ersetzen. `Debuggen > Kompilieren` ausführen. Zugangsdaten nicht umbenennen oder neu anlegen, wenn `CredentialTarget` auf den vorhandenen Eintrag zeigt. Bei geänderter Konfiguration Outlook neu starten, damit Token- und Indexcaches zur neuen Einstellung passen.

Ein Self-Test verwendet pro Test neutrale Defaults und synthetische temporäre Dateien. Die echten Einstellungen werden nach Ende des Tests wieder gelesen. `RunLiveTests` ist ausdrücklich getrennt; dessen Ticket-IDs und externe Referenz werden über die optionalen `Live*`-Schlüssel festgelegt. Die Vorlagennummern aus Unit-Tests sind keine Vorgaben für Live-Abfragen.

## PowerShell-Smoketest

Der lesende Smoketest verlangt API-Basis und Credential-Ziel explizit. Er verwendet keinen voreingestellten Tenant:

```powershell
.\tools\ninja-smoketest.ps1 -ApiBase 'https://tenant.example.invalid/v2/' -CredentialTarget 'NinjaRMM-Outlook-Classic/API'
```

Die URL im Beispiel muss durch die eigene Basis ersetzt werden. Offene/geschlossene Ticket-IDs werden bei Bedarf als Parameter angegeben; keine IDs aus internen Fällen in das Repository übernehmen.
