# Provider-Inventar (Auslagerung 30% der Systeme)

Sammelt pro Server alle Angaben, die ein neuer Provider fuer Sizing und Machbarkeit braucht,
und fasst sie zu Excel-Dateien zusammen. Ein Scan pro Server liefert alles, auch den
Security-Teil. Der Security-Teil wird erst beim Zusammenfuehren herausgefiltert, es gibt keinen
zweiten Scan.

```
[Server 1..90]                          [Ihr Rechner]
Get-sqmProviderInventory.ps1   ->  JSON je Server  ->  Merge-sqmProviderInventory.ps1
(lokal, standalone, nur lesend)    (Rohdaten)          -> Provider_Inventory_<Datum>.xlsx   (ohne Security)
                                                       -> Security_Inventory_<Datum>.xlsx   (nur -IncludeSecurity)
                                                       -> Pruefbericht_<Datum>.txt
```

## Schritt 1: Scan auf jedem Server

Datei `Get-sqmProviderInventory.ps1` auf den Server kopieren (Share, Mail) und dort starten.
Kein Modul und keine Installation noetig (kein dbatools). Windows PowerShell 5.1 genuegt.
Als lokaler Administrator starten, sonst fehlen Teile der OS-Daten (Volumes, Dienste, Energieplan).

```powershell
.\Get-sqmProviderInventory.ps1 -OutputPath D:\Inventory
```

Mit Anwendung und Obergruppe (optional, kann auch spaeter in der CSV stehen):

```powershell
.\Get-sqmProviderInventory.ps1 -OutputPath D:\Inventory -Application "SAP|DWH" -ParentGroup "Finanzen"
```

| Parameter | Bedeutung |
|---|---|
| `-SqlInstance` | Instanz(en), Standard: alle SQL-Dienste des Servers. Bei Cluster-Instanzen (FCI) den virtuellen Netzwerknamen angeben. |
| `-SqlCredential` | SQL-Login, falls Windows-Anmeldung nicht geht. |
| `-IncludeAgentJobs` | Agent-Jobs im Detail (Name, Owner, Zeitplan, letzter Lauf, externe Referenzen). Ohne den Schalter nur die Anzahl. |
| `-SkipSecurity` | Keine Logins, DB-User, Rollen, Rechte erfassen. |
| `-SkipPerDatabase` | Nicht in jede DB wechseln (kein belegter Platz, keine DB-User, keine Cross-DB-Referenzen). Nur bei extrem vielen DBs. |
| `-SampleSeconds` | Dauer der Lastmessung, Standard 10 s, 0 = keine. |
| `-BackupHistoryDays` | Zeitraum fuer Backupdauer und -groesse, Standard 30. |

Das Ergebnis ist **eine JSON-Datei je Server** (`<Server>_<Datum>.json`). Diese Dateien sammeln
Sie in einem Ordner.

Der Scan liest nur Metadaten (Systemsichten, DMVs, msdb), keine Anwenderdaten. Befehle von
Agent-Job-Schritten werden **nicht** gespeichert (koennten Kennwoerter enthalten), nur erkannte
externe Bezuege (UNC-Pfade, Linked Server, OPENQUERY). Nicht lesbare Bereiche (Rechte,
alte SQL-Version) brechen den Scan nicht ab, sie stehen in der Liste `errors` der JSON-Datei und
im Pruefbericht.

### Was erfasst wird

| Bereich | Inhalt |
|---|---|
| OS und Ausbau | Betriebssystem, Version, Build, Edition, Installationsart, Virtuell/Physisch, Hersteller/Modell, Sockets/Kerne/logische CPUs, RAM, Pagefile, Energieplan, letzter Boot, letzter Hotfix, IPs, Zeitzone |
| Platten | je Volume Groesse, belegt, frei, Dateisystem, Blockgroesse; physische Datentraeger; SQL-Dateien, IOPS, MB/s und Latenz je Volume |
| SQL Server | Version, CU, Edition, Collation, sichtbare Kerne (Lizenz), Max/Min Memory, MAXDOP, Cluster, AlwaysOn, TCP-Port, Endpoints, Dienste |
| Last | CPU avg/p95/max (letzte ca. 4 h aus der DMV), Page Life Expectancy, Memory, Batch Requests/s, Transaktionen/s, Verbindungen |
| Datenbanken | Daten-/Logdatei-Groesse, belegt, Compat-Level, Recovery Model, Collation, Owner, TDE, Wachstum MB/Monat, Replikation, CLR, Service Broker; tempdb ja, andere Systemdatenbanken nein |
| Backup | je DB und Typ letzte Zeit, Dauer avg/max, Groesse, komprimierte Groesse, Anzahl, Ziel |
| Abhaengigkeiten | Linked Server, Cross-DB- und Cross-Server-Referenzen, Replikation, externe Bezuege in Agent-Jobs |
| HA und DR | Availability Groups, Replikate, Listener, Cluster-Knoten, Log Shipping, Mirroring |
| Komponenten | SSIS/SSRS/SSAS-Dienste, Fremdsoftware-Dienste (Heuristik), Credentials, Proxies, Zertifikate, TDE, Database Mail |
| Security (immer erfasst) | Logins (Windows/Gruppe/SQL), Serverrechte, DB-User (inkl. verwaist), eigene DB-Rollen, Rollenmitglieder, explizite DB-Rechte |

## Schritt 2: Mapping-CSV (Anwendung, Obergruppe, ...)

Anwendung und Obergruppe pflegen Sie auf Serverebene in einer CSV (Trennzeichen `;`, UTF-8).
Vorlage aus den gescannten Servern erzeugen (vorhandene Zeilen bleiben erhalten):

```powershell
.\Merge-sqmProviderInventory.ps1 -InputPath D:\Inventory -MappingPath D:\Inventory\Mapping.csv -NewMappingTemplate
```

Spalten: `Server;Anwendung;Obergruppe;Schutzbedarf;Kritikalitaet;RPO;RTO;Wartungsfenster;Ansprechpartner;Bemerkung`.
Mehrere Anwendungen pro Server: mit `|` trennen. Die CSV gewinnt gegenueber Werten aus dem Scan.
Aenderungen an der Zuordnung brauchen **keinen neuen Scan**, nur einen neuen Merge.

Schutzbedarf, RPO/RTO, Wartungsfenster und Ansprechpartner lassen sich nicht scannen, sie
kommen nur aus dieser CSV.

## Schritt 3: Zusammenfuehren

```powershell
# Fuer den Provider
.\Merge-sqmProviderInventory.ps1 -InputPath D:\Inventory -MappingPath D:\Inventory\Mapping.csv

# Mit Security (Schritt 2, NICHT an den Provider)
.\Merge-sqmProviderInventory.ps1 -InputPath D:\Inventory -MappingPath D:\Inventory\Mapping.csv -IncludeSecurity

# Servernamen und Domaene durch Kennungen ersetzen (SRV001 ...)
.\Merge-sqmProviderInventory.ps1 -InputPath D:\Inventory -MappingPath D:\Inventory\Mapping.csv -Anonymize
```

`Provider_Inventory_<Datum>.xlsx` enthaelt keine Logins, User oder Rollen. Blaetter:
Zusammenfassung (je Obergruppe und gesamt, mit Anteil in %), Server, Platten, Datentraeger,
SQL-Instanzen, Datenbanken, Dateien, Abhaengigkeiten, HA-DR, Komponenten, Agent-Jobs (falls erfasst).

`Security_Inventory_<Datum>.xlsx` (nur mit `-IncludeSecurity`): Zusammenfassung, Logins,
Server-Rechte, DB-User, DB-Rollen, Rollenmitglieder, DB-Berechtigungen. Standardmaessig sind
System-Logins (`sa`, `NT SERVICE\*`, `NT AUTHORITY\*`) und sysadmin-Mitglieder ausgefiltert
(`-IncludeSystemAndSysadmin` hebt das auf). Die Zusammenfassung zeigt je Instanz Windows-Logins,
AD-Gruppen und SQL-Logins. Das ist die Grundlage fuer die Frage, ob die AD-Anbindung beim
neuen Provider noetig ist.

Mehrere Scans desselben Servers: der neueste gewinnt, die aelteren werden im Pruefbericht genannt.

## Pruefbericht

`Pruefbericht_<Datum>.txt` listet vor der Weitergabe:

- `FEHLT`: Server im Mapping, aber kein Scan
- `ZUORDNUNG`: Anwendung oder Obergruppe fehlt (in den Blaettern als `NICHT ZUGEORDNET`), Server nicht im Mapping
- `ALT`: Scan aelter als `-MaxAgeDays` (Standard 30)
- `FEHLER`: Instanz nicht erreichbar, Datei nicht lesbar
- `LUECKE`: Bereiche, die der Scan nicht lesen konnte
- `HINWEIS`: z.B. SQL-Dienst laeuft erst kurz, Lastwerte sind dann nicht belastbar

## Grenzen und Hinweise

- **Last ist eine Momentaufnahme.** CPU stammt aus der Historie der letzten ca. 4 Stunden, IOPS und
  Latenz sind Mittelwerte seit dem SQL-Start, Batch Requests/s und Transaktionen/s zusaetzlich aus
  einer kurzen Messung. Spitzen (Monatsabschluss) sind damit nicht sicher erfasst. Die Spalte
  `Uptime h` zeigt, wie belastbar die Werte sind.
- **Wachstum** wird aus der monatlichen Groesse der Full-Backups der letzten 12 Monate errechnet
  (belegte Daten, nicht Dateigroesse). Ohne Backup-Historie bleibt die Spalte leer.
- **Lizenz-Kerne** sind die sichtbaren Online-Scheduler der Instanz. Lizenzmodell und Software
  Assurance lassen sich nicht auslesen.
- **Fremdsoftware-Dienste** sind eine Heuristik (automatische Dienste ausserhalb von Windows und
  Microsoft-Pfaden), nicht jede Monitoring- oder Backup-Software wird erkannt.
- **`-Anonymize`** ersetzt Servernamen, Instanznamen und Domaene in allen Textzellen. Datenbank-,
  Anwendungs- und Jobnamen bleiben, ebenso Pfade und Linked-Server-Namen, die einen anderen Text
  enthalten. Vor der Weitergabe die Excel-Datei durchsehen.
- Anonymize-Schluessel (`Anonymize_Zuordnung_*.csv`), Security-Datei und die Roh-JSON-Dateien gehen
  nicht an den Provider.
- Nicht erfasst: Firewall-Regeln, DNS-Eintraege, Quell-IPs der Clients, Zertifikate im OS-Speicher,
  SSIS-Paketinhalt. Diese Angaben muessen aus dem Netzwerk- und Betriebsumfeld kommen.

## Tests

```powershell
pwsh -NoProfile -Command "Invoke-Pester .\tests\Merge.Tests.ps1"
```

Die Tests laufen gegen einen synthetischen Scan in `tests\fixtures`.
