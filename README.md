
# CygateScript.ps1



Ett skript för att underlätta arbetet med uppgradering av Visma´s produkt Personec P HRM.
Skriptet har funktioner som underlättar och sparar tid samt gör att uppgraderingsstegen oavsett konsult blir standardiserade.



## Start av skriptet - Första gången
Det finns två sätt att plocka ner skriptet beroende på hur kunden tillåter internetåtkomst.
-	Nedladdning via powershell
```Powershell
Invoke-WebRequest 'https://github.com/Dambergc/Vismascript/releases/latest/download/CygateScript.ps1' -OutFile D:\Visma\Install\Backup\CygateScript.ps1 -Verbose
```
Gå till Github till adressen https://github.com/dambergc/vismascript och går till Releases och ladda ner filen CygateScript.ps1 under senaste Release.

![Välj rätt i github.](https://github.com/DambergC/VismaScript/blob/fc34252a991c24617acf65b4885b48fc6f5ca5e3/ReadMEPictures/github.png)

Är det första gången så behöver du lägga till lite värden, spara BIGRAM och sedan skapa en backupmapp.

![Ändra versionsnummer](https://github.com/DambergC/VismaScript/blob/d49564e9d1f41395cf79a763b86de1d6fb58c7e5/ReadMEPictures/edit_versions.png)
- VPBS Version 1 - Används i sökvägen för att ladda ner VPBS från Visma
  - (visma-pbs.s3.eu-central-1.amazonaws.com/<ins>**251**</ins>/VPBSDownload252.zip)
- VPBS Version 2 - Används i sökvägen för att ladda ner VPBS från Visma
  - (visma-pbs.s3.eu-central-1.amazonaws.com/251/VPBSDownload<ins>**252**</ins>.zip)
- Release Version - Används i skapandet av SQL queries, pekar ut sökvägen till Vismas SQL skript. ***Ska inte förväxlas med FDN som inte hanteras i skriptet***.
- PPP Version - Används vid skapandet av SQL Queries
- PUD Version - Används vid skapandet av SQL Queries
- PFH Version - Används vid skapandet av SQL Queries

## Uppstart av skript
För att starta skriptet så starta ett eleverat Powershell fönster och sedan skriv sökvägen till skriptet

```Powershell
D:\visma\install\backup\CygateScript.ps1
```
Skulle det finnas kund som har sin installation på annan enhet en den som är default (D:\) så hanteras det av skriptet och då är sökvägen självklart samma som ovan men med ändring av enhetsbokstaven som du får ändra manuellt.

![Hur man startar skriptet.](https://github.com/DambergC/VismaScript/blob/f3c6cc2df2366acbac5a6164e5042a71ea5b0712/ReadMEPictures/CMDstart.png)

Om det skulle ha varit en uppgradering av skriptet tillgänglig så laddas den ned och sedan så kommer du få en dialogruta om att ny version av skriptet nedladdad och att du ombedes att starta om skriptet.
## Visma Services Trusted Users
För att få köra skriptet så är det samma krav som på Public Installer, du måste vara medlem i den lokala gruppen ”Visma Services Trusted Users” Hur du lägger till dig i gruppen går att göra det manuellt via Computer Management eller via powershell som du hittar i vårt uppgraderingsdokument.
## Några Prereq för att komma åt vissa funktioner.
För att komma åt funktionerna i skriptet så krävs att du väljer BIGRAM och BACKUKFOLDER under den knappen ”Bigram And BackupFolder”. Efter att du valt det så är övriga knappar öppna så att du kan komma åt underliggande funktioner.

Det BIGRAM och BackupFolder du väljer kommer följa med dig under hela tiden så länge du har skriptet igång. Skulle det stängas ner av misstag så måste du åter välja det vid uppstart.

## Logfil
Alla val man gör i skriptet skrivs i en logfil som ligger under d:\visma\install\backup. Logfilen används innan man valt BackupFolder och måste därför ligga under Visma\Backup och dessa logfiler får vid behov rensas manuellt. Det skapas en logfil för varje dag man kör skriptet.
## Funktioner
De funktioner i skriptet är framtagna för att standardisera och effektivisera arbetet för konsult vid uppgradering och felsökning av Personec P installation hos kund.

![Funktioner](https://github.com/DambergC/VismaScript/blob/d49564e9d1f41395cf79a763b86de1d6fb58c7e5/ReadMEPictures/Funktioner.png)

### FileBackup 
Filbackup är ett krav och MÅSTE köras inför varje uppgradering oavsett Major eller Minor och har man fler BIGRAM så räcker det att köra backupen en gång.

Det som kopieras undan är allt under wwwroot och programs med undantag av *.log filer som exkluderas pga storleken.

Man ska INTE avkryptera miljön innan backup då vid behov så har vi  en funktion där vi kan avkryptera backup för att komma åt värden. Detta är av säkerhet då okrypterade backuper innehåller känslig info om konton och lösenord i kundens miljö.

![Filebackup](https://github.com/DambergC/VismaScript/blob/d49564e9d1f41395cf79a763b86de1d6fb58c7e5/ReadMEPictures/Filebackup.png)

Backupen körs med hjälp av RoboCopy och loggarna för backupen sparas under BackupFolder-mappen.

## Inventering
Inventering av systemet där följande saker inventeras:

![Inventering](https://github.com/DambergC/VismaScript/blob/d49564e9d1f41395cf79a763b86de1d6fb58c7e5/ReadMEPictures/Inventory.png)

### system
***Kan köras när som, dock innan avinstallation***
-	Vad som är installerat och vilka versioner som är kopplade till Visma
-	Vilka applikationspooler som är igång och med vilka konton som dom körs med – Om det är en webserver.
-	Vilka tjänster som är igång och hur dom är konfigurerade. Ibland så kör kunden t.ex. Batchtjänsten med ett AD-konto vilket kunden behöver information om att återställa efter uppgraderingen.

### Password
***Kan bara köras efter att backup är genomförd***

Lösenorden till vissa utpekade konton som vi ibland kan ha behov att ha tillgång till inventeras från den backup som körts. Att detta ska fungera så måste följande vara uppfyllt.
-	Backup körd och BackupFolder vald
-	Avkryptering av Backupen gjord

### Settings
***Kan bara köras efter att backup är genomförd***

Följande inventeras just nu av Settings
-	Om värdet useSSO är true eller false
-	Om värdet Multitenant är true och false
-	Om det finns en license.json i backupen

Om det är mer värden som kan vara nytta för oss tekniker under en uppgradering så är det enkelt att lägga till mer saker att inventera.

## SQL Queries

Här skapas SQL Queries som underlättar vårt arbete. Viktigt är att rätt värden är satt under Bigram och Backupfolder.

![SQL queries](https://github.com/DambergC/VismaScript/blob/d49564e9d1f41395cf79a763b86de1d6fb58c7e5/ReadMEPictures/SQLqueries.png)

### DBupgrade

SQL Queries för att uppgradera PPP PUD och PUF.

***Textfil sparas under backupmappen du valt***

### QRRead

När man installerar Quick Report för första gången så behöver man sätta upp några konton i SQL.

***Skickas till urklipp, ej till fil***

### Change Password SQL

Vid byte av lösenord hos kund så är det enklare att köra en SQL querie så byts lösenorden på en gång.

***Skickas till urklipp, ej till fil***

Följande konton byts:
- MODUL_DashBoardUser
- MODUL_MenuUser
- MODUL_NeptuneAdmin
- MODUL_NeptuneReportUser
- MODUL_NeptuneUser
- MODUL_QuickReportAdmin
- MODUL_QuickReportUser
- Visma_CurrencyUser
- Visma_CommonStorageUser
- Visma_SchedulerUser
- rspdbuser

Sedan så behöver du byta i PIN manuellt.

### Nytt eget NA_Admin konto

Då konton stoppas från inloggning så behöver vi ett konto hos kunden i Neptune.

Denna SQL query skapar en användare med de rättigheter som vi behöver för att komma in systemet.

Du svarar på frågorna så skapas sedan SQL Query som skickas till ditt urklipp.

Första gången du loggar på med kontot så använder du lösenordet Vism@Cyg@te!!.

***Se till att byta till ett eget lösenord det första du gör***

## Stop And Start

Här stoppar du och startar tjänster och site.

När du stoppar tjänster så skapas en RestorePoint.xml som sparas i den valda Backupkatalogen.

När du sedan väljer att starta tjänsterna så återställer skriptet tjänsterna till hur dom var innan du gjorde uppgraderingen.

Två nya funktioner finns nu i skriptet som är kopplade till tjänsterna.

![Stop and Start](https://github.com/DambergC/VismaScript/blob/224017569cb62235422ec55eeef3ff26406d34fe/ReadMEPictures/StopStart.png)

### Load RestoreFile

Här så får du välja att öppna RestorePoint.xml om det ligger en i Backupkatalogen.då får du en enkel överblick hur kundens tjänster hade för status.

## Check after uppgradera

För att kunna jämföra hur det ser ut efter uppgraderingen så kan du få en enkel överblick som du sedan kan jämföra med när du innan valt att ladda in RestorePoint.xml

## CleanUp

Denna funktion är helt ny och den går bara att köra om du gjort en avinstallation. Om du inte gjort avinstallation så kommer du inte kunna resna något med koppling till Visma produkter.

De knappar som är öppna är Temp Asp.NET, Inetpub Logs och Install Catalog.

Alla knappar under CleanUp är kopplade till vad i vårt dokument säger vi ska rensa.

![CleanUP](https://github.com/DambergC/VismaScript/blob/224017569cb62235422ec55eeef3ff26406d34fe/ReadMEPictures/CleanUp.png)

### Temp Asp.NET

Rensar följande kataloger:

- "C:\Windows\Microsoft.NET\Framework\v4.0.30319\Temporary ASP.NET
- "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\Temporary ASP.NET Files"

### Inetpub

Rensar följande kataloger:

- C:\inetpub\logs\logfiles

***OBS! Det finns kunder som har installerat det på d:\inetpub\logs\logfiles då tar inte denna funktion bort filerna.*** 

### Install Catalog

Denna funktion rensar egentligen samma sak som som VPBS där du väljer att ta bort nedladdade versioner av Valda produkter.

Följande kataloger rensas
-	Support
-	SAP
-	Preinstall
-	Microsoft
-	HRM
-	FDN med undantag för PIN

## Tools

Sista knappen på huvudsidan för vårt skript är en samling av verktyg som vi kan ha behov av under en felsökning eller en uppgradering.

![Tools](https://github.com/DambergC/VismaScript/blob/224017569cb62235422ec55eeef3ff26406d34fe/ReadMEPictures/Tools.png)

### Cert Permissions & Cert Thumbprint

Här under kan du sätta defaulträttigheter som krävs för certifikat för Personec P och få ut Thumbprint för valt certifikat.

### Reset WWW PUD PPP

Ibland så råkar kund ändra på rättigheter i filstrukturen under WWW för PUD och PPP. Denna funktion återställer till de rättigheter som behövs för att Personec P ska fungera.

### Download IISCrypto

Nedladdning av verktygen för att sätta site till Strict.

Laddas ner till d:\visma\install

### Download VPBS

Nedladdning av nedladdningsverktyget för Visma.

OBS! måste har rätt värden under Bigram och BackupFolder.

### Decrypt and Encrypt Backup 

Två knappar för att avkryptera eller kryptera en vald backupkatalog.

### TLS Check

Kontrollerar om miljön är satt till Strict och endast accepterar TLS 1.2

Skriptet kollar efter 1.0 och  1.1 och rapporterar tillbaka hur status är.

![TLS Check](https://github.com/DambergC/VismaScript/blob/224017569cb62235422ec55eeef3ff26406d34fe/ReadMEPictures/TLSCheck.png)

### TLS 1.2 Regfix

Skjuter in de två registervärden som behövs för TLS 1.2

### CiceronFix

Mer info kommer när den funktionen är klar

## Övriga funktioner

Utöver nämnda funktioner så finns det några funktioner som kan vara bara att veta om.

### Load Logfile

Här väljer du vilken logfil du vill läsa och se vad som är gjort och inventerat.

### Cloes And Encrypt

När du är klar med uppgraderingen och vill avsluta skriptet på ett korrekt sätt så väljer du denna knapp. Då får du valet att kryptera backupfilernas web.config.
## Authors

- [@DambergC](https://www.github.com/DambergC)

