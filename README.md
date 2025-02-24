
# CygateScript.ps1

Ett skript för att underlätta arbetet med uppgradering av Visma´s produkt Personec P HRM.
Skriptet har funktioner som underlättar och sparar tid samt gör att uppgraderingsstegen oavsett konsult blir standardiserade.



## Start av skriptet - Första gången
Det finns två sätt att plocka ner skriptet beroende på hur kunden tillåter internetåtkomst.
-	Nedladdning via powershell
    - Invoke-WebRequest 'https://github.com/Dambergc/Vismascript/releases/latest/download/CygateScript.ps1' -OutFile D:\Visma\Install\Backup\CygateScript.ps1 -Verbose
    - b.	Gå till Github till adressen https://github.com/dambergc/vismascript och går till Releases och ladda ner filen CygateScript.ps1 under senaste Release.

![Välj rätt i github.](https://github.com/DambergC/VismaScript/blob/fc34252a991c24617acf65b4885b48fc6f5ca5e3/ReadMEPictures/github.png)

Är det första gången så behöver du lägga till lite värden, spara BIGRAM och sedan skapa en backupmapp.

![Ändra versionsnummer](https://github.com/DambergC/VismaScript/blob/d49564e9d1f41395cf79a763b86de1d6fb58c7e5/ReadMEPictures/edit_versions.png)

## Uppstart av skript
För att starta skriptet så starta ett eleverat Powershell fönster och sedan skriv sökvägen till skriptet

```Powershell
D:\visma\install\backup\CygateScript.ps1
```
Skulle det finnas kund som har sin installation på annan enhet en den som är default (D:\) så hanteras det av skriptet och då är sökvägen självklart samma som ovan men med ändring av enhetsbokstaven.

![Hur man startar skriptet.](https://github.com/DambergC/VismaScript/blob/f3c6cc2df2366acbac5a6164e5042a71ea5b0712/ReadMEPictures/CMDstart.png)

Om det skulle ha varit en uppgradering av skriptet tillgänglig så laddas den ned och sedan så kommer du få en dialogruta om att ny version av skriptet nedladdad och att du ombedes att starta om skriptet.
## Visma Services Trusted Users
För att få köra skriptet så är det samma krav som på Public Installer, du måste vara medlem i den lokala gruppen ”Visma Services Trusted Users” Hur du lägger till dig i gruppen går att göra det manuellt via Computer Management eller via powershell som du hittar i vårt uppgraderingsdokument.
## Krav på att välja BIGRAM och BACKUPFOLDER
För att komma åt funktionerna i skriptet så krävs att du väljer BIGRAM och BackupFolder under den knappen ”Bigram And BackupFolder”. Efter att du valt det så är övriga knappar öppna så att du kan komma åt underliggande funktioner.

Det BIGRAM och BackupFolder du väljer kommer följa med dig under hela tiden så länge du har skriptet igång. Skulle det stängas ner av misstag så måste du åter välja det vid uppstart.

## Logfil
Allt som väljs och de val man gör i skriptet skrivs i en logfil som ligger under d:\visma\install\backup. Då logfilen används innan man valt BackupFolder så ligger den mer centralt. Dessa logfiler får vid behov rensas manuellt. Det skapas en logfil för varje dag man kör skriptet.
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





## Authors

- [@DambergC](https://www.github.com/DambergC)

