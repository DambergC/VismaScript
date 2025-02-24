
# CygateScript.ps1

Ett skript för att underlätta arbetet med uppgradering av Visma´s produkt Personec P HRM.
Skriptet har funktioner som underlättar och sparar tid samt gör att uppgraderingsstegen oavsett konsult blir standardiserade.



## Authors

- [@DambergC](https://www.github.com/DambergC)


## Installation

Första gången man ska ladda ner skriptets så kör man följande kommandorad i Powershell

```Powershell
Invoke-WebRequest 'https://github.com/Dambergc/Vismascript/releases/latest/download/CygateScript.ps1' -OutFile D:\Visma\Install\Backup\CygateScript.ps1 -Verbose
```
    
## Uppstart av skript
För att starta skriptet så starta ett eleverat Powershell fönster och sedan skriv sökvägen till skriptet

```Powershell
D:\visma\install\backup\CygateScript.ps1
```
Skulle det finnas kund som har sin installation på annan enhet en den som är default (D:\) så hanteras det av skriptet och då är sökvägen självklart samma som ovan men med ändring av enhetsbokstaven.

Om det skulle ha varit en uppgradering av skriptet tillgänglig så laddas den ned och sedan så kommer du få en dialogruta om att ny version av skriptet nedladdad och att du ombedes att starta om skriptet.
## Visma Services Trusted Users
För att få köra skriptet så är det samma krav som på Public Installer, du måste vara medlem i den lokala gruppen ”Visma Services Trusted Users” Hur du lägger till dig i gruppen går att göra det manuellt via Computer Management eller via powershell som du hittar i vårt uppgraderingsdokument.
## Krav på att välja BIGRAM och BACKUPFOLDER
För att komma åt funktionerna i skriptet så krävs att du väljer BIGRAM och BackupFolder under den knappen ”Bigram And BackupFolder”. Efter att du valt det så är övriga knappar öppna så att du kan komma åt underliggande funktioner.

Det BIGRAM och BackupFolder du väljer kommer följa med dig under hela tiden så länge du har skriptet igång. Skulle det stängas ner av misstag så måste du åter välja det vid uppstart.

## Logfil
Allt som väljs och de val man gör i skriptet skrivs i en logfil som ligger under d:\visma\install\backup. Då logfilen används innan man valt BackupFolder så ligger den mer centralt. Dessa logfiler får vid behov rensas manuellt. Det skapas en logfil för varje dag man kör skriptet.

 
