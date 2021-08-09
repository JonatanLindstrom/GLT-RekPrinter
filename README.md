# GLT RekPrinter :package: :page_with_curl:

Ett program för att splitta upp reklistan från Pyramid till en sida per enhet. För att hålla programmer fungerande behöver variabeln "placeMap" uppdateras när en enhet tillkommer, ändras eller tas bort.

## Installation :open_file_folder:

Antingen kan programmet köras via en byggd applikation som hittas [här](https://github.com/JonatanLindstrom/GLT-RekPrinter/releases "RekPrinter releases") (rekommenderas). Det går även att hämta hem koden och köra den från kommandotolken vilket kräver något mer installation men går att ändra efter behov.

### Klonad kod :inbox_tray:

Börja med att klona koden från GitHub, hela koden kan laddas ned [här](https://github.com/JonatanLindstrom/GLT-RekPrinter/archive/master.zip).

Om datorn ej har python installerat, hämta det och installera från deras [officiella hemsida](https://www.python.org/)

Då programmet bygger på andra paket, kör följande kommando i kommandotolken:

``` pip install openpyxl ```

Efter detta går programmet att köra genom att i en kommandotolk öppnad i filens mapp skriva

``` python RekPrinter.py```

### Snabbkommando

För att underlätta efter installationen går det att skapa en .bat fil med följande innehåll:

``` python C:/FullständigAdressTillFilen/RekPrinter.py ```

Denna fil kan då läggas var som helst (ex. skrivbordet) och vid körning starta programmet.
