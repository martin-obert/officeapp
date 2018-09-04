# Office App

spustielné .exe je ve složce dist OfficeApp.zip

###Nastavení App.config
### Sekce - ExcelFormattersSection
formatter
- id: unikátní identifikátor (string)
- keyword: klíčové slovo, které se vysktuje v dané buňce
- replacement: náhlrada za keyword
- range: specifikuje daný range (zatím odzkoušeno jen na jedné buňce)

###Příklady:
Nahradí "Werkstückname:" za "Werkstückname"
<formatter id="2" keyword="Werkstückname:" replacement="Werkstückname" />

Nastaví buňce v oblastii c7 až c7 formát data i času na hh:mm:ss
<formatter id="12" range="c7:c7" numberFormat="hh:mm:ss" />

### Sekce - appSettings
- filename: název souborů, kterých se formátování týká (bez čísla a přípony na konci)
- directory: cesta ke složce, kde se soubory nacházejí
- range: oblast, která se bude zpracovávat