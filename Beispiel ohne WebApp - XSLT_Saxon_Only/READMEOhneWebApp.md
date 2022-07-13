# Projekt 4: invoice_project_cmp Beispiel ohne WebApp


## Durchführung:

### 1. Ausführen des Saxon Kommandozeilenbefehls (sheet1 wird direkt in Ordner eingebunden) (--Achtung, falls sheet1.xml in /ExcelInhalt/xl/worksheets existiert --> löschen)
```
java -cp .\saxon-he-10.3.jar net.sf.saxon.Transform -t -s:.\invoiceExample.xml  -xsl:.\ExcelCreator.xsl -o:.\ExcelInhalt\xl\worksheets\sheet1.xml
```
### 2. Manuelles zippen des Inhaltes des Ordners ExcelInhalt (nicht den Ordner an sich!) 

### 3. Umbenennen der .zip in .xlsx

bei Fehlern wird Excel seine Selbstreparatur nutzen (erlauben)




