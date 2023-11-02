# Tiss_Exams
Diese Skript dient zum auslesen der Prüfungstermine im TISS.

# Usage
Verwendeten Pakete sind in requirements.txt zu finden. 
Installation unter Windows mittels 'pip install -r requirements.txt'.

Das Skript verwendet Selenium with Python um die Prüfungsdaten auszulesen. Hierfür muss ein Browse-Driver installiert werden. 
Standardmäßig wird Chrome verwendet. 
Die benötigten Driver können hier gefunden werden.(https://selenium-python.readthedocs.io/installation.html#drivers). Seit Chrome Version 115 ist die Versionskontrolle nicht mehr nötig und es kann die aktuele Version des Chromedriver heruntergeladen werden. Diese wird in den selben Ordner gespeichert wie main.py.

Die LVA Nummern werde aus einer Datei ausgelesen (LVA-Nummern) und müssen dem Format in der XLSX-Datei entsprechen. 

# ID
Da die Tabelle mit den Prüfungsterminen dynamisch erzeugt wird muss jedes Semester die ID der Tabelle überprüft werden. Hierzu muss das Entwicklertool verwendet werden. Mit dessen hilfe kann die ID gefunden werden. Im folgenden Bild ist ein Beispiel hierfür zu sehen. Diese Id muss dann im main.py file der variable 'id' zugeordnet werden.
![id_example](/img/find_id.PNG)

# Log Datei
Hier werden die LVA-Nummern abgelegt bei denen es keine Prüfungstermine gibt oder bei welchen die LVA Nummer falsch sind. 

# Lizenz
Das komplette Projekt steht unter der Creative Commons Lizenz CC-BY-SA und darf benutzt, kopiert und verändert werden, sofern das veränderte Werk unter der gleichen Lizenz weitergegeben wird und die Namen der Urheber genannt werden.

#### made with Love by Teamwolke
