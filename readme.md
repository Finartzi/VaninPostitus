# Vanin Postitus [![img.png](res/Vaninlogo.png)](https://vanin.yhdistysavain.fi/)
@author: arto.paeivinen@gmail.com

Copyright: [MIT](mit_licence.md)

Vantaan invalidit ry:n tarve postittaa jäsentiedotteita liittyy myös postituskustannuksiin.
Tämän ohjelman idea on muokata Microsoft Excel tiedostoa seuraavasti:

1. Jäsentiedot tuodaan ensin Excelin rekisteriin "Tietokanta"
2. Ohjelma luo rekisteristä "Tietokanta" myös "tietokanta.json" tiedoston, jossa nämä tiedot ovat ikäänkuin varmuuskopioituna ja ovat siten palautettavissa, jos itse rekisteri muuttuu.
2. Ohjelma kopioi Tietokannasta ensin ne jäsenet uuteen rekisteriin "Sähköpostitse", jotka haluavat jäsentiedotteen sähköpostitse
3. Ne jäsenet, jotka edelleen haluavat postitse toimitettavan jäsentiedotteen, ohjelma kopioi rekisteriin "Postitse"
4. Ohjelma lajittelee postitse toimitettavien jäsenten osoitteiden mukaan jäsenet niin, että samaan postitusosoitteeseen ei menisi kuin yksi jäsentiedote 
5. Ohjelma luo postitettavien osoitetiedot tarroiksi tulostettaviksi käyttäen Microsoft Word ohjelmaa.
6. Ohjelma myös luo sähköpostituslistan valmiiksi (Template) siten, että vastaanottajat luetellaan kohtaan "BCC" (blind carbon copy), jottei sähköpostin vastaanottajat suoraan voisi tunnistaa toisia vastaanottajia. Tässä on kuitenkin muistutettava, että sähköpostijärjestelmässä on tunnettu vika, jonka johdosta muut sähköpostiosoitteet voivat paljastua, jos tällaiseen "BCC"-postaukseen vastataan "vastaa kaikille" -vaihtoehdolla.

Ohjelman luomisessa on käytetty kuvitteellisia osoitetietoja ohjelman testaamiseksi.

## Käyttöön liittyvää

Excel-tiedosto toimiakseen tässä ohjelmassa ei saa olla jo valmiiksi auki, muutoin syntyy "PermissionError: [Errno 13] Permission denied" virheilmoitus.

Tämä ohjelma on tehty Python 3.10.11 ja PyPi 23.1.2 versioilla, mutta uudemmat käyvät aina.

## Ohjelmointiympäristö

- PyCharm 2022.2.5 (Community Edition)  = Ilmaisversio 😍 
- Build #PC-222.4554.11, built on March 15, 2023
- Runtime version: 17.0.6+7-b469.82 amd64
- VM: OpenJDK 64-Bit Server VM by JetBrains s.r.o.
- Windows 11 10.0
- GC: G1 Young Generation, G1 Old Generation
- Memory: 2048M
- Cores: 4

### Mistä Python-moduulit saa? [![img.png](res/Pythonlogo.png)](https://www.python.org/downloads/windows/)
- https://www.python.org/downloads/windows/

Lisäksi tarvitaan PyPi-moduulit
- https://phoenixnap.com/kb/install-pip-windows

#### Ohjeita
Sivulta voi valita kielen, ja version, jolla dokumentaatiota haluaa seurata
- https://docs.python.org/3/using/windows.html

### Mistä PyCharm-ohjelman saa? [![img.png](res/JetBeans_PyCharm.png)](https://www.jetbrains.com/pycharm/download/#section=windows)

Ilmaisversio on mainittu "**Community**" - versio 
https://www.jetbrains.com/pycharm/download/#section=windows
Lisäksi suosittelen .md-tiedostojen käsittelyyn "Obsidian"-ohjelmaa (https://obsidian.md/download)
.md-tekstitiedostot ovat helposti muunnettavissa html-sivuiksi, jolloin ne ovat edustavia esitellä myös nettiselaimessa.
- https://markdowntohtml.com/
- https://adamtheautomator.com/convert-markdown-to-html/
- https://notepad-plus-plus.org/downloads/
- https://github.com/notepad-plus-plus/nppPluginList
- https://github.com/mohzy83/NppMarkdownPanel (Plugin Notepad++ -ohjelmaan lisäosana)
- https://github.com/mohzy83/NppMarkdownPanel/releases

## Lisätietoja

- https://www.addictivetips.com/windows-tips/fix-running-scripts-is-disabled-on-this-system-powershell-on-windows-10/
- https://stackoverflow.com/questions/46896093/how-to-activate-virtual-environment-from-windows-10-command-prompt
- https://python.land/virtual-environments/virtualenv
- https://www.datacamp.com/tutorial/python-excel-tutorial
- https://stackoverflow.com/questions/60044233/converting-excel-into-json-using-python
- https://automatetheboringstuff.com/chapter12/
- https://openpyxl.readthedocs.io/en/latest/

### Todettuja ongelmia
- Moduuli "xlrd" ei toimi uusien Exceltaulukoiden kanssa, sillä se ei tue formaattia .xlsx, vaan vain vanhempaa .xls -formaattia.
- Edellä mainitusta syystä onkin käytettävä isoa "pandas"-moduulia, mutta se ei ole kevyt ohjelman kannalta. Siksi käytän openpyxl-moduulia, vaikka se ei annakaan valmiita työkaluja tietojen edelleen käsittelyyn.

### Tekniikoita tutkittavaksi myöhemmin
- jsonpickle (https://pypi.org/project/jsonpickle/)
- marshmallow (https://marshmallow.readthedocs.io/en/stable/)

