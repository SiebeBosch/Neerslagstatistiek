# Neerslagstatistiek
Neerslagstatistieken voor Nederland: implementatie van de publicaties van STOWA

## Inleiding
Deze git-repository bevat een implementatie van de neerslagstatistieken voor Nederland zoals gepubliceerd door STOWA in de jaren 2015, 2019 en 2024.
De resultaten van de macro's en scripts in deze repository vormen input voor o.a.:

* De Nieuwe Stochastentool (https://github.com/SiebeBosch/DeNieuweStochastentool)
* Meteobase (https://github.com/SiebeBosch/meteobase).

## Werking
De kansverdelingsfuncties zoals door STOWA gepubliceerd worden omgezet in discrete klassen met neerslagvolume en bijbehorende kans. Andersom gebeurt ook: volume als functie van kans.

## Implementaties
De neerslagstatistieken ontsluiten we op de volgende manieren:

* Een Excel-macro. Zie de folder Excel. Deze macro geeft als functie van neerslagvolume de bijbehorende kans/herhalingstijd. De uitkomsten vormen op hun beurt weer input voor De Nieuwe Stochastentool

### Excel
Het Excel-document bevat op dit moment tabellen voor de publicatiejaren 2015 en 2019:

Het tabblad "KVD" bevat de parameterwaarden voor de kansverdelingsfuncties. Aan de grootheid 'neerslagvolume' is in de genoemde publicaties de GEV-kansdichtheidsfunctie gefit (Generalized Extreme Values). Het tabblad bevat voor verschillende duren, seizoenen en klimaatscenario's drie parameterwaarden:

* mu: de locatieparameter
* sigma: de schaalparameter
* kappa: de vormparameter

Voor de publicatie uit 2019 komt daar nog de dispersiecoëfficiënt bij.

Het tabblad 'Resultaten' bevat de complete lijst aan volumes en bijbehorende frequenties. Deze lijst is bedoeld om te worden geïmporteerd in de database die door De Nieuwe Stochastentool wordt gebruikt.





# Literatuur







