# Neerslagstatistiek
Neerslagstatistieken voor Nederland: implementatie van de publicaties van STOWA

## Inleiding
Deze git-repository bevat een implementatie van de neerslagstatistieken voor Nederland zoals gepubliceerd door STOWA in de jaren 2014, 2019 en 2024.
De resultaten van de macro's en scripts in deze repository vormen input voor o.a.:

* De Nieuwe Stochastentool (https://github.com/SiebeBosch/DeNieuweStochastentool)
* Meteobase (https://github.com/SiebeBosch/meteobase).

## Werking
De kansverdelingsfuncties zoals door STOWA gepubliceerd worden omgezet in discrete klassen met neerslagvolume en bijbehorende kans. Andersom gebeurt ook: volume als functie van kans.

## Implementaties
De neerslagstatistieken ontsluiten we op de volgende manieren:

* Een Excel-macro. Zie de folder Excel. Deze macro geeft als functie van neerslagvolume de bijbehorende kans/herhalingstijd. De uitkomsten vormen op hun beurt weer input voor De Nieuwe Stochastentool








