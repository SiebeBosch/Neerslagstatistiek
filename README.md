# Neerslagstatistiek van Nederland
Neerslagstatistieken voor Nederland: implementatie van de publicaties van STOWA uit 2015/2018, 2019 en 2024.
Samengesteld door Siebe Bosch, met dataleveringen en ondersteuning door Dorien Lugt, Robin Nicolai en Rudolf Versteeg.

Datum: 23-5-2024

## Inleiding
Deze git-repository bevat een implementatie van de neerslagstatistieken voor Nederland zoals gepubliceerd door STOWA in de jaren 2015 2018 (update korte duren), 2019 en 2024.
De resultaten van de macro's en scripts in deze repository vormen input voor o.a.:

* De Nieuwe Stochastentool (https://github.com/SiebeBosch/DeNieuweStochastentool)
* Meteobase (https://github.com/SiebeBosch/meteobase).

## Werking
De kansverdelingsfuncties zoals door STOWA gepubliceerd kunnen met behulp van de scripts in deze repository worden omgezet in discrete klassen met neerslagvolume en bijbehorende klassenfrequentie (kans). Andersom kan ook: volume als functie van kans.

## Implementaties
De neerslagstatistieken ontsluiten we op de volgende manieren:

* Een Excel-macro. Zie de folder Excel. Deze macro geeft als functie van neerslagvolume de bijbehorende kans/herhalingstijd. De uitkomsten vormen op hun beurt weer input voor De Nieuwe Stochastentool
* Een Jupyter-notebook. Zie de folder Jupyter. Dit notebook berekent de multipliers om komen van neerslagvolumes uit 2019 scenario Huidg naar een 2024 klimaatscenario en produceert daarbij Excel-bestanden met volumes

### Excel
De basis voor het Excel-macro wat we hieronder bespreken wordt gevormd door de VBA-functies in het bestand STOWA_Neerslagstatistiek.bas. Dit bestand bevat alle functies waarmee gebruikers van Excel herhalingstijden en overschrijdingskansen kunnen opvragen voor elk van de scenario's zoals door STOWA gepubliceerd.

#### Werkboeken
_Statistiek van neerslagvolumes ten behoeve van De Nieuwe Stochastentool:_
* Neerslagstatistieken_2015.xlsm: bevat de neerslagstatistieken conform publicatie door STOWA in 2015
* Neerslagstatistieken_2019.xlsm: bevat de neerslagstatistieken conform publicatie door STOWA in 2019.
* Neerslagstatistieken_2024.xlsm: bevat de neerslagstatistieken conform publicatie door STOWA in 2024.

  Statistieken klimaat 'huidig' ongewijzigd tov publicatie 2019. Voor de zichtjaren wordt het 'verandergetal) (een multiplier) berekend die een functie is van de verwachte temperatuursstijging en het seizoen.
  Op het tabblad 'Start' kan de gebruiker de statistieken exporteren naar CSV ten behoeve van opname in de database van De Nieuwe Stochastentool.

_Statistiek van neerslapatronen ten behoeve van De Nieuwe Stochastentool:_
* Overzicht_patronen_2019.xlsm: dit zijn de patronen en hun kansen zoals gepubliceerd in 2019. Ze zijn ook van toepassing op de scenario's van 2024 en de kansen zijn identiek voor alle klimaatscenario's (zie pag 156 van het STOWA-rapport NEERSLAGSTATISTIEK EN -REEKSEN VOOR HET WATERBEHEER 2019). Het werkboek bevat een aantal buttons waarmee de patronen en hun kansen kunnen worden omgezet naar tekst die in een CSV-file geplakt kunnen worden. Van daaruit kunt u ze importeren in De Nieuwe Stochastentool.

#### VBA
Alle Excel-werkboeken gebruiken hetzelfde bronbestand met VBA-functies: STOWA_Neerslagstatistiek.bas.

Let op: in alle functies gebruiken we zichtjaar 2014 om het 'huidige klimaat' mee aan te duiden; ook al is dit mogelijk niet langer opportuun. Dit om consistentie tussen de verschillende publicatiejaren te kunnen behouden.

_Publicatie neerslagstatistieken 2024_:
* STOWA2024_JAARROND_V: berekent het jaarrond overschrijdingsvolume, gegeven de duur in minuten, herhalingstijd in jaren, zichtjaar (2014, 2033, 2050, 2100, 2150) en scenario (L, M, H)
* STOWA2024_JAARROND_T: berekent de herhalingstijd voor wintergebeurtenissen, gegeven het volume in millimeters, de duur in minuten, zichtjaar (2014, 2033, 2050, 2100, 2150) en scenario (L, M, H) 
* STOWA2024_NDJF_V: berekent het winteroverschrijdingsvolume (nov, dec, jan, feb), gegeven de duur in minuten, herhalingstijd in jaren, zichtjaar (2014, 2033, 2050, 2100, 2150) en scenario (L, M, H)
* STOWA2024_NDJF_T: berekent de herhalingstijd voor wintergebeurtenissen, gegeven het volume in millimeters, de duur in minuten, zichtjaar (2014, 2033, 2050, 2100, 2150) en scenario (L, M, H)

_Publicatie neerslagstatistieken 2019_:
* STOWA2019_JAARROND_V: berekent het jaarrond overschrijdingsvolume, gegeven de duur in minuten, herhalingstijd in jaren, zichtjaar (2014, 2030, 2050, 2085) en scenario (GL, GH, WL, WH)
* STOWA2019_JAARROND_T: berekent de herhalingstijd voor wintergebeurtenissen, gegeven het volume in millimeters, de duur in minuten, zichtjaar (2014, 2030, 2050, 2085) en scenario (GL, GH, WL, WH) 
* STOWA2019_NDJF_V: berekent het winteroverschrijdingsvolume (nov, dec, jan, feb), gegeven de duur in minuten, herhalingstijd in jaren, zichtjaar (2014, 2030, 2050, 2085) en scenario (GL, GH, WL, WH)
* STOWA2019_NDJF_T: berekent de herhalingstijd voor wintergebeurtenissen, gegeven het volume in millimeters, de duur in minuten, zichtjaar (2014, 2030, 2050, 2085) en scenario (GL, GH, WL, WH)

_Publicatie neerslagstatistieken 2015/2018_:
* STOWA2015_2018_JAARROND_V: berekent het jaarrond overschrijdingsvolume, gegeven de duur in minuten, herhalingstijd in jaren, zichtjaar (2014, 2030, 2050, 2085) en scenario (GL, GH, WL, WH)
* STOWA2015_2018_JAARROND_T: berekent de herhalingstijd voor wintergebeurtenissen, gegeven het volume in millimeters, de duur in minuten, zichtjaar (2014, 2030, 2050, 2085) en scenario (GL, GH, WL, WH) 
* STOWA2015_JAARROND_V: berekent het winteroverschrijdingsvolume (nov, dec, jan, feb), gegeven de duur in minuten, herhalingstijd in jaren, zichtjaar (2014, 2030, 2050, 2085) en scenario (GL, GH, WL, WH)
* STOWA2015_JAARROND_T: berekent de herhalingstijd voor wintergebeurtenissen, gegeven het volume in millimeters, de duur in minuten, zichtjaar (2014, 2030, 2050, 2085) en scenario (GL, GH, WL, WH)


### Jupyter
De Jupyter-notebooks berekenen het verandergetal (multiplier) dat nodig is om van het neerslagvolume onder scenario 2019_Huidig te komen tot het volume onder de scenario's 2024.
Het oorspronkelijke notebook werd opgesteld en aangeleverd door HKV-Lijn-In-Water.

### R
De map met R bevat het oppervlaktereductie-script van Aart Overeem (KNMI). Met dit script wordt extreme neerslag zoals gemeten op een puntlocatie (bijv. meetstations) gecorrigeerd wanneer deze moet worden toegepast over een groter gebied.

## Literatuur

STOWA. (2015). Nieuwe neerslagstatistieken voor het waterbeheer: extreme neerslaggebeurtenissen nemen toe (Publicatienummer 2015-10a). Verkregen van [https://www.stowa.nl/publicaties/nieuwe-neerslagstatistieken-voor-het-waterbeheer-extreme-neerslaggebeurtenissen-nemen](https://www.stowa.nl/sites/default/files/assets/PUBLICATIES/Publicaties%202015/STOWA%202015-10A.pdf)

STOWA. (2018). Neerslagstatistieken voor korte duren (Publicatienummer 2018-12). Verkregen van https://www.stowa.nl/sites/default/files/assets/PUBLICATIES/Publicaties%202018/STOWA%202018-12%20HR.pdf

STOWA. (2019). Neerslagstatistiek en reeksen voor het waterbeheer 2019 (Publicatienummer 2019-19). Verkregen van [https://www.stowa.nl/publicaties/neerslagstatistiek-en-reeksen-voor-het-waterbeheer-2019](https://www.stowa.nl/sites/default/files/assets/PUBLICATIES/Publicaties%202019/STOWA%202019-19%20neerslagstatistieken.pdf).

STOWA. (2024). Neerslagstatistiek Nog in te vullen





