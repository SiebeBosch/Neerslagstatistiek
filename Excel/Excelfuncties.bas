Attribute VB_Name = "Excelfuncties"

'Deze declaratie is een timer. Kan worden aangeroepen met de call Sleep(miliseconden)

Option Explicit


'-------------------------'
'Auteur: Siebe Bosch      '
'Hydroconsult             '
'Lulofsstraat 55, unit 47 '
'2521 AL Den Haag         '
'siebe@hydroconsult.nl    '
'0617682689               '
'-------------------------'

'VERSIE 4.37
'bovenstaande regel is bedoeld om het declareren van elke variabele verplicht te maken
'Beschikbare functies en routines:

'getallenreeksen:
'GOALSEEKRANGE               - voert een goalseek uit op een range
'INTERPOLATE                 - Interpoleert tussen twee xy-punten, blockinterpolation is optioneel
'EXTRAPOLATE                 - Extrapoleert lineair vanaf twee xy-punten
'FITLINEAR_A                 - berekent a in y = ax + b gegeven twee coordinaten
'FITLINEAR_B                 - berekent b in y = ax + b gegeven twee coordinaten
'INTERPOLATEFROMRANGE        - interpoleert een gegeven X in een range van X-waarden en een range van Y-waarden
'INTERPOLATERANGEFROMRANGE   - interpoleert voor een hele reeks getallen (in een range) uit een gegeven range van X en Y-waarden
'INTERPOLATEFROMRANGEPLUS    - interpoleert een gegeven X in een range van X-waarden en een range van Y-waarden, gegeven een ID in een range met ID's
'KLEINSTEKWADRATENMETHODE    - geeft het kleinstekwadratenverschil tussen een gemeten en berekende reeks
'ISSTRINGARRAYEMPTY          - test of een array van strings leeg is
'SORTARRAY                   - Sorteert een array met getallen in oplopende volgorde
'HEAPSORT                    - Creeert een array met de indexnummers van de gesorteerde input-array
'SORTCOLLECTIONBYKEY         - Creeert een array met de indexnummers van een op key gesorteerde collection van objecten
'COLLECTIONGETIDX            - geeft het indexnummer terug van een gegeveen waarde in een collectie
'COLLECTIONCONTAINS          - geeft een boolean terug met het antwoord of een collectie een gegeven waarde bevat
'RANDOM                      - creeert een random integer tussen een gespecificeerd minimum en maximum
'RANDOMDOUBLE                - creeert een random double tussen een gespecificeerd minimum en maximum
'MAXIMUM                     - geeft het maximum van twee getallen terug
'FROMWORKSHEET               - haalt data uit een range van een werkblad en stopt die in een array
'ARRAYVARIANTTOWORKSHEET     - schrijft data uit een array van het type variant naar een werkblad
'ARRAYDATETOWORKSHEET        - schrijft data uit een array van het type date naar een werkblad
'ARRAYSINGLETOWORKSHEET      - schrijft data uit een array van het type single naar een werkblad
'TIMESERIES2ARRAYS           - leest een tijdreeks van het werkblad in twee arrays
'USERSELECTRANGE             - laat de gebruiker een range op het werkblad selecteren
'RANGEADDRESSFROMRC          - definieert een bereik op basis van rij- en kolomnummers
'RANGECOLIDXFROMMAXVAL       - geeft het kolomnummer terug dat hoort bij de hoogste waarde uit een bereik
'ASSIGNVALUEBYMONTH          - geeft een waarde terug, afhankelijk van de maand van een gegeven datum
'CLASSIFYNUMBERBYCLASS       - classificeert een getal naar een gegeven ondergrens, bovengrens en increment
'DESIGNTWOPARTTEMPORALPATTERNS  - genereert een serie tijdsafhankelijke patronen, gegeven een duur (tijdstappen) en stapgrootte (percentiel)
'DESIGNTEMPORALBLOCKPATTERNS - genereert een serie blokpatronen, gegeven een duur (tijdstappen) en blockgrootte

'graphics:
'SVGCOMPONENTTRAPEZIUMSHARP  - genereert een trapeziumvorm en schrijft het in SVG-formaat (zonder headers)

'collections:
'MAXFROMCOLLECTION           - retourneert de maximumwaarde uit een collectie
'MINFROMCOLLECTION           - retourneert de minimumwaarde uit een collectie
'AVGFROMCOLLECTION           - retourneert de gemiddelde waarde uit een collectie

'kansrekening:
'GEVCUM                      - berekent de cumulatieve kansdichtheid volgens de Gegeneraliseerde Extremewaardenverdeling (GEV)
'GENPARETOCDF                - berekent de cumulatieve kansdichtheid volgens de Gegeneraliseerde Pareto-kansverdeling (GENPAR)
'EXP2PCUM                    - berekent de cumulatieve kansdichtheid volgens de tweepunts-exponentiele verdeling (EXP2P)
'CONDWEIBULLCDF              - berekent de cumulatieve kansdichtheid volgens de conditionele weibull-verdeling (CONDWEIBULL)
'GENPARETOCDF                - berekent de cumulatieve kansdichtheid volgens de gegeneraliseerde pareto-verdeling
'BEREKENSTOCHASTVOLUMEKLASSE       - berekent de aangepaste frequentie wanneer bepaalde stochasten voor NeerslagVolume zijn uitgeschakeld
'BEREKENSTOCHASTPATROONKLASSE      - berekent de aangepaste frequentie wanneer bepaalde stochasten NeerslagPatroon zijn uitgeschakeld
'HERH2KLASSEFREQ             - berekent de frequentie van een klasse gegeven herhalingstijd van vorige, huidige en volgende klasse en duur
'HERHFROMSTOCHASTICRESULT    - berekent de waterhoogte behorende bij een gegeven herhalingtijd op basis van de uitkomsten van een stochastenanalyse
'KLASSEFREQUENTIEUITHERHALINGSTIJD  - berekent de frequentie van een klasse gegeven de onder- en bovengrens van volumes uitgedrukt in herhalingstijd
'KLASSEKANSUITOVERSCHRIJDINGSKANSEN   - berekent de kans voor een klasse, gegeven omringende overschrijdingskansen
'CLASSIFYDURATIONS           - classificeert gebeurtenissen naar hun duur door te zoeken naar de duur van een overschrijding

'meetkundig:
'OPPERVLAKAFGEPLATTECIRKEL   - berkent het oppervlak van een afgeplatte cirkel
'NATTEOMTREKAFGEPLATTECIRKEL - berekent de contactomtrek van een afgeplatte cirkel
'ELLIPSBREEDTE               - berekent bij verschillende hoogtes de breedte van een ellips
'ARCSIN                      - berekent de inverse sinus
'ARCCOS                      - berekent de inverse cosinus
'ARCTAN                      - berekent de inverse tangens
'ROTATEPOINT                 - roteert een xy-coordinaat rond een vastgelegd nulpunt en verplaatst het
'D2R                         - graden naar radialen
'R2G                         - radialen naar graden
'LINEANGLEDEGREES            - berekent de hoek van een lijn tussen twee xy-co-ordinaten
'POINTDISTANCE               - berekent de afstand tussen twee x,y coordinaten
'POINTINPOLYGON              - berekent of een gegeven punt binnen een polygoon ligt
'NEARESTPOINT                - zoekt gegeven een XY-coordinaat het dichtstbijzijnde punt uit een range
'POOLCOORDINAATX             - geeft X terug gegeven een hoek alpha (ten opzichte van noord-as) en lengte L
'POOLCOORDINAATY             - geeft Y terug gegeven een hoek alpha (ten opzichte van noord-as) en lengte L
'PYTHAGORAS                  - geeft lengte schuine kant terug
'PYTHAGORAS_INVERSE          - geeft lengte van een rechte kant terug
'TRAPEZIUMAREA               - geeft het oppervlak van een trapezium (bestaande uit een rechthoek met daarbovenop een driehoek) terug.

'wiskundig:
'MILEAGEONEUP                - verhoogt de 'kilometerstand' in een array met één
'MEETSCONDITION              - checkt of een waarde voldoet aan een bepaalde condition (bijv. ">= -0.52")
'RMSE                        - berekent de Root Mean Square Error voor twee ranges

'objecten in Excel
'CLEARCOMBOBOX               - verwijdert alle items uit een combobox
'GetShapeByNameFromWorksheet - vraagt een shape van een gegeven werkblad op, op basis van zijn naam.

'grafieken
'MAKESCATTERCHART
'MAKECHART
'EXPORTCHART                 - exporteert een grafiek naar een .png-bestand in dezelfde folder als de applicatie.

'datum- en tijdfuncties
'DAYSINMONTH                 - Geeft het aantal dagen in de maand van een gespecificeerde datum terug
'DAYSINMONTH2                - Geeft het aantal dagen in de maand van een gespecificeerd maandnummer
'ISLEAPYEAR                  - Geeft terug of een gegeven jaar een schrikkeljaar is (TRUE/FALSE)
'KWARTAAL                    - Geeft het kwartaalnummer van een datum terug
'ZOMERWINTERHALFJAAR         - Geeft het halfjaar van een datum terug
'METEOROLOGISCHSEIZOEN       - Geeft het meteorologisch seizoen terug waarin een opgegeven datum ligt
'METEOROLOGISCHHALFJAAR      - Geeft het meteorologische halfjaar terug van een opgegeven datum
'HYDROLOGISCHSEIZOEN         - Geeft het hydrologisch seizoen terug waarin een opgegeven datum ligt
'DOUBLE2DATETIMESTRING       - transformeert een getal (double) naar een datum/tijd-string
'DATEEXISTS                  - controleert of een opgegeven combinatie van dag, maand en jaar valide is
'DAYNUMBER                   - geeft het dagnummer in het jaar van een gegeven datum terug
'CALCDAYQUARTER              - geeft datum + eerste uur binnen het zesuurswindow van een gegeven datumtijd terug
'DATEFROMSTRING              - maakt een datum aan op basis van een string
'DATE2TEXT                   - maakt een string aan op basis van een datum
'TIMEFROMSTRING              - maakt een tijd aan op basis van een string
'DATEANDTIMEFROMSTRINGS      - maak een datum en tijd aan op basis van twee strings
'DECADE                      - berekent het decadenummer gegeven een datum
'SOBEKTIMETABLESTRING        - maakt een tijdstabel voor SOBEK in de Control.def

'werkbladfuncties
'HOR_ZOEKEN_DOUBLE           - laat de gebruiker zoeken op bass van waarden in de eerste TWEE rijen
'VERT_HORIZ_ZOEKEN           - Geef kolomnaam en rijnaam op, en krijg de inhoud van de bijbehorende cel terug
'VERT_ZOEKEN_DOUBLE          - laat de gebruiker zoeken op basis van waarden in de eerste TWEE kolommen
'VERT_ZOEKEN_TRIPLE          - laat de gebruiker zoeken op basis van waarden in de eerste DRIE kolommen
'VERT_ZOEKEN_QUADRUPLE       - laat de gebruiker zoeken op basis van waarden in de eerste VIER kolommen
'VERT_ZOEKEN_MIN             - Geeft de minimumwaarde terug uit een range waarin eenzelfde ID meerdere Parent1n voorkomt
'VERT_ZOEKEN_MAX             - Geeft de maximumwaarde terug uit een range waarin eenzelfde ID meerdere Parent1n voorkomt
'VERT_ZOEKEN_MODUS           - Geeft de meest voorkomende waarde terug uit een range waarin eenzelfde ID meerdere Parent1n voorkomt
'VERT_ZOEKEN_SOM             - Sommeert alle waarden uit kolom Y die gevonden worden bij een opgegeven zoekterm in kolom X
'VERT_ZOEKEN_NEARESTXY                  - Lookup in een range met X,Y en Waarde, waarbij de waarde van het meest dichtbijzijnde object wordt teruggegeven
'FindColumnInRange           - geeft de kolomindex van een range terug, gegeven een waarde die gezocht wordt. optioneel geeft hij een lege terug indien niet gevonden
'FindRowInRange              - geeft de rijindex van een range terug, gegeven een waarde die gezocht wordt. optioneel geeft hij een lege terug indien niet gevonden
'AVERAGEFROMRANGE            - geeft de gemiddelde waarde uit een range terug
'MINFROMRANGE                - geeft de laatste waarde uit een range terug
'MAXFROMRANGE                - geeft de kleinste waarde uit een range terug
'FIRSTFROMRANGE              - geeft de eerste waarde uit een range terug
'LASTFROMRANGE               - geeft de laatste waarde uit een range terug
'MOSTCOMMONFROMRANGE         - Geeft de meest voorkomende waarde terug uit een range
'GEWOGEN_GEMIDDELDE          - Geeft van een reeks voor elke ID de gewogen gemiddelde waarde terug op basis van bijv. meerdere waarde- en oppervlakteparen
'VERT_ZOEKEN_GROOTSTEAANDEELHOUDER      - Geeft terug voor welke 'aandeelhouder' de som van de waarden het grootst is, gegeven een objectID
'HEADERBYMAXIMUMVALUE        - Geeft voor een gegeven range de titel terug die staat boven de kolom met de grootste waarde
'WORKSHEETEXISTS             - Returns 'true if a worksheet exists
'SUMRANGE                    - Geeft de som van de inhoud van een range op een werkblad terug
'FRACTIONOFDAYSUM            - Geeft voor een bepaalde cel voor een bepaalde datum/tijd de fractie van de totale dagsom terug
'ISRANGEASCENDING            - Checkt of een opgegeven range een oplopende volgorde heeft
'MINYFROMXYRANGE             - Geeft de minimum Y-waarde terug uit een range met daarin X en Y waarden. Optioneel zoekrange beperken van Xmin tot Xmax
'MAXYFROMXYRANGE             - Geeft de maximum Y-waarde terug uit een range met daarin X en Y waarden. Optioneel zoekrange beperken van Xmin tot Xmax
'CONCATENATEALGEBRAIC        - Veegt termen uit een reeks samen tot een algebraische formule, bijv "X + Y + Z"
'CONCATENATEWITHDELIMITER    - Veegt waarden uit een reeks van cellen samen tot een string, met een gegeven delimiter
'ADDWORKSHEET                - Voegt een nieuw werkblad toe aan het huidige werkboek.
'FINDCOLUMNONWORKSHEET       - Zoekt het kolomnummer voor een gegeven header op een werkblad
'UNPIVOTMULTIHEADER          - ontvlecht een tabel met meerdere headers in zowel rij als kolom en zet hem klaar voor pivot-doeleinden
'UNPIVOTTABLE                - ontvlecht een tabel met één kolomheader en één rijheader en schrijft hem in COLHEADER, ROWHEADER, VALUE formaat
'UNPIVOT                     - converteert een 2D-tabel naar een Header1, Header2, Waarde-tabel voor pivot-doeleinden
'UNPIVOT2CSV                 - converteert een 2D-tabel naar een Header1, Header2, Waarde-tabel in een .csv-bestand
'RANGE2CSV                   - schrijft de gegevens uit een range naar een csv file.
'GOALSEEKTRIPLE              - zoekt het optimum voor een cel waarvan de waarde een functie is van drie variabelen
'GOALSEEKDOUBLE              - zoekt het optimum voor een cel waarvan de waarde een functie is van twee variabelen
'COLUMN_NUMBER               - zoekt het kolomnummer uit een reeks, gegeven een gezochte celinhoud
'PRINTARRAY                  - schrijft een array naar het werkblad
'RANGEVERTASCENDING           - checkt of een range in vertikale richting oploopt
'VALUEFROMCELLADDRESS        - geeft de waarde uit een cel terug gegeven zijn rij- en kolomnummer

'werkbladroutines voor ranges (hiervoor moet je wel een knop inbouwen)
'AGGREGERENNAARUREN          - aggregeert een tijdreeks met waarden naar hele uren
'COUNTSEQUENTIALEXCEEDANCES  - telt het aantal achtereenvolgende overschrijdingen van een drempelwaarde in een gegeven range (betaande uit een kolom)
'AGGREGEREN                  - aggregeert een tijdreeks door een vast aantal rijen over te slaan
'AGGREGATEFROMRANGE          - aggregeert een gegeven kolom uit een range op basis van waarden in een andere kolom en een opgegeven aggregatiemethode
'AGGREGATERANGECONDITIONALLY - aggregeert een range op basis van een geselecteerde kolom en een specificatie van de aggregatiemethode per kolom, maar met een voorwaarde voor de waarden uit een andere kolom
'COLUMNFROMRANGE             - geeft een kolom uit een range terug als range. Houdt ook rekening met multi-area ranges!
'CONDITIONALSUBRANGE         - geeft een subrange uit een range terug, waarvoor aan een gegeven voorwaarde wordt voldaan
'GETASCIIGRIDVALUES          - geeft voor gegeven X en Y coordinaten de bijbehorende waarde uit een ASCIIGRID terug
'GETROWCOLFROMASCIIGRID      - geeft voor gegeven grid-dimensies het rij- en kolomnummer behorende bij een X- en Y-coordinaat terug
'RANGEWITHHEADER2THREECOLRANGE - converteert een reeks met header en daaronder X en Y naar een reeks met drie kolommen: ID, X, Y
'WEAVETABLESBLOCKINTERPOLATION - weeft twee tabellen met datum/waarde ineen, werkend met blokinterpolatie. Handig voor gemaalactiviteiten
'TRUNCATERANGEBYEMPTYROWS      - kapt een gegeven range af op lege rijen

'eenheidsconversies
'CELCIUS2KELVIN              - converteert graden celcius naar kelvin
'KELVIN2CELCIUS              - converteert graden kelvin naar celcius
'FORMATROMAN                 - converteert een getal (integer) naar Romeins formaat
'LSHA2MMPD                   - converteert liter/seconde/ha naar mm/d
'MMPD2LSHA                   - converteert mm/d naar liter/seconde/ha
'M3PS2MMPD                   - converteert m3/s naar mm/d
'MMPD2M3PS                   - converteert mm/d naar m3/s
'MMPU2M3PS                   - converteert mm/u naar m3/s
'M3PS2MMPU                   - converteert m3/s naar mm/u

'geografisch
'RD2LATLONG                  - converteert een coordinaat in RD naar LAT/LONG
'RD2LAT                      - converteert een coordinaat in RD naar LAT
'RD2LON                      - converteert een coordinaat in RD naar LONG
'RD2WGS84                    - converteert een coordinaat in RD naar WGS84 (LAT/LONG)
'WGS842RD                    - converteert een WGS84-coordinaat (LAT/LONG) naar RD
'WGS842X                     - converteert een WGS84-coordinaat (LAT/LONG) naar RD, X-coordinaat
'WGS842Y                     - converteert een WGS84-coordinaat (LAT/LONG) naar RD, Y-coordinaat
'RD2BESSEL                   - van RD naar besselfunctie
'BESSEL2WGS84                - van Besselfunctie naar latlong
'WGS84DEG2DECIMAL            - converteert latlong in graden naar decimalen
'WGS84DEG2LATDECIMAL         - converteert latlong in graden naar latitude in decimalen
'WGS84DEG2LONDECIMAL         - converteert latlong in graden naar longitude in decimalen
'WGS842BESSEL                - van latlong naar besselfunctie
'BESSEL2RD                   - van Besselfunctie naar RD

'hydrologisch
'GETIJDEN_SINUS              - Berekent de waterstand van een getijdenslag voor elk gewenst tijdstip .
'QSTUW                       - Berekent het debiet over een rechthoekige stuw
'QABUTMENTBRIDGE             - berekent het debiet door een landhoofdbrug
'BREEDTE_STUW                - berekent de gewenste breedte van een stuw, gegeven de afvoer in mm/d en overstortende straal
'WEIRSUBMERGED               - Berekent of een stuw verdronken is
'QHEVEL                      - Berekent het debiet door een hevel
'QDUIKERRECHTHOEK            - Berekent het debiet door een rechthoekige duiker
'QDUIKER                     - Berekent het debiet door een duiker
'QORIFICE                    - Berekent het debiet door een schuif met gegeven dH, breedte en openingshoogte
'DHGEVULDERONDEDUIKER        - Berekent het verval over een ronde duiker die geheel gevuld is met water
'DHRONDEDUIKER               - Berekent het verval over een ronde duiker die al dan niet geheel gevuld is
'WIDTHORIFICE                - Berekent de benodigde breedte van een niet-verdronken onderlaat gegeven een gevraagd debiet en drempelhoogte
'GETAVGMINFROMTIDE           - haalt de gemiddelde laagwaterstand uit een getijdenreeks
'GETAVGMAXFROMTIDE           - haalt de gemiddelde hoogwaterstand uit een getijdenreeks
'WINDRICHTING                - geeft de windrichting terug (N, NO, O, ZO, etc.) als functie van de hoek in graden. Optioneel in graden (0,45,90,135,180,225,270,315,360)
'TIDALMINMAXFROMSERIES       - haalt per getijdenslag de hoogste en laagste waterstand binnen en schrijft deze naar een range
'TIDALLOWSFROMSERIES         - haalt per getijdenslag de laagste waterstand binnen en schrijft deze naar een range
'LGN5TONBW                   - converteert LGN code naar de benodigde landgebruikscode voor NBW-toetsing (zoals afgeleid voor waterschap Noorderzijlvest)
'LGN2SOBEK                   - converteert LGN code naar landgebruiksnummer in SOBEK (1=grass, 2=potatoes etc.)
'ERNSTRecord                 - schrijft een ERNST-record weg
'BOD2CAPSIM                  - converteert bodemcode (letter + cijfer) naar CAPSIM bodemtypenummer voor SOBEK
'GHGGLG2GT                   - converteert een gegeven GHG en GLG (m - maaiveld) naar Grondwatertrap
'HYDROZOMERWINTER            - berekent of een datum in de hydrologische zomer/winter valt
'EVAPDEBRUINKEIJMAN          - openwaterverdamping volgens de bruin-keijman
'EVAPMAKKINK                 - referentiegewasverdamping volgens Makkink
'MAKKINK2OPENWATER           - converteert makkinkverdamping naar openwaterverdamping
'OPENWATEREVAPFACTOR         - berekent gegeven de datum de 'gewasfactor' openwaterverdamping terug
'EVAPDAY2HOUR                - deaggregeert etmaalverdampingssom naar uurwaarden
'HOURLYEVAPORATIONFRACTION   - bepaalt, gegeven het uur van de dag, de fractie van de etmaalverdampingssom op basis sinus van 6 tot 18)
'HydraulicRadius             - berekent de hydraulische straal
'Manning2Chezy                      - Converteert n_manning naar chezy ruwheid
'Chezy2Manning                      - converteert chezy naar n_manning ruwheid
'Chezy                              - berekent Q voor een gegeven Chezy-waarde en bodemverhang
'MaatgevendeAfvoer                  - berekent de maatgevende afvoer op basis van neerslagintensiteit en oppervlak
'NEERSLAGPATROON                    - berekent het type neerslagpatroon volgens STOWA 2004 (Nieuwe neerslagstatistiek voor waterbeheerders)


'statistiek
'CORRELATIONBYWINDOW                - berekent de correlatiecoefficient r tussen twee ranges met gegeven window size
'EXPONENTIELEVERDELINGCDF           - berekent de cumulatieve voor de Exponentiele verdeling
'GUMBELFIT                          - fit de gumbel-kansverdeling aan een gegeven reeks van maxima, gebruikmakend van de Maximum Likelihood
'GUMBELVERDELINGSFUNCTIE            - berekent de ONDERschrijdingskans van een bepaalde parameterwaarde op basis van opgegeven GUMBEL-parameters en parameterwaarde X
'GUMBELINVERSE                      - berekent de parameterwaarde die hoort bij een gegeven ONDERschrijdinskans volgens de GUMBEL kansverdeling type I
'GENPARETOINVERSE                   - berekent (iteratief) de inverse van de Generalized Pareto CDF
'GENPARETOCDF                       - berekent de cumulatieve onderschrijdingskans volgens Gen. Pareto
'GEVVERDELINGSFUNCTIE               - berekent de ONDERschrijdingskans van een bepaalde parameterwaarde op basis van een opgegeven GEV-kansverdeling
'GEVINVERSE                         - berekent de parameterwaarde die hoort bij een gegeven ONDERschrijdingskans volgens de GEV-kansverdeling
'CALCNEERSLAGSTATS                  - berekent de statistische parameters (GEV-kansdichtheidsfunctie) van een bui, gegeven gebiedsoppervlak en neerslagduur
'CALCHERHALINGSTIJD                 - berekent gegeven neerslagduur (uren), gebiedsoppervlak (km2) en volume (mm) de herhalingstijd
'CALCNEERSLAGVOLUME                 - berekent gegeven herhalingstijd, neerslagduur en gebiedsoppervlak het bijbehorende neerslagvolume
'PRECIPITATIONAREAREDUCTION         - reduceert het neerslagvolume in een reeks als functie van duur, herhalingstijd en gebiedsoppervlak
'ANNUALMAXIMUMPRECIPITATIONEVENTS   - extraheert uit een neerslagreeks en een gegeven neerslagduur de jaarmaxima, zowel voor zomer, winter als jaarrond
'PLOTTINGPOSITIONFROMANNUALMAXIMA   - berekent de plotting position (herhalingstijd in jaren) voor een getal adhv een lijst met jaarmaxima
'IDENTIFYPRECIPITATIONEVENTSPOT     - extraheert uit een neerslagreeks gegeven een POT-waarde en neerslagduur de neerslagvolumes
'CLASSIFYEVENTS                     - schrijft rangnummers weg voor gebeurtenissen met gegeven duur die binnen een klasse > X en < Y vallen
'RANKNUMBEROFEXCEEDANCES            - schrijft de rangnummers van overschrijdingen van een vooraf opgegeven drempel naar een nieuwe kolom
'POTANALYSISSUM                     - indexeert de extreemste gebeurtenissen met vooraf opgegeven duur door hun som en indexnummers naar naastgelegen kolommen te schrijven
'POTANALYSISMAX                     - indexeert de extreemste gebeurtenissen met vooraf opgegeven duur door hun maximum en indexnummers naar naastgelegen kolommen te schrijven
'CALCULATEEXTREMEEVENTS             - extraheert uit een neerslagreeks alle zwaarste buien gegeven neerslagduur
'NASH_SUTCLIFFE                     - berekent de nash-sutcliffe-coëfficiënt voor twee reeksen
'NASH_SUTCLIFFE_FAST
'FILTERBASEFLOW                     - filtert baseflow of interflow uit een opgegeven reeks met afvoeren
'HOOGHOUDT_q                        - berekent de stationaire q voor een situatie met twee drains volgens de formule van Hooghoudt
'HOOGHOUDT_L                        - berekent de drainafstand tussen twee drains volgens de formule van Hooghoudt, gegeven q en opbolling
'YZHYDRAULICPROPERTIES              - berekent de hydraulische eigenschappen van een gegeven YZ-profiel: A, P en R as functie van de diepte

'sobek
'READSPECIFICHISRESULTS      - leest resultaten uit een HIS-file met vooraf opgegeven locatie en parameter
'READHISLOCPARTIM            - leest de Locaties, Parameters en Tijdstappen in van een HIS-file
'GETNODESTATSFROMSOBEK       - haalt voor alle objecten in calcpnt.his x,y,min,max,avg,first en last op
'MERGESTORAGETABLES          - voegt twee bergingstabellen samen (hoogte/oppervlak). Beide tabellen moeten een collection of clsLevelAreaPair zijn
'INTERPOLATEFROMSTORAGETABLE - interpoleert uit een bergingstabel (hoogte/oppervlak) invoer:hoogte, uitvoer:oppervlak. tabel moet collection of clsLevelAreaPair zijn
'PARSESOBEKFILE              - parst een sobek inputfile en schrijft het resultaat naar een vooraf opgegeven locatie
'PARSESOBEKTABLE             - parst een sobek tabel en geeft een array met de resultaten terug
'PARSEBYSINGLECHAR           - parst een string op basis van 1 karakter per keer
'MAKESOBEKTARGETLEVELTABLE   - maakt een tabel met zomer- en winterstreefpeilen
'READBUIFILE                 - leest een .bui file van SOBEK in en schrijft de data naar het werkblad
'WRITEBUIFILE                - leest neerslagdata van het werkblad en schrijft een .bui file voor SOBEK weg
'WRITERKSFILE                - leest neerslagdata van het werkblad en schrijft een .rks file voor SOBEK weg
'WRITEPRNFILE                - leest een tijdtabel van het werkblad en schrijft een .prn file voor SOBEK weg
'WRITEPRNFILES               - leest een tijdtabel van het werkblad en schrijft meerdere .prn files voor SOBEK weg
'WRITERRBOUNDARYDATA         - schrijft bound3b.3b en bound3b.tbl
'GETDELWAQID                 - Genereert het DELWAQ-ID gegeven het segmentnummer
'IDFROMSTRING                - extraheert een ID uit een string, gegeven een prefix en/of een afbreekstring
'REMOVEPOSTFIX               - verwijdert een postfix uit een string
'WRITESTOCHASTXMLFILE        - schrijft locaties en bijbehorende herhalingstijden en waterhoogtes weg in XML zodat de toetsingstool ze kan inlezen
'REPLACEDATESINSETTINGSDAT   - vervangt de start- en einddatum van een simulatie in het bestand settings.dat
'REPLACEDATESINDELFT3BINI    - vervangt de start- en einddatum van een simulatie in het bestand delft_3b.ini
'WriteUnpaved3BRecord        - schrijft een unpaved.3b record
'TrapeziumRangeToYZProfiles  - creëert dwarsprofielinformatie in YZ-formaat op basis van data in trapezium-kentallen (inclusief plasberm)

'overige modellen
'WRITEWAGMODINPUT            - schrijft een .dat file voor het wageningenmodel, met neerslag, verdamping en gemeten afvoeren (optioneel)
'READWAGMODOUTPUT            - leest een .00P-bestand van het Wageningen model in en schrijft de inhoud naar het actieve werkblad.
'WRITEPCRASTERXYZ            - schrijft een .xyz file ten behoeve van PCRASTER, die op zijn beurt weer een inundatiegrid kan opstellen

'meteofucties
'MAKKINKAVG                  - geeft voor een gegeven dag in het jaar de meerjarig gemiddelde potentiele gewasverdamping volgens Makkink terug
'DAYSTOHOURS                 - disaggregeert etmaalwaarden naar uurwaarden. Opties "none" (voor bijv. temperatuur) en "divide"
'EVAPDAYTOHOUR               - disaggregeert etmaalverdampingssommen tot uursommen, gebaseerd op een sinusoide
'NEERSLAGTEKORT              - berekent het neerslagtekort op basis van een tijdstap met neerslag, verdamping en het tekort van de vorige tijdstap
'HIRLAMTRANSLATE             - converteert HIRLAM-voorspellingsrasters met neerslag

'stringbewerkingen
'PARSESTRING                 - parst een string op basis van een te specificeren deelstring
'TEXTSNIPPET                 - deelt een string op in drie delen, gegeven twee karakterposities
'MULTIPARSE                  - parst in een keer het n'de element uit een string
'PARSENUMERIC                - parst net zo lang een karakter tot het volgende niet langer numeriek is
'BNASTRING                   - creeert een string voor een BNA-file. Vraagt om ID en X- en Y-coordinaat
'WAGMODSTASTRING             - creeert een meteo-string voor de .STA-file van het Wageningenmodel
'WALRUSDATSTRING             - creeert een meteo-string voor de .DAT-file van het WALRUS-model
'VERWIJDERDAGNAAMUITDATUM    - verwijdert de naam van de dag uit een string
'MAKEXMLTOKEN                - maakt van een tokenID en de waarde een tokenID="waarde" string
'STRINGPOSITIE               - geeft het positienummer van de eerstvoorkomende string van een opgegeven type op
'REPLACESTRING               - vervangt een opgegeven deelstring van een string door een andere string, dus niet op basis van positie
'REPLACESTRINGINALLFILES     - vervangt een string in alle files in de huidige directory, eventueel incl. subdirectories
'DOUBLEIDSINSTRINGCOLLECTION - checkt of een collectie met strings dubbele waarden bevat (boolean)
'TRIMUSINGCUSTOMSTRING       - voert een VBA.Trim uit met een opgegeven karakter ipv standaard de spatie
'UnifyString                 - uniformeert een string door te VBA.Trimmen en altijd de uppercase te gebruiken. Te gebruiken als Key in collections
'ISBANKNUMBER                - herkent of een string een bankrekeningnummer is
'MATCHWILDCARD               - checkt of een gegeven ID matcht met een gegeven structuur met wildcards
'CONCATENATECOMBINATIONS     - maakt een lijst met alle unieke stringcombinaties door te permuteren

'importeren van bestanden
'READHMCZDATA                               - Leest waterstanden van het Hydro Meteo Centrum (ASCII formaat) in
'READASCIIGRID                              - Leest een Arc/Info grid in
'WriteASCIIGridFromEquation                 - Schrijft een Arc/Info grid op basis van z = ax + by + c
'WriteASCIIGridFromMultipleEquations        - Schrijft een Arc/Info grid op basis van meerdere z = ax + by + c vergelijkingen
'WRITEASCIIGRID                             - Schrijft een Arc/Info grid
'ASCII2XYZ                                  - converteert een Arc/Info grid naar een bestand met XYZ-waardes
'READMT940                                  - leest een MT940-file in, dat rekeningoverzichten bevat (o.a. ABN-AMRO)
'READENTIRETEXTFILE                         - leest de volledige inhoud van een tekstbestand naar het geheugen

'GIS-bewerkingen
'JoinNodes                   - maakt een nieuwe knoopID aan voor meerdere xy-knopen als ze dicht genoeg bijeen liggen
'FindNearestObjectInRange    - zoekt het ID van het dichstbijzijnde object uit een lijst (bijv. Meteo-stations) op basis van XY-coordinaten

'bestanden
'OPENSINGLEFILE              - open file dialog box
'LISTFILESINFOLDER           - produceert een collection van alle bestanden in een directory
'DIRECTORYEXISTS             - geeft terug of een directory bestaat
'CONTAINSKEY                 - geeft terug of een gegeven key onderdeel uitmaakt van een collection (WEKT NIET!!!!)
'CONTAINSKEY_BYOBJECTID      - geeft terug of een gegeven ID onderdeel uitmaakt van een collection met objecten die een element ID hebben
'DELETESHAPEFILE             - verwijdert een shapefile inclusief zijn bijbehorende bestanden (shx, dbf, shp)
'MOVEFILE                    - verplaatst een bestand van dir1 naar dir2
'DIRECTORYCOPY               - kopieert een directory incl. subdirs en inhoud naar een andere dir.
'FOLDERBROWSER               - presenteert een folder browser dialog
'REPLACEINFILE               - vervangt een opgegeven string overal in een tekstbestand
'REPLACEINSTRING             - vervangt een opgegeven string overal in een string

'Binaire functies
'Binary to Hex               - BinToHex(BinNum As String)
'Binary to Octal             - BinToOct(BinNum As String)
'Binary to Decimal           - BinToDec(BinNum As String)
'Hex to Binary               - HexToBin(HexNum As String)
'Octal to Binary             - OctToBin(OctNum As String)
'Decimal to Binary           - DecToBin(DecNum As String)


'overig
'RUNDOEVENTS                 - voert de optie doEvents uit voor een opgegeven aantal seconden, zodat andere processen even de ruimte krijgen
'SLEEP                       - laat de uitvoering van de macro een gespecificeerd aantal miliseconden wachten
'SHELLANDWAIT                - voert executables via de command line uit en wacht tot ze klaar zijn
'FINANCIELECATEGORIE         - rubriceert op basis van omschrijving uitgaven en inkomsten
'FILEEXISTS                  - controleert of een bestand bestaat.
'IB2011                      - berekent ruwweg de inkomstenbelasting voor 2011 op basis van opgegeven bruto inkomen


Public Enum enmKVDParameter
    dispersiecoefficient = 0
    locatieparameter = 1
    schaalparameter = 2
    vormparameter = 3
    waarde = 4
    terugkeertijd = 5
End Enum


Public Enum enmAggregateMethod
  Average = 1
  Most = 2
  Smallest = 3
  Largest = 4
  first = 5
  Last = 6
  Sum = 7
End Enum

Private Const STATUS_PENDING = &H103&
Private Const PROCESS_QUERY_INFORMATION = &H400
Public Const pi As Double = 3.141592

Public Function Interpolate(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Optional BlockInterpolate As Boolean = False) As Double

Dim Y3 As Double 'de geïnterpoleerde waarde die we straks in de cel gaan zetten
If X3 < Minimum(X1, X2) Then
  Y3 = -999
ElseIf X3 > Maximum(X1, X2) Then
  Y3 = -999
Else
  If BlockInterpolate = True Then
    Interpolate = Y1
  Else
   Interpolate = Y1 + (Y2 - Y1) / (X2 - X1) * (X3 - X1)
  End If
End If
End Function

Public Function GoalSeekRange(GoalCellsRange As Range, GoalValuesRange As Range, AdjustRange As Range) As Boolean
    If GoalCellsRange.Rows.Count <> GoalValuesRange.Rows.Count Or GoalCellsRange.Rows.Count <> AdjustRange.Rows.Count Then
        MsgBox ("Error: number of rows must be equal for all ranges that are passed to function GoalSeekRange.")
        GoalSeekRange = False
    ElseIf GoalCellsRange.Columns.Count > 1 Or GoalValuesRange.Columns.Count > 1 Or AdjustRange.Columns.Count > 1 Then
        MsgBox ("Error: each range passed to the function GoalSeekRange must only consist of one column.")
        GoalSeekRange = False
    Else
        Dim i As Integer
        For i = 1 To GoalCellsRange.Rows.Count
            GoalCellsRange.Cells(i, 1).GoalSeek Goal:=GoalValuesRange.Cells(i, 1), ChangingCell:=AdjustRange.Cells(i, 1)
        Next
    End If
End Function

Public Function Extrapolate(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double) As Double
'extrapolates linearly

Dim Y3 As Double, Rico As Double
If X3 > X2 Then
  Rico = (Y2 - Y1) / (X2 - X1)
  Extrapolate = Y2 + (X3 - X2) * Rico
ElseIf X3 < X1 Then
  Rico = (Y2 - Y1) / (X2 - X1)
  Extrapolate = Y1 - (X1 - X3) * Rico
Else
  Extrapolate = -999
End If

End Function

Public Function FitLinear_a(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
  'creates a straight line between two XY-co-ordinates and returns a (from y = ax + b)
  FitLinear_a = (Y2 - Y1) / (X2 - X1)
End Function

Public Function FitLinear_b(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
  Dim A As Double
  A = (Y2 - Y1) / (X2 - X1)
  FitLinear_b = Y2 - A * X2
End Function

Public Function InterpolateFromRange(x As Double, XRange As Range, YRange As Range, Optional ExtrapolateBelow As Boolean = True, Optional ExtrapolateAbove As Boolean = True, Optional BlockInterpolation As Boolean = False, Optional CheckIfAscending As Boolean = True) As Variant
  Dim r As Long, r2 As Long, startr As Long, stepsize As Long
  If XRange.Count <> YRange.Count Then
    InterpolateFromRange = "Error: X and Y range must be of equal size."
    Exit Function
  ElseIf XRange.Columns.Count > 1 Then
    InterpolateFromRange = "Error: column for X values must consist of one column."
    Exit Function
  ElseIf YRange.Columns.Count > 1 Then
    InterpolateFromRange = "Error: column for Y values must consit of one column."
    Exit Function
  End If
  
  If CheckIfAscending Then
    If Not IsRangeAscending(XRange) Then
      InterpolateFromRange = "Error: column containing X values must be ascending."
      Exit Function
    End If
 End If

If x <= XRange(1, 1).Value Then
  If ExtrapolateBelow = True Then
    InterpolateFromRange = YRange(1, 1).Value
    Exit Function
  Else
    InterpolateFromRange = 0
    Exit Function
  End If
ElseIf x >= XRange(XRange.Count, 1).Value Then
  If ExtrapolateAbove = True Then
    InterpolateFromRange = YRange(XRange.Count, 1).Value
    Exit Function
  Else
    InterpolateFromRange = 0
    Exit Function
  End If
ElseIf XRange.Count > 1 Then
  
  If XRange.Count > 100000 Then
    stepsize = 10000
  ElseIf XRange.Count > 10000 Then
    stepsize = 1000
  ElseIf XRange.Count > 1000 Then
    stepsize = 100
  ElseIf XRange.Count > 100 Then
    stepsize = 10
  Else
    stepsize = 1
  End If

  For r = 1 To XRange.Count Step stepsize
    If XRange(r, 1).Value > x Or r > XRange.Count - stepsize Then
      startr = r - stepsize
      For r2 = startr To XRange.Count
        If x >= XRange(r2 - 1, 1).Value And x <= XRange(r2, 1).Value Then
          InterpolateFromRange = Interpolate(XRange(r2 - 1, 1).Value, YRange(r2 - 1, 1).Value, XRange(r2, 1).Value, YRange(r2, 1).Value, x, BlockInterpolation)
          Exit Function
        End If
      Next
    End If
  Next
Else
  InterpolateFromRange = "Error: outside range."
End If

End Function


Public Function InterpolateRangeFromRange(XYRange As Range, ResultsRange As Range, Optional BlockInterpolation As Boolean = False) As Variant
  
  Dim i As Long, j As Long
  Dim r As Long, c As Long
  Dim XLookup As Variant, CurX As Variant, NextX As Variant
  
  'first some checks:
  If XYRange.Columns.Count <> 2 Then
    MsgBox ("Input range must consist of two columns: one containing X-values; one containing Y-values")
  ElseIf ResultsRange.Columns.Count <> 2 Then
    MsgBox ("Results range must consist of two columns: one containing X-values; one for the computed Y-values")
  ElseIf RANGEVERTASCENDING(XYRange) = False Then
    MsgBox ("Input range must be ascending.")
  Else
  
    'read the input range
    Dim XYData As Variant
    ReDim XYData(XYRange.Rows.Count, XYRange.Columns.Count)
    XYData = XYRange
  
    'read the output range
    Dim Results As Variant
    ReDim Results(ResultsRange.Rows.Count, ResultsRange.Columns.Count)
    Results = ResultsRange
    
    For r = 1 To UBound(Results, 1)
      XLookup = Results(r, 1)
      
      If XLookup < XYData(1, 1) Then
        Results(r, 2) = XYData(1, 2)
      ElseIf XLookup > XYData(UBound(XYData, 1), 1) Then
        Results(r, 2) = XYData(UBound(XYData, 1), 2)
      Else
        For i = 1 To UBound(XYData, 1)
          CurX = XYData(i, 1)
          NextX = XYData(i + 1, 1)
          
          If CurX <= XLookup And NextX >= XLookup Then
            Results(r, 2) = Interpolate(XYData(i, 1), XYData(i, 2), XYData(i + 1, 1), XYData(i + 1, 2), XLookup, BlockInterpolation)
            Exit For
          End If
        Next
      End If
    Next
  End If
  
  Call PrintArray(Results, ResultsRange)

End Function

Public Function InterpolateFromRangePlus(ID As String, x As Double, IDRange As Range, XRange As Range, YRange As Range, Optional ExtrapolateBelow As Boolean = True, Optional ExtrapolateAbove As Boolean = True, Optional BlockInterpolation As Boolean = False, Optional CheckIfAscending As Boolean = True) As Variant
  Dim r As Long, r2 As Long, startr As Long, stepsize As Long
  If IDRange.Count <> XRange.Count Then
    InterpolateFromRangePlus = "Error: ID and X range must be of equal size."
    Exit Function
  ElseIf XRange.Count <> YRange.Count Then
    InterpolateFromRangePlus = "Error: X and Y range must be of equal size."
    Exit Function
  ElseIf XRange.Columns.Count > 1 Then
    InterpolateFromRangePlus = "Error: column for X values must consist of one column."
    Exit Function
  ElseIf YRange.Columns.Count > 1 Then
    InterpolateFromRangePlus = "Error: column for Y values must consit of one column."
    Exit Function
  End If
  
  Dim StartRow As Long, EndRow As Long
  Dim startfound As Boolean
    
  'first find the start- and endrow for the given ID
  For r = 1 To IDRange.Count
    If UCase(Trim(IDRange(r, 1))) = UCase(Trim(ID)) And startfound = False Then
      startfound = True
      StartRow = r
    ElseIf startfound = True And IDRange(r, 1) <> ID Then
      EndRow = r - 1
      Exit For
    End If
  Next
  
  If x <= XRange(StartRow, 1).Value Then
    If ExtrapolateBelow = True Then
      InterpolateFromRangePlus = YRange(StartRow, 1).Value
      Exit Function
    Else
      InterpolateFromRangePlus = 0
      Exit Function
    End If
  ElseIf x >= XRange(EndRow, 1).Value Then
    If ExtrapolateAbove = True Then
      InterpolateFromRangePlus = YRange(EndRow, 1).Value
      Exit Function
    Else
      InterpolateFromRangePlus = 0
      Exit Function
    End If
  ElseIf (EndRow - StartRow) > 1 Then
    For r = StartRow To EndRow
      If XRange(r, 1).Value > x Then
        InterpolateFromRangePlus = Interpolate(XRange(r - 1, 1).Value, YRange(r - 1, 1).Value, XRange(r, 1).Value, YRange(r, 1).Value, x, BlockInterpolation)
        Exit Function
      End If
    Next
  Else
    InterpolateFromRangePlus = "Error: outside range."
 End If

End Function

Public Function KleinsteKwadratenMethode(GemetenDatum As Range, GemetenWaarden, BerekendDatum As Range, BerekendWaarden As Range) As Double

'deze functie berekent het kleinstekwadratenverschil tussen een berekende en gemeten reeks
Dim D As Double, v As Double
Dim D1 As Double, v1 As Double
Dim D2 As Double, v2 As Double
Dim v3 As Double
Dim Sum As Double
Dim r As Long, c As Long
Dim r2 As Long, c2 As Long

Sum = 0
For r = 1 To GemetenDatum.Rows.Count
  D = GemetenDatum.Cells(r, 1)
  v = GemetenWaarden.Cells(r, 1)
  
  If D >= BerekendDatum.Cells(1, 1) And D <= BerekendDatum.Cells(BerekendDatum.Rows.Count, 1) Then
  
    For r2 = 1 To BerekendDatum.Rows.Count - 1
      D1 = BerekendDatum.Cells(r2, 1)
      v1 = BerekendWaarden.Cells(r2, 1)
      D2 = BerekendDatum.Cells(r2 + 1, 1)
      v2 = BerekendWaarden.Cells(r2 + 1, 1)
      
      If D1 <= D And D2 >= D Then
        v3 = Interpolate(D1, v1, D2, v2, D)
        Sum = Sum + (v3 - v) ^ 2
        Exit For
      End If
    Next
  
  End If
Next
KleinsteKwadratenMethode = Sum

End Function

Function IsStringArrayEmpty(anArray() As String)

Dim i As Integer
On Error Resume Next
i = UBound(anArray, 1)
If err.Number = 0 Then
    IsStringArrayEmpty = False
Else
    IsStringArrayEmpty = True
End If

End Function

Public Function GETARRAYSORTIDX(myArr() As Double) As Long()
    Dim IdxArr() As Long, DoneArr() As Boolean, i As Long
    ReDim IdxArr(LBound(myArr), UBound(myArr()))
    ReDim DoneArr(LBound(myArr()), UBound(myArr()))
        
    For i = LBound(myArr()) To UBound(myArr())
      IdxArr(i) = GETMAXIDXFROMARRAY(myArr, DoneArr)
      DoneArr(i) = True
    Next
    GETARRAYSORTIDX = IdxArr

End Function

Public Function GETMAXIDXFROMARRAY(ByRef myArr() As Double, ByRef DoneArr() As Boolean) As Long
  Dim i As Long, myMax As Double, myMaxIdx As Long
  myMax = -999999999
  For i = LBound(myArr()) To UBound(myArr())
    If myArr(i) >= myMax And DoneArr(i) = False Then
      myMax = myArr(i)
      myMaxIdx = i
    End If
  Next
  GETMAXIDXFROMARRAY = myMaxIdx
End Function


' This routine uses the "heap sort" algorithm to sort a VB collection.
' It returns the sorted collection.
' Author: Christian d'Heureuse (www.source-code.biz)
Public Function SortCollection(ByVal c As Collection) As Collection
   Dim n As Long: n = c.Count
   If n = 0 Then Set SortCollection = New Collection: Exit Function
   ReDim Index(0 To n - 1) As Long                    ' allocate index array
   Dim i As Long, m As Long
   For i = 0 To n - 1: Index(i) = i + 1: Next         ' fill index array
   For i = n \ 2 - 1 To 0 Step -1                     ' generate ordered heap
      Heapify c, Index, i, n
      Next
   For m = n To 2 Step -1                             ' sort the index array
      Exchange Index, 0, m - 1                        ' move highest element to top
      Heapify c, Index, 0, m - 1
      Next
   Dim c2 As New Collection
   For i = 0 To n - 1: c2.Add c.Item(Index(i)): Next  ' fill output collection
   Set SortCollection = c2
End Function
   
' Heapsort routine.
' Returns a sorted Index array for the Keys array.
' Author: Christian d'Heureuse (www.source-code.biz)
Public Function HeapSort(Keys)
   Dim Base As Long: Base = LBound(Keys)                    ' array index base
   Dim n As Long: n = UBound(Keys) - LBound(Keys) + 1       ' array size
   Dim Index() As Long
   ReDim Index(Base To Base + n - 1) As Long                ' allocate index array
   Dim i As Long, m As Long
   For i = 0 To n - 1: Index(Base + i) = Base + i: Next     ' fill index array
   For i = n \ 2 - 1 To 0 Step -1                           ' generate ordered heap
      Heapify Keys, Index, i, n
      Next
   For m = n To 2 Step -1
      Exchange Index, 0, m - 1                              ' move highest element to top
      Heapify Keys, Index, 0, m - 1
   Next
   HeapSort = Index
End Function

Public Function SortCollectionByKey(myCollection As Collection) As Long()
  'in order to sort a collection of items by its key we'll first create an array that contains all keys
  'then we'll sort that array using Christian d'Heureuse's Heapsort-routine, which we'll return
  'this means that the function will return an array that contains the index numbers for the sorted keys
  
  'IMPORTANT: IN VBA it is NOT possible to retrieve the actual key. Therefore make sure you also store the key
  'as an element of the object within the collection!
  
  Dim SortMe() As Variant, i As Long
  ReDim SortMe(1 To myCollection.Count)
    
  For i = 1 To myCollection.Count
    SortMe(i) = myCollection.Item(i).key
  Next
  
  SortCollectionByKey = HeapSort(SortMe)

End Function


Public Function Random(lowerbound As Integer, upperbound As Integer) As Integer
  'geeft een random getal terug tussen twee gespecificeerde boundarywaaren (hele getallen)
  Randomize
  Random = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Function

Public Function RandomDouble(lowerbound As Double, upperbound As Double) As Double
  
  'creeer een random integer tussen 0 en 32000
  Dim myRnd As Integer
  myRnd = Random(0, 32000)
  
  'transformeer deze terug naar een waarde tussen min en max
  RandomDouble = lowerbound + myRnd / 32000 * (upperbound - lowerbound)

End Function



Public Function Maximum(val1 As Double, val2 As Double) As Double
  If val1 > val2 Then
    Maximum = val1
  Else
    Maximum = val2
  End If
End Function

Public Function Minimum(val1 As Double, val2 As Double) As Double
  If val1 < val2 Then
    Minimum = val1
  Else
    Minimum = val2
  End If
End Function

Public Function ARRAYFROMWORKSHEET(SheetName As String, StartRow As Long, StartCol As Long, EndCol As Long) As Variant()
  'Author: Siebe Bosch
  'Date : 1-9-2013
  'Description: extracts data from a worksheet and puts it into an array
  Dim curSheet As String, i As Long, j As Long, r As Long, c As Long
  Dim EndRow As Long
  Dim myArray() As Variant
  curSheet = ActiveSheet.Name
  Worksheets(SheetName).Activate
  r = StartRow - 1
  
  'find the last record
  While Not ActiveSheet.Cells(r + 1, StartCol) = ""
    r = r + 1
  Wend
  EndRow = r
  ReDim myArray(1 To EndRow - StartRow + 1, 1 To EndCol - StartCol + 1)
      
  r = 0
  For i = StartRow To EndRow
    c = 0
    r = r + 1
    For j = StartCol To EndCol
      c = c + 1
      myArray(r, c) = ActiveSheet.Cells(i, j)
    Next
  Next

  Worksheets(curSheet).Activate
  ARRAYFROMWORKSHEET = myArray

End Function

Public Sub ARRAYVARIANTTOWORKSHEET(SheetName As String, myArray() As Variant, StartRow As Long, StartCol As Long)
  
  Dim curSheet As String, Header As String, r As Long, c As Long, i As Long, j As Long
  curSheet = ActiveSheet.Name
  Worksheets(SheetName).Activate
    
  'write the data to the worksheet
  r = StartRow - 1
  c = StartCol
  For i = 1 To UBound(myArray, 1)
    r = r + 1
    c = StartCol
    For j = 1 To UBound(myArray, 2)
      c = c + 1
      ActiveSheet.Cells(r, c) = myArray(i, j)
    Next
  Next
  
  Worksheets(curSheet).Activate

End Sub

Public Sub ARRAYDATETOWORKSHEET(SheetName As String, myArray() As Date, StartRow As Long, StartCol As Long)
  
  Dim curSheet As String, Header As String, r As Long, c As Long, i As Long, j As Long
  curSheet = ActiveSheet.Name
  Worksheets(SheetName).Activate
    
  'write the data to the worksheet
  r = StartRow - 1
  c = StartCol
  For i = 1 To UBound(myArray, 1)
    r = r + 1
    c = StartCol
    ActiveSheet.Cells(r, c) = myArray(i)
  Next
  
  Worksheets(curSheet).Activate

End Sub

Public Sub ARRAYSINGLETOWORKSHEET(SheetName As String, myArray() As Single, StartRow As Long, StartCol As Long)
  
  Dim curSheet As String, Header As String, r As Long, c As Long, i As Long, j As Long
  curSheet = ActiveSheet.Name
  Worksheets(SheetName).Activate
    
  'write the data to the worksheet
  r = StartRow - 1
  c = StartCol
  For i = 1 To UBound(myArray, 1)
    r = r + 1
    c = StartCol
    ActiveSheet.Cells(r, c) = myArray(i)
  Next
  
  Worksheets(curSheet).Activate

End Sub


Public Sub TIMESERIES2ARRAYS(myRange As Range, ByRef Dates() As Date, ByRef Vals() As Single)
  Dim r As Long
  ReDim Dates(1 To myRange.Rows.Count)
  ReDim Vals(1 To myRange.Rows.Count)
  
  For r = 1 To myRange.Rows.Count
    Dates(r) = myRange.Cells(r, 1)
    Vals(r) = myRange.Cells(r, 2)
  Next

End Sub

Public Function USERSELECTRANGE() As Range
  Set USERSELECTRANGE = Application.InputBox(Prompt:="Please Select Range", Title:="Range Select", Type:=8)
  Call Application.GoTo(USERSELECTRANGE)
End Function

Public Function RANGEADDRESSFROMRC(ByRef SheetName As String, r1 As Long, C1 As Long, r2 As Long, c2 As Long) As Range
  Dim MySheet As Worksheet
  Dim i As Long
  For i = 1 To Worksheets.Count
    If Worksheets(i).Name = SheetName Then
      Set RANGEADDRESSFROMRC = Worksheets(i).Range(Cells(r1, C1).Address, Cells(r2, c2).Address)
    End If
  Next
End Function

Public Function RANGECOLIDXFROMMAXVAL(myRange As Range) As Long
  'Returns the column indexnumber for the largest value in a given range
  'note: this is not the worksheet column number but the column number from within the range!
  Dim myMax As Variant, myVal As Variant, myCol As Long
  Dim r As Long, c As Long
  myMax = -99999999999999#
  
  For c = 1 To myRange.Columns.Count
    For r = 1 To myRange.Rows.Count
      myVal = myRange.Cells(r, c).Value
      If myVal >= myMax Then
        myCol = c
        myMax = myVal
      End If
    Next
  Next
  RANGECOLIDXFROMMAXVAL = myCol

End Function

Public Function ASSIGNVALUEBYMONTH(myDate As Date, Jan As Double, Feb As Double, Mar As Double, Apr As Double, May As Double, Jun As Double, Jul As Double, Aug As Double, Sep As Double, Oct As Double, Nov As Double, Dec As Double) As Double
  Dim myMonth As Integer, myVal As Double
  myMonth = Month(myDate)
  Select Case myMonth
    Case Is = 1
      myVal = Jan
    Case Is = 2
      myVal = Feb
    Case Is = 3
      myVal = Mar
    Case Is = 4
      myVal = Apr
    Case Is = 5
      myVal = May
    Case Is = 6
      myVal = Jun
    Case Is = 7
      myVal = Jul
    Case Is = 8
      myVal = Aug
    Case Is = 9
      myVal = Sep
    Case Is = 10
      myVal = Oct
    Case Is = 11
      myVal = Nov
    Case Is = 12
      myVal = Dec
  End Select
  ASSIGNVALUEBYMONTH = myVal
End Function

Public Function CLASSIFYNUMBERBYCLASS(myVal As Double, ByVal minVal As Double, ByVal maxVal As Double, ByVal classSize As Double) As String
  Dim Done As Boolean, curVal As Double
  curVal = minVal
  
  If myVal < curVal Then
    CLASSIFYNUMBERBYCLASS = "< " & curVal
    Done = True
    Exit Function
  End If
  
  While Not Done
    curVal = curVal + classSize
    If myVal < curVal Then
      CLASSIFYNUMBERBYCLASS = curVal - classSize & " - " & curVal
      Done = True
      Exit Function
    End If
  Wend
  
  If Done = False Then
      CLASSIFYNUMBERBYCLASS = curVal - classSize & " - " & curVal
  End If
  
End Function

Public Sub DESIGNTEMPORALBLOCKPATTERNS(nTimesteps As Integer, nBlocks As Integer, PercentageIncrement As Integer, StartRow As Integer, EndRow As Integer)
    Dim PercentageSum As Double
    Dim Mileage() As Integer, i As Integer, Sum As Integer
    Dim r As Integer, c As Integer
    r = StartRow
    c = StartCol

    If Math.Round(nTimesteps / nBlocks, 0) <> nTimesteps / nBlocks Then
        MsgBox ("Error: choose another number of blocks.")
    Else
        Dim nBlockSteps As Integer, Percentage As Integer
        nBlockSteps = nTimesteps / nBlocks
        ReDim Mileage(1 To nBlocks)
                        
        While MileageOneUp(0, 100, Mileage)
            
            Sum = 0
            For i = 1 To Mileage.Count
                Sum = Sum + Mileage(i)
            Next
            
            If Sum = 100 Then
                c = c + 1
                r = StartRow
                For iblock = 1 To nBlocks
                    For iStep = 1 To nBlockSteps
                        r = r + 1
                        ActiveSheet.Cells(r, StartCol) = "Fraction"
                        ActiveSheet.Cells(r, c) = Mileage(iblock)
                    Next
                Next
                
            End If
            
        Wend
                        
                    
    End If
    
    
End Sub

Public Sub DESIGNTWOPARTTEMPORALPATTERNS(nTimesteps As Integer, PercentageIncrement As Integer, TimestepIncrement As Integer, StartRow As Integer, StartCol As Integer)
    'this function designs temporal patterns that consist of two lineair sections: start to mid and mid to end
    'hereby the start and end do NOT necessarily have to be 0
    'we'll start by incrementing the centerpoint in time
    
    '...............´;&%/.I%%&?í´...................................................'
    '...........,»%%%%%%/.I%%%%%%%%%&?í,............................................
    '.......^=%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%=;'.....................................
    '.....,''''''''''''',.I%%%%%%%%%%%%%%%%%%%%%%%%*/'..............................
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/^.......................
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%I»~´...............
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%&?~´........
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%&=í,.
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%=.
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%?.
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%?.
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%?.
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%?.
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%?.
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%?.
    Dim Ts As Integer
    Dim LeftTSFraction As Double, RightTSFraction As Double
    Dim p1Done As Boolean, p2Done As Boolean, p3Done As Boolean
    Dim p1 As Double, p2 As Double, p3 As Double
    Dim r As Integer, c As Integer
    Dim SecondPartRemaining As Double
    Dim patnum As Integer, Check As Double
    
    c = StartCol
    
    For Ts = TimestepIncrement To (nTimesteps - TimestepIncrement) Step TimestepIncrement
        'now we can iterate through all possibilities for the first point
            
        LeftTSFraction = Ts / nTimesteps
        RightTSFraction = (nTimesteps - Ts) / nTimesteps
            
        p1Done = False
        While Not p1Done
            
            'now that p1 is known we can iterate through p2 (the center point)
            p2Done = False
            p2 = 0
            While Not p2Done
                            
              'now that p2 is known we can calculate p3 (right point)
              'each part consists of a rectangle + a triangle on top
              Dim FirstPart As Double, SecondPart As Double
              FirstPart = TrapeziumArea(p1, p2, LeftTSFraction)
              SecondPart = 1 - FirstPart
              If SecondPart < 0 Then
                'the current combination of p1 with all possible values for p2 has been fully investigated, so move on
                p2Done = True
                p1Done = True
              Else
                'first decide if p2 is the hightest or lowest for our second part
                If p2 * RightTSFraction <= SecondPart Then
                    'we now know that p3 must lie higher than p2
                    SecondPartRemaining = SecondPart - RightTSFraction * p2
                    '(p3 - p2) * RightTSFraction = SecondPartRemaining
                    'p3 - p2 = SecondPartRemaining / RightTSFraction
                    'p3 = p2 + SecondPartRemaining / RightTSFraction
                    p3 = p2 + SecondPartRemaining / RightTSFraction * 2
                Else
                    'we now know that p3 lies lower than p2
                    '(p2 * RightTSFraction) + (p3 - p2) * RightTSFraction = SecondPart
                    'RightTSFraction * (p2 + (p3 - p2)/2) = SecondPart
                    '(p2 + (p3 - p2)/2) = SecondPart/RightTSFraction
                    '(p3 - p2)/2 = SecondPart/RightTSFraction - p2
                    'p3/2 - p2/2 = SecondPart/RightTSFraction - p2
                    'p3/2 = SecondPart/RightTSFraction - p2 + p2/2
                    'p3/2 = SecondPart/RightTSFraction - p2/2
                    'p3 = 2 * (SecondPart/RightTSFraction - p2/2)
                    p3 = 2 * (SecondPart / RightTSFraction - p2 / 2)
                End If
                
                'final check!
                Check = TrapeziumArea(p1, p2, LeftTSFraction) + TrapeziumArea(p2, p3, RightTSFraction)
                If Round(Check, 2) <> 1 Then Stop
                
                If p3 < 0 Then
                    'for the current p2 we have reached the limit regarding p3
                    p2Done = True
                Else
                    'write the result!
                    patnum = patnum + 1
                    ActiveSheet.Cells(StartRow, StartCol) = "ID"
                    ActiveSheet.Cells(StartRow, c) = "pattern" & patnum
                    ActiveSheet.Cells(StartRow + 1, StartCol) = "timestep peak"
                    ActiveSheet.Cells(StartRow + 1, c) = Ts
                    ActiveSheet.Cells(StartRow + 2, StartCol) = "fractionFirst"
                    ActiveSheet.Cells(StartRow + 2, c) = p1
                    ActiveSheet.Cells(StartRow + 3, StartCol) = "fractionPeak"
                    ActiveSheet.Cells(StartRow + 3, c) = p2
                    ActiveSheet.Cells(StartRow + 4, StartCol) = "fractionLast"
                    ActiveSheet.Cells(StartRow + 4, c) = p3
                    ActiveSheet.Cells(StartRow + 5, StartCol) = "Checksum"
                    ActiveSheet.Cells(StartRow + 5, c) = Check
                    c = c + 1
                End If
                
                
              End If
              p2 = p2 + PercentageIncrement / 100
            Wend
            
          p1 = p1 + PercentageIncrement / 100
      Wend
        
    
    Next

End Sub

Public Function TrapeziumArea(Y1 As Double, Y2 As Double, Width As Double) As Double
    TrapeziumArea = (Minimum(Y1, Y2) + (Maximum(Y1, Y2) - Minimum(Y1, Y2)) / 2) * Width
End Function

Public Function USERSELECTCELL() As Range
  Set USERSELECTCELL = Application.InputBox(Prompt:="Select first Data Cell", Title:="Cell Select", Type:=8)
  Call Application.GoTo(USERSELECTCELL)
End Function

Public Function SVGICONHEADER(Optional ByVal ScaleFactor As Double = 1) As String
    'in de basis is een SVG-icon 20 x 20 pixels. Met de ScaleFactor kun je hem optioneel groter maken
    SVGICONHEADER = "<svg width='" & 20 * ScaleFactor & "' height='" & ScaleFactor * 20 & "' xmlns='http://www.w3.org/2000/svg'>"
End Function

Public Function SVGICONFOOTER() As String
    SVGICONFOOTER = "</svg>"
End Function

Public Function SVGTRAPEZIUMSHARP(Optional ByVal ScaleFactor As Double = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Double = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    SVGTRAPEZIUMSHARP = SVGICONHEADER(ScaleFactor) & SVGSHAPETRAPEZIUMSHARP(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGICONFOOTER
End Function

Public Function SVGTRAPEZIUMBLUNT(Optional ByVal ScaleFactor As Double = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Double = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    SVGTRAPEZIUMBLUNT = SVGICONHEADER(ScaleFactor) & SVGSHAPETRAPEZIUMBLUNT(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGICONFOOTER
End Function

Public Function SVGCULVERT(Optional ByVal ScaleFactor As Double = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Double = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    SVGCULVERT = SVGICONHEADER(ScaleFactor) & SVGSHAPETRAPEZIUMBLUNT(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGSHAPECIRCLE(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGICONFOOTER
End Function

Public Function SVGABUTMENTBRIDGE(Optional ByVal ScaleFactor As Double = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Double = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    SVGABUTMENTBRIDGE = SVGICONHEADER(ScaleFactor) & SVGSHAPETRAPEZIUMSHARP(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGSHAPEABUTMENTBRIDGE(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGICONFOOTER
End Function

Public Function SVGPILLARBRIDGE(Optional ByVal ScaleFactor As Double = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Double = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    SVGPILLARBRIDGE = SVGICONHEADER(ScaleFactor) & SVGSHAPETRAPEZIUMSHARP(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGSHAPEPILLARBRIDGE(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGICONFOOTER
End Function

Public Function SVGRECTANGULARWEIR(Optional ByVal ScaleFactor As Double = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Double = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    SVGRECTANGULARWEIR = SVGICONHEADER(ScaleFactor) & SVGSHAPETRAPEZIUMSHARP(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGSHAPERECTANGULARWEIR(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGICONFOOTER
End Function

Public Function SVGORIFICE(Optional ByVal ScaleFactor As Double = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Double = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    SVGORIFICE = SVGICONHEADER(ScaleFactor) & SVGSHAPETRAPEZIUMSHARP(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGSHAPEORIFICE(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGICONFOOTER
End Function

Public Function SVGPUMP(Optional ByVal ScaleFactor As Double = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Double = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    Dim SVGSTRING As String
    SVGSTRING = SVGICONHEADER(ScaleFactor)
    SVGSTRING = SVGSTRING & SVGSHAPECIRCLE(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth)
    SVGSTRING = SVGSTRING & SVGSHAPEPUMP(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth)
    SVGPUMP = SVGSTRING & SVGICONFOOTER
End Function

Public Function SVGFISH(Optional ByVal ScaleFactor As Double = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Double = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    Dim SVGSTRING As String
    SVGSTRING = SVGICONHEADER(ScaleFactor)
    SVGSTRING = SVGSTRING & SVGSHAPETRAPEZIUMSHARP(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth)
    SVGSTRING = SVGSTRING & SVGSHAPEFISH(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth)
    SVGFISH = SVGSTRING & SVGICONFOOTER
End Function


Public Function SVGSHAPEPUMP(Optional ByVal ScaleFactor As Double = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Double = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    Dim Schoep1 As String, Schoep2 As String, Schoep3 As String
    Dim SVGSTRING As String
    'definitie van een Bézier curve in SVG: C X1 Y1 X2 Y2 x y waarbij: X1,Y1 = handle van het startpunt, X2,Y2 = handle van het eindpunt, x,y = eindpunt
    
    Dim i As Integer
    Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Xstart As Double, Ystart As Double, Xend As Double, Yend As Double
        
    For i = 0 To 5
        Xstart = 10 * ScaleFactor
        Ystart = 5 * ScaleFactor
        Xend = (10 + 3) * ScaleFactor
        Yend = Ystart
        X1 = Xstart + 0.5 * ScaleFactor
        Y1 = Ystart + 0.5 * ScaleFactor
        X2 = Xend - 0.5 * ScaleFactor
        Y2 = Yend + 0.5 * ScaleFactor
        
        'rotate the endpoint and the handles
        Call RotatePoint(Xend, Yend, Xstart, Ystart, 360 / 6 * i, Xend, Yend)
        Call RotatePoint(X1, Y1, Xstart, Ystart, 360 / 5 * i, X1, Y1)
        Call RotatePoint(X2, Y1, Xstart, Ystart, 360 / 5 * i, X2, Y2)
        
        SVGSTRING = SVGSTRING & "<path d='M " & Xstart & " " & Ystart & " C " & X1 & " " & Y1 & " " & X2 & " " & Y2 & " " & Xend & " " & Yend & " ' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "' fill='transparent'/>"
    Next
    
    SVGSHAPEPUMP = SVGSTRING
End Function



Public Function SVGSHAPEFISH(Optional ByVal ScaleFactor As Double = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Double = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    Dim SVGSTRING As String
    Dim Xstart As Double, Xend As Double, Ystart As Double, Yend As Double, X1 As Double, X2 As Double, Y1 As Double, Y2 As Double
    
    'buik
    Xstart = 5 * ScaleFactor
    Ystart = (5 + 2) * ScaleFactor
    X1 = (5 + 1) * ScaleFactor
    Y1 = 5 * ScaleFactor
    Xend = 15 * ScaleFactor
    Yend = 5 * ScaleFactor
    X2 = (15 - 2) * ScaleFactor
    Y2 = (5 - 4) * ScaleFactor
    SVGSTRING = "<path d='M " & Xstart & " " & Ystart & " C " & X1 & " " & Y1 & " " & X2 & " " & Y2 & " " & Xend & " " & Yend
    
    'rug
    Xstart = 5 * ScaleFactor
    Ystart = (5 - 2) * ScaleFactor
    X1 = (5 + 1) * ScaleFactor
    Y1 = 5 * ScaleFactor
    Xend = 15 * ScaleFactor
    Yend = 5 * ScaleFactor
    X2 = (15 - 2) * ScaleFactor
    Y2 = (5 + 4) * ScaleFactor
    SVGSTRING = SVGSTRING & " C " & X2 & " " & Y2 & " " & X1 & " " & Y1 & " " & Xstart & " " & Ystart & " Z'" & " stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "' fill-opacity='" & FillOpacity & "' fill='" & FillColor & "'/>"
    
    SVGSHAPEFISH = SVGSTRING
End Function


Public Function SVGSHAPEORIFICE(Optional ByVal ScaleFactor As Double = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Double = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    'we schematize an orifice using a shape and a rectangle
    Dim SVGSTRING As String
    SVGSTRING = "<path d='M 0 0 l " & 7 * ScaleFactor & " " & 0 & " l 0 " & 8 * ScaleFactor & " l " & 6 * ScaleFactor & " 0 l 0 " & -8 * ScaleFactor & " l " & 7 * ScaleFactor & " 0 l " & -5 * ScaleFactor & " " & 10 * ScaleFactor & " l " & -10 * ScaleFactor & " 0 " & " Z' fill-opacity='" & FillOpacity & "' fill='" & FillColor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>" & "<path d='M " & 7 * ScaleFactor & " " & 5 * ScaleFactor & " l 0 " & 5 * ScaleFactor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>" & "<path d='M " & 13 * ScaleFactor & " " & 5 * ScaleFactor & " l 0 " & 5 * ScaleFactor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>"
    SVGSHAPEORIFICE = SVGSTRING & "<path d='M " & 7 * ScaleFactor & " 0 l 0 " & 4 * ScaleFactor & " l " & 6 * ScaleFactor & " 0 l 0 " & -4 * ScaleFactor & " Z' fill-opacity='" & FillOpacity & "' fill='" & FillColor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>" & "<path d='M " & 7 * ScaleFactor & " " & 5 * ScaleFactor & " l 0 " & 5 * ScaleFactor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>" & "<path d='M " & 13 * ScaleFactor & " " & 5 * ScaleFactor & " l 0 " & 5 * ScaleFactor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>"
End Function

Public Function SVGSHAPERECTANGULARWEIR(Optional ByVal ScaleFactor As Double = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Double = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    'we schematize a rectangular weir using a shape and two vertical stripes
    SVGSHAPERECTANGULARWEIR = "<path d='M " & 2 * ScaleFactor & " " & 4 * ScaleFactor & " l " & 3 * ScaleFactor & " " & 6 * ScaleFactor & " l " & 10 * ScaleFactor & " 0 l " & 3 * ScaleFactor & " " & -6 * ScaleFactor & " l " & -5 * ScaleFactor & " 0 l 0 " & 3 * ScaleFactor & " l " & -6 * ScaleFactor & " 0 l 0 " & -3 * ScaleFactor & " Z' fill-opacity='" & FillOpacity & "' fill='" & FillColor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>" & "<path d='M " & 7 * ScaleFactor & " " & 5 * ScaleFactor & " l 0 " & 5 * ScaleFactor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>" & "<path d='M " & 13 * ScaleFactor & " " & 5 * ScaleFactor & " l 0 " & 5 * ScaleFactor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>"
End Function

Public Function SVGSHAPEPILLARBRIDGE(Optional ByVal ScaleFactor As Double = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Double = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    'a typical trapezium would be width 20 and height 10
    SVGSHAPEPILLARBRIDGE = "<path d='M 0 0 l " & 3 * ScaleFactor & " " & 6 * ScaleFactor & " l 0 " & -4 * ScaleFactor & " l " & 6 * ScaleFactor & " 0 l 0 " & 8 * ScaleFactor & " l " & 2 * ScaleFactor & " 0 " & " l 0 " & -8 * ScaleFactor & " l " & 6 * ScaleFactor & " 0 l " & " 0 " & 4 * ScaleFactor & " l " & 3 * ScaleFactor & " " & -6 * ScaleFactor & " Z' fill-opacity='" & FillOpacity & "' fill='" & FillColor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>"
End Function

Public Function SVGSHAPEABUTMENTBRIDGE(Optional ByVal ScaleFactor As Double = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Double = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    'a typical trapezium would be width 20 and height 10
    SVGSHAPEABUTMENTBRIDGE = "<path d='M 0 0 l " & 3 * ScaleFactor & " " & 6 * ScaleFactor & " l 0 " & -4 * ScaleFactor & " l " & 14 * ScaleFactor & " " & 0 & " l " & " 0 " & 4 * ScaleFactor & " l " & 3 * ScaleFactor & " " & -6 * ScaleFactor & " Z' fill-opacity='" & FillOpacity & "' fill='" & FillColor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>"
End Function

Public Function SVGSHAPECIRCLE(Optional ByVal ScaleFactor As Double = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Double = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    SVGSHAPECIRCLE = "<circle cx='" & 10 * ScaleFactor & "' cy='" & 5 * ScaleFactor & "' r='" & 3 * ScaleFactor & "' fill-opacity='0.5' fill='kleur' stroke='kleur' stroke-width='" & StrokeWidth & "'/>"
End Function

Public Function SVGSHAPETRAPEZIUMSHARP(Optional ByVal ScaleFactor As Double = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Double = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    'a typical trapezium would be width 20 and height 10
    SVGSHAPETRAPEZIUMSHARP = "<path d='M 0 0 l " & 5 * ScaleFactor & " " & 10 * ScaleFactor & " l " & 10 * ScaleFactor & " 0 " & " l " & 5 * ScaleFactor & " " & -10 * ScaleFactor & " Z' fill-opacity='" & FillOpacity & "' fill='" & FillColor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>"
End Function

Public Function SVGSHAPETRAPEZIUMBLUNT(Optional ByVal ScaleFactor As Double = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Double = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    'a typical trapezium would be width 20 and height 10
    SVGSHAPETRAPEZIUMBLUNT = "<path d='M " & 1 * ScaleFactor & " 0 l 0 " & 2 * ScaleFactor & " l " & 4 * ScaleFactor & " " & 8 * ScaleFactor & " l " & 10 * ScaleFactor & " " & 0 & " l " & 4 * ScaleFactor & " " & -8 * ScaleFactor & " l " & 0 & " " & -2 * ScaleFactor & " Z' fill-opacity='" & FillOpacity & "' fill='" & FillColor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>"
End Function

Public Function MAXFROMCOLLECTION(myColl As Collection) As Double
  Dim Val As Double, max As Double, i As Long
  max = -999999999999#
  For i = 1 To myColl.Count
    Val = myColl(i)
    If Val > max Then max = Val
  Next
  MAXFROMCOLLECTION = max
End Function

Public Function MINFROMCOLLECTION(myColl As Collection) As Double
  Dim Val As Double, Min As Double, i As Long
  Min = 999999999999#
  For i = 1 To myColl.Count
    Val = myColl(i)
    If Val < Min Then Min = Val
  Next
  MINFROMCOLLECTION = Min
End Function

Public Function AVGFROMCOLLECTION(myColl As Collection) As Double
  Dim Val As Double, Sum As Double, i As Long
  For i = 1 To myColl.Count
    Val = myColl(i)
    Sum = Sum + Val
  Next
  AVGFROMCOLLECTION = Sum / myColl.Count
End Function


Public Function GENPARETOPDF(mu As Double, Sigma As Double, kappa As Double, x As Double) As Double
  'calculates the cumulative probability density according to the Generalized Pareto probability distribution
  Dim par As Double
  par = (x - mu) / Sigma

    GENPARETOPDF = 1 / Sigma * (1 + kappa * par) ^ -(1 / kappa + 1)
   
End Function

Public Function GENPARETOCDF(mu As Double, Sigma As Double, kappa As Double, x As Double) As Double
  'calculates the cumulative probability density according to the Generalized Pareto probability distribution
  Dim par As Double
  par = (x - mu) / Sigma

  If kappa = 0 Then
    GENPARETOCDF = 1 - Exp(-par)
  Else
    GENPARETOCDF = 1 - (1 + kappa * par) ^ (-1 / kappa)
  End If
   
End Function

Public Function CONDWEIBULLCDF(alpha As Double, beta As Double, gamma As Double, x As Double) As Double
  'calculates the cumulative probability density according to the Conditional Weibull probability distribution
  CONDWEIBULLCDF = 1 - Math.Exp(-((x - gamma) / beta) ^ alpha)
End Function

Function BerekenStochastVolumeKlasse(rKlasseFreq As Range, rStochastInUse As Range, rCurCell As Range) As Variant

Dim rCell As Range
Dim vResult
Dim CurRow As Long, i As Long, n As Long, r As Long, r1 As Long
Dim StartRow As Long, EndRow As Long
Dim Inuse() As Boolean, Freq() As Double
Dim Done As Boolean, rad As Integer
Dim rLow As Integer, rHigh As Integer

CurRow = rCurCell.Row
n = rKlasseFreq.Count
StartRow = rKlasseFreq.Row
EndRow = StartRow + n - 1
ReDim Inuse(StartRow To EndRow)
ReDim Freq(StartRow To EndRow)

'Kijk welke stochasten in gebruik zijn
r = StartRow - 1
For Each rCell In rStochastInUse
  r = r + 1
  If rCell.Value = "a" Then
    Inuse(r) = True
  Else
    Inuse(r) = False
  End If
Next

'inventariseer voor iedere klasse de frequentie
r = StartRow - 1
For Each rCell In rKlasseFreq
  r = r + 1
  Freq(r) = rCell.Value
Next

'doorloop de range en zoek bij inactieve cellen naar de dichtstbijzijnde actieve broeders
r = StartRow - 1
For Each rCell In rKlasseFreq
  r = r + 1
  
  If Inuse(r) = False Then
    rLow = 0
    rHigh = 0
    For r1 = r - 1 To StartRow Step -1
      If Inuse(r1) Then
        rLow = r1
        Exit For
      End If
    Next
      
    For r1 = r + 1 To EndRow
      If Inuse(r1) Then
        rHigh = r1
        Exit For
      End If
    Next
  
    'herverdeel de frequentie van de inactieve klasse
    If rLow = 0 Then
      Freq(rHigh) = Freq(rHigh) + Freq(r)
      Freq(r) = 0
    ElseIf rHigh = 0 Then
      Freq(rLow) = Freq(rLow) + Freq(r)
      Freq(r) = 0
    ElseIf Math.Abs(rHigh - r) = Math.Abs(r - rLow) Then
      'divide equally
      Freq(rHigh) = Freq(rHigh) + Freq(r) / 2
      Freq(rLow) = Freq(rLow) + Freq(r) / 2
      Freq(r) = 0
    ElseIf Math.Abs(rHigh - r) > Math.Abs(r - rLow) Then
      'low is nearest so assign all frequency to that one
      Freq(rLow) = Freq(rLow) + Freq(r)
      Freq(r) = 0
    ElseIf Math.Abs(rHigh - r) < Math.Abs(rLow - r) Then
      Freq(rHigh) = Freq(rHigh) + Freq(r)
      Freq(r) = 0
    End If
  End If
Next

If Freq(CurRow) = 0 Then
  BerekenStochastVolumeKlasse = ""
Else
  BerekenStochastVolumeKlasse = Freq(CurRow)
End If


End Function

Function BerekenStochastPatroonKlasse(rPatroonNaam As Range, rPatroonKans As Range, rStochastInUse As Range, rCurCell As Range) As Variant

Dim Inuse(1 To 7) As Boolean
Dim Kans(1 To 7) As Double
Dim Naam(1 To 7) As String

Dim StartCol As Integer
Dim c As Integer, C1 As Integer
Dim rCell As Range
Dim pSum As Double
Dim curCol As Double, CurIdx As Integer

curCol = rCurCell.Column
StartCol = rPatroonNaam.Column
CurIdx = curCol - StartCol + 1

'inventariseer de namen
c = 0
For Each rCell In rPatroonNaam
  c = c + 1
  Naam(c) = VBA.UCase(rCell)
Next

'inventariseer de kansen
c = 0
For Each rCell In rPatroonKans
  c = c + 1
  Kans(c) = rCell
Next

'inventariseer het gebruik
c = 0
For Each rCell In rStochastInUse
  c = c + 1
  If rCell = "a" Then
    Inuse(c) = True
  Else
    Inuse(c) = False
  End If
Next

'bereken de som van kansen van de actieve patronen
pSum = 0
For c = 1 To 7
  If Inuse(c) Then pSum = pSum + Kans(c)
Next

'herverdeel de ongebruikte kansen naar rato over alle actieve patronen
For c = 1 To 7
  If Inuse(c) Then
    Kans(c) = Kans(c) / pSum
  Else
    Kans(c) = 0
  End If
Next


If Kans(CurIdx) = 0 Then
  BerekenStochastPatroonKlasse = ""
Else
  BerekenStochastPatroonKlasse = Kans(CurIdx)
End If


End Function

Public Function HERH2KLASSEFREQ(PrevH As Variant, curH As Variant, nextH As Variant, DurationHours As Integer) As Double
  'computes the frequency of a class given its return period AND the return period of the previous and next class
  'the result is based on the average return period between the surrounding classes
  Dim ExceedanceFrequencyLower As Double, ExceedanceFrequencyUpper As Double
  
  If Not IsNumeric(PrevH) Then PrevH = 0
  If Not IsNumeric(nextH) Then nextH = 0
  If Not IsNumeric(curH) Then curH = 0
  
  If curH = 0 Then
    'invalid return period!
    ExceedanceFrequencyLower = 0
    ExceedanceFrequencyUpper = 0
  ElseIf PrevH = 0 Then
    'this is the first class!
    ExceedanceFrequencyLower = 365.25 * 24 / DurationHours
    ExceedanceFrequencyUpper = 1 / ((curH + nextH) / 2)
  ElseIf nextH = 0 Then
    'this is the last class!
    ExceedanceFrequencyLower = 1 / ((PrevH + curH) / 2)
    ExceedanceFrequencyUpper = 0
  Else
    ExceedanceFrequencyLower = 1 / ((PrevH + curH) / 2)
    ExceedanceFrequencyUpper = 1 / ((curH + nextH) / 2)
  End If
  HERH2KLASSEFREQ = ExceedanceFrequencyLower - ExceedanceFrequencyUpper
  
End Function

Public Function HERHFROMSTOCHASTICRESULT(HERH As Double, WLEventNumRange As Range, WLValueRange As Range, FreqEventNumRange As Range, FreqValueRange As Range) As Double
  'this function computes the exceedance level for a given return period.
  'it expects two ranges with resp. event numbers and corresponding water levels,
  'and two ranges with resp. event numbers and corresponding frequencies
  Dim rWL As Long, rFreq As Long
  
  Dim WLValues() As Double, WLEventNums() As Integer, WLSortedIdx() As Long
  Dim FreqValues() As Double, FreqEventNums() As Integer
  Dim WLSorted() As Double, Herhalingstijd() As Double
  Dim FreqSum As Double, i As Long, j As Long
  Dim myEventNum As Integer, myWL As Double, myFreq As Double
  
  'input
  ReDim WLValues(1 To WLValueRange.Rows.Count)
  ReDim WLEventNums(1 To WLEventNumRange.Rows.Count)
  ReDim FreqValues(1 To FreqValueRange.Rows.Count)
  ReDim FreqEventNums(1 To FreqEventNumRange.Rows.Count)
  
  'output
  ReDim WLSorted(1 To WLValueRange.Rows.Count)
  ReDim Herhalingstijd(1 To WLValueRange.Rows.Count)
  
  If WLValueRange.Rows.Count <> WLEventNumRange.Rows.Count Then
    MsgBox ("Error: number of rows in water level range must be equal to that in the event number range.")
  ElseIf FreqValueRange.Rows.Count <> FreqEventNumRange.Rows.Count Then
    MsgBox ("Error: number of rows in frequency value range must be equal to that in the event number range.")
  Else
  
    'read the water levels
    For rWL = 1 To WLEventNumRange.Rows.Count
      WLValues(rWL) = WLValueRange.Cells(rWL, 1)
      WLEventNums(rWL) = WLEventNumRange.Cells(rWL, 1)
    Next
    
    'read the frequencies
    For rFreq = 1 To FreqEventNumRange.Rows.Count
      FreqValues(rFreq) = FreqValueRange.Cells(rFreq, 1)
      FreqEventNums(rFreq) = FreqEventNumRange.Cells(rFreq, 1)
    Next
    
    'create an array with the index number for the water levels in ascending order
    WLSortedIdx = HeapSort(WLValues)
    
    'walk through the water levels in descending order
    For i = UBound(WLSortedIdx) To 1 Step -1
      myEventNum = WLEventNums(WLSortedIdx(i))
      myWL = WLValues(WLSortedIdx(i))
      
      'find the frequency corresponding with this event
      For j = 1 To UBound(FreqEventNums)
        If FreqEventNums(j) = myEventNum Then
          myFreq = FreqValues(j)
          Exit For
        End If
      Next
      
      FreqSum = FreqSum + myFreq
      WLSorted(i) = myWL
      Herhalingstijd(i) = 1 / FreqSum
    Next
    
  End If
  
  'interpolate between the two surrounding Return Periods.
  For i = 1 To UBound(WLSorted) - 1
    If Herhalingstijd(i) <= HERH And Herhalingstijd(i + 1) >= HERH Then
      HERHFROMSTOCHASTICRESULT = Interpolate(Herhalingstijd(i), WLSorted(i), Herhalingstijd(i + 1), WLSorted(i + 1), HERH)
      Exit Function
    End If
  Next
  
End Function

Public Function KLASSEFREQUENTIEUITHERHALINGSTIJD(FrequentieSom As Double, HerhOndergrens As Variant, HerhBovengrens As Variant, VolgendeHerh As Variant) As Double
  If Not IsNumeric(HerhOndergrens) Or HerhOndergrens = "" Then
    'bereken klassefrequentie voor de onderste klasse
    KLASSEFREQUENTIEUITHERHALINGSTIJD = FrequentieSom - 1 / HerhBovengrens
  ElseIf Not IsNumeric(VolgendeHerh) Or VolgendeHerh = "" Then
    'er is geen volgende klasse, dus bereken hier het restant van de frequenties
    KLASSEFREQUENTIEUITHERHALINGSTIJD = 1 / HerhOndergrens
  Else
    KLASSEFREQUENTIEUITHERHALINGSTIJD = (1 / HerhOndergrens) - (1 / HerhBovengrens)
  End If
End Function

Public Function KLASSEKANSUITOVERSCHRIJDINGSKANSEN(Vorige As Variant, Huidige As Variant, Volgende As Variant) As Double
  If Not IsNumeric(Vorige) Or Vorige = "" Then
    KLASSEKANSUITOVERSCHRIJDINGSKANSEN = 1 - Huidige
  ElseIf Not IsNumeric(Volgende) Or Volgende = "" Then
    KLASSEKANSUITOVERSCHRIJDINGSKANSEN = Vorige
  Else
    KLASSEKANSUITOVERSCHRIJDINGSKANSEN = Vorige - Huidige
  End If
End Function

Public Sub CLASSIFYDURATIONS(ValuesRange As Range, threshold As Double, resultsRow As Integer, resultsCol As Integer)
  'deze routine onderzoekt welke duur (aantal tijdstappen) gebeurtenissen in een reeks hebben
  'argumenten: het bereik waarin de getallen staan en de drempelwaarde waarboven een gebeurtenis wordt 'gedetecteerd'
  
  Dim i As Long, j As Long, Values() As Double, Inuse() As Boolean, Durations() As Integer
  Dim n As Integer, maxn As Integer
  ReDim Values(1 To ValuesRange.Rows.Count)
  ReDim Inuse(1 To ValuesRange.Rows.Count)
  ReDim Durations(1 To ValuesRange.Rows.Count)
  For i = 1 To ValuesRange.Rows.Count
    Values(i) = ValuesRange.Cells(i, 1)
    Inuse(i) = False
  Next
  
  Dim Index() As Long
  Index = HeapSort(Values)
  
  For i = UBound(Index) To 1 Step -1
    If Values(Index(i)) > threshold And Inuse(Index(i)) = False Then
      n = 1
      Inuse(Index(i)) = True
      'move backwards to find the start of the event
      For j = Index(i) - 1 To 1 Step -1
        If Inuse(j) = True Then Exit For
        If Values(j) <= threshold Then Exit For
        n = n + 1
        Inuse(j) = True
      Next
      'move forwards to find the end of the event
      For j = Index(i) + 1 To ValuesRange.Rows.Count
        If Inuse(j) = True Then Exit For
        If Values(j) <= threshold Then Exit For
        n = n + 1
        Inuse(j) = True
      Next
      
      'keep track of the longest event found
      If n > maxn Then maxn = n
      
      'we've found an event and identified its duration. Store it in a histogram
      Durations(n) = Durations(n) + 1
    
    End If
  Next
  
  ReDim Preserve Durations(1 To maxn)
  
  'write the histogram to the results sheet
  For i = 1 To UBound(Durations)
    ActiveSheet.Cells(resultsRow + i - 1, resultsCol) = i
    ActiveSheet.Cells(resultsRow + i - 1, resultsCol + 1) = Durations(i)
  Next
    
  
End Sub

Private Sub Heapify(Keys, Index() As Long, ByVal i1 As Long, ByVal n As Long)
   ' Heap order rule: a[i] >= a[2*i+1] and a[i] >= a[2*i+2]
   Dim Base As Long: Base = LBound(Index)
   Dim nDiv2 As Long: nDiv2 = n \ 2
   Dim i As Long: i = i1
   Do While i < nDiv2
      Dim K As Long: K = 2 * i + 1
      If K + 1 < n Then
         If Keys(Index(Base + K)) < Keys(Index(Base + K + 1)) Then K = K + 1
         End If
      If Keys(Index(Base + i)) >= Keys(Index(Base + K)) Then Exit Do
      Exchange Index, i, K
      i = K
      Loop
   End Sub

Private Sub Exchange(A() As Long, ByVal i As Long, ByVal j As Long)
   Dim Base As Long: Base = LBound(A)
   Dim Temp As Long: Temp = A(Base + i)
   A(Base + i) = A(Base + j)
   A(Base + j) = Temp
   End Sub

Private Function GenerateArrayWithRandomValues()
   Dim n As Long: n = 1 + Rnd * 100
   ReDim A(0 To n - 1) As Long
   Dim i As Long
   For i = LBound(A) To UBound(A)
      A(i) = Rnd * 1000
      Next
   GenerateArrayWithRandomValues = A
   End Function

Private Sub VerifyIndexIsSorted(Keys, Index)
   Dim i As Long
   For i = LBound(Index) To UBound(Index) - 1
      If Keys(Index(i)) > Keys(Index(i + 1)) Then
         err.Raise vbObjectError, , "Index array is not sorted!"
         End If
      Next
   End Sub


Public Function OPPERVLAKAFGEPLATTECIRKEL(r As Double, Y_center As Double, Y_snede As Double) As Double
  'R = straal, Y_center = hoogte VBA.Middelpunt cirkel, Y_snede = hoogte waar de cirkel is afgesneden
  Dim O_cirkel As Double, O_taartpunt As Double, O_driehoek As Double
  Dim Hoogte As Double, Breedte As Double, Hoek As Double, pi As Double
  
  pi = 3.141592
  Hoogte = Y_snede - Y_center
  O_cirkel = pi * r ^ 2
  
  If Hoogte >= r Then
    'volledig gevulde cirkel
    OPPERVLAKAFGEPLATTECIRKEL = O_cirkel
  ElseIf Hoogte <= -1 * r Then 'lege cirkel
    OPPERVLAKAFGEPLATTECIRKEL = 0
  Else
    'de taartpunt die eruit wordt geknipt
    Breedte = Sqr(r ^ 2 - Hoogte ^ 2) 'pythagoras
    Hoek = 2 * ArcCos(Hoogte / r)
    O_taartpunt = Hoek / (2 * pi) * O_cirkel
    
    'de driehoek die weer moet worden toegevoegd
    O_driehoek = 2 * Hoogte * Breedte / 2
    
    OPPERVLAKAFGEPLATTECIRKEL = O_cirkel - O_taartpunt + O_driehoek
  End If
  
End Function

Public Function RotatePoint(ByVal Xold As Double, ByVal Yold As Double, ByVal XOrigin As Double, ByVal YOrigin As Double, ByVal degrees As Double, ByRef Xnew As Double, ByRef Ynew As Double) As Boolean
 Dim r As Double, theta As Double, dY As Double, dX As Double, Direction As Double
 'roteert een punt ten opzichte van zijn oorsprong
  
 dY = (Yold - YOrigin)
 dX = (Xold - XOrigin)
 r = Sqr(dX ^ 2 + dY ^ 2)
 
 If dX = 0 Then dX = 0.00000000000001
 theta = Math.Atn(dY / dX)
    
 Xnew = r * Math.Cos(theta - DEG2RAD(degrees)) + XOrigin
 Ynew = r * Math.Sin(theta - DEG2RAD(degrees)) + YOrigin
 RotatePoint = True
End Function
  
Public Function DEG2RAD(ByVal angle As Double) As Double
  'graden naar radialen
  DEG2RAD = angle / 180 * pi
End Function

Public Function RAD2DEG(ByVal angle As Double) As Double
  'radialen naar graden
  RAD2DEG = angle * 180 / pi
End Function

Public Function LINEANGLEDEGREES(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double) As Double
  'berekent de hoek van een lijn tussen twee xy co-ordinaten
  Dim dX As Double, dY As Double
  
  dX = VBA.Abs(X2 - X1)
  dY = VBA.Abs(Y2 - Y1)
  
  If dX = 0 Then
    If dY = 0 Then
      LINEANGLEDEGREES = 0
    ElseIf Y2 > Y1 Then
      LINEANGLEDEGREES = 0
    ElseIf Y2 < Y1 Then
      LINEANGLEDEGREES = 180
    End If
  ElseIf dY = 0 Then
    If X2 > X1 Then
      LINEANGLEDEGREES = 90
    ElseIf X2 < X1 Then
      LINEANGLEDEGREES = 270
    End If
  Else
    If X2 > X1 And Y2 > Y1 Then 'eerste kwadrant
      LINEANGLEDEGREES = R2D(VBA.Atn(dX / dY))
    ElseIf X2 > X1 And Y2 < Y1 Then 'tweede kwadrant
      LINEANGLEDEGREES = 90 + R2D(VBA.Atn(dY / dX))
    ElseIf X2 < X1 And Y2 < Y1 Then 'derde kwadrant
      LINEANGLEDEGREES = 180 + R2D(VBA.Atn(dX / dY))
    Else 'vierde kwadrant
      LINEANGLEDEGREES = 270 + R2D(VBA.Atn(dX / dY))
    End If
  End If
  
End Function

Public Function PointDistance(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
  PointDistance = VBA.Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
End Function

Public Function PointInPolygon(ByVal x As Double, ByVal y As Double, VerticesX As Collection, VerticesY As Collection) As Boolean
Dim pt As Integer
Dim total_angle As Double

  'Add up the angles between the point in question and adjacent points on the polygon taken in order.
  'If the total of all the angles is 2 * PI or -2 * PI, then the point is inside the polygon.
  'If the total is zero, the point is outside. You can verify this intuitively with some simple examples using squares or triangles.

    ' Get the angle between the point and the
    ' first and last vertices.
    total_angle = GetAngle(VerticesX(VerticesX.Count), VerticesY(VerticesY.Count), x, y, VerticesX(1), VerticesY(1))
    
    ' Add the angles from the point to each other pair of vertices.
    For pt = 1 To VerticesX.Count - 1
      total_angle = total_angle + GetAngle(VerticesX(pt), VerticesY(pt), x, y, VerticesX(pt + 1), VerticesY(pt + 1))
    Next pt

    ' The total angle should be 2 * PI or -2 * PI if
    ' the point is in the polygon and close to zero
    ' if the point is outside the polygon.
    PointInPolygon = (Abs(total_angle) > pi)
End Function

Public Function NearestPoint(ByVal x As Double, ByVal y As Double, ByVal myRange As Range, ByVal XCol As Integer, ByVal YCol As Integer, ByVal ReturnCol As Integer, HasHeader As Boolean)

  Dim r As Long, minDist As Double, myDist As Double
  Dim StartRow As Integer
  Dim myX As Double, myY As Double, minID As String
  minDist = 99999999
  
  If HasHeader Then
    StartRow = 2
  Else
    StartRow = 1
  End If
  
  For r = StartRow To myRange.Rows.Count
    myX = myRange.Cells(r, XCol)
    myY = myRange.Cells(r, YCol)
    myDist = Math.Sqr((myX - x) ^ 2 + (myY - y) ^ 2)
    If myDist < minDist Then
      minDist = myDist
      minID = myRange.Cells(r, ReturnCol)
    End If
  Next

  NearestPoint = minID

End Function

Public Function PoolCoordinaatX(ByVal alpha As Double, Length As Double) As Double
  'geeft de x-coordinaat terug, gegeven poolcoordinaat (alpha, lengte).
  'Let op: de hoek alhpa is gedefinieerd vanaf de vertikale as, NIET vanaf de horizontale!
  Dim rad As Double
  rad = D2R(alpha)
  PoolCoordinaatX = Sin(rad) * Length
End Function

Public Function PoolCoordinaatY(ByVal alpha As Double, Length As Double) As Double
  'geeft de y-coordinaat terug, gegeven poolcoordinaat (alpha, lengte).
  'Let op: de hoek alhpa is gedefinieerd vanaf de vertikale as, NIET vanaf de horizontale!
  Dim rad As Double
  rad = D2R(alpha)
  PoolCoordinaatY = Cos(rad) * Length
End Function

Public Function PYTHAGORAS(ByVal A As Double, b As Double) As Double
  PYTHAGORAS = Math.Sqr(A ^ 2 + b ^ 2)
End Function

Public Function PYTHAGORAS_INVERSE(ByVal A As Double, c As Double) As Double
  'c = schuine zijde, a = rechte zijde
  'a^2 + b^2 = c ^2
  'b^2 = c^2 - a ^2
  'b = sqr(c^2 - a^2)
  PYTHAGORAS_INVERSE = Math.Sqr(c ^ 2 - A ^ 2)
End Function


Public Function MileageOneUp(startNum As Integer, endNum As Integer, ByRef myArray() As Integer) As Boolean
  'werkt als een kilometerteller. Als het hectometergetal boven z'n maximum komt, springt hij terug naar nul
  'en gaat het getalletje ervoor een omhoog et cetera. Produceert TRUE bij succes
  'produceert FALSE als hij aan z'n eind is gekomen en niet verder kan ophogen
  Dim nElements As Integer
  nElements = UBound(myArray)
  Dim i As Long, j As Long
  Dim Done As Boolean
  Done = True
  
  '---------------------------------------------------------------------
  'for the very first run, the following is crucial:
  'if necessary initialize the array to be equal to its start number
  'EXCEPT for the last number. That should start one lower than startnum
  '---------------------------------------------------------------------
  For i = 1 To nElements - 1
    If myArray(i) < startNum Then myArray(i) = startNum
  Next
  If myArray(nElements) < startNum - 1 Then myArray(nElements) = startNum - 1
  '---------------------------------------------------------------------

  'also check if the counter is already at its maxumum. if so, quit
  For i = 1 To nElements
    If myArray(i) < endNum Then Done = False
    Exit For
  Next
  If Done Then
    MileageOneUp = False
    Exit Function
  End If

  'there is still some room for moving further
  For i = nElements To 1 Step -1
    If myArray(i) < endNum Then
      myArray(i) = myArray(i) + 1
      MileageOneUp = True
      Exit Function
    ElseIf myArray(i) = endNum And i = 1 Then
      MileageOneUp = False
      Exit Function
    ElseIf myArray(i) = endNum And i > 1 Then
      myArray(i) = startNum 'reset de waarde naar de basisstand en draai de voorgaande een omhoog
      For j = i - 1 To 1 Step -1
        If myArray(j) < endNum Then
          myArray(j) = myArray(j) + 1
          MileageOneUp = True
          Exit Function
        Else
          myArray(j) = startNum 'reset ook deze voorgaande waarde naar de basisstand en ga weer door naar degene hiervoor
        End If
      Next
    End If
  Next

End Function

Public Function MeetsCondition(ByVal myVal As Double, ByVal Condition As String) As Boolean
  Dim Operator As String, Operand As Double
  
  'tests a value to a certain conditions
  Condition = VBA.Trim(Condition)
  
  'if no condition specified, exit straight away. Always true
  If Condition = "" Then
    MeetsCondition = True
    Exit Function
  End If
  
  'check validity of the condition string
  If InStr(1, Condition, " ") <= 0 Then
    MsgBox ("Error: condition must contain a space between operator and operand: " & Condition)
    End
  End If
  
  'parse the string to retrieve operator and operand
  Operator = ParseString(Condition)
  Operand = Condition
  
  'perform the check
  Select Case Operator
    Case Is = "<"
      If myVal < Operand Then MeetsCondition = True
    Case Is = "<="
      If myVal <= Operand Then MeetsCondition = True
    Case Is = ">"
      If myVal > Operand Then MeetsCondition = True
    Case Is = ">="
      If myVal >= Operand Then MeetsCondition = True
    Case Is = "<>"
      If myVal <> Operand Then MeetsCondition = True
    Case Is = "="
      If myVal = Operand Then MeetsCondition = True
    Case Else
      MsgBox ("Error: operand not (yet) supported in condition: " & Operand & " " & Operator)
      End
  End Select

End Function

Public Function RMSE(Range1 As Range, Range2 As Range) As Double
    Dim i As Long, myRMSE As Double
    If Range1.Rows.Count <> Range2.Rows.Count Then
        MsgBox ("Ranges should have equal length")
    Else
        For i = 1 To Range1.Rows.Count
            myRMSE = myRMSE + (Range1.Cells(i, 0) - Range2.Cells(i, 0)) ^ 2
        Next
        RMSE = Math.Sqr(myRMSE / Range1.Rows.Count)
    End If
End Function

Public Function GetShapeByNameFromWorksheet(ByRef MySheet As Worksheet, MyName As String) As Shape
  'finds the shape with a given name on the active worksheet
  Dim myShape As Shape
  For Each myShape In MySheet.Shapes
    If myShape.Name = MyName Then
      Set GetShapeByNameFromWorksheet = myShape
    End If
  Next
End Function


Public Function TrapeziumRangeToYZProfiles(Source As Range, IDCol As Integer, XCol As Integer, YCol As Integer, ProfileHeight As Double, BedLevelCol As Integer, BedWidthCol As Integer, LeftSlopeCol As Integer, RightSlopeCol As Integer, WetBermLeftBob1Col As Integer, WetBermLeftBob2Col As Integer, WetBermLeftWidthCol As Integer, WetBermLeftSlopeCol As Integer, WetBermRightBob1Col As Integer, WetBermRightBob2Col As Integer, WetBermRightWidthCol As Integer, WetBermRightSlopeCol As Integer, TargetSheet As String) As Boolean
    'this function creates YZ-cross section tables from a given range with trapezium information
    'in this function we assume that nothing is known of the surface level or surface width
    Dim rs As Integer, cs As Integer 'row and column source
    Dim rt As Integer, ct As Integer 'row and column target
    Dim i As Integer, nErrors As Integer
    Dim BedLevel As Double, BedWidth As Double
    Dim SlopeLeft As Double, SlopeRight As Double
    Dim SurfaceLevel As Double
    Dim WetBermLeftBob1 As Double, WetBermLeftBob2 As Double, WetBermLeftWidth As Double
    Dim WetBermRightBob1 As Double, WetBermRightBob2 As Double, WetBermRightWidth As Double
    Dim WetBermLeftSlope As Double, WetBermRightSlope As Double
    Dim YVals As Collection
    Dim ZVals As Collection
    Dim XCOORD As Double, YCOORD As Double, ID As String
    
    'initialize the target sheet
    rt = rt + 1
    Worksheets(TargetSheet).Cells(rt, 1) = "ID"
    Worksheets(TargetSheet).Cells(rt, 2) = "XCOORD"
    Worksheets(TargetSheet).Cells(rt, 3) = "YCOORD"
    Worksheets(TargetSheet).Cells(rt, 4) = "Y"
    Worksheets(TargetSheet).Cells(rt, 5) = "X"
           
    'process each cross section
    For rs = 2 To Source.Rows.Count 'skip the header
        nErrors = 0
        Set YVals = New Collection
        Set ZVals = New Collection
        If Source.Cells(rs, IDCol) = "" Then nErrors = nErrors + 1 Else ID = Source.Cells(rs, IDCol)
        If Source.Cells(rs, XCol) = "" Then nErrors = nErrors + 1 Else XCOORD = Source.Cells(rs, XCol)
        If Source.Cells(rs, YCol) = "" Then nErrors = nErrors + 1 Else YCOORD = Source.Cells(rs, YCol)
        If Source.Cells(rs, BedLevelCol) = "" Then nErrors = nErrors + 1 Else BedLevel = Source.Cells(rs, BedLevelCol)
        If Source.Cells(rs, BedWidthCol) = "" Then nErrors = nErrors + 1 Else BedWidth = Source.Cells(rs, BedWidthCol)
                
        If nErrors = 0 Then
            If Source.Cells(rs, LeftSlopeCol) <> "" Then SlopeLeft = Source.Cells(rs, LeftSlopeCol) Else SlopeLeft = Source.Cells(rs, RightSlopeCol)
            If Source.Cells(rs, RightSlopeCol) <> "" Then SlopeRight = Source.Cells(rs, RightSlopeCol) Else SlopeRight = SlopeLeft
            
            'add the leftmost coordinate
            SurfaceLevel = BedLevel + ProfileHeight
            Call YVals.Add(0)
            Call ZVals.Add(SurfaceLevel)
            
            'in case we have a left wet berm, add it
            If Source.Cells(rs, WetBermLeftBob1Col) <> "" And Source.Cells(rs, WetBermLeftBob2Col) <> "" And Source.Cells(rs, WetBermLeftWidthCol) > 0 Then
                'calculate the distance of the start of the wet berm
                WetBermLeftBob1 = Source.Cells(rs, WetBermLeftBob1Col)
                WetBermLeftBob2 = Source.Cells(rs, WetBermLeftBob2Col)
                WetBermLeftWidth = Source.Cells(rs, WetBermLeftWidthCol)
                If Source.Cells(rs, WetBermLeftSlopeCol) <> "" Then WetBermLeftSlope = Source.Cells(rs, WetBermLeftSlopeCol) Else WetBermLeftSlope = SlopeLeft
                            
                Call YVals.Add(WetBermLeftSlope * (SurfaceLevel - WetBermLeftBob2))
                Call ZVals.Add(WetBermLeftBob2)
                Call YVals.Add(YVals.Item(YVals.Count) + WetBermLeftWidth)
                Call ZVals.Add(WetBermLeftBob1)
                Call YVals.Add(YVals.Item(YVals.Count) + SlopeLeft * (WetBermLeftBob1 - BedLevel))
                Call ZVals.Add(BedLevel)
            Else
                Call YVals.Add(SlopeLeft * (SurfaceLevel - BedLevel))
                Call ZVals.Add(BedLevel)
            End If
            
            'bed width
            Call YVals.Add(YVals.Item(YVals.Count) + BedWidth)
            Call ZVals.Add(BedLevel)
            
            'in case we have a right wet berm, add it
            If Source.Cells(rs, WetBermRightBob1Col) <> "" And Source.Cells(rs, WetBermRightBob2Col) <> "" And Source.Cells(rs, WetBermRightWidthCol) > 0 Then
                'calculate the distance of the start of the wet berm
                WetBermRightBob1 = Source.Cells(rs, WetBermRightBob1Col)
                WetBermRightBob2 = Source.Cells(rs, WetBermRightBob2Col)
                WetBermRightWidth = Source.Cells(rs, WetBermRightWidthCol)
                If Source.Cells(rs, WetBermRightSlopeCol) <> "" Then WetBermRightSlope = Source.Cells(rs, WetBermRightSlopeCol) Else WetBermRightSlope = SlopeRight
                
                Call YVals.Add(YVals.Item(YVals.Count) + (WetBermRightBob1 - BedLevel) * SlopeRight)
                Call ZVals.Add(WetBermRightBob1)
                Call YVals.Add(YVals.Item(YVals.Count) + WetBermRightWidth)
                Call ZVals.Add(WetBermRightBob2)
                Call YVals.Add(YVals.Item(YVals.Count) + (SurfaceLevel - WetBermRightBob2) * WetBermRightSlope)
                Call ZVals.Add(SurfaceLevel)
            Else
                Call YVals.Add(YVals.Item(YVals.Count) + SlopeRight * (SurfaceLevel - BedLevel))
                Call ZVals.Add(SurfaceLevel)
            End If
            
            For i = 1 To YVals.Count
                rt = rt + 1
                Worksheets(TargetSheet).Cells(rt, 1) = ID
                Worksheets(TargetSheet).Cells(rt, 2) = XCOORD
                Worksheets(TargetSheet).Cells(rt, 3) = YCOORD
                Worksheets(TargetSheet).Cells(rt, 4) = YVals(i)
                Worksheets(TargetSheet).Cells(rt, 5) = ZVals(i)
            Next
        End If
    Next


End Function


Public Function GetAngle(ByVal Ax As Single, ByVal Ay As Single, ByVal Bx As Single, ByVal By As Single, ByVal Cx As Single, ByVal Cy As Single) As Single
' Return the angle ABC.
' Returns a value between PI and -PI.
' Note that the value is the opposite of what you might expect because Y coordinates increase downward.
    Dim dot_product As Single
    Dim cross_product As Single

    ' Get the dot product and cross product.
    dot_product = DotProduct(Ax, Ay, Bx, By, Cx, Cy)
    cross_product = CrossProductLength(Ax, Ay, Bx, By, Cx, Cy)

    ' Calculate the angle.
    GetAngle = ATan2(cross_product, dot_product)
End Function

Public Function ATan2(ByVal Opp As Single, ByVal adj As Single) As Single
  Dim angle As Single
  ' Return the angle with tangent opp/hyp. The returned
  ' value is between PI and -PI.

  ' Get the basic angle.
  If Abs(adj) < 0.0001 Then
    angle = pi / 2
  Else
    angle = Abs(Atn(Opp / adj))
  End If

  ' See if we are in quadrant 2 or 3.
  If adj < 0 Then
    'angle > PI/2 or angle < -PI/2.
    angle = pi - angle
  End If

  'See if we are in quadrant 3 or 4.
  If Opp < 0 Then
    angle = -angle
  End If

  'Return the result.
  ATan2 = angle

End Function


Private Function DotProduct(ByVal Ax As Single, ByVal Ay As Single, ByVal Bx As Single, ByVal By As Single, ByVal Cx As Single, ByVal Cy As Single) As Single
  ' Return the dot product AB · BC.
  ' Note that AB · BC = |AB| * |BC| * Cos(theta).
  Dim BAx As Single
  Dim BAy As Single
  Dim BCx As Single
  Dim BCy As Single
    
  ' Get the vectors' coordinates.
  BAx = Ax - Bx
  BAy = Ay - By
  BCx = Cx - Bx
  BCy = Cy - By
    
  ' Calculate the dot product.
  DotProduct = BAx * BCx + BAy * BCy

End Function

Public Function CrossProductLength( _
    ByVal Ax As Single, ByVal Ay As Single, _
    ByVal Bx As Single, ByVal By As Single, _
    ByVal Cx As Single, ByVal Cy As Single _
  ) As Single

  ' Return the cross product AB x BC.
  ' The cross product is a vector perpendicular to AB
  ' and BC having length |AB| * |BC| * Sin(theta) and
  ' with direction given by the VBA.Right-hand rule.
  ' For two vectors in the X-Y plane, the result is a
  ' vector with X and Y components 0 so the Z component
  ' gives the vector's length and direction.

  Dim BAx As Single
  Dim BAy As Single
  Dim BCx As Single
  Dim BCy As Single

  ' Get the vectors' coordinates.
  BAx = Ax - Bx
  BAy = Ay - By
  BCx = Cx - Bx
  BCy = Cy - By

  ' Calculate the Z coordinate of the cross product.
  CrossProductLength = BAx * BCy - BAy * BCx

End Function

Public Function NATTEOMTREKAFGEPLATTECIRKEL(r As Double, Y_center As Double, Y_snede As Double) As Double
  Dim Hoogte As Double, Breedte As Double, Hoek As Double
  Dim Omtrek_cirkel As Double
  
  Omtrek_cirkel = 2 * pi * r
  Hoogte = Y_snede - Y_center
  
  If Hoogte >= r Then        'volledige cirkel
    NATTEOMTREKAFGEPLATTECIRKEL = 2 * pi * r
  ElseIf Hoogte <= -1 * r Then   'lege cirkel
    NATTEOMTREKAFGEPLATTECIRKEL = 0
  Else                                  'de hoek van de taartpunt die eruit wordt geknipt (radialen)
    Breedte = Sqr(r ^ 2 - Hoogte ^ 2) 'pythagoras
    Hoek = 2 * ArcCos(Hoogte / r)
    NATTEOMTREKAFGEPLATTECIRKEL = (2 * pi - Hoek) * r
  End If
  
End Function

Public Function EllipsBreedte(Breedte As Double, Hoogte As Double, h As Double) As Double
  'h is gedefinieerd als de hoogte vanaf de bodem van de ellips
  'een ellips voldoet aan de vgl x^2/a^2 + y^2/b^2 = 1
  'waarbij het brandpunt van de ellips als nulpunt moet worden beschouwd, a de halve breedte is en b de halve hoogte
  Dim A As Double
  Dim b As Double
  Dim y As Double 'hoogte y tov brandpunt
  Dim x As Double
  
  b = Hoogte / 2
  A = Breedte / 2
  
  y = h - b
  
  If h >= 0 And h <= Hoogte Then
    x = Sqr((1 - y ^ 2 / b ^ 2) * A ^ 2)
    EllipsBreedte = x * 2
  Else
    EllipsBreedte = -999
  End If

End Function

' Inverse Sinus
Function ArcSin(x As Double) As Double
  ArcSin = Atn(x / Sqr(-x * x + 1))
End Function

'Inverse Cosinus
Function ArcCos(x As Double) As Double
  ArcCos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Function

'inverse tangent
Function ArcTan(x As Double) As Double
  ArcTan = Atn(x)
End Function

Public Function ArcTan2(ByVal x As Double, ByVal y As Double) As Double
  
  'Code from www.visiblevisual.com
  If x = 0 And y = 0 Then
    ATan2 = 0
  Else
    If x = 0 Then x = 0.00000000001
    ATan2 = Atn(y / x) - pi * (x < 0)
  End If
  End Function

End Function

Public Function DaysInMonth(myDate)

  Dim NextMonth, EndOfMonth
  NextMonth = DateAdd("m", 1, myDate)
  EndOfMonth = NextMonth - DatePart("d", NextMonth)
  DaysInMonth = DatePart("d", EndOfMonth)

End Function

Public Function DaysInMonth2(myMonth As Integer, myYear As Integer, Optional AlwaysInclude29Feb As Boolean = False)

  If myMonth = 1 Then
    DaysInMonth2 = 31
  ElseIf myMonth = 2 Then
    If AlwaysInclude29Feb Then
      DaysInMonth2 = 29
    ElseIf IsLeapYear(myYear) Then
      DaysInMonth2 = 29
    Else
      DaysInMonth2 = 28
    End If
  ElseIf myMonth = 3 Then
    DaysInMonth2 = 31
  ElseIf myMonth = 4 Then
    DaysInMonth2 = 30
  ElseIf myMonth = 5 Then
    DaysInMonth2 = 31
  ElseIf myMonth = 6 Then
    DaysInMonth2 = 30
  ElseIf myMonth = 7 Then
    DaysInMonth2 = 31
  ElseIf myMonth = 8 Then
    DaysInMonth2 = 31
  ElseIf myMonth = 9 Then
    DaysInMonth2 = 30
  ElseIf myMonth = 10 Then
    DaysInMonth2 = 31
  ElseIf myMonth = 11 Then
    DaysInMonth2 = 30
  ElseIf myMonth = 12 Then
    DaysInMonth2 = 31
  End If
End Function

Public Function IsLeapYear(myYear As Integer) As Boolean
  If VBA.Round(myYear / 4, 0) = myYear / 4 Then
    IsLeapYear = True
  Else
    IsLeapYear = False
  End If
End Function

Public Function Kwartaal(myDate)
  Select Case Month(myDate)
  Case Is = 1
    Kwartaal = 1
  Case Is = 2
    Kwartaal = 1
  Case Is = 3
    Kwartaal = 1
  Case Is = 4
    Kwartaal = 2
  Case Is = 5
    Kwartaal = 2
  Case Is = 6
    Kwartaal = 2
  Case Is = 7
    Kwartaal = 3
  Case Is = 8
    Kwartaal = 3
  Case Is = 9
    Kwartaal = 3
  Case Is = 10
    Kwartaal = 4
  Case Is = 11
    Kwartaal = 4
  Case Is = 12
    Kwartaal = 4
  End Select
End Function

Public Function Halfjaar(myDate As Date) As String
  Select Case Month(myDate)
    
  Case Is = 1
    Halfjaar = Year(myDate) - 1 & "-" & VBA.Right(Year(myDate), 2) & " winter"
  Case Is = 2
    Halfjaar = Year(myDate) - 1 & "-" & VBA.Right(Year(myDate), 2) & " winter"
  Case Is = 3
    Halfjaar = Year(myDate) - 1 & "-" & VBA.Right(Year(myDate), 2) & " winter"
  Case Is = 4
    Halfjaar = Year(myDate) & " zomer"
  Case Is = 5
    Halfjaar = Year(myDate) & " zomer"
  Case Is = 6
    Halfjaar = Year(myDate) & " zomer"
  Case Is = 7
    Halfjaar = Year(myDate) & " zomer"
  Case Is = 8
    Halfjaar = Year(myDate) & " zomer"
  Case Is = 9
    Halfjaar = Year(myDate) & " zomer"
  Case Is = 10
    Halfjaar = Year(myDate) & "-" & VBA.Right(Year(myDate) + 1, 2) & " winter"
  Case Is = 11
    Halfjaar = Year(myDate) & "-" & VBA.Right(Year(myDate) + 1, 2) & " winter"
  Case Is = 12
    Halfjaar = Year(myDate) & "-" & VBA.Right(Year(myDate) + 1, 2) & " winter"
  End Select
  
End Function

Public Function METEOROLOGISCHSEIZOEN(myDate As Date) As String
  'geeft het meteorologische seizoen van een datum terug
  If Month(myDate) <= 2 Or Month(myDate) = 12 Then
    METEOROLOGISCHSEIZOEN = "winter"
  ElseIf Month(myDate) < 6 Then
    METEOROLOGISCHSEIZOEN = "lente"
  ElseIf Month(myDate) < 9 Then
    METEOROLOGISCHSEIZOEN = "zomer"
  ElseIf Month(myDate) < 12 Then
    METEOROLOGISCHSEIZOEN = "herfst"
  End If
End Function

Public Function METEOROLOGISCHHALFJAAR(myDate As Date) As String
  'geeft het meteorologische halfjaar van een datum terug
  If Month(myDate) <= 3 Then
    METEOROLOGISCHHALFJAAR = "winter"
  ElseIf Month(myDate) <= 9 Then
    METEOROLOGISCHHALFJAAR = "zomer"
  Else
    METEOROLOGISCHHALFJAAR = "winter"
  End If
End Function

Public Function HYDROLOGISCHSEIZOEN(myDate As Date, WinZomMonth As Long, WinZomDay As Long, ZomWinMonth As Long, ZomWinDay As Long) As String
  'geeft het hydrologisch seizoen van een datum terug
  If Month(myDate) < WinZomMonth Then
    HYDROLOGISCHSEIZOEN = "winter"
  ElseIf Month(myDate) > ZomWinMonth Then
    HYDROLOGISCHSEIZOEN = "winter"
  ElseIf Month(myDate) > WinZomMonth And Month(myDate) < ZomWinMonth Then
    HYDROLOGISCHSEIZOEN = "zomer"
  ElseIf Month(myDate) = WinZomMonth Then
    If Day(myDate) >= WinZomDay Then
      HYDROLOGISCHSEIZOEN = "zomer"
    Else
      HYDROLOGISCHSEIZOEN = "winter"
    End If
  ElseIf Month(myDate) = ZomWinMonth Then
    If Day(myDate) >= ZomWinDay Then
      HYDROLOGISCHSEIZOEN = "winter"
    Else
      HYDROLOGISCHSEIZOEN = "zomer"
    End If
  End If
  
End Function

Public Function DOUBLE2DATETIMESTRING(myDate As Double, Optional DateSeparator As String = "/", Optional TimeSeparator As String = ":", Optional DateTimeSeparator As String = "-", Optional YearLen As Long = 4, Optional YearOrder As Integer = 1, Optional MonthOrder As Integer = 2, Optional DayOrder As Integer = 3, Optional HourOrder As Integer = 4, Optional MinuteOrder As Integer = 5, Optional SecondOrder As Integer = 6) As String
Dim YearStr As String
Dim MonthStr As String
Dim DayStr As String
Dim HourStr As String
Dim MinuteStr As String
Dim SecondStr As String

If YearOrder + MonthOrder + DayOrder + HourOrder + MinuteOrder + SecondOrder <> 21 Then
  DOUBLE2DATETIMESTRING = "Error, invalid order specified for datetime-elements"
  Exit Function
Else
  If YearLen = 2 Then
    YearStr = VBA.Format(Year(myDate), "00")
  ElseIf YearLen = 4 Then
    YearStr = VBA.Format(Year(myDate), "0000")
  Else
    DOUBLE2DATETIMESTRING = "Error, year must be in 2 or 4 digits, e.g. 12 or 2012"
    Exit Function
  End If
  
  MonthStr = VBA.Format(Month(myDate), "00")
  DayStr = VBA.Format(Day(myDate), "00")
  HourStr = VBA.Format(Hour(myDate), "00")
  MinuteStr = VBA.Format(Minute(myDate), "00")
  SecondStr = VBA.Format(Second(myDate), "00")
  
  If YearOrder = 1 And MonthOrder = 2 And DayOrder = 3 And HourOrder = 4 And MinuteOrder = 5 And SecondOrder = 6 Then
    DOUBLE2DATETIMESTRING = YearStr & DateSeparator & MonthStr & DateSeparator & DayStr & DateTimeSeparator & HourStr & TimeSeparator & MinuteStr & TimeSeparator & SecondStr
    Exit Function
  ElseIf YearOrder = 3 And MonthOrder = 2 And DayOrder = 1 And HourOrder = 4 And MinuteOrder = 5 And SecondOrder = 6 Then
    DOUBLE2DATETIMESTRING = DayStr & DateSeparator & MonthStr & DateSeparator & YearStr & DateTimeSeparator & HourStr & TimeSeparator & MinuteStr & TimeSeparator & SecondStr
    Exit Function
  Else
    DOUBLE2DATETIMESTRING = "Error: specified order of date-time elements not (yet) supported."
    Exit Function
  End If
End If

End Function

Public Function DateExists(myYear As Long, myMonth As Long, myDay As Long) As Boolean

DateExists = True
If myDay < 1 Or myDay > 31 Then
  DateExists = False
ElseIf myMonth < 1 Or myMonth > 12 Then
  DateExists = False
ElseIf myMonth = 4 Or myMonth = 6 Or myMonth = 9 Or myMonth = 11 Then
  If myDay > 30 Then
    DateExists = False
  End If
ElseIf myMonth = 2 Then
  If myDay > 29 Then
    DateExists = False
  ElseIf myDay > 28 Then  'alleen geldig bij een schrikkeljaar
    If myYear / 4 <> Round(myYear / 4, 0) Then
      DateExists = False
    End If
  End If
End If

End Function

Public Function DayNumber(myDate As Date, AlwaysInclude29Feb As Boolean) As Integer
  Dim myMonth As Integer
  Dim myNum As Integer
  Dim i As Integer
  
  For i = 1 To 12
    If i = Month(myDate) Then
      myNum = myNum + Day(myDate)
      DayNumber = myNum
      Exit Function
    Else
      myNum = myNum + DaysInMonth2(i, Year(myDate), AlwaysInclude29Feb)
    End If
  Next
  myMonth = Month(myDate)
End Function

Public Function DATEHOURWINDOW(myDate As Double) As Double
  Dim myHour As Integer
  myHour = Hour(myDate)
  
  DATEHOUR = DateSerial(Year(myDate), Month(myDate), Day(myDate))
  DATEHOUR = DATEHOUR + myHour / 24
          
End Function

Public Function DATETWOHOURWINDOW(myDate As Variant) As Double
  'Author: Siebe Bosch
  'Description: returns the date + the two-hour-window of the day a certain datetime-value falls in
  Dim myHour As Integer
  myHour = Hour(myDate)
  DATETWOHOURWINDOW = DateSerial(Year(myDate), Month(myDate), Day(myDate))
  
  If myHour < 2 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 1 / 24
  ElseIf myHour < 4 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 3 / 24
  ElseIf myHour < 6 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 5 / 24
  ElseIf myHour < 8 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 7 / 24
  ElseIf myHour < 10 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 9 / 24
  ElseIf myHour < 12 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 11 / 24
  ElseIf myHour < 14 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 13 / 24
  ElseIf myHour < 16 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 15 / 24
  ElseIf myHour < 18 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 17 / 24
  ElseIf myHour < 20 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 19 / 24
  ElseIf myHour < 22 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 21 / 24
  Else
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 23 / 24
  End If
  
End Function

Public Function DATEFOURHOURWINDOW(myDate As Double) As Double
  'Author: Siebe Bosch
  'Description: returns the date + the quarter of the day a certain datetime-value falls in
  Dim myHour As Integer
  myHour = Hour(myDate)
  DATEFOURHOURWINDOW = DateSerial(Year(myDate), Month(myDate), Day(myDate))
  
  If myHour < 6 Then
    DATEFOURHOURWINDOW = DATEFOURHOURWINDOW + 3 / 24  '3 is the middle between 0 and 6
  ElseIf myHour < 12 Then
    DATEFOURHOURWINDOW = DATEFOURHOURWINDOW + 9 / 24  '9 is the middle between 6 and 12
  ElseIf myHour < 18 Then
    DATEFOURHOURWINDOW = DATEFOURHOURWINDOW + 12 / 24 '15 is the middle between 12 and 18
  Else
    DATEFOURHOURWINDOW = DATEFOURHOURWINDOW + 21 / 24 '21 is the middle between 18 and 24
  End If
  
End Function

Public Function DATEFROMSTRING(myDate As String, dateFormat As String) As Double
   Dim myYear As Integer, myMonth As Integer, myDay As Integer, myHour As Integer, myMinute As Integer, mySecond As Integer
   
   Select Case dateFormat
     Case Is = "yyyymmddhh"
       myYear = Left(myDate, 4)
       myMonth = Left(Right(myDate, 6), 2)
       myDay = Left(Right(myDate, 4), 2)
       myHour = Right(myDate, 2)
    Case Is = "yyyymmdd"
      myYear = Left(myDate, 4)
      myMonth = Left(Right(myDate, 4), 2)
      myDay = Right(myDate, 2)
   End Select
   
   'corrigeer wanneer 00 uur als 24 wordt weergegeven
   If myHour = 24 Then
     myHour = 0
     myDay = myDay + 1
     If myDay > DaysInMonth2(myMonth, myYear) Then
       myDay = 1
       myMonth = myMonth + 1
       If myMonth > 12 Then
         myMonth = 1
         myYear = myYear + 1
       End If
     End If
   End If
      
   DATEFROMSTRING = DateValue(myYear & "-" & myMonth & "-" & myDay) + TimeValue(myHour & ":" & myMinute & ":" & mySecond)
End Function

Public Function DATE2TEXT(myDate As Date, Formatting As String, MidnightAs24 As Boolean) As String
    Dim myStr As String, hpos As Integer
    
    If MidnightAs24 Then
        If Hour(myDate) = 0 And Minute(myDate) = 0 And Second(myDate) = 0 Then
            myStr = Format(myDate - 1, Formatting)
            hpos = InStr(1, Formatting, "hh", vbTextCompare)
            myStr = Left(myStr, hpos - 1) & "24" & Right(myStr, hpos + 2)
        End If
            myStr = Format(myDate, Formatting)
        Else
    End If
    DATE2TEXT = myStr
    
End Function


Public Function TIMEFROMSTRING(myDate As String, timeFormat As String) As Double
  Dim myHour As Integer, myMinute As Integer, mySecond As Integer
   
  Select Case Trim(LCase(timeFormat))
    Case Is = "hm"
      If VBA.Len(myDate) = 2 Then
        myHour = 0
        myMinute = myDate
      ElseIf VBA.Len(myDate) = 3 Then
        myHour = Left(myDate, 1)
        myMinute = Right(myDate, 2)
      ElseIf VBA.Len(myDate) = 4 Then
        myHour = Left(myDate, 2)
        myMinute = Right(myDate, 2)
      End If
    Case Is = "hhmm"
      myHour = VBA.Left(myDate, 2)
      myMinute = VBA.Right(myDate, 2)
    Case Is = "hhmmss"
      myHour = VBA.Left(myDate, 2)
      myMinute = VBA.Mid(myDate, 3, 2)
      mySecond = VBA.Right(myDate, 2)
   End Select
      
   TIMEFROMSTRING = TimeValue(myHour & ":" & myMinute & ":" & mySecond)
End Function


Public Function DATEANDTIMEFROMSTRINGS(myDateStr As String, myTimeStr As String, dateFormat As String, timeFormat As String) As Double
  Dim myDay As Integer, myMonth As Integer, myYear As Integer
  Dim myHour As Integer, myMinute As Integer, mySecond As Integer
  Dim myDate As Date
  
  myDate = DATEFROMSTRING(myDateStr, dateFormat)
   
  'timeformat doen we handmatig ivm het mogelijk voorkomen van 24:00
  Select Case timeFormat
    Case Is = "hhmm"
      If VBA.Len(myTimeStr) = 2 Then
        myHour = 0
        myMinute = myTimeStr
      ElseIf VBA.Len(myTimeStr) = 3 Then
        myHour = Left(myTimeStr, 1)
        myMinute = Right(myTimeStr, 2)
      ElseIf VBA.Len(myTimeStr) = 4 Then
        myHour = Left(myTimeStr, 2)
        myMinute = Right(myTimeStr, 2)
      End If
   End Select
      
   If myHour = 24 Then
     myDate = myDate + 1
     myHour = 0
   End If
      
   DATEANDTIMEFROMSTRINGS = myDate + TimeValue(myHour & ":" & myMinute & ":" & mySecond)
   
   
End Function

Public Function Decade(myDate As Date) As Integer
    'returns the decade number for a given date.
    'note: 36 is the highest possible value. All values will be limited to 36
    Decade = Maximum(1, Minimum(WorksheetFunction.RoundUp((myDate - DateSerial(Year(myDate), 1, 1)) / 10, 0), 36))
End Function

Public Function SOBEKTIMETABLESTRING(WinZomMonthStart As Integer, WinZomDayStart As Integer, WinZomMonthEnd As Integer, WinZomDayEnd As Integer, ZomWinMonthStart As Integer, ZomWinDayStart As Integer, ZomWinMonthEnd As Integer, ZomWinDayEnd As Integer, WinVal As Double, ZomVal As Double) As String
    Dim myStr As String
    Dim i As Integer
        myStr = "'1906/01/01;00:00:00' " & WinVal & " <"
    For i = 1906 To 2050
        myStr = myStr & vbCrLf & "'" & i & "/" & Format(WinZomMonthStart, "00") & "/" & Format(WinZomDayStart, "00") & ";00:00:00' " & WinVal & " <"
        myStr = myStr & vbCrLf & "'" & i & "/" & Format(WinZomMonthEnd, "00") & "/" & Format(WinZomDayEnd, "00") & ";00:00:00' " & ZomVal & " <"
        myStr = myStr & vbCrLf & "'" & i & "/" & Format(ZomWinMonthStart, "00") & "/" & Format(ZomWinDayStart, "00") & ";00:00:00' " & ZomVal & " <"
        myStr = myStr & vbCrLf & "'" & i & "/" & Format(ZomWinMonthEnd, "00") & "/" & Format(ZomWinDayEnd, "00") & ";00:00:00' " & WinVal & " <"
    Next
    SOBEKTIMETABLESTRING = myStr
End Function

Public Function DecadeStartDate(Year As Integer, Decade As Integer) As Date
    DecadeStartDate = DateSerial(Year, 1, 1) - 10 + Decade * 10
End Function

Public Function PAIR_LOOKUP_WORKSHEET(Bereik As Range, IDRow As Integer, LookupValColRelative As Integer, ReturnValColRelative As Integer, LookupID As String, LookupVal As Double) As Double
    'this function looks up a value in tables where every item has TWO columns
    'specify the row that contains the ID you're looking a value for
    'specify in which column, relative to the column that contains the ID, the Lookup Value is located
    'specify in which column, relative to the column that contains the ID, the Return Value is located
    Dim c As Integer, r As Integer
    c = 0
    Do While Not Bereik.Cells(IDRow, c + 1) = ""
      c = c + 1
      If Bereik.Cells(IDRow, c) = LookupID Then
        Do While Not Bereik.Cells(r + 1, c) = ""
            r = r + 1
            If Bereik.Cells(r, c + LookupValColRelative) = LookupVal Then
              PAIR_LOOKUP_WORKSHEET = Bereik.Cells(r, c + ReturnValColRelative)
              Exit Function
            End If
        Loop
        Exit Do
      End If
    Loop
    End Function

Public Sub PAIR_LOOKUP_ACTIVEX(InputRange As Range, InputIDRow As Integer, OutputRange As Range, ProgressRange As Range)
    'this function looks up a value in tables where every item has TWO columns
    'in the input range we expect the column containing the looked up value to be the same as the one containing the ID
    'in the input range we expect the column containing the return value to be immediately to the right of the lookup column
    Dim ro As Integer, co As Integer, ri As Integer, ci As Integer, n As Long
    Dim LookupID As String, LookupVal As Double
    Dim MaxNum As Long
    MaxNum = (OutputRange.Rows.Count - 1) * (OutputRange.Columns.Count - 1)
    For ro = 2 To OutputRange.Rows.Count
        For co = 2 To OutputRange.Columns.Count
            n = n + 1
            ProgressRange.Cells(1, 1) = n / MaxNum
            LookupID = OutputRange.Cells(ro, 1)
            LookupVal = OutputRange.Cells(1, co)
            
            For ci = 1 To InputRange.Columns.Count
                If InputRange.Cells(InputIDRow, ci) = LookupID Then
                    For ri = 1 To InputRange.Rows.Count
                        If InputRange.Cells(ri, ci) = LookupVal Then
                            OutputRange.Cells(ro, co) = InputRange.Cells(ri, ci + 1)
                            Exit For
                        End If
                    Next
                    Exit For
                End If
            Next
        Next
        DoEvents
    Next

    End Sub


Public Function VERT_HORIZ_ZOEKEN(Bereik As Range, ZoekVerticaal As String, ZoekHorizontaal As String) As Variant

Dim Kolom As Variant
Dim Rij As Variant
Dim KolomTeller As Integer
Dim RijTeller As Integer
Dim ZoekKolom As Integer
Dim ZoekRij As Integer

    KolomTeller = 0
    RijTeller = 0

    For Each Kolom In Bereik.Columns
        KolomTeller = KolomTeller + 1
        If UCase(Kolom.Columns.Cells(1, 1).Value) = UCase(ZoekHorizontaal) Then
            ZoekKolom = KolomTeller
        End If
    Next Kolom
    
    For Each Rij In Bereik.Rows
        RijTeller = RijTeller + 1
        If UCase(Rij.Rows.Cells(1, 1).Value) = UCase(ZoekVerticaal) Then
            ZoekRij = RijTeller
        End If
    Next Rij

    If ZoekKolom = 0 Or ZoekRij = 0 Then
        VERT_HORIZ_ZOEKEN = 0
    Else
        VERT_HORIZ_ZOEKEN = Bereik.Cells(ZoekRij, ZoekKolom).Value
    End If

End Function

Public Function VERT_ZOEKEN_DOUBLE(SeekValue1 As Variant, SeekValue2 As Variant, myRange As Range, ReturnCol As Long) As Variant
  'Deze functie is een uitbreiding op vertikaal zoeken, namelijk dat hij zoekt op basis van twee criteria: een waarde in kol1 en een in kol2
  Dim r As Long
  
  VERT_ZOEKEN_DOUBLE = Null
  If myRange.Columns.Count < ReturnCol Then Exit Function
  If ReturnCol < 3 Then Exit Function
  
  For r = 1 To myRange.Rows.Count
    If Trim(UCase(myRange.Cells(r, 1))) = Trim(UCase(SeekValue1)) And Trim(UCase(myRange.Cells(r, 2))) = Trim(UCase(SeekValue2)) Then
      VERT_ZOEKEN_DOUBLE = myRange.Cells(r, ReturnCol)
      Exit Function
    End If
  Next

End Function


Public Function HOR_ZOEKEN_DOUBLE(SeekValue1 As Variant, SeekValue2 As Variant, myRange As Range, ReturnRow As Long) As Variant
  'Deze functie is een uitbreiding op horizontaal zoeken, namelijk dat hij zoekt op basis van twee criteria: een waarde in rij1 en een in rij2
  Dim c As Long
  
  HOR_ZOEKEN_DOUBLE = Null
  If myRange.Rows.Count < ReturnRow Then Exit Function
  If ReturnRow < 3 Then Exit Function
  
  For c = 1 To myRange.Columns.Count
    If myRange.Cells(1, c) = SeekValue1 And myRange.Cells(2, c) = SeekValue2 Then
      HOR_ZOEKEN_DOUBLE = myRange.Cells(ReturnRow, c)
      Exit Function
    End If
  Next
  
  HOR_ZOEKEN_DOUBLE = ""
  
End Function

Public Function VERT_ZOEKEN_TRIPLE(SeekValue1 As Variant, SeekValue2 As Variant, SeekValue3 As Variant, myRange As Range, ReturnCol As Long) As Variant
  'Deze functie is een uitbreiding op vertikaal zoeken, namelijk dat hij zoekt op basis van DRIE criteria: een waarde in kol1 en een in kol2, een in Kol3
  Dim r As Long
  
  VERT_ZOEKEN_TRIPLE = Null
  If myRange.Columns.Count < ReturnCol Then Exit Function
  If ReturnCol < 4 Then Exit Function
  
  For r = 1 To myRange.Rows.Count
    If myRange.Cells(r, 1) = SeekValue1 And myRange.Cells(r, 2) = SeekValue2 And myRange.Cells(r, 3) = SeekValue3 Then
      VERT_ZOEKEN_TRIPLE = myRange.Cells(r, ReturnCol)
      Exit Function
    End If
  Next

End Function


Public Function VERT_ZOEKEN_QUADRUPLE(SeekValue1 As Variant, SeekValue2 As Variant, SeekValue3 As Variant, SeekValue4 As Variant, myRange As Range, ReturnCol As Long) As Variant
  'Deze functie is een uitbreiding op vertikaal zoeken, namelijk dat hij zoekt op basis van VIER criteria: een waarde in kol1 en een in kol2, een in Kol3 en een in Kol4
  Dim r As Long
  
  VERT_ZOEKEN_QUADRUPLE = Null
  If myRange.Columns.Count < ReturnCol Then Exit Function
  If ReturnCol < 5 Then Exit Function
  
  For r = 1 To myRange.Rows.Count
    If myRange.Cells(r, 1) = SeekValue1 And myRange.Cells(r, 2) = SeekValue2 And myRange.Cells(r, 3) = SeekValue3 And myRange.Cells(r, 4) = SeekValue4 Then
      VERT_ZOEKEN_QUADRUPLE = myRange.Cells(r, ReturnCol)
      Exit Function
    End If
  Next

End Function

Public Function VERT_ZOEKEN_MIN(ID As String, myRange As Range, ValueColIdx As Long, Optional SkipZero As Boolean = False) As Variant
 Dim r As Long, c As Long, myMin As Double, n As Long
 
 For r = 1 To myRange.Rows.Count
   If myRange.Cells(r, 1) = ID And Not (myRange.Cells(r, ValueColIdx) = 0 And SkipZero = True) Then
     n = n + 1
     If n = 1 Then
       myMin = myRange.Cells(r, ValueColIdx)
     Else
       If myRange.Cells(r, ValueColIdx) < myMin Then myMin = myRange.Cells(r, ValueColIdx)
     End If
   End If
  Next
  
  If n = 0 Then
    VERT_ZOEKEN_MIN = Nothing
  Else
    VERT_ZOEKEN_MIN = myMin
  End If
  
End Function

Public Function VERT_ZOEKEN_MAX(ID As String, myRange As Range, ValueColIdx As Long, Optional SkipZero As Boolean = False) As Double
 Dim r As Long, c As Long, myMax As Double, n As Long
 
 For r = 1 To myRange.Rows.Count
   If myRange.Cells(r, 1) = ID And Not (myRange.Cells(r, ValueColIdx) = 0 And SkipZero = True) Then
     n = n + 1
     If n = 1 Then
       myMax = myRange.Cells(r, ValueColIdx)
     Else
       If myRange.Cells(r, ValueColIdx) > myMax Then myMax = myRange.Cells(r, ValueColIdx)
     End If
   End If
  Next
  
  If n = 0 Then
    VERT_ZOEKEN_MAX = -999
  Else
    VERT_ZOEKEN_MAX = myMax
  End If

End Function

Public Function VERT_ZOEKEN_GROOTSTEAANDEELHOUDER(ID As String, myRange As Range, AandelhoudersColIdx As Long, ValueColIdx As Long, Optional Absoluut As Boolean = False) As Variant
 Dim r As Long, c As Long, i As Long, GrAand As String, mySum As Double
 Dim Aandeelhouders() As String
 Dim WaardeSom() As Double
 
 ReDim Aandeelhouders(1 To myRange.Rows.Count)  'dimensioneer op het meest pessimistische geval: alleen maar unieke waarden
 ReDim WaardeSom(1 To myRange.Rows.Count)       'dimensioneer op het meest pessimistische geval: alleen maar unieke waarden
 
 For r = 1 To myRange.Rows.Count
   If myRange.Cells(r, 1) = ID Then
     For i = 1 To UBound(Aandeelhouders())
       If Aandeelhouders(i) = myRange.Cells(r, AandelhoudersColIdx) Or Aandeelhouders(i) = "" Then
         Aandeelhouders(i) = myRange.Cells(r, AandelhoudersColIdx)
         If Not Absoluut Then
           WaardeSom(i) = WaardeSom(i) + myRange.Cells(r, ValueColIdx)
         Else
           WaardeSom(i) = WaardeSom(i) + VBA.Abs(myRange.Cells(r, ValueColIdx))
         End If
         Exit For
       End If
     Next
   End If
 Next

 For i = 1 To UBound(Aandeelhouders())
   If WaardeSom(i) > mySum Then
     mySum = WaardeSom(i)
     GrAand = Aandeelhouders(i)
   End If
 Next
  
 VERT_ZOEKEN_GROOTSTEAANDEELHOUDER = GrAand
  
End Function

Public Function HEADERBYMAXIMUMVALUE(HeaderRange As Range, ValuesRange As Range) As Variant
 'this function returns the header value that corresponds with the column containing the largest value in a given range
 Dim myVal As Variant
 Dim maxVal As Variant
 Dim Header As Variant
 maxVal = -9.99E+101
 Dim c As Long
 
 If HeaderRange.Columns.Count <> ValuesRange.Columns.Count Then
   HEADERBYMAXIMUMVALUE = "Number of columns for header and values must be equal"
 ElseIf HeaderRange.Rows.Count > 1 Then
   HEADERBYMAXIMUMVALUE = "Header Range can only have one row"
 ElseIf HeaderRange.Rows.Count > 1 Then
   HEADERBYMAXIMUMVALUE = "Values Range can only have one row"
 Else
   For c = 1 To ValuesRange.Columns.Count
     If ValuesRange.Cells(1, c) > maxVal Then
       maxVal = ValuesRange.Cells(1, c)
       Header = HeaderRange.Cells(1, c)
     End If
   Next
 End If
 
 HEADERBYMAXIMUMVALUE = Header
  
End Function

Public Function VERT_ZOEKEN_NEARESTXY(x As Double, y As Double, XYVALRANGE As Range, ReturnColIdx As Long, Optional XColIdx As Long = 1, Optional YColIdx As Long = 2) As Double
  'zoekt voor een gegeven X en Y het meest dichtbijzijnde object uit een range met X,Y en geeft de waarde uit een gespecificeerde kolom terug
  Dim r As Long
  Dim Dist As Double, minDist As Double
  Dim dX As Double, dY As Double
  Dim minDistVal As Double  'de waarde die moet worden teruggegeven
  minDist = 9999999999999#
  
  For r = 1 To XYVALRANGE.Rows.Count
    dX = x - XYVALRANGE.Cells(r, XColIdx)
    dY = y - XYVALRANGE.Cells(r, YColIdx)
    Dist = VBA.Math.Sqr(dX ^ 2 + dY ^ 2)
    If Dist < minDist Then
      minDist = Dist
      minDistVal = XYVALRANGE.Cells(r, ReturnColIdx)
    End If
  Next
  VERT_ZOEKEN_NEARESTXY = minDistVal
  
End Function


Public Function VERT_ZOEKEN_MODUS(ID As String, myRange As Range, ValueColIdx As Long) As Variant
 'geeft de meest voorkomende waarde terug behorende bij een vooraf vastgesteld ID
 Dim r As Long, c As Long, i As Long, Found As Boolean, nPoints As Long, myModus As String, MaxNum As Long
 Dim Values() As String
 Dim n() As Long
 
 ReDim Values(1 To myRange.Rows.Count)  'dimensioneer op het meest pessimistische geval: alleen maar unieke waarden
 ReDim n(1 To myRange.Rows.Count)       'dimensioneer op het meest pessimistische geval: alleen maar unieke waarden
 
 For r = 1 To myRange.Rows.Count
   If myRange.Cells(r, 1) = ID Then
     For i = 1 To UBound(Values())
       If Values(i) = myRange.Cells(r, ValueColIdx) Or Values(i) = "" Then
         Values(i) = myRange.Cells(r, ValueColIdx)
         n(i) = n(i) + 1
         Exit For
       End If
     Next
   End If
  Next
    
  MaxNum = 0
  For i = 1 To UBound(Values())
    If n(i) > MaxNum Then
      MaxNum = n(i)
      myModus = Values(i)
    End If
  Next
  VERT_ZOEKEN_MODUS = myModus

End Function

Public Function VERT_ZOEKEN_SOM(ID As String, myRange As Range, ValueColIdx As Long) As Variant
  'geeft de som terug van alle waarden uit kolomnr ValueColIdx achter een ID in kolom 1 met vooraf vastgestelde waarde
  Dim r As Long, mySum As Double
 
  mySum = 0
  For r = 1 To myRange.Rows.Count
   If myRange.Cells(r, 1) = ID Then mySum = mySum + myRange.Cells(r, ValueColIdx)
  Next
  VERT_ZOEKEN_SOM = mySum

End Function

Public Function FindColumnInRange(myRange As Range, SeekValue As Variant, assignEmptyColumnIfNotFound As Boolean) As Long
  'deze functie geeft de kolomindex terug, gegeven een gezochte waarde.
  'Let op: de range mag slechts 1 cel hoog zijn.
  Dim FirstEmpty As Long 'de eerst lege kolom die hij tegenkomt
  Dim c As Long
  
  For c = 1 To myRange.Columns.Count
    If myRange.Cells(1, c) = SeekValue Then
      FindColumnInRange = c
      Exit Function
    ElseIf myRange.Cells(1, c) = "" Then
      If FirstEmpty = 0 Then FirstEmpty = c
    End If
  Next
  
  'als hij hier aankomt, heeft hij niets gevonden. Ken dus een nieuwe kolom toe voor de data
  If assignEmptyColumnIfNotFound = True Then
    If FirstEmpty > 0 Then
      FindColumnInRange = FirstEmpty
    Else
      'tel een bij de laatste op
      FindColumnInRange = c + 1
    End If
  Else
    FindColumnInRange = 0
  End If
  
End Function

Public Function FindRowInRange(myRange As Range, SeekValue As Variant, assignEmptyRowIfNotFound As Boolean) As Long
  'deze functie geeft de rijindex terug, gegeven een gezochte waarde.
  'Let op: de range mag slechts 1 cel breed zijn.
  Dim FirstEmpty As Long 'de eerst lege rij die hij tegenkomt
  Dim r As Long
  
  For r = 1 To myRange.Rows.Count
    If myRange.Cells(r, 1) = SeekValue Then
      FindRowInRange = r
      Exit Function
    ElseIf myRange.Cells(r, 1) = "" Then
      If FirstEmpty = 0 Then FirstEmpty = r
    End If
  Next
  
  'als hij hier aankomt, heeft hij niets gevonden. Ken dus een nieuwe rij toe voor de data
  If assignEmptyRowIfNotFound = True Then
    If FirstEmpty > 0 Then
      FindRowInRange = FirstEmpty
    Else
      'tel een bij de laatste op
      FindRowInRange = r + 1
    End If
  Else
    FindRowInRange = 0
  End If
  
End Function

Public Function AVERAGEFROMRANGE(myRange As Range, c As Integer, ConditionalColumn As Integer, Condition As String, Optional ByVal UseFirstIfNoFound As Boolean = True) As Variant
  Dim r As Integer, A As Integer, n As Integer
  Dim subRange As Range
  Dim condVal As Variant, myVal As Variant, mySum As Variant
  
  If myRange.areas.Count > 1 Then
    MsgBox ("Error: ranges with multiple areas are not (yet) supported in function AVERAGEFROMRANGE.")
    End
  End If
  
  If Condition <> "" And ConditionalColumn > 0 Then
    For r = 1 To myRange.Rows.Count
      myVal = myRange.Cells(r, c)
      condVal = myRange.Cells(r, ConditionalColumn)
      If MeetsCondition(condVal, Condition) Then
        n = n + 1
        mySum = mySum + myVal
      End If
    Next
    
    If n > 0 Then
      AVERAGEFROMRANGE = mySum / n
    Else
      If UseFirstIfNoFound Then
        AVERAGEFROMRANGE = myRange.Cells(1, c)
      Else
        AVERAGEFROMRANGE = -999
      End If
    End If
    
  Else
    AVERAGEFROMRANGE = Application.WorksheetFunction.Average(myRange.Range(myRange.Cells(1, c), myRange.Cells(myRange.Rows.Count, c)))
  End If

End Function

Public Function MINFROMRANGE(myRange As Range, c As Integer, Optional ConditionalColumn As Integer = 0, Optional Condition As String = "") As Variant
  Dim minVal As Variant
  minVal = 9999999999999#

  If myRange.areas.Count > 1 Then
    MsgBox ("Error: ranges with multiple areas are not (yet) supported in function MINFROMRANGE.")
    End
  End If
  
  For r = 1 To myRange.Rows.Count
    myVal = myRange.Cells(r, c)
    condVal = myRange.Cells(r, ConditionalColumn)
    If MeetsCondition(condVal, Condition) Then
      If myVal < minVal Then minVal = myVal
    End If
  Next

  MINFROMRANGE = minVal
End Function

Public Function MAXFROMRANGE(myRange As Range, c As Integer, Optional ConditionalColumn As Integer = 0, Optional Condition As String = "") As Variant
  Dim maxVal As Variant
  maxVal = -9999999999999#

  If myRange.areas.Count > 1 Then
    MsgBox ("Error: ranges with multiple areas are not (yet) supported in function MAXFROMRANGE.")
    End
  End If
  
  For r = 1 To myRange.Rows.Count
    myVal = myRange.Cells(r, c)
    condVal = myRange.Cells(r, ConditionalColumn)
    If MeetsCondition(condVal, Condition) Then
      If myVal > maxVal Then maxVal = myVal
    End If
  Next

  MAXFROMRANGE = maxVal
End Function

Public Function FIRSTFROMRANGE(myRange As Range, c As Integer, Optional ConditionalColumn As Integer = 0, Optional Condition As String = "") As Variant
  Dim r As Integer, condVal As Variant, myVal As Variant
  If myRange.areas.Count > 1 Then
    MsgBox ("Error: ranges with multiple areas are not (yet) supported in function FIRSTFROMRANGE.")
    End
  End If
  
  For r = 1 To myRange.Rows.Count
    myVal = myRange.Cells(r, c)
    condVal = myRange.Cells(r, ConditionalColumn)
    If MeetsCondition(condVal, Condition) Then
      FIRSTFROMRANGE = myVal
      Exit Function
    End If
  Next
  
End Function

Public Function LASTFROMRANGE(myRange As Range, c As Integer, Optional ConditionalColumn As Integer = 0, Optional Condition As String = "") As Variant
  Dim r As Integer, condVal As Variant, myVal As Variant

  If myRange.areas.Count > 1 Then
    MsgBox ("Error: ranges with multiple areas are not (yet) supported in function LASTFROMRANGE.")
    End
  End If

  For r = myRange.Rows.Count To 1 Step -1
    myVal = myRange.Cells(r, c)
    condVal = myRange.Cells(r, ConditionalColumn)
    If MeetsCondition(condVal, Condition) Then
      LASTFROMRANGE = myVal
      Exit Function
    End If
  Next
  
  End Function

Public Function MOSTCOMMONFROMRANGE(myRange As Range, c As Integer, Optional ConditionalColumn As Integer = 0, Optional Condition As String = "", Optional ByVal UseFirstIfNoFound As Boolean = True) As Variant
Dim r As Long, i As Long, A As Long, Found As Boolean
Dim myVals() As Variant, myNumbers() As Long
Dim MaxNum As Long, myVal As Variant, condVal As Variant, n As Long

'This function vinds the most common value in a range.

If myRange.areas.Count > 1 Then
  MsgBox ("Error: ranges with multiple areas are not (yet) supported in function LASTFROMRANGE.")
  End
End If

n = 0

For r = 1 To myRange.Rows.Count

  myVal = myRange.Cells(r, c).Value
  condVal = myRange.Cells(r, ConditionalColumn).Value

  If MeetsCondition(condVal, Condition) Then
    
    Found = False
    If n > 0 Then
      For i = 1 To UBound(myVals)
        If myVal = myVals(i) Then
          myVals(i) = myRange.Cells(r, c)
          myNumbers(i) = myNumbers(i) + 1
          Found = True
          Exit For
        End If
      Next
    End If
    
    'if the value was not yet found in the array, add it
    If Not Found Then
      n = n + 1
      ReDim Preserve myVals(1 To n)
      ReDim Preserve myNumbers(1 To n)
      myVals(n) = myRange.Cells(r, c)
      myNumbers(n) = 1
    End If
  
  End If
Next

MaxNum = 0

If n > 0 Then
  For i = 1 To UBound(myVals)
    If myNumbers(i) > MaxNum Then
      myVal = myVals(i)
      MaxNum = myNumbers(i)
    End If
  Next
ElseIf UseFirstIfNoFound Then
  myVal = myRange.Cells(1, c)
Else
  myVal = -999
End If

MOSTCOMMONFROMRANGE = myVal

End Function

Public Sub GEWOGEN_GEMIDDELDE(myRange As Range, IDColIdx As Long, ValColIdx As Long, WeightValColIdx, resultsRow As Long, resultsCol As Long, Optional HasHeader As Boolean = True)
  'berekent een gewogen gemiddelde waarde voor ieder ID, gewogen naar bijv. oppervlaktes
  Dim myID As String, checkID As Variant, myResult As Double, myWeight As Variant, SumOfWeights As Double
  Dim IDsDone As Collection, IDDone As Boolean
  Dim Vals As Collection, Weights As Collection
  Dim r As Long, r2 As Long, r3 As Long, c As Long, i As Long, StartRow As Long
  
  r3 = resultsRow
  c = resultsCol
  ActiveSheet.Cells(r3, c) = "ID"
  ActiveSheet.Cells(r3, c + 1) = "Gewogen gemiddelde"
  
  If HasHeader Then
    StartRow = 2
  Else
    StartRow = 1
  End If
  
  Set IDsDone = New Collection
  
  For r = StartRow To myRange.Rows.Count
    myID = ActiveSheet.Cells(r, IDColIdx)
    IDDone = False
    For Each checkID In IDsDone
      If checkID = myRange.Cells(r, IDColIdx) Then
        IDDone = True
        Exit For
      End If
    Next
    
    If Not IDDone Then
      Set Vals = New Collection
      Set Weights = New Collection
      Vals.Add (ActiveSheet.Cells(r, ValColIdx))
      Weights.Add (ActiveSheet.Cells(r, WeightValColIdx))
      For r2 = r + 1 To myRange.Rows.Count
        If ActiveSheet.Cells(r2, IDColIdx) = myID Then
          Vals.Add (ActiveSheet.Cells(r2, ValColIdx))
          Weights.Add (ActiveSheet.Cells(r2, WeightValColIdx))
        End If
      Next
      
      SumOfWeights = 0
      'bereken de som van alle gewichten
      For Each myWeight In Weights
        SumOfWeights = SumOfWeights + myWeight
      Next
      
      myResult = 0
      'bereken de gewogen waarde
      For i = 1 To Vals.Count
        If SumOfWeights <> 0 Then
          myResult = myResult + Vals(i) * Weights(i) / SumOfWeights
        Else
          myResult = 0
        End If
      Next
      
      'schrijf weg
      r3 = r3 + 1
      ActiveSheet.Cells(r3, c) = myID
      ActiveSheet.Cells(r3, c + 1) = myResult
      
     Call IDsDone.Add(myID)
    End If
    
  Next
  
End Sub

Public Sub AGGREGEREN(myRange As Range, resultsRow As Long, resultsCol As Long, ExportEachnRows As Long)
  'Assumes date/time in first column and a header row.
  Dim r As Long, c As Long, r2 As Long, c2 As Long
  
  r2 = resultsRow
  c2 = resultsCol
  ActiveSheet.Cells(r2, c2) = "Datum/Tijd"
  For c = 2 To myRange.Columns.Count
    ActiveSheet.Cells(r2, c2 + c - 1) = myRange.Cells(1, c)
  Next
    
  For r = 2 To myRange.Rows.Count Step ExportEachnRows
    DoEvents
    r2 = r2 + 1
    
    ActiveSheet.Cells(r2, c2) = myRange(r, 1)  'write date/time
    For c = 2 To myRange.Columns.Count
      ActiveSheet.Cells(r2, c2 + c - 1) = myRange(r, c)
    Next
  Next
End Sub

Public Sub AGGREGATEFROMRANGE(RangeIncludingHeader As Range, AggregateByColumn As Integer, AggregateColumn As Integer, myMethod As enmAggregateMethod, resultsRow As Integer, resultsCol As Integer)
  Dim r1 As Long, r2 As Long, r As Long
  Dim StartRow As Long, EndRow As Long
  Dim myVal As Variant, Col1Val As Variant
  Dim subRange As Range
  Dim i As Long
  
  Dim curVal As Variant
  curVal = ""
  
  'write the results header
  ActiveSheet.Cells(resultsRow, resultsCol) = RangeIncludingHeader.Cells(1, AggregateByColumn)
  ActiveSheet.Cells(resultsRow, resultsCol + 1) = RangeIncludingHeader.Cells(1, AggregateColumn)
  
  'walk through the data and find unique blocks based on the AggregateColumn
  For r1 = 2 To RangeIncludingHeader.Rows.Count
    If RangeIncludingHeader.Cells(r1, AggregateByColumn) <> curVal And RangeIncludingHeader.Cells(r1, AggregateByColumn) <> "" Then
       curVal = RangeIncludingHeader.Cells(r1, AggregateByColumn)
       StartRow = r1
       For r2 = r1 + 1 To RangeIncludingHeader.Rows.Count
       
         'as soon as the next row in the aggregatebycolumn colummn changes, exit the loop and compute the aggregated value
         If RangeIncludingHeader.Cells(r2, AggregateByColumn) <> RangeIncludingHeader.Cells(r1, AggregateByColumn) Then
           EndRow = r2 - 1
           r1 = EndRow
           Exit For
         End If
       Next
       
       If EndRow > StartRow Then
       
       resultsRow = resultsRow + 1
       Set subRange = RangeIncludingHeader.Range(RangeIncludingHeader.Cells(StartRow, AggregateColumn), RangeIncludingHeader.Cells(EndRow, AggregateColumn))
         
         Col1Val = RangeIncludingHeader.Cells(StartRow, AggregateByColumn)
         
         If myMethod = Average Then
'           If Application.WorksheetFunction.Sum(SubRange) > 0 Then
'             myVal = Application.Average(SubRange)
'           Else
'             myVal = 0
'           End If
            Dim mySum As Double
            mySum = 0
            For i = StartRow To EndRow
              mySum = mySum + RangeIncludingHeader.Cells(i, AggregateColumn)
            Next
            If mySum > 0 Then
              myVal = mySum / (EndRow - StartRow + 1)
            Else
              myVal = 0
            End If
         ElseIf myMethod = first Then
           myVal = RangeIncludingHeader.Cells(StartRow, AggregateColumn).Value
         ElseIf myMethod = Last Then
           myVal = RangeIncludingHeader.Cells(EndRow, AggregateColumn).Value
         ElseIf myMethod = Largest Then
           myVal = Application.max(subRange)
         ElseIf myMethod = Smallest Then
           myVal = Application.Min(subRange)
         ElseIf myMethod = Most Then
           myVal = MOSTCOMMONFROMRANGE(subRange, 1)
         ElseIf myMethod = Sum Then
            myVal = Application.Sum(subRange)
         End If
         
        ActiveSheet.Cells(resultsRow, resultsCol) = Col1Val
        ActiveSheet.Cells(resultsRow, resultsCol + 1) = myVal
       
       End If
         
    End If
  Next

End Sub


Public Sub AGGREGATERANGECONDITIONALLY(myRange As Range, AggregateColumn As Integer, AggregateMethod() As enmAggregateMethod, ConditionalColumn As Integer, Condition() As String, resultsRow As Integer, resultsCol As Integer)
  Dim r1 As Integer, r2 As Integer, r As Integer, c As Integer, A As Integer
  Dim StartRow As Integer, EndRow As Integer
  Dim myMethod As enmAggregateMethod
  Dim myVal As Variant, subRange As Range, myCond As String
    
  If myRange.Columns.Count <> UBound(AggregateMethod) Then
    MsgBox ("Error: array for aggregation method must have same dimensions as number of columns in range.")
    End
  End If
  
  Dim curVal As Variant
  curVal = ""
  
  'write the results header
  For c = 1 To myRange.Columns.Count
    ActiveSheet.Cells(resultsRow, resultsCol + c - 1) = myRange.Cells(1, c)
  Next
  
  'walk through the data and find unique blocks based on the AggregateColumn
  For r1 = 2 To myRange.Rows.Count
    If ActiveSheet.Cells(r1, AggregateColumn) <> curVal And ActiveSheet.Cells(r1, AggregateColumn) <> "" Then
       curVal = ActiveSheet.Cells(r1, AggregateColumn)
       StartRow = r1
       For r2 = r1 + 1 To myRange.Rows.Count
         If ActiveSheet.Cells(r2, AggregateColumn) <> ActiveSheet.Cells(r1, AggregateColumn) Then
           EndRow = r2 - 1
           Exit For
         End If
       Next
       
       resultsRow = resultsRow + 1
       
       'create a subrange for the block we're in
       Set subRange = myRange.Range(myRange.Cells(StartRow, 1), myRange.Cells(EndRow, myRange.Columns.Count))
       
       For c = 1 To subRange.Columns.Count
         myMethod = AggregateMethod(c)
         myCond = Condition(c)
         
         If myMethod = Average Then
           myVal = AVERAGEFROMRANGE(subRange, c, ConditionalColumn, myCond)
         ElseIf myMethod = first Then
           myVal = FIRSTFROMRANGE(subRange, c, ConditionalColumn, myCond)
         ElseIf myMethod = Last Then
           myVal = LASTFROMRANGE(subRange, c, ConditionalColumn, myCond)
         ElseIf myMethod = Largest Then
           myVal = MAXFROMRANGE(subRange, c, ConditionalColumn, myCond)
         ElseIf myMethod = Smallest Then
           myVal = MINFROMRANGE(subRange, c, ConditionalColumn, myCond)
         ElseIf myMethod = Most Then
           myVal = MOSTCOMMONFROMRANGE(subRange, c, ConditionalColumn, myCond)
         End If
         
        ActiveSheet.Cells(resultsRow, resultsCol + c - 1) = myVal
         
       Next
    End If
  Next

End Sub

Public Function COLUMNFROMRANGE(myRange As Range, ColNum As Integer) As Range
  Dim myArea As Range, newRange As Range, subRange As Range
  Dim A As Integer
  
  For A = 1 To myRange.areas.Count
    Set myArea = myRange.areas(A)
    Set subRange = myArea.Range(myArea.Cells(-1, ColNum), myArea.Cells(myArea.Rows.Count - 2, ColNum))
    If newRange Is Nothing Then
      Set newRange = subRange
    Else
      Set newRange = Union(newRange, subRange)
    End If
  Next
    
  Set COLUMNFROMRANGE = newRange

End Function

Public Function CONDITIONALSUBRANGE(ByVal myRange As Range, ByVal ConditionColumn As Integer, ByVal Condition As String) As Range
  'this function applies a given condition to a range and only returns the rows for wich the condtion is met
  'conditions can be: "> x, >= x, < x, <= x, = x, <> x
  Dim newRange As Range, Range2 As Range
  Dim r As Integer
  Dim Operator As String, Operand As Double
  Dim myVal As Double, Inuse As Boolean
  
  Condition = VBA.Trim(Condition)
  
  If InStr(1, Condition, " ") <= 0 Then
    MsgBox ("Condition is not valid. Must contain space between operator and operand: " & Condition)
    End
  End If
  
  
  Operator = ParseString(Condition, " ")
  Operand = Condition
  
  For r = 1 To myRange.Rows.Count
  
    'decide whether the condition is met for this row
    Inuse = False
    myVal = myRange.Cells(r, ConditionColumn)
    Select Case Operator
      Case Is = ">"
         If myVal > Operand Then Inuse = True
      Case Is = ">="
         If myVal >= Operand Then Inuse = True
      Case Is = "<"
         If myVal < Operand Then Inuse = True
      Case Is = "<="
         If myVal <= Operand Then Inuse = True
      Case Is = "<>"
         If myVal <> Operand Then Inuse = True
      Case Is = "="
         If myVal = Operand Then Inuse = True
      Case Else
        MsgBox ("Error: operator in conditional formatting was not recognized or is not supported: " & Operator)
        End
    End Select
    
    'if the condition is met, add the row to our new range
    Dim n As Integer
    If Inuse = True Then
      n = n + 1
      If newRange Is Nothing Then
        Set newRange = myRange.Range(myRange.Cells(r, 1), myRange.Cells(r, myRange.Columns.Count))
      Else
        Set Range2 = myRange.Range(myRange.Cells(r, 1), myRange.Cells(r, myRange.Columns.Count))
        Set newRange = Union(newRange, Range2)
      End If
    End If
  Next
  
  Set CONDITIONALSUBRANGE = newRange

End Function



Public Sub AGGREGERENNAARUREN(DATETIMERANGE As Range, ValRange As Range, ProgressRange As Range, resultrow As Long, resultcol As Long, Optional HeleUren As Boolean = True)
Dim r As Long, r2 As Long, c2 As Long, lastProgress As Double, Progress As Double

'voorbeeld aanroep AGGREGERENNAARUREN(Range(Cells(2, 1), Cells(2000, 1)), Range(Cells(2, 2), Cells(2000, 2)), 2, 3)
r2 = resultrow
c2 = resultcol
ActiveSheet.Cells(r2, c2) = "Datum/Tijd"
ActiveSheet.Cells(r2, c2 + 1) = "Waarde"


If HeleUren = True Then
  For r = 1 To DATETIMERANGE.Rows.Count
    Progress = r / DATETIMERANGE.Rows.Count * 100
    If Round(Progress, 0) > Round(lastProgress, 0) Then
      ProgressRange.Value = Progress
      DoEvents
      lastProgress = Progress
    End If
    
    If Minute(DATETIMERANGE(r, 1)) = 0 Then
      r2 = r2 + 1
      ActiveSheet.Cells(r2, c2) = DATETIMERANGE(r, 1)
      ActiveSheet.Cells(r2, c2 + 1) = ValRange(r, 1)
    End If
  Next
Else
  MsgBox ("Optie nog niet ondersteund")
End If

End Sub

Public Function CountSequentialExceedances(ValRange As Range, threshold As Double) As Integer
  Dim r As Integer, n As Integer, nMax As Integer
  
  If ValRange.Columns.Count > 1 Then
    MsgBox ("Error in function CountSequentialExceedances. Range can have no more than one column.")
  Else
    For r = 1 To ValRange.Rows.Count
      If ValRange.Cells(r, 1) > threshold Then
        n = n + 1
        If n > nMax Then nMax = n
      Else
        n = 0
      End If
    Next
  End If
  
  CountSequentialExceedances = nMax

End Function


Public Sub GETASCIIGRIDVALUES(path As String, XYVALRANGE As Range, Optional XColIdx As Long = 1, Optional YColIdx As Long = 2, Optional ValColIdx As Long = 3)
  'haalt voor gegeven XY-coordinaten de bijbehorende waarde uit een ASCII-grid en schrijft deze naar het werkblad
  Dim Data() As Double
  Dim nCols As Long, nRows As Long, xllcorner As Double, yllcorner As Double, cellsize As Double, nodata_value As Double
  Dim x As Double, y As Double, Val As Double, rowIdx As Long, colIdx As Long
  Dim yulcorner As Double, xlrcorner As Double
  Dim r As Long, c As Long
  
  Call READASCIIGRID(path, nCols, nRows, xllcorner, yllcorner, cellsize, nodata_value, Data)
  yulcorner = yllcorner + cellsize * nRows
  xlrcorner = xllcorner + cellsize * nCols
    
  For r = 1 To XYVALRANGE.Rows.Count
   
    x = XYVALRANGE(r, XColIdx)
    y = XYVALRANGE(r, YColIdx)
    
    If x >= xllcorner And x <= xlrcorner And y > yllcorner And y < yulcorner Then
      colIdx = Application.WorksheetFunction.RoundUp((x - xllcorner) / cellsize, 0)
      rowIdx = Application.WorksheetFunction.RoundUp((yulcorner - y) / cellsize, 0)
      XYVALRANGE.Cells(r, ValColIdx) = Data(rowIdx, colIdx)
    Else
      XYVALRANGE.Cells(r, ValColIdx) = nodata_value
    End If
  Next
  
End Sub

Public Sub getRowColFromASCIIGRID(xllcenter As Double, yllcenter As Double, nCols As Long, nRows As Long, dX As Double, dY As Double, x As Double, y As Double, ByRef myRow As Long, ByRef myCol As Long)
  Dim xllcorner As Double, yllcorner As Double, yurcorner As Double
  xllcorner = xllcenter - dX / 2
  yllcorner = yllcenter - dY / 2
  yurcorner = yllcorner + dY * nRows
  
  myCol = Application.WorksheetFunction.RoundUp((x - xllcorner) / dX, 0)
  If myCol <= 0 Or myCol > nCols Then myCol = 0
  
  myRow = Application.WorksheetFunction.RoundUp((yurcorner - y) / dY, 0)
  If myRow <= 0 Or myRow > nRows Then myRow = 0
  
End Sub


Public Sub RANGEWITHHEADER2THREECOLRANGE(myRange As Range, HeaderTitle As String, resultsRow As Long, resultsCol As Long)
  'deze routine converteert een reeks waarin X en Y data staan en waarboven telkens een header staat naar een reeks met ID, X en Y in drie kolommen
  'dus van:
         'ID MyID
         'X1 Y1
         'X2 Y2
  'naar:
         'ID X1 Y1
         'ID X2 Y2
  Dim myID As String
  Dim rowIdx As Long
  Dim r As Long
  Dim c As Long
  
  r = resultsRow - 1
  c = resultsCol
  
  For rowIdx = 1 To myRange.Rows.Count
    If myRange.Cells(rowIdx, 1) = HeaderTitle Then
      myID = myRange.Cells(rowIdx, 2)
    Else
      r = r + 1
      ActiveSheet.Cells(r, c) = myID
      ActiveSheet.Cells(r, c + 1) = myRange.Cells(rowIdx, 1)
      ActiveSheet.Cells(r, c + 2) = myRange.Cells(rowIdx, 2)
    End If
  Next

End Sub

Public Sub WEAVETABLESBLOCKINTERPOLATION(myTable1 As Range, myTable2 As Range, resultsRow As Long, resultsCol As Long)
  
  'deze routine weeft twee tabellen (met verspringende x-waarden) ineen
  'gaat standaard uit van blokinterpolatie en als voorgaande waarden ontbreken 0
  Dim Table1() As Variant
  Dim Table2() As Variant
  
  'zorg dat beide ranges geen lege cellen in de eerste kolom bevatten
  'Set myTable1 = TRUNCATERANGEBYEMPTYROWS(myTable1) ROUTINE BEVAT FOUT
  'Set myTable2 = TRUNCATERANGEBYEMPTYROWS(myTable2)
  
  'no 2D-array because the first dimension cannot be resized with redim preserve
  Dim Table3 As Variant
  
  Dim maxRows As Long
  Dim Row As Long, Col As Long
  Dim i1 As Long, i2 As Long, i3 As Long
  Dim Table1Done As Boolean, Table2Done As Boolean, Done As Boolean
  Dim LastVal1 As Double, LastVal2 As Double
  Dim NextVal1 As Double, NextVal2 As Double
  
  Table1 = myTable1
  Table2 = myTable2
  LastVal1 = -9999
  LastVal2 = -9999
  NextVal1 = -9999
  NextVal2 = -9999
  
  maxRows = UBound(Table1, 1) + UBound(Table2, 1)
  ReDim Table3(1 To maxRows, 1 To 3)
  
  If Table1(1, 1) <> Table2(1, 1) Then
    MsgBox ("Error: beide tabellen moeten starten met dezelfde x-waarde")
    End
  End If

  i1 = 1
  i2 = 1
  i3 = 1
  
  Table3(i3, 1) = Table1(i1, 1)
  Table3(i3, 2) = Table1(i1, 2)
  Table3(i3, 3) = Table2(i2, 2)
  
  'nu de rest
  While Not (Table1Done And Table2Done)
    
    'If i3 = 159 Then Stop
    
    If i1 >= UBound(Table1, 1) Then Table1Done = True
    If i2 >= UBound(Table2, 1) Then Table2Done = True
    
    If Table1Done And Table2Done Then
      'do nothing
    ElseIf Table1Done And Not Table2Done And i2 < UBound(Table2, 1) Then
      'finish table 2
      i2 = i2 + 1
      i3 = i3 + 1
      
      Table3(i3, 1) = Table2(i2, 1)
      Table3(i3, 2) = Table1(i1, 2)
      Table3(i3, 3) = Table2(i2, 2)
    ElseIf Table2Done And Not Table1Done And i1 < UBound(Table1, 1) Then
      'finish table1
      i1 = i1 + 1
      i3 = i3 + 1
      Table3(i3, 1) = Table1(i1, 1)
      Table3(i3, 2) = Table1(i1, 2)
      Table3(i3, 3) = Table2(i2, 2)
    ElseIf i1 < UBound(Table1, 1) And i2 < UBound(Table2, 1) Then
      NextVal1 = Table1(i1 + 1, 1)
      NextVal2 = Table2(i2 + 1, 1)
      
      If NextVal1 < NextVal2 Then
        'move one up in table 1
        i1 = i1 + 1
        i3 = i3 + 1
        Table3(i3, 1) = Table1(i1, 1)
        Table3(i3, 2) = Table1(i1, 2)
        Table3(i3, 3) = Table2(i2, 2) 'de vorige waarde uit tabel 2 is nog altijd van toepassing
      ElseIf NextVal2 < NextVal1 Then
        'move one up in table 2
        i2 = i2 + 1
        i3 = i3 + 1
        Table3(i3, 1) = Table2(i2, 1)
        Table3(i3, 2) = Table1(i1, 2) 'de vorige waarde uit tabel 1 is nog altijd van toepassing
        Table3(i3, 3) = Table2(i2, 2)
      ElseIf NextVal1 = NextVal2 Then
        'move one up in both tables
        i1 = i1 + 1
        i2 = i2 + 1
        i3 = i3 + 1
        Table3(i3, 1) = Table1(i1, 1)
        Table3(i3, 2) = Table1(i1, 2)
        Table3(i3, 3) = Table2(i2, 2)
      End If
    End If
        
  Wend
  
  'ReDim Preserve Table3(1 To i3, 1 To 3)
  
  'write the woven table to the worksheet
  Row = resultsRow
  Col = resultsCol
  ActiveSheet.Cells(Row, Col) = "X"
  ActiveSheet.Cells(Row, Col + 1) = "YTable1"
  ActiveSheet.Cells(Row, Col + 2) = "YTable2"
  
  Row = Row + 1
  
  Call PrintArray(Table3, ActiveSheet.Range(Cells(Row, Col), Cells(Row, Col)))
  
  Exit Sub
End Sub

Public Function TRUNCATERANGEBYEMPTYROWS(ByRef myRange As Range) As Range
  Dim StartRow As Long, EndRow As Long
  Dim r As Long, i As Long
  
  For i = 1 To myRange.Rows.Count
    If myRange.Cells(i, 1) <> "" Then
      StartRow = i
      Exit For
    End If
  Next
  
  For i = myRange.Rows.Count To 1 Step -1
    If myRange.Cells(i, 1) <> "" Then
      EndRow = i
      Exit For
    End If
  Next
  
  Set TRUNCATERANGEBYEMPTYROWS = myRange.Range(myRange.Cells(StartRow, 1), myRange.Cells(EndRow, myRange.Columns.Count))
  
End Function

Public Function GETIJDEN_SINUS(Amplitude As Double, Periode As Double, TijdstipNul As Double, Evenwichtswaterstand As Double, DatumTijd As Double) As Double
    GETIJDEN_SINUS = Amplitude / 2 * Sin(2 * 3.1415 / Periode * (DatumTijd - TijdstipNul)) + Evenwichtswaterstand
End Function

Public Function BRIDGE_TABULATED_PROFILE_FROM_YZ(ProfileYRange As Range, ProfileZRange As Range, LIDYRANGE As Range, LIDZRANGE As Range, ResultsStartRow As Integer, ResultsStartCol As Integer) As Boolean
    
    'NIET AF!!!!!
    Dim ElevationList(1 To ProfileYRange.Rows.Count + LIDYRANGE.Rows.Count)
    Dim WidthList(1 To ProfileYRange.Rows.Count + LIDYRANGE.Rows.Count)
    
    If ProfileYRange.Rows.Count <> ProfileZRange.Rows.Count Then
        MsgBox ("Error: range of profile Z values must be of equal length as range of Y values.")
        BRIDGE_TABULATED_PROFILE_FROM_YZ = False
    ElseIf LIDYRANGE.Rows.Count <> LIDZRANGE.Rows.Count Then
        MsgBox ("Error: range of lid Z values must be of equal length as range of Y values.")
        BRIDGE_TABULATED_PROFILE_FROM_YZ = False
    Else
        Dim LowestIdx As Integer
        Dim LowestZ As Double
        LowestIdx = -1
        LowestZ = 9E+99
        For i = 1 To ProfileYRange.Rows.Count
            If ProfileZRange.Cells(i, 1) < LowestZ Then
                LowestIdx = i
                LowestZ = ProfileZRange.Cells(i, 1)
            End If
        Next
        'write the lowest point to our worksheet
        ElevationList(1) = LowestZ
        WidthList(1) = 0
    End If

    'probe left and right
    Dim Done As Boolean
    While Not Done
        
    Wend

End Function

Public Function QHEVEL(Diameter As Double, Lengte As Double, Chezy As Double, muin As Double, muUit As Double, muBuig As Double, dH As Double) As Double

Dim A As Double
Dim P As Double
Dim Friction As Double
Dim mu As Double

A = pi * (Diameter / 2) ^ 2
P = 2 * pi * (Diameter / 2)

Friction = (2 * 9.81 * Lengte) / (Chezy ^ 2 * A / P)
mu = 1 / (Sqr(muin + muUit + Friction + muBuig))

QHEVEL = mu * A * Sqr(2 * 9.81 * dH)

End Function

Public Function QDUIKER(Diameter As Double, Lengte As Double, Chezy As Double, muin As Double, muUit As Double, dH As Double) As Double

Dim A As Double
Dim P As Double
Dim Friction As Double
Dim mu As Double

A = pi * (Diameter / 2) ^ 2
P = 2 * pi * (Diameter / 2)

Friction = (2 * 9.81 * Lengte) / (Chezy ^ 2 * A / P)
mu = 1 / (Sqr(muin + muUit + Friction))

QDUIKER = mu * A * Sqr(2 * 9.81 * dH)

End Function

Public Function QDUIKERRECHTHOEK(BOB As Double, Breedte As Double, Hoogte As Double, Lengte As Double, Chezy As Double, muin As Double, muUit As Double, h1 As Double, h2 As Double) As Double

Dim A As Double
Dim P As Double
Dim Friction As Double
Dim mu As Double


If h1 >= Hoogte + BOB Then
  'geheel gevuld
  A = Breedte * Hoogte
  P = Breedte * 2 + Hoogte * 2
Else
  'gedeeltelijk gevuld
  A = Breedte * (h1 - BOB)
  P = Breedte + (h1 - BOB) * 2
End If

Friction = (2 * 9.81 * Lengte) / (Chezy ^ 2 * A / P)
mu = 1 / (Sqr(muin + muUit + Friction))

QDUIKERRECHTHOEK = mu * A * Sqr(2 * 9.81 * (h1 - h2))

End Function


Public Function QORIFICE(z As Double, W As Double, gh As Double, mu As Double, cw As Double, h1 As Double, h2 As Double) As Double
'Z = crest level
'W = width
'gh = gate height (openningshoogte)
'mu = contraction coef (standaard 0.63)
'cw = lateral contraction coef
'h1 = waterstand bovenstrooms
'h2 = waterstand benedenstrooms
'ce = afvoercoefficient. standaard 1.5

Dim Af As Double
Dim ce As Double
Dim g As Double
Dim u As Double 'stroomsnelheid over de kruin. Moet eigenlijk iteratief worden bepaald maar ik zet hem even op 1
u = 1
ce = 1.5
g = 9.81

'bepaal of hij verdronken of vrij is
If (h1 - z) >= (3 / 2 * gh) Then   'orifice flow
  If h2 <= (z + gh) Then 'free orifice flow
    Af = W * mu * gh
    QORIFICE = cw * W * mu * gh * VBA.Sqr(2 * g * (h1 - (z + mu * gh)))
  ElseIf h2 > (z + gh) Then 'submerged orifice flow
    Af = W * mu * gh
    QORIFICE = cw * W * mu * gh * VBA.Sqr(2 * g * (h1 - h2))
  End If
ElseIf (h1 - z) < (3 / 2 * gh) Then 'weir flow
  If (h1 - z) > (3 / 2 * (h2 - z)) Then 'free weir flow
    Af = W * 2 / 3 * (h1 - z)
    QORIFICE = cw * W * 2 / 3 * VBA.Sqr(2 / 3 * g * (h1 - z) ^ 3 / 2)
  ElseIf (h1 - z) <= (3 / 2 * (h2 - z)) Then 'submerged weir flow
    Af = W * (h1 - z - u ^ 2 / (2 * g))
    QORIFICE = ce * cw * W * (h1 - z - (u ^ 2 / (2 * g))) * VBA.Sqr(2 * g * (h1 - h2))
  End If
Else
  MsgBox ("Error: kon niet bepalen of orifice verdronken of vrij was.")
End If


End Function

Public Function QSTUW(Breedte As Double, DischCoef As Double, h1 As Double, h2 As Double, z As Double, Optional LatContrCoef As Double = 1) As Double
  Dim Hup As Double, Hdown As Double, Multiplier As Double

'Free flow: als h2 - z < 2/3 * (h1 -z)
If h1 >= h2 Then
  Hup = h1
  Hdown = h2
  Multiplier = 1
Else
  Hup = h2
  Hdown = h1
  Multiplier = -1
End If

If Hup <= z Then
  QSTUW = 0
ElseIf Hdown < z Or (Hdown - z) < 2 / 3 * (Hup - z) Then
  'Free flow: Q = c * B * 2/3 * SQRT(2/3 * g) * (h1 - z)^1.5
  QSTUW = Multiplier * DischCoef * LatContrCoef * Breedte * 2 / 3 * Sqr(2 / 3 * 9.81) * (Hup - z) ^ 1.5
Else
  'Drowned flow: Q = c * B * (h2 -z) * SQRT(2 * g *(h1 - h2))
  QSTUW = Multiplier * DischCoef * LatContrCoef * Breedte * (Hdown - z) * Sqr(2 * 9.81 * (Hup - Hdown))
End If

End Function


Public Function QABUTMENTBRIDGE(h1 As Double, h2 As Double, ProfileVerticalShift As Double, BridgeVerticalShift As Double, muinlet As Double, outletlosscoef As Double, Length As Double, nManning As Double, BridgeTableProfileZRange As Range, BridgeTableProfileWRange As Range, ProfileYRange As Range, ProfileZRange As Range, Optional ByVal MaximizeMu As Boolean = False) As Double
    
    'dit is een complexe functie omdat t.b.v. de outlet loss ook het benedenstroomse profiel bekend moet zijn.
    'let op: om numerieke redenen wordt in SOBEK de uiteindelijke mu gemaximaliseerd op 1. Dit is niet altijd terecht en kan resulteren in een overschatting van het verval over de brug.
    Dim mu As Double
    Dim muoutlet As Double
    Dim mufric As Double
    Dim Chezy As Double
    
    'hydraulic variables
    Dim ABridge As Double, PBridge As Double, RBridge As Double
    Dim AProfile As Double, PProfile As Double, RProfile As Double
    
    
    'determine the water depth inside the bridge and inside the profile
    Dim BridgeDepth As Double
    Dim ProfileDepth As Double
    BridgeDepth = h1 - BridgeTableProfileZRange.Cells(1, 1) - ProfileVerticalShift
    ProfileDepth = h2 - Application.WorksheetFunction.Min(ProfileZRange) - ProfileVerticalShift
    
    'determine the hydraulic properties
    Call TabulatedHydraulicProperties(BridgeTableProfileZRange, BridgeTableProfileWRange, BridgeDepth, ABridge, PBridge, RBridge)
    Call YZHydraulicProperties(ProfileYRange, ProfileZRange, ProfileDepth, AProfile, PProfile, RProfile)
    Chezy = Manning2Chezy(nManning, RBridge)
    
    'calculate the energy loss coefficients
    muoutlet = outletlosscoef * (1 - ABridge / AProfile) ^ 2
    mufric = (2 * 9.81 * Length) / (Chezy ^ 2 * RBridge)
    
    mu = 1 / Math.Sqr(muinlet + muoutlet + mufric)
    If MaximizeMu Then mu = Application.WorksheetFunction.Min(mu, 1)
    
    QABUTMENTBRIDGE = mu * ABridge * Math.Sqr(2 * 9.81 * (h1 - h2))
    

End Function


Public Function TabulatedWettedPerimeterByWaterlevel(Waterlevel As Double, ZRange As Range, WRange As Range, ProfileVerticalShift As Double) As Double
    Dim A As Double
    Dim P As Double
    Dim r As Double
    Dim Depth As Double
    Depth = Waterlevel - ZRange.Cells(1, 1) - ProfileVerticalShift
    Call TabulatedHydraulicProperties(ZRange, WRange, Depth, A, P, r)
    TabulatedWettedPerimeterByWaterlevel = P
End Function

Public Function TabulatedHydraulicRadiusByWaterlevel(Waterlevel As Double, ZRange As Range, WRange As Range, ProfileVerticalShift As Double) As Double
    Dim A As Double
    Dim P As Double
    Dim r As Double
    Dim Depth As Double
    Depth = Waterlevel - ZRange.Cells(1, 1) - ProfileVerticalShift
    Call TabulatedHydraulicProperties(ZRange, WRange, Depth, A, P, r)
    TabulatedHydraulicRadiusByWaterlevel = r
End Function


Public Function BREEDTE_STUW(AreaM2 As Double, Q_MMPD As Double, DischCoef As Double, OVStraalM As Double, Optional LatContrCoef As Double = 1) As Double
  
  'Free flow: Q = c * B * 2/3 * SQRT(2/3 * g) * (h1 - z)^1.5
  Dim Q As Double
  Q = AreaM2 * Q_MMPD / 1000 / 24 / 3600
  BREEDTE_STUW = Q / (DischCoef * LatContrCoef * 2 / 3 * Sqr(2 / 3 * 9.81) * OVStraalM ^ 1.5)
  
End Function
Public Function WeirSubmerged(h1 As Double, h2 As Double, z As Double) As Boolean

'Free flow: als h2 - z < 2/3 * (h1 -z)
If h2 - z < 2 / 3 * (h1 - z) Then
  WeirSubmerged = False
Else
  WeirSubmerged = True
End If

End Function

Public Function DHGEVULDERONDEDUIKER(Q As Double, D As Double, L As Double, n_manning As Double, zi As Double, zo As Double) As Double
  
  'Q = mu * A * SQR(2 * g * dh)
  'dh = Q^2 / (mu^2 * A^2 * 2g)
  
  Dim mu As Double
  Dim Chezy As Double
  Dim zf As Double 'ruwheidsverlies
  Dim A As Double, P As Double, r As Double 'natte doorsnede, natte omtrek en hydraulische straal
  
    'we gaan uit van een volledig gevulde duiker
    A = 3.141592 * (D / 2) ^ 2
    P = 2 * 3.141592 * (D / 2)
    r = A / P
    Chezy = Manning2Chezy(n_manning, r)
    zf = (2 * 9.81 * L) / (Chezy ^ 2 * r)
    mu = 1 / Sqr(zi + zo + zf)
    DHGEVULDERONDEDUIKER = Q ^ 2 / (mu ^ 2 * A ^ 2 * 2 * 9.81)
  
End Function

Public Function DHRONDEDUIKER(Q As Double, D As Double, L As Double, h1 As Double, BOB1 As Double, n_manning As Double, zi As Double, zo As Double) As Double
  
  'Q = mu * A * SQR(2 * g * dh)
  'dh = Q^2 / (mu^2 * A^2 * 2g)
  
  Dim mu As Double
  Dim Chezy As Double
  Dim zf As Double 'ruwheidsverlies
  Dim A As Double, P As Double, r As Double 'natte doorsnede, natte omtrek en hydraulische straal
  
    'we gaan uit van een volledig gevulde duiker
    A = OPPERVLAKAFGEPLATTECIRKEL(D / 2, BOB1 + D / 2, h1)
    P = NATTEOMTREKAFGEPLATTECIRKEL(D / 2, BOB1 + D / 2, h1)
    r = A / P
    Chezy = Manning2Chezy(n_manning, r)
    zf = (2 * 9.81 * L) / (Chezy ^ 2 * r)
    mu = 1 / Sqr(zi + zo + zf)
    DHRONDEDUIKER = Q ^ 2 / (mu ^ 2 * A ^ 2 * 2 * 9.81)
  
End Function

Public Function WidthOrifice(Q As Double, h1 As Double, Drempel As Double, ContrCoef As Double, DisCoef As Double) As Double
  'berekent de benodigde breedte van een niet-verdronken orifice gegeven een debiet en drempelhoogte
  'Q = mu * c * B * d * SQR(2 * g * (h1 - (z + mu*d)))
  'we gaan uit van een onderkant-schuif die hoger ligt dan het waterpeil, dus d = h1-z
  'B = Q / (mu * c * (h1-z) * sqr(2 * g * (h1-z)))
  
  WidthOrifice = Q / (ContrCoef * DisCoef * (h1 - Drempel) * Sqr(2 * 9.81 * (h1 - Drempel)))
  

End Function

Public Function HydraulicRadius(A As Double, P As Double) As Double
  'calculates the hydraulic radius from wetted area and wetted perimeter
  If P <= 0 Or A <= 0 Then
    HydraulicRadius = 0
  Else
    HydraulicRadius = A / P
  End If
End Function

Public Function Manning2Chezy(n_manning As Double, r As Double) As Double
  'computes the Chezy roughness value from manning's coefficient and the hydraulic radius
  Manning2Chezy = r ^ (1 / 6) / n_manning
End Function
Public Function Chezy2Manning(Chezy As Double, r As Double) As Double
  Chezy2Manning = r ^ (1 / 5) / Chezy
End Function
Public Function Chezy(c As Double, Depth As Double, Width As Double, bedslope As Double) As Double
    Dim Q As Double
    Q = Depth * Width * c * Sqr((Depth * Width) / (Width + 2 * Depth) * bedslope)
    Chezy = Q
End Function

Public Function MaatgevendeAfvoer(Oppervlak_ha As Double, Optional Aanvoer_lpspha As Double = 1.5) As Double
  'oppervlak in ha
  'Aanvoer in l/s/ha
  'resultaat in m3/s
  
  MaatgevendeAfvoer = Oppervlak_ha * Aanvoer_lpspha / 1000

End Function

Public Function NeerslagPatroon(ValueRange As Range) As String
  'classificeert een neerslagpatroon (duren 24, 48, 96, 126, 196 uur) als een van de 7 patronen die STOWA onderscheidt
  'in "Nieuwe Neerslagstatistiek voor Waterbeheerders"
  'ga ervan uit dat de gegevens als uurcijfers worden aangeleverd.
  'we delen de patronen op in drie delen. Aan de hand van de onderlinge verhoudingen bepalen we de classificatie
  
  Dim D(1 To 8)
  Dim Gegevens As New Collection
  Dim i As Long
  
  Dim r As Long
  For r = 1 To ValueRange.Rows.Count
    Gegevens.Add ValueRange.Cells(r, 1)
  Next
  
  Dim Sum As Double, Dmax As Double, PairMax As Double, QuatMax As Double
  Dim myPair As Double, myQuat As Double
  
  If Gegevens.Count = 24 Then
    Dmax = 0
    PairMax = 0
    QuatMax = 0
    'deel de periode op in drie vakken
    For i = 1 To Gegevens.Count / 8
      D(1) = D(1) + Gegevens(i)
    Next
    If D(1) > Dmax Then Dmax = D(1)
    
    For i = Gegevens.Count / 8 + 1 To Gegevens.Count * 2 / 8
      D(2) = D(2) + Gegevens(i)
    Next
    If D(2) > Dmax Then Dmax = D(2)
    
    For i = Gegevens.Count * 2 / 8 + 1 To Gegevens.Count * 3 / 8
      D(3) = D(3) + Gegevens(i)
    Next
    If D(3) > Dmax Then Dmax = D(3)
    
    For i = Gegevens.Count * 3 / 8 + 1 To Gegevens.Count * 4 / 8
      D(4) = D(4) + Gegevens(i)
    Next
    If D(4) > Dmax Then Dmax = D(4)
    
    For i = Gegevens.Count * 4 / 8 + 1 To Gegevens.Count * 5 / 8
      D(5) = D(5) + Gegevens(i)
    Next
    If D(5) > Dmax Then Dmax = D(5)
    
    For i = Gegevens.Count * 5 / 8 + 1 To Gegevens.Count * 6 / 8
      D(6) = D(6) + Gegevens(i)
    Next
    If D(6) > Dmax Then Dmax = D(6)
    
    For i = Gegevens.Count * 6 / 8 + 1 To Gegevens.Count * 7 / 8
      D(7) = D(7) + Gegevens(i)
    Next
    If D(7) > Dmax Then Dmax = D(7)
    
    For i = Gegevens.Count * 7 / 8 + 1 To Gegevens.Count
      D(8) = D(8) + Gegevens(i)
    Next
    If D(8) > Dmax Then Dmax = D(8)
    
    Sum = D(1) + D(2) + D(3) + D(4) + D(5) + D(6) + D(7) + D(8)
    
    'doorloop alle mogelijke paren
    For i = 1 To 7
      myPair = D(i) + D(i + 1)
      If myPair > PairMax Then PairMax = myPair
    Next
    
    'doorloop alle combinaties van 4
    For i = 1 To 5
      myQuat = D(i) + D(i + 1) + D(i + 2) + D(i + 3)
      If myQuat > QuatMax Then QuatMax = myQuat
    Next
    
    'type hoog
    If PairMax > 0.85 * Sum Then
      NeerslagPatroon = "hoog"
    ElseIf PairMax > 0.7 * Sum Then
      NeerslagPatroon = "middelhoog"
    ElseIf PairMax > 0.55 * Sum Then
      NeerslagPatroon = "middellaag"
    ElseIf QuatMax > 0.6 * Sum Then
      NeerslagPatroon = "laag"
    ElseIf D(2) > 0.25 * Sum And D(7) > 0.25 * Sum Then
      NeerslagPatroon = "kort"
    ElseIf D(1) > 0.25 * Sum And D(6) > 0.25 * Sum Then
      NeerslagPatroon = "kort"
    ElseIf D(3) > 0.25 * Sum And D(8) > 0.25 * Sum Then
      NeerslagPatroon = "kort"
    ElseIf D(1) > 0.25 * Sum And D(8) > 0.25 * Sum Then
      NeerslagPatroon = "lang"
    Else
      NeerslagPatroon = "uniform"
    End If
    
  Else
    MsgBox ("Error: alleen neerslagduur 24 uur in uurcijfers wordt geaccepteerd.")
    NeerslagPatroon = ""
  End If
  
  

End Function

Public Function GUMBELFIT(OK As Boolean, xVals As Range, x0Range As Range, betaRange As Range) As Boolean
  Dim maxiter As Integer
  Dim threshold As Double, mu As Double, Sigma As Double, beta_out As Double
  Dim dif As Double, olddif As Double, step As Double
  Dim i As Integer, dir As Integer, n As Integer
  Dim x0 As Double, beta As Double
        
  n = xVals.Rows.Count
  Dim x() As Double
  ReDim x(1 To xVals.Rows.Count)
  For i = 1 To xVals.Rows.Count
    x(i) = xVals.Cells(i, 1)
  Next
      
  threshold = 0.00001
  maxiter = 50
  mu = Application.WorksheetFunction.Average(xVals)
  Sigma = Application.WorksheetFunction.StDev(xVals)
        
  'first estimate of beta
  beta = Sigma / 1.2825
  
  'failsafe
  If (beta = 0) Then beta = 0.01

  'calc beta
  dif = 999
  dir = 1
  step = 0.5
  i = 0
  OK = True
  While (OK And (Math.Abs(dif) > threshold))
    i = i + 1
    OK = True
    Call calc_pars_Gumbel(OK, beta_out, x0, x, beta, mu)
    If (OK) Then
        olddif = dif
        dif = Abs(beta - beta_out)
        If (olddif < dif) Then
         dir = -dir
        End If
        If (dif >= threshold) Then
          beta = beta - (dif * step * dir)
        End If
        If (i > maxiter) Then
          OK = False
        End If
    End If
 
    If (beta = 0) Then OK = False
  Wend
  
  betaRange.Cells(1, 1) = beta_out
  x0Range.Cells(1, 1) = x0
  GUMBELFIT = True
  
End Function

Public Function GumbelFitMLE(xVals As Range, muCell As Range, sigmaCell As Range, likelihoodCell As Range, AkaikeCell As Range) As Boolean
    'this function fits a gumbel distribution by finding the maximum likelihood
    'it sets the cell value for mu, sigma and the maximum likelihood
    
    Dim muMax As Double, muMin As Double
    Dim sMax As Double, sMin As Double
    Dim iMu As Integer, iSigma As Integer
    Dim mu As Double, Sigma As Double
    Dim Likelihood As Double, bestLikelihood As Double
    Dim BestMuIdx As Integer, bestSigmaIdx As Integer
    Dim bestMu As Double, bestSigma As Double
    Dim iIter As Integer
    
    'initialize mu and sigma
    muMax = Application.WorksheetFunction.max(xVals)
    muMin = Application.WorksheetFunction.Min(xVals)
    sMax = muMax - muMin
    sMin = 0
    bestLikelihood = -9E+99
    
    For iIter = 1 To 50
        For iMu = 1 To 10
            mu = muMin + (muMax - muMin) / 10 * (iMu - 0.5) 'take the centerpoint from the current section
            For iSigma = 1 To 10
                Sigma = sMin + (sMax - sMin) / 10 * (iSigma - 0.5) 'take the centerpoint from the current selection
                Likelihood = GumbelLogLikelihood(xVals, mu, Sigma)
                'Likelihood = GumbelLikelihood(xVals, mu, sigma)
                If Likelihood > bestLikelihood Then
                    BestMuIdx = iMu
                    bestSigmaIdx = iSigma
                    bestMu = mu
                    bestSigma = Sigma
                    bestLikelihood = Likelihood
                End If
            Next
        Next
        
        'another iteration complete. Narrow down based on the best result
        muMin = muMin + (muMax - muMin) / 10 * (BestMuIdx - 1.5)
        muMax = muMin + (muMax - muMin) / 10 * (BestMuIdx + 1.5)
        sMin = sMin + (sMax - sMin) / 10 * (bestSigmaIdx - 1.5)
        sMax = sMin + (sMax - sMin) / 10 * (bestSigmaIdx + 1.5)
                
    Next
    
    muCell.Cells(1, 1) = bestMu
    sigmaCell.Cells(1, 1) = bestSigma
    likelihoodCell.Cells(1, 1) = bestLikelihood
    AkaikeCell.Cells(1, 1) = Akaike(bestLikelihood, 2)
    GumbelFitMLE = True
    
End Function

Public Function GumbelLogLikelihood(xVals As Range, mu As Double, Sigma As Double) As Double
    'this function computes the log likelihood for a given dataset and Gumbel distribution
    Dim myResult As Double
    Dim P As Double
    Dim r As Integer, z As Double
    For r = 1 To xVals.Rows.Count
        P = GUMBELPDF(mu, Sigma, xVals.Cells(r, 1))
        If P = 0 Then
          myResult = -9E+99
          Exit For
        Else
          myResult = myResult + Math.Log(P)
        End If
    Next
    GumbelLogLikelihood = myResult
End Function

Public Function GEVLogLikelihood(xVals As Range, mu As Double, Sigma As Double, Zeta As Double) As Double
    'this function computes the log likelihood for a given dataset and Gumbel distribution
    Dim myResult As Double
    Dim P As Double
    Dim r As Integer, z As Double
    For r = 1 To xVals.Rows.Count
        P = GEVPDF(mu, Sigma, Zeta, xVals.Cells(r, 1))
        If P = 0 Then
          myResult = -9E+99
          Exit For
        Else
          myResult = myResult + Math.Log(P)
        End If
    Next
    GEVLogLikelihood = myResult
End Function

Public Function GenParetoLogLikelihood(xVals As Range, mu As Double, Sigma As Double, kappa As Double) As Double
    'this function computes the log likelihood for a given dataset and Generalized Pareto distribution
    Dim myResult As Double
    Dim P As Double
    Dim r As Integer, z As Double
    For r = 1 To xVals.Rows.Count
        P = GENPARETOPDF(mu, Sigma, kappa, xVals.Cells(r, 1))
        If P = 0 Then
          myResult = -9E+99
          Exit For
        Else
          myResult = myResult + Math.Log(P)
        End If
    Next
    GenParetoLogLikelihood = myResult
End Function

Public Function Akaike(LogLikelihood As Double, nParameters As Integer) As Double
    Akaike = -2 * LogLikelihood + 2 * nParameters
End Function

Public Function Akaike2(Likelihood As Double, nParameters As Integer) As Double
    Akaike = -2 * Math.Log(Likelihood) + 2 * nParameters
End Function


Public Function calc_pars_Gumbel(OK As Boolean, BetaOut As Double, X0Out As Double, x() As Double, BetaIn As Double, muin As Double) As Boolean
Dim sum_xi_exp_m_xi_dbeta As Double
Dim sum_exp_m_xi_dbeta As Double
Dim i As Integer

OK = True

sum_xi_exp_m_xi_dbeta = 0
sum_exp_m_xi_dbeta = 0
For i = 1 To UBound(x)
    If (BetaIn = 0) Then
        BetaOut = 0.01
        Exit For
    End If
    sum_exp_m_xi_dbeta = sum_exp_m_xi_dbeta + kf_exp(-x(i) / BetaIn)
    sum_xi_exp_m_xi_dbeta = sum_xi_exp_m_xi_dbeta + x(i) * kf_exp(-x(i) / BetaIn)
    If (sum_exp_m_xi_dbeta = 0) Then
      OK = False
      Exit For
    End If
Next
BetaOut = muin - sum_xi_exp_m_xi_dbeta / sum_exp_m_xi_dbeta
X0Out = -BetaOut * Math.Log(sum_exp_m_xi_dbeta / UBound(x))
calc_pars_Gumbel = True

End Function

Public Function kf_exp(v As Double) As Double
    'een e-macht met beveiliging tegen crashen
    If (v > 700) Then v = 700
    If (v < -708) Then
        kf_exp = 0
    Else
        kf_exp = Math.Exp(1) ^ v
    End If
End Function

Public Function GUMBELPDF(mu As Double, Sigma As Double, x As Double) As Double
  '------------------------------------------------------------------------------------------------
  'Datum: 9-11-2010
  'Auteur: Siebe Bosch
  'Deze routine berekent de kansdichtheid volgens Gumbel
  'Kansdichtheidsfunctie: f(x) = e^-x * e^(-e^(-x))
  '------------------------------------------------------------------------------------------------
    
  Dim z As Double
  z = (x - mu) / Sigma
  GUMBELPDF = 1 / Sigma * Math.Exp(1) ^ (-(z + Math.Exp(1) ^ (-z)))
  
End Function

Public Function GUMBELCDF(mu As Double, Sigma As Double, x As Double) As Double
  '------------------------------------------------------------------------------------------------
  'Datum: 9-11-2010
  'Auteur: Siebe Bosch
  'Deze routine berekent de ONDERschrijdingskans van een bepaalde parameterwaarde volgens Gumbel type 1
  'dit betekent gewoon dat we de verdelingsfunctie gaan berekenen (= de integraal van de kansdichtheidsfunctie)
  'Kansdichtheidsfunctie: f(x) = e^-x * e^(-e^(-x))
  'Kansverdelingsfunctie: F(x) = e^(-e^((mu-x)/sigma))
  '------------------------------------------------------------------------------------------------
  
  Dim e As Double 'natuurlijke logaritme
  e = Math.Exp(1)
  
  GUMBELCDF = e ^ (-1 * e ^ ((mu - x) / Sigma))

End Function

Public Function CORRELATIONBYWINDOW(Range1 As Range, Range2 As Range, WindowSize As Integer) As Double
    Dim i As Integer, j As Integer
    Dim Sum1 As Double, Sum2 As Double
    Dim Array1 As Collection, Array2 As Collection
    Set Array1 = New Collection
    Set Array2 = New Collection
    If Range1.Rows.Count <> Range2.Rows.Count Then
        MsgBox ("Error: ranges should have equal dimensions.")
    Else
        For i = 1 To Range1.Rows.Count Step WindowSize
            Sum1 = 0
            Sum2 = 0
            For j = i To i + Minimum(Range1.Rows.Count, WindowSize - 1)
                Sum1 = Sum1 + Range1.Cells(j, 1)
                Sum2 = Sum2 + Range2.Cells(j, 1)
            Next
            Array1.Add (Sum1)
            Array2.Add (Sum2)
        Next
    End If
    
    'calculate the correlation coefficient
    CORRELATIONBYWINDOW = CorrelationCoef(Array1, Array2)
        
End Function

Public Function CorrelationCoef(Coll1 As Collection, Coll2 As Collection) As Double
    
    'start by finding the average value for both collections
    Dim Avg1 As Double, Avg2 As Double
    Dim i As Integer
    For i = 1 To Coll1.Count
        Avg1 = Avg1 + Coll1(i)
        Avg2 = Avg2 + Coll2(i)
    Next
    Avg1 = Avg1 / Coll1.Count
    Avg2 = Avg2 / Coll2.Count
    
    'now compute the correlation coefficient
    Dim Teller As Double, Noemer1 As Double, Noemer2 As Double, Noemer As Double
    For i = 1 To Coll1.Count
        Teller = Teller + (Coll1.Item(i) - Avg1) * (Coll2.Item(i) - Avg2)
        Noemer1 = Noemer1 + (Coll1.Item(i) - Avg1) ^ 2
        Noemer2 = Noemer2 + (Coll2.Item(i) - Avg2) ^ 2
    Next
    Noemer = Math.Sqr(Noemer1 * Noemer2)
    CorrelationCoef = Teller / Noemer
    
End Function

Public Function StdevAvgRatio(myRange As Range) As Double
    'computes StDev / Average for a given range.
    StdevAvgRatio = Application.WorksheetFunction.StDev(myRange) / Application.WorksheetFunction.Average(myRange)
End Function


Public Function EXPONENTIELEVERDELINGCDF(lambda As Double, x As Double, threshold As Double) As Double
    If x > threshold Then
        EXPONENTIELEVERDELINGCDF = 1 - Math.Exp(1) ^ (-lambda * (x - threshold))
    Else
        EXPONENTIELEVERDELINGCDF = 0
    End If
End Function

Public Function GUMBELINVERSE(P As Double, mu As Double, Sigma As Double) As Double
  '------------------------------------------------------------------------------------------------
  'Datum: 9-11-2010
  'Auteur: Siebe Bosch
  'Deze routine berekent de waarde X gegeven de ONDERschrijdingskans p volgens Gumbel type 1
  'dit betekent gewoon dat we de verdelingsfunctie gaan berekenen (= de integraal van de kansdichtheidsfunctie)
  'Kansdichtheidsfunctie: f(x) = e^-x * e^(-e^(-x))
  'Kansverdelingsfunctie: F(x) = e^(-e^((mu-x)/sigma))
  '------------------------------------------------------------------------------------------------
    
    GUMBELINVERSE = mu - Sigma * (Math.Log(-1 * Math.Log(P)))

End Function

Public Function GENPARETOINVERSE(p_ond As Double, mu As Double, Sigma As Double, kappa As Double) As Double
  '------------------------------------------------------------------------------------------------
  'Datum: 25-7-2018
  'Auteur: Siebe Bosch
  'Deze routine berekent de waarde X gegeven de ONDERschrijdingskans p volgens Generalized Pareto
  'Cumulatieve kansdichtheidsfunctie: F(x) = 1-(1+kz)^-1/k waarin:
  'z = (x-mu)/sigma
  'LET OP: het is niet gelukt om deze formule te inverteren, dus we lossen het iteratief op
  '------------------------------------------------------------------------------------------------
  
   'we weten op welke onderschrijdingskans we willen uitkomen en zoeken daarbij de X
   'laten we X iteratief zoeken tussen mu - 10sigma en mu + 10sigma
   Dim iIter As Integer, iSlice As Integer
   Dim Xmin As Double, Xmax As Double, Slice As Double
   Dim Xcur As Double, Pcur As Double, Pbest As Double, BestSlice As Integer, Xbest As Double
   
   Xmin = mu - 10 * Sigma
   Xmax = mu + 10 * Sigma
   Pbest = -1
   
   For iIter = 1 To 10
     'split the range in ten slices
     Slice = (Xmax - Xmin) / 10
     For iSlice = 1 To 10
        Xcur = Xmin + (iSlice - 0.5) * Slice
        Pcur = GENPARETOCDF(mu, Sigma, kappa, Xcur)
        If Math.Abs(Pcur - p_ond) < Math.Abs(Pbest - p_ond) Then
            Pbest = Pcur
            Xbest = Xcur
            BestSlice = iSlice
        End If
     Next
     
     'narrow down the search window and move on to the next iteration
     'build in some extra security by surrounding slices in the next iteration
     Xmin = Xbest - 2 * Slice
     Xmax = Xbest + 2 * Slice
     
   Next
   
    GENPARETOINVERSE = Xbest
  
End Function

Public Function Log10(myVal As Double) As Double
    'omdat VBA met LOG de natuurlijke logaritme bedoelt, voeren we hier een conversie uit naar 10Log
    Log10 = Log(myVal) / Log(10)
End Function



Public Function GLOCDF(mu As Double, Sigma As Double, teta As Double, x As Double) As Double
  '------------------------------------------------------------------------------------------------
  'Datum: 5-11-2018
  'Auteur: Siebe Bosch
  'Deze routine berekent de ONDERschrijdingskans van een bepaalde parameterwaarde volgens de GLO-verdeling (Generalized Logistic)
  'dit betekent gewoon dat we de verdelingsfunctie gaan berekenen (= de integraal van de kansdichtheidsfunctie)
  '------------------------------------------------------------------------------------------------
  
  Dim z As Double, T As Double
  z = (x - mu) / Sigma
    
  If teta = 0 Then
    GLOCDF = (1 + Exp(-z)) ^ -1
  Else
    GLOCDF = (1 + (1 - teta * z) ^ (1 / teta)) ^ -1
  End If
    
End Function

Public Function GLOINVERSE(mu As Double, Sigma As Double, teta As Double, Value As Double) As Double
  '------------------------------------------------------------------------------------------------
  'Datum: 24-12-2018
  'Auteur: Siebe Bosch
  'Deze routine berekent de waarde X gegeven een ONDERschrijdingskans en een GLO-kansverdeling (Generalized Logistic)
  'dit betekent gewoon dat we de verdelingsfunctie gaan berekenen (= de integraal van de kansdichtheidsfunctie)
  '------------------------------------------------------------------------------------------------
  
  If teta = 0 Then
    GLOINVERSE = mu - Sigma * Math.Log(1 / Value - 1)
  Else
    GLOINVERSE = mu + Sigma * ((1 - (1 / Value - 1) ^ teta) / teta)
  End If
  
End Function


Public Function GEVPDF(mu As Double, Sigma As Double, Zeta As Double, x As Double) As Double
  Dim e As Double 'natuurlijke logaritme
  Dim z As Double, T As Double
  
  e = Math.Exp(1)
  z = (x - mu) / Sigma
    
  If Zeta = 0 Then
    T = e ^ -z
  Else
    T = (1 + Zeta * z) ^ (-1 / Zeta)
  End If
  
  GEVPDF = 1 / Sigma * T ^ (Zeta + 1) * e ^ -T
  
End Function


Public Function GEVCDFOLD(mu As Double, Sigma As Double, K As Double, x As Double) As Double
  'calculates the cumulative probability density according to the GEV-probability distribution

   Dim z As Double
   'Dim arg1 As Double
   'Dim arg2 As Double
   
   z = (x - mu) / Sigma
   'arg1 = (1 + k * z)
   'arg2 = -1 / k
   
   If K <> 0 Then
     GEVCDF = Exp(-1 * (1 + K * z) ^ (-1 / K)) 'this is the original one
     'GEVCDF = Exp(-1 * arg1 ^ arg2)      'edit: this was necessary to prevent an invalid procedure call due to the numbers inside
   Else
     GEVCDF = Exp(-1 * Math.Exp(-z))
   End If
   
End Function

Public Function GEVCDF(mu As Double, Sigma As Double, Zeta As Double, x As Double) As Double
  '------------------------------------------------------------------------------------------------
  'Datum: 9-11-2010
  'Datum: 21-3-2020 het minteken van Zeta omgekeerd. Nu consistent met de STOWA-notatie
  'Auteur: Siebe Bosch
  'Deze routine berekent de ONDERschrijdingskans van een bepaalde parameterwaarde volgens de GEV-verdeling (Gegeneraliseerde Extreme Waarden)
  'dit betekent gewoon dat we de verdelingsfunctie gaan berekenen (= de integraal van de kansdichtheidsfunctie)
  'Kansverdelingsfunctie:    F(x;\mu,\sigma,\xi) = \exp\left\{-\left[1+\xi\left(\frac{x-\mu}{\sigma}\right)\right]^{-1/\xi}\right\}
  '------------------------------------------------------------------------------------------------
  
  Dim e As Double 'natuurlijke logaritme
  Dim z As Double, T As Double
  
  e = Math.Exp(1)
  z = (x - mu) / Sigma
    
  If Zeta = 0 Then
    T = e ^ -z
  Else
    T = (1 - Zeta * z) ^ (1 / Zeta)
  End If
  
  GEVCDF = e ^ -T
  
End Function

Public Function GEVINVERSE(mu As Double, Sigma As Double, Zeta As Double, Value As Double) As Double
  '------------------------------------------------------------------------------------------------
  'Datum: 9-11-2010
  'Auteur: Siebe Bosch
  'Deze routine berekent de ONDERschrijdingskans p van een bepaalde parameterwaarde volgens GEV-verdeling
  'dit betekent gewoon dat we de verdelingsfunctie gaan berekenen (= de integraal van de kansdichtheidsfunctie)
  '------------------------------------------------------------------------------------------------

  GEVINVERSE = mu + Sigma * (((-1 * Application.WorksheetFunction.Ln(Value)) ^ (Zeta) - 1) / -Zeta)

End Function

Public Function EXPPDF(lambda As Double, y As Double, Value As Double) As Double
    EXPPDF = lambda * Math.Exp(-lambda * (Value - y))
End Function

Public Function EXPCDF(lambda As Double, y As Double, Value As Double) As Double
    EXPCDF = 1 - Math.Exp(-lambda * (Value - y))
End Function

Public Function EXPINVERSE(lambda As Double, y As Double, p_ond As Double) As Double
    EXPINVERSE = -(Math.Log(1 - p_ond)) / lambda + y
End Function

  Public Sub calcNeerslagStats(ByVal duration As Integer, ByVal Area As Double, ByRef mu As Double, ByRef gamma As Double, ByRef kappa As Double)
    'deze functie berekent de statistische parameters van de kansdichtheidsfunctie voor neerslagvolume in Nederland:
    'neerslagvolume voldoet namelijk aan de GEV-kansverdeling (Gegeneraliseerde Extremewaardenverdeling)
    'mu = locatieparameter' gamma = schaalparameter, kappa = vormparameter
    'waarden voor a1, a2, b1, b2 en c zijn aangeleverd door HKV-lijn in water
    'document Actuele extreme neerslagstatistiek en neerslag- en verdampingsreeksen, van 7 juli 2011: PR2197.10
    'originele bronvermelding: Overeem, A., T.A. Buishand, I. Holleman en R. Uijlenhoet, Extreme-value modeling of areal rainfall from weather radar, Water Resour. Res., 2010, 46, W09514, doi:10.1029/2009wr008517
    Dim y As Double
    Dim a1 As Double, a2 As Double, b1 As Double, b2 As Double, c As Double

    a1 = 17.92  'was in 2009 17.92
    a2 = 0.225  'was in 2009 0.225
    b1 = -3.57  'was in 2009 -3.57
    b2 = 0.427  'was in 2009 0.43
    c = 0.128   'was in 2009 0.128
    mu = a1 * duration ^ a2 + b1 * Area ^ c + (b2 * Area ^ c) * Math.Log(duration)

    a1 = 0.337  'was in 2009 0.344
    a2 = -0.018 'was in 2009 -0.025
    b1 = -0.014 'was in 2009 -0.016
    b2 = 0      'was in 2009 0.0003
    c = 0       'was in 2009 0
    y = a1 + a2 * Math.Log(duration) + b1 * Math.Log(Area) + b2 * duration * Math.Log(Area)
    gamma = y * mu

    a1 = -0.206 'was in 2009 -0.206
    a2 = 0      'was in 2009 0
    b1 = 0.018  'was in 2009 0.022
    b2 = 0      'was in 2009 -0.004
    c = 0       'was in 2009 0
    kappa = a1 + b1 * Math.Log(Area) + b2 * Math.Log(duration) * Math.Log(Area)

  End Sub


Public Function calcHerhalingstijd(ByVal Volume As Double, ByVal Duur As Integer, ByVal Area As Double) As Double
    'berekent de herhalingstijd van een bui, gegeven Volume, Duur en gebiedsoppervlak)
    
    'Volume in mm
    'Duur in uren
    'Area in km2
    
    Dim mu As Double, gamma As Double, kappa As Double
    Dim F_jaar As Double 'overschrijdingsfrequentie op jaarbasis
    
    Call calcNeerslagStats(Duur, Area, mu, gamma, kappa)        'bereken de kansdichtheidsparameters
    F_jaar = (1 - kappa / gamma * (Volume - mu)) ^ (1 / kappa)  'frequentie in aantal keren / jaar
    calcHerhalingstijd = 1 / F_jaar                             'bereken de herhalingstijd van de gebeurtenis

    'onderstaand is een test of de terugrekening weer hetzelfde volume genereert
    'Dim myVol = calcNeerslagVolume(Area, Duration, ARI, mu, gamma, kappa)

  End Function
  
  Public Function calcNeerslagVolume(ByVal Area As Double, ByVal Duur As Integer, ByVal Herhalingstijd As Double) As Double
    'Deze functie rekent terug. Gegeven duur, Oppervlak en overschrijdingskans
    'rekent hij het volume over een oppervlak groter dan puntneerslag uit
    Dim F_jaar As Double
    Dim mu As Double, gamma As Double, kappa As Double
    Call calcNeerslagStats(Duur, Area, mu, gamma, kappa)        'bereken de kansdichtheidsparameters
    
    F_jaar = 1 / Herhalingstijd
    calcNeerslagVolume = mu + gamma / kappa * (1 - F_jaar ^ kappa)

  End Function
  
  Public Sub PrecipitationAreaReduction(ValuesRange As Range, CorrRange As Range, ActivityRange As Range, ProgressRange As Range, ByVal minHerh As Single, Optional ByVal Area As Double = 6)
    'deze routine identificeert individuele buien uit een tijdreeks met uurlijkse neerslagsommen in Nederland
    'Oppervlak in km2
    'minHerh = minimum Herhalingstijd in jaren
    'voor puntneerslag houden we een standaardoppervlakte van 6 km2 aan
    Dim i As Integer, r As Long, K As Long, Dur As Integer, mySum As Double, myNextSum As Double, h As Single
    Dim SkipEvent As Boolean, HERH() As Double, EventSum() As Double, duration() As Integer
    Dim myMu As Double, myGamma As Double, myKappa As Double 'probability function parameters
    Dim subRange As Range, CorrSum As Double
    
    'opschonen bestaand resultaat en herdimensioneren arrays
    Call CorrRange.ClearContents
    ReDim HERH(1 To ValuesRange.Rows.Count)
    ReDim EventSum(1 To ValuesRange.Rows.Count)
    ReDim duration(1 To ValuesRange.Rows.Count)

    'doorloop alle neerslagduren van 1, 2, 4, 8, 12 en 24 uur
    For i = 1 To 6
      Select Case i
        Case Is = 1
          Dur = 1
        Case Is = 2
          Dur = 2
        Case Is = 3
          Dur = 4
        Case Is = 4
          Dur = 8
        Case Is = 5
          Dur = 12
        Case Is = 6
          Dur = 24
      End Select
    
      ActivityRange.Cells(1, 1) = "Analyseren neerslagduur " & Dur & " uur."
      DoEvents

      'doorloop de gecorrigeerde neerslagwaarden en onderscheid buien hierbinnen
      For r = 1 To ValuesRange.Rows.Count - 1
            
        mySum = Application.WorksheetFunction.Sum(Range(ValuesRange.Cells(r, 1), ValuesRange.Cells(r + Dur - 1, 1)))
        myNextSum = Application.WorksheetFunction.Sum(Range(ValuesRange.Cells(r + 1, 1), ValuesRange.Cells(r + Dur, 1)))

        If myNextSum < mySum Then 'nu weten we dat we een losse bui te pakken hebben
          'bereken de overschrijdingskans van deze puntneerslagsom en haal on the fly ook de bijbehorende Herhalingstijd binnen
          ProgressRange.Cells(1, 1) = r / ValuesRange.Rows.Count
          DoEvents
          
          h = calcHerhalingstijd(mySum, Dur, 6)

          'alleen als de herhalingstijd > minimum is, schrijven we hem weg
          If h >= minHerh Then

            'doorloop eerst de lijst met herhalingstijden om te checken of hij al is toegekend
            SkipEvent = False 'initialiseer SkipEvent
            For K = r To r + Dur - 1
              If HERH(K) > h Then
                'helaas, een gebeurtenis met kortere duur had al een grotere herhalingstijd. We skippen deze bui voor de huidige duur
                SkipEvent = True
                Exit For
              End If
            Next

            'als deze gebeurtenis nog niet is overruled door een zeldzamer herhalingstijd bij kortere duur:
            'leg de herhalingstijd vast!
            If Not SkipEvent Then
              For K = r To r + Dur - 1
                HERH(K) = h                'leg voor deze bui de herhalingstijd vast
                duration(K) = Dur         'leg voor deze bui de neerslagduur vast
                EventSum(K) = mySum       'leg voor deze bui de neerslagsom vast
              Next
            End If
            'Bui is afgehandeld, dus zet r aan het einde van de bui
            r = r + Dur - 1
          End If
        End If
      Next
      
    Next
    
    'update de voortgangsindicatoren
    ProgressRange.Cells(1, 1) = 0
    ActivityRange.Cells(1, 1) = "Berekent gecorrigeerde neerslagvolumes."
    DoEvents
    
    'doorloop nu alle cellen om de gecorrigeerde neerslagvolumes te berekenen en weg te schrijven
    For K = 1 To ValuesRange.Rows.Count
      If HERH(K) > 1 Then
        ProgressRange.Cells(1, 1) = K / ValuesRange.Rows.Count
        CorrSum = calcNeerslagVolume(Area, duration(K), HERH(K))
        CorrRange.Cells(K, 1) = ValuesRange.Cells(K, 1) * CorrSum / EventSum(K)
        DoEvents
      Else
        'geen correctie; neem oorspronkelijke waarde over
        CorrRange.Cells(K, 1) = ValuesRange.Cells(K, 1)
      End If
    Next
    
    'update de voortgangsindicatoren
    ProgressRange.Cells(1, 1) = 100
    ActivityRange.Cells(1, 1) = "Klaar."
    DoEvents
    

  End Sub

Public Sub ANNUALMAXIMUMPRECIPITATIONEVENTS(HeaderRow As Integer, DateCol As Integer, ValCol As Integer, duration As Integer, resultsRow As Integer, resultsCol As Integer, ProgressRange As Range)
    'Deze subroutine loopt door een volledige tijdreeks met neerslagvolumes en extraheert de maxima per jaar en seizoen
    
    Dim ValSubRange As Range
    Dim DateSubRange As Range
    Dim DateValRange As Range
    Dim i As Long, r As Long
    Dim myDate As Date, myYear As Integer, mySeizoen As String
    Dim mySum As Double
    Dim MergeCells As Range
    
    Dim StartYear As Integer
    Dim EndYear As Integer
    
    'set de range
    r = HeaderRow
    While Not ActiveSheet.Cells(r + 1, DateCol) = ""
      r = r + 1
    Wend
    Set DateValRange = Range(ActiveSheet.Cells(HeaderRow + 1, DateCol), ActiveSheet.Cells(r, ValCol))
    
    StartYear = Year(DateValRange.Cells(1, DateCol))
    EndYear = Year(DateValRange.Cells(DateValRange.Rows.Count, DateCol))

    Dim JaarMaximaZOM() As Double
    Dim JaarMaximaWin() As Double
    Dim JaarMaxima() As Double
    ReDim JaarMaximaZOM(StartYear To EndYear)
    ReDim JaarMaximaWin(StartYear To EndYear)
    ReDim JaarMaxima(StartYear To EndYear)
    
    For i = 1 To DateValRange.Rows.Count - duration + 1
      Set ValSubRange = DateValRange.Range(DateValRange.Cells(i, 2), DateValRange.Cells(i + duration - 1, 2))
      Set DateSubRange = DateValRange.Range(DateValRange.Cells(i, 1), DateValRange.Cells(i + duration - 1, 1))
      myDate = DateValRange.Cells(i, 1)
      myYear = Year(myDate)
      mySeizoen = METEOROLOGISCHHALFJAAR(myDate)
      mySum = Application.WorksheetFunction.Sum(ValSubRange)
      If mySum > JaarMaxima(myYear) Then
        JaarMaxima(myYear) = mySum
        ProgressRange = i / DateValRange.Rows.Count
        DoEvents
      End If
      If VBA.LCase(mySeizoen) = "zomer" Then
        If mySum > JaarMaximaZOM(myYear) Then JaarMaximaZOM(myYear) = mySum
      ElseIf VBA.LCase(mySeizoen) = "winter" Then
        If mySum > JaarMaximaWin(myYear) Then JaarMaximaWin(myYear) = mySum
      End If
    Next
        
    'create a section header and merge the cells
    r = resultsRow
    Set MergeCells = Range(Cells(r, resultsCol), Cells(r, resultsCol + 3))
    MergeCells.Merge
    ActiveSheet.Cells(r, resultsCol) = duration & "h"
    
    'write the column headers
    r = r + 1
    ActiveSheet.Cells(r, resultsCol) = "jaar"
    ActiveSheet.Cells(r, resultsCol + 1) = "jaarrond"
    ActiveSheet.Cells(r, resultsCol + 2) = "zomer"
    ActiveSheet.Cells(r, resultsCol + 3) = "winter"
    
    'write the results
    For i = StartYear To EndYear
      If JaarMaxima(i) > 0 Then
        r = r + 1
        ActiveSheet.Cells(r, resultsCol) = i
        ActiveSheet.Cells(r, resultsCol + 1) = JaarMaxima(i)
        ActiveSheet.Cells(r, resultsCol + 2) = JaarMaximaZOM(i)
        ActiveSheet.Cells(r, resultsCol + 3) = JaarMaximaWin(i)
      End If
    Next

End Sub

Public Function PLOTTINGPOSITIONFROMANNUALMAXIMA(myVal As Double, ValuesRange As Range) As Double
   Dim r As Long, n As Long, i As Long, F As Double, h As Double, P As Double
   n = ValuesRange.Rows.Count
   Dim curVal As Double
   
   'writes the return period in the second column of the range
  If ValuesRange.Columns.Count <> 1 Then
    MsgBox ("Range must contain only one column, containing the annual maxima.")
  Else
  
    'calculate the index number for the given value
    i = 0
    For r = 1 To ValuesRange.Rows.Count
      curVal = ValuesRange.Cells(r, 1)
      If curVal >= myVal Then i = i + 1
    Next
       
    'calculate the return period based on the index number
    P = (i - 0.3) / (n + 0.4) 'plotting position
    F = -Math.Log(1 - P) 'exceedance frequency in times per year
    PLOTTINGPOSITIONFROMANNUALMAXIMA = 1 / F 'return period
   
   End If
   
End Function

Public Sub IDENTIFYPRECIPITATIONEVENTSPOT(DateTimeCol As Long, ValueCol As Long, StartRow As Long, EndRow As Long, duration As Integer, POT As Double, resultsRow As Integer, resultsCol As Integer, ProgressRange As Range)
    'Deze subroutine loopt door een volledige tijdreeks met neerslagvolumes en de totaalvolumes die een bepaalde POT-waarde overschrijden
    
    Dim i As Long, j As Long, r As Long, c As Long
    Dim myYear As Integer, mySeizoen As String
    
    Dim PrevRange As Range, CurRange As Range, NextRange As Range
    Dim PrevSum As Double, CurSum As Double, NextSum As Double
    Dim Zomer As Collection, Winter As Collection, Jaarrond As Collection
    Dim myDate As Date
    
    Set Zomer = New Collection
    Set Winter = New Collection
    Set Jaarrond = New Collection
    
    For r = StartRow + 1 To EndRow - duration - 2
    
      ProgressRange.Cells(1, 1) = (r - StartRow) / (EndRow - StartRow)
      DoEvents
    
      Set PrevRange = ActiveSheet.Range(Cells(r - 1, ValueCol), Cells(r + duration - 2, ValueCol))
      Set CurRange = ActiveSheet.Range(Cells(r, ValueCol), Cells(r + duration - 1, ValueCol))
      Set NextRange = ActiveSheet.Range(Cells(r + 1, ValueCol), Cells(r + duration, ValueCol))
      PrevSum = WorksheetFunction.Sum(PrevRange)
      CurSum = WorksheetFunction.Sum(CurRange)
      NextSum = WorksheetFunction.Sum(NextRange)
      
      If CurSum > PrevSum And CurSum > NextSum And CurSum > POT Then
        myDate = ActiveSheet.Cells(r, DateTimeCol)
        myYear = Year(myDate)
        mySeizoen = METEOROLOGISCHHALFJAAR(myDate)
        r = r + duration - 1                            'skip deze bui nu we hem geidentificeerd hebben
        If mySeizoen = "zomer" Then
          Call Zomer.Add(CurSum, Str(myDate))
          Call Jaarrond.Add(CurSum, Str(myDate))
        ElseIf mySeizoen = "winter" Then
          Call Winter.Add(CurSum, Str(myDate))
          Call Jaarrond.Add(CurSum, Str(myDate))
        End If
      End If
    Next
    
    r = resultsRow
    c = resultsCol
    ActiveSheet.Cells(r, c) = "Zomer"
    For i = 1 To Zomer.Count
      r = r + 1
      ActiveSheet.Cells(r, c) = Zomer.Item(i)
    Next
        
    r = resultsRow
    c = c + 1
    ActiveSheet.Cells(r, c) = "Winter"
    For i = 1 To Winter.Count
      r = r + 1
      ActiveSheet.Cells(r, c) = Winter.Item(i)
    Next
    
    r = resultsRow
    c = c + 1
    ActiveSheet.Cells(r, c) = "Jaarrond"
    For i = 1 To Jaarrond.Count
      r = r + 1
      ActiveSheet.Cells(r, c) = Jaarrond.Item(i)
    Next
    
End Sub

Public Sub CLASSIFYEVENTS(myRange As Range, duration As Integer, ContainsHeader As Boolean, ClassMin As Double, ClassMax As Double)
  'verplicht: 1e kolom = datum, 2e kolom = waarde, 3e kolom = resultaat
  Dim r As Long, i As Long, Done As Boolean
  Dim ValRange As Range, resRange As Range, myCell As Range
  Dim Sum As Double, maxSum As Double, maxIdx As Long
  Dim RankNum As Long, Ranks() As Long
  ReDim Ranks(1 To duration)
  
  Dim StartRow As Integer
  If ContainsHeader Then
    StartRow = 2
  Else
    StartRow = 1
  End If
  
  'remove old results
  Set resRange = myRange.Range(myRange.Cells(StartRow, 3), myRange.Cells(myRange.Rows.Count, 3))
  Call resRange.ClearContents
  
  While Not Done
    Done = True
    Sum = 0
    maxSum = 0
    For r = StartRow To myRange.Rows.Count - duration
      Set ValRange = myRange.Range(myRange.Cells(r, 2), myRange.Cells(r + duration - 1, 2))
      Set resRange = myRange.Range(myRange.Cells(r, 3), myRange.Cells(r + duration - 1, 3))
      
      If Application.WorksheetFunction.Sum(resRange) = 0 And Application.WorksheetFunction.Sum(ValRange) >= ClassMin And Application.WorksheetFunction.Sum(ValRange) <= ClassMax Then
        Sum = Application.WorksheetFunction.Sum(ValRange)
        If Sum > maxSum Then
          maxIdx = r
          maxSum = Sum
          Done = False
        End If
      End If
    Next
    
    If Done = False Then
      RankNum = RankNum + 1
      Set resRange = myRange.Range(myRange.Cells(maxIdx, 3), myRange.Cells(maxIdx + duration - 1, 3))
      For i = 1 To duration
        Ranks(i) = RankNum
      Next
      resRange.Value = Ranks
    End If
    
  Wend
End Sub

Public Sub MAXFROMMOVINGWINDOW(DATARANGEWITHCOLHEADER As Range, Durations As Collection, ResultsSheet As String, resultsRow As Integer, resultsCol As Integer)
    
    Dim curSheet As Worksheet, newSheet As Worksheet
    Dim CurSheetName As String
    CurSheetName = ActiveSheet.Name
    Set curSheet = ActiveWorkbook.Sheets(CurSheetName)
        
    'first create a new worksheet for the results
    If Not WorkSheetExists(ResultsSheet) Then
        Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = ResultsSheet
        Set newSheet = ActiveWorkbook.Sheets(ResultsSheet)
    Else
        MsgBox ("Worksheet " & ResultsSheet & " already exists. Please remove the old one first.")
        Exit Sub
    End If
    
    Dim iDur As Integer, myDuration As Integer
    Dim r As Long, c As Long, i As Integer, r2 As Integer, c2 As Integer
    Dim ID As String
    Dim myMax As Double, mySum As Double
    r2 = resultsRow
    
    For iDur = 1 To Durations.Count
        r2 = r2 + 1
        c2 = resultsCol
        myDuration = Durations(iDur)
        newSheet.Cells(r2, resultsCol) = myDuration                          'write the duration to the results sheet
        For c = 1 To DATARANGEWITHCOLHEADER.Columns.Count
            myMax = 0
            c2 = c2 + 1
            ID = DATARANGEWITHCOLHEADER.Cells(1, c)
            newSheet.Cells(resultsRow, c2) = ID                      'write the id to the results sheet
            For r = 2 To DATARANGEWITHCOLHEADER.Rows.Count
                mySum = 0
                For i = 0 To myDuration - 1
                    mySum = mySum + DATARANGEWITHCOLHEADER.Cells(r + i, c)
                Next
                If mySum > myMax Then myMax = mySum
             Next
             newSheet.Cells(r2, c2) = myMax
        Next
    Next
    
    MsgBox ("Operation complete.")
            
End Sub


Public Sub RANKNUMBEROFEXCEEDANCESBYMOVINGWINDOW(StartRowInclHeader As Integer, DataCol As Integer, RankCol As Integer, resultsCol As Integer, MovingWindowSize As Integer, threshold As Double, ProgressRange As Range, Optional ByVal OnlyIfSequential As Boolean = False)
  'This routine finds the number of exceedances of a given threshold in a given window size and
  'classifies them, using a moving window.
  Dim r As Long, c As Long, EndRow As Long
  Dim ValRange As Range, RankRange As Range, TempRange As Range
  Dim maxFound As Integer, maxRow As Integer
  Dim maxExceedances As Integer, RankNum As Integer
  
  'find the last row
  EndRow = StartRowInclHeader
  While Not ActiveSheet.Cells(EndRow + MovingWindowSize, DataCol) = ""
    EndRow = EndRow + 1
  Wend
  
  'cleanup the results range
  Set TempRange = ActiveSheet.Range(Cells(StartRowInclHeader + 1, RankCol), Cells(EndRow + MovingWindowSize - 1, RankCol))
  Call TempRange.ClearContents
  Set TempRange = ActiveSheet.Range(Cells(StartRowInclHeader + 1, resultsCol), Cells(EndRow + MovingWindowSize - 1, resultsCol))
  Call TempRange.ClearContents
  
  'write the results header
  ActiveSheet.Cells(StartRowInclHeader, RankCol) = "Rangnummer"
  ActiveSheet.Cells(StartRowInclHeader, resultsCol) = "Aantal aaneengesloten overschrijdingen"
  
  'start searching for threshold exceedances, using the moving window. Start with the largest numer of exceedances, end with the lowest ones
  maxFound = 999 'initialization
  While maxFound > 0
    
    maxFound = 0
    
    'find the window with the highest number of exceedances
    For r = StartRowInclHeader + 1 To EndRow
      Set RankRange = Range(ActiveSheet.Cells(r, RankCol), ActiveSheet.Cells(r + MovingWindowSize - 1, RankCol))
      
      If Application.WorksheetFunction.Sum(RankRange) = 0 Then
        Set ValRange = Range(ActiveSheet.Cells(r, DataCol), ActiveSheet.Cells(r + MovingWindowSize - 1, DataCol))
        'count the number of exceedances of the threshold
        If OnlyIfSequential Then
          maxExceedances = CountSequentialExceedances(ValRange, threshold)
        Else
          maxExceedances = Application.WorksheetFunction.CountIf(ValRange, "> " & threshold)
        End If
              
        'if the number exceeds the previously found number we'll overwrite the previous value
        If maxExceedances > maxFound Then
          maxFound = maxExceedances
          maxRow = r
        End If
      End If
    Next
    
    'write the ranking number found to the results column
    If maxFound > 0 Then
      RankNum = RankNum + 1
      
      Set TempRange = ActiveSheet.Range(Cells(maxRow, RankCol), Cells(maxRow + MovingWindowSize - 1, RankCol))
      TempRange.Value = RankNum
      Set TempRange = ActiveSheet.Range(Cells(maxRow, resultsCol), Cells(maxRow + MovingWindowSize - 1, resultsCol))
      TempRange.Value = maxFound
    End If
    
    'update the progress indicator
    ProgressRange.Value = RankNum & " of maximum " & (EndRow - StartRowInclHeader) / MovingWindowSize
    DoEvents
  
  Wend
    
End Sub
    
Public Sub RANKNUMBEROFEXCEEDANCESBYINTERVAL(StartRowInclHeader As Integer, DataCol As Integer, RankCol As Integer, resultsCol As Integer, IntervalSize As Integer, threshold As Double, Optional ByVal OnlyIfSequential As Boolean = False)
  'This routine finds the number of exceedances of a given threshold in a given window size and
  'classifies them, using a fixed interval
  Dim r As Long, c As Long, StartRow As Long, EndRow As Long
  Dim ValRange As Range, RankRange As Range, TempRange As Range
  Dim maxFound As Integer, maxRow As Integer
  Dim nExceedances As Integer, blockNum As Integer
  
  'find the last row
  EndRow = StartRowInclHeader
  StartRow = StartRowInclHeader + 1
  While Not ActiveSheet.Cells(EndRow + IntervalSize, DataCol) = ""
    EndRow = EndRow + 1
  Wend
  
  'cleanup the results range
  Set TempRange = ActiveSheet.Range(Cells(StartRowInclHeader + 1, RankCol), Cells(EndRow + IntervalSize - 1, RankCol))
  Call TempRange.ClearContents
  Set TempRange = ActiveSheet.Range(Cells(StartRowInclHeader + 1, resultsCol), Cells(EndRow + IntervalSize - 1, resultsCol))
  Call TempRange.ClearContents
  
  'write the results header
  ActiveSheet.Cells(StartRowInclHeader, RankCol) = "Rangnummer"
  ActiveSheet.Cells(StartRowInclHeader, resultsCol) = "Aantal aaneengesloten overschrijdingen"
  
  For r = StartRow To EndRow Step IntervalSize
    blockNum = blockNum + 1
    Set ValRange = ActiveSheet.Range(Cells(r, DataCol), Cells(r + IntervalSize - 1, DataCol))
    
    If OnlyIfSequential Then
      nExceedances = CountSequentialExceedances(ValRange, threshold)
    Else
      nExceedances = Application.WorksheetFunction.CountIf(ValRange, "> " & threshold)
    End If
    
    Set TempRange = ActiveSheet.Range(Cells(r, RankCol), Cells(r + IntervalSize - 1, RankCol))
    TempRange.Value = blockNum
    Set TempRange = ActiveSheet.Range(Cells(r, resultsCol), Cells(r + IntervalSize - 1, resultsCol))
    TempRange.Value = nExceedances
  Next
        
End Sub

Public Sub POTANALYSISSUM(HeaderRow As Integer, DateCol As Integer, ValCol As Integer, EventIndexCol As Integer, EventSumCol As Integer, EventExceedanceCol As Integer, duration As Integer, MinimumTimeStepsBetweenEvents As Integer, PotExceedanceFrequencyPerYear As Integer, IncludeSummer As Boolean, IncludeWinter As Boolean, ProgressRange As Range)
  '---------------------------------------------------------------------------------------------------------------------------------------------------
  'Datum: 29-7-2014
  'Auteur: Siebe Bosch
  'Deze routine indexeert de zwaarste neerslaggebeurtenissen uit een opgegeven range en schrijft de indexnummers naar een naastgelegen resultatenkolom
  'Bovendien maakt hij een overzicht van alle bijkomende volumes
  '---------------------------------------------------------------------------------------------------------------------------------------------------
  Dim i As Long, myIdx As Integer, r As Long, lastRow As Long
  Dim subValRange As Range, subIdxRange As Range
  Dim maxSum As Double, maxIdx As Long, idxSum As Double
  Dim startDate As Date, endDate As Date, nDays As Long
  Dim EventSum() As Double, MaxEvents As Long
  Dim maxCol As Integer
  Dim DateValResultsRange As Range
  
  'zoek het bereik van de gegevens
  r = HeaderRow
  While Not ActiveSheet.Cells(r + 1, DateCol) = ""
    r = r + 1
  Wend
  maxCol = WorksheetFunction.max(DateCol, ValCol, EventIndexCol, EventSumCol)
  Set DateValResultsRange = ActiveSheet.Range(Cells(HeaderRow + 1, DateCol), Cells(r, maxCol))
  
  ActiveSheet.Cells(HeaderRow, EventIndexCol) = "Index for " & duration & "h"
  
  'zoek de start- en einddatum en bereken het gewenste aantal overschrijdingen van de POT-waarde
  startDate = DateValResultsRange.Cells(1, DateCol)
  endDate = DateValResultsRange.Cells(DateValResultsRange.Rows.Count, DateCol)
  nDays = endDate - startDate
  MaxEvents = nDays / 365.25 * PotExceedanceFrequencyPerYear
  ReDim EventSum(1 To MaxEvents)
  
  'opschonen oude resultaten
  Set subIdxRange = DateValResultsRange.Range(Cells(1, EventIndexCol), Cells(DateValResultsRange.Rows.Count, EventIndexCol))
  Call subIdxRange.ClearContents
  Set subIdxRange = DateValResultsRange.Range(Cells(1, EventSumCol), Cells(DateValResultsRange.Rows.Count, EventSumCol))
  Call subIdxRange.ClearContents
  Set subIdxRange = DateValResultsRange.Range(Cells(1, EventExceedanceCol), Cells(DateValResultsRange.Rows.Count, EventExceedanceCol))
  Call subIdxRange.ClearContents
  
  'create a moving window array that contains the sum of each window
  Dim movingWindowSum() As Double, inUseSum() As Integer
  ReDim movingWindowSum(1 To DateValResultsRange.Rows.Count)
  ReDim inUseSum(1 To DateValResultsRange.Rows.Count)
  For i = 1 To DateValResultsRange.Rows.Count - duration + 1
    movingWindowSum(i) = Application.WorksheetFunction.Sum(DateValResultsRange.Range(DateValResultsRange.Cells(i, 2), DateValResultsRange.Cells(i + duration - 1, 2)))
  Next
    
  'next walk through the moving window array to find the highest volumes, starting with the maximum (rank 1) and moving up in rank (=lower volume)
  For myIdx = 1 To MaxEvents
    ProgressRange = myIdx / MaxEvents
    DoEvents
    maxSum = 0
    
    'make a distinction between summer and winter if required
    If IncludeSummer = True And IncludeWinter = True Then
      For i = 1 To DateValResultsRange.Rows.Count - duration + 1
        If movingWindowSum(i) = 0 Then
          i = i + duration - 1
        ElseIf movingWindowSum(i) > maxSum And inUseSum(i) = 0 Then
          maxSum = movingWindowSum(i)
          maxIdx = i
        End If
      Next
    ElseIf IncludeSummer = True Then
      For i = 1 To DateValResultsRange.Rows.Count - duration + 1
        If movingWindowSum(i) = 0 Then
          i = i + duration - 1
        ElseIf movingWindowSum(i) > maxSum And inUseSum(i) = 0 Then
          If METEOROLOGISCHHALFJAAR(DateValResultsRange.Cells(i, 1)) = "zomer" Then
            maxSum = movingWindowSum(i)
            maxIdx = i
          End If
        End If
      Next
    ElseIf IncludeWinter = True Then
      For i = 1 To DateValResultsRange.Rows.Count - duration + 1
        If movingWindowSum(i) = 0 Then
          i = i + duration - 1
        ElseIf movingWindowSum(i) > maxSum And inUseSum(i) = 0 Then
          If METEOROLOGISCHHALFJAAR(DateValResultsRange.Cells(i, 1)) = "winter" Then
          maxSum = movingWindowSum(i)
          maxIdx = i
          End If
        End If
      Next
    End If
    
    'write the index number to the worksheet
    Set subIdxRange = DateValResultsRange.Range(DateValResultsRange.Cells(maxIdx, EventIndexCol), DateValResultsRange.Cells(maxIdx + duration - 1, EventIndexCol))
    subIdxRange = myIdx
    
    'zet de relevante velden in de inUse-array plus uitloop voor de minimumruimte tussen twee events op 'bezet'. Let op: ook een stuk terug! de array bevat immers een vooruitblik
    For i = maxIdx To Application.WorksheetFunction.Min(maxIdx + duration - 1 + MinimumTimeStepsBetweenEvents, DateValResultsRange.Rows.Count)
      inUseSum(i) = 1
    Next
    For i = (maxIdx) To Application.WorksheetFunction.max(1, (maxIdx - duration + 1 - MinimumTimeStepsBetweenEvents)) Step -1
      inUseSum(i) = 1
    Next
        
    'store the event sum in the array
    EventSum(myIdx) = maxSum
  Next
  
  'finally write the event sums and threshold exceedance sums to the worksheet
  r = HeaderRow
  ActiveSheet.Cells(r, EventSumCol) = "Volume." & duration & "h"
  ActiveSheet.Cells(r, EventExceedanceCol) = "Exceedance." & duration & "h"
  For myIdx = 1 To MaxEvents
    r = r + 1
    ActiveSheet.Cells(r, EventSumCol) = EventSum(myIdx)
    ActiveSheet.Cells(r, EventExceedanceCol) = EventSum(myIdx) - EventSum(MaxEvents)
  Next

End Sub

Public Sub POTANALYSISMAX(HeaderRow As Integer, DateCol As Integer, ValCol As Integer, EventIndexCol As Integer, EventMaxCol As Integer, EventExceedanceCol As Integer, duration As Integer, MinimumTimeStepsBetweenEvents As Integer, PotExceedanceFrequencyPerYear As Integer, IncludeSummer As Boolean, IncludeWinter As Boolean, ProgressRange As Range)
  '---------------------------------------------------------------------------------------------------------------------------------------------------
  'Datum: 29-7-2014
  'Auteur: Siebe Bosch
  'Deze routine indexeert de zwaarste gebeurtenissen (op basis van maximum) uit een opgegeven range en schrijft de indexnummers naar een naastgelegen resultatenkolom
  'Bovendien maakt hij een overzicht van alle bijkomende volumes
  '---------------------------------------------------------------------------------------------------------------------------------------------------
  Dim i As Long, myIdx As Integer, r As Long, lastRow As Long
  Dim subValRange As Range, subIdxRange As Range
  Dim maxVal As Double, maxIdx As Long, idxMax As Double
  Dim startDate As Date, endDate As Date, nDays As Long
  Dim EventMax() As Double, MaxEvents As Long
  Dim maxCol As Integer
  Dim DateValResultsRange As Range
  
  'zoek het bereik van de gegevens
  r = HeaderRow
  While Not ActiveSheet.Cells(r + 1, DateCol) = ""
    r = r + 1
  Wend
  maxCol = WorksheetFunction.max(DateCol, ValCol, EventIndexCol, EventMaxCol)
  Set DateValResultsRange = ActiveSheet.Range(Cells(HeaderRow + 1, DateCol), Cells(r, maxCol))
  
  ActiveSheet.Cells(HeaderRow, EventIndexCol) = "Index for " & duration & "h"
  
  'zoek de start- en einddatum en bereken het gewenste aantal overschrijdingen van de POT-waarde
  startDate = DateValResultsRange.Cells(1, DateCol)
  endDate = DateValResultsRange.Cells(DateValResultsRange.Rows.Count, DateCol)
  nDays = endDate - startDate
  MaxEvents = nDays / 365.25 * PotExceedanceFrequencyPerYear
  ReDim EventMax(1 To MaxEvents)
  
  'opschonen oude resultaten
  Set subIdxRange = DateValResultsRange.Range(Cells(1, EventIndexCol), Cells(DateValResultsRange.Rows.Count, EventIndexCol))
  Call subIdxRange.ClearContents
  Set subIdxRange = DateValResultsRange.Range(Cells(1, EventMaxCol), Cells(DateValResultsRange.Rows.Count, EventMaxCol))
  Call subIdxRange.ClearContents
  Set subIdxRange = DateValResultsRange.Range(Cells(1, EventExceedanceCol), Cells(DateValResultsRange.Rows.Count, EventExceedanceCol))
  Call subIdxRange.ClearContents
  
  'create a moving window array that contains the Max of each window
  Dim movingWindowMax() As Double, inUseMax() As Integer
  ReDim movingWindowMax(1 To DateValResultsRange.Rows.Count)
  ReDim inUseMax(1 To DateValResultsRange.Rows.Count)
  For i = 1 To DateValResultsRange.Rows.Count - duration + 1
    movingWindowMax(i) = Application.WorksheetFunction.max(DateValResultsRange.Range(DateValResultsRange.Cells(i, 2), DateValResultsRange.Cells(i + duration - 1, 2)))
  Next
    
  'next walk through the moving window array to find the highest volumes, starting with the maximum (rank 1) and moving up in rank (=lower volume)
  For myIdx = 1 To MaxEvents
    ProgressRange = myIdx / MaxEvents
    DoEvents
    maxVal = 0
    
    'make a distinction between Summer and winter if required
    If IncludeSummer = True And IncludeWinter = True Then
      For i = 1 To DateValResultsRange.Rows.Count - duration + 1
        If movingWindowMax(i) = 0 Then
          i = i + duration - 1
        ElseIf movingWindowMax(i) > maxVal And inUseMax(i) = 0 Then
          maxVal = movingWindowMax(i)
          maxIdx = i
        End If
      Next
    ElseIf IncludeSummer = True Then
      For i = 1 To DateValResultsRange.Rows.Count - duration + 1
        If movingWindowMax(i) = 0 Then
          i = i + duration - 1
        ElseIf movingWindowMax(i) > maxVal And inUseMax(i) = 0 Then
          If METEOROLOGISCHHALFJAAR(DateValResultsRange.Cells(i, 1)) = "zomer" Then
            maxVal = movingWindowMax(i)
            maxIdx = i
          End If
        End If
      Next
    ElseIf IncludeWinter = True Then
      For i = 1 To DateValResultsRange.Rows.Count - duration + 1
        If movingWindowMax(i) = 0 Then
          i = i + duration - 1
        ElseIf movingWindowMax(i) > maxVal And inUseMax(i) = 0 Then
          If METEOROLOGISCHHALFJAAR(DateValResultsRange.Cells(i, 1)) = "winter" Then
          maxVal = movingWindowMax(i)
          maxIdx = i
          End If
        End If
      Next
    End If
    
    'write the index number to the worksheet
    Set subIdxRange = DateValResultsRange.Range(DateValResultsRange.Cells(maxIdx, EventIndexCol), DateValResultsRange.Cells(maxIdx + duration - 1, EventIndexCol))
    subIdxRange = myIdx
    
    'zet de relevante velden in de inUse-array plus uitloop voor de minimumruimte tussen twee events op 'bezet'. Let op: ook een stuk terug! de array bevat immers een vooruitblik
    For i = maxIdx To Application.WorksheetFunction.Min(maxIdx + duration - 1 + MinimumTimeStepsBetweenEvents, DateValResultsRange.Rows.Count)
      inUseMax(i) = 1
    Next
    For i = (maxIdx) To Application.WorksheetFunction.max(1, (maxIdx - duration + 1 - MinimumTimeStepsBetweenEvents)) Step -1
      inUseMax(i) = 1
    Next
        
    'store the event Max in the array
    EventMax(myIdx) = maxVal
  Next
  
  'finally write the event Maxs and threshold exceedance Maxs to the worksheet
  r = HeaderRow
  ActiveSheet.Cells(r, EventMaxCol) = "Max." & duration & "h"
  ActiveSheet.Cells(r, EventExceedanceCol) = "Exceedance." & duration & "h"
  For myIdx = 1 To MaxEvents
    r = r + 1
    ActiveSheet.Cells(r, EventMaxCol) = EventMax(myIdx)
    ActiveSheet.Cells(r, EventExceedanceCol) = EventMax(myIdx) - EventMax(MaxEvents)
  Next

End Sub


Public Sub CalculateExtremeEvents(Volumes() As Variant, DuurInUse() As Long, Duur As Long, dIdx As Long, nExtremen As Long, ProgressRange As Range, StartRow As Long)
  
  Dim r As Long, c As Long, i As Long, j As Long, K As Long, rMax As Long
  Dim NSLMaxSom As Double, mySom As Double 'rMax is het rijnummer van het record van de hoogste dagneerslag, NSLMax de dagsom
  Dim SumRange As Range, Inuse As Long
    
  '----------------------------------------------------------------------------------------------------------------------------------
  'Datum: 1-11-2010
  'Auteur: Siebe Bosch
  'Deze routine zoekt in een array van neerslagvolumes welke neerslaggebeurtenissen met een opgegeven duur de 1000 zwaarste zijn
  'op het werkblad schrijft hij daarna het volgnummer van de zwaarte weg. 1 = zwaarste, 1000 = minst zware
  '----------------------------------------------------------------------------------------------------------------------------------
  For i = 1 To nExtremen                            'bepaal de zwaarste neerslaggebeurtenissen met de opgegeven duur
    ProgressRange.Cells(1, 1) = i / nExtremen
    DoEvents
    
    NSLMaxSom = 0                                   'initialiseer de maximum neerslagsom
    For j = 1 To UBound(Volumes(), 1) - Duur + 1    'doorloop de hele reeks en zoek naar de zwaarste neerslagsom over het opgegeven aantal uren
      
      mySom = Volumes(j, 1)                         'de som begint altijd met het eerste record
      Inuse = DuurInUse(j, dIdx)                    'controleer of dit record niet al aangemerkt is als "inuse" oftewel: al in een maximum verwerkt
      If Inuse = 0 Then                             'alleen als dit record nog niet in gebruik is
        For K = 1 To Duur - 1                       'doorloop de rest van de uurgegevens voor de opgegeven neerslagduur
          mySom = mySom + Volumes(j + K, 1)         'sommeer ze
          Inuse = DuurInUse(j + K, dIdx)            'wederom controle of geen van de opgetelde volumes al in gebruik waren
          If Inuse > 0 Then
            j = j + K + Duur - 1                    'als een record binnen de komende duur al in gebruik is, kunnen we de teller meteen doorzetten tot voorbij de hele neerslaggebeurtenis
            Exit For                                  'als een record al in gebruik is, kunnen we deze loop meteen al verlaten
          End If
        Next
        If mySom > NSLMaxSom And Inuse = 0 Then       'alleen als de som over de duur groter is dan het totnogtoe geregistreerde maximum EN geen van de records is in gebruik, gaan we door
          rMax = j + StartRow - 1                     'registreer het rijnummer dat hoort bij de gevonden duur met maximum
          NSLMaxSom = mySom                           'let het gevonden maximum ook als zodanig vast
        End If
      End If
    Next
    
    'nu we de (i-1)-na zwaarste gebeurtenis hebben gevonden, voegen we hem toe aan de collectie
    'door een vlaggetje met het volgnummer naast het record te zetten staat hij ook meteen te boek als reeds verwerkt
    For r = rMax To rMax + Duur - 1
      ActiveSheet.Cells(r, 2 + dIdx) = i
      DuurInUse(r - StartRow, dIdx) = i
    Next
    
  Next
End Sub

Public Function NASH_SUTCLIFFE(myRange As Range, Datumcol As Integer, MeasCol As Integer, ValsCol As Integer, ContainsHeader As Boolean, Optional ByVal Log As Boolean = False) As Double
  
  On Error GoTo Errorhandling
  
  Dim nObserved As Long
  Dim Sum As Double, sumLog As Double, AvgObserved As Double, AvgLogObserved As Double
  Dim sumTeller As Double, sumNoemer As Double
  Dim ErrStr As String, r As Long, StartRow As Integer
  
  If ContainsHeader Then
    StartRow = 2
  Else
    StartRow = 1
  End If
  
  Sum = 0
  nObserved = 0
  For r = StartRow To myRange.Rows.Count
    nObserved = nObserved + 1
    Sum = Sum + myRange.Cells(r, MeasCol)
    If myRange.Cells(r, MeasCol) > 0 Then sumLog = sumLog + Math.Log(myRange.Cells(r, MeasCol)) 'log-NS
  Next
  
  'calculate the average
  If nObserved = 0 Then
    ErrStr = "No measured data found to compare computed data with. Please check from- and to-dates and time series with measured data."
    GoTo Errorhandling
  Else
    AvgObserved = Sum / nObserved
    AvgLogObserved = sumLog / nObserved
  End If
  
  For r = StartRow To myRange.Rows.Count
    If Not Log Then
      sumTeller = sumTeller + (myRange.Cells(r, MeasCol) - myRange.Cells(r, ValsCol)) ^ 2
      sumNoemer = sumNoemer + (myRange.Cells(r, MeasCol) - AvgObserved) ^ 2
    Else
      sumTeller = sumTeller + (Math.Log(myRange.Cells(r, MeasCol)) - Math.Log(myRange.Cells(r, ValsCol))) ^ 2
      sumNoemer = sumNoemer + (Math.Log(myRange.Cells(r, MeasCol)) - AvgLogObserved) ^ 2
    End If
  Next
  
  NASH_SUTCLIFFE = 1 - (sumTeller / sumNoemer)
  Exit Function
  
Errorhandling:
  MsgBox ("Error in function calcNashSutcliffe. " & ErrStr)
  End
  
  
End Function


Public Function NASH_SUTCLIFFE_FAST(Observed As Range, Computed As Range, Optional ByVal Log As Boolean = False) As Double
  
  On Error GoTo Errorhandling
  
  Dim Sum As Double, sumLog As Double, AvgObserved As Double, AvgLogObserved As Double
  Dim sumTeller As Double, sumNoemer As Double
  Dim ErrStr As String, r As Long
    
  Sum = 0
  For r = 1 To Observed.Rows.Count
    Sum = Sum + Observed.Cells(r, 1)
    If Observed.Cells(r, 1) > 0 Then sumLog = sumLog + Math.Log(Observed.Cells(r, 1)) 'log-NS
  Next
  
  'calculate the average
  If Observed.Rows.Count = 0 Then
    ErrStr = "No measured data found to compare computed data with. Please check from- and to-dates and time series with measured data."
    GoTo Errorhandling
  Else
    AvgObserved = Sum / Observed.Rows.Count
    AvgLogObserved = sumLog / Observed.Rows.Count
  End If
  
  For r = 1 To Computed.Rows.Count
    If Not Log Then
      sumTeller = sumTeller + (Observed.Cells(r, 1) - Computed.Cells(r, 1)) ^ 2
      sumNoemer = sumNoemer + (Observed.Cells(r, 1) - AvgObserved) ^ 2
    Else
      sumTeller = sumTeller + (Math.Log(Observed.Cells(r, 1)) - Math.Log(Computed.Cells(r, 1))) ^ 2
      sumNoemer = sumNoemer + (Math.Log(Observed.Cells(r, 1)) - AvgLogObserved) ^ 2
    End If
  Next
  
  NASH_SUTCLIFFE_FAST = 1 - (sumTeller / sumNoemer)
  Exit Function
  
Errorhandling:
  MsgBox ("Error in function calcNashSutcliffe. " & ErrStr)
  End
  
  
End Function


Public Sub FILTERBASEFLOW(ByRef ValRangeNoHeader As Range, K As Double, W As Double, BaseflowCol As Integer, InterFlowCol As Integer)
  'this routine filters the baseflow out of the total discharge
  'it does so by applying the method by prof. Patrick Willems (Leuven University) as implemented in his tool Wetspro
  Dim alpha As Double, A As Double, b As Double, c As Double, v As Double
  Dim i As Long, iPar As Long
  
  Dim TotalFlow() As Double
  Dim InterFlow() As Double
  Dim BaseFlow() As Double
  Dim prevTotalFlow As Double, prevInterFlow As Double, prevBaseFlow As Double
  
  ReDim TotalFlow(ValRangeNoHeader.Rows.Count)
  ReDim InterFlow(ValRangeNoHeader.Rows.Count)
  ReDim BaseFlow(ValRangeNoHeader.Rows.Count)
  
  For iPar = 1 To 3
    For i = 1 To ValRangeNoHeader.Count
  
      'retrieve the total, inter and baseflow from the previous timestep
      If i = 1 Then
        prevTotalFlow = 0
        prevInterFlow = 0
        prevBaseFlow = 0
      Else
        prevTotalFlow = TotalFlow(i - 1)
        prevInterFlow = InterFlow(i - 1)
        prevBaseFlow = BaseFlow(i - 1)
      End If
    
      If iPar = 1 Then 'total flow
        TotalFlow(i) = ValRangeNoHeader.Cells(i, 1)
      ElseIf iPar = 2 Then 'interflow
        alpha = Math.Exp(-1 / K)
        v = (1 - W) / W
        A = ((2 + v) * alpha - v) / (2 + v - v * alpha)
        b = 2 / (2 + v - v * alpha)
        c = 0.5 * v
        'curFlow.InterFlow = a * prevFlow.InterFlow + b * (curFlow.Value - alpha * prevFlow.Value)
        InterFlow(i) = A * prevInterFlow + b * (TotalFlow(i) - alpha * prevTotalFlow)

      ElseIf iPar = 3 Then  'baseflow
        alpha = Math.Exp(-1 / K)
        v = (1 - W) / W
        A = ((2 + v) * alpha - v) / (2 + v - v * alpha)
        b = 2 / (2 + v - v * alpha)
        c = 0.5 * v
        BaseFlow(i) = alpha * prevBaseFlow + c * (1 - alpha) * (prevInterFlow + InterFlow(i))
        
      End If
    Next
  Next
  
  ActiveSheet.Cells(ValRangeNoHeader.Cells(1, 1).Row - 1, BaseflowCol) = "Baseflow"
  ActiveSheet.Cells(ValRangeNoHeader.Cells(1, 1).Row - 1, InterFlowCol) = "Interflow"
  
  For i = 1 To ValRangeNoHeader.Count
    ActiveSheet.Cells(ValRangeNoHeader.Cells(i, 1).Row, BaseflowCol) = BaseFlow(i)
    ActiveSheet.Cells(ValRangeNoHeader.Cells(i, 1).Row, InterFlowCol) = InterFlow(i)
  Next
  
  
End Sub

Public Function HOOGHOUDT_q(k1 As Double, k2 As Double, D As Double, L As Double, h As Double) As Double
  'k1 = doorlatendheid bovenste laag
  'k2 = doorlatendheid onderste laag
  'Dikte gedraineerde laag
  'L = afstand tussen de drains
  'h = maximale opbolling (m) tussen de drains
  'q = stationaire specifieke afvoer (m/s)
  'let op: K1 en K2 mogen alleen verschillen als de drainagemiddelen exact op de scheidingslaag liggen!
  
  HOOGHOUDT_q = (8 * k2 * D * h + 4 * k1 * h ^ 2) / L ^ 2
  
End Function

Public Function HOOGHOUDT_L(k1 As Double, k2 As Double, D As Double, Q As Double, h As Double) As Double
  'k1 = doorlatendheid bovenste laag
  'k2 = doorlatendheid onderste laag
  'Dikte gedraineerde laag
  'L = afstand tussen de drains
  'h = maximale opbolling (m) tussen de drains
  'q = stationaire specifieke afvoer (m/s)
  'let op: K1 en K2 mogen alleen verschillen als de drainagemiddelen exact op de scheidingslaag liggen!
  
  HOOGHOUDT_L = Sqr((8 * k2 * D * h + 4 * k1 * h ^ 2) / Q)
  
End Function

Public Function YZConveyanceManningSegmentedByDepth(YRange As Range, ZRange As Range, ManningValue As Double, ResultsRange As Range, DepthColIdx As Integer, KColIdx As Integer)
    Dim Row As Integer
    Dim K As Double, Depth As Double
    For Row = 1 To ResultsRange.Rows.Count
        Depth = ResultsRange.Cells(Row, DepthColIdx)
        If Depth > 0 Then
            K = YZConveyanceManningSegmented(YRange, ZRange, ManningValue, Depth)
            ResultsRange.Cells(Row, KColIdx) = K
        Else
            ResultsRange.Cells(Row, KColIdx) = 0
        End If
    Next
End Function

Function subRange(r As Range, startPos As Integer, endPos As Integer) As Range
    Set subRange = r.Parent.Range(r.Cells(startPos), r.Cells(endPos))
End Function

Public Function YZConveyanceManningSegmented(ByVal YRange As Range, ByVal ZRange As Range, ManningValue As Double, Depth As Double) As Double
    'this function calculates conveyance K for a given YZ cross section, using the 'segmented' method
    Dim Row As Integer, minZ As Double, Level As Double
    Dim curY As Double, nextY As Double, CurZ As Double, nextZ As Double
    Dim yxsect As Double
    Dim A() As Double, P() As Double, r() As Double, K() As Double
    Dim newYRange As Range
    Dim newZRange As Range
    minZ = 9E+99
    YZConveyanceManningSegmented = 0
    
    'truncate the ranges if the table stops before the end
    Dim i As Integer
    For i = 1 To YRange.Rows.Count
        If YRange.Cells(i, 1) = "" Then
            Set newYRange = subRange(YRange, 1, i - 1)
            Set newZRange = subRange(ZRange, 1, i - 1)
            Exit For
        End If
    Next
    
    Set YRange = newYRange
    Set ZRange = newZRange
    
    ReDim A(1 To YRange.Rows.Count - 1)
    ReDim P(1 To YRange.Rows.Count - 1)
    ReDim r(1 To YRange.Rows.Count - 1)
    ReDim K(1 To YRange.Rows.Count - 1)
    
    'first find the lowest point and calculate the waterlevel at the current depth
    For Row = 1 To YRange.Rows.Count
        If ZRange.Cells(Row, 1) < minZ Then minZ = ZRange.Cells(Row, 1)
    Next
    Level = minZ + Depth
    
    'now calculate the hydraulic properties
    For Row = 1 To YRange.Rows.Count - 1
        A(Row) = 0
        P(Row) = 0
        r(Row) = 0
        K(Row) = 0
        curY = YRange.Cells(Row, 1)
        CurZ = ZRange.Cells(Row, 1)
        nextY = YRange.Cells(Row + 1, 1)
        nextZ = ZRange.Cells(Row + 1, 1)
        If Level > CurZ And Level > nextZ Then
            P(Row) = P(Row) + PYTHAGORAS(nextZ - CurZ, nextY - curY)
            A(Row) = A(Row) + (nextY - curY) * (Level - Maximum(nextZ, CurZ)) + (nextY - curY) * Math.Abs(nextZ - CurZ) / 2
        ElseIf Level > CurZ Then
            yxsect = Interpolate(CurZ, curY, nextZ, nextY, Level)
            P(Row) = P(Row) + PYTHAGORAS(Level - CurZ, yxsect - curY)
            A(Row) = A(Row) + (Level - CurZ) * Math.Abs(yxsect - curY)
        ElseIf Level > nextZ Then
            yxsect = Interpolate(CurZ, curY, nextZ, nextY, Level)
            P(Row) = P(Row) + PYTHAGORAS(Level - nextZ, nextY - yxsect)
            A(Row) = A(Row) + (Level - nextZ) * Math.Abs(nextY - yxsect)
        End If
         
        If P(Row) > 0 Then
            r(Row) = A(Row) / P(Row)
            K(Row) = A(Row) * r(Row) ^ (1 / 6) / ManningValue * Sqr(r(Row))
        Else
            r(Row) = 0
            K(Row) = 0
        End If
    Next
    
    For Row = 1 To YRange.Rows.Count - 1
        YZConveyanceManningSegmented = YZConveyanceManningSegmented + K(Row)
    Next
    
    
End Function

Public Function YZHydraulicPropertiesByDepth(YRange As Range, ZRange As Range, ResultsRange As Range, DepthColIdx As Integer, AColIdx As Integer, PColIdx As Integer, RColIdx As Integer) As Boolean
    'this function calculates hydraulic properties for a given YZ cross section, as a function of given depths:
    'A (wetted area)
    'P (wetted perimeter)
    'R (hydraulic radius)
    On Error GoTo errorhandler
    Dim Row As Integer
    Dim A As Double, P As Double, r As Double, Depth As Double
    For Row = 1 To ResultsRange.Rows.Count
        Depth = ResultsRange.Cells(Row, DepthColIdx)
        If Depth > 0 Then
            Call YZHydraulicProperties(YRange, ZRange, Depth, A, P, r)
            ResultsRange.Cells(Row, AColIdx).Value = Format(A, "0.00")
            ResultsRange.Cells(Row, PColIdx).Value = P
            ResultsRange.Cells(Row, RColIdx).Value = r
        End If
    Next
    YZHydraulicPropertiesByDepth = True
errorhandler:
    MsgBox ("Error")
    
End Function


Public Function YZWettedArea(YRange As Range, ZRange As Range, Depth As Double) As Double
    'this function calculates the wetted area  for a given YZ cross section and a given depth:
    'A (wetted area)
    'P (wetted perimeter)
    'R (hydraulic radius)
    Dim A As Double
    Dim P As Double
    Dim r As Double
    Call YZHydraulicProperties(YRange, ZRange, Depth, A, P, r)
    YZWettedArea = A
End Function

Public Function TabulatedWettedAreaByDepth(YRange As Range, ZRange As Range, Depth As Double) As Double
    'this function calculates the wetted area  for a given YZ cross section and a given depth:
    'A (wetted area)
    'P (wetted perimeter)
    'R (hydraulic radius)
    Dim A As Double
    Dim P As Double
    Dim r As Double
    Call TabulatedHydraulicProperties(YRange, ZRange, Depth, A, P, r)
    TabulatedWettedAreaByDepth = A
End Function


Public Function TabulatedWettedAreaByWaterlevel(Waterlevel As Double, ZRange As Range, WRange As Range, ProfileVerticalShift As Double) As Double
    Dim A As Double
    Dim P As Double
    Dim r As Double
    Dim Depth As Double
    Depth = Waterlevel - ZRange.Cells(1, 1) - ProfileVerticalShift
    Call TabulatedHydraulicProperties(ZRange, WRange, Depth, A, P, r)
    TabulatedWettedAreaByWaterlevel = A
End Function


Public Function YZHydraulicRadius(YRange As Range, ZRange As Range, Depth As Double) As Double
    'this function calculates the hydraulic radius for a given YZ cross section and a given depth:
    'A (wetted area)
    'P (wetted perimeter)
    'R (hydraulic radius)
    Dim A As Double
    Dim P As Double
    Dim r As Double
    Call YZHydraulicProperties(YRange, ZRange, Depth, A, P, r)
    YZHydraulicRadius = r
End Function


Public Function YZHydraulicProperties(YRange As Range, ZRange As Range, Depth As Double, ByRef A As Double, ByRef P As Double, ByRef r As Double) As Boolean
    'this function calculates hydraulic properties for a given YZ cross section, as a function of given depths:
    'A (wetted area)
    'P (wetted perimeter)
    'R (hydraulic radius)
    
    Dim Row As Integer, minZ As Double, Level As Double
    Dim curY As Double, nextY As Double, CurZ As Double, nextZ As Double
    Dim yxsect As Double
    minZ = 9E+99
    A = 0
    P = 0
    r = 0
    
    'first find the lowest point and calculate the waterlevel at the current depth
    For Row = 1 To YRange.Rows.Count
        If ZRange.Cells(Row, 1) < minZ Then minZ = ZRange.Cells(Row, 1)
    Next
    Level = minZ + Depth
    
    'now calculate the hydraulic properties
    For Row = 1 To YRange.Rows.Count - 1
        curY = YRange.Cells(Row, 1)
        CurZ = ZRange.Cells(Row, 1)
        nextY = YRange.Cells(Row + 1, 1)
        nextZ = ZRange.Cells(Row + 1, 1)
        If Level > CurZ And Level > nextZ Then
            P = P + PYTHAGORAS(nextZ - CurZ, nextY - curY)
            A = A + (nextY - curY) * (Level - Maximum(nextZ, CurZ)) + (nextY - curY) * Math.Abs(nextZ - CurZ) / 2
        ElseIf Level > CurZ Then
            yxsect = Interpolate(CurZ, curY, nextZ, nextY, Level)
            P = P + PYTHAGORAS(Level - CurZ, yxsect - curY)
            A = A + (Level - CurZ) * Math.Abs(yxsect - curY)
        ElseIf Level > nextZ Then
            yxsect = Interpolate(CurZ, curY, nextZ, nextY, Level)
            P = P + PYTHAGORAS(Level - nextZ, nextY - yxsect)
            A = A + (Level - nextZ) * Math.Abs(nextY - yxsect)
        End If
     Next
     If P > 0 Then
        r = A / P
     Else
        r = 0
     End If
    YZHydraulicProperties = True
End Function

Public Function TabulatedHydraulicProperties(ZRange As Range, WRange As Range, Depth As Double, ByRef A As Double, ByRef P As Double, ByRef r As Double) As Boolean
    'this function calculates hydraulic properties for a given tabulated cross section, as a function of given depth:
    'A (wetted area)
    'P (wetted perimeter)
    'R (hydraulic radius)
    Dim BedLevel As Double
    Dim PrevZ As Double
    Dim CurZ As Double
    Dim PrevW As Double
    Dim CurW As Double
    Dim PrevD As Double
    Dim CurD As Double
    Dim WidthAtDepth As Double
    Dim Row As Integer
        
    BedLevel = ZRange.Cells(1, 1)
    For Row = 2 To ZRange.Rows.Count
        PrevZ = ZRange.Cells(Row - 1, 1)
        CurZ = ZRange.Cells(Row, 1)
        PrevW = WRange.Cells(Row - 1, 1)
        CurW = WRange.Cells(Row, 1)
        PrevD = PrevZ - BedLevel
        CurD = CurZ - BedLevel
                        
        'zolang de waterdiepte groter is dan de 'diepte' van het huidige segment moeten we doorgaan
        If Depth >= CurD Then
            'add the entire segment
            A = A + (CurZ - PrevZ) * (Application.WorksheetFunction.Min(CurW, PrevW) + Math.Abs(CurW - PrevW) / 2)
            P = P + 2 * PYTHAGORAS((CurZ - PrevZ), Math.Abs(CurW - PrevW) / 2)
        ElseIf Depth >= PrevD Then
            'add part of the segment
            WidthAtDepth = Interpolate(CurD, CurW, PrevD, PrevW, Depth)
            A = A + (Depth - PrevD) * (Application.WorksheetFunction.Min(WidthAtDepth, PrevW) + Math.Abs(WidthAtDepth - PrevW) / 2)
            P = P + 2 * PYTHAGORAS((Depth - PrevD), Math.Abs(WidthAtDepth - PrevW) / 2)
            Exit For
        End If
    Next
    If P > 0 Then r = A / P
    TabulatedHydraulicProperties = True
End Function

Public Function NELENSCHUURMANSFYSIEKVOORKOMENTONBW(GRIDCODE As Integer) As Integer
  'resultaat: 0= openwater, 1 = akkerbouw, 2 = akkerbouw hoogwaardig, 3 = gras, 4 = natuur, 5 = stedelijk
  
  Select Case GRIDCODE
  Case Is = 1 'Dak
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 5
  Case Is = 2 'Zand (constatering: zand in natuur is al afgedekt (Duin, Heide), dus we nemen aan dat het hier stedelijk zand betreft)
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 5
  Case Is = 3 'Half verhard
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 5
  Case Is = 4 'Erf (boerenerf heeft naar ons idee geen recht op T=100 norm. Dus T=10 gehanteerd)
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 3
  Case Is = 5 'Gesloten verharding
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 5
  Case Is = 6 'Onverhard
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 3
  Case Is = 7 'Open verharding
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 5
  Case Is = 8 'Groenvoorziening. Aanname: de T=100 geldt alleen boven de dorpelhoogte, dus alle plantsoenen zijn naar ons idee T=10
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 3
  Case Is = 9 'Naaldbos. Dit interpreteren wij als natuur
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 4
  Case Is = 10 'Gras
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 3
  Case Is = 11 'Grasland. Dit interpreteren wij agrarisch gras, dus T=10
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 3
  Case Is = 12 'Struiken. Wij interpreteren dit als natuur, dus geen norm.
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 3
  Case Is = 13 'Natuurterreinen
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 4
  Case Is = 14 'Boomteelt. Interpreteren wij als hoogwaardige akkerbouw (T=50)
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 2
  Case Is = 15 'Duin
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 4
  Case Is = 16 'Heide
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 4
  Case Is = 17 'Bouwland. Wij interpreteren dit als akkerbouw, dus T=25
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 1
  Case Is = 18 'Houtwal. Omringt normaalgesproken bouwland, dus we hanteren T=25
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 1
  Case Is = 19 'Loofbos. Interpreteren wij als natuur
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 4
  Case Is = 20 'Rietland. Interpreteren wij als natuur
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 4
  Case Is = 21 'Fruitteelt. Interpreteren wij als hoogwaardige land- en tuinbouw
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 2
  Case Is = 22 'Gemengd bos. Interpreteren wij als natuur
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 4
  Case Is = 23 'Braakland. Is vrijwel altijd stedelijk, maar niet verhard. Wij interpreteren dit als T=10
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 3
  Case Is = 26 'Moeras. Interpreteren wij als natuur
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 4
  Case Is = 27 'Kwelder. Natuur
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 4
  Case Is = 28 'Waterberm. We hanteren de norm voor water (dus geen norm)
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 0
  Case Is = 29 'Water. We hanteren de norm voor water (dus geen norm)
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 0
  Case Is = 30 'Overig. Kan vanalles zijn. We hanteren T=10
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 3
  Case Is = 253 'Onbekend. Kan vanalles zijn. We hanteren T=10
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 3
  Case Is = 254 'Water.We hanteren de norm voor water (dus geen norm)
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 0
  End Select
  
End Function

Public Function LGN5TONBW(LGN5Code As Integer) As Integer
  'resultaat: 0= openwater, 1 = akkerbouw, 2 = akkerbouw hoogwaardig, 3 = gras, 4 = natuur, 5 = stedelijk
  
  Select Case LGN5Code
  Case Is = 1
    LGN5TONBW = 3
  Case Is = 2
    LGN5TONBW = 1
  Case Is = 3
    LGN5TONBW = 1
  Case Is = 4
    LGN5TONBW = 1
  Case Is = 5
    LGN5TONBW = 1
  Case Is = 6
    LGN5TONBW = 1
  Case Is = 8
    LGN5TONBW = 2
  Case Is = 9
    LGN5TONBW = 2
  Case Is = 10
    LGN5TONBW = 2
  Case Is = 11
    LGN5TONBW = 3
  Case Is = 12
    LGN5TONBW = 3
  Case Is = 16
    LGN5TONBW = 0
  Case Is = 17
    LGN5TONBW = 0
  Case Is = 18
    LGN5TONBW = 5
  Case Is = 19
    LGN5TONBW = 5
  Case Is = 20
    LGN5TONBW = 3
  Case Is = 21
    LGN5TONBW = 3
  Case Is = 22
    LGN5TONBW = 5
  Case Is = 23
    LGN5TONBW = 3
  Case Is = 24
    LGN5TONBW = 3
  Case Is = 25
    LGN5TONBW = 5
  Case Is = 26
    LGN5TONBW = 5
  Case Is = 30
    LGN5TONBW = 0
  Case Is = 35
    LGN5TONBW = 0
  Case Is = 36
    LGN5TONBW = 0
  Case Is = 37
    LGN5TONBW = 0
  Case Is = 38
    LGN5TONBW = 0
  Case Is = 39
    LGN5TONBW = 0
  Case Is = 40
    LGN5TONBW = 0
  Case Is = 41
    LGN5TONBW = 0
  Case Is = 42
    LGN5TONBW = 0
  Case Is = 43
    LGN5TONBW = 0
  Case Is = 45
    LGN5TONBW = 0
  Case Is = 46
    LGN5TONBW = 0
  End Select
  
End Function

Public Function LGN2SOBEK(LGNCODE As Long) As Long

'1 = grass
'2 = corn
'3 = potatoes
'4 = sugarbeet
'5 = grain
'6 = miscellaneous
'7 = non-arable land
'8 = greenhouse area
'9 = orchard
'10 = bulbous plants
'11 = foliage forest
'12 = pine forest
'13 = nature
'14 = fallow
'15 = vegetables
'16 = flowers

'zelf toegevoegd:
'17 = water
'18 = verhard

Select Case LGNCODE
  Case Is = 1 'gras
    LGN2SOBEK = 1
  Case Is = 2 'mais
    LGN2SOBEK = 2
  Case Is = 3 'aardappelen
    LGN2SOBEK = 3
  Case Is = 4 'suikerbiet
    LGN2SOBEK = 4
  Case Is = 5 'graan
    LGN2SOBEK = 5
  Case Is = 6 'overige landbouwgewassen
    LGN2SOBEK = 6
  Case Is = 8 'kassen
    LGN2SOBEK = 8
  Case Is = 9 'boomgaard
    LGN2SOBEK = 9
  Case Is = 10 'bollenteelt
    LGN2SOBEK = 10
  Case Is = 11 'loofbos
    LGN2SOBEK = 11
  Case Is = 12 'naaldbos
    LGN2SOBEK = 12
  Case Is = 16 'zoet water
    LGN2SOBEK = 17
  Case Is = 17 'zout water
    LGN2SOBEK = 17
  Case Is = 18 'stedelijk bebouwd
    LGN2SOBEK = 18
  Case Is = 19 'bebouwd buitengebied
    LGN2SOBEK = 18
  Case Is = 20 'loofbos in bebouwd gebied
    LGN2SOBEK = 1
  Case Is = 21 'naaldbos in bebouwd gebied
    LGN2SOBEK = 1
  Case Is = 22 'bos met dichte bebouwing
    LGN2SOBEK = 18
  Case Is = 23 'gras in bebouwd gebied
    LGN2SOBEK = 1
  Case Is = 24 'kale grond in bebouwd buitengebied
    LGN2SOBEK = 1
  Case Is = 25 'hoofdwegen en spoorwegen
    LGN2SOBEK = 18
  Case Is = 26 'bebouwing in agrarisch gebied
    LGN2SOBEK = 18
  Case Is = 28 'Gras in secundair bebouwd gebied
    LGN2SOBEK = 1
  Case Is = 30 'kwelders
    LGN2SOBEK = 13
  Case Is = 35 'open stuifzand
    LGN2SOBEK = 13
  Case Is = 36 'heide
    LGN2SOBEK = 13
  Case Is = 37 'matig vergraste heide
    LGN2SOBEK = 13
  Case Is = 38 'sterk vergraste heide
    LGN2SOBEK = 13
  Case Is = 39 'hoogveen
    LGN2SOBEK = 13
  Case Is = 40 'bos in hoogveen
    LGN2SOBEK = 13
  Case Is = 41 'overige moerasvegetatie
    LGN2SOBEK = 13
  Case Is = 42 'rietvegetatie
    LGN2SOBEK = 13
  Case Is = 43 'bos in moerasgebied
    LGN2SOBEK = 13
  Case Is = 45 'overig open begroeid natuurgebied
    LGN2SOBEK = 13
  Case Is = 46 'kale grond in natuurgebied
    LGN2SOBEK = 13
  Case Is = 61 'boomkwekerijen
    LGN2SOBEK = 9
  Case Is = 61 'fruitkwekerijen
    LGN2SOBEK = 9
End Select

End Function

Public Function ERNSTRecord(ID As String, a1 As Double, a2 As Double, a3 As Double, a4 As Double, lv1 As Double, lv2 As Double, lv3 As Double, ainf As Double) As String
  ERNSTRecord = "ERNS id '" & ID & "' nm '" & ID & "' cvi " & ainf & " cvo " & a1 & " " & a2 & " " & a3 & " " & a4 & " lv " & lv1 & " " & lv2 & " " & lv3 & " cvs 1 erns"
End Function

Public Function Bod2ZandKleiVeen(bc As String) As String
        Dim CapSimCode As Long
        CapSimCode = BOD2CAPSIM(bc)
        Select Case CapSimCode
                Case Is = 101
                        Bod2ZandKleiVeen = "VEEN"
                Case Is = 102
                        Bod2ZandKleiVeen = "VEEN"
                Case Is = 103
                        Bod2ZandKleiVeen = "VEEN"
                Case Is = 104
                        Bod2ZandKleiVeen = "VEEN"
                Case Is = 105
                        Bod2ZandKleiVeen = "VEEN"
                Case Is = 106
                        Bod2ZandKleiVeen = "VEEN"
                Case Is = 107
                        Bod2ZandKleiVeen = "ZAND"
                Case Is = 108
                        Bod2ZandKleiVeen = "ZAND"
                Case Is = 109
                        Bod2ZandKleiVeen = "ZAND"
                Case Is = 110
                        Bod2ZandKleiVeen = "ZAND"
                Case Is = 111
                        Bod2ZandKleiVeen = "ZAND"
                Case Is = 112
                        Bod2ZandKleiVeen = "ZAND"
                Case Is = 113
                        Bod2ZandKleiVeen = "ZAND"
                Case Is = 114
                        Bod2ZandKleiVeen = "ZAND"
                Case Is = 115
                        Bod2ZandKleiVeen = "ZAND"
                Case Is = 116
                        Bod2ZandKleiVeen = "KLEI"
                Case Is = 117
                        Bod2ZandKleiVeen = "KLEI"
                Case Is = 118
                        Bod2ZandKleiVeen = "KLEI"
                Case Is = 119
                        Bod2ZandKleiVeen = "KLEI"
                Case Is = 120
                        Bod2ZandKleiVeen = "KLEI"
                Case Is = 121
                        Bod2ZandKleiVeen = "KLEI"
        End Select
End Function

Public Function GHGGLG2GT(GHGmMV As Double, GLGmMV As Double) As String
    If GHGmMV < 0.2 Or GLGmMV < 0.5 Then
        GHGGLG2GT = "I"
    
    End If
    
End Function

Public Function BOD2CAPSIM(bc As String) As Long
'converteert bodemtypes uit de Bodemkaart Nederland naar het corresponderende CAPSIM bodemnummer in SOBEK

'knip de grondwatertrap eraf!
bc = ParseString(bc, "-")
If bc = "|a GROEVE" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|b AFGRAV" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|c OPHOOG" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|d EGAL" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|e VERWERK" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|f TERP" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|g MOERAS" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|g WATER" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|h BEBOUW" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|h DIJK" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|i BOVLAND" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|j MYNSTRT" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "AAK" Then BOD2CAPSIM = 119 'afgegraven klei
If bc = "AAKp" Then BOD2CAPSIM = 119
If bc = "AAP" Then BOD2CAPSIM = 105
If bc = "ABk" Then BOD2CAPSIM = 119
If bc = "ABkt" Then BOD2CAPSIM = 119
If bc = "ABl" Then BOD2CAPSIM = 121
If bc = "ABv" Then BOD2CAPSIM = 105
If bc = "ABvg" Then BOD2CAPSIM = 105
If bc = "ABvt" Then BOD2CAPSIM = 105
If bc = "ABvx" Then BOD2CAPSIM = 105
If bc = "ABz" Then BOD2CAPSIM = 113
If bc = "ABzt" Then BOD2CAPSIM = 113
If bc = "AD" Then BOD2CAPSIM = 113      'Duin- en kweldergronden
If bc = "AEk9" Then BOD2CAPSIM = 116
If bc = "AEm5" Then BOD2CAPSIM = 115
If bc = "AEm8" Then BOD2CAPSIM = 116
If bc = "AEm9" Then BOD2CAPSIM = 116
If bc = "AEm9A" Then BOD2CAPSIM = 116
If bc = "AEp6A" Then BOD2CAPSIM = 116
If bc = "AEp7A" Then BOD2CAPSIM = 116
If bc = "AFk" Then BOD2CAPSIM = 113
If bc = "AGm9C" Then BOD2CAPSIM = 117   'hollebollige, gemoerde zeekleigronden
If bc = "AFz" Then BOD2CAPSIM = 113
If bc = "Aha" Then BOD2CAPSIM = 121     'glauconiethellingronden
If bc = "AHa" Then BOD2CAPSIM = 121     'glauconiethellingronden
If bc = "AHc" Then BOD2CAPSIM = 121
If bc = "AHk" Then BOD2CAPSIM = 121
If bc = "AHl" Then BOD2CAPSIM = 121
If bc = "Ahs" Then BOD2CAPSIM = 121
If bc = "AHs" Then BOD2CAPSIM = 121
If bc = "AHt" Then BOD2CAPSIM = 121
If bc = "AHv" Then BOD2CAPSIM = 121
If bc = "AHz" Then BOD2CAPSIM = 121
If bc = "AK" Then BOD2CAPSIM = 119
If bc = "AKp" Then BOD2CAPSIM = 119
If bc = "ALu" Then BOD2CAPSIM = 116
If bc = "AM" Then BOD2CAPSIM = 119
If bc = "AMm" Then BOD2CAPSIM = 115
If bc = "AO" Then BOD2CAPSIM = 119
If bc = "AOg" Then BOD2CAPSIM = 119
If bc = "AOp" Then BOD2CAPSIM = 119
If bc = "AOv" Then BOD2CAPSIM = 119
If bc = "AP" Then BOD2CAPSIM = 101
If bc = "App" Then BOD2CAPSIM = 102
If bc = "AQ" Then BOD2CAPSIM = 107
If bc = "AR" Then BOD2CAPSIM = 119
If bc = "AS" Then BOD2CAPSIM = 107
If bc = "aVc" Then BOD2CAPSIM = 101
If bc = "AVk" Then BOD2CAPSIM = 105
If bc = "AVo" Then BOD2CAPSIM = 101
If bc = "aVp" Then BOD2CAPSIM = 102
If bc = "aVpg" Then BOD2CAPSIM = 102
If bc = "aVpx" Then BOD2CAPSIM = 102
If bc = "aVs" Then BOD2CAPSIM = 101
If bc = "aVz" Then BOD2CAPSIM = 102
If bc = "aVzt" Then BOD2CAPSIM = 102
If bc = "aVzx" Then BOD2CAPSIM = 102
If bc = "AWg" Then BOD2CAPSIM = 116
If bc = "AWo" Then BOD2CAPSIM = 106
If bc = "AWv" Then BOD2CAPSIM = 106
If bc = "AZ1" Then BOD2CAPSIM = 114
If bc = "AZW0A" Then BOD2CAPSIM = 107
If bc = "AZW0Al" Then BOD2CAPSIM = 107
If bc = "AZW0Av" Then BOD2CAPSIM = 107
If bc = "AZW1A" Then BOD2CAPSIM = 119
If bc = "AZW1Ar" Then BOD2CAPSIM = 119
If bc = "AZW1Aw" Then BOD2CAPSIM = 119
If bc = "AZW5A" Then BOD2CAPSIM = 119
If bc = "AZW6A" Then BOD2CAPSIM = 119
If bc = "AZW6Al" Then BOD2CAPSIM = 116
If bc = "AZW6Alv" Then BOD2CAPSIM = 118
If bc = "AZW7Al" Then BOD2CAPSIM = 116
If bc = "AZW7A" Then BOD2CAPSIM = 116
If bc = "AZW7Alw" Then BOD2CAPSIM = 116
If bc = "AZW7Alwp" Then BOD2CAPSIM = 119
If bc = "AZW8A" Then BOD2CAPSIM = 116
If bc = "AZW8Al" Then BOD2CAPSIM = 116
If bc = "AZW8Alw" Then BOD2CAPSIM = 116
If bc = "bEZ21" Then BOD2CAPSIM = 112
If bc = "bEZ21g" Then BOD2CAPSIM = 112
If bc = "bEZ21x" Then BOD2CAPSIM = 112
If bc = "bEZ23" Then BOD2CAPSIM = 112
If bc = "bEZ23g" Then BOD2CAPSIM = 112
If bc = "bEZ23t" Then BOD2CAPSIM = 112
If bc = "bEZ23x" Then BOD2CAPSIM = 112
If bc = "bEZ30" Then BOD2CAPSIM = 112
If bc = "bEZ30x" Then BOD2CAPSIM = 112
If bc = "bgMn15C" Then BOD2CAPSIM = 115
If bc = "bgMn25C" Then BOD2CAPSIM = 115
If bc = "bgMn53C" Then BOD2CAPSIM = 117
If bc = "BKd25" Then BOD2CAPSIM = 115
If bc = "BKd25x" Then BOD2CAPSIM = 115
If bc = "BKd26" Then BOD2CAPSIM = 115
If bc = "BKh25" Then BOD2CAPSIM = 115
If bc = "BKh25x" Then BOD2CAPSIM = 115
If bc = "BKh26" Then BOD2CAPSIM = 115
If bc = "BKh26x" Then BOD2CAPSIM = 115
If bc = "BLb6" Then BOD2CAPSIM = 121
If bc = "BLb6g" Then BOD2CAPSIM = 121
If bc = "BLb6k" Then BOD2CAPSIM = 121
If bc = "BLb6s" Then BOD2CAPSIM = 121
If bc = "BLd5" Then BOD2CAPSIM = 121
If bc = "BLd5g" Then BOD2CAPSIM = 121
If bc = "BLd5t" Then BOD2CAPSIM = 121
If bc = "BLd6" Then BOD2CAPSIM = 121
If bc = "BLd6m" Then BOD2CAPSIM = 121
If bc = "BLh5" Then BOD2CAPSIM = 121
If bc = "BLh5m" Then BOD2CAPSIM = 121
If bc = "BLh6" Then BOD2CAPSIM = 121
If bc = "BLh6g" Then BOD2CAPSIM = 121
If bc = "BLh6m" Then BOD2CAPSIM = 121
If bc = "BLh6s" Then BOD2CAPSIM = 121
If bc = "BLn5" Then BOD2CAPSIM = 121
If bc = "BLn5m" Then BOD2CAPSIM = 121
If bc = "BLn5t" Then BOD2CAPSIM = 121
If bc = "BLn6" Then BOD2CAPSIM = 121
If bc = "BLn6g" Then BOD2CAPSIM = 121
If bc = "BLn6m" Then BOD2CAPSIM = 121
If bc = "BLn6s" Then BOD2CAPSIM = 121
If bc = "bMn15A" Then BOD2CAPSIM = 115
If bc = "bMn15C" Then BOD2CAPSIM = 115
If bc = "bMn25A" Then BOD2CAPSIM = 115
If bc = "bMn25C" Then BOD2CAPSIM = 115
If bc = "bMn35A" Then BOD2CAPSIM = 116
If bc = "bMn45A" Then BOD2CAPSIM = 117
If bc = "bMn56Cp" Then BOD2CAPSIM = 119
If bc = "bMn85C" Then BOD2CAPSIM = 116
If bc = "bMn86C" Then BOD2CAPSIM = 117
If bc = "bRn46C" Then BOD2CAPSIM = 117
If bc = "BZd23" Then BOD2CAPSIM = 113
If bc = "BZd24" Then BOD2CAPSIM = 113
If bc = "cHd21" Then BOD2CAPSIM = 108
If bc = "cHd21g" Then BOD2CAPSIM = 110
If bc = "cHd21x" Then BOD2CAPSIM = 111
If bc = "cHd23" Then BOD2CAPSIM = 113
If bc = "cHd23x" Then BOD2CAPSIM = 111
If bc = "cHd30" Then BOD2CAPSIM = 114
If bc = "cHn21" Then BOD2CAPSIM = 109
If bc = "cHn21g" Then BOD2CAPSIM = 110
If bc = "cHn21t" Then BOD2CAPSIM = 111
If bc = "cHn21w" Then BOD2CAPSIM = 111
If bc = "cHn21x" Then BOD2CAPSIM = 111
If bc = "cHn23" Then BOD2CAPSIM = 113
If bc = "cHn23g" Then BOD2CAPSIM = 110
If bc = "cHn23t" Then BOD2CAPSIM = 111
If bc = "cHn23wx" Then BOD2CAPSIM = 111
If bc = "cHn23x" Then BOD2CAPSIM = 111
If bc = "cHn30" Then BOD2CAPSIM = 114
If bc = "cHn30g" Then BOD2CAPSIM = 114
If bc = "cY21" Then BOD2CAPSIM = 109
If bc = "cY21g" Then BOD2CAPSIM = 110
If bc = "cY21x" Then BOD2CAPSIM = 111
If bc = "cY23" Then BOD2CAPSIM = 113
If bc = "cY23g" Then BOD2CAPSIM = 113
If bc = "cY23x" Then BOD2CAPSIM = 111
If bc = "cY30" Then BOD2CAPSIM = 114
If bc = "cY30g" Then BOD2CAPSIM = 114
If bc = "cZd21" Then BOD2CAPSIM = 108
If bc = "cZd21g" Then BOD2CAPSIM = 110
If bc = "cZd23" Then BOD2CAPSIM = 113
If bc = "cZd30" Then BOD2CAPSIM = 114
If bc = "dgMn58Cv" Then BOD2CAPSIM = 117
If bc = "dgMn83C" Then BOD2CAPSIM = 117
If bc = "dgMn88Cv" Then BOD2CAPSIM = 117
If bc = "dhVb" Then BOD2CAPSIM = 101
If bc = "dhVk" Then BOD2CAPSIM = 106
If bc = "dhVr" Then BOD2CAPSIM = 101
If bc = "dkVc" Then BOD2CAPSIM = 103
If bc = "dMn86C" Then BOD2CAPSIM = 117
If bc = "dMv41C" Then BOD2CAPSIM = 118
If bc = "dMv61C" Then BOD2CAPSIM = 118
If bc = "dpVc" Then BOD2CAPSIM = 103
If bc = "dVc" Then BOD2CAPSIM = 101
If bc = "dVd" Then BOD2CAPSIM = 101
If bc = "dVk" Then BOD2CAPSIM = 106
If bc = "dVr" Then BOD2CAPSIM = 101
If bc = "dWo" Then BOD2CAPSIM = 106
If bc = "dWol" Then BOD2CAPSIM = 106
If bc = "EK19" Then BOD2CAPSIM = 115
If bc = "EK19p" Then BOD2CAPSIM = 119
If bc = "EK19x" Then BOD2CAPSIM = 115
If bc = "EK76" Then BOD2CAPSIM = 117
If bc = "EK79" Then BOD2CAPSIM = 116
If bc = "EK79v" Then BOD2CAPSIM = 116
If bc = "EK79w" Then BOD2CAPSIM = 116
If bc = "EL5" Then BOD2CAPSIM = 121
If bc = "eMn12Ap" Then BOD2CAPSIM = 119
If bc = "eMn15A" Then BOD2CAPSIM = 115
If bc = "eMn15Ap" Then BOD2CAPSIM = 119
If bc = "eMn22A" Then BOD2CAPSIM = 119
If bc = "eMn22Ap" Then BOD2CAPSIM = 119
If bc = "eMn25A" Then BOD2CAPSIM = 115
If bc = "eMn25Ap" Then BOD2CAPSIM = 119
If bc = "eMn25Av" Then BOD2CAPSIM = 118
If bc = "eMn35A" Then BOD2CAPSIM = 116
If bc = "eMn35Ap" Then BOD2CAPSIM = 119
If bc = "eMn35Av" Then BOD2CAPSIM = 118
If bc = "eMn35Awp" Then BOD2CAPSIM = 119
If bc = "eMn45A" Then BOD2CAPSIM = 117
If bc = "eMn45Ap" Then BOD2CAPSIM = 117
If bc = "eMn45Av" Then BOD2CAPSIM = 118
If bc = "eMn52Cg" Then BOD2CAPSIM = 119
If bc = "eMn52Cp" Then BOD2CAPSIM = 119
If bc = "eMn52Cwp" Then BOD2CAPSIM = 119
If bc = "eMn56Av" Then BOD2CAPSIM = 118
If bc = "eMn82A" Then BOD2CAPSIM = 119
If bc = "eMn82Ap" Then BOD2CAPSIM = 119
If bc = "eMn82C" Then BOD2CAPSIM = 119
If bc = "eMn82Cp" Then BOD2CAPSIM = 119
If bc = "eMn86A" Then BOD2CAPSIM = 117
If bc = "eMn86Av" Then BOD2CAPSIM = 118
If bc = "eMn86C" Then BOD2CAPSIM = 117
If bc = "eMn86Cv" Then BOD2CAPSIM = 118
If bc = "eMn86Cw" Then BOD2CAPSIM = 117
If bc = "eMo20A" Then BOD2CAPSIM = 119
If bc = "eMo20Ap" Then BOD2CAPSIM = 119
If bc = "eMo80A" Then BOD2CAPSIM = 116
If bc = "eMo80Ap" Then BOD2CAPSIM = 119
If bc = "eMo80C" Then BOD2CAPSIM = 116
If bc = "eMo80Cv" Then BOD2CAPSIM = 118
If bc = "eMOb72" Then BOD2CAPSIM = 119
If bc = "eMOb75" Then BOD2CAPSIM = 116
If bc = "eMOo05" Then BOD2CAPSIM = 115
If bc = "eMv41C" Then BOD2CAPSIM = 118
If bc = "eMv51A" Then BOD2CAPSIM = 118
If bc = "eMv61C" Then BOD2CAPSIM = 118
If bc = "eMv61Cp" Then BOD2CAPSIM = 118
If bc = "eMv81A" Then BOD2CAPSIM = 118
If bc = "eMv81Ap" Then BOD2CAPSIM = 118
If bc = "epMn55A" Then BOD2CAPSIM = 115
If bc = "epMn85A" Then BOD2CAPSIM = 116
If bc = "epMo50" Then BOD2CAPSIM = 115
If bc = "epMo80" Then BOD2CAPSIM = 116
If bc = "epMv81" Then BOD2CAPSIM = 118
If bc = "epRn56" Then BOD2CAPSIM = 117
If bc = "epRn59" Then BOD2CAPSIM = 119
If bc = "epRn86" Then BOD2CAPSIM = 117
If bc = "eRn45A" Then BOD2CAPSIM = 117
If bc = "eRn46A" Then BOD2CAPSIM = 117
If bc = "eRn46Av" Then BOD2CAPSIM = 118
If bc = "eRn47C" Then BOD2CAPSIM = 117
If bc = "eRn52A" Then BOD2CAPSIM = 119
If bc = "eRn66A" Then BOD2CAPSIM = 117
If bc = "eRn66Av" Then BOD2CAPSIM = 118
If bc = "eRn82A" Then BOD2CAPSIM = 119
If bc = "eRn94C" Then BOD2CAPSIM = 117
If bc = "eRn95A" Then BOD2CAPSIM = 116
If bc = "eRn95Av" Then BOD2CAPSIM = 118
If bc = "eRo40A" Then BOD2CAPSIM = 117
If bc = "eRv01A" Then BOD2CAPSIM = 118
If bc = "eRv01C" Then BOD2CAPSIM = 118
If bc = "EZ50A" Then BOD2CAPSIM = 107
If bc = "EZ50Av" Then BOD2CAPSIM = 107
If bc = "EZg21" Then BOD2CAPSIM = 112
If bc = "EZg21g" Then BOD2CAPSIM = 112
If bc = "EZg21v" Then BOD2CAPSIM = 112
If bc = "EZg21w" Then BOD2CAPSIM = 112
If bc = "EZg23" Then BOD2CAPSIM = 112
If bc = "EZg23g" Then BOD2CAPSIM = 112
If bc = "EZg23t" Then BOD2CAPSIM = 112
If bc = "EZg23tw" Then BOD2CAPSIM = 112
If bc = "EZg23w" Then BOD2CAPSIM = 112
If bc = "EZg23wg" Then BOD2CAPSIM = 112
If bc = "EZg23wt" Then BOD2CAPSIM = 112
If bc = "EZg30" Then BOD2CAPSIM = 112
If bc = "EZg30g" Then BOD2CAPSIM = 112
If bc = "EZg30v" Then BOD2CAPSIM = 112
If bc = "fABk" Then BOD2CAPSIM = 119
If bc = "fAFk" Then BOD2CAPSIM = 119
If bc = "fAFz" Then BOD2CAPSIM = 113
If bc = "faVc" Then BOD2CAPSIM = 101
If bc = "faVz" Then BOD2CAPSIM = 102
If bc = "faVzt" Then BOD2CAPSIM = 102
If bc = "FG" Then BOD2CAPSIM = 114
If bc = "fHn21" Then BOD2CAPSIM = 109
If bc = "fhVc" Then BOD2CAPSIM = 101
If bc = "fhVd" Then BOD2CAPSIM = 101
If bc = "fhVz" Then BOD2CAPSIM = 102
If bc = "fiVc" Then BOD2CAPSIM = 105
If bc = "fiVz" Then BOD2CAPSIM = 105
If bc = "fiWp" Then BOD2CAPSIM = 105
If bc = "fiWz" Then BOD2CAPSIM = 105
If bc = "FK" Then BOD2CAPSIM = 121
If bc = "FKk" Then BOD2CAPSIM = 121
If bc = "fkpZg23" Then BOD2CAPSIM = 119
If bc = "fkpZg23g" Then BOD2CAPSIM = 120
If bc = "fkpZg23t" Then BOD2CAPSIM = 119
If bc = "fKRn1" Then BOD2CAPSIM = 119
If bc = "fKRn1g" Then BOD2CAPSIM = 120
If bc = "fKRn2g" Then BOD2CAPSIM = 120
If bc = "fKRn8" Then BOD2CAPSIM = 119
If bc = "fKRn8g" Then BOD2CAPSIM = 120
If bc = "fkVc" Then BOD2CAPSIM = 103
If bc = "fkVs" Then BOD2CAPSIM = 103
If bc = "fkVz" Then BOD2CAPSIM = 104
If bc = "fkWz" Then BOD2CAPSIM = 104
If bc = "fkWzg" Then BOD2CAPSIM = 104
If bc = "fkZn21" Then BOD2CAPSIM = 119
If bc = "fkZn23" Then BOD2CAPSIM = 119
If bc = "fkZn23g" Then BOD2CAPSIM = 120
If bc = "fkZn30" Then BOD2CAPSIM = 120
If bc = "fMn56Cp" Then BOD2CAPSIM = 119
If bc = "fMn56Cv" Then BOD2CAPSIM = 118
If bc = "fpLn5" Then BOD2CAPSIM = 121
If bc = "fpRn59" Then BOD2CAPSIM = 119
If bc = "fpRn86" Then BOD2CAPSIM = 117
If bc = "fpVc" Then BOD2CAPSIM = 103
If bc = "fpVs" Then BOD2CAPSIM = 103
If bc = "fpVz" Then BOD2CAPSIM = 104
If bc = "fpZg21" Then BOD2CAPSIM = 109
If bc = "fpZg21g" Then BOD2CAPSIM = 110
If bc = "fpZg23" Then BOD2CAPSIM = 113
If bc = "fpZg23g" Then BOD2CAPSIM = 113
If bc = "fpZg23t" Then BOD2CAPSIM = 111
If bc = "fpZg23x" Then BOD2CAPSIM = 111
If bc = "fpZn21" Then BOD2CAPSIM = 109
If bc = "fpZn23tg" Then BOD2CAPSIM = 111
If bc = "fRn15C" Then BOD2CAPSIM = 115
If bc = "fRn62C" Then BOD2CAPSIM = 119
If bc = "fRn62Cg" Then BOD2CAPSIM = 120
If bc = "fRn95C" Then BOD2CAPSIM = 116
If bc = "fRo60C" Then BOD2CAPSIM = 116
If bc = "fRv01C" Then BOD2CAPSIM = 118
If bc = "fVc" Then BOD2CAPSIM = 101
If bc = "fvWz" Then BOD2CAPSIM = 102
If bc = "fvWzt" Then BOD2CAPSIM = 102
If bc = "fvWztx" Then BOD2CAPSIM = 102
If bc = "fVz" Then BOD2CAPSIM = 102
If bc = "fZn21" Then BOD2CAPSIM = 107
If bc = "fZn21g" Then BOD2CAPSIM = 107
If bc = "fZn23" Then BOD2CAPSIM = 113
If bc = "fZn23-F" Then BOD2CAPSIM = 113
If bc = "fZn23g" Then BOD2CAPSIM = 113
If bc = "fzVc" Then BOD2CAPSIM = 105
If bc = "fzVz" Then BOD2CAPSIM = 105
If bc = "fzVzt" Then BOD2CAPSIM = 105
If bc = "fzWp" Then BOD2CAPSIM = 105
If bc = "fzWz" Then BOD2CAPSIM = 105
If bc = "fzWzt" Then BOD2CAPSIM = 105
If bc = "gbEZ21" Then BOD2CAPSIM = 112
If bc = "gbEZ30" Then BOD2CAPSIM = 112
If bc = "gcHd30" Then BOD2CAPSIM = 114
If bc = "gcHn21" Then BOD2CAPSIM = 109
If bc = "gcHn30" Then BOD2CAPSIM = 114
If bc = "gcY21" Then BOD2CAPSIM = 109
If bc = "gcY23" Then BOD2CAPSIM = 113
If bc = "gcY30" Then BOD2CAPSIM = 114
If bc = "gcZd30" Then BOD2CAPSIM = 114
If bc = "gHd21" Then BOD2CAPSIM = 108
If bc = "gHd30" Then BOD2CAPSIM = 114
If bc = "gHn21" Then BOD2CAPSIM = 109
If bc = "gHn21t" Then BOD2CAPSIM = 111
If bc = "gHn21x" Then BOD2CAPSIM = 111
If bc = "gHn23" Then BOD2CAPSIM = 113
If bc = "gHn23x" Then BOD2CAPSIM = 111
If bc = "gHn30" Then BOD2CAPSIM = 114
If bc = "gHn30t" Then BOD2CAPSIM = 114
If bc = "gHn30x" Then BOD2CAPSIM = 114
If bc = "gKRd1" Then BOD2CAPSIM = 119
If bc = "gKRd7" Then BOD2CAPSIM = 119
If bc = "gKRn1" Then BOD2CAPSIM = 119
If bc = "gKRn2" Then BOD2CAPSIM = 119
If bc = "gLd6" Then BOD2CAPSIM = 121
If bc = "gLh6" Then BOD2CAPSIM = 121
If bc = "gMK" Then BOD2CAPSIM = 115
If bc = "gMn15C" Then BOD2CAPSIM = 115
If bc = "gMn25C" Then BOD2CAPSIM = 115
If bc = "gMn25Cv" Then BOD2CAPSIM = 115
If bc = "gMn52C" Then BOD2CAPSIM = 119
If bc = "gMn52Cp" Then BOD2CAPSIM = 119
If bc = "gMn52Cw" Then BOD2CAPSIM = 119
If bc = "gMn53C" Then BOD2CAPSIM = 117
If bc = "gMn53Cp" Then BOD2CAPSIM = 119
If bc = "gMn53Cpx" Then BOD2CAPSIM = 119
If bc = "gMn53Cv" Then BOD2CAPSIM = 118
If bc = "gMn53Cw" Then BOD2CAPSIM = 117
If bc = "gMn53Cwp" Then BOD2CAPSIM = 119
If bc = "gMn58C" Then BOD2CAPSIM = 117
If bc = "gMn58Cv" Then BOD2CAPSIM = 117
If bc = "nkZn50A" Then BOD2CAPSIM = 119
If bc = "gMn82C" Then BOD2CAPSIM = 119
If bc = "gMn83C" Then BOD2CAPSIM = 117
If bc = "gMn83Cp" Then BOD2CAPSIM = 117
If bc = "gMn83Cv" Then BOD2CAPSIM = 118
If bc = "gMn83Cw" Then BOD2CAPSIM = 117
If bc = "gMn83Cwp" Then BOD2CAPSIM = 117
If bc = "gMn85C" Then BOD2CAPSIM = 116
If bc = "gMn85Cv" Then BOD2CAPSIM = 118
If bc = "gMn85Cwl" Then BOD2CAPSIM = 116
If bc = "gMn88C" Then BOD2CAPSIM = 117
If bc = "gMn88Cl" Then BOD2CAPSIM = 117
If bc = "gMn88Clv" Then BOD2CAPSIM = 118
If bc = "gMn88Cv" Then BOD2CAPSIM = 118
If bc = "gMn88Cw" Then BOD2CAPSIM = 117
If bc = "gpZg23x" Then BOD2CAPSIM = 111
If bc = "gpZg30" Then BOD2CAPSIM = 114
If bc = "gpZn21" Then BOD2CAPSIM = 109
If bc = "gpZn21x" Then BOD2CAPSIM = 111
If bc = "gpZn23x" Then BOD2CAPSIM = 111
If bc = "gpZn30" Then BOD2CAPSIM = 114
If bc = "gRd10A" Then BOD2CAPSIM = 119
If bc = "gRn15A" Then BOD2CAPSIM = 119
If bc = "gRn94Cv" Then BOD2CAPSIM = 117
If bc = "gtZd30" Then BOD2CAPSIM = 114
If bc = "gvWp" Then BOD2CAPSIM = 102
If bc = "gY21" Then BOD2CAPSIM = 109
If bc = "gY21g" Then BOD2CAPSIM = 109
If bc = "gY23" Then BOD2CAPSIM = 113
If bc = "gY30" Then BOD2CAPSIM = 114
If bc = "gY30-F" Then BOD2CAPSIM = 114
If bc = "gY30-G" Then BOD2CAPSIM = 114
If bc = "gZb30" Then BOD2CAPSIM = 114
If bc = "gZd21" Then BOD2CAPSIM = 107
If bc = "gZd30" Then BOD2CAPSIM = 114
If bc = "gzEZ21" Then BOD2CAPSIM = 112
If bc = "gzEZ23" Then BOD2CAPSIM = 112
If bc = "gzEZ30" Then BOD2CAPSIM = 112
If bc = "gZn30" Then BOD2CAPSIM = 114
If bc = "Hd21" Then BOD2CAPSIM = 108
If bc = "Hd21g" Then BOD2CAPSIM = 108
If bc = "Hd21x" Then BOD2CAPSIM = 108
If bc = "Hd23" Then BOD2CAPSIM = 113
If bc = "Hd23g" Then BOD2CAPSIM = 110
If bc = "Hd23x" Then BOD2CAPSIM = 111
If bc = "Hd30" Then BOD2CAPSIM = 114
If bc = "Hd30g" Then BOD2CAPSIM = 114
If bc = "hEV" Then BOD2CAPSIM = 101
If bc = "Hn21" Then BOD2CAPSIM = 109
If bc = "Hn21-F" Then BOD2CAPSIM = 109
If bc = "Hn21g" Then BOD2CAPSIM = 110
If bc = "Hn21gx" Then BOD2CAPSIM = 110
If bc = "Hn21t" Then BOD2CAPSIM = 111
If bc = "Hn21v" Then BOD2CAPSIM = 109
If bc = "Hn21w" Then BOD2CAPSIM = 109
If bc = "Hn21wg" Then BOD2CAPSIM = 109
If bc = "Hn21x" Then BOD2CAPSIM = 111
If bc = "Hn21x-F" Then BOD2CAPSIM = 111
If bc = "Hn21xg" Then BOD2CAPSIM = 111
If bc = "Hn23" Then BOD2CAPSIM = 113
If bc = "Hn23-F" Then BOD2CAPSIM = 113
If bc = "Hn23g" Then BOD2CAPSIM = 110
If bc = "Hn23t" Then BOD2CAPSIM = 111
If bc = "Hn23x" Then BOD2CAPSIM = 111
If bc = "Hn23x-F" Then BOD2CAPSIM = 111
If bc = "Hn23xg" Then BOD2CAPSIM = 111
If bc = "Hn30" Then BOD2CAPSIM = 114
If bc = "Hn30g" Then BOD2CAPSIM = 114
If bc = "Hn30x" Then BOD2CAPSIM = 114
If bc = "hRd10A" Then BOD2CAPSIM = 119
If bc = "hRd10C" Then BOD2CAPSIM = 119
If bc = "hRd90A" Then BOD2CAPSIM = 116
If bc = "hVb" Then BOD2CAPSIM = 101
If bc = "hVc" Then BOD2CAPSIM = 101
If bc = "hVcc" Then BOD2CAPSIM = 101
If bc = "hVd" Then BOD2CAPSIM = 101
If bc = "hVk" Then BOD2CAPSIM = 106
If bc = "hVkl" Then BOD2CAPSIM = 106
If bc = "hVr" Then BOD2CAPSIM = 101
If bc = "hVs" Then BOD2CAPSIM = 101
If bc = "hVsc" Then BOD2CAPSIM = 101
If bc = "hVz" Then BOD2CAPSIM = 102
If bc = "hVzc" Then BOD2CAPSIM = 102
If bc = "hVzg" Then BOD2CAPSIM = 102
If bc = "hVzx" Then BOD2CAPSIM = 102
If bc = "hZd20A" Then BOD2CAPSIM = 107
If bc = "iVc" Then BOD2CAPSIM = 105
If bc = "iVp" Then BOD2CAPSIM = 105
If bc = "iVpc" Then BOD2CAPSIM = 105
If bc = "iVpg" Then BOD2CAPSIM = 105
If bc = "iVpt" Then BOD2CAPSIM = 105
If bc = "iVpx" Then BOD2CAPSIM = 105
If bc = "iVs" Then BOD2CAPSIM = 105
If bc = "iVz" Then BOD2CAPSIM = 105
If bc = "iVzg" Then BOD2CAPSIM = 105
If bc = "iVzt" Then BOD2CAPSIM = 105
If bc = "iVzx" Then BOD2CAPSIM = 105
If bc = "iWp" Then BOD2CAPSIM = 105
If bc = "iWpc" Then BOD2CAPSIM = 105
If bc = "iWpg" Then BOD2CAPSIM = 105
If bc = "iWpt" Then BOD2CAPSIM = 105
If bc = "iWpx" Then BOD2CAPSIM = 105
If bc = "iWz" Then BOD2CAPSIM = 105
If bc = "iWzt" Then BOD2CAPSIM = 105
If bc = "iWzx" Then BOD2CAPSIM = 105
If bc = "kcHn21" Then BOD2CAPSIM = 119
If bc = "kgpZg30" Then BOD2CAPSIM = 120
If bc = "kHn21" Then BOD2CAPSIM = 119
If bc = "kHn21g" Then BOD2CAPSIM = 120
If bc = "kHn21x" Then BOD2CAPSIM = 119
If bc = "kHn23" Then BOD2CAPSIM = 119
If bc = "kHn23x" Then BOD2CAPSIM = 119
If bc = "kHn30" Then BOD2CAPSIM = 120
If bc = "KK" Then BOD2CAPSIM = 121
If bc = "KM" Then BOD2CAPSIM = 121
If bc = "kMn43C" Then BOD2CAPSIM = 117
If bc = "kMn43Cp" Then BOD2CAPSIM = 117
If bc = "kMn43Cpx" Then BOD2CAPSIM = 117
If bc = "kMn43Cv" Then BOD2CAPSIM = 118
If bc = "kMn43Cwp" Then BOD2CAPSIM = 117
If bc = "kMn48C" Then BOD2CAPSIM = 117
If bc = "kMn48Cl" Then BOD2CAPSIM = 117
If bc = "kMn48Clv" Then BOD2CAPSIM = 118
If bc = "kMn48Cv" Then BOD2CAPSIM = 118
If bc = "kMn48Cvl" Then BOD2CAPSIM = 118
If bc = "kMn48Cw" Then BOD2CAPSIM = 117
If bc = "kMn63C" Then BOD2CAPSIM = 117
If bc = "kMn63Cp" Then BOD2CAPSIM = 119
If bc = "kMn63Cpx" Then BOD2CAPSIM = 119
If bc = "kMn63Cv" Then BOD2CAPSIM = 118
If bc = "kMn63Cwp" Then BOD2CAPSIM = 119
If bc = "kMn68C" Then BOD2CAPSIM = 117
If bc = "kMn68Cl" Then BOD2CAPSIM = 117
If bc = "kMn68Cv" Then BOD2CAPSIM = 118
If bc = "kpZg20A" Then BOD2CAPSIM = 119
If bc = "kpZg21" Then BOD2CAPSIM = 119
If bc = "kpZg21g" Then BOD2CAPSIM = 120
If bc = "kpZg23" Then BOD2CAPSIM = 119
If bc = "kpZg23g" Then BOD2CAPSIM = 120
If bc = "kpZg23t" Then BOD2CAPSIM = 119
If bc = "kpZg23x" Then BOD2CAPSIM = 119
If bc = "kpZn21" Then BOD2CAPSIM = 119
If bc = "kpZn21g" Then BOD2CAPSIM = 120
If bc = "kpZn23" Then BOD2CAPSIM = 119
If bc = "kpZn23x" Then BOD2CAPSIM = 119
If bc = "KRd1" Then BOD2CAPSIM = 119
If bc = "KRd1g" Then BOD2CAPSIM = 120
If bc = "KRd7" Then BOD2CAPSIM = 119
If bc = "KRd7g" Then BOD2CAPSIM = 120
If bc = "KRn1" Then BOD2CAPSIM = 119
If bc = "KRn1g" Then BOD2CAPSIM = 120
If bc = "KRn2" Then BOD2CAPSIM = 119
If bc = "KRn2g" Then BOD2CAPSIM = 120
If bc = "KRn2w" Then BOD2CAPSIM = 119
If bc = "KRn8" Then BOD2CAPSIM = 119
If bc = "KRn8g" Then BOD2CAPSIM = 120
If bc = "KS" Then BOD2CAPSIM = 115
If bc = "kSn13A" Then BOD2CAPSIM = 119
If bc = "kSn13Av" Then BOD2CAPSIM = 119
If bc = "kSn13Aw" Then BOD2CAPSIM = 119
If bc = "kSn14A" Then BOD2CAPSIM = 119
If bc = "kSn14Ap" Then BOD2CAPSIM = 119
If bc = "kSn14Av" Then BOD2CAPSIM = 119
If bc = "kSn14Aw" Then BOD2CAPSIM = 119
If bc = "kSn14Awp" Then BOD2CAPSIM = 119
If bc = "KT" Then BOD2CAPSIM = 115
If bc = "kVb" Then BOD2CAPSIM = 103
If bc = "kVc" Then BOD2CAPSIM = 103
If bc = "kVcc" Then BOD2CAPSIM = 103
If bc = "kVd" Then BOD2CAPSIM = 103
If bc = "kVk" Then BOD2CAPSIM = 106
If bc = "kVr" Then BOD2CAPSIM = 103
If bc = "kVs" Then BOD2CAPSIM = 103
If bc = "kVsc" Then BOD2CAPSIM = 103
If bc = "kVz" Then BOD2CAPSIM = 104
If bc = "kVzc" Then BOD2CAPSIM = 104
If bc = "kVzx" Then BOD2CAPSIM = 104
If bc = "kWp" Then BOD2CAPSIM = 104
If bc = "kWpg" Then BOD2CAPSIM = 104
If bc = "kWpx" Then BOD2CAPSIM = 104
If bc = "kWz" Then BOD2CAPSIM = 104
If bc = "kWzg" Then BOD2CAPSIM = 104
If bc = "kWzx" Then BOD2CAPSIM = 104
If bc = "KX" Then BOD2CAPSIM = 115
If bc = "kZb21" Then BOD2CAPSIM = 119
If bc = "kZb23" Then BOD2CAPSIM = 119
If bc = "kZn10A" Then BOD2CAPSIM = 119
If bc = "kZn10Av" Then BOD2CAPSIM = 119
If bc = "kZn21" Then BOD2CAPSIM = 119
If bc = "kZn21g" Then BOD2CAPSIM = 120
If bc = "kZn21p" Then BOD2CAPSIM = 119
If bc = "kZn21r" Then BOD2CAPSIM = 119
If bc = "kZn21w" Then BOD2CAPSIM = 119
If bc = "kZn21x" Then BOD2CAPSIM = 119
If bc = "kZn23" Then BOD2CAPSIM = 119
If bc = "kZn30" Then BOD2CAPSIM = 120
If bc = "kZn30A" Then BOD2CAPSIM = 120
If bc = "kZn30Ar" Then BOD2CAPSIM = 120
If bc = "kZn30x" Then BOD2CAPSIM = 120
If bc = "kZn40A" Then BOD2CAPSIM = 119
If bc = "kZn40Ap" Then BOD2CAPSIM = 119
If bc = "kZn40Av" Then BOD2CAPSIM = 119
If bc = "kZn50A" Then BOD2CAPSIM = 119
If bc = "kZn50Ap" Then BOD2CAPSIM = 119
If bc = "kZn50Ar" Then BOD2CAPSIM = 119
If bc = "Ld5" Then BOD2CAPSIM = 121
If bc = "Ld5g" Then BOD2CAPSIM = 121
If bc = "Ld5m" Then BOD2CAPSIM = 121
If bc = "Ld5t" Then BOD2CAPSIM = 121
If bc = "Ld6" Then BOD2CAPSIM = 121
If bc = "Ld6a" Then BOD2CAPSIM = 121
If bc = "Ld6g" Then BOD2CAPSIM = 121
If bc = "Ld6k" Then BOD2CAPSIM = 121
If bc = "Ld6m" Then BOD2CAPSIM = 121
If bc = "Ld6s" Then BOD2CAPSIM = 121
If bc = "Ld6t" Then BOD2CAPSIM = 121
If bc = "Ldd5" Then BOD2CAPSIM = 121
If bc = "Ldd5g" Then BOD2CAPSIM = 121
If bc = "Ldd6" Then BOD2CAPSIM = 121
If bc = "Ldh5" Then BOD2CAPSIM = 121
If bc = "Ldh5g" Then BOD2CAPSIM = 121
If bc = "Ldh5t" Then BOD2CAPSIM = 121
If bc = "Ldh6" Then BOD2CAPSIM = 121
If bc = "Ldh6m" Then BOD2CAPSIM = 121
If bc = "lFG" Then BOD2CAPSIM = 114
If bc = "lFK" Then BOD2CAPSIM = 121
If bc = "lFKk" Then BOD2CAPSIM = 121
If bc = "Lh5" Then BOD2CAPSIM = 121
If bc = "Lh5g" Then BOD2CAPSIM = 121
If bc = "Lh6" Then BOD2CAPSIM = 121
If bc = "Lh6g" Then BOD2CAPSIM = 121
If bc = "Lh6s" Then BOD2CAPSIM = 121
If bc = "lKK" Then BOD2CAPSIM = 116
If bc = "lKM" Then BOD2CAPSIM = 116
If bc = "lKRd7" Then BOD2CAPSIM = 119
If bc = "lKS" Then BOD2CAPSIM = 121
If bc = "Ln5" Then BOD2CAPSIM = 121
If bc = "Ln5g" Then BOD2CAPSIM = 121
If bc = "Ln5m" Then BOD2CAPSIM = 121
If bc = "Ln5t" Then BOD2CAPSIM = 121
If bc = "Ln6" Then BOD2CAPSIM = 121
If bc = "Ln6a" Then BOD2CAPSIM = 121
If bc = "Ln6m" Then BOD2CAPSIM = 121
If bc = "Ln6t" Then BOD2CAPSIM = 121
If bc = "Lnd5" Then BOD2CAPSIM = 121
If bc = "Lnd5g" Then BOD2CAPSIM = 121
If bc = "Lnd5m" Then BOD2CAPSIM = 121
If bc = "Lnd5t" Then BOD2CAPSIM = 121
If bc = "Lnd6" Then BOD2CAPSIM = 121
If bc = "Lnd6v" Then BOD2CAPSIM = 121
If bc = "Lnh6" Then BOD2CAPSIM = 121
If bc = "MA" Then BOD2CAPSIM = 116
If bc = "mcY23" Then BOD2CAPSIM = 113
If bc = "mcY23x" Then BOD2CAPSIM = 111
If bc = "mHd23" Then BOD2CAPSIM = 113
If bc = "mHn21x" Then BOD2CAPSIM = 111
If bc = "mHn23x" Then BOD2CAPSIM = 111
If bc = "MK" Then BOD2CAPSIM = 116
If bc = "mKK" Then BOD2CAPSIM = 116
If bc = "mKRd7" Then BOD2CAPSIM = 119
If bc = "mKX" Then BOD2CAPSIM = 115
If bc = "mLd6s" Then BOD2CAPSIM = 121
If bc = "mLh6s" Then BOD2CAPSIM = 121
If bc = "Mn12A" Then BOD2CAPSIM = 119
If bc = "Mn12Ap" Then BOD2CAPSIM = 119
If bc = "Mn12Av" Then BOD2CAPSIM = 119
If bc = "Mn12Awp" Then BOD2CAPSIM = 119
If bc = "Mn15A" Then BOD2CAPSIM = 115
If bc = "Mn15Ap" Then BOD2CAPSIM = 119
If bc = "Mn15Av" Then BOD2CAPSIM = 118
If bc = "Mn15Aw" Then BOD2CAPSIM = 115
If bc = "Mn15Awp" Then BOD2CAPSIM = 119
If bc = "Mn15C" Then BOD2CAPSIM = 115
If bc = "Mn15Clv" Then BOD2CAPSIM = 118
If bc = "Mn15Cv" Then BOD2CAPSIM = 118
If bc = "Mn15Cw" Then BOD2CAPSIM = 115
If bc = "Mn22A" Then BOD2CAPSIM = 119
If bc = "Mn22Alv" Then BOD2CAPSIM = 115
If bc = "Mn22Ap" Then BOD2CAPSIM = 119
If bc = "Mn22Av" Then BOD2CAPSIM = 115
If bc = "Mn22Aw" Then BOD2CAPSIM = 119
If bc = "Mn22Awp" Then BOD2CAPSIM = 119
If bc = "Mn22Ax" Then BOD2CAPSIM = 119
If bc = "Mn25A" Then BOD2CAPSIM = 115
If bc = "Mn25Alv" Then BOD2CAPSIM = 115
If bc = "Mn25Ap" Then BOD2CAPSIM = 119
If bc = "Mn25Av" Then BOD2CAPSIM = 118
If bc = "Mn25Aw" Then BOD2CAPSIM = 115
If bc = "Mn25Awp" Then BOD2CAPSIM = 119
If bc = "Mn25C" Then BOD2CAPSIM = 115
If bc = "Mn25Cp" Then BOD2CAPSIM = 119
If bc = "Mn25Cv" Then BOD2CAPSIM = 118
If bc = "Mn25Cw" Then BOD2CAPSIM = 115
If bc = "Mn35A" Then BOD2CAPSIM = 116
If bc = "Mn35Ap" Then BOD2CAPSIM = 119
If bc = "Mn35Av" Then BOD2CAPSIM = 118
If bc = "Mn35Aw" Then BOD2CAPSIM = 116
If bc = "Mn35Awp" Then BOD2CAPSIM = 119
If bc = "Mn35Ax" Then BOD2CAPSIM = 116
If bc = "Mn45A" Then BOD2CAPSIM = 117
If bc = "Mn45Ap" Then BOD2CAPSIM = 119
If bc = "Mn45Av" Then BOD2CAPSIM = 118
If bc = "Mn52C" Then BOD2CAPSIM = 119
If bc = "Mn52Cp" Then BOD2CAPSIM = 119
If bc = "Mn52Cpx" Then BOD2CAPSIM = 119
If bc = "Mn52Cwp" Then BOD2CAPSIM = 119
If bc = "Mn52Cx" Then BOD2CAPSIM = 119
If bc = "Mn56A" Then BOD2CAPSIM = 117
If bc = "Mn56Ap" Then BOD2CAPSIM = 119
If bc = "Mn56Av" Then BOD2CAPSIM = 118
If bc = "Mn56Aw" Then BOD2CAPSIM = 117
If bc = "Mn56C" Then BOD2CAPSIM = 117
If bc = "Mn56Cp" Then BOD2CAPSIM = 119
If bc = "Mn56Cv" Then BOD2CAPSIM = 118
If bc = "Mn56Cwp" Then BOD2CAPSIM = 119
If bc = "Mn82A" Then BOD2CAPSIM = 119
If bc = "Mn82Ap" Then BOD2CAPSIM = 119
If bc = "Mn82C" Then BOD2CAPSIM = 119
If bc = "Mn82Cp" Then BOD2CAPSIM = 119
If bc = "Mn82Cpx" Then BOD2CAPSIM = 119
If bc = "Mn82Cwp" Then BOD2CAPSIM = 119
If bc = "Mn85C" Then BOD2CAPSIM = 116
If bc = "Mn85Clwp" Then BOD2CAPSIM = 119
If bc = "Mn85Cp" Then BOD2CAPSIM = 119
If bc = "Mn85Cv" Then BOD2CAPSIM = 118
If bc = "Mn85Cw" Then BOD2CAPSIM = 116
If bc = "Mn85Cwp" Then BOD2CAPSIM = 119
If bc = "Mn86A" Then BOD2CAPSIM = 117
If bc = "Mn86Al" Then BOD2CAPSIM = 117
If bc = "Mn86Av" Then BOD2CAPSIM = 118
If bc = "Mn86Aw" Then BOD2CAPSIM = 117
If bc = "Mn86C" Then BOD2CAPSIM = 117
If bc = "Mn86Cl" Then BOD2CAPSIM = 117
If bc = "Mn86Clv" Then BOD2CAPSIM = 117
If bc = "Mn86Clw" Then BOD2CAPSIM = 117
If bc = "Mn86Clwp" Then BOD2CAPSIM = 119
If bc = "Mn86Cp" Then BOD2CAPSIM = 119
If bc = "Mn86Cv" Then BOD2CAPSIM = 118
If bc = "Mn86Cw" Then BOD2CAPSIM = 117
If bc = "Mn86Cwp" Then BOD2CAPSIM = 119
If bc = "Mo10A" Then BOD2CAPSIM = 115
If bc = "Mo10Av" Then BOD2CAPSIM = 115
If bc = "Mo20A" Then BOD2CAPSIM = 115
If bc = "Mo20Av" Then BOD2CAPSIM = 115
If bc = "Mo50C" Then BOD2CAPSIM = 115
If bc = "Mo80A" Then BOD2CAPSIM = 116
If bc = "Mo80Ap" Then BOD2CAPSIM = 119
If bc = "Mo80Av" Then BOD2CAPSIM = 118
If bc = "Mo80C" Then BOD2CAPSIM = 116
If bc = "Mo80Cl" Then BOD2CAPSIM = 116
If bc = "Mo80Cp" Then BOD2CAPSIM = 119
If bc = "Mo80Cv" Then BOD2CAPSIM = 118
If bc = "Mo80Cvl" Then BOD2CAPSIM = 118
If bc = "Mo80Cw" Then BOD2CAPSIM = 116
If bc = "Mo80Cwp" Then BOD2CAPSIM = 119
If bc = "MOb12" Then BOD2CAPSIM = 119
If bc = "MOb15" Then BOD2CAPSIM = 115
If bc = "MOb72" Then BOD2CAPSIM = 119
If bc = "MOb75" Then BOD2CAPSIM = 116
If bc = "MOo02" Then BOD2CAPSIM = 119
If bc = "MOo02v" Then BOD2CAPSIM = 119
If bc = "MOo05" Then BOD2CAPSIM = 115
If bc = "Mv41C" Then BOD2CAPSIM = 118
If bc = "Mv41Cl" Then BOD2CAPSIM = 118
If bc = "Mv41Cp" Then BOD2CAPSIM = 118
If bc = "Mv41Cv" Then BOD2CAPSIM = 118
If bc = "Mv51A" Then BOD2CAPSIM = 118
If bc = "Mv51Al" Then BOD2CAPSIM = 118
If bc = "Mv51Ap" Then BOD2CAPSIM = 118
If bc = "Mv61C" Then BOD2CAPSIM = 118
If bc = "Mv61Cl" Then BOD2CAPSIM = 118
If bc = "Mv61Cp" Then BOD2CAPSIM = 118
If bc = "Mv81A" Then BOD2CAPSIM = 118
If bc = "Mv81Al" Then BOD2CAPSIM = 118
If bc = "Mv81Ap" Then BOD2CAPSIM = 118
If bc = "mY23" Then BOD2CAPSIM = 113
If bc = "mY23x" Then BOD2CAPSIM = 111
If bc = "mZb23x" Then BOD2CAPSIM = 111
If bc = "MZk" Then BOD2CAPSIM = 121
If bc = "MZz" Then BOD2CAPSIM = 107
If bc = "nAO" Then BOD2CAPSIM = 119
If bc = "nkZn21" Then BOD2CAPSIM = 119
If bc = "nkZn50Ab" Then BOD2CAPSIM = 119
If bc = "nMn15A" Then BOD2CAPSIM = 115
If bc = "nMn15Av" Then BOD2CAPSIM = 115
If bc = "nMo10A" Then BOD2CAPSIM = 115
If bc = "nMo10Av" Then BOD2CAPSIM = 118
If bc = "nMo80A" Then BOD2CAPSIM = 116
If bc = "nMo80Aw" Then BOD2CAPSIM = 116
If bc = "nMv61C" Then BOD2CAPSIM = 118
If bc = "npMo50l" Then BOD2CAPSIM = 115
If bc = "npMo80l" Then BOD2CAPSIM = 116
If bc = "nSn13A" Then BOD2CAPSIM = 113
If bc = "nSn13Av" Then BOD2CAPSIM = 113
If bc = "nvWz" Then BOD2CAPSIM = 102
If bc = "nZn21" Then BOD2CAPSIM = 107
If bc = "nZn40A" Then BOD2CAPSIM = 107
If bc = "nZn50A" Then BOD2CAPSIM = 107
If bc = "nZn50Ab" Then BOD2CAPSIM = 107
If bc = "ohVb" Then BOD2CAPSIM = 101
If bc = "ohVc" Then BOD2CAPSIM = 101
If bc = "ohVk" Then BOD2CAPSIM = 106
If bc = "ohVs" Then BOD2CAPSIM = 101
If bc = "opVb" Then BOD2CAPSIM = 103
If bc = "opVc" Then BOD2CAPSIM = 103
If bc = "opVk" Then BOD2CAPSIM = 106
If bc = "opVs" Then BOD2CAPSIM = 103
If bc = "pKRn1" Then BOD2CAPSIM = 119
If bc = "pKRn1g" Then BOD2CAPSIM = 120
If bc = "pKRn2" Then BOD2CAPSIM = 119
If bc = "pKRn2g" Then BOD2CAPSIM = 120
If bc = "pLn5" Then BOD2CAPSIM = 121
If bc = "pLn5g" Then BOD2CAPSIM = 121
If bc = "pMn52A" Then BOD2CAPSIM = 119
If bc = "pMn52C" Then BOD2CAPSIM = 119
If bc = "pMn52Cp" Then BOD2CAPSIM = 119
If bc = "pMn55A" Then BOD2CAPSIM = 115
If bc = "pMn55Av" Then BOD2CAPSIM = 118
If bc = "pMn55Aw" Then BOD2CAPSIM = 115
If bc = "pMn55C" Then BOD2CAPSIM = 115
If bc = "pMn55Cp" Then BOD2CAPSIM = 119
If bc = "pMn56C" Then BOD2CAPSIM = 117
If bc = "pMn56Cl" Then BOD2CAPSIM = 117
If bc = "pMn82A" Then BOD2CAPSIM = 119
If bc = "pMn82C" Then BOD2CAPSIM = 119
If bc = "pMn85A" Then BOD2CAPSIM = 116
If bc = "pMn85Aw" Then BOD2CAPSIM = 116
If bc = "pMn85C" Then BOD2CAPSIM = 116
If bc = "pMn85Cv" Then BOD2CAPSIM = 118
If bc = "pMn86C" Then BOD2CAPSIM = 117
If bc = "pMn86Cl" Then BOD2CAPSIM = 117
If bc = "pMn86Cv" Then BOD2CAPSIM = 118
If bc = "pMn86Cw" Then BOD2CAPSIM = 117
If bc = "pMn86Cwl" Then BOD2CAPSIM = 117
If bc = "pMo50" Then BOD2CAPSIM = 115
If bc = "pMo50l" Then BOD2CAPSIM = 115
If bc = "pMo50w" Then BOD2CAPSIM = 115
If bc = "pMo80" Then BOD2CAPSIM = 116
If bc = "pMo80l" Then BOD2CAPSIM = 116
If bc = "pMo80v" Then BOD2CAPSIM = 118
If bc = "pMv51" Then BOD2CAPSIM = 118
If bc = "pMv81" Then BOD2CAPSIM = 118
If bc = "pMv81l" Then BOD2CAPSIM = 118
If bc = "pMv81p" Then BOD2CAPSIM = 118
If bc = "pRn56" Then BOD2CAPSIM = 119
If bc = "pRn56p" Then BOD2CAPSIM = 119
If bc = "pRn56v" Then BOD2CAPSIM = 118
If bc = "pRn56wp" Then BOD2CAPSIM = 119
If bc = "pRn59" Then BOD2CAPSIM = 119
If bc = "pRn59p" Then BOD2CAPSIM = 119
If bc = "pRn59t" Then BOD2CAPSIM = 119
If bc = "pRn59w" Then BOD2CAPSIM = 119
If bc = "pRn86" Then BOD2CAPSIM = 117
If bc = "pRn86p" Then BOD2CAPSIM = 119
If bc = "pRn86t" Then BOD2CAPSIM = 117
If bc = "pRn86v" Then BOD2CAPSIM = 118
If bc = "pRn86w" Then BOD2CAPSIM = 117
If bc = "pRn86wp" Then BOD2CAPSIM = 119
If bc = "pRn89" Then BOD2CAPSIM = 118
If bc = "pRn89v" Then BOD2CAPSIM = 118
If bc = "pRv81" Then BOD2CAPSIM = 118
If bc = "pVb" Then BOD2CAPSIM = 103
If bc = "pVc" Then BOD2CAPSIM = 103
If bc = "pVcc" Then BOD2CAPSIM = 103
If bc = "pVd" Then BOD2CAPSIM = 103
If bc = "pVk" Then BOD2CAPSIM = 106
If bc = "pVr" Then BOD2CAPSIM = 103
If bc = "pVs" Then BOD2CAPSIM = 103
If bc = "pVsc" Then BOD2CAPSIM = 103
If bc = "pVsl" Then BOD2CAPSIM = 103
If bc = "pVz" Then BOD2CAPSIM = 104
If bc = "pVzx" Then BOD2CAPSIM = 104
If bc = "pZg20A" Then BOD2CAPSIM = 107
If bc = "pZg20Ar" Then BOD2CAPSIM = 107
If bc = "pZg21" Then BOD2CAPSIM = 109
If bc = "pZg21g" Then BOD2CAPSIM = 110
If bc = "pZg21r" Then BOD2CAPSIM = 111
If bc = "pZg21t" Then BOD2CAPSIM = 111
If bc = "pZg21w" Then BOD2CAPSIM = 109
If bc = "pZg21x" Then BOD2CAPSIM = 111
If bc = "pZg23" Then BOD2CAPSIM = 113
If bc = "pZg23g" Then BOD2CAPSIM = 113
If bc = "pZg23r" Then BOD2CAPSIM = 113
If bc = "pZg23t" Then BOD2CAPSIM = 111
If bc = "pZg23w" Then BOD2CAPSIM = 113
If bc = "pZg23x" Then BOD2CAPSIM = 111
If bc = "pZg30" Then BOD2CAPSIM = 114
If bc = "pZg30p" Then BOD2CAPSIM = 114
If bc = "pZg30r" Then BOD2CAPSIM = 114
If bc = "pZg30x" Then BOD2CAPSIM = 114
If bc = "pZn21" Then BOD2CAPSIM = 109
If bc = "pZn21g" Then BOD2CAPSIM = 110
If bc = "pZn21t" Then BOD2CAPSIM = 111
If bc = "pZn21tg" Then BOD2CAPSIM = 109
If bc = "pZn21v" Then BOD2CAPSIM = 109
If bc = "pZn21x" Then BOD2CAPSIM = 111
If bc = "pZn23" Then BOD2CAPSIM = 113
If bc = "pZn23g" Then BOD2CAPSIM = 110
If bc = "pZn23gx" Then BOD2CAPSIM = 110
If bc = "pZn23t" Then BOD2CAPSIM = 111
If bc = "pZn23v" Then BOD2CAPSIM = 113
If bc = "pZn23w" Then BOD2CAPSIM = 113
If bc = "pZn23x" Then BOD2CAPSIM = 111
If bc = "pZn23x-F" Then BOD2CAPSIM = 111
If bc = "pZn30" Then BOD2CAPSIM = 114
If bc = "pZn30g" Then BOD2CAPSIM = 114
If bc = "pZn30r" Then BOD2CAPSIM = 114
If bc = "pZn30w" Then BOD2CAPSIM = 114
If bc = "pZn30x" Then BOD2CAPSIM = 114
If bc = "Rd10A" Then BOD2CAPSIM = 119
If bc = "Rd10Ag" Then BOD2CAPSIM = 119
If bc = "Rd10C" Then BOD2CAPSIM = 119
If bc = "Rd10Cg" Then BOD2CAPSIM = 120
If bc = "Rd10Cm" Then BOD2CAPSIM = 119
If bc = "Rd10Cp" Then BOD2CAPSIM = 119
If bc = "Rd90A" Then BOD2CAPSIM = 116
If bc = "Rd90C" Then BOD2CAPSIM = 116
If bc = "Rd90Cg" Then BOD2CAPSIM = 120
If bc = "Rd90Cm" Then BOD2CAPSIM = 116
If bc = "Rd90Cp" Then BOD2CAPSIM = 119
If bc = "Rn14C" Then BOD2CAPSIM = 117
If bc = "Rn15A" Then BOD2CAPSIM = 115
If bc = "Rn15C" Then BOD2CAPSIM = 115
If bc = "Rn15Cg" Then BOD2CAPSIM = 115
If bc = "Rn15Ct" Then BOD2CAPSIM = 115
If bc = "Rn15Cw" Then BOD2CAPSIM = 115
If bc = "Rn42C" Then BOD2CAPSIM = 119
If bc = "Rn42Cg" Then BOD2CAPSIM = 119
If bc = "Rn42Cp" Then BOD2CAPSIM = 119
If bc = "Rn44C" Then BOD2CAPSIM = 117
If bc = "Rn44Cv" Then BOD2CAPSIM = 118
If bc = "Rn44Cw" Then BOD2CAPSIM = 117
If bc = "Rn45A" Then BOD2CAPSIM = 117
If bc = "Rn45C" Then BOD2CAPSIM = 117
If bc = "Rn46A" Then BOD2CAPSIM = 117
If bc = "Rn46Av" Then BOD2CAPSIM = 118
If bc = "Rn46Aw" Then BOD2CAPSIM = 117
If bc = "Rn47C" Then BOD2CAPSIM = 117
If bc = "Rn47Cg" Then BOD2CAPSIM = 120
If bc = "Rn47Cp" Then BOD2CAPSIM = 119
If bc = "Rn47Cv" Then BOD2CAPSIM = 118
If bc = "Rn47Cw" Then BOD2CAPSIM = 117
If bc = "Rn47Cwp" Then BOD2CAPSIM = 119
If bc = "Rn52A" Then BOD2CAPSIM = 120
If bc = "Rn52Ag" Then BOD2CAPSIM = 120
If bc = "Rn62C" Then BOD2CAPSIM = 119
If bc = "Rn62Cg" Then BOD2CAPSIM = 120
If bc = "Rn62Cp" Then BOD2CAPSIM = 119
If bc = "Rn62Cwp" Then BOD2CAPSIM = 119
If bc = "Rn66A" Then BOD2CAPSIM = 117
If bc = "Rn66Av" Then BOD2CAPSIM = 118
If bc = "Rn67C" Then BOD2CAPSIM = 117
If bc = "Rn67Cg" Then BOD2CAPSIM = 120
If bc = "Rn67Cp" Then BOD2CAPSIM = 119
If bc = "Rn67Cv" Then BOD2CAPSIM = 118
If bc = "Rn67Cwp" Then BOD2CAPSIM = 119
If bc = "Rn82A" Then BOD2CAPSIM = 119
If bc = "Rn82Ag" Then BOD2CAPSIM = 120
If bc = "Rn94C" Then BOD2CAPSIM = 117
If bc = "Rn94Cv" Then BOD2CAPSIM = 118
If bc = "Rn95A" Then BOD2CAPSIM = 116
If bc = "Rn95Av" Then BOD2CAPSIM = 118
If bc = "Rn95C" Then BOD2CAPSIM = 116
If bc = "Rn95Cg" Then BOD2CAPSIM = 120
If bc = "Rn95Cm" Then BOD2CAPSIM = 116
If bc = "Rn95Cp" Then BOD2CAPSIM = 119
If bc = "Ro40A" Then BOD2CAPSIM = 117
If bc = "Ro40Av" Then BOD2CAPSIM = 118
If bc = "Ro40C" Then BOD2CAPSIM = 117
If bc = "Ro40Cv" Then BOD2CAPSIM = 118
If bc = "Ro40Cw" Then BOD2CAPSIM = 117
If bc = "Ro60A" Then BOD2CAPSIM = 116
If bc = "Ro60C" Then BOD2CAPSIM = 116
If bc = "ROb72" Then BOD2CAPSIM = 119
If bc = "ROb75" Then BOD2CAPSIM = 116
If bc = "Rv01A" Then BOD2CAPSIM = 118
If bc = "Rv01C" Then BOD2CAPSIM = 118
If bc = "Rv01Cg" Then BOD2CAPSIM = 118
If bc = "Rv01Cp" Then BOD2CAPSIM = 118
If bc = "saVc" Then BOD2CAPSIM = 101
If bc = "saVz" Then BOD2CAPSIM = 102
If bc = "sHn21" Then BOD2CAPSIM = 109
If bc = "shVz" Then BOD2CAPSIM = 102
If bc = "skVc" Then BOD2CAPSIM = 103
If bc = "skWz" Then BOD2CAPSIM = 104
If bc = "Sn13A" Then BOD2CAPSIM = 113
If bc = "Sn13Ap" Then BOD2CAPSIM = 113
If bc = "Sn13Av" Then BOD2CAPSIM = 113
If bc = "Sn13Aw" Then BOD2CAPSIM = 113
If bc = "Sn13Awp" Then BOD2CAPSIM = 113
If bc = "Sn14A" Then BOD2CAPSIM = 113
If bc = "Sn14Ap" Then BOD2CAPSIM = 113
If bc = "Sn14Av" Then BOD2CAPSIM = 113
If bc = "spVc" Then BOD2CAPSIM = 103
If bc = "spVz" Then BOD2CAPSIM = 104
If bc = "sVc" Then BOD2CAPSIM = 101
If bc = "sVk" Then BOD2CAPSIM = 106
If bc = "sVp" Then BOD2CAPSIM = 102
If bc = "sVs" Then BOD2CAPSIM = 101
If bc = "svWp" Then BOD2CAPSIM = 102
If bc = "svWz" Then BOD2CAPSIM = 102
If bc = "svWzt" Then BOD2CAPSIM = 102
If bc = "sVz" Then BOD2CAPSIM = 102
If bc = "sVzt" Then BOD2CAPSIM = 102
If bc = "sVzx" Then BOD2CAPSIM = 102
If bc = "tZd21" Then BOD2CAPSIM = 107
If bc = "tZd21g" Then BOD2CAPSIM = 110
If bc = "tZd21v" Then BOD2CAPSIM = 107
If bc = "tZd23" Then BOD2CAPSIM = 113
If bc = "tZd30" Then BOD2CAPSIM = 114
If bc = "U4546nr005" Then BOD2CAPSIM = 109 'in omschrijving erachter stond cHn21 (veldpodzol, lemig fijn zand)
If bc = "U4546nr113" Then BOD2CAPSIM = 109 'in de omschrijving erachter stond Hn21 (veldpodzol, zwak lemig fijn zand)
If bc = "U4546nr013" Then BOD2CAPSIM = 109 'in de omschrijving erachter stond Hn21g (veldpodzol, zwak lemig fijn zand)
If bc = "U4546nr127" Then BOD2CAPSIM = 112 'in de omschrijving erachter stond zEZ21 (hoge, zwarte enkeerdgronden, leemarm en zwak lemig fijn zand)
If bc = "U4546nr017" Then BOD2CAPSIM = 112 'in de omschrijving erachter stond pZn21g (gooreerdgronden, lemarm en zwak lemig fijn zand)
If bc = "U4546nr003" Then BOD2CAPSIM = 109 'in de omschrijving erachter stond cHn21 (laarpodzolgronden, leemarm en zwak lemig fijn zand)
If bc = "uWz" Then BOD2CAPSIM = 113 'moerige eerdgronden
If bc = "Vb" Then BOD2CAPSIM = 101
If bc = "Vc" Then BOD2CAPSIM = 101
If bc = "Vd" Then BOD2CAPSIM = 101
If bc = "Vk" Then BOD2CAPSIM = 106
If bc = "Vo" Then BOD2CAPSIM = 101
If bc = "Vp" Then BOD2CAPSIM = 102
If bc = "Vpx" Then BOD2CAPSIM = 102
If bc = "Vr" Then BOD2CAPSIM = 101
If bc = "Vs" Then BOD2CAPSIM = 101
If bc = "Vsc" Then BOD2CAPSIM = 101
If bc = "vWp" Then BOD2CAPSIM = 102
If bc = "vWpg" Then BOD2CAPSIM = 102
If bc = "vWpt" Then BOD2CAPSIM = 102
If bc = "vWpx" Then BOD2CAPSIM = 102
If bc = "vWz" Then BOD2CAPSIM = 102
If bc = "vWzg" Then BOD2CAPSIM = 102
If bc = "vWzr" Then BOD2CAPSIM = 102
If bc = "vWzt" Then BOD2CAPSIM = 102
If bc = "vWzx" Then BOD2CAPSIM = 102
If bc = "Vz" Then BOD2CAPSIM = 102
If bc = "Vzc" Then BOD2CAPSIM = 102
If bc = "Vzg" Then BOD2CAPSIM = 102
If bc = "Vzt" Then BOD2CAPSIM = 102
If bc = "Vzx" Then BOD2CAPSIM = 102
If bc = "Wg" Then BOD2CAPSIM = 106
If bc = "Wgl" Then BOD2CAPSIM = 106
If bc = "Wo" Then BOD2CAPSIM = 106
If bc = "Wol" Then BOD2CAPSIM = 106
If bc = "Wov" Then BOD2CAPSIM = 106
If bc = "Y21" Then BOD2CAPSIM = 109
If bc = "Y21g" Then BOD2CAPSIM = 110
If bc = "Y21x" Then BOD2CAPSIM = 111
If bc = "Y23" Then BOD2CAPSIM = 113
If bc = "Y23b" Then BOD2CAPSIM = 113
If bc = "Y23g" Then BOD2CAPSIM = 110
If bc = "Y23x" Then BOD2CAPSIM = 111
If bc = "Y30" Then BOD2CAPSIM = 114
If bc = "Y30x" Then BOD2CAPSIM = 114
If bc = "Zb20A" Then BOD2CAPSIM = 107
If bc = "Zb21" Then BOD2CAPSIM = 109
If bc = "Zb21g" Then BOD2CAPSIM = 110
If bc = "Zb23" Then BOD2CAPSIM = 113
If bc = "Zb23g" Then BOD2CAPSIM = 113
If bc = "Zb23t" Then BOD2CAPSIM = 111
If bc = "Zb23x" Then BOD2CAPSIM = 111
If bc = "Zb30" Then BOD2CAPSIM = 114
If bc = "Zb30A" Then BOD2CAPSIM = 114
If bc = "Zb30g" Then BOD2CAPSIM = 114
If bc = "Zd20A" Then BOD2CAPSIM = 107
If bc = "Zd20Ab" Then BOD2CAPSIM = 107
If bc = "Zd21" Then BOD2CAPSIM = 107
If bc = "Zd21g" Then BOD2CAPSIM = 107
If bc = "Zd23" Then BOD2CAPSIM = 113
If bc = "Zd30" Then BOD2CAPSIM = 114
If bc = "Zd30A" Then BOD2CAPSIM = 114
If bc = "zEZ21" Then BOD2CAPSIM = 112
If bc = "zEZ21g" Then BOD2CAPSIM = 112
If bc = "zEZ21t" Then BOD2CAPSIM = 112
If bc = "zEZ21w" Then BOD2CAPSIM = 112
If bc = "zEZ21x" Then BOD2CAPSIM = 112
If bc = "zEZ23" Then BOD2CAPSIM = 112
If bc = "zEZ23g" Then BOD2CAPSIM = 112
If bc = "zEZ23t" Then BOD2CAPSIM = 112
If bc = "zEZ23w" Then BOD2CAPSIM = 112
If bc = "zEZ23x" Then BOD2CAPSIM = 112
If bc = "zEZ30" Then BOD2CAPSIM = 112
If bc = "zEZ30g" Then BOD2CAPSIM = 112
If bc = "zEZ30x" Then BOD2CAPSIM = 112
If bc = "zgHd30" Then BOD2CAPSIM = 114
If bc = "zgMn15C" Then BOD2CAPSIM = 115
If bc = "zgMn88C" Then BOD2CAPSIM = 117
If bc = "zgY30" Then BOD2CAPSIM = 114
If bc = "zHd21" Then BOD2CAPSIM = 108
If bc = "zHd21g" Then BOD2CAPSIM = 108
If bc = "zHn21" Then BOD2CAPSIM = 108
If bc = "zHn23" Then BOD2CAPSIM = 109
If bc = "zhVk" Then BOD2CAPSIM = 106
If bc = "zKRn1g" Then BOD2CAPSIM = 120
If bc = "zKRn2" Then BOD2CAPSIM = 119
If bc = "zkVc" Then BOD2CAPSIM = 103
If bc = "zkWp" Then BOD2CAPSIM = 104
If bc = "zMn15A" Then BOD2CAPSIM = 115
If bc = "zMn22Ap" Then BOD2CAPSIM = 119
If bc = "zMn25Ap" Then BOD2CAPSIM = 119
If bc = "zMn56Cp" Then BOD2CAPSIM = 117
If bc = "zMo10A" Then BOD2CAPSIM = 115
If bc = "zMv41C" Then BOD2CAPSIM = 118
If bc = "zMv61C" Then BOD2CAPSIM = 118
If bc = "Zn10A" Then BOD2CAPSIM = 107
If bc = "Zn10Ap" Then BOD2CAPSIM = 107
If bc = "Zn10Av" Then BOD2CAPSIM = 107
If bc = "Zn10Aw" Then BOD2CAPSIM = 107
If bc = "Zn10Awp" Then BOD2CAPSIM = 107
If bc = "Zn21" Then BOD2CAPSIM = 107
If bc = "Zn21-F" Then BOD2CAPSIM = 107
If bc = "Zn21g" Then BOD2CAPSIM = 107
If bc = "Zn21-H" Then BOD2CAPSIM = 107
If bc = "Zn21p" Then BOD2CAPSIM = 107
If bc = "Zn21r" Then BOD2CAPSIM = 107
If bc = "Zn21t" Then BOD2CAPSIM = 107
If bc = "Zn21v" Then BOD2CAPSIM = 107
If bc = "Zn21w" Then BOD2CAPSIM = 107
If bc = "Zn21x" Then BOD2CAPSIM = 107
If bc = "Zn21x-F" Then BOD2CAPSIM = 107
If bc = "Zn23" Then BOD2CAPSIM = 113
If bc = "Zn23-F" Then BOD2CAPSIM = 113
If bc = "Zn23g" Then BOD2CAPSIM = 113
If bc = "Zn23g-F" Then BOD2CAPSIM = 113
If bc = "Zn23-H" Then BOD2CAPSIM = 113
If bc = "Zn23p" Then BOD2CAPSIM = 113
If bc = "Zn23r" Then BOD2CAPSIM = 113
If bc = "Zn23t" Then BOD2CAPSIM = 111
If bc = "Zn23x" Then BOD2CAPSIM = 111
If bc = "Zn30" Then BOD2CAPSIM = 114
If bc = "Zn30A" Then BOD2CAPSIM = 114
If bc = "Zn30Ab" Then BOD2CAPSIM = 114
If bc = "Zn30Ag" Then BOD2CAPSIM = 114
If bc = "Zn30Ar" Then BOD2CAPSIM = 114
If bc = "Zn30g" Then BOD2CAPSIM = 114
If bc = "Zn30r" Then BOD2CAPSIM = 114
If bc = "Zn30v" Then BOD2CAPSIM = 114
If bc = "Zn30x" Then BOD2CAPSIM = 114
If bc = "Zn40A" Then BOD2CAPSIM = 107
If bc = "Zn40Ap" Then BOD2CAPSIM = 107
If bc = "Zn40Ar" Then BOD2CAPSIM = 107
If bc = "Zn40Av" Then BOD2CAPSIM = 107
If bc = "Zn50A" Then BOD2CAPSIM = 107
If bc = "Zn50Ab" Then BOD2CAPSIM = 107
If bc = "Zn50Ap" Then BOD2CAPSIM = 107
If bc = "Zn50Ar" Then BOD2CAPSIM = 107
If bc = "Zn50Aw" Then BOD2CAPSIM = 107
If bc = "zpZn23w" Then BOD2CAPSIM = 113
If bc = "zRd10A" Then BOD2CAPSIM = 119
If bc = "zRn15C" Then BOD2CAPSIM = 115
If bc = "zRn47Cwp" Then BOD2CAPSIM = 117
If bc = "zRn62C" Then BOD2CAPSIM = 119
If bc = "zSn14A" Then BOD2CAPSIM = 113
If bc = "zVc" Then BOD2CAPSIM = 105
If bc = "zVp" Then BOD2CAPSIM = 105
If bc = "zVpg" Then BOD2CAPSIM = 105
If bc = "zVpt" Then BOD2CAPSIM = 105
If bc = "zVpx" Then BOD2CAPSIM = 105
If bc = "zVs" Then BOD2CAPSIM = 105
If bc = "zVz" Then BOD2CAPSIM = 105
If bc = "zVzg" Then BOD2CAPSIM = 105
If bc = "zVzt" Then BOD2CAPSIM = 105
If bc = "zVzx" Then BOD2CAPSIM = 105
If bc = "zWp" Then BOD2CAPSIM = 105
If bc = "zWpg" Then BOD2CAPSIM = 105
If bc = "zWpt" Then BOD2CAPSIM = 105
If bc = "zWpx" Then BOD2CAPSIM = 105
If bc = "zWz" Then BOD2CAPSIM = 105
If bc = "zWzg" Then BOD2CAPSIM = 105
If bc = "zWzt" Then BOD2CAPSIM = 105
If bc = "zWzx" Then BOD2CAPSIM = 105
If bc = "zY21" Then BOD2CAPSIM = 108
If bc = "zY21g" Then BOD2CAPSIM = 108
If bc = "zY23" Then BOD2CAPSIM = 109
If bc = "zY30" Then BOD2CAPSIM = 114

End Function


Public Sub RunDoEvents(PauseTime As Long)
  Dim start, Finish, TotalTime
  start = Timer ' Set start time.
  Do While Timer < start + PauseTime
    DoEvents ' Yield to other processes.
  Loop
End Sub

Public Function ShellandWait(ExeFullPath As String, _
                                Optional TimeOutValue As Long = 0, _
                                Optional CheckReturnCode As Boolean = False, Optional ByRef ReturnCodeFile As String) As Boolean
    
    Dim lInst As Long
    Dim lStart As Long
    Dim lTimeToQuit As Long
    Dim sExeName As String
    Dim lProcessId As Long
    Dim lExitCode As Long
    Dim bPastMidnight As Boolean
    Dim ExeDirectory As String
    
    On Error GoTo errorhandler

    'paths with .'s and or spaces go wrong. So fix it here by surrounding them with double quotes
    lStart = CLng(Timer)
    sExeName = Trim(ExeFullPath)
    
    'set the directory where the executable resides as the current dir
    ExeDirectory = GetDirectory(sExeName)
    Call ChDir(ExeDirectory)
    
    If Left(sExeName, 1) <> Chr(34) Or Right(sExeName, 1) <> Chr(34) Then
      sExeName = Chr(34) & sExeName
      sExeName = sExeName & Chr(34)
    End If

    'Deal with timeout being reset at VBA.Midnight
    If TimeOutValue > 0 Then
        If lStart + TimeOutValue < 86400 Then
            lTimeToQuit = lStart + TimeOutValue
        Else
            lTimeToQuit = (lStart - 86400) + TimeOutValue
            bPastMidnight = True
        End If
    End If

    lInst = Shell(sExeName, vbMinimizedNoFocus)
    
    lProcessId = OpenProcess(PROCESS_QUERY_INFORMATION, False, lInst)

    Do
        Call GetExitCodeProcess(lProcessId, lExitCode)
        DoEvents
        If TimeOutValue And Timer > lTimeToQuit Then
            If bPastMidnight Then
                 If Timer < lStart Then Exit Do
            Else
                 Exit Do
            End If
    End If
    Loop While lExitCode = STATUS_PENDING
    
    If CheckReturnCode Then
      If FileExists(ReturnCodeFile) Then
        Dim hf As Long
        Dim st As String
        hf = FreeFile
        Open ReturnCodeFile For Input As #hf
        Input #hf, st
        Close #hf
        
        ShellandWait = (Val(st) = 0)
      Else
        ShellandWait = False
      End If
    Else
      ShellandWait = True
    End If
    
    Exit Function
   
errorhandler:
ShellandWait = False
Exit Function
End Function

Public Function HydroZomerWinter(myDate As Double, Optional SkipFromMonth As Long = 0, Optional SkipFromDay As Long = 0, Optional SkipToMonth As Integer = 0, Optional SkipToDay As Integer = 0) As String
  
  'integriteitscontrole
  If SkipToMonth < SkipFromMonth Then
    HydroZomerWinter = "error in function HydroZomerWinter"
    Exit Function
  End If
  
  'check eerst of hij geskipped moet worden
  If Month(myDate) = SkipFromMonth Then
    If Day(myDate) >= SkipFromDay Then
      HydroZomerWinter = "overgeslagen"
      Exit Function
    End If
  ElseIf Month(myDate) = SkipToMonth Then
    If Day(myDate) <= SkipToDay Then
      HydroZomerWinter = "overgeslagen"
      Exit Function
    End If
  ElseIf Month(myDate) > SkipFromMonth And Month(myDate) < SkipToMonth Then
    HydroZomerWinter = "overgeslagen"
    Exit Function
  End If
  
  Select Case Month(myDate)
  Case Is = 1
    HydroZomerWinter = "winter"
  Case Is = 2
    HydroZomerWinter = "winter"
  Case Is = 3
    HydroZomerWinter = "winter"
  Case Is = 4
    If Day(myDate) < 15 Then
      HydroZomerWinter = "winter"
    Else
      HydroZomerWinter = "zomer"
    End If
  Case Is = 5
    HydroZomerWinter = "zomer"
  Case Is = 6
    HydroZomerWinter = "zomer"
  Case Is = 7
    HydroZomerWinter = "zomer"
  Case Is = 8
    HydroZomerWinter = "zomer"
  Case Is = 9
    HydroZomerWinter = "zomer"
  Case Is = 10
    If Day(myDate) < 15 Then
      HydroZomerWinter = "zomer"
    Else
      HydroZomerWinter = "winter"
    End If
  Case Is = 11
    HydroZomerWinter = "winter"
  Case Is = 12
    HydroZomerWinter = "winter"
  End Select
End Function

Public Function EVAPMAKKINK(Kin As Double, Tdag As Double, Tmin As Double, Tmax As Double, P As Double) As Double
  Dim esat As Double
  Dim s As Double
  Dim y As Double
  Dim lambdaE As Double

  'lambda = verdampingwarmte van water (2.45E06 J/kg)
  'E = verdampingsflux (kg/m2/s)
  'a_accent = constante (ongeveer 1.1)
  'Rn = nettostraling (W/m2)
  's = afgeleide van ew bij luchttemperatuur T (Pa/K), dus s = dew/dT
  'y = psychrometerconstante in Pa/K
  'G = bodemwarmtestroom
  'Beta = 10 W/m2


  esat = Verzadigingsdampdruk(Tmin, Tmax)
  s = DampDrukGradient(esat, Tdag)
  y = PsychrometerConstante(P, VerdampingswarmteWater(Tdag))

  lambdaE = 0.65 * s / (s + y) * Kin
  
  'converteer naar mm/d
  EVAPMAKKINK = lambdaE / 2450000 / 1000 * 1000 * 3600 * 24 ' =lmbdaE * 0.035

End Function

Public Sub MAKKINK2OPENWATER(StartRow As Integer, DateCol As Integer, ValCol As Integer, resultsRow As Integer, resultsCol As Integer)
  'This routine converts evaporation according to Makkink (referentiegewasverdamping) to evaporation of openwater bodies
  
  Dim r As Long, c_dat As Long, c_val As Long
  Dim r_res As Long, c_res As Long
  Dim myDate As Date, myVal As Double
  r = StartRow - 1
  c_dat = DateCol
  c_val = ValCol
  r_res = resultsRow - 1
  c_res = resultsCol
  
  While Not ActiveSheet.Cells(r + 1, c_dat) = ""
    r = r + 1
    r_res = r_res + 1
    
    myDate = ActiveSheet.Cells(r, c_dat)
    myVal = ActiveSheet.Cells(r, c_val)
    
    ActiveSheet.Cells(r_res, c_res) = myDate
    ActiveSheet.Cells(r_res, c_res + 1) = myVal * OPENWATEREVAPFACTOR(myDate)
  Wend

End Sub

Public Sub EVAPDAY2HOUR(StartRow As Integer, DateCol As Integer, ValCol As Integer, resultsRow As Integer, resultsCol As Integer)
  'spreads daily evaporation sum out over 24 hours within the day
  Dim r As Long, c_dat As Long, c_val As Long
  Dim r_res As Long, c_res As Long, i As Integer
  Dim myDate As Date, myVal As Double
  r = StartRow - 1
  c_dat = DateCol
  c_val = ValCol
  r_res = resultsRow - 1
  c_res = resultsCol
  
  While Not ActiveSheet.Cells(r + 1, c_dat) = ""
    r = r + 1
    myDate = ActiveSheet.Cells(r, c_dat)
    myVal = ActiveSheet.Cells(r, c_val)
    
    For i = 1 To 24
     r_res = r_res + 1
     ActiveSheet.Cells(r_res, c_res) = myDate + (i - 1) / 24
     ActiveSheet.Cells(r_res, c_res + 1) = myVal * HOURLYEVAPORATIONFRACTION(i)
    Next
    
  Wend
End Sub

Public Function HOURLYEVAPORATIONFRACTION(h As Integer) As Double

  If h <= 6 Then
    HOURLYEVAPORATIONFRACTION = 0
  ElseIf h >= 18 Then
    HOURLYEVAPORATIONFRACTION = 0
  ElseIf h = 7 Or h = 17 Then
    HOURLYEVAPORATIONFRACTION = 0.03
  ElseIf h = 8 Or h = 16 Then
    HOURLYEVAPORATIONFRACTION = 0.07
  ElseIf h = 9 Or h = 15 Then
    HOURLYEVAPORATIONFRACTION = 0.09
  ElseIf h = 10 Or h = 14 Then
    HOURLYEVAPORATIONFRACTION = 0.11
  ElseIf h = 11 Or h = 13 Then
    HOURLYEVAPORATIONFRACTION = 0.13
  ElseIf h = 12 Then
    HOURLYEVAPORATIONFRACTION = 0.14
  Else
  HOURLYEVAPORATIONFRACTION = 0
  End If
  
End Function

Public Function OPENWATEREVAPFACTOR(myDate As Date) As Double
  'retrieves the openwater evaporation multiplication w.r.t. Makkink evaporation for a given date
  
  Dim MonthDay As String
  MonthDay = VBA.Trim(VBA.Str(Month(myDate))) & "_" & VBA.Trim(VBA.Str(Day(myDate)))
  
  
  Select Case MonthDay
Case Is = "1_1"
  OPENWATEREVAPFACTOR = 0.5
Case Is = "1_2"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_3"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_4"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_5"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_6"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_7"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_8"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_9"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_10"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_11"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_12"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_13"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_14"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_15"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_16"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_17"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_18"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_19"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_20"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_21"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_22"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_23"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_24"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_25"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_26"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_27"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_28"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_29"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_30"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_31"
OPENWATEREVAPFACTOR = 0.7
Case Is = "2_1"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_2"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_3"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_4"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_5"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_6"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_7"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_8"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_9"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_10"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_11"
OPENWATEREVAPFACTOR = 1
Case Is = "2_12"
OPENWATEREVAPFACTOR = 1
Case Is = "2_13"
OPENWATEREVAPFACTOR = 1
Case Is = "2_14"
OPENWATEREVAPFACTOR = 1
Case Is = "2_15"
OPENWATEREVAPFACTOR = 1
Case Is = "2_16"
OPENWATEREVAPFACTOR = 1
Case Is = "2_17"
OPENWATEREVAPFACTOR = 1
Case Is = "2_18"
OPENWATEREVAPFACTOR = 1
Case Is = "2_19"
OPENWATEREVAPFACTOR = 1
Case Is = "2_20"
OPENWATEREVAPFACTOR = 1
Case Is = "2_21"
OPENWATEREVAPFACTOR = 1
Case Is = "2_22"
OPENWATEREVAPFACTOR = 1
Case Is = "2_23"
OPENWATEREVAPFACTOR = 1
Case Is = "2_24"
OPENWATEREVAPFACTOR = 1
Case Is = "2_25"
OPENWATEREVAPFACTOR = 1
Case Is = "2_26"
OPENWATEREVAPFACTOR = 1
Case Is = "2_27"
OPENWATEREVAPFACTOR = 1
Case Is = "2_28"
OPENWATEREVAPFACTOR = 1
Case Is = "2_29"
OPENWATEREVAPFACTOR = 1
Case Is = "3_1"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_2"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_3"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_4"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_5"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_6"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_7"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_8"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_9"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_10"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_11"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_12"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_13"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_14"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_15"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_16"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_17"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_18"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_19"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_20"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_21"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_22"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_23"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_24"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_25"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_26"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_27"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_28"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_29"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_30"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_31"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_1"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_2"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_3"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_4"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_5"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_6"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_7"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_8"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_9"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_10"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_11"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_12"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_13"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_14"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_15"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_16"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_17"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_18"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_19"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_20"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_21"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_22"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_23"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_24"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_25"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_26"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_27"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_28"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_29"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_30"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_1"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_2"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_3"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_4"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_5"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_6"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_7"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_8"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_9"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_10"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_11"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_12"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_13"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_14"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_15"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_16"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_17"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_18"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_19"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_20"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_21"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_22"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_23"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_24"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_25"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_26"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_27"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_28"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_29"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_30"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_31"
OPENWATEREVAPFACTOR = 1.3
Case Is = "6_1"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_2"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_3"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_4"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_5"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_6"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_7"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_8"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_9"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_10"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_11"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_12"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_13"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_14"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_15"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_16"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_17"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_18"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_19"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_20"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_21"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_22"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_23"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_24"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_25"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_26"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_27"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_28"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_29"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_30"
OPENWATEREVAPFACTOR = 1.31
Case Is = "7_1"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_2"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_3"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_4"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_5"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_6"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_7"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_8"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_9"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_10"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_11"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_12"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_13"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_14"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_15"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_16"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_17"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_18"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_19"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_20"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_21"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_22"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_23"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_24"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_25"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_26"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_27"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_28"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_29"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_30"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_31"
OPENWATEREVAPFACTOR = 1.24
Case Is = "8_1"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_2"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_3"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_4"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_5"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_6"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_7"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_8"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_9"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_10"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_11"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_12"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_13"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_14"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_15"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_16"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_17"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_18"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_19"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_20"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_21"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_22"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_23"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_24"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_25"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_26"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_27"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_28"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_29"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_30"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_31"
OPENWATEREVAPFACTOR = 1.18
Case Is = "9_1"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_2"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_3"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_4"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_5"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_6"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_7"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_8"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_9"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_10"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_11"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_12"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_13"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_14"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_15"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_16"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_17"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_18"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_19"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_20"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_21"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_22"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_23"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_24"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_25"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_26"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_27"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_28"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_29"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_30"
OPENWATEREVAPFACTOR = 1.17
Case Is = "10_1"
OPENWATEREVAPFACTOR = 1
Case Is = "10_2"
OPENWATEREVAPFACTOR = 1
Case Is = "10_3"
OPENWATEREVAPFACTOR = 1
Case Is = "10_4"
OPENWATEREVAPFACTOR = 1
Case Is = "10_5"
OPENWATEREVAPFACTOR = 1
Case Is = "10_6"
OPENWATEREVAPFACTOR = 1
Case Is = "10_7"
OPENWATEREVAPFACTOR = 1
Case Is = "10_8"
OPENWATEREVAPFACTOR = 1
Case Is = "10_9"
OPENWATEREVAPFACTOR = 1
Case Is = "10_10"
OPENWATEREVAPFACTOR = 1
Case Is = "10_11"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_12"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_13"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_14"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_15"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_16"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_17"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_18"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_19"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_20"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_21"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_22"
OPENWATEREVAPFACTOR = 0.8
Case Is = "10_23"
OPENWATEREVAPFACTOR = 0.8
Case Is = "10_24"
OPENWATEREVAPFACTOR = 0.8
Case Is = "10_25"
OPENWATEREVAPFACTOR = 0.8
Case Is = "10_26"
OPENWATEREVAPFACTOR = 0.8
Case Is = "10_27"
OPENWATEREVAPFACTOR = 0.8
Case Is = "10_28"
OPENWATEREVAPFACTOR = 0.8
Case Is = "10_29"
OPENWATEREVAPFACTOR = 0.8
Case Is = "10_30"
OPENWATEREVAPFACTOR = 0.8
Case Is = "10_31"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_1"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_2"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_3"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_4"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_5"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_6"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_7"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_8"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_9"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_10"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_11"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_12"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_13"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_14"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_15"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_16"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_17"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_18"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_19"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_20"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_21"
OPENWATEREVAPFACTOR = 0.6
Case Is = "11_22"
OPENWATEREVAPFACTOR = 0.6
Case Is = "11_23"
OPENWATEREVAPFACTOR = 0.6
Case Is = "11_24"
OPENWATEREVAPFACTOR = 0.6
Case Is = "11_25"
OPENWATEREVAPFACTOR = 0.6
Case Is = "11_26"
OPENWATEREVAPFACTOR = 0.6
Case Is = "11_27"
OPENWATEREVAPFACTOR = 0.6
Case Is = "11_28"
OPENWATEREVAPFACTOR = 0.6
Case Is = "11_29"
OPENWATEREVAPFACTOR = 0.6
Case Is = "11_30"
OPENWATEREVAPFACTOR = 0.6
Case Is = "12_1"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_2"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_3"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_4"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_5"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_6"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_7"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_8"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_9"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_10"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_11"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_12"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_13"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_14"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_15"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_16"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_17"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_18"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_19"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_20"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_21"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_22"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_23"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_24"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_25"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_26"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_27"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_28"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_29"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_30"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_31"
OPENWATEREVAPFACTOR = 0.5
Case Else
OPENWATEREVAPFACTOR = 0
End Select
  
End Function

Public Function EVAPDEBRUINKEIJMAN(Datum As Double, Kin As Double, Tdag As Double, Tmin As Double, Tmax As Double, P As Double, SP As Double, UG As Double, D As Double) As Double
  
  'Kin = dagsom globale straling (W/m2)
  'Tdag = gemiddelde dagtemperatuur (Celcius en K)
  'Tmin = minimum dagtemperatuur (celcius en K)
  'Tmax = maximum dagtemperatuur (Celcius en K)
  'p = luchtdruk kPa op hoogte z0
  'SP = percentage v d langst mogelijke zonneschijn
  'UG = percentage luchtvochtigheid
  'd = gemiddelde waterdiepte over het gebied
  
  'a_accent = constante (-)
  'beta = 10 W/m2
  'dTdt() verandering watertemperatuur in de tijd K/s.
  'RN = nettostraling in W/m2
  'literatuur:
  'Futurewater, 2006, Berekening openwaterverdamping, in opdracht van wetterskip fryslan
  'STOWA 2009, verbetering bepaling actuele verdamping voor het strategisch waterbeheer, definitiestudie
  
  'OmrekeningSfAcTOren vOOr mm, mJ en WATTS.
  '                       mm d^(-1)             mJ * m^(-2) * d^(-1)  W m^(-2)
  '  1 mm d-1             1.000                 2.451                 28.368
  '  1 MJ m-2 d-1         0.408                 1.000                 11.574
  '  1 W m-2              0.035                 0.086                 1.000
  
  Dim a_accent As Double    'ongeveer 1.1
  Dim beta As Double        'beta = 10 W/m2
  Dim dTdt(1 To 12) As Double 'verandering watertemperatuur in de tijd K/s.
  Dim g As Double
  Dim lambdaE As Double     'de verdampingswarmteflux. converteren we m.b.v. de tabel naar mm/d
  Dim Rn As Double          'nettostraling in W/m2
  Dim esat As Double        'verzadigingsdampdruk (hPA)
  Dim ez As Double          'dampdruk (hPa)
  Dim Nrel As Double        'relatieve zonneschijn
  Dim LNetto As Double      'netto langgolvige straling
  Dim y As Double           'psychrometerconstante
  Dim s As Double           'dampdrukgradiënt
  Dim RH As Double          'relative humidity (-)
  
  'vergelijking van de bruin-Keijman:
  'lambda * E = a_accent * s/(s + y) *(Rn - G) + Beta
  
  'waarin:
  'lambda = verdampingwarmte van water (2.45E06 J/kg)
  'E = verdampingsflux (kg/m2/s)
  'a_accent = constante (ongeveer 1.1)
  'Rn = nettostraling (W/m2)
  's = afgeleide van ew bij luchttemperatuur T (Pa/K), dus s = dew/dT
  'y = psychrometerconstante in Pa/K
  'G = bodemwarmtestroom
  'Beta = 10 W/m2
  
  '-------------------------------------------------------------------
  'bereken eerst de bodemwarmtestroom G (W/m2)
  g = BodemWarmteStroom(Datum, D)
    
  'bereken verzadigingsdampdruk esat (hPa) en dampdruk e(z) op hoogte z
  esat = Verzadigingsdampdruk(Tmin, Tmax)
  RH = UG / 100     'relatieve luchtvochtigheid
  ez = RH * esat
  
  'bereken de relatieve zonneschijnduur en de netto langgolvige straling (in W/m2)
  Nrel = SP / 100
  LNetto = NettoLanggolvigeStraling(Tmax, Tmin, Nrel, ez)
  
  'bereken Rn (netto straling) (W/m2)
  Rn = NettoStraling(Kin, LNetto)
  
  a_accent = 1.1
  beta = 10
  y = PsychrometerConstante(P, VerdampingswarmteWater(Tdag))
  s = DampDrukGradient(esat, Tdag)

  lambdaE = a_accent * s / (s + y) * (Rn - g) + beta
  
  'converteer naar mm/d
  EVAPDEBRUINKEIJMAN = 0.035 * lambdaE
  
End Function

Public Function Verzadigingsdampdruk(Tmin As Double, Tmax As Double) As Double
  'rekent de verzadigingsdampdruk uit a.d.h.v. minimum en maximum luchttemperatuur
  'eenheid: kPa
  Verzadigingsdampdruk = 0.305 * (Exp(17.27 * Tmin / Celcius2Kelvin(Tmin)) + Exp(17.27 * Tmax / Celcius2Kelvin(Tmax)))
End Function

Public Function BodemWarmteStroom(Datum As Double, D As Double) As Double
  
  'deze functie berekent de bodemwarmtestroom G in W/m2 ten behoeve van de
  'berekening van openwaterverdamping met de De Bruin, Keijman - formule
    
  Dim dTdt(1 To 12) As Double
  Dim rho_water As Double   'dichtheid water = 1000 kg/m3
  Dim c_water As Double     'soortelijke warmte water = 4200 J/kg/K
  'd = gemiddelde waterdiepte in het gebied
  
  rho_water = 1000
  c_water = 4200
  
  'temperatuurveranderingen in de tijd K/s
  'bron: Futurewater, 2006 Tabel A.2
  dTdt(1) = -0.000000746714
  dTdt(2) = 0.000000373357
  dTdt(3) = 0.00000119732
  dTdt(4) = 0.00000112007
  dTdt(5) = 0.00000192901
  dTdt(6) = 0.00000112007
  dTdt(7) = 0.000000385802
  dTdt(8) = 0.000000373357
  dTdt(9) = -0.00000112007
  dTdt(10) = -0.00000115741
  dTdt(11) = -0.00000224014
  dTdt(12) = -0.00000115741

  BodemWarmteStroom = rho_water * c_water * D * dTdt(Month(Datum))
   
End Function

Public Function NettoStraling(Kin As Double, LNetto As Double) As Double
'deze functie berekent de nettostraling in W/m2 ten behoeve van verdampingsberekeningen

Dim albedo As Double
albedo = 0.06 'voor water

NettoStraling = (1 - albedo) * Kin - LNetto

End Function

Public Function NettoLanggolvigeStraling(Tmax As Double, Tmin As Double, Nrel As Double, ez As Double) As Double
  'deze functie berekent de netto langgolvige straling t.b.v. verdampingsberekeningen
  'in W/m2
  
  'Tmax = maximale dagtemperatuur
  'Tmin =  minimale dagtemperatuur
  'ez = dampdruk op hoogte z (kPa)
  'Nrel = relatieve zonneschijnduur (-)
  
  Dim sbconst As Double 'stephan bolzmann constante
  sbconst = 0.000000004903 'MJ/K^4/m2/d
  NettoLanggolvigeStraling = sbconst * ((Celcius2Kelvin(Tmax) ^ 4 + Celcius2Kelvin(Tmin) ^ 4) / 2) * (0.34 - 0.14 * Sqr(ez)) * (0.1 + 0.9 * Nrel) * 11.574

End Function

Public Function DampDrukGradient(esat As Double, Tdag As Double) As Double
  'deze functie berekent de dampdrukgradiënt s bij gemiddelde dagluchttemperatuur T zoals die gebruikt wordt bij verdampingsberekeningen
  'eenheid: kPa/K
  
  DampDrukGradient = 4098 * esat / Celcius2Kelvin(Tdag) ^ 2
End Function

Public Function PsychrometerConstante(P As Double, lambda As Double) As Double
  'deze functie berekent de psychrometerconstante y
  'lambda = verdampingswarmte van water bij de gemiddelde dagtemperatuur (MJ/kg)
  'p = luchtdruk (hPa) op hoogte z0
  
  PsychrometerConstante = 0.00163 * P / lambda

End Function

Public Function VerdampingswarmteWater(Tdag As Double) As Double
  'deze functie berekent de verdampingswarmte van water (lambda) in MJ/kg bij daggemiddelde temperatuur T in celcius
  VerdampingswarmteWater = 2.501 - 0.002361 * Tdag
End Function

Public Sub MakeScatterChart(XaxisTitle As String, YAxisTitle As String, MeasTimeRange As Range, MeasDataRange As Range, SobekTimeRange As Range, SobekDataRange As Range, Title As String, Optional minX As Double = -999, Optional maxX As Double = -999)
    
  Charts.Add
  With ActiveChart
    
    'maak de eerste sobek case de basis voor deze grafiek
    '.ChartType = xlXYScatterLinesNoMarkers
    .ChartType = xlXYScatter
    .SetSourceData Source:=Union(SobekTimeRange, SobekDataRange), PlotBy:=xlColumns
    
    Call .SetElement(msoElementChartTitleAboveChart)
    '.HasTitle = True
    .ChartTitle.Text = Title
    
    .Axes(xlValue).CrossesAt = -1000  'zorg dat de x-as altijd zo laag mogelijk de y-as snijdt
    .Axes(xlCategory).HasTitle = True
    .Axes(xlCategory).AxisTitle.Characters.Text = XaxisTitle
    .Axes(xlCategory).TickLabels.NumberFormat = "dd/mm/yy"
    
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = YAxisTitle
    .Name = Title
    
    'voeg SOBEK resultaten toe
    .SeriesCollection.NewSeries
    .SeriesCollection(1).Name = "Berekend"
    .SeriesCollection(1).XValues = SobekTimeRange
    .SeriesCollection(1).Values = SobekDataRange
    .SeriesCollection(1).MarkerSize = 5
    .SeriesCollection(1).MarkerStyle = xlMarkerStyleDash
    
    'tot slot, voeg de meetgegevens toe als serie aan de grafiek
    .SeriesCollection.NewSeries
    .SeriesCollection(2).ChartType = xlXYScatter
    .SeriesCollection(2).Name = "Gemeten"
    .SeriesCollection(2).XValues = MeasTimeRange
    .SeriesCollection(2).Values = MeasDataRange
      
    'opmaak
    .SeriesCollection(2).MarkerBackgroundColorIndex = 40
    .SeriesCollection(2).MarkerForegroundColorIndex = 3
    .SeriesCollection(2).MarkerStyle = xlMarkerStyleDash
    .SeriesCollection(2).Smooth = False
    .SeriesCollection(2).MarkerSize = 5
    .SeriesCollection(2).Shadow = False
    
    If minX <> -999 Then .Axes(xlCategory).MinimumScale = minX
    If maxX <> -999 Then .Axes(xlCategory).MaximumScale = maxX
    
  End With
  

End Sub

Public Sub MakeChart(XaxisTitle As String, YAxisTitle As String, MeasTimeRange As Range, MeasDataRange As Range, SobekTimeRange As Range, SobekDataRange As Range, Title As String, Optional minX As Double = -999, Optional maxX As Double = -999)
    
  Charts.Add
  With ActiveChart
    
    'maak de eerste sobek case de basis voor deze grafiek
    .ChartType = xlXYScatterLinesNoMarkers
    .SetSourceData Source:=Union(SobekTimeRange, SobekDataRange), PlotBy:=xlColumns
    
    Call .SetElement(msoElementChartTitleAboveChart)
    '.HasTitle = True
    .ChartTitle.Text = Title
    
    .Axes(xlValue).CrossesAt = -1000  'zorg dat de x-as altijd zo laag mogelijk de y-as snijdt
    .Axes(xlCategory).HasTitle = True
    .Axes(xlCategory).AxisTitle.Characters.Text = XaxisTitle
    .Axes(xlCategory).TickLabels.NumberFormat = "dd/mm/yy"
    
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = YAxisTitle
    .Name = Title
    .Axes(xlCategory, xlPrimary).TickLabels.Orientation = xlUpward
    
    'voeg SOBEK resultaten toe
    .SeriesCollection.NewSeries
    .SeriesCollection(1).Name = "Berekend"
    .SeriesCollection(1).XValues = SobekTimeRange
    .SeriesCollection(1).Values = SobekDataRange
    
    'tot slot, voeg de meetgegevens toe als serie aan de grafiek
    .SeriesCollection.NewSeries
    .SeriesCollection(2).ChartType = xlXYScatter
    .SeriesCollection(2).Name = "Gemeten"
    .SeriesCollection(2).XValues = MeasTimeRange
    .SeriesCollection(2).Values = MeasDataRange
      
    'opmaak
    .SeriesCollection(2).MarkerBackgroundColorIndex = 40
    .SeriesCollection(2).MarkerForegroundColorIndex = 3
    .SeriesCollection(2).MarkerStyle = xlCircle
    .SeriesCollection(2).Smooth = False
    .SeriesCollection(2).MarkerSize = 2
    .SeriesCollection(2).Shadow = False
    
    If minX <> -999 Then .Axes(xlCategory).MinimumScale = minX
    If maxX <> -999 Then .Axes(xlCategory).MaximumScale = maxX
    
  End With
  
End Sub

Sub ExportChart(ChartIndex As Integer, myFileNameNoExtension As String)
    
    Dim myChart As Chart
    Set myChart = ActiveWorkbook.Charts(ChartIndex)

    'myFileName = "myChart.png"
    

    On Error Resume Next
    Kill ThisWorkbook.path & "\" & myFileNameNoExtension
    On Error GoTo 0

    myChart.Export FileName:=ThisWorkbook.path & "\" & myFileNameNoExtension & ".png", Filtername:="PNG"

    MsgBox "OK"
End Sub

Public Function FileExists(path As String) As Boolean
  'controleert of een bestand bestaat
  If VBA.Trim(path) = "" Then
    FileExists = False
  Else
    FileExists = (VBA.dir(path) > "")
  End If
End Function

Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then 'See above
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If
End Sub

Public Function IB(Bruto As Double) As Double

Dim SchaalMax(1 To 3) As Double
Dim SchaalPerc(1 To 4) As Double

SchaalMax(1) = 18628
SchaalMax(2) = 33436
SchaalMax(3) = 55694

SchaalPerc(1) = 33
SchaalPerc(2) = 41.95
SchaalPerc(3) = 42
SchaalPerc(3) = 52

If Bruto <= SchaalMax(1) Then
  IB = IB + Bruto * SchaalPerc(1) / 100
ElseIf Bruto <= SchaalMax(2) Then
  IB = IB + SchaalMax(1) * SchaalPerc(1) / 100
  IB = IB + (Bruto - SchaalMax(1)) * SchaalPerc(2) / 100
ElseIf Bruto <= SchaalMax(3) Then
  IB = IB + SchaalMax(1) * SchaalPerc(1) / 100
  IB = IB + (SchaalMax(2) - SchaalMax(1)) * SchaalPerc(2) / 100
  IB = IB + (Bruto - SchaalMax(2)) * SchaalPerc(3) / 100
ElseIf Bruto > SchaalMax(3) Then
  IB = IB + SchaalMax(1) * SchaalPerc(1) / 100
  IB = IB + (SchaalMax(2) - SchaalMax(1)) * SchaalPerc(2) / 100
  IB = IB + (SchaalMax(3) - SchaalMax(2)) * SchaalPerc(3) / 100
  IB = IB + (Bruto - SchaalMax(3)) * SchaalPerc(4) / 100
End If

End Function


Public Function TIDALMINMAXFROMARRAY(myArray() As Variant, RoundHours As Boolean) As Variant()
  'Author: Siebe Bosch
  'Date: 1-9-2013
  'Description: extracts the tidal minima and maxima from a 2D-array with date/time and water levels
  Dim i As Long, j As Long, K As Long, c As Long, n As Long, Timestep As Integer, SearchRadius As Integer
  Dim curVal As Double, IsMin As Boolean, IsMax As Boolean, Header As String, curDate As Date
  Timestep = (myArray(3, 1) - myArray(2, 1)) * 24 * 60 'in minutes
  n = UBound(myArray, 1) * (UBound(myArray, 2) - 1)
    
  'diminsioning the arrays
  Dim TidalArray() As Variant
  ReDim TidalArray(1 To n, 1 To 4)
  Dim FinalArray() As Variant
  
  'setting the search radius
  If Timestep <= 10 Then
    SearchRadius = 30 '10 minutes timestep. Detect tidal value by comparing -5 hours and + 5 hours
  ElseIf Timestep <= 15 Then
    SearchRadius = 20 '15 minutes timestep. Detect tidal value by comparing -5 hours +5 hours
  ElseIf Timestep < 60 Then
    SearchRadius = 5 '1 hour timestep. Detect tidal value by comparing -5 and +5 hours
  End If
  
  'walk through the array and search for tides
  For c = 2 To UBound(myArray, 2)
    For i = 2 + SearchRadius To UBound(myArray) - SearchRadius   'we start at row 2 since the first row contains headers
      Header = myArray(1, c)
      curDate = myArray(i, 1)
      curVal = myArray(i, c)
      
      IsMin = True
      IsMax = True
      
      'search backward
      For j = i - 1 To i - SearchRadius Step -1
        If myArray(j, c) >= curVal Then IsMax = False       'note: the >= here and the > in the next section is importand in case of equal values!
        If myArray(j, c) <= curVal Then IsMin = False       'note: the <= here and the < in the next section is importand in case of equal values!
        If IsMax = False And IsMin = False Then Exit For
      Next
      
      'search forward
      For j = i + 1 To i + SearchRadius Step 1
        If myArray(j, c) > curVal Then IsMax = False
        If myArray(j, c) < curVal Then IsMin = False
        If IsMax = False And IsMin = False Then Exit For
      Next
      
      'identify whether this point is a tidal min or max
      If IsMin Or IsMax Then
        K = K + 1
        TidalArray(K, 1) = Header
        If RoundHours Then
          TidalArray(K, 2) = DATETWOHOURWINDOW(myArray(i, 1))   'since the timing of the computed and observed peak may difference, we'll introduce a certain bandwidth
        Else
          TidalArray(K, 2) = myArray(i, 1)
        End If
        TidalArray(K, 3) = curVal
        If IsMin Then
          TidalArray(K, 4) = "Laag"
        ElseIf IsMax Then
          TidalArray(K, 4) = "Hoog"
        End If
      End If
    Next
  Next
  
  'truncate the tidal array to match the actual number of tides
  ReDim FinalArray(1 To K, 1 To 4)
  For i = 1 To K
    FinalArray(i, 1) = TidalArray(i, 1)
    FinalArray(i, 2) = TidalArray(i, 2)
    FinalArray(i, 3) = TidalArray(i, 3)
    FinalArray(i, 4) = TidalArray(i, 4)
  Next
  
  TIDALMINMAXFROMARRAY = FinalArray

End Function


Public Function TIDALMINMAXSEQUENCE(myRange As Range, resultsRow As Integer, resultsCol As Integer) As Variant

  Dim i As Long, j As Long, K As Long, n As Long, r As Long, c As Long
  Dim Location As String, myVal As Double, Timestep As Integer, SearchRadius As Integer
  Dim ValRange As Range, DateRange As Range
  Dim lastMinDate As Double, lastMaxDate As Double, lastDate As Double
  Dim lastMinVal As Double, lastMaxVal As Double
  Dim lastMinIdx As Long, lastMaxIdx As Long, nextIdx As Long
  Dim myDate As Double, FirstDone As Boolean, LocationDone As Boolean
  Dim IsMin As Boolean, IsMax As Boolean, curVal As Double
  
  'Date: 30-8-2013
  'Author: Siebe Bosch
  'Description: subdivides a day into 4 quarters and exports the tidal min or max within that quarter
  
  r = resultsRow
  c = resultsCol
  
  ActiveSheet.Cells(r, c) = "Datum/Tijd"
  ActiveSheet.Cells(r, c + 1) = myRange.Cells(1, 2) 'location name
  
  'first analyse the timestep in order to specify a window size.
  'then derive a search radius (expressed in n_timesteps) for each following min/max
  Timestep = (myRange.Cells(3, 1) - myRange.Cells(2, 1)) * 24 * 60 'in minutes
  If Timestep <= 10 Then
    SearchRadius = 12 '10 minutes timestep. Detect tidal value by comparing -2 hours and + 2 hours
  ElseIf Timestep <= 15 Then
    SearchRadius = 8 '15 minutes timestep. Detect tidal value by comparing -2 hours +2 hours
  ElseIf Timestep < 60 Then
    SearchRadius = 2 '1 hour timestep. Detect tidal value by comparing -2 and +2 hours
  End If
    
  If myRange.Columns.Count < 2 Then
    TIDALMINMAXSEQUENCE = "Error: data range must contain at least two columns: date/time and values"
  ElseIf myRange.Rows.Count < 2 Then
    TIDALMINMAXSEQUENCE = "Error: data range must contain a sufficient number of rows"
  Else
    For i = 2 + SearchRadius To myRange.Rows.Count - SearchRadius
      curVal = myRange.Cells(i, 2)
      If myRange.Cells(i + SearchRadius, 1) = 0 Then Exit For 'reached the end of the timeseries
      
      IsMin = True
      IsMax = True
      For j = i - 1 To i - SearchRadius Step -1
        If myRange.Cells(j, 2) >= curVal Then IsMax = False
        If myRange.Cells(j, 2) <= curVal Then IsMin = False
        If IsMax = False And IsMin = False Then Exit For
      Next
      For j = i + 1 To i + SearchRadius Step 1
        If myRange.Cells(j, 2) >= curVal Then IsMax = False
        If myRange.Cells(j, 2) <= curVal Then IsMin = False
        If IsMax = False And IsMin = False Then Exit For
      Next
      If IsMin Or IsMax Then
        r = r + 1
        ActiveSheet.Cells(r, c) = DATEHOUR(myRange.Cells(i, 1))
        ActiveSheet.Cells(r, c + 1) = curVal
      End If
    Next
  End If
  TIDALMINMAXSEQUENCE = True
End Function

Public Function WindRichting(angle As Double, ReturnNumeric As Boolean) As Variant
  If (angle = 0 Or angle > 360) Then
    WindRichting = "Windstil/Variabel"
  ElseIf (angle < 22.5 Or angle >= 337.5) Then
    If ReturnNumeric Then
      WindRichting = 0
    Else
      WindRichting = "N"
    End If
  ElseIf (angle < 67.5 And angle >= 22.5) Then
    If ReturnNumeric Then
      WindRichting = 45
    Else
      WindRichting = "NO"
    End If
  ElseIf (angle < 112.5 And angle >= 67.5) Then
    If ReturnNumeric Then
      WindRichting = 90
    Else
      WindRichting = "O"
    End If
  ElseIf (angle < 157.5 And angle >= 112.5) Then
    If ReturnNumeric Then
      WindRichting = 135
    Else
      WindRichting = "ZO"
    End If
  ElseIf (angle < 202.5 And angle >= 157.5) Then
    If ReturnNumeric Then
      WindRichting = 180
    Else
      WindRichting = "Z"
    End If
  ElseIf (angle < 247.5 And angle >= 202.5) Then
    If ReturnNumeric Then
      WindRichting = 225
    Else
      WindRichting = "ZW"
    End If
  ElseIf (angle < 292.5 And angle >= 247.5) Then
    If ReturnNumeric Then
      WindRichting = 270
    Else
      WindRichting = "W"
    End If
  ElseIf (angle < 337.5 And angle >= 292.5) Then
    If ReturnNumeric Then
      WindRichting = 315
    Else
      WindRichting = "NW"
    End If
  Else
    WindRichting = "Windstil/Variabel"
  End If
End Function

Public Function EXTRACTHARMONICFROMRANGE(myRange As Range, myPeriodDays As Double, resultsRow As Integer, resultsCol As Integer) As Boolean
  'this function extracts a harmonic (sinusoideal function) from a range with date/time (first column) and values (second column). (E.g. tidal movement)
  'for a given period (days) of the harmonic to extract
  'it does so by minimizing the RMS between observed and computed values
  'the remaining timeseries is written to the worksheet as well as the amplitude of the harmonic found
  
  'first calculate the average value inside the range
  Dim avgVal As Double, minVal As Double, maxVal As Double
  avgVal = Application.WorksheetFunction.Sum(Range(myRange.Cells(1, 2), myRange.Cells(myRange.Rows.Count, 2))) / myRange.Rows.Count
  minVal = Application.WorksheetFunction.Min(Range(myRange.Cells(1, 2), myRange.Cells(myRange.Rows.Count, 2)))
  maxVal = Application.WorksheetFunction.max(Range(myRange.Cells(1, 2), myRange.Cells(myRange.Rows.Count, 2)))
  
  ActiveSheet.Cells(resultsRow, resultsCol) = "gem:"
  ActiveSheet.Cells(resultsRow, resultsCol) = "min:"
  ActiveSheet.Cells(resultsRow, resultsCol) = "max:"
  ActiveSheet.Cells(resultsRow, resultsCol + 1) = avgVal
  ActiveSheet.Cells(resultsRow, resultsCol + 1) = minVal
  ActiveSheet.Cells(resultsRow, resultsCol + 1) = maxVal
  
  
End Function

Public Function TIDALMINMAXFROMSERIES(myRange As Range, resultsRow As Integer, resultsCol As Integer) As Boolean

  Dim i As Long, j As Long, K As Long, n As Long, r As Long, c As Long
  Dim Location As String, myVal As Double, Timestep As Double, SearchRadius As Integer
  Dim ValRange As Range, DateRange As Range
  Dim lastMinDate As Double, lastMaxDate As Double, lastDate As Double, startDate As Double
  Dim lastMinVal As Double, lastMaxVal As Double
  Dim lastMinIdx As Long, lastMaxIdx As Long, nextIdx As Long
  Dim myDate As Double, FirstDone As Boolean, LocationDone As Boolean
  
  r = resultsRow
  c = resultsCol
  
  ActiveSheet.Cells(r, c) = "Location"
  ActiveSheet.Cells(r, c + 1) = "Date/time low"
  ActiveSheet.Cells(r, c + 2) = "Low"
  ActiveSheet.Cells(r, c + 3) = "Date/time high"
  ActiveSheet.Cells(r, c + 4) = "High"
  
  'first analyse the timestep in order to specify a window size.
  'then derive a search radius (expressed in n_timesteps) for each following min/max
  Timestep = myRange.Cells(3, 1) - myRange.Cells(2, 1)
  SearchRadius = Application.WorksheetFunction.RoundUp((2 / 24) / Timestep, 0)
  
  If myRange.Columns.Count < 2 Then
    TIDALMINMAXFROMSERIES = "Error: data range must contain at least two columns: date/time and values"
  ElseIf myRange.Rows.Count < 2 Then
    TIDALMINMAXFROMSERIES = "Error: data range must contain a sufficient number of rows"
  Else
    For j = 2 To myRange.Columns.Count
      
      Location = myRange.Cells(1, j)
      FirstDone = False
      LocationDone = False
      
      i = 2
      startDate = myRange.Cells(i, 1)
      lastMinVal = myRange.Cells(i, j)
      lastMaxVal = myRange.Cells(i, j)
      lastMinDate = myRange.Cells(i, 1)
      lastMaxDate = myRange.Cells(i, 1)
      lastDate = myRange.Cells(i, 1)
      
      'first find the minimum and maximum in the first 13.1 hours which is a little longer than one tidal wave (12.5 h)
      While Not FirstDone = True
        myDate = myRange.Cells(i, 1)
        myVal = myRange.Cells(i, j)
        If (myDate - startDate) > (13.1 / 24) Then
          FirstDone = True
        ElseIf i > myRange.Rows.Count Then
          FirstDone = True
        Else
          If myVal < lastMinVal Then
            lastMinVal = myVal
            lastMinDate = myDate
            lastMinIdx = i
          ElseIf myVal > lastMaxVal Then
            lastMaxVal = myVal
            lastMaxDate = myDate
            lastMaxIdx = i
          End If
        End If
        i = i + 1
      Wend
      
      'write the initial results
      r = r + 1
      ActiveSheet.Cells(r, c) = Location
      ActiveSheet.Cells(r, c + 1) = lastMinDate
      ActiveSheet.Cells(r, c + 2) = lastMinVal
      ActiveSheet.Cells(r, c + 3) = lastMaxDate
      ActiveSheet.Cells(r, c + 4) = lastMaxVal
      
      'now that we have a startlocation we can start looking for the next minima and maxima
      While Not LocationDone
        lastMinVal = 99999999
        lastMaxVal = -99999999
        
        'find the next minimum approximately 12.5 hours later
        nextIdx = Math.Round(lastMinIdx + (12.5 / 24) / Timestep, 0)
        For i = nextIdx - SearchRadius To nextIdx + SearchRadius
          If i > myRange.Rows.Count Then
            LocationDone = True
          Else
            myDate = myRange.Cells(i, 1)
            myVal = myRange.Cells(i, j)
            If myVal < lastMinVal Then
              lastMinDate = myDate
              lastMinVal = myVal
              lastMinIdx = i
            End If
          End If
        Next
        
        'find the next maximum
        nextIdx = Math.Round(lastMaxIdx + (12.5 / 24) / Timestep, 0)
        For i = nextIdx - SearchRadius To nextIdx + SearchRadius
          If i > myRange.Rows.Count Then
            LocationDone = True
          Else
            myDate = myRange.Cells(i, 1)
            myVal = myRange.Cells(i, j)
            If myVal > lastMaxVal Then
              lastMaxDate = myDate
              lastMaxVal = myVal
              lastMaxIdx = i
            End If
          End If
        Next
        
        'write the results
        r = r + 1
        ActiveSheet.Cells(r, c) = Location
        ActiveSheet.Cells(r, c + 1) = lastMinDate
        ActiveSheet.Cells(r, c + 2) = lastMinVal
        ActiveSheet.Cells(r, c + 3) = lastMaxDate
        ActiveSheet.Cells(r, c + 4) = lastMaxVal
      
      Wend
    Next
  End If
  TIDALMINMAXFROMSERIES = True

End Function

Public Function TIDALLOWSFROMSERIES(myRange As Range, resultsRow As Integer, resultsCol As Integer) As Boolean

  Dim i As Long, j As Long, K As Long, n As Long, r As Long, c As Long
  Dim Location As String, myVal As Double, Timestep As Double, SearchRadius As Integer
  Dim ValRange As Range, DateRange As Range
  Dim lastMinDate As Double, lastMaxDate As Double, lastDate As Double, startDate As Double
  Dim lastMinVal As Double, lastMaxVal As Double
  Dim lastMinIdx As Long, lastMaxIdx As Long, nextIdx As Long
  Dim myDate As Double, FirstDone As Boolean, LocationDone As Boolean
  
  r = resultsRow
  c = resultsCol
  
  ActiveSheet.Cells(r, c) = "Location"
  ActiveSheet.Cells(r, c + 1) = "Date/time"
  ActiveSheet.Cells(r, c + 2) = "Low tide"
  
  'first analyse the timestep in order to specify a window size.
  'then derive a search radius (expressed in n_timesteps) for each following min/max
  Timestep = myRange.Cells(3, 1) - myRange.Cells(2, 1)
  SearchRadius = Application.WorksheetFunction.RoundUp((2 / 24) / Timestep, 0)
  
  If myRange.Columns.Count < 2 Then
    TIDALLOWSFROMSERIES = "Error: data range must contain at least two columns: date/time and values"
  ElseIf myRange.Rows.Count < 2 Then
    TIDALLOWSFROMSERIES = "Error: data range must contain a sufficient number of rows"
  Else
    For j = 2 To myRange.Columns.Count
      
      Location = myRange.Cells(1, j)
      FirstDone = False
      LocationDone = False
      
      i = 2
      startDate = myRange.Cells(i, 1)
      lastMinVal = myRange.Cells(i, j)
      lastMinDate = myRange.Cells(i, 1)
      lastDate = myRange.Cells(i, 1)
      
      'first find the minimum in the first 13.1 hours which is a little longer than one tidal wave (12.5 h)
      While Not FirstDone = True
        myDate = myRange.Cells(i, 1)
        myVal = myRange.Cells(i, j)
        If (myDate - startDate) > (13.1 / 24) Then
          FirstDone = True
        ElseIf i > myRange.Rows.Count Then
          FirstDone = True
        Else
          If myVal < lastMinVal Then
            lastMinVal = myVal
            lastMinDate = myDate
            lastMinIdx = i
          End If
        End If
        i = i + 1
      Wend
      
      'write the initial results
      r = r + 1
      ActiveSheet.Cells(r, c) = Location
      ActiveSheet.Cells(r, c + 1) = lastMinDate
      ActiveSheet.Cells(r, c + 2) = lastMinVal
      
      'now that we have a startlocation we can start looking for the next minima
      While Not LocationDone
        lastMinVal = 99999999
        
        'find the next minimum approximately 12.5 hours later
        nextIdx = Math.Round(lastMinIdx + (12.5 / 24) / Timestep, 0)
        For i = nextIdx - SearchRadius To nextIdx + SearchRadius
          If i > myRange.Rows.Count Then
            LocationDone = True
          Else
            myDate = myRange.Cells(i, 1)
            myVal = myRange.Cells(i, j)
            If myVal < lastMinVal Then
              lastMinDate = myDate
              lastMinVal = myVal
              lastMinIdx = i
            End If
          End If
        Next
        
        'write the results
        If Not LocationDone Then
          r = r + 1
          ActiveSheet.Cells(r, c) = Location
          ActiveSheet.Cells(r, c + 1) = lastMinDate
          ActiveSheet.Cells(r, c + 2) = lastMinVal
        End If
      
      Wend
    Next
  End If
  TIDALLOWSFROMSERIES = True

End Function

Public Function TIDALLOWSFROMARRAYS(Dates() As Date, Vals() As Single, ByRef DatesLow() As Date, ByRef ValsLow() As Single) As Boolean

  Dim i As Long, j As Long
  Dim myVal As Double, Timestep As Double, SearchRadius As Integer
  Dim lastMinDate As Double, lastDate As Double, startDate As Double
  Dim lastMinVal As Double
  Dim lastMinIdx As Long, nextIdx As Long
  Dim myDate As Double, FirstDone As Boolean, Done As Boolean
  
  'initialize the size of the output array to be the same as the input array. We'll redim again later
  ReDim DatesLow(1 To UBound(Dates))
  ReDim ValsLow(1 To UBound(Vals))
    
  'first analyse the timestep in order to specify a window size.
  'then derive a search radius (expressed in n_timesteps) for each following min/max
  Timestep = Dates(2) - Dates(1)
  SearchRadius = Application.WorksheetFunction.RoundUp((2 / 24) / Timestep, 0)
  
  FirstDone = False
      
  startDate = Dates(1)
  lastMinVal = Vals(1)
  lastMinDate = Dates(1)
  lastDate = Dates(1)
  i = 1
      
  'first find the minimum in the first 13.1 hours which is a little longer than one tidal wave (12.5 hours)
  While Not FirstDone = True
    myDate = Dates(i)
    myVal = Vals(i)
    
    If (myDate - startDate) > (13.1 / 24) Then
      FirstDone = True
    ElseIf i > UBound(Dates) Then
      FirstDone = True
    Else
      If myVal < lastMinVal Then
        lastMinVal = myVal
        lastMinDate = myDate
        lastMinIdx = i
      End If
    End If
    i = i + 1
  Wend
      
  'write the initial results
  j = j + 1
  DatesLow(j) = lastMinDate
  ValsLow(j) = lastMinVal
      
  'now that we have a startlocation we can start looking for the next minima
  While Not Done
      
    'initialize the minimum value
    lastMinVal = 99999999
        
    'find the next low tide approximately 12.5 hours later
    nextIdx = Math.Round(lastMinIdx + (12.5 / 24) / Timestep, 0)
    For i = nextIdx - SearchRadius To nextIdx + SearchRadius
      If i > UBound(Dates) Then
        Done = True
      Else
        myDate = Dates(i)
        myVal = Vals(i)
        
        If myVal < lastMinVal Then
          lastMinDate = myDate
          lastMinVal = myVal
          lastMinIdx = i
        End If
      End If
    Next
        
    'write the results
    If Not Done Then
      j = j + 1
      DatesLow(j) = lastMinDate
      ValsLow(j) = lastMinVal
    End If
  Wend
  
  'define the upper boundary of the output arrays
  ReDim Preserve DatesLow(1 To j)
  ReDim Preserve ValsLow(1 To j)
  
  TIDALLOWSFROMARRAYS = True

End Function

Public Function getAvgMaxFromTide(myRange As Range) As Double

Dim Cumulative As Double, myVal As Double
Dim i As Long, n As Long

If myRange.Columns.Count = 1 Then
  For i = 3 To myRange.Rows.Count - 2
    If IsNumeric(myRange.Cells(i, 1)) Then
      myVal = myRange.Cells(i, 1)
      If myRange.Cells(i - 1, 1) < myVal And myVal >= myRange.Cells(i + 1, 1) And myRange.Cells(i - 2, 1) < myVal And myVal >= myRange.Cells(i + 2, 1) Then
        n = n + 1
        Cumulative = Cumulative + myVal
      End If
    End If
  Next i
End If

getAvgMaxFromTide = Cumulative / n

End Function

Public Function getAvgMinFromTide(myRange As Range) As Double

Dim Cumulative As Double, myVal As Double
Dim i As Long, n As Long

If myRange.Columns.Count = 1 Then
  For i = 3 To myRange.Rows.Count - 2
    If IsNumeric(myRange.Cells(i, 1)) Then
      myVal = myRange.Cells(i, 1)
      If myRange.Cells(i - 1, 1) > myVal And myVal <= myRange.Cells(i + 1, 1) And myRange.Cells(i - 2, 1) > myVal And myVal <= myRange.Cells(i + 2, 1) Then
        n = n + 1
        Cumulative = Cumulative + myVal
      End If
    End If
  Next i
End If

getAvgMinFromTide = Cumulative / n

End Function

Public Sub READHMCZDATA(path As String, TargetSheet As String, StartRow As Long, StartCol As Long, IntervalMinutes As Long)
  Dim fn As Long, fileContent As String, FileRecords() As String
  Dim i As Long, r As Long
  Dim spc1 As Long, spc2 As Long, spc3 As Long
  Dim dateStr As String, TimeStr As String, valstr As String
  Dim Uur As Long, Minuut As Long
  Dim Tijd As Double
  
  r = 0
  fn = FreeFile
  
  Open path For Input As #fn
  fileContent = Input(VBA.LOF(fn), #fn)
  FileRecords = Split(fileContent, vbLf)
  Close (fn)
  
  If WorkSheetExists(TargetSheet) Then
    For i = 0 To UBound(FileRecords())
      Dim tmpStr As String
      tmpStr = FileRecords(i)
      
      If VBA.Mid(tmpStr, 3, 1) = "-" Then
        dateStr = ParseString(tmpStr, " ")
        TimeStr = ParseString(tmpStr, " ")
        valstr = ParseString(tmpStr, " ")
        Uur = Val(ParseString(TimeStr, ":"))
        Minuut = Val(ParseString(TimeStr, ":"))
        
        Dim DateConvert As Date
        DateConvert = VBA.Format(dateStr, "dd-mm-yyyy")
        Tijd = TimeSerial(Uur, Minuut, 0)
          
        If IsNumeric(valstr) And Minuut / IntervalMinutes = VBA.Round(Minuut / IntervalMinutes, 0) Then
          r = r + 1
          Worksheets(TargetSheet).Cells(StartRow + r, StartCol) = DateConvert + Tijd
          Worksheets(TargetSheet).Cells(StartRow + r, StartCol + 1) = valstr / 100
        End If
      End If
      
    Next
  Else
    MsgBox ("Target worksheet does not exist")
  End If

  
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ReadSpecificHISResults(HisFile As String, myLoc As String, myPar As String, AllowLeftParMatch As Boolean, AllowLocWildCards As Boolean, ValueSelection As String, Multiplier As Double) As Collection

' this routine is an example of how you can use ODSSVR20.DLL to read a his file
' give the variable myHis the same structure as the class clsODSServer from the ODSSVR20.DLL library
Dim myHis As ODSSVR20.clsODSServer
Set myHis = New ODSSVR20.clsODSServer

Set ReadSpecificHISResults = New Collection
Dim myResult As clsDateValPair

Dim Values() As Single
Dim Loc() As String, par() As String, Tim() As Double
Dim LocDef() As String, ParDef() As String, TimDef() As Double
Dim nLoc As Long, nPar As Long, nTim As Long
Dim myVal As Double

' iRes is just a number that will be returned by the ODSSVR20 library. 0 means: the function was successfully called
' iLoc, iPar and iTim are paramters that will later run from 1 to the number of resp. Locations, Parameters and Timesteps
Dim iRes As Long, iLoc As Long, iPar As Long, iTim As Long
  
myHis.KeepFilesOpen = True
  
myHis.Add HisFile, HisFile, True, True                '.Add is a method that the ODSSVR library supports.
If Not myHis.Item(HisFile).Exists Then                'give an error message if hisfile does not exist
  MsgBox "HisFile does not exist: " & HisFile
  Exit Function
End If
  
'hereunder we give you two options:
'1 - to read the whole file
'2 - to first read all locations, parameters & dates/times, and then to choose for which location and parameter to retrieve data
'We've commented out the first option, but feel free to retrieve that one by removing the "'" characters
  
'----------option1-----------------------------------------------
'read the whole file
'The argument "Hisfile" goes into the function, the rest of the parameters are returned to you by the function
'if the hisfile has properly been read, the function returns a value of 0 for itself
'iRes = myHis.GetAllData(Values(), nLoc, nPar, nTim, Hisfile, , Loc(), Par(), Tim())
'If iRes <> 0 Then MsgBox "Function call GetAllData not successful."
'----------/option1-----------------------------------------------
  
'--------option 2------------------------------------------------
'first read the locations
iRes = myHis.GetLoc(HisFile, , nLoc, , Loc())
If iRes <> 0 Then MsgBox "Function call GetLoc not successful."
  
' read the parameters
iRes = myHis.GetPar(HisFile, , nPar, , par())
If iRes <> 0 Then MsgBox "Function call GetPar not successful."
  
' read the dates/times
iRes = myHis.GetTime(HisFile, , nTim, , , Tim())
If iRes <> 0 Then MsgBox "Function call GetTime not successful."

' only read the values for Loc(1) en Par(1). Because of this we'lpl redimension these variables such that they can only contain one value
ReDim LocDef(1 To 1), ParDef(1 To 1) As String, TimDef(1 To nTim) As Double

'Walk through all parameters and check if the current parameter matches the requested parameter
For iPar = LBound(par()) To UBound(par())
  If (VBA.LCase(Left(par(iPar), VBA.Len(myPar))) = VBA.LCase(myPar) And AllowLeftParMatch = True) Or (LCase(par(iPar)) = LCase(myPar)) Then

    For iLoc = LBound(Loc()) To UBound(Loc())
  
      'Check if the current location matches the requested location
      If (LCase(Loc(iLoc)) = LCase(myLoc)) Or (AllowLocWildCards = True And MATCHWILDCARD(Loc(iLoc), myLoc, False) = True) Then
      
        LocDef(1) = Loc(iLoc)
        ParDef(1) = par(iPar)
        TimDef = Tim
  
        'iRes = myHis.GetData(Values(), nLoc, nPar, nTim, HisFile, strlocdef:=LocDef(), strpardef:=ParDef(), dblTimdef:=TimDef(), strLocLst:=Loc(), strparlst:=Par(), dbltimlst:=Tim())
        iRes = myHis.GetData(Values(), nLoc, nPar, nTim, HisFile, strlocdef:=LocDef(), strpardef:=ParDef(), dblTimdef:=TimDef())
        If iRes <> 0 Then MsgBox "Function call GetData not successful."
        
        For iTim = LBound(TimDef()) To UBound(TimDef())
          myVal = Values(1, 1, iTim)
          
          'de selectie toepassen
          If VBA.LCase(ValueSelection) = "< 0" Then
            myVal = Minimum(myVal, 0)
          ElseIf VBA.LCase(ValueSelection) = "> 0" Then
            myVal = Maximum(myVal, 0)
          ElseIf VBA.LCase(ValueSelection) = "absolute" Then
            myVal = Math.Abs(myVal)
          ElseIf ValueSelection = "" Then
            'do nothing
          Else
            MsgBox ("Error: value selection was not recognized " & ValueSelection)
            End
          End If
        
          If ReadSpecificHISResults.Count >= iTim Then
            'add to existing value
            Set myResult = ReadSpecificHISResults.Item(iTim)
            myResult.Value = myResult.Value + myVal * Multiplier
          Else
            'create a new datavalue pair and add it to the collection
            Set myResult = New clsDateValPair
            myResult.Datum = Tim(iTim)
            myResult.Value = myVal * Multiplier
            Call ReadSpecificHISResults.Add(myResult)
          End If
        Next
      End If
    Next
    Exit For
  End If
Next

myHis.CloseFiles
myHis.Delete HisFile
Set myHis = Nothing
Erase Loc, par, Tim, LocDef, ParDef, TimDef, Values
  
End Function

Public Function ReadHISLocParTim(HisFile As String, ByRef Loc() As String, ByRef par() As String, ByRef Tim() As Double) As Boolean

' this routine is an example of how you can use ODSSVR20.DLL to read a his file
' give the variable myHis the same structure as the class clsODSServer from the ODSSVR20.DLL library
Dim myHis As ODSSVR20.clsODSServer
Set myHis = New ODSSVR20.clsODSServer

Set ReadSpecificHISResults = New Collection
Dim myResult As clsDateValPair
  
myHis.KeepFilesOpen = True
  
myHis.Add HisFile, HisFile, True, True                '.Add is a method that the ODSSVR library supports.
If Not myHis.Item(HisFile).Exists Then                'give an error message if hisfile does not exist
  MsgBox "HisFile does not exist: " & HisFile
  ReadHISLocParTim = False
End If
    
'first read the locations
iRes = myHis.GetLoc(HisFile, , nLoc, , Loc())
If iRes <> 0 Then MsgBox "Function call GetLoc not successful."
  
' read the parameters
iRes = myHis.GetPar(HisFile, , nPar, , par())
If iRes <> 0 Then MsgBox "Function call GetPar not successful."
  
' read the dates/times
iRes = myHis.GetTime(HisFile, , nTim, , , Tim())
If iRes <> 0 Then MsgBox "Function call GetTime not successful."

myHis.CloseFiles
myHis.Delete HisFile
Set myHis = Nothing

ReadHISLocParTim = True
  
End Function

Public Function getNodeStatsFromSobekCase(SbkCaseDir As String, ParIdx As Long) As Collection
  'geeft voor iedere locatie in een hisfile de laatste (in tijd) waarde terug
  Dim HisFile As String, tpFile As String
  HisFile = VBA.replace(SbkCaseDir & "\calcpnt.his", "\\", "\")
  tpFile = VBA.replace(SbkCaseDir & "\network.tp", "\\", "\")
  
  Dim myHis As ODSSVR20.clsODSServer
  Dim TpFileContent As clsNetworkTPFileContent
  Set myHis = New ODSSVR20.clsODSServer
  Dim Results As Collection
  Set Results = New Collection
  
  Set TpFileContent = New clsNetworkTPFileContent
  Call TpFileContent.Read(tpFile)
  
  Dim Values() As Single
  Dim Loc() As String, par() As String, Tim() As Double
  Dim LocDef() As String, ParDef() As String, TimDef() As Double
  Dim nLoc As Long, nPar As Long, nTim As Long

  ' iRes is just a number that will be returned by the ODSSVR20 library. 0 means: the function was successfully called
  ' iLoc, iPar and iTim are paramters that will later run from 1 to the number of resp. Locations, Parameters and Timesteps
  Dim iRes As Long, iLoc As Long, iPar As Long, iTim As Long
  Dim myLoc As clsSBKNodeStats
  Dim Min As Double, max As Double, avg As Double, mySum As Double
  
  myHis.KeepFilesOpen = True
  
  myHis.Add HisFile, HisFile, True, True                '.Add is a method that the ODSSVR library supports.
  If Not myHis.Item(HisFile).Exists Then                'give an error message if hisfile does not exist
    MsgBox "HisFile does not exist: " & HisFile
    Exit Function
  End If
    
  'read the whole file
  'The argument "Hisfile" goes into the function, the rest of the parameters are returned to you by the function
  'if the hisfile has properly been read, the function returns a value of 0 for itself
  iRes = myHis.GetAllData(Values(), nLoc, nPar, nTim, HisFile, , Loc(), par(), Tim())
  If iRes <> 0 Then
    MsgBox "Function call GetAllData not successful."
  Else
    For iLoc = 1 To UBound(Loc())
      max = -99999999999#
      Min = 99999999999#
      mySum = 0
      Set myLoc = New clsSBKNodeStats
      myLoc.ID = Loc(iLoc)
      myLoc.par = par(ParIdx)
      myLoc.first = Values(iLoc, ParIdx, 1)
      myLoc.Last = Values(iLoc, ParIdx, UBound(Tim))
      For iTim = 1 To UBound(Tim())
        mySum = mySum + Values(iLoc, ParIdx, iTim)
        If Values(iLoc, ParIdx, iTim) < Min Then Min = Values(iLoc, ParIdx, iTim)
        If Values(iLoc, ParIdx, iTim) > max Then max = Values(iLoc, ParIdx, iTim)
      Next
      If UBound(Tim()) > 0 Then myLoc.avg = mySum / UBound(Tim())
      myLoc.Min = Min
      myLoc.max = max
      
      'zoek nu in de Network.TP file content de X- en Y-coördinaat op
      Dim myNode As clsCFReachNode
      Set myNode = TpFileContent.FindNode(myLoc.ID)
      If Not myNode Is Nothing Then
        myLoc.x = myNode.x
        myLoc.y = myNode.y
      End If
      Call Results.Add(myLoc)
    Next
  End If
  
  Set getNodeStatsFromSobekCase = Results
    
  myHis.CloseFiles
  myHis.Delete HisFile
  Set myHis = Nothing
  Erase Loc, par, Tim, LocDef, ParDef, TimDef, Values
End Function

Public Function MergeStorageTables(Table1 As Collection, Table2 As Collection) As Collection
  'voegt twee hoogte/oppervlaktabellen samen
  'beide tabellen moeten een collection zijn van clsLevelAreaPair
  Dim iTable1 As Long
  Dim iTable2 As Long
  Dim Table1Done As Boolean, Table2Done As Boolean
  iTable1 = 0
  iTable2 = 0
  Dim test1 As clsLevelAreaPair
  Dim test2 As clsLevelAreaPair
  Dim newPair As clsLevelAreaPair
  
  Dim newTable As Collection
  Set newTable = New Collection
  
  If Table1.Count = 0 Then
    Set MergeStorageTables = Table2
    Exit Function
  ElseIf Table2.Count = 0 Then
    Set MergeStorageTables = Table1
    Exit Function
  End If
  
  'zet eerst een lijst met alle levels uit tabellen 1 en 2 op
  While Not (Table1Done And Table2Done)
    If Not Table1Done Then Set test1 = Table1(iTable1 + 1)
    If Not Table2Done Then Set test2 = Table2(iTable2 + 1)
    If Table1Done Then 'tabel 1 is al helemaal doorlopen; maak tabel2 in z'n eentje verder af
      iTable2 = iTable2 + 1
      Set newPair = New clsLevelAreaPair
      newPair.Level = Table2(iTable2).Level
      Call newTable.Add(newPair)
      If iTable2 = Table2.Count Then Table2Done = True
    ElseIf Table2Done Then 'tabel 2 is al helemaal doorlopen; maak tabel1 in z'n eentje verder af
      iTable1 = iTable1 + 1
      Set newPair = New clsLevelAreaPair
      newPair.Level = Table1(iTable1).Level
      Call newTable.Add(newPair)
      If iTable1 = Table1.Count Then Table1Done = True
    ElseIf test1.Level <= test2.Level Then
      iTable1 = iTable1 + 1
      Set newPair = New clsLevelAreaPair
      newPair.Level = Table1(iTable1).Level
      Call newTable.Add(newPair)
      If iTable1 = Table1.Count Then Table1Done = True
    Else
      iTable2 = iTable2 + 1
      Set newPair = New clsLevelAreaPair
      newPair.Level = Table2(iTable2).Level
      Call newTable.Add(newPair)
      If iTable2 = Table2.Count Then Table2Done = True
    End If
  Wend
  
  'bereken nu VBA.Middels interpolatie voor elk van de levels het bijbehorende oppervlak
  For Each newPair In newTable
    newPair.Area = InterpolateFromStorageTable(newPair.Level, Table1) + InterpolateFromStorageTable(newPair.Level, Table2)
  Next
  
  Set MergeStorageTables = newTable
  Exit Function
End Function

Public Function InterpolateFromStorageTable(myLevel As Double, myTable As Collection) As Double
  'deze functie interpoleert een level binnen een level/area table
  'geeft bijbehorend oppervlak terug
  'het is een specifieke functie omdat je aan de onderkant niet extrapoleert en aan de bovenkant het oppervlak constant houdt
  Dim myPair As clsLevelAreaPair
  Dim minVal As Double, maxVal As Double
  Dim lPair As clsLevelAreaPair, uPair As clsLevelAreaPair
  Dim i As Long
  
  Set myPair = myTable(1)
  minVal = myPair.Level
  Set myPair = myTable(myTable.Count)
  maxVal = myPair.Level
  
  If myLevel < minVal Then
    InterpolateFromStorageTable = 0 ' voor alle waarden onder de tabel: geef nul terug
    Exit Function
  ElseIf myLevel >= maxVal Then
    Set myPair = myTable(myTable.Count)
    InterpolateFromStorageTable = myPair.Area ' voor alle waarden boven de tabel: geef maximum waarde terug
    Exit Function
  Else
    'voor alle waarden binnen de tabel: interpoleren
    For i = 1 To myTable.Count - 1
      Set lPair = myTable(i)
      Set uPair = myTable(i + 1)
      If myLevel >= lPair.Level And myLevel < uPair.Level Then
        InterpolateFromStorageTable = Interpolate(lPair.Level, lPair.Area, uPair.Level, uPair.Area, myLevel)
        Exit Function
      End If
    Next
  End If
End Function


Public Function ParseSobekRecords(myPath As String, myToken As String) As Collection
  Dim fn As Long, myStr As String
  Dim fileContent As String, records As Collection
  Set records = New Collection
  
  fileContent = ReadEntireTextFile(myPath)
  records = Split(fileContent, myToken & " ", , vbBinaryCompare)
  
End Function

Public Sub ParseSobekFile(myPath As String, resultsRow As Long, resultsCol As Long)
  'leest de inhoud van de Sobek (bijv. Network.CR) file in en schrijft die naar een opgegeven locatie
  Dim fn As Long, myStr As String
  Dim r As Long, c As Long
  fn = FreeFile
  Open myPath For Input As #fn
  
  r = resultsRow - 1
  While Not EOF(fn)
    Line Input #fn, myStr
    r = r + 1
    c = resultsCol - 1
    While Not myStr = ""
      c = c + 1
      ActiveSheet.Cells(r, c) = ParseString(myStr, " ")
    Wend
  Wend

  Close (fn)
End Sub

Public Function ParseSobekTable(ByRef myRecord As String) As Double()
  
  'zoek allereerst naar "TBLE"
  Dim start As Boolean, endsign As Boolean, Done As Boolean
  Dim tmpRecord As String, tmpStr As String
  Dim r, c, nRow, nCol As Long
  myRecord = VBA.replace(myRecord, vbCrLf, " ")
  myRecord = VBA.replace(myRecord, "  ", " ")
  tmpRecord = myRecord
  
  Dim myTable() As Double
  
  c = 0
  r = 1 'ga ervan uit dat de tabel ten minste een rij bezit
  
  'eerst gaan we de dimensies van de tabel vaststellen
  nRow = 0
  nCol = 0
  Done = False
  While Not Done
    tmpStr = ParseString(tmpRecord, " ")
    If tmpStr = "TBLE" Then start = True                'begintoken voor tabel gevonden
    If tmpStr = "<" Then
      endsign = True                 'afsluitend teken voor tabelrij gevonden
      nRow = nRow + 1                'een rij gevonden, dus meteen het tellertje bijhouden
    End If
    If endsign = False And IsNumeric(tmpStr) Then nCol = nCol + 1
    If tmpRecord = "" Or tmpStr = "tble" Then Done = True 'tabel is compleet
  Wend
  
  'nu gaan we de tabel vullen
  ReDim myTable(1 To nRow, 1 To nCol)
  r = 1
  Done = False
  While Not Done
    tmpStr = ParseString(myRecord, " ")
    If tmpStr = "TBLE" Then start = True
    If tmpStr = "<" Then
      r = r + 1
      c = 0
    ElseIf IsNumeric(tmpStr) Then
      c = c + 1
      myTable(r, c) = Val(tmpStr)
    End If
    If myRecord = "" Or tmpStr = "tble" Then Done = True 'tabel is compleet
  Wend
  ParseSobekTable = myTable

End Function


Public Function ParseBySingleChar(ByRef myString As String) As String
  If VBA.Len(myString) > 0 Then
    ParseBySingleChar = VBA.Left(myString, 1)
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
  Else
    ParseBySingleChar = ""
  End If
End Function


Public Sub MakeSobekTargetLevelTable(ID As String, ZP As Double, WP As Double, StartYear As Long, EndYear As Long, resultsRow As Long, resultsCol As Long)
  Dim i As Long, r As Long, c As Long
  Dim myDate As Double
  
  r = resultsRow
  c = resultsCol
    
  ActiveSheet.Cells(r, c) = "ID"
  ActiveSheet.Cells(r, c + 1) = "datum"
  ActiveSheet.Cells(r, c + 2) = "tijd"
  ActiveSheet.Cells(r, c + 3) = "waarde"
  For i = StartYear To EndYear
    
    If i = StartYear Then
      r = r + 1
      ActiveSheet.Cells(r, c) = ID
      ActiveSheet.Cells(r, c + 1) = VBA.DateSerial(i, 1, 1)
      ActiveSheet.Cells(r, c + 2) = VBA.TimeSerial(0, 0, 0)
      ActiveSheet.Cells(r, c + 3) = WP
    End If
    
    r = r + 1
    ActiveSheet.Cells(r, c) = ID
    ActiveSheet.Cells(r, c + 1) = VBA.DateSerial(i, 3, 31)
    ActiveSheet.Cells(r, c + 2) = VBA.TimeSerial(23, 59, 0)
    ActiveSheet.Cells(r, c + 3) = WP
    r = r + 1
    ActiveSheet.Cells(r, c) = ID
    ActiveSheet.Cells(r, c + 1) = VBA.DateSerial(i, 4, 1)
    ActiveSheet.Cells(r, c + 2) = VBA.TimeSerial(0, 0, 0)
    ActiveSheet.Cells(r, c + 3) = ZP
    r = r + 1
    ActiveSheet.Cells(r, c) = ID
    ActiveSheet.Cells(r, c + 1) = VBA.DateSerial(i, 9, 30)
    ActiveSheet.Cells(r, c + 2) = VBA.TimeSerial(23, 59, 0)
    ActiveSheet.Cells(r, c + 3) = ZP
    r = r + 1
    ActiveSheet.Cells(r, c) = ID
    ActiveSheet.Cells(r, c + 1) = VBA.DateSerial(i, 10, 1)
    ActiveSheet.Cells(r, c + 2) = VBA.TimeSerial(0, 0, 0)
    ActiveSheet.Cells(r, c + 3) = WP
    
    If i = EndYear Then
      r = r + 1
      ActiveSheet.Cells(r, c) = ID
      ActiveSheet.Cells(r, c + 1) = VBA.DateSerial(i + 1, 1, 1)
      ActiveSheet.Cells(r, c + 2) = VBA.TimeSerial(0, 0, 0)
      ActiveSheet.Cells(r, c + 3) = WP
    End If
    
  Next

End Sub

Public Sub READBUIFILE(myPath As String, resultsRow As Long, resultsCol As Long)
  Dim fn As Long, myStr As String
  fn = FreeFile
  Dim RainfallData(1, 1) As Double
  Dim nStations As Long, nEvents As Long, Timestep As Long
  Dim HeaderRead As Long
  Dim myDate As Double, myYear As Long, myMonth As Long, myDay As Long, myHour As Long, myMinute As Long, mySecond As Long
  Dim r As Long, c As Long
  Dim i As Long
      
  r = resultsRow
  c = resultsCol

  'reads a .bui file (SOBEK rainfall event) and writes it to the worksheet
  Open myPath For Input As #fn
  While Not EOF(fn)
    Line Input #fn, myStr
    If VBA.Trim(VBA.LCase(myStr)) = "*aantal stations" Then
      Line Input #fn, myStr
      nStations = Val(myStr)
      HeaderRead = HeaderRead + 1
    ElseIf VBA.Trim(VBA.LCase(myStr)) = "*namen van stations" Then
      For i = 1 To nStations
        Line Input #fn, myStr
        ActiveSheet.Cells(r, c + i) = VBA.replace(myStr, "'", "")
      Next
      HeaderRead = HeaderRead + 1
    ElseIf VBA.Trim(VBA.LCase(myStr)) = "*en het aantal seconden per waarnemingstijdstap" Then
      Line Input #fn, myStr
      nEvents = VBA.Val(ParseString(myStr, " "))
      Timestep = VBA.Val(ParseString(myStr, " ")) / 3600 'convert to hours
      HeaderRead = HeaderRead + 1
    ElseIf VBA.Left(myStr, 1) = "*" Then
      'commentaarregel
    ElseIf HeaderRead >= 3 Then  'geen commentaarregels meer
      If HeaderRead = 3 Then
        myYear = VBA.Left(myStr, 4)
        myMonth = VBA.Mid(myStr, 5, 2)
        myDay = VBA.Mid(myStr, 7, 2)
        myHour = VBA.Mid(myStr, 9, 2)
        myMinute = VBA.Mid(myStr, 11, 2)
        mySecond = VBA.Mid(myStr, 13, 2)
        myDate = DateSerial(myYear, myMonth, myDay) + TimeSerial(myHour, myMinute, mySecond)
        'set it back one timestep before reading the first line
        myDate = myDate - Timestep / 24
        HeaderRead = HeaderRead + 1
      Else
        'nieuw record
        r = r + 1
        i = 0
        myDate = myDate + Timestep / 24
        ActiveSheet.Cells(r, c) = myDate
        While Not myStr = ""
          i = i + 1
          ActiveSheet.Cells(r, c + i) = VBA.Val(ParseString(myStr, " "))
        Wend
      End If
    End If
  Wend
  Close (fn)
End Sub

Public Sub WriteBuiFile(path As String, DataBlock As Range, TSSecs As Integer, ProgressRange As Range)

Dim startDate As Double, endDate As Double
Dim DurDays As Long, DurHours As Long, DurMins As Long, DurSecs As Long, rest As Double, TotDur As Double
Dim fn As Long
Dim r As Long, c As Long, i As Long
Dim stations As Collection
Set stations = New Collection
Dim myDate As Double, myStr As String
Dim TSMins As Long

TSMins = TSSecs / 60

'belangrijk: bovenste rij bevat neerslagstations, linker kolom bevat datum/tijd

'haal gegevens voor de bui op
startDate = DataBlock.Cells(2, 1)
endDate = DataBlock.Cells(DataBlock.Rows.Count, 1)

TotDur = (DataBlock.Rows.Count - 1) * TSMins 'totale duur in minuten
DurDays = WorksheetFunction.RoundDown(TotDur / 60 / 24, 0)
rest = TotDur - DurDays * 24 * 60
DurHours = WorksheetFunction.RoundDown(rest / 60, 0)
rest = rest - DurHours * 60
DurMins = WorksheetFunction.RoundDown(rest, 0)
rest = rest - DurMins
DurSecs = rest * 60

'enventariseer de neerslagstations
For c = 2 To DataBlock.Columns.Count
  stations.Add DataBlock.Cells(1, c)
Next

fn = FreeFile
Open path For Output As #fn
  Print #fn, "*Name of this file: " & path
  Print #fn, "*Date and time of construction: "
  Print #fn, "1"
  Print #fn, "*Aantal stations"
  Print #fn, stations.Count
  Print #fn, "*Namen van stations"
  Dim myStation As Variant
  For Each myStation In stations
    Print #fn, "'" & myStation & "'"
  Next
  Print #fn, "*Aantal gebeurtenissen (omdat het 1 bui betreft is dit altijd 1)"
  Print #fn, "*en het aantal seconden per waarnemingstijdstap"
  Print #fn, " 1  3600 "
  Print #fn, "*Elke commentaarregel wordt begonnen met een * (asteriks)."
  Print #fn, "*Eerste record bevat startdatum en -tijd, lengte van de gebeurtenis in dd hh mm ss"
  Print #fn, "*Het VBA.Format is: yyyymmdd:hhmmss:ddhhmmss"
  Print #fn, "*Daarna voor elk station de neerslag in mm per tijdstap."
  Print #fn, " " & Year(startDate) & " " & Month(startDate) & " " & Day(startDate) & " " & Hour(startDate) & " " & Minute(startDate) & " " & Second(startDate) & " " & DurDays & " " & DurHours & " " & DurMins & " " & DurSecs

  For r = 2 To DataBlock.Rows.Count
  
    If Math.Round(r / 1000, 0) * 1000 = r Then
      ProgressRange.Cells(1, 1) = (r - 1) / DataBlock.Rows.Count
    End If
  
    myDate = DataBlock.Cells(r, 1)
    If myDate >= startDate And myDate <= endDate Then
      myStr = ""
      For i = 1 To stations.Count
        myStr = myStr & " " & DataBlock.Cells(r, 1 + i)
      Next
      myStr = VBA.Trim(myStr)
      Print #fn, myStr
    End If
  Next


Close (fn)
End Sub

Public Sub WriteRKSFile(path As String, DataBlock As Range, StartEndDatesBlock As Range, ProgressRange As Range, TSSecs As Integer)

Dim startDate As Double, endDate As Double
Dim DurDays As Long, DurHours As Long, DurMins As Long, DurSecs As Long, rest As Double, TotDur As Double
Dim fn As Long
Dim r As Long, c As Long, i As Long, j As Long, K As Long
Dim stations As Collection
Set stations = New Collection
Dim myDate As Double, myStr As String
Dim TSMins As Long

TSMins = TSSecs / 60

'IMPORTANT: bovenste rij DataBlock bevat neerslagstations, linker kolom bevat datum/tijd
'IMPORTANT: StartEndDatesBlock bevat 3 kolommen: links het nummer van de bui, midden de startdatum, rechts einddatum. Geen header

If StartEndDatesBlock.Columns.Count <> 3 Then
  MsgBox ("Fout: StartEndDatesBlock moet drie kolommen bevatten: buinummer, startdatum, einddatum")
  End
End If

If Not IsDate(StartEndDatesBlock.Cells(1, 2)) Then
  MsgBox ("Fout: StartEndDatesBlock mag geen header bevatten. Begin meteen met de eerste bui, met in kolommen 2 en 3 start- en einddatum van de bui.")
End If

'check chronologische volgorde events
For i = 2 To StartEndDatesBlock.Rows.Count
  If StartEndDatesBlock.Cells(i, 2) <= StartEndDatesBlock.Cells(i - 1, 2) Then
    MsgBox ("Fout: het blok met start- en einddatums van buien moet in chronologische volgorde staan.")
    End
  End If
Next

'enventariseer de neerslagstations
For c = 2 To DataBlock.Columns.Count
  stations.Add DataBlock.Cells(1, c)
Next

fn = FreeFile
Open path For Output As #fn
  Print #fn, "*Name of this file: " & path
  Print #fn, "* Gebruik de default dataset voor overige invoer (altijd 1 bij bui, 0 bij reeks)"
  Print #fn, "0"
  Print #fn, "*Aantal stations"
  Print #fn, stations.Count
  Print #fn, "*Namen van de stations"
  Dim myStation As Variant
  For Each myStation In stations
    Print #fn, "'" & myStation & "'"
  Next
  Print #fn, "* Number of events in series and time step size [s]"
  Print #fn, StartEndDatesBlock.Rows.Count & " " & TSSecs
  
  'read each of the start- and enddates
  For i = 1 To StartEndDatesBlock.Rows.Count
    ProgressRange.Cells(1, 1) = i / StartEndDatesBlock.Rows.Count
  
    Print #fn, "* Event " & StartEndDatesBlock.Cells(i, 1) & " duration   " & (StartEndDatesBlock.Cells(i, 3) - StartEndDatesBlock.Cells(i, 2)) * 24 & " [hours]"
    Print #fn, "* Start date and time of the event: yyyy mm dd hh mm ss"
    Print #fn, "* Duration of the event           : dd hh mm ss"
    Print #fn, "* Rainfall value per time step [mm/time step]"
    
    'haal gegevens voor de bui op
    startDate = StartEndDatesBlock.Cells(i, 2)
    endDate = StartEndDatesBlock.Cells(i, 3)

    'bereken de duur van deze bui in dagen, uren, minuten en seconden
    TotDur = (StartEndDatesBlock.Cells(i, 3) - StartEndDatesBlock.Cells(i, 2)) * 24 * 60 'totale duur in minuten
    DurDays = WorksheetFunction.RoundDown(TotDur / 60 / 24, 0)
    rest = TotDur - DurDays * 24 * 60
    DurHours = WorksheetFunction.RoundDown(rest / 60, 0)
    rest = rest - DurHours * 60
    DurMins = WorksheetFunction.RoundDown(rest, 0)
    rest = rest - DurMins
    DurSecs = rest * 60
    
    Print #fn, " " & Year(startDate) & " " & Month(startDate) & " " & Day(startDate) & " " & Hour(startDate) & " " & Minute(startDate) & " " & Second(startDate) & " " & DurDays & " " & DurHours & " " & DurMins & " " & DurSecs
        
    'zoek nu in het datablok de startdatum en schrijf waarden weg
    For j = 2 To DataBlock.Rows.Count
      If DataBlock.Cells(j, 1) >= startDate Then
        If DataBlock.Cells(j, 1) < endDate Then
          myStr = ""
          For K = 1 To stations.Count
            myStr = myStr & " " & DataBlock.Cells(j, 1 + K)
          Next
          myStr = VBA.Trim(myStr)
          Print #fn, myStr
        Else
          Exit For
        End If
      End If
    Next
  Next
  
Close (fn)
End Sub



Public Function WritePRNFile(path As String, DateValueRange As Range, IncludesHeader As Boolean, DateColIdx As Integer, ValColIdx As Integer) As Boolean

Dim i As Long, fn As Long, myYear As Integer, myMonth As Integer, myDay As Integer, myHour As Integer, myMin As Integer, mySec As Integer, myVal As Double
fn = FreeFile
Open path For Output As #fn

Dim StartRow As Long
If IncludesHeader Then
  StartRow = 2
Else
  StartRow = 1
End If

'"1998/01/01;00:00:00" 9.1 <
For i = StartRow To DateValueRange.Rows.Count
  myYear = Year(DateValueRange.Cells(i, DateColIdx))
  myMonth = Month(DateValueRange.Cells(i, DateColIdx))
  myDay = Day(DateValueRange.Cells(i, DateColIdx))
  myHour = Hour(DateValueRange.Cells(i, DateColIdx))
  myMin = Minute(DateValueRange.Cells(i, DateColIdx))
  mySec = Second(DateValueRange.Cells(i, DateColIdx))
  
  If IsNumeric(DateValueRange.Cells(i, ValColIdx)) Then
    myVal = DateValueRange.Cells(i, ValColIdx)
    Print #fn, Chr(34) & Format(myYear, "0000") & "/" & Format(myMonth, "00") & "/" & Format(myDay, "00") & ";" & Format(myHour, "00") & ":" & Format(myMin, "00") & ":" & Format(mySec, "00") & Chr(34) & " " & myVal & " <"
  End If
  
Next

Close (fn)
WritePRNFile = True

End Function

Public Function WritePRNFiles(OutputDir As String, DateValueRange As Range) As Variant

Dim i As Long, j As Long, fn As Long, myYear As Integer, myMonth As Integer, myDay As Integer, myHour As Integer, myMin As Integer, mySec As Integer, myVal As Double
Dim path As String

'errorhandling
If DateValueRange.Columns.Count < 2 Then
  WritePRNFiles = "Error: range must contain at least two columns"
  Exit Function
End If


For j = 2 To DateValueRange.Columns.Count
  path = OutputDir & "\" & DateValueRange.Cells(1, j) & ".prn"
  fn = FreeFile
  Open path For Output As #fn
  For i = 2 To DateValueRange.Rows.Count
    myYear = Year(DateValueRange.Cells(i, 1))
    myMonth = Month(DateValueRange.Cells(i, 1))
    myDay = Day(DateValueRange.Cells(i, 1))
    myHour = Hour(DateValueRange.Cells(i, 1))
    myMin = Minute(DateValueRange.Cells(i, 1))
    mySec = Second(DateValueRange.Cells(i, 1))
    
    If IsNumeric(DateValueRange.Cells(i, j)) Then
      myVal = DateValueRange.Cells(i, j)
      Print #fn, Chr(34) & Format(myYear, "0000") & "/" & Format(myMonth, "00") & "/" & Format(myDay, "00") & ";" & Format(myHour, "00") & ":" & Format(myMin, "00") & ":" & Format(mySec, "00") & Chr(34) & " " & myVal & " <"
    End If
  Next
  
  Close (fn)
Next

WritePRNFiles = "Complete"


End Function


Public Function WRITERRBOUNDARYDATA(myRange As Range, File3B As String, FileTBL As String) As String
  'Author: Siebe Bosch
  'Date: 21-6-2013
  'first column must contain ID
  'second column must contain summer target level
  'third column must contain winter target level
  Dim r As Long, c As Long, fn1 As Long, fn2 As Long
  Dim ID As String, ZP As Double, WP As Double
  
  fn1 = FreeFile
  Open File3B For Output As #fn1
  
  fn2 = FreeFile
  Open FileTBL For Output As #fn2
  
  For r = 1 To myRange.Rows.Count
    ID = myRange.Cells(r, 1)
    ZP = myRange.Cells(r, 2)
    WP = myRange.Cells(r, 3)
    
    'BOUN id 'rrcf121212' bl 1 'rrcf121212' is 0 boun
    Print #fn1, "BOUN id '" & ID & "' bl 1 '" & ID & "' is 0 boun"
    
    '    bn_t ID 'rrcf121212' nm 'rrcf121212' PDIN 1 1 '365;00:00:00' pdin TBLE
    '    '2000/01/01;00:00:00' -0.5 0 <
    '    '2000/04/15;00:00:00' -0.25 0 <
    '    '2000/10/15;00:00:00' -0.5 0 <
    '    tble bn_t
    
    Print #fn2, "BN_T id '" & ID & "' nm '" & ID & "' PDIN 1 1 '365;00:00:00' pdin TBLE"
    Print #fn2, "'2000/01/01;00:00:00' " & WP & " 0 <"
    Print #fn2, "'2000/04/15;00:00:00' " & ZP & " 0 <"
    Print #fn2, "'2000/10/15;00:00:00' " & WP & " 0 <"
    Print #fn2, "tble bn_t"
  
  Next
  
  
  
  Close (fn1)
  Close (fn2)
  
  WRITERRBOUNDARYDATA = "COMPLETE"
    
End Function


Public Function getDelwaqID(myNum As Integer) As String
  Dim myStr As String, myNumStr As String
  Dim i As Long
  myStr = "Segment"
  myNumStr = VBA.Trim(VBA.Str(myNum))
  For i = VBA.Len(myNumStr) + 1 To 5
    myStr = myStr & " "
  Next
  myStr = myStr & myNumStr
  getDelwaqID = myStr
End Function

Public Function IDFROMSTRING(myStr As String, Optional Prefix As String = "", Optional CutoffString As String = "") As String
  Dim CutOffPos As Long
  Dim PrefixPos As Long
  
  PrefixPos = InStr(1, myStr, Prefix)
  CutOffPos = InStr(1, myStr, CutoffString)
  
  If Prefix = "" And CutoffString = "" Then                           'geen prefix of afbreekstring opgegeven, dus retourneer de hele string
    IDFROMSTRING = myStr
  ElseIf Prefix <> "" And CutoffString = "" Then                      'wel prefix opgegeven maar geen afbreekstring
    If PrefixPos > 0 Then
      IDFROMSTRING = VBA.Right(myStr, VBA.Len(myStr) - PrefixPos + 1)         'prefix aangetroffen
    Else                                                              'prefix niet aangetroffen
      IDFROMSTRING = ""
    End If
  ElseIf Prefix = "" And CutoffString <> "" Then                      'geen prefix opgegeven, maar wel een afbreekstring
    If CutOffPos > 0 Then
      IDFROMSTRING = VBA.Left(myStr, CutOffPos - 1)                       'afbreekstring aangetroffen
    Else
      IDFROMSTRING = myStr                                            'afbreekstring niet aangetroffen, dus retourneer de hele string
    End If
  ElseIf Prefix <> "" And CutoffString <> "" Then                     'zowel prefix als afbreekstring opgegeven
    If PrefixPos > 0 And CutOffPos > 0 And CutOffPos > PrefixPos Then 'prefix en afbreekstring aangetroffen en afbreekstring ligt achter prefix
      IDFROMSTRING = VBA.Mid(myStr, PrefixPos, (CutOffPos - PrefixPos + 1))
    ElseIf prefixpost > 0 And CutOffPos = 0 Then
      IDFROMSTRING = VBA.Right(myStr, VBA.Len(myStr) - PrefixPos + 1)         'prefix aangetroffen, maar afbreekstring niet
    ElseIf PrefixPos = 0 And CutOffPos > 0 Then
      IDFROMSTRING = VBA.Left(myStr, CutOffPos - 1)                       'afbreekstring aangetroffen
    Else
      IDFROMSTRING = ""
    End If
  End If
  
End Function

Public Function RemovePostFix(myStr As String, Postfix As String) As String
  Dim Pos As Integer
  Pos = InStr(myStr, Postfix)
  If Pos > 0 Then
    RemovePostFix = Left(myStr, Pos - 1)
  Else
    RemovePostFix = myStr
  End If
End Function

Public Sub WRITESTOCHASTXMLFILE(myRange As Range, myPath As String)
  Dim r As Long, c As Long
  Dim fn As Long
  Dim myID As String, myAlias As String
  Dim myHerh As Double, myH As Double
  
  fn = FreeFile
  Open myPath For Output As #fn
  
  Print #fn, "<stochasticAnalysis>"
  
  For r = 2 To myRange.Rows.Count
    myID = myRange.Cells(r, 1)
    myAlias = myRange.Cells(r, 2)
    Print #fn, "  <location>"
    Print #fn, "    <id>" & myID & "</id>"
    Print #fn, "    <alias>" & myAlias & "</alias>"
    For c = 3 To myRange.Columns.Count
      myHerh = myRange.Cells(1, c)
      myH = myRange.Cells(r, c)
      Print #fn, "    <result>"
      Print #fn, "      <frequencyEvent>" & 1 / (myRange.Columns.Count - 2) & "</frequencyEvent>"
      Print #fn, "      <returnPeriodInYears>" & myHerh & "</returnPeriodInYears>"
      Print #fn, "      <jobname>" & myID & "_" & myAlias & "</jobname>"
      Print #fn, "      <exceedanceWaterLevel>" & myH & "</exceedanceWaterLevel>"
      Print #fn, "    </result>"
    Next
    Print #fn, "  </location>"
  Next
  Print #fn, "</stochasticAnalysis>"
  
  Close (fn)
  
End Sub

Public Sub ReplaceDatesInSettingsDat(templateFile As String, Outfile As String, startDate As Date, endDate As Date)
  'this routine replaces the start- and end date of a simulation in the settings.dat file
  'NOTE: it might be that the Delft_3B.INI file needs adjustment too!
  
  Dim fn As Integer, i As Integer, tmpStr As String
  fn = FreeFile
  Dim fileContent As String, FileRecords() As String
  
  Open templateFile For Input As #fn
    fileContent = Input(VBA.LOF(fn), #fn)
    FileRecords = Split(fileContent, vbCrLf)
  Close (fn)
    
  fn = FreeFile
  Open Outfile For Output As #fn
  For i = 0 To UBound(FileRecords) - 1
    tmpStr = replace(FileRecords(i), vbCrLf, "")
    If InStr(1, tmpStr, "BeginYear") > 0 Then
      Print #fn, "BeginYear=" & Year(startDate)
    ElseIf InStr(1, tmpStr, "BeginMonth") > 0 Then
      Print #fn, "BeginMonth=" & Month(startDate)
    ElseIf InStr(1, tmpStr, "BeginDay") > 0 Then
      Print #fn, "BeginDay=" & Day(startDate)
    ElseIf InStr(1, tmpStr, "BeginHour") > 0 Then
      Print #fn, "BeginHour=" & Hour(startDate)
    ElseIf InStr(1, tmpStr, "BeginMinute") > 0 Then
      Print #fn, "BeginMinute=" & Minute(startDate)
    ElseIf InStr(1, tmpStr, "BeginSecond") > 0 Then
      Print #fn, "BeginSecond=" & Second(startDate)
    ElseIf InStr(1, tmpStr, "EndYear") > 0 Then
      Print #fn, "EndYear=" & Year(endDate)
    ElseIf InStr(1, tmpStr, "EndMonth") > 0 Then
      Print #fn, "EndMonth=" & Month(endDate)
    ElseIf InStr(1, tmpStr, "EndDay") > 0 Then
      Print #fn, "EndDay=" & Day(endDate)
    ElseIf InStr(1, tmpStr, "EndHour") > 0 Then
      Print #fn, "EndHour=" & Hour(endDate)
    ElseIf InStr(1, tmpStr, "EndMinute") > 0 Then
      Print #fn, "EndMinute=" & Minute(endDate)
    ElseIf InStr(1, tmpStr, "EndSecond") > 0 Then
      Print #fn, "EndSecond=" & Second(endDate)
    Else
      Print #fn, tmpStr
    End If
  Next
  Close (fn)
   
End Sub

Public Sub ReplaceDatesInDelft3BINI(templateFile As String, Outfile As String, startDate As Date, endDate As Date)
  'this routine replaces the start- and end date of a simulation in the delft_3b.ini file
  'NOTE: it might be that the settings.dat file needs adjustment too!
  
  Dim fn As Integer, i As Integer, tmpStr As String
  fn = FreeFile
  Dim fileContent As String, FileRecords() As String
  
  Open templateFile For Input As #fn
    fileContent = Input(VBA.LOF(fn), #fn)
    FileRecords = Split(fileContent, vbCrLf)
  Close (fn)
    
  fn = FreeFile
  Open Outfile For Output As #fn
  For i = 0 To UBound(FileRecords) - 1
    tmpStr = replace(FileRecords(i), vbCrLf, "")
    If InStr(1, tmpStr, "StartTime") = 1 Then
      Print #fn, "StartTime='" & Format(Year(startDate), "0000") & "/" & Format(Month(startDate), "00") & "/" & Format(Day(startDate), "00") & ";" & Format(Hour(startDate), "00") & ":" & Format(Minute(startDate), "00") & ":" & Format(Second(startDate), "00") & "'"
    ElseIf InStr(1, tmpStr, "EndTime") = 1 Then
      Print #fn, "EndTime='" & Format(Year(endDate), "0000") & "/" & Format(Month(endDate), "00") & "/" & Format(Day(endDate), "00") & ";" & Format(Hour(endDate), "00") & ":" & Format(Minute(endDate), "00") & ":" & Format(Second(endDate), "00") & "'"
    Else
      Print #fn, tmpStr
    End If
  Next
  Close (fn)
   
End Sub

Public Function WritePaved3BRecord(ID As String, ar As Double, lv As Double, StorDef As String, MeteoStation As String, DWADef As String, SewageSystem As Integer, POCMixm3ps As Double, POCDWFm3ps As Double, MixToObjectType As Integer, DWFToObjectType As Integer) As String
Dim myRecord As String
Dim i As Integer
'SewageSystem: 0=mixed 1=separated 2 = improved separated
'MixToObjectType: 0 = openwater, 1 = boundary, 2 = WWTP
'DWFToObjectType: 0 = openwater, 1 = boundary, 3 = WWTP

'PAVE ID 'J1-11vh' ar 155890 lv -4.41 ss 1 sd 'stor_J1-11' qc 0 0 0 qo 1 1 ms 'J1-11' aaf 1 is 0 np 0 dw 'alg_dwa' ro 1 ru 0.2 qh '' pave
myRecord = "PAVE id '" & ID & "' ar " & ar & " lv " & lv & " sd '" & StorDef & "' ss " & SewageSystem & " qc 0 " & POCMixm3ps & " " & POCDWFm3ps & " qo " & MixToObjectType & " " & DWFToObjectType & " ms '" & MeteoStation & "' aaf 1 is 0 np 0 dw '" & DWADef & "' ro 1 ru 0.2 qh '' pave"
WritePaved3BRecord = myRecord
End Function

Public Function WriteUnpaved3BRecord(ID As String, gwArea As Double, areas As Range, lv As Double, StorDef As String, ErnstDef As String, SeepDef As String, InfDef As String, SoilType As Integer, ig As Double, MaxGW As Double, MeteoStation As String) As String
Dim myRecord As String
Dim i As Integer

'J1-11ovl' na 16 ar 2083950 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 ga 2083950 lv -4.41 co 3 rc 0 sd 'sd_J1-11ovl' ad 'P04_2' ed 'ad_J1-11ovl' sp 'sp_J1-11ovl' ic 'ic_J1-11ovl' bt 119 ig 0 1.32 su 0 '' gl 10 mg -4.41 ms 'J1-11' is 0 aaf 1 unpv
If areas.Rows.Count <> 16 Then
    MsgBox ("Landuse table requires exactly 16 entries.")
Else
    myRecord = "UNPV id '" & ID & "' na 16 ar"
    For i = 1 To areas.Rows.Count
        myRecord = myRecord & " " & areas(i, 1)
    Next
    myRecord = myRecord & " ga " & gwArea
    myRecord = myRecord & " lv " & lv & " co 3 rc 1 sd '" & StorDef & "' ad '' ed '" & ErnstDef & "' sp '" & SeepDef & "' ic '" & InfDef & "' bt " & SoilType & " ig 0 " & ig & " su 0 '' gl 10 mg " & MaxGW & " ms '" & MeteoStation & "' is 0 aaf 1 unpv"
End If
WriteUnpaved3BRecord = myRecord

End Function

 

Public Sub WriteWagModInput(path As String, StartRow As Long, DateCol As Long, PrecCol As Long, EvapCol As Long, Optional MeasCol As Long = 0)
  'schrijft een .dat file voor het Wageningen-model
  Dim fn As Long, i As Long
  Dim r As Long
  Dim myYear As String, myMonth As String, myDay As String, myHour As String
  Dim myPrec As String, myEvap As String, myMeas As String
  
  fn = FreeFile
  r = StartRow - 1
  Open path For Output As #fn
    Print #fn, "Deze file is geschreven met de ExcelFuncties van Hydroconsult.nl"
    Print #fn, "op: " & Now
    Print #fn, "                                      <-------P<-----ETp<------Qm"
    Print #fn, "datum#                                <----[mm]<----[mm]<----[mm]"
    While Not ActiveSheet.Cells(r + 1, DateCol) = ""
      r = r + 1
      myYear = VBA.Format(Year(ActiveSheet.Cells(r, DateCol)), "0000")
      myMonth = VBA.Format(Month(ActiveSheet.Cells(r, DateCol)), "00")
      myDay = VBA.Format(Day(ActiveSheet.Cells(r, DateCol)), "00")
      myHour = VBA.Str(Hour(ActiveSheet.Cells(r, DateCol)))
      myPrec = VBA.Format(ActiveSheet.Cells(r, PrecCol), "0.000")
      myEvap = VBA.Format(ActiveSheet.Cells(r, EvapCol), "0.000")
      If MeasCol > 0 Then
        myMeas = VBA.Format(ActiveSheet.Cells(r, MeasCol), "0.000")
      Else
        myMeas = "0.000"
      End If
      
      While Not VBA.Len(myHour) >= 3
        myHour = " " & myHour
      Wend
      While Not VBA.Len(myPrec) >= 14
        myPrec = " " & myPrec
      Wend
      While Not VBA.Len(myEvap) >= 9
        myEvap = " " & myEvap
      Wend
      While Not VBA.Len(myMeas) >= 9
        myMeas = " " & myMeas
      Wend
      Print #fn, myYear & "/" & myMonth & "/" & myDay & myHour & myPrec & myEvap & myMeas
      
    Wend
    
  Close (fn)
End Sub

Public Sub READWAGMODOUTPUT(Path00P As String, StartRow As Integer, StartCol As Integer, ByRef Progress As Range)
    Dim fn As Long, c As Integer, r As Long, AddPos As Integer, AddC As Integer
    Dim myStr As String, tmpStr As String
    Dim nLines As Long, iLine As Long
    r = StartRow
    c = StartCol
        
    fn = FreeFile
    Open Path00P For Input As #fn
                
    'count the number of lines
    While Not EOF(fn)
      Line Input #fn, myStr
      nLines = nLines + 1
    Wend
    
    Close #fn
    
    Open Path00P For Input As #fn
    Line Input #fn, myStr
    Line Input #fn, myStr
    Line Input #fn, myStr
    Line Input #fn, myStr
    
    'read the header and write it to the worksheet
    AddC = 0
    While Not myStr = ""
        tmpStr = ParseStringByMultiSpaces(myStr)
        ActiveSheet.Cells(r, c + AddC) = VBA.Trim(tmpStr)
        AddC = AddC + 1
    Wend
    
    'read the file content and write it to the worksheet
    While Not EOF(fn)
        AddC = 0
        iLine = iLine + 1
        
        'update the progress cell
        Progress.Cells(1, 1) = iLine / nLines
        DoEvents
        
        r = r + 1
        Line Input #fn, myStr
        While Not myStr = ""
            tmpStr = ParseStringByMultiSpaces(myStr)
            ActiveSheet.Cells(r, c + AddC) = VBA.Trim(tmpStr)
            AddC = AddC + 1
        Wend
    Wend
        

End Sub

Public Function ParseStringByMultiSpaces(ByRef myStr As String) As String
    Dim CharacterFound As Boolean
    Dim i As Integer
    For i = 1 To VBA.Len(myStr)
        If CharacterFound = False And VBA.Mid$(myStr, i, 1) <> " " Then
            CharacterFound = True
        ElseIf CharacterFound = True And VBA.Mid$(myStr, i, 1) = " " Then
            ParseStringByMultiSpaces = Left(myStr, i - 1)
            myStr = VBA.Right$(myStr, VBA.Len(myStr) - i + 1)
            Exit Function
        End If
    Next
    ParseStringByMultiSpaces = myStr
    myStr = ""
End Function

Public Sub WRITEPCRASTERXYZ(ResultsFile As String, DATARANGE As Range, XColIdx As Integer, YColIdx As Integer, ValColIdx As Integer)
  Dim fn As Long
  Dim r As Long
   
  fn = FreeFile
  Open ResultsFile For Output As #fn
    Print #fn, "field data"
    Print #fn, "3"
    Print #fn, "xcoord"
    Print #fn, "ycoord"
    Print #fn, "max"
    For r = 1 To DATARANGE.Rows.Count
      Print #fn, DATARANGE.Cells(r, XColIdx) & " " & DATARANGE.Cells(r, YColIdx) & " " & DATARANGE.Cells(r, ValColIdx)
    Next
  Close (fn)
End Sub

Public Function CropFact(myDate As Date, Crop As String) As Double
Dim CropIdx As Integer, DayNum As Integer
Dim Fact() As String, DayFacts() As String 'record voor 1 dag, gesplitst
Dim DayVals() As Double
ReDim Fact(1 To 366)

DayNum = DayNumber(myDate, True)

If LCase(Crop) = "grass" Then
  CropIdx = 1
ElseIf LCase(Crop) = "corn" Then
  CropIdx = 2
ElseIf LCase(Crop) = "potatoes" Then
  CropIdx = 3
ElseIf LCase(Crop) = "sugarbeet" Then
  CropIdx = 4
ElseIf LCase(Crop) = "grain" Then
  CropIdx = 5
ElseIf LCase(Crop) = "miscellaneous" Then
  CropIdx = 6
ElseIf LCase(Crop) = "non-arable land" Then
  CropIdx = 7
ElseIf LCase(Crop) = "greenhouse area" Then
  CropIdx = 8
ElseIf LCase(Crop) = "orchard" Then
  CropIdx = 9
ElseIf LCase(Crop) = "bulbous plants" Then
  CropIdx = 10
ElseIf LCase(Crop) = "foliage forest" Then
  CropIdx = 11
ElseIf LCase(Crop) = "pine forest" Then
  CropIdx = 12
ElseIf LCase(Crop) = "nature" Then
  CropIdx = 13
ElseIf LCase(Crop) = "fallow" Then
  CropIdx = 14
ElseIf LCase(Crop) = "vegetables" Then
  CropIdx = 15
ElseIf LCase(Crop) = "flowers" Then
  CropIdx = 16
End If

Fact(1) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(2) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(3) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(4) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(5) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(6) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(7) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(8) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(9) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(10) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(11) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(12) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(13) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(14) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(15) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(16) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(17) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(18) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(19) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(20) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(21) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(22) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(23) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(24) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(25) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(26) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(27) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(28) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(29) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(30) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(31) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(32) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(33) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(34) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(35) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(36) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(37) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(38) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(39) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(40) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(41) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(42) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(43) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(44) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(45) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(46) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(47) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(48) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(49) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(50) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(51) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(52) = "0.95,0.40,0.40,0.40,0.40,0.95,0.40,0.00,0.40,0.40,0.90,1.20,0.95,0.40,0.40,0.00"
Fact(53) = "0.95,0.40,0.40,0.40,0.40,0.95,0.40,0.00,0.40,0.40,0.90,1.20,0.95,0.40,0.40,0.00"
Fact(54) = "0.95,0.40,0.40,0.40,0.40,0.95,0.40,0.00,0.40,0.40,0.90,1.20,0.95,0.40,0.40,0.00"
Fact(55) = "0.95,0.40,0.40,0.40,0.40,0.95,0.40,0.00,0.40,0.40,0.90,1.20,0.95,0.40,0.40,0.00"
Fact(56) = "0.95,0.40,0.40,0.40,0.40,0.95,0.40,0.00,0.40,0.40,0.90,1.20,0.95,0.40,0.40,0.00"
Fact(57) = "0.95,0.40,0.40,0.40,0.40,0.95,0.40,0.00,0.40,0.40,0.90,1.20,0.95,0.40,0.40,0.00"
Fact(58) = "0.95,0.40,0.40,0.40,0.40,0.95,0.40,0.00,0.40,0.40,0.90,1.20,0.95,0.40,0.40,0.00"
Fact(59) = "0.95,0.40,0.40,0.40,0.40,0.95,0.40,0.00,0.40,0.40,0.90,1.20,0.95,0.40,0.40,0.00"
Fact(60) = "0.95,0.40,0.40,0.40,0.40,0.95,0.40,0.00,0.40,0.40,0.90,1.20,0.95,0.40,0.40,0.00"
Fact(61) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(62) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(63) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(64) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(65) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(66) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(67) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(68) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(69) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(70) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(71) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(72) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(73) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(74) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(75) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(76) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(77) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(78) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(79) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(80) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(81) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(82) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(83) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(84) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(85) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(86) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(87) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(88) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(89) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(90) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(91) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(92) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(93) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(94) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(95) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(96) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(97) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(98) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(99) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(100) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(101) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(102) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(103) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(104) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(105) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(106) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(107) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(108) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(109) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(110) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(111) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(112) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(113) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(114) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(115) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(116) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(117) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(118) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(119) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(120) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(121) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(122) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(123) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(124) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(125) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(126) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(127) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(128) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(129) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(130) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(131) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(132) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(133) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(134) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(135) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(136) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(137) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(138) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(139) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(140) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(141) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(142) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(143) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(144) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(145) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(146) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(147) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(148) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(149) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(150) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(151) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(152) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(153) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(154) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(155) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(156) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(157) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(158) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(159) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(160) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(161) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(162) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(163) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(164) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(165) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(166) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(167) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(168) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(169) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(170) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(171) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(172) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(173) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(174) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(175) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(176) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(177) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(178) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(179) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(180) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(181) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(182) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(183) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(184) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(185) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(186) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(187) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(188) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(189) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(190) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(191) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(192) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(193) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(194) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(195) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(196) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(197) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(198) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(199) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(200) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(201) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(202) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(203) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(204) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(205) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(206) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(207) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(208) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(209) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(210) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(211) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(212) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(213) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(214) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(215) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(216) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(217) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(218) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(219) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(220) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(221) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(222) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(223) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(224) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(225) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(226) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(227) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(228) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(229) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(230) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(231) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(232) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(233) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(234) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(235) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(236) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(237) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(238) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(239) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(240) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(241) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(242) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(243) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(244) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.05,1.20,0.90,0.25,0.25,0.00"
Fact(245) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(246) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(247) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(248) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(249) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(250) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(251) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(252) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(253) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(254) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(255) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(256) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(257) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(258) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(259) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(260) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(261) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(262) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(263) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(264) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(265) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(266) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(267) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(268) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(269) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(270) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(271) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(272) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(273) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(274) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(275) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(276) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(277) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(278) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(279) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(280) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(281) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(282) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(283) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(284) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(285) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(286) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(287) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(288) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(289) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(290) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(291) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(292) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(293) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(294) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(295) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(296) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(297) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(298) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(299) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(300) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(301) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(302) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(303) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(304) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(305) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(306) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(307) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(308) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(309) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(310) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(311) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(312) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(313) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(314) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(315) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(316) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(317) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(318) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(319) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(320) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(321) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(322) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(323) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(324) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(325) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(326) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(327) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(328) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(329) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(330) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(331) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(332) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(333) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(334) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(335) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(336) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(337) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(338) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(339) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(340) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(341) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(342) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(343) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(344) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(345) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(346) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(347) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(348) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(349) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(350) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(351) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(352) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(353) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(354) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(355) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(356) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(357) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(358) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(359) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(360) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(361) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(362) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(363) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(364) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(365) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(366) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"

DayFacts = VBA.Split(Fact(DayNum), ",")
CropFact = DayFacts(CropIdx - 1)


End Function

Public Function MAKKINKAVG(myDate As Date) As Double

Dim MAK() As Double
ReDim MAK(1 To 12, 1 To 31)
Dim myDay As Integer, myMonth As Integer
myDay = Day(myDate)
myMonth = Month(myMonth)

MAK(1, 1) = 0.2
MAK(1, 2) = 0.167
MAK(1, 3) = 0.197
MAK(1, 4) = 0.24
MAK(1, 5) = 0.163
MAK(1, 6) = 0.22
MAK(1, 7) = 0.227
MAK(1, 8) = 0.22
MAK(1, 9) = 0.19
MAK(1, 10) = 0.2
MAK(1, 11) = 0.22
MAK(1, 12) = 0.23
MAK(1, 13) = 0.29
MAK(1, 14) = 0.253
MAK(1, 15) = 0.23
MAK(1, 16) = 0.207
MAK(1, 17) = 0.267
MAK(1, 18) = 0.317
MAK(1, 19) = 0.243
MAK(1, 20) = 0.267
MAK(1, 21) = 0.233
MAK(1, 22) = 0.19
MAK(1, 23) = 0.257
MAK(1, 24) = 0.247
MAK(1, 25) = 0.203
MAK(1, 26) = 0.323
MAK(1, 27) = 0.273
MAK(1, 28) = 0.29
MAK(1, 29) = 0.353
MAK(1, 30) = 0.343
MAK(1, 31) = 0.39
MAK(2, 1) = 0.357
MAK(2, 2) = 0.403
MAK(2, 3) = 0.37
MAK(2, 4) = 0.403
MAK(2, 5) = 0.427
MAK(2, 6) = 0.367
MAK(2, 7) = 0.357
MAK(2, 8) = 0.44
MAK(2, 9) = 0.467
MAK(2, 10) = 0.433
MAK(2, 11) = 0.44
MAK(2, 12) = 0.437
MAK(2, 13) = 0.597
MAK(2, 14) = 0.56
MAK(2, 15) = 0.487
MAK(2, 16) = 0.603
MAK(2, 17) = 0.5
MAK(2, 18) = 0.517
MAK(2, 19) = 0.587
MAK(2, 20) = 0.617
MAK(2, 21) = 0.583
MAK(2, 22) = 0.647
MAK(2, 23) = 0.697
MAK(2, 24) = 0.713
MAK(2, 25) = 0.67
MAK(2, 26) = 0.713
MAK(2, 27) = 0.647
MAK(2, 28) = 0.69
MAK(2, 29) = 0.729
MAK(3, 1) = 0.753
MAK(3, 2) = 0.68
MAK(3, 3) = 0.76
MAK(3, 4) = 0.727
MAK(3, 5) = 0.9
MAK(3, 6) = 0.907
MAK(3, 7) = 0.793
MAK(3, 8) = 0.903
MAK(3, 9) = 0.807
MAK(3, 10) = 0.973
MAK(3, 11) = 0.837
MAK(3, 12) = 1
MAK(3, 13) = 0.917
MAK(3, 14) = 0.977
MAK(3, 15) = 0.89
MAK(3, 16) = 0.98
MAK(3, 17) = 0.94
MAK(3, 18) = 0.99
MAK(3, 19) = 0.903
MAK(3, 20) = 1.127
MAK(3, 21) = 1.083
MAK(3, 22) = 1.06
MAK(3, 23) = 1.163
MAK(3, 24) = 1.157
MAK(3, 25) = 1.18
MAK(3, 26) = 1.173
MAK(3, 27) = 1.223
MAK(3, 28) = 1.293
MAK(3, 29) = 1.42
MAK(3, 30) = 1.343
MAK(3, 31) = 1.32
MAK(4, 1) = 1.283
MAK(4, 2) = 1.35
MAK(4, 3) = 1.473
MAK(4, 4) = 1.28
MAK(4, 5) = 1.38
MAK(4, 6) = 1.403
MAK(4, 7) = 1.48
MAK(4, 8) = 1.473
MAK(4, 9) = 1.89
MAK(4, 10) = 1.747
MAK(4, 11) = 1.643
MAK(4, 12) = 1.553
MAK(4, 13) = 1.817
MAK(4, 14) = 1.893
MAK(4, 15) = 1.877
MAK(4, 16) = 1.707
MAK(4, 17) = 1.84
MAK(4, 18) = 1.787
MAK(4, 19) = 1.87
MAK(4, 20) = 1.92
MAK(4, 21) = 1.847
MAK(4, 22) = 2.193
MAK(4, 23) = 1.84
MAK(4, 24) = 2.273
MAK(4, 25) = 2.333
MAK(4, 26) = 2#
MAK(4, 27) = 2.203
MAK(4, 28) = 2.067
MAK(4, 29) = 2.22
MAK(4, 30) = 2.267
MAK(5, 1) = 2.243
MAK(5, 2) = 2.323
MAK(5, 3) = 2.23
MAK(5, 4) = 2.26
MAK(5, 5) = 2.337
MAK(5, 6) = 2.18
MAK(5, 7) = 2.303
MAK(5, 8) = 2.4
MAK(5, 9) = 2.553
MAK(5, 10) = 2.403
MAK(5, 11) = 2.647
MAK(5, 12) = 2.687
MAK(5, 13) = 2.583
MAK(5, 14) = 2.783
MAK(5, 15) = 2.803
MAK(5, 16) = 2.91
MAK(5, 17) = 2.793
MAK(5, 18) = 2.84
MAK(5, 19) = 3.007
MAK(5, 20) = 2.707
MAK(5, 21) = 2.547
MAK(5, 22) = 2.953
MAK(5, 23) = 2.727
MAK(5, 24) = 2.737
MAK(5, 25) = 2.72
MAK(5, 26) = 2.887
MAK(5, 27) = 2.723
MAK(5, 28) = 2.737
MAK(5, 29) = 2.93
MAK(5, 30) = 3.157
MAK(5, 31) = 2.91
MAK(6, 1) = 3.063
MAK(6, 2) = 2.783
MAK(6, 3) = 2.39
MAK(6, 4) = 2.773
MAK(6, 5) = 2.94
MAK(6, 6) = 2.663
MAK(6, 7) = 2.533
MAK(6, 8) = 2.853
MAK(6, 9) = 3.09
MAK(6, 10) = 3.123
MAK(6, 11) = 2.867
MAK(6, 12) = 3.263
MAK(6, 13) = 3.353
MAK(6, 14) = 3.06
MAK(6, 15) = 3.02
MAK(6, 16) = 2.807
MAK(6, 17) = 3.063
MAK(6, 18) = 2.74
MAK(6, 19) = 2.877
MAK(6, 20) = 3.023
MAK(6, 21) = 3.16
MAK(6, 22) = 2.59
MAK(6, 23) = 3.15
MAK(6, 24) = 2.757
MAK(6, 25) = 2.76
MAK(6, 26) = 3.053
MAK(6, 27) = 2.613
MAK(6, 28) = 2.673
MAK(6, 29) = 2.683
MAK(6, 30) = 3.293
MAK(7, 1) = 3.067
MAK(7, 2) = 3.01
MAK(7, 3) = 3.163
MAK(7, 4) = 3.4
MAK(7, 5) = 3.34
MAK(7, 6) = 3.173
MAK(7, 7) = 3.327
MAK(7, 8) = 3.173
MAK(7, 9) = 2.947
MAK(7, 10) = 3.013
MAK(7, 11) = 3.103
MAK(7, 12) = 3.383
MAK(7, 13) = 3.033
MAK(7, 14) = 2.887
MAK(7, 15) = 2.88
MAK(7, 16) = 2.513
MAK(7, 17) = 2.757
MAK(7, 18) = 2.683
MAK(7, 19) = 2.713
MAK(7, 20) = 2.643
MAK(7, 21) = 2.613
MAK(7, 22) = 2.8
MAK(7, 23) = 2.997
MAK(7, 24) = 2.787
MAK(7, 25) = 2.653
MAK(7, 26) = 2.453
MAK(7, 27) = 2.54
MAK(7, 28) = 2.72
MAK(7, 29) = 2.943
MAK(7, 30) = 2.85
MAK(7, 31) = 2.85
MAK(8, 1) = 2.717
MAK(8, 2) = 2.763
MAK(8, 3) = 2.787
MAK(8, 4) = 2.85
MAK(8, 5) = 2.747
MAK(8, 6) = 2.98
MAK(8, 7) = 2.77
MAK(8, 8) = 2.44
MAK(8, 9) = 2.67
MAK(8, 10) = 2.597
MAK(8, 11) = 2.53
MAK(8, 12) = 2.573
MAK(8, 13) = 2.707
MAK(8, 14) = 2.797
MAK(8, 15) = 2.653
MAK(8, 16) = 2.557
MAK(8, 17) = 2.393
MAK(8, 18) = 2.52
MAK(8, 19) = 2.59
MAK(8, 20) = 2.447
MAK(8, 21) = 2.47
MAK(8, 22) = 2.28
MAK(8, 23) = 2.407
MAK(8, 24) = 2.4
MAK(8, 25) = 2.427
MAK(8, 26) = 2.383
MAK(8, 27) = 2.273
MAK(8, 28) = 2.263
MAK(8, 29) = 2.32
MAK(8, 30) = 2.22
MAK(8, 31) = 1.957
MAK(9, 1) = 1.877
MAK(9, 2) = 1.88
MAK(9, 3) = 1.877
MAK(9, 4) = 1.887
MAK(9, 5) = 1.86
MAK(9, 6) = 1.987
MAK(9, 7) = 2#
MAK(9, 8) = 1.977
MAK(9, 9) = 1.787
MAK(9, 10) = 1.673
MAK(9, 11) = 1.657
MAK(9, 12) = 1.71
MAK(9, 13) = 1.577
MAK(9, 14) = 1.547
MAK(9, 15) = 1.49
MAK(9, 16) = 1.48
MAK(9, 17) = 1.487
MAK(9, 18) = 1.523
MAK(9, 19) = 1.68
MAK(9, 20) = 1.57
MAK(9, 21) = 1.547
MAK(9, 22) = 1.483
MAK(9, 23) = 1.497
MAK(9, 24) = 1.437
MAK(9, 25) = 1.177
MAK(9, 26) = 1.263
MAK(9, 27) = 1.333
MAK(9, 28) = 1.403
MAK(9, 29) = 1.343
MAK(9, 30) = 1.093
MAK(10, 1) = 1.327
MAK(10, 2) = 1.257
MAK(10, 3) = 1.163
MAK(10, 4) = 1.213
MAK(10, 5) = 1.09
MAK(10, 6) = 0.98
MAK(10, 7) = 1.017
MAK(10, 8) = 0.937
MAK(10, 9) = 0.917
MAK(10, 10) = 1
MAK(10, 11) = 1.037
MAK(10, 12) = 1
MAK(10, 13) = 0.913
MAK(10, 14) = 0.997
MAK(10, 15) = 0.85
MAK(10, 16) = 0.837
MAK(10, 17) = 0.877
MAK(10, 18) = 0.85
MAK(10, 19) = 0.91
MAK(10, 20) = 0.8
MAK(10, 21) = 0.817
MAK(10, 22) = 0.817
MAK(10, 23) = 0.743
MAK(10, 24) = 0.827
MAK(10, 25) = 0.677
MAK(10, 26) = 0.68
MAK(10, 27) = 0.623
MAK(10, 28) = 0.553
MAK(10, 29) = 0.643
MAK(10, 30) = 0.533
MAK(10, 31) = 0.5
MAK(11, 1) = 0.577
MAK(11, 2) = 0.507
MAK(11, 3) = 0.483
MAK(11, 4) = 0.53
MAK(11, 5) = 0.507
MAK(11, 6) = 0.453
MAK(11, 7) = 0.443
MAK(11, 8) = 0.43
MAK(11, 9) = 0.44
MAK(11, 10) = 0.38
MAK(11, 11) = 0.343
MAK(11, 12) = 0.487
MAK(11, 13) = 0.43
MAK(11, 14) = 0.367
MAK(11, 15) = 0.35
MAK(11, 16) = 0.323
MAK(11, 17) = 0.31
MAK(11, 18) = 0.32
MAK(11, 19) = 0.273
MAK(11, 20) = 0.323
MAK(11, 21) = 0.3
MAK(11, 22) = 0.287
MAK(11, 23) = 0.243
MAK(11, 24) = 0.297
MAK(11, 25) = 0.253
MAK(11, 26) = 0.227
MAK(11, 27) = 0.247
MAK(11, 28) = 0.267
MAK(11, 29) = 0.293
MAK(11, 30) = 0.25
MAK(12, 1) = 0.257
MAK(12, 2) = 0.21
MAK(12, 3) = 0.257
MAK(12, 4) = 0.23
MAK(12, 5) = 0.25
MAK(12, 6) = 0.24
MAK(12, 7) = 0.223
MAK(12, 8) = 0.207
MAK(12, 9) = 0.23
MAK(12, 10) = 0.187
MAK(12, 11) = 0.21
MAK(12, 12) = 0.183
MAK(12, 13) = 0.177
MAK(12, 14) = 0.207
MAK(12, 15) = 0.187
MAK(12, 16) = 0.16
MAK(12, 17) = 0.19
MAK(12, 18) = 0.177
MAK(12, 19) = 0.19
MAK(12, 20) = 0.19
MAK(12, 21) = 0.21
MAK(12, 22) = 0.163
MAK(12, 23) = 0.177
MAK(12, 24) = 0.177
MAK(12, 25) = 0.187
MAK(12, 26) = 0.183
MAK(12, 27) = 0.197
MAK(12, 28) = 0.187
MAK(12, 29) = 0.16
MAK(12, 30) = 0.193
MAK(12, 31) = 0.18

MAKKINKAVG = MAK(myMonth, myDay)

End Function

Public Sub DAYSTOHOURS(myRange As Range, resultsRow As Long, resultsol As Long, compOption As String)
    Dim i As Long, j As Long, h As Long
    Dim curDate As Date, curVal As Double
    Dim newDate As Date, newVal As Double
    
    'compOption can have the following values:
    'none
    'divide
    
    
    ActiveSheet.Cells(resultsRow, resultsCol) = "Datum/Tijd"
    ActiveSheet.Cells(resultsRow, resultsCol + 1) = "Waarde"
    If myRange.Columns.Count <> 2 Then
      MsgBox ("Error: het bereik met gegevens moet twee kolommen bevatten: datum en waarde")
    ElseIf compOption = "none" Or compOption = "divide" Then
    
      For i = 1 To myRange.Rows.Count
        curDate = myRange.Cells(i, 1)
        curVal = myRange.Cells(i, 2)
        For h = 0 To 23
          newDate = curDate + h / 24
          If compOption = "divide" Then
            newVal = curVal / 24
          Else
            newVal = curVal
          End If
          resultsRow = resultsRow + 1
          ActiveSheet.Cells(resultsRow, resultsCol) = newDate
          ActiveSheet.Cells(resultsRow, resultsCol + 1) = newVal
        Next
      Next
    Else
      MsgBox ("Error: de variabele compOption moet een van de volgende waarden hebben: none of divide")
    End If
    
End Sub


Public Sub EVAPDAYTOHOUR(DateValuesRange As Range, resultsRow As Long, resultsCol As Long)
  'deze routine disaggregeert etmaalverdampingssommen naar uurcijfers
  'en hanteert hiervoor een sinusfunctie
  Dim r1 As Long, r2 As Long, r3 As Long
  Dim myDate As Date, myVal As Double
  Dim newDate As Date, newVal As Double
  Dim cyclus As Double
  
  ActiveSheet.Cells(resultsRow, resultsCol) = "Datum/Tijd"
  ActiveSheet.Cells(resultsRow, resultsCol + 1) = "Uurwaarde verdamping"
  r3 = resultsRow
  
  For r1 = 1 To DateValuesRange.Rows.Count
    If IsDate(DateValuesRange.Cells(r1, 1)) Then
      myDate = DateValuesRange.Cells(r1, 1)
      myVal = DateValuesRange.Cells(r1, 2)
      
      For r2 = 0 To 23
        cyclus = (-6 + r2) / 24 * 2 * 3.141592 'de positie in de dagelijkse cyclus
        newVal = myVal / 24 * (Math.Sin(cyclus) + 1)
        r3 = r3 + 1
        ActiveSheet.Cells(r3, resultsCol) = myDate + r2 / 24
        ActiveSheet.Cells(r3, resultsCol + 1) = newVal
      Next
    End If
  Next

End Sub

Public Function Neerslagtekort(P As Double, e As Double, LastTekort As Double, GewasFactor As Double) As Double
  'berekent het neerslagtekort van een gegeven tijdstip met neerslag en verdamping
  Dim NewTekort As Double
  NewTekort = LastTekort + e * GewasFactor - P 'neerslagtekort = vorig tekort - neerslag + verdamping
  If NewTekort < 0 Then NewTekort = 0          'aanname: overtollige neerslag wordt meteen afgevoerd, dus een reset naar 0
  Neerslagtekort = NewTekort
End Function

Public Function HIRLAMTRANSLATE(GDALBinDir As String, SourceDir As String, TargetDir As String, SourceProj As String, TargetProj As String, GegevensBandCurrentFiles As Integer, GegevensBandPredictionFiles, myDate As Double)
  Dim i As Long, j As Long, K As Long, L As Long
  Dim InFile As String, outDir As String, Outfile As String, outFile2 As String
  Dim dateStr As String, curDateStr As String, tmpStr As String, curDate As Double, predictHour As Integer
  Dim myCollection As Collection
  
  Call ShellandWait("setx PATH " & Chr(34) & "C:\GDAL\bin" & Chr(34))
  
  Set myCollection = New Collection
  Set myCollection = ListFilesInFolder(SourceDir)
  For i = 1 To myCollection.Count
    
    'leid de huidige datum/tijd af
    curDateStr = myCollection(i)
    tmpStr = ParseString(curDateStr, "_")
    tmpStr = ParseString(curDateStr, "_")
    tmpStr = ParseString(curDateStr, "_")
    tmpStr = ParseString(curDateStr, "_")
    tmpStr = ParseString(curDateStr, "_")
    tmpStr = ParseString(curDateStr, "_")
    curDate = DATEFROMSTRING(tmpStr, "yyyymmddhh")
    
    If curDate = myDate Then
    
      'leid de voorspelhorizon van dit bestand af
      tmpStr = ParseString(curDateStr, "_")
      predictHour = tmpStr
    
      'maak nu een uitvoerdirectory aan voor deze datum/tijd, transformeer het bestand en schrijf het ernaar weg
      dateStr = Year(curDate) & VBA.Format(Month(curDate), "00") & VBA.Format(Day(curDate), "00") & VBA.Format(Hour(curDate), "00")
      InFile = SourceDir & "\" & myCollection(i)
      
      outDir = TargetDir & "\" & dateStr & "\"
      If Not DirectoryExists(outDir) Then Call VBA.MkDir(outDir)

      Outfile = outDir & VBA.Format(predictHour, "000") & ".tif"
      outFile2 = outDir & VBA.Format(predictHour, "000") & ".asc"
    
      'set the path environment and convert the grids
      Call ShellandWait(Chr(34) & GDALBinDir & "\gdal\apps\gdalwarp.exe" & Chr(34) & " -s_srs " & Chr(34) & SourceProj & Chr(34) & " -t_srs " & Chr(34) & TargetProj & Chr(34) & " " & Chr(34) & InFile & Chr(34) & " " & Chr(34) & Outfile & Chr(34))
      If predictHour = 0 Then
        Call ShellandWait(Chr(34) & GDALBinDir & "\gdal\apps\gdal_translate.exe" & Chr(34) & " -b " & GegevensBandCurrentFiles & " -of AAIGrid " & Chr(34) & Outfile & Chr(34) & " " & Chr(34) & outFile2 & Chr(34))
      Else
        Call ShellandWait(Chr(34) & GDALBinDir & "\gdal\apps\gdal_translate.exe" & Chr(34) & " -b " & GegevensBandPredictionFiles & " -of AAIGrid " & Chr(34) & Outfile & Chr(34) & " " & Chr(34) & outFile2 & Chr(34))
      End If
    
    End If
    
  Next

End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Sub READASCIIGRID(path As String, ByRef nCols As Long, ByRef nRows As Long, ByRef xllcorner As Double, ByRef yllcorner As Double, ByRef cellsize As Double, ByRef nodata_value As Double, ByRef Data() As Double)
  
  Dim fn As Long, myStr As String, tmpStr As String
  Dim r As Long, c As Long
  Dim spcpos As Long
  fn = FreeFile
  
  If FileExists(path) Then
    Open path For Input As #fn
    While Not EOF(fn)
      Line Input #fn, myStr
      myStr = VBA.Trim(myStr)
      If InStr(1, myStr, "/*") > 0 Then
        'commentaarregel
      ElseIf InStr(1, VBA.LCase(myStr), "ncols") > 0 Then
        tmpStr = ParseString(myStr, " ")
        nCols = VBA.Val(myStr)
      ElseIf InStr(1, VBA.LCase(myStr), "nrows") > 0 Then
        tmpStr = ParseString(myStr, " ")
        nRows = VBA.Val(myStr)
      ElseIf InStr(1, VBA.LCase(myStr), "xllcorner") > 0 Then
        tmpStr = ParseString(myStr, " ")
        xllcorner = VBA.Val(myStr)
      ElseIf InStr(1, VBA.LCase(myStr), "yllcorner") > 0 Then
        tmpStr = ParseString(myStr, " ")
        yllcorner = VBA.Val(myStr)
      ElseIf InStr(1, VBA.LCase(myStr), "cellsize") > 0 Then
        tmpStr = ParseString(myStr, " ")
        cellsize = VBA.Val(myStr)
      ElseIf InStr(1, VBA.LCase(myStr), "nodata_value") > 0 Then
        tmpStr = ParseString(myStr, " ")
        nodata_value = VBA.Val(myStr)
        ReDim Data(1 To nRows, 1 To nCols)
      Else
        r = r + 1
        c = 0
        While Not myStr = ""
          c = c + 1
          Data(r, c) = VBA.Val(ParseString(myStr, " "))
        Wend
      End If
    Wend
    Close (fn)
Else
  MsgBox ("Error: opgegeven bestand bestaat niet.")
End If

End Sub


Public Sub WriteASCIIGridFromEquation(FileName As String, A As Double, b As Double, c As Double, Xmin As Double, Xmax As Double, ymin As Double, ymax As Double, cellsize As Integer)
    'applies the formula for a 2D plane in order to build the grid: z = ax + by + c
    Dim Row As Long, Col As Long, x As Double, y As Double, z As Double
    Dim maxr As Long, maxc As Long, myLine As String
    maxr = (ymax - ymin) / cellsize
    maxc = (Xmax - Xmin) / cellsize
    Dim fn As Long
    fn = FreeFile
    Open FileName For Output As #fn
    
    Print #fn, "ncols        " & maxc
    Print #fn, "nrows        " & maxr
    Print #fn, "xllcorner    " & Xmin
    Print #fn, "yllcorner    " & ymin
    Print #fn, "cellsize     " & cellsize
    Print #fn, "NODATA_value " & -999
    
    For Row = maxr To 1 Step -1
        myLine = ""
        y = ymin + Row - 1
        For Col = 1 To maxc
            x = Xmin + Col * cellsize - cellsize / 2   'we gaan uit van cell centered
            myLine = myLine & " " & Format(A * x + b * y + c, "0.0000")
        Next
        Print #fn, myLine
    Next
    Close (fn)

End Sub

Public Sub WriteASCIIGridFromMultipleEquations(FileName As String, aVals As Range, bVals As Range, cVals As Range, xMinVals As Range, xMaxVals As Range, yMinVals As Range, yMaxVals As Range, xMinGrid As Double, xMaxGrid As Double, yMinGrid As Double, yMaxGrid As Double, cellsize As Double, nodata_value As Double, useLowest As Boolean, NullOutEdges As Boolean)
    'applies the formula for a 2D plane in order to build multiple grids: z = ax + by + c
    'for the entire domain it takes the highest z-value from all supplied plains
    'NOTE: the a-values, b-values and c-values are sought within COLUMNS.
    Dim Row As Long, Col As Long, x As Double, y As Double, z() As Double, zUse As Double, i As Integer
    ReDim z(1 To aVals.Count)
    Dim maxr As Long, maxc As Long, myLine As String
    maxr = (yMaxGrid - yMinGrid) / cellsize
    maxc = (xMaxGrid - xMinGrid) / cellsize
    Dim fn As Long
    
    If aVals.Rows.Count > 1 Or bVals.Rows.Count > 1 Or cVals.Rows.Count > 1 Then
        MsgBox ("Error: het bereik met a-, b-, en c-waarden mag alleen uit kolommen bestaan.")
        End
    ElseIf aVals.Columns.Count <> bVals.Columns.Count Or aVals.Columns.Count <> cVals.Columns.Count Then
        MsgBox ("Error: het bereik voor a-, b-, en c-waarden moet gelijk zijn.")
        End
    End If
    
    fn = FreeFile
    Open FileName For Output As #fn
    
    Print #fn, "ncols        " & maxc
    Print #fn, "nrows        " & maxr
    Print #fn, "xllcorner    " & xMinGrid
    Print #fn, "yllcorner    " & yMinGrid
    Print #fn, "cellsize     " & cellsize
    Print #fn, "NODATA_value " & nodata_value
    
    For Row = maxr To 1 Step -1
        myLine = ""
        y = yMinGrid + (Row - 0.5) * cellsize 'we gaan uit van cell centered
        For Col = 1 To maxc
            x = xMinGrid + (Col - 0.5) * cellsize   'we gaan uit van cell centered
            If useLowest Then zUse = 999 Else zUse = -999 'initialize the z-value that will be used.
            For i = 1 To aVals.Count
                If x >= xMinVals(, i) And x <= xMaxVals(, i) And y >= yMinVals(, i) And y <= yMaxVals(, i) Then
                    z(i) = aVals(, i) * x + bVals(, i) * y + cVals(, i)
                    If useLowest = True And z(i) < zUse Then zUse = z(i)
                    If useLowest = False And z(i) > zUse Then zUse = z(i)
                End If
                If NullOutEdges Then
                    If Row = 1 Then
                        zUse = nodata_value
                    ElseIf Row = maxr Then
                        zUse = nodata_value
                    ElseIf Col = 1 Then
                        zUse = nodata_value
                    ElseIf Col = maxc Then
                        zUse = nodata_value
                    End If
                End If
                
            Next
            myLine = myLine & " " & Format(zUse, "0.0000")
        Next
        Print #fn, myLine
    Next
    Close (fn)

End Sub

Public Sub GRIDINTEGERS(path As String, ByRef nCols As Long, ByRef nRows As Long, ByRef xllcorner As Double, ByRef yllcorner As Double, ByRef cellsize As Double, ByRef nodata_value As Double, ByRef Data() As Integer)
  
  Dim fn As Long, myStr As String
  Dim i As Long, j As Long
  fn = FreeFile
  
  Open path For Output As #fn
  Print #fn, "ncols         " & nCols
  Print #fn, "nrows         " & nRows
  Print #fn, "xllcorner     " & xllcorner
  Print #fn, "yllcorner     " & yllcorner
  Print #fn, "cellsize      " & cellsize
  Print #fn, "NODATA_value  " & nodata_value
        
  For i = 1 To nRows
    myStr = ""
    For j = 1 To nCols - 1
      myStr = myStr & Data(i, j) & " "
    Next
    Print #fn, myStr & Data(i, j)
  Next
  Close (fn)

End Sub
Public Sub ASCII2XYZ(ASCPath As String, XYZPath As String)
  Dim fn As Long, r As Long, c As Long, x As Double, y As Double, z As Double
  Dim nCols As Long, nRows As Long, xllcorner As Double, yllcorner As Double, cellsize As Double, nodatavalue As Double, Data() As Double
  'converteert een ASCII grid in een bestand met X Y Z
     
  Call READASCIIGRID(ASCPath, nCols, nRows, xllcorner, yllcorner, cellsize, nodatavalue, Data)
  
  fn = FreeFile
  Open XYZPath For Output As #fn
  
  For r = 1 To nRows
    y = yllcorner + (nRows - r + 0.5) * cellsize
    For c = 1 To nCols
      x = xllcorner + cellsize * (c - 0.5)
      z = Data(r, c)
      If Not z = nodatavalue Then Print #fn, x & " " & y & " " & z
    Next
  Next
  Close (fn)
  
End Sub

Public Sub READMT940(path As String, StartRow As Integer, StartCol As Integer)
  'MT940 is een bestandsformaat voor rekeningafschriften, o.a. gebruikt door ABN AMRO
  Dim fn As Long, i As Long, r As Long, c As Long, myStr As String, tmpStr As String, CD As String 'credit debet
  Dim mult As Integer
  fn = FreeFile
  r = StartRow - 1
  c = StartCol
  
  ActiveSheet.Range(Cells(StartRow, StartCol), Cells(StartRow + 1000000, StartCol + 10)).ClearContents
      
  Open path For Input As #fn
  
  Close #fn
    
  If FileExists(path) Then
    Open path For Input As #fn
    While Not EOF(fn)
      Line Input #fn, myStr
      myStr = replace(VBA.Trim(myStr), ",", ".")
      
'      If myStr = "940" Then 'er begint een nieuw afschrift
'        r = r + 1
'      ElseIf Left(myStr, 4) = ":60F" Then 'beginsaldo afschrift
'        ActiveSheet.Cells(r, c) = ParseString(myStr, "EUR")
'      ElseIf Left(myStr, 4) = ":62F " Then 'eindsaldo afschrift"
'        ActiveSheet.Cells(r, c) = ParseString(myStr, "EUR")
'      ElseIf Left(myStr, 3) = ":20" Then 'banknaam
'        ActiveSheet.Cells(r, c + 1) = MultiParse(myStr, 2, ":")
'      ElseIf Left(myStr, 3) = ":25" Then 'rekeningnummer
'        ActiveSheet.Cells(r, c + 2) = MultiParse(myStr, 2, ":")
'      ElseIf Left(myStr, 3) = ":28" Then 'afschriftnummer
'        ActiveSheet.Cells(r, c + 3) = MultiParse(myStr, 2, ":")
      If Left(myStr, 3) = ":61" Then 'bedrag
        r = r + 1
        Call ParseString(myStr, ":")
        tmpStr = ParseNumeric(myStr)                  'datumtijdstring: jjmmdduumm
        ActiveSheet.Cells(r, c + 2) = "20" & VBA.Left(tmpStr, 2) & "-" & VBA.Left(VBA.Right(tmpStr, 8), 2) & "-" & VBA.Left(VBA.Right(tmpStr, 6), 2)
        ActiveSheet.Cells(r, c + 3) = VBA.Left(VBA.Right(tmpStr, 4), 2) & ":" & VBA.Right(tmpStr, 2)
        CD = VBA.Left(myStr, 1)                        'D=debet, C=credit
        myStr = Right(myStr, VBA.Len(myStr) - 1)     'restant = bedrag + een of andere code
        If CD = "D" Then
          mult = -1
        ElseIf CD = "C" Then
          mult = 1
        End If
        ActiveSheet.Cells(r, c + 4) = ParseNumeric(myStr) * mult
      ElseIf Left(myStr, 3) = ":86" Then
        Call ParseString(myStr, ":")
        
        'eerst de bankautomaat of het rekeningnummer identificeren
        tmpStr = ParseString(myStr, " ")
        If tmpStr = "BEA" Or tmpStr = "GEA" Then
          tmpStr = tmpStr & " " & ParseString(myStr, " ")
          ActiveSheet.Cells(r, c + 5) = VBA.Trim(tmpStr)
        ElseIf tmpStr = "GIRO" Then
          tmpStr = tmpStr & " " & ParseString(myStr, " ")
          ActiveSheet.Cells(r, c + 5) = VBA.Trim(tmpStr)
        ElseIf IsBankNumber(tmpStr) Then
          ActiveSheet.Cells(r, c + 5) = VBA.Trim(tmpStr)
        End If
        
        tmpStr = ParseString(myStr, " ")
        If VBA.InStr(tmpStr, "/") > 0 Then
          ActiveSheet.Cells(r, c + 6) = VBA.Trim(tmpStr)
        Else
          myStr = tmpStr & " " & myStr
        End If
        ActiveSheet.Cells(r, c + 7) = myStr
      End If
    Wend
  Close (fn)
Else
  MsgBox ("Error: opgegeven bestand bestaat niet.")
End If
  
End Sub

Public Function MATCHWILDCARD(myStr As String, myMask As String, CaseSensitive As Boolean) As Boolean
  'Date: 8-12-2013
  'Author: Siebe Bosch
  'Description: matches a given string with a string with wildcards
  'Note: only tested for SOMETHING* so far.
  Dim tmpMask As String, tmpStr As String, checkStr As String, i As Integer, startPos As Integer
  Dim maskPart As String, partPos As Integer
  
  'if case insensitive, convert both strings to uppercase
  If CaseSensitive = False Then
    myStr = VBA.UCase(myStr)
    myMask = VBA.UCase(myMask)
  End If
  
  'create a new string that consists of asteriskses only and that has the length of myStr
  For i = 1 To VBA.Len(myStr)
    checkStr = checkStr & "*"
  Next
  
  'now start parsing the mask in order to find its components (disregarding the wildcards for now)
  startPos = 1
  tmpMask = myMask
  While Not tmpMask = ""
    maskPart = ParseString(tmpMask, "*")
    partPos = InStr(startPos, myStr, maskPart, vbBinaryCompare)
    If partPos > 0 Then
      'embed the string we found in checkStr, at the exact same location
      checkStr = Left(checkStr, partPos - 1) & maskPart & VBA.Right(checkStr, VBA.Len(checkStr) - (partPos - 1) - VBA.Len(maskPart))
    End If
  Wend
  
  'now that we have a checkStr that only consists of * and parts from the mask, we can reduce it to its minimum
  'and check whether it matches our original mask
  While InStr(1, checkStr, "**") > 0
    checkStr = VBA.replace(checkStr, "**", "*")
  Wend
  
  If checkStr = myMask Then
    MATCHWILDCARD = True
  Else
    MATCHWILDCARD = False
  End If

End Function

Public Sub CONCATENATECOMBINATIONS(myRange As Range, StartRow As Integer, resultsCol As Integer)
    'this function combines all unique values from all columns in the given range (except for empty cells)
    'and concatenates them
    Dim r As Long, n As Long
    Dim Found As Boolean
    r = StartRow
    
    Dim i As Long, j As Long
    Dim myStr As String
    Dim newArray() As Integer
    ReDim newArray(1 To myRange.Columns.Count)
    While MileageOneUp(1, myRange.Rows.Count, newArray)
    
        If newArray(1) = 21 And newArray(2) = 21 And newArray(3) = 20 Then Stop
        
        n = n + 1
        Found = True
        myStr = ""
        For j = 1 To myRange.Columns.Count
            If myRange.Cells(newArray(j), j) = "" Then Found = False
            myStr = myStr & myRange.Cells(newArray(j), j)
        Next
        If Found Then
          r = r + 1
          ActiveSheet.Cells(r, resultsCol) = myStr
        End If
        If n > myRange.Rows.Count ^ myRange.Columns.Count Then End
    Wend

End Sub

Public Function ReadEntireTextFile(myPath) As String
  Dim fn As Long, myStr As String
  Dim fileContent As String

  'reads the entire file to memory
  Open myPath For Input As #fn
  
  If FileExists(myPath) Then
    Open myPath For Input As #fn
    fileContent = VBA.Input(LOF(ifn), ifn)
    Close #fn
  Else
    MsgBox ("Error: file does not exist: " & myPath)
    End
  End If
    
  'return the result
  ReadEntireTextFile = fileContent

End Function

Public Sub JoinNodes(myRange As Range, IDCol As Long, XCol As Long, YCol As Long, rIDcol As Long, rXcol As Long, rYcol As Long, Mergedistance As Double, Optional ResultsNodePrefix As String = "", Optional BNACol As Long = 0)
  'maakt nieuwe knopen aan door knopen die dicht bijeen liggen samen te voegen. Handig als lozingspunten van meerdere afwateringseenheden dicht bijeen liggen.
  Dim JoinedNodes As New Collection
  Dim JoinedNode As clsMultiNodeObject, Node As clsNode
  Dim i As Long, j As Long, K As Long, n As Long, myDist As Double
  Dim Found As Boolean
    
  For i = 1 To myRange.Rows.Count
    If i = 1 Then
      n = 1
      Set JoinedNode = New clsMultiNodeObject
      JoinedNode.ID = ResultsNodePrefix & n
      Set Node = New clsNode
      Node.ID = myRange.Cells(i, IDCol).Value
      Node.x = myRange.Cells(i, XCol)
      Node.y = myRange.Cells(i, YCol)
      Call JoinedNode.AddNode(Node)
      Call JoinedNodes.Add(JoinedNode)
    Else
      Found = False
      Set Node = New clsNode
      Node.ID = myRange.Cells(i, IDCol)
      Node.x = myRange.Cells(i, XCol)
      Node.y = myRange.Cells(i, YCol)
      
      For j = 1 To JoinedNodes.Count
        Set JoinedNode = JoinedNodes(j)
        myDist = PointDistance(JoinedNode.XAvg, JoinedNode.YAvg, Node.x, Node.y)
        If myDist <= Mergedistance Then
          Found = True
          Call JoinedNode.AddNode(Node)
        End If
      Next
      If Not Found Then
        n = n + 1
        Set JoinedNode = New clsMultiNodeObject
        JoinedNode.ID = ResultsNodePrefix & n
        Call JoinedNode.AddNode(Node)
        Call JoinedNodes.Add(JoinedNode)
      End If
    End If
  Next

  'schrijf de resultaten weg
  For i = 1 To myRange.Rows.Count
    For j = 1 To JoinedNodes.Count
      Set JoinedNode = JoinedNodes(j)
      For K = 1 To JoinedNode.Nodes.Count
        Set Node = JoinedNode.Nodes(K)
        If myRange.Cells(i, IDCol) = Node.ID Then
          myRange.Cells(i, rIDcol) = JoinedNode.ID
          myRange.Cells(i, rXcol) = JoinedNode.XAvg
          myRange.Cells(i, rYcol) = JoinedNode.YAvg
          Exit For
        End If
      Next
    Next
  Next
  
  'optie BNA-string wegschrijven
  If BNACol > 0 Then
    For j = 1 To JoinedNodes.Count
      Set JoinedNode = JoinedNodes(j)
      myRange.Cells(j, BNACol) = BNAString(JoinedNode.ID, JoinedNode.XAvg, JoinedNode.YAvg)
    Next
  End If
End Sub


'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'-----------------------------------------STRINGBEWERKINGEN--------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------

Public Function VERWIJDERDAGNAAMUITDATUM(myString As String) As String
  myString = VBA.LCase(myString)
  myString = VBA.replace(myString, "maandag", "")
  myString = VBA.replace(myString, "dinsdag", "")
  myString = VBA.replace(myString, "woensdag", "")
  myString = VBA.replace(myString, "donderdag", "")
  myString = VBA.replace(myString, "vrijdag", "")
  myString = VBA.replace(myString, "zaterdag", "")
  myString = VBA.replace(myString, "zondag", "")
  myString = VBA.Trim(myString)
  VERWIJDERDAGNAAMUITDATUM = myString

End Function

Public Function MAKEXMLTOKEN(myToken As String, myValue As String) As String
  MAKEXMLTOKEN = myToken & "=" & VBA.Str(34) & myValue & VBA.Str(34)
End Function

Public Function getDoubleFromXMLRecord(xmlStr As String, TokenID As String) As Double
  Dim Result As String
  Result = VBA.LCase(xmlStr)
  Result = VBA.replace(Result, "<" & VBA.LCase(TokenID) & ">", "")
  Result = VBA.replace(Result, "</" & VBA.LCase(TokenID) & ">", "")
  Result = VBA.Trim(Result)
  getDoubleFromXMLRecord = Result
End Function


Public Function STRINGPOSITIE(SearchString As String, SeekString As String, Optional startPos As Long = 1) As Long
  Dim myPos As Long
  myPos = InStr(startPos, SearchString, SeekString)
  STRINGPOSITIE = myPos
End Function

Public Function ReplaceString(SearchStr As String, FindStr As String, ReplaceStr As String) As String
  ReplaceString = VBA.replace(SearchStr, FindStr, ReplaceStr, , , vbTextCompare)
End Function

Public Sub REPLACESTRINGINALLFILES(SearchDir As String, FindStr As String, ReplaceStr As String)
  Dim myCollection As Collection, myFile As String, myContent As String, Found As Boolean
  Dim fn As Long, of As Long, i As Long
  
  Set myCollection = New Collection
  Set myCollection = ListFilesInFolder(SearchDir)
  For i = 1 To myCollection.Count
    myFile = SearchDir & "\" & myCollection.Item(i)
    myFile = ReplaceString(myFile, "\\", "\")       'make sure we only have one backslash at a time in the path
    Found = False
    fn = FreeFile
    Open myFile For Input As #fn
      If LOF(fn) > 0 Then
        myContent = Input(LOF(fn), fn)
        If InStr(1, myContent, FindStr, vbTextCompare) > 0 Then
          myContent = ReplaceString(myContent, FindStr, ReplaceStr)
          Found = True
        End If
      End If
    Close
    
    If Found Then
      of = FreeFile
      Open myFile For Output As #of
      Print #of, myContent
      Close #of
    End If
  Next
  
End Sub

Public Function DOUBLEIDSINSTRINGCOLLECTION(myCollection As Collection, ByRef doubleStr As String) As Boolean
  'checkt of een collectie van strings dubbele waarden bevat
  Dim i As Long, j As Long
  
  DOUBLEIDSINSTRINGCOLLECTION = False
  For i = 1 To myCollection.Count
    For j = i + 1 To myCollection.Count
      If myCollection(i) = myCollection(j) Then
        doubleStr = myCollection(i)
        DOUBLEIDSINSTRINGCOLLECTION = True
        Exit Function
      End If
    Next
  Next

End Function

Public Function TRIMUSINGCUSTOMSTRING(myStr As String, myTrimStr As String, Optional CaseSensitive As Boolean = False) As String

If Not CaseSensitive Then
  While VBA.Left(VBA.LCase(myStr), VBA.Len(myTrimStr)) = VBA.LCase(myTrimStr)
    myStr = VBA.Right(myStr, VBA.Len(myStr) - VBA.Len(myTrimStr))
  Wend
  While VBA.Right(VBA.LCase(myStr), VBA.Len(myTrimStr)) = VBA.LCase(myTrimStr)
    myStr = VBA.Left(myStr, VBA.Len(myStr) - VBA.Len(myTrimStr))
  Wend
Else
  While VBA.Left(myStr, VBA.Len(myTrimStr)) = myTrimStr
    myStr = VBA.Right(myStr, VBA.Len(myStr) - VBA.Len(myTrimStr))
  Wend
  While VBA.Right(myStr, VBA.Len(myTrimStr)) = myTrimStr
    myStr = VBA.Left(myStr, VBA.Len(myStr) - VBA.Len(myTrimStr))
  Wend
End If

TRIMUSINGCUSTOMSTRING = myStr

End Function

Public Function UnifyString(myStr As String) As String
  'deze functie uniformeert eeen string door de uppercase te nemen en hem te VBA.Trimmen.
  'handig om te gebruiken als key in collections
  UnifyString = VBA.UCase(VBA.Trim(myStr))
End Function

Public Function IsBankNumber(myStr As String) As Boolean
  myStr = VBA.Trim(myStr)
  If Mid(myStr, 3, 1) = "." And VBA.Mid(myStr, 6, 1) = "." And VBA.Mid(myStr, 9, 1) = "." Then
    IsBankNumber = True
  Else
    IsBankNumber = False
  End If
End Function

Public Function FindNearestObjectInRange(x As Double, y As Double, SearchListRange As Range, IDColIdx As Long, XColIdx As Long, YColIdx As Long) As String

Dim Dist As Double, tmpDist As Double, ID As String, tmpID As String, r As Long, c As Long

'initialiseren
Dist = Sqr((x - SearchListRange.Cells(1, XColIdx)) ^ 2 + (y - SearchListRange.Cells(1, YColIdx)) ^ 2)
ID = SearchListRange.Cells(1, IDColIdx)

For r = 2 To SearchListRange.Rows.Count
  tmpDist = Sqr((x - SearchListRange.Cells(r, XColIdx)) ^ 2 + (y - SearchListRange.Cells(r, YColIdx)) ^ 2)
  tmpID = SearchListRange.Cells(r, IDColIdx)
  If tmpDist < Dist Then
    Dist = tmpDist
    ID = tmpID
  End If
Next

FindNearestObjectInRange = ID

End Function

'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'-----------------------------------------BESTANDEN----------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------

Public Function OpenSingleFile() As String
  Dim Filter As String, Title As String
  Dim FilterIndex As Integer
  Dim FileName As Variant
  
  ' File filters
  Filter = "MT940 Files (*.sta),*.sta, All Files (*.*),*.*"
  FilterIndex = 3

  ' Set Dialog Caption
  Title = "Selecteer een bestand."
  'ChDrive ("C")
  'ChDir ("E:\Chapters\chap14")
  With Application
    ' Set File Name to selected File
    FileName = .GetOpenFilename(Filter, FilterIndex, Title)
    ' Reset Start Drive/Path
    ChDrive (VBA.Left(.DefaultFilePath, 1))
    ChDir (.DefaultFilePath)
  End With

 ' Exit on Cancel
 If FileName = False Then
    MsgBox "No file was selected."
    Exit Function
  End If
  OpenSingleFile = FileName
End Function

Public Function ListFilesInFolder(SourceFolderName As String, Optional EXT As String = "*") As Collection
  'The macro example below assumes that your VBA project has added a reference to the Microsoft Scripting Runtime library.
  'You can do this from within the VBE by selecting the menu Etra, References and selecting Microsoft Scripting Runtime.
  
  ' lists information about the files in SourceFolder
  ' example: ListFilesInFolder "C:\FolderName\", True
  Dim myFile As String
  Dim myCollection As Collection
  Set myCollection = New Collection
  
  myFile = dir$(SourceFolderName & "\*." & EXT)
  Do While myFile <> ""
    myCollection.Add myFile
    myFile = dir$
  Loop
  Set ListFilesInFolder = myCollection

End Function

Public Function DirectoryExists(DName As String) As Boolean

Dim sDummy As String
On Error Resume Next

If VBA.Right(DName, 1) <> "\" Then DName = DName & "\"
sDummy = dir$(DName & "*.*", vbDirectory)
DirectoryExists = Not (sDummy = "")

End Function

Public Function CONTAINSKEY(ByRef Col As Collection, ByVal key As Variant) As Boolean

Dim obj As Variant
On Error GoTo err
  CONTAINSKEY = True
  obj = Col(key)
  Exit Function
err:
  CONTAINSKEY = False

End Function

Public Function CONTAINSKEY_BYOBJECTID(ByRef Col As Collection, ByVal ID As String) As Boolean

'uses the .ID element of the objects in a collection as a key
'this is because VBA has no way of retrieving objects from a collection by Key
'note: this only works if the elements of the collection actually HAVE an element named ID

Dim i As Long
For i = 1 To Col.Count
  If VBA.Trim(VBA.UCase(Col.Item(i).ID)) = VBA.Trim(VBA.UCase(ID)) Then
    CONTAINSKEY_BYOBJECTID = True
    Exit Function
  End If
Next

'not found
CONTAINSKEY_BYOBJECTID = False


End Function

Public Sub DELETESHAPEFILE(path As String)
  Dim myPath As String
  myPath = path
  If FileExists(myPath) Then Call DeleteFile(myPath)
  myPath = replace(path, ".shp", ".dbf")
  If FileExists(myPath) Then Call DeleteFile(myPath)
  myPath = replace(path, ".shp", ".shx")
  If FileExists(myPath) Then Call DeleteFile(myPath)
  myPath = replace(path, ".shp", ".prj")
  If FileExists(myPath) Then Call DeleteFile(myPath)
End Sub

Public Sub MoveFile(FromDir As String, ToDir As String, FileName As String)
  Dim FromFile As String, ToFile As String
  FromFile = FromDir & "\" & FileName
  ToFile = ToDir & "\" & FileName

  If FileExists(FromFile) Then
    If DirectoryExists(ToDir) Then
      Call FileCopy(FromFile, ToFile)
      Call Kill(FromFile)
    Else
      MsgBox ("Error: target directory does not exist:" & ToDir)
    End If
  Else
      MsgBox ("Error: file does not exist:" & FromFile)
  End If

End Sub

Public Sub DIRECTORYCOPY(FromDir As String, ToDir As String)
  'This example copy all files and subfolders from FromPath to ToPath.
  'Note: If ToPath already exist it will overwrite existing files in this folder
  'if ToPath not exist it will be made for you.
    Dim FSO As Object
    Set FSO = CreateObject("scripting.filesystemobject")

    If FSO.FolderExists(FromDir) = False Then
        MsgBox FromDir & " doesn't exist"
        Exit Sub
    End If
    FSO.CopyFolder Source:=FromDir, Destination:=ToDir
End Sub

Public Function FOLDERBROWSER(strPath As String) As String
  Dim fldr As FileDialog
  Dim sItem As String
  Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
  With fldr
    .Title = "Select a Folder"
    .AllowMultiSelect = False
    .InitialFileName = strPath
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
  End With
NextCode:
  FOLDERBROWSER = sItem
  Set fldr = Nothing
End Function

Public Sub ReplaceInFile(InFile As String, Outfile As String, ReplaceString As String, ReplaceByString As String)
  
  Dim fn As Long, fn2 As Long, myStr As String
  fn = FreeFile
  Open InFile For Input As #fn
  fn2 = FreeFile
  Open Outfile For Output As #fn2
  
  While Not EOF(fn)
    Line Input #fn, myStr
    myStr = replace(ReplaceString, myStr, ReplaceByString)
    Print #fn2, myStr
  Wend
  
  Close (fn)
  Close (fn2)

End Sub

Public Function ReplaceInString(SourceStr As String, ReplaceStr As String, ReplaceBy As String) As String
    ReplaceInString = replace(SourceStr, ReplaceStr, ReplaceBy)
End Function

Function FileNameFromPath(path) As String
    FileNameFromPath = Right(path, InStrRev(path, "\" - 1))
End Function
Function DirFromPath(path) As String
   DirFromPath = Left(path, InStrRev(path, "\"))
End Function

Function GetDirectory(path) As String
   GetDirectory = Left(path, InStrRev(path, "\"))
End Function

Function WorkSheetExists(wksName As String) As Boolean
  'checks of een worksheet al bestaat
  On Error Resume Next
  WorkSheetExists = CBool(Len(Worksheets(wksName).Name) > 0)
End Function

Public Function SumRange(myRange As Range) As Double
    Dim CurCell As Object
    Dim mySum As Double
    For Each CurCell In myRange
      mySum = mySum + CurCell.Value
    Next
    SumRange = mySum
    Exit Function
End Function

Public Function FRACTIONOFDAYSUM(myDateTimeCell As Range, DateTimeCol As Long, valuesCol As Long) As Double
  'Deze functie rekent uit welk aandeel van de dagsom in een bepaalde cel staat
  'Dit betekent dat je moet opgeven: de kolom waarin datum/tijd staat, de kolom waarin de bijbehorende waarden staan
  'én natuurlijk de cel met de datum/tijd waarvoor je de fractie wilt weten en de cel waarin de waarde staat.
  'de functie deelt de waarde uit de gezochte cel door de som van de waarden van alle cellen die op dezelfde datum vallen
  
  Dim myDay As Double
  Dim myYear As Double
  Dim mySum As Double
  Dim myCell As Object
  Dim myValue As Double
  Dim nCells As Long
  Dim r As Long
  Dim Done As Boolean
  
  myDay = Day(myDateTimeCell.Value)
  myYear = Year(myDateTimeCell.Value)
  myValue = ActiveSheet.Cells(myDateTimeCell.Row, valuesCol).Value
  mySum = myValue
  nCells = 1
  
  If myDateTimeCell.Count <> 1 Then
    MsgBox ("Error: één cel selecteren voor huidige datum/tijd")
  End If
  
  'we lopen vanaf de gevraagde cel omhoog tot de datum verschilt
  r = myDateTimeCell.Row
  Done = False
  While Not Done
    r = r - 1
    If r > 0 And IsDate(ActiveSheet.Cells(r, DateTimeCol)) Then
      If Day(ActiveSheet.Cells(r, DateTimeCol)) = myDay And Year(ActiveSheet.Cells(r, DateTimeCol)) = myYear Then
        nCells = nCells + 1
        mySum = mySum + ActiveSheet.Cells(r, valuesCol)
      Else
        Done = True
      End If
    Else
      Done = True
    End If
  Wend
  
  'en nu omlaag
  r = myDateTimeCell.Row
  Done = False
  While Not Done
    r = r + 1
    If IsDate(ActiveSheet.Cells(r, DateTimeCol)) Then
      If Day(ActiveSheet.Cells(r, DateTimeCol)) = myDay And Year(ActiveSheet.Cells(r, DateTimeCol)) = myYear Then
        nCells = nCells + 1
        mySum = mySum + ActiveSheet.Cells(r, valuesCol)
      Else
        Done = True
      End If
    Else
      Done = True
    End If
  Wend
  
  If mySum = 0 Then
    FRACTIONOFDAYSUM = 1 / nCells
  Else
    FRACTIONOFDAYSUM = myValue / mySum
  End If

End Function

Public Function IsRangeAscending(myRange As Range) As Boolean
'checkt of een range (1e kolom) een oplopende volgorde heeft
Dim r As Long
IsRangeAscending = True
  If myRange.Rows.Count > 1 Then
    For r = 2 To myRange.Rows.Count
      If myRange.Cells(r, 1).Value < myRange.Cells(r - 1, 1).Value Then
        IsRangeAscending = False
      End If
    Next
  Else
    IsRangeAscending = True
  End If
End Function


Public Function MinYFromXYRange(myWorksheet As String, myXRange As Range, myYRange As Range, Optional fromX As Double = -10000000000000#, Optional toX As Double = 10000000000000#) As Double
Dim Row As Long, curSheet As String
'retrieves te lowest Y value from a Range with X and Y values
'XcolIdx is the index number of the column within the range in which the X values can be found
'YColIdx is the index number of the column within the range in which the Y values can be found
'fromX and toX are optional and can be used to restrict the search to the part of the range where X falls between these values
curSheet = ActiveWorkbook.ActiveSheet.Name

If myXRange.Rows.Count <> myYRange.Rows.Count Then
  MsgBox ("Error in function MinYFromXYRange. Ranges must be of equal length.")
  Exit Function
ElseIf myXRange.Columns.Count <> 1 Then
  MsgBox ("Error in function MinYFromXYRange. Range containing X values must consist of only one column.")
  Exit Function
ElseIf myYRange.Columns.Count <> 1 Then
  MsgBox ("Error in function MinYFromXYRange. Range containing Y values must consist of only one column.")
  Exit Function
End If

MinYFromXYRange = 10000000000000#
Worksheets(myWorksheet).Activate
For Row = 1 To myXRange.Rows.Count
  If IsNumeric(myXRange.Cells(Row, 1)) And IsNumeric(myYRange.Cells(Row, 1)) Then
    If myYRange.Cells(Row, 1) < MinYFromXYRange And myXRange.Cells(Row, 1) >= fromX And myXRange.Cells(Row, 1) <= toX Then
      MinYFromXYRange = myYRange.Cells(Row, 1)
    End If
  Else
    'MsgBox ("Error in function MinYFromXYRange: non numeric value encountered in row index " & row & " of the data range.")
    'Exit Function
  End If
Next Row

End Function

Public Function MaxYFromXYRange(myWorksheet As String, myXRange As Range, myYRange As Range, Optional fromX As Double = -10000000000000#, Optional toX As Double = 10000000000000#) As Double
Dim Row As Long, curSheet As String
'retrieves te highest Y value from a Range with X and Y values
'XcolIdx is the index number of the column within the range in which the X values can be found
'YColIdx is the index number of the column within the range in which the Y values can be found
'fromX and toX are optional and can be used to restrict the search to the part of the range where X falls between these values
curSheet = ActiveWorkbook.ActiveSheet.Name

If myXRange.Rows.Count <> myYRange.Rows.Count Then
  MsgBox ("Error in function MaxYFromXYRange. Ranges must be of equal length.")
  Exit Function
ElseIf myXRange.Columns.Count <> 1 Then
  MsgBox ("Error in function MaxYFromXYRange. Range containing X values must consist of only one column.")
  Exit Function
ElseIf myYRange.Columns.Count <> 1 Then
  MsgBox ("Error in function MaxYFromXYRange. Range containing Y values must consist of only one column.")
  Exit Function
End If

MaxYFromXYRange = -10000000000000#
Worksheets(myWorksheet).Activate
For Row = 1 To myXRange.Rows.Count
  If IsNumeric(myXRange.Cells(Row, 1)) And IsNumeric(myYRange.Cells(Row, 1)) Then
    If myYRange.Cells(Row, 1) > MaxYFromXYRange And myXRange.Cells(Row, 1) >= fromX And myXRange.Cells(Row, 1) <= toX Then
      MaxYFromXYRange = myYRange.Cells(Row, 1)
    End If
  Else
    'MsgBox ("Error in function MaxYFromXYRange: non numeric value encountered in row index " & row & " of the data range.")
    'Exit Function
  End If
Next Row

End Function

Public Function CONCATENATEALGEBRAIC(myRange As Range, AlgebraString As String) As String
  Dim i As Long, Result As String
  If myRange.Columns.Count <> 1 Then
    MsgBox ("Error in function CONCATENATEALGEBRAIC. Range must consist of one column.")
  Else
   Result = myRange.Rows(1)
   For i = 2 To myRange.Rows.Count
     Result = Result & " " & AlgebraString & " " & myRange.Rows(i)
   Next
   CONCATENATEALGEBRAIC = Result
  End If
End Function

Public Function CONCATENATEWITHDELIMITER(myRange As Range, Delimiter As String) As String
  Dim Result As String
  Dim r As Long, c As Long
  
  If Delimiter = "\t" Then Delimiter = vbTab
  
  For r = 1 To myRange.Rows.Count
    For c = 1 To myRange.Columns.Count
      If r = 1 And c = 1 Then
        Result = myRange.Cells(r, c)
      Else
        Result = Result & Delimiter & myRange.Cells(r, c)
      End If
    Next
  Next
  CONCATENATEWITHDELIMITER = Result

End Function

Public Sub AddWorkSheet(SheetName As String)

  If WorkSheetExists(SheetName) Then
    Application.DisplayAlerts = False
    Worksheets(SheetName).Delete
    Application.DisplayAlerts = True
    Worksheets.Add
    ActiveSheet.Name = SheetName
  Else
    Worksheets.Add
    ActiveSheet.Name = SheetName
  End If

End Sub

Public Function FindColumnOnWorkSheet(SheetName As String, Header As String, Row As Long, Optional GiveWarning As Boolean) As Long
Dim Col As Long

FindColumnOnWorkSheet = 0
For Col = 1 To 100
  If VBA.LCase(Worksheets(SheetName).Cells(Row, Col)) = VBA.LCase(Header) Then
    FindColumnOnWorkSheet = Col
    Exit For
  End If
Next Col

If FindColumnOnWorkSheet = 0 And GiveWarning Then
  MsgBox ("Column " & Header & " not found.")
End If

End Function

Public Function UnpivotMultiHeader(ByRef myRange As Range, nHeaderRows As Integer, nHeaderCols As Integer, ResultsStartRow As Integer, ResultsStartCol As Integer) As Boolean
    'this routine converts a given table including multiple headers into a pivot-ready table
    Dim r As Long, c As Long
    Dim Row As Integer, Col As Integer
    Dim RowHeaderIdx As Integer
    Dim ColHeaderIdx As Integer
    
    Row = ResultsStartRow
    Col = ResultsStartCol
    For r = (nHeaderRows + 1) To myRange.Rows.Count
        For c = nHeaderCols + 1 To myRange.Columns.Count
            Col = ResultsStartCol
            For RowHeaderIdx = 1 To nHeaderRows
                ActiveSheet.Cells(Row, Col) = myRange.Cells(RowHeaderIdx, c)
                Col = Col + 1
            Next
            For ColHeaderIdx = 1 To nHeaderCols
                ActiveSheet.Cells(Row, Col) = myRange.Cells(r, ColHeaderIdx)
                Col = Col + 1
            Next
            ActiveSheet.Cells(Row, Col) = myRange.Cells(r, c)
            Row = Row + 1
        Next
    Next

End Function

Public Function UnpivotTable(ByRef myRange As Range, ColumnHeaderVariableName As String, RowHeaderVariableName As String, ValueHeaderName As String) As Boolean
    'a very simple unpivot
    Dim curSheet As Worksheet, newSheet As Worksheet
    Dim CurSheetName As String
    Dim NewSheetName As String
    CurSheetName = ActiveSheet.Name
    NewSheetName = CurSheetName & ".UnPivot"
    Set curSheet = ActiveWorkbook.Sheets(CurSheetName)
    
    'create a new worksheet for the results
    If Not WorkSheetExists(NewSheetName) Then
        Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = NewSheetName
        Set newSheet = ActiveWorkbook.Sheets(NewSheetName)
        UnpivotTable = True
    Else
        MsgBox ("Worksheet " & NewSheetName & " already exists. Please remove the old one first.")
        UnpivotTable = False
    End If
    
    Dim r2 As Integer
    r2 = 1
    newSheet.Cells(r2, 1) = ColumnHeaderVariableName
    newSheet.Cells(r2, 2) = RowHeaderVariableName
    newSheet.Cells(r2, 3) = ValueHeaderName
    

     
    Dim r As Integer, c As Integer
    For r = 2 To myRange.Rows.Count
        For c = 2 To myRange.Columns.Count
            r2 = r2 + 1
            newSheet.Cells(r2, 1) = myRange.Cells(1, c)
            newSheet.Cells(r2, 2) = myRange.Cells(r, 1)
            newSheet.Cells(r2, 3) = myRange.Cells(r, c)
        Next
    Next
    
    
End Function

Public Function UnPivot2(ByRef myRange As Range, ValuesStartCol As Integer, ValuesEndCol As Integer, ResultsStartRow As Integer, ResultsStartCol As Integer)
    Dim r As Integer, c As Integer
    Dim Header1 As String, Header2 As String, Value As Double
    Dim resrow As Integer, rescol As Integer
    
    resrow = ResultsStartRow
    rescol = ResultsStartCol
    ActiveSheet.Cells(resrow, rescol) = "Header1"
    ActiveSheet.Cells(resrow, rescol + 1) = "Header2"
    ActiveSheet.Cells(resrow, rescol + 2) = "Value"
            
    For r = 2 To myRange.Rows.Count
        Header1 = myRange.Cells(r, 1)
        For c = ValuesStartCol To ValuesEndCol
            Header2 = myRange.Cells(1, c)
            Value = myRange.Cells(r, c)
            resrow = resrow + 1
            ActiveSheet.Cells(resrow, rescol) = Header1
            ActiveSheet.Cells(resrow, rescol + 1) = Header2
            ActiveSheet.Cells(resrow, rescol + 2) = Value
        Next
    Next

End Function


Public Function UnPivot(ByRef myRange As Range, HeaderColNum As Integer, UnpivotColNum As Integer, ValueColNum As Integer) As Boolean
  'This routine creates a new worksheet and writes data from the original sheet in an unpivoted way
  'Note: as input a range with exactly three (3) columns is required! One for the row headers in the result, one for the new columns and one containing the values
  Dim r As Long, c As Long, r2 As Long, c2 As Long
  Dim nRow As Integer, nCol As Integer
  nRow = myRange.Rows.Count
  nCol = myRange.Columns.Count
  Dim myArray() As Variant
  Dim CurSheetName As String
  Dim NewSheetName As String
  Dim PivotList As Collection
  Set PivotList = New Collection
  Dim HeaderList As Collection
  Set HeaderList = New Collection
      
  Dim curSheet As Worksheet, newSheet As Worksheet
  CurSheetName = ActiveSheet.Name
  NewSheetName = CurSheetName & ".UnPivot"
  Set curSheet = ActiveWorkbook.Sheets(CurSheetName)
    
  'since we've been given a column that needs unpivoting, we'll first create a collection of all unique elements inside that column
  For r = 2 To myRange.Rows.Count
    If CollectionGetIndex(PivotList, myRange.Cells(r, UnpivotColNum)) = 0 Then Call PivotList.Add(myRange.Cells(r, UnpivotColNum))
  Next
  
  'figure out the range at which the header value is unique. After that it starts repeating itself, so quit
  For r = 2 To myRange.Rows.Count
    If CollectionGetIndex(HeaderList, myRange.Cells(r, HeaderColNum)) = 0 Then
        Call HeaderList.Add(myRange.Cells(r, HeaderColNum))
    Else
        Exit For
    End If
  Next
  
  'redim the results array
  ReDim myArray(HeaderList.Count + 1, PivotList.Count + 1)
  
  'write the column headers
  For c = 1 To PivotList.Count
    myArray(1, c + 1) = PivotList(c)
  Next

  'write the row headers
  For r = 1 To HeaderList.Count
    myArray(r + 1, 1) = HeaderList(r)
  Next

  'write the block
  Dim Header As Variant
  Dim PivotHeader As Variant
  Dim myValue As Variant
  Dim myR As Integer, myC As Integer
  
  For r = 2 To myRange.Rows.Count
    Header = myRange.Cells(r, HeaderColNum)
    PivotHeader = myRange.Cells(r, UnpivotColNum)
    myValue = myRange.Cells(r, ValueColNum)
      
    'now find the row and column number
    myR = CollectionGetIndex(HeaderList, Header) + 1
    myC = CollectionGetIndex(PivotList, PivotHeader) + 1
    myArray(myR, myC) = myValue
  Next
    
  If Not WorkSheetExists(NewSheetName) Then
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = NewSheetName
    Set newSheet = ActiveWorkbook.Sheets(NewSheetName)
    newSheet.Range(ActiveSheet.Cells(1, 1), ActiveSheet.Cells(UBound(myArray, 1) + 1, UBound(myArray, 2) + 1)) = myArray
    UnPivot = True
  Else
    MsgBox ("Worksheet " & NewSheetName & " already exists. Please remove the old one first.")
    UnPivot = False
  End If

  
End Function

Public Function CollectionContains(Col As Collection, key As Variant) As Boolean
Dim obj As Variant
On Error GoTo err
    CollectionContains = True
    obj = Col(key)
    Exit Function
err:
    CollectionContains = False
End Function

Public Function CollectionGetIndex(ByRef Col As Collection, key As Variant) As Integer
    Dim Idx As Integer
    For Idx = 1 To Col.Count
        If Col(Idx) = key Then
            CollectionGetIndex = Idx
            Exit Function
        End If
    Next
    CollectionGetIndex = 0
End Function


Public Sub UnPivot2CSV(ByRef myRange As Range, StartDataCol As Integer, ResultsFile As String, Delimiter As String, DataColName As String)
  'This routine creates a new worksheet and writes data from the original sheet in an unpivoted way to csv
  Dim r As Long, c As Long, fn As Long
  Dim BaseStr As String, DataStr As String, myStr As String
  
  fn = FreeFile
  Open ResultsFile For Output As #fn
  
  'WRITE THE HEADER
  myStr = myRange.Cells(1, 1)
  If StartDataCol > 2 Then
    For c = 2 To StartDataCol - 1
      myStr = myStr & "," & myRange.Cells(1, c)
    Next
  End If
  myStr = myStr & "," & DataColName
  Print #fn, myStr
  
  'WRITE THE DATA
  For r = 2 To myRange.Rows.Count
    myStr = myRange.Cells(r, 1)
    If StartDataCol > 2 Then
      For c = 2 To StartDataCol - 1
        BaseStr = BaseStr & Delimiter & myRange.Cells(r, c)
      Next
    End If
    
    For c = StartDataCol To myRange.Columns.Count
      If myRange.Cells(r, c) <> "" Then
        DataStr = myRange.Cells(1, c)
        Print #fn, BaseStr & Delimiter & DataStr
      End If
    Next
  Next
  
  Close (fn)
  
End Sub

Public Sub Range2CSV(ByRef myRange As Range, ResultsFile As String, Delimiter As String, Append As Boolean, WriteHeader As Boolean)
  'This routine creates a new worksheet and writes data from the original sheet in an unpivoted way to csv
  Dim r As Long, c As Long, fn As Long, tmpStr As String
  
  fn = FreeFile
  If Append Then
      Open ResultsFile For Append As #fn
  Else
      Open ResultsFile For Output As #fn
  End If
  
  'first write the header
  If WriteHeader Then
    tmpStr = myRange.Cells(1, 1)
    For c = 2 To myRange.Columns.Count
      tmpStr = tmpStr & Delimiter & myRange.Cells(1, c)
    Next
    Print #fn, tmpStr
  End If
  
  'next write the data
  For r = 2 To myRange.Rows.Count
    DoEvents
    tmpStr = myRange.Cells(r, 1)
    For c = 2 To myRange.Columns.Count
      tmpStr = tmpStr & Delimiter & myRange.Cells(r, c)
    Next
    Print #fn, tmpStr
  Next
  Close (fn)
  
End Sub

Public Sub GoalSeekMultiple(ByRef GoalCell As Range, myGoal As Double, ByRef myRange As Range)
  'this function attempts to optimize a cell by adjusting values in multiple cells
  'it is a fairly simple approach, so it won't always work!!!
  'the routine optimizes by adjusting only one cell at a time
  Dim r As Long, c As Long
  For r = 1 To myRange.Rows.Count
    For c = 1 To myRange.Columns.Count
      GoalSeekMultiple = GoalCell.GoalSeek(myGoal, myRange.Cells(r, c))
    Next
  Next

End Sub


Public Sub GoalSeekTriple(ByRef GoalCell As Range, myGoal As Double, Adjust As Range, l1 As Double, u1 As Double, l2 As Double, u2 As Double, l3 As Double, u3 As Double, nIterations As Integer)
  Dim r As Long, c As Long, i As Integer, j As Long, K As Long, nIter As Integer
  Dim minI As Integer, minJ As Integer, minK As Integer
  Dim Range1 As Double, Range2 As Double, range3 As Double
  Dim rowIdx As Integer, colIdx As Integer
  Dim myErr As Double, minErr As Double
  
  If Adjust.Count <> 3 Then
    MsgBox ("Error: range must contain 3 cells.")
    End
  End If
  
  Range1 = u1 - l1
  Range2 = u2 - l2
  range3 = u3 - l3
  
  Dim Results(10, 10, 10) As Variant
  
  For nIter = 1 To nIterations
  
    For i = 1 To 10
      Adjust.Cells(1, 1) = l1 + (i - 0.5) * (u1 - l1) / 10
      For j = 1 To 10
      
        If Adjust.Rows.Count > 1 Then
          Adjust.Cells(2, 1) = l2 + (j - 0.5) * (u2 - l2) / 10
        ElseIf Adjust.Columns.Count > 1 Then
          Adjust.Cells(1, 2) = l2 + (j - 0.5) * (u2 - l2) / 10
        End If
      
        For K = 1 To 10
          If Adjust.Rows.Count > 1 Then
            Adjust.Cells(3, 1) = l3 + (K - 0.5) * (u3 - l3) / 10
          ElseIf Adjust.Columns.Count > 1 Then
            Adjust.Cells(1, 3) = l3 + (K - 0.5) * (u3 - l3) / 10
          End If
        
          'set the values for the 10x10x10 matrix
          If IsNumeric(GoalCell.Value) Then
            Results(i, j, K) = GoalCell.Value
          Else
            Results(i, j, K) = 99999999
          End If
        Next
      Next
    Next
    
    'find the value that's closest to the target
    minErr = 99999999
    For i = 1 To 10
      For j = 1 To 10
        For K = 1 To 10
           myErr = Math.Abs(Results(i, j, K) - myGoal)
           If myErr < minErr Then
             minI = i
             minJ = j
             minK = K
             minErr = myErr
           End If
        Next
      Next
    Next
    
    'set the final value
    If Adjust.Rows.Count > 1 Then
      Adjust.Cells(1, 1) = l1 + (minI - 0.5) * (u1 - l1) / 10
      Adjust.Cells(2, 1) = l2 + (minJ - 0.5) * (u2 - l2) / 10
      Adjust.Cells(3, 1) = l3 + (minK - 0.5) * (u3 - l3) / 10
    ElseIf Adjust.Columns.Count > 1 Then
      Adjust.Cells(1, 1) = l1 + (minI - 0.5) * (u1 - l1) / 10
      Adjust.Cells(1, 2) = l2 + (minJ - 0.5) * (u2 - l2) / 10
      Adjust.Cells(1, 3) = l3 + (minK - 0.5) * (u3 - l3) / 10
    End If
    
    'adjust the boundaries to initiate the next iteration
    l1 = l1 + (minI - 1) * Range1 / 10
    u1 = l1 + Range1 / 10
    l2 = l2 + (minJ - 1) * Range2 / 10
    u2 = l2 + Range2 / 10
    l3 = l3 + (minK - 1) * range3 / 10
    u3 = l3 + range3 / 10
    Range1 = u1 - l1
    Range2 = u2 - l2
    range3 = u3 - l3
  
  Next
  

  

End Sub

Public Sub GoalSeekDouble(ByRef GoalCell As Range, myGoal As Double, Adjust As Range, l1 As Double, u1 As Double, l2 As Double, u2 As Double, nIterations As Integer)
  Dim r As Long, c As Long, i As Integer, j As Long, nIter As Integer
  Dim minI As Integer, minJ As Integer
  Dim Range1 As Double, Range2 As Double
  Dim rowIdx As Integer, colIdx As Integer
  Dim myErr As Double, minErr As Double
  
  Range1 = u1 - l1
  Range2 = u2 - l2
  
  If Adjust.Count <> 2 Then
    MsgBox ("Error: range must contain 2 cells.")
    End
  End If

  
  Dim Results(10, 10) As Variant
  
  For nIter = 1 To nIterations
  
    For i = 1 To 10
      Adjust.Cells(1, 1) = l1 + (i - 0.5) * (u1 - l1) / 10
      For j = 1 To 10
        If Adjust.Rows.Count > 1 Then
          Adjust.Cells(2, 1) = l2 + (j - 0.5) * (u2 - l2) / 10
        ElseIf Adjust.Columns.Count > 1 Then
          Adjust.Cells(1, 2) = l2 + (j - 0.5) * (u2 - l2) / 10
        End If
      
        'set the values for the 10x10 matrix
        If IsNumeric(GoalCell.Value) Then
          Results(i, j) = GoalCell.Value
        Else
          Results(i, j) = 99999999
        End If
      Next
    Next
    
    'find the value that's closest to the target
    minErr = 99999999
    For i = 1 To 10
      For j = 1 To 10
        myErr = Math.Abs(Results(i, j) - myGoal)
        If myErr < minErr Then
          minI = i
          minJ = j
          minErr = myErr
        End If
      Next
    Next
    
    'set the final value
    If Adjust.Rows.Count > 1 Then
      Adjust.Cells(1, 1) = l1 + (minI - 0.5) * (u1 - l1) / 10
      Adjust.Cells(2, 1) = l2 + (minJ - 0.5) * (u2 - l2) / 10
    ElseIf Adjust.Columns.Count > 1 Then
      Adjust.Cells(1, 1) = l1 + (minI - 0.5) * (u1 - l1) / 10
      Adjust.Cells(1, 2) = l2 + (minJ - 0.5) * (u2 - l2) / 10
    End If
    
    'adjust the boundaries to initiate the next iteration
    l1 = l1 + (minI - 1) * Range1 / 10
    u1 = l1 + Range1 / 10
    l2 = l2 + (minJ - 1) * Range2 / 10
    u2 = l2 + Range2 / 10
    Range1 = u1 - l1
    Range2 = u2 - l2
  Next

End Sub

Public Function COLUMN_NUMBER(ByVal myVal As Variant, ByVal myRange As Range) As Long
  Dim c As Long
  For c = 1 To myRange.Columns.Count
    If myRange.Cells(1, c) = myVal Then
      COLUMN_NUMBER = c
      Exit Function
    End If
  Next
  COLUMN_NUMBER = 0
End Function

Public Sub PrintArray(ByRef Data As Variant, ByRef Cl As Range)
    Cl.Resize(UBound(Data, 1), UBound(Data, 2)) = Data
End Sub

Public Function RANGEVERTASCENDING(ByRef myRange As Range, Optional AllowEqualValues As Boolean = True) As Boolean
  Dim i As Long
  
  If AllowEqualValues Then
    For i = 1 To myRange.Rows.Count - 1
      If myRange(i, 1) > myRange(i + 1, 1) Then
        RANGEVERTASCENDING = False
        Exit Function
      End If
    Next
  Else
    For i = 1 To myRange.Rows.Count - 1
      If myRange(i, 1) >= myRange(i + 1, 1) Then
        RANGEVERTASCENDING = False
        Exit Function
      End If
    Next
  End If
  
  RANGEVERTASCENDING = True
  
End Function

Public Function VALUEFROMCELLADDRESS(Row As Integer, Col As Integer, SheetName As String) As Variant
    'deze functie geeft een celwaarde terug op basis van het adres (rij, kolom)
    'merk op dat het ook direct als werkbladfunctie kan: INDIRECT(ADDRESS(RIJ, KOLOM,,,WERKBLADNAAM))
    VALUEFROMCELLADDRESS = ActiveWorkbook.Sheets(SheetName).Cells(Row, Col)
End Function

Public Function FormatRoman(ByVal n As Integer) As String
   ' Author: Christian d'Heureuse (www.source-code.biz)
   If n = 0 Then VBA.FormatRoman = "0": Exit Function
      ' There is no roman symbol for 0, but we don't want to return an empty string.
   Const r = "IVXLCDM"              ' roman symbols
   Dim i As Integer: i = Abs(n)
   Dim s As String, P As Integer
   For P = 1 To 5 Step 2
      Dim D As Integer: D = i Mod 10: i = i \ 10
      Select Case D                 ' VBA.Format a decimal digit
         Case 0 To 3: s = String(D, VBA.Mid(r, P, 1)) & s
         Case 4:      s = VBA.Mid(r, P, 2) & s
         Case 5 To 8: s = VBA.Mid(r, P + 1, 1) & String(D - 5, VBA.Mid(r, P, 1)) & s
         Case 9:      s = VBA.Mid(r, P, 1) & VBA.Mid(r, P + 2, 1) & s
         End Select
      Next
   s = String(i, "M") & s           ' VBA.Format thousands
   If n < 0 Then s = "-" & s        ' insert sign if negative (non-standard)
   VBA.FormatRoman = s

End Function

Public Function LSHA2MMPD(myVal As Double) As Double
  Dim newVal As Double
  newVal = myVal * 3600 * 24 / 10000
  LSHA2MMPD = newVal
End Function

Public Function MMPD2LSHA(myVal As Double) As Double
  Dim newVal As Double
  newVal = myVal / 3600 / 24 * 10000
  MMPD2LSHA = newVal
End Function

Public Function M3PS2MMPD(CAP As Double, Opp As Double) As Double
  'cap in m3/s
  'opp in m2
  M3PS2MMPD = CAP / Opp * 1000 * 3600 * 24
End Function

Public Function M3PS2MMPU(CAP As Double, Opp As Double) As Double
  'cap in m3/s
  'opp in m2
  If Opp > 0 Then
    M3PS2MMPU = CAP / Opp * 1000 * 3600
  Else
    M3PS2MMPU = 0
  End If
End Function

Public Function MMPU2M3PS(CAP As Double, Opp As Double) As Double
  'cap in mm/u
  'opp in m2
  MMPU2M3PS = CAP / 3600 / 1000 * Opp
End Function

Public Function MMPD2M3PS(CAP As Double, Opp As Double) As Double
  'cap in mm/d
  'opp in m2
  MMPD2M3PS = CAP / 1000 / 24 / 3600 * Opp
End Function


Public Function Celcius2Kelvin(Celcius As Double)
  Celcius2Kelvin = Celcius + 273.15
End Function
Public Function Kelvin2Celcius(Kelvin As Double)
  Kelvin2Celcius = Kelvin - 273.15
End Function

Public Function RD2LATLONG(x As Double, y As Double, Optional ByRef Latitude As Double = 0, Optional ByRef Longitude As Double = 0) As String
  Dim dX As Double, dY As Double
  Dim SomN As Double, SomE As Double

  dX = (x - 155000) * 10 ^ (-5)
  dY = (y - 463000) * 10 ^ (-5)
  SomN = (3235.65389 * dY) + (-32.58297 * dX ^ 2) + (-0.2475 * dY ^ 2) + (-0.84978 * dX ^ 2 * dY) + (-0.0655 * dY ^ 3) + (-0.01709 * dX ^ 2 * dY ^ 2) + (-0.00738 * dX) + (0.0053 * dX ^ 4) + (-0.00039 * dX ^ 2 * dY ^ 3) + (0.00033 * dX ^ 4 * dY) + (-0.00012 * dX * dY)
  SomE = (5260.52916 * dX) + (105.94684 * dX * dY) + (2.45656 * dX * dY ^ 2) + (-0.81885 * dX ^ 3) + (0.05594 * dX * dY ^ 3) + (-0.05607 * dX ^ 3 * dY) + (0.01199 * dY) + (-0.00256 * dX ^ 3 * dY ^ 2) + (0.00128 * dX * dY ^ 4) + (0.00022 * dY ^ 2) + (-0.00022 * dX ^ 2) + (0.00026 * dX ^ 5)
  Latitude = 52.15517 + (SomN / 3600)
  Longitude = 5.387206 + (SomE / 3600)
 
  RD2LATLONG = Latitude & ";" & Longitude

End Function

Public Function RD2LAT(x As Double, y As Double) As Double
  Dim Latitude As Double, Longitude As Double
  Call RD2LATLONG(x, y, Latitude, Longitude)
  RD2LAT = Latitude
End Function
Public Function RD2LON(x As Double, y As Double) As Double
  Dim Latitude As Double, Longitude As Double
  Call RD2LATLONG(x, y, Latitude, Longitude)
  RD2LON = Longitude
End Function

Public Function RD2WGS84(x As Double, y As Double, Optional ByRef Lat As Double = 0, Optional ByRef Lon As Double = 0) As String
  'converteert RD-coordinaten naar Lat/Long (WGS84)
  'maakt gebruik van de routines van Ejo Schrama: schrama @geo.tudelft.nl
  Dim phi As Double
  Dim lambda As Double
  Call RD2BESSEL(x, y, phi, lambda)
  Call BESSEL2WGS84(phi, lambda, Lat, Lon)
  RD2WGS84 = Lat & "," & Lon
End Function

Public Function WGS842RD(Lat As Double, Lon As Double, Optional ByRef x As Double = 0, Optional ByRef y As Double = 0) As String
  'converteert WGS84-coordinaten (Lat/Long) naar RD
  'maakt gebruik van de routines van Ejo Schrama: schrama @geo.tudelft.nl
  Dim phiBes As Double
  Dim LambdaBes As Double
  Call WGS842BESSEL(Lat, Lon, phiBes, LambdaBes)
  Call BESSEL2RD(phiBes, LambdaBes, x, y)
  WGS842RD = x & "," & y
  
End Function

Public Function WGS842X(Lat As Double, Lon As Double) As Double
  'converteert WGS84-coordinaten (Lat/Long) naar RD (alleen de X-coordinaat)
  'maakt gebruik van de routines van Ejo Schrama: schrama @geo.tudelft.nl
  Dim x As Double, y As Double
  Dim phiBes As Double
  Dim LambdaBes As Double
  Call WGS842BESSEL(Lat, Lon, phiBes, LambdaBes)
  Call BESSEL2RD(phiBes, LambdaBes, x, y)
  WGS842X = x
  
End Function

Public Function WGS842Y(Lat As Double, Lon As Double) As Double
  'converteert WGS84-coordinaten (Lat/Long) naar RD (alleen de Y-coordinaat)
  'maakt gebruik van de routines van Ejo Schrama: schrama @geo.tudelft.nl
  Dim x As Double, y As Double
  Dim phiBes As Double
  Dim LambdaBes As Double
  Call WGS842BESSEL(Lat, Lon, phiBes, LambdaBes)
  Call BESSEL2RD(phiBes, LambdaBes, x, y)
  WGS842Y = y
  
End Function

Public Function WGS84DEG2DECIMAL(Deg As String) As String
  'converts WGS84 coordinates from degrees to decimal
  Dim tmpStr As String, Pos As Integer, startPos As Integer
  Dim DegNB As Double, MinNB As Double, SecNB As Double, Northing As Double
  Dim DegOL As Double, MinOL As Double, SecOL As Double, Easting As Double
  Deg = VBA.Trim(Deg)
  
  'find out where the actual value for northing begins and clean the string up
  For Pos = 1 To Len(Deg)
   If IsNumeric(VBA.Mid(Deg, Pos, 1)) Then
     startPos = Pos
     Exit For
   End If
  Next
  Deg = VBA.Right(Deg, Len(Deg) - startPos + 1)
  
  'determine the coordinate for northing
  DegNB = Val(ParseString(Deg, "°", 0))
  MinNB = Val(ParseString(Deg, "'", 0)) / 60
  SecNB = Val(ParseString(Deg, Chr(34), 0)) / 3600
  Northing = DegNB + MinNB + SecNB
  
  'find out where the actual value for easting begins
  For Pos = 1 To Len(Deg)
   If IsNumeric(VBA.Mid(Deg, Pos, 1)) Then
     startPos = Pos
     Exit For
   End If
  Next
  Deg = VBA.Right(Deg, Len(Deg) - startPos + 1)
  
  'retrieve the coordinates for Easting
  DegOL = Val(ParseString(Deg, "°", 0))
  MinOL = Val(ParseString(Deg, "'", 0)) / 60
  SecOL = Val(ParseString(Deg, Chr(34), 0)) / 3600
  Easting = DegOL + MinOL + SecOL
    
  WGS84DEG2DECIMAL = Northing & "," & Easting

End Function

Public Function WGS84DEG2LATDECIMAL(Deg As String) As String
  'converts WGS84 coordinates from degrees to Latitude in Decimals
  Dim Decimals As String
  Decimals = WGS84DEG2DECIMAL(Deg)
  WGS84DEG2LATDECIMAL = ParseString(Decimals, ",")
End Function

Public Function WGS84DEG2LONDECIMAL(Deg As String) As String
  'converts WGS84 coordinates from degrees to Latitude in Decimals
  Dim Decimals As String, tmpStr As String
  Decimals = WGS84DEG2DECIMAL(Deg)
  tmpStr = ParseString(Decimals, ",")
  WGS84DEG2LONDECIMAL = Decimals
End Function

Public Sub RD2BESSEL(x As Double, y As Double, ByRef phi As Double, ByRef lambda As Double)

'converteert RD-coordinaten naar phi en lambda voor een Bessel-functie
'code is geheel gebaseerd op de routines van Ejo Schrama's software:
'schrama@geo.tudelft.nl

Dim x0 As Double
Dim y0 As Double
Dim K As Double
Dim bigr As Double
Dim m As Double
Dim n As Double
Dim lambda0 As Double
Dim phi0 As Double
Dim l0 As Double
Dim b0 As Double
Dim e As Double
Dim A As Double

Dim d_1 As Double, d_2 As Double, r As Double, sa As Double, ca As Double, psi As Double, cpsi As Double, spsi As Double
Dim sb As Double, cb As Double, b As Double, sdl As Double, dl As Double, W As Double, Q As Double, phiprime As Double
Dim dq As Double, i As Long, pi As Double

x0 = 155000
y0 = 463000
K = 0.9999079
bigr = 6382644.571
m = 0.003773953832
n = 1.00047585668

pi = Application.WorksheetFunction.pi
'pi = 3.14159265358979
lambda0 = pi * 2.99313271611111E-02
phi0 = pi * 0.289756447533333
l0 = pi * 2.99313271611111E-02
b0 = pi * 0.289561651383333

e = 0.08169683122
A = 6377397.155

d_1 = x - x0
d_2 = y - y0
r = Sqr(d_1 ^ 2 + d_2 ^ 2)

If r <> 0 Then
  sa = d_1 / r
  ca = d_2 / r
Else
  sa = 0
  ca = 0
End If

psi = Application.WorksheetFunction.ATan2(K * 2 * bigr, r) * 2
cpsi = Cos(psi)
spsi = Sin(psi)

sb = ca * Cos(b0) * spsi + Sin(b0) * cpsi
d_1 = sb
cb = Sqr(1 - d_1 ^ 2)
b = Application.WorksheetFunction.Acos(cb)
sdl = sa * spsi / cb
dl = Application.WorksheetFunction.Asin(sdl)
lambda = dl / n + lambda0
W = Application.WorksheetFunction.Ln(Tan(b / 2 + pi / 4))
Q = (W - m) / n

phi = Atn(Exp(1) ^ Q) * 2 - pi / 2 'phi prime
For i = 1 To 4
  dq = e / 2 * Application.WorksheetFunction.Ln((e * Sin(phi) + 1) / (1 - e * Sin(phi)))
  phi = Atn(Exp(1) ^ (Q + dq)) * 2 - pi / 2
Next

lambda = lambda / pi * 180
phi = phi / pi * 180

End Sub

Public Sub BESSEL2WGS84(phi As Double, lambda As Double, ByRef PhiWGS As Double, ByRef LamWGS As Double)
  Dim dphi As Double, dlam As Double, phicor As Double, lamcor As Double

  dphi = phi - 52
  dlam = lambda - 5
  phicor = (-96.862 - dphi * 11.714 - dlam * 0.125) * 0.00001
  lamcor = (dphi * 0.329 - 37.902 - dlam * 14.667) * 0.00001
  PhiWGS = phi + phicor
  LamWGS = lambda + lamcor


End Sub

Public Sub WGS842BESSEL(PhiWGS As Double, LamWGS As Double, ByRef phi As Double, ByRef lambda As Double)
  Dim dphi As Double, dlam As Double, phicor As Double, lamcor As Double

  dphi = PhiWGS - 52
  dlam = LamWGS - 5
  phicor = (-96.862 - dphi * 11.714 - dlam * 0.125) * 0.00001
  lamcor = (dphi * 0.329 - 37.902 - dlam * 14.667) * 0.00001
  phi = PhiWGS - phicor
  lambda = LamWGS - lamcor
  
End Sub

Public Sub BESSEL2RD(phiBes As Double, lamBes As Double, ByRef x As Double, ByRef y As Double)

'converteert Lat/Long van een Bessel-functie naar X en Y in RD
'code is geheel gebaseerd op de routines van Ejo Schrama's software:
'schrama@geo.tudelft.nl

Dim x0 As Double
Dim y0 As Double
Dim K As Double
Dim bigr As Double
Dim m As Double
Dim n As Double
Dim lambda0 As Double
Dim phi0 As Double
Dim l0 As Double
Dim b0 As Double
Dim e As Double
Dim A As Double

Dim d_1 As Double, d_2 As Double, r As Double, sa As Double, ca As Double, psi As Double, cpsi As Double, spsi As Double
Dim sb As Double, cb As Double, b As Double, sdl As Double, dl As Double, W As Double, Q As Double, phiprime As Double
Dim dq As Double, i As Long, pi As Double, phi As Double, lambda As Double, s2psihalf As Double, cpsihalf As Double, spsihalf As Double
Dim tpsihalf As Double

x0 = 155000
y0 = 463000
K = 0.9999079
bigr = 6382644.571
m = 0.003773953832
n = 1.00047585668

pi = Application.WorksheetFunction.pi
'pi = 3.14159265358979
lambda0 = pi * 2.99313271611111E-02
phi0 = pi * 0.289756447533333
l0 = pi * 2.99313271611111E-02
b0 = pi * 0.289561651383333

e = 0.08169683122
A = 6377397.155

phi = phiBes / 180 * pi
lambda = lamBes / 180 * pi

Q = Application.WorksheetFunction.Ln(Tan(phi / 2 + pi / 4))
dq = e / 2 * Application.WorksheetFunction.Ln((e * Sin(phi) + 1) / (1 - e * Sin(phi)))
Q = Q - dq
W = n * Q + m
b = Atn(Exp(1) ^ W) * 2 - pi / 2
dl = n * (lambda - lambda0)
d_1 = Sin((b - b0) / 2)
d_2 = Sin(dl / 2)
s2psihalf = d_1 * d_1 + d_2 * d_2 * Cos(b) * Cos(b0)
cpsihalf = Sqr(1 - s2psihalf)
spsihalf = Sqr(s2psihalf)
tpsihalf = spsihalf / cpsihalf
spsi = spsihalf * 2 * cpsihalf
cpsi = 1 - s2psihalf * 2
sa = Sin(dl) * Cos(b) / spsi
ca = (Sin(b) - Sin(b0) * cpsi) / (Cos(b0) * spsi)
r = K * 2 * bigr * tpsihalf
x = Round(r * sa + x0, 0)
y = Round(r * ca + y0, 0)

End Sub

Public Function MultiParse(ByRef myString As String, returnInstanceNumber As Integer, Optional Delimiter As String = " ", Optional QuoteHandlingFlag As Long = 1) As String
  Dim tmpString As String, i As Long
  For i = 1 To returnInstanceNumber
    tmpString = ParseString(myString, Delimiter, QuoteHandlingFlag)
  Next
  MultiParse = tmpString
End Function

Public Function ParseNumeric(ByRef myString As String) As String
  Dim i As Integer, myChar As String, Done As Boolean
  'knabbelt net zo lang een karakter van de linker kant van een string af tot het resultaat niet langer numeriek is
  
  While Not Done
    myChar = VBA.Left(myString, 1)
    If Not (IsNumeric(myChar) Or myChar = ".") Then
      Exit Function
    Else
      ParseNumeric = ParseNumeric & myChar
      myString = Right(myString, VBA.Len(myString) - 1)
    End If
  Wend
  
End Function

Public Function ParseString(ByRef myString As String, Optional Delimiter As String = " ", Optional QuoteHandlingFlag As Long = 1) As String

Dim Pos As Long, quoteEven As Boolean
quoteEven = True

'Quotehandlingflag: default = 1
'0 = items between quotes are NOT being treated as separate items (parsing also between quotes)
'1 = items between single quotes are being treated as separate items (no parsing between single quotes)
'2 = items between double quotes are being treated as separate items (no parsing between double quotes)

Dim i As Long
For i = 1 To VBA.Len(myString)
  
  'als we een quote tegenkomen, houden we bij of het even of oneven is. zo weten we of we een omsloten object hebben
  If VBA.Left(myString, 1) = "'" And QuoteHandlingFlag = 1 Then
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If quoteEven Then
      quoteEven = False
      'we voegen niets toe aan de geparste string want quotes zelf doen niet mee
    Else
      quoteEven = True
      'we hebben een omsloten object gevonden dus is weer even
      If VBA.Len(myString) > 0 Then myString = VBA.Right(myString, VBA.Len(myString) - 1)
      Exit Function
    End If
  
  ElseIf VBA.Left(myString, 1) = VBA.Chr(34) And QuoteHandlingFlag = 2 Then 'double quote encountered
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If quoteEven Then
      quoteEven = False
      'we voegen niets toe aan de geparste string want quotes zelf doen niet mee
    Else
      quoteEven = True
      'we hebben een omsloten object gevonden dus is weer even
      If VBA.Len(myString) > 0 Then myString = VBA.Right(myString, VBA.Len(myString) - 1)
      Exit Function
    End If
  'als het teken gelijk is aan de delimiter, kijken we of we al geldige tekens hadden gevonden
  'zo ja, wegschrijven
  ElseIf VBA.Left(myString, 1) = Delimiter And QuoteHandlingFlag = 1 And quoteEven = True Then
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If VBA.Len(ParseString) > 0 Then
      Exit Function
    End If
  ElseIf VBA.Left(myString, 1) = Delimiter And QuoteHandlingFlag = 2 And quoteEven = True Then
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If VBA.Len(ParseString) > 0 Then
      Exit Function
    End If
  ElseIf VBA.Left(myString, 1) = Delimiter And QuoteHandlingFlag = 0 Then
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If VBA.Len(ParseString) > 0 Then
      Exit Function
    End If
  Else
    'hier gebeurt het werkelijke parsen
    ParseString = ParseString & VBA.Left(myString, 1)
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
  End If
Next

End Function

Public Function ParseStringPlus(ByRef myString As String, ByRef QuotesFound As Boolean, Optional Delimiter As String = " ", Optional QuoteHandlingFlag As Long = 1) As String

Dim Pos As Long, quoteEven As Boolean
quoteEven = True
QuotesFound = False

'Differences with ParseString:
'- Uses a byref boolean to return whether an item surrounded by quoted was found

'Quotehandlingflag: default = 1
'0 = items between quotes are NOT being treated as separate items (parsing also between quotes)
'1 = items between single quotes are being treated as separate items (no parsing between single quotes)
'2 = items between double quotes are being treated as separate items (no parsing between double quotes)

Dim i As Long
For i = 1 To VBA.Len(myString)
  
  'als we een quote tegenkomen, houden we bij of het even of oneven is. zo weten we of we een omsloten object hebben
  If VBA.Left(myString, 1) = "'" And QuoteHandlingFlag = 1 Then
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If quoteEven Then
      quoteEven = False
      'we voegen niets toe aan de geparste string want quotes zelf doen niet mee
    Else
      quoteEven = True
      'we hebben een omsloten object gevonden dus is weer even
      If VBA.Len(myString) > 0 Then
        myString = VBA.Right(myString, VBA.Len(myString) - 1)
        QuotesFound = True
        Exit Function
      End If
    End If
  
  ElseIf VBA.Left(myString, 1) = VBA.Chr(34) And QuoteHandlingFlag = 2 Then 'double quote encountered
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If quoteEven Then
      quoteEven = False
      'we voegen niets toe aan de geparste string want quotes zelf doen niet mee
    Else
      quoteEven = True
      'we hebben een omsloten object gevonden dus is weer even
      If VBA.Len(myString) > 0 Then
        myString = VBA.Right(myString, VBA.Len(myString) - 1)
        QuotesFound = True
        Exit Function
      End If
    End If
  'als het teken gelijk is aan de delimiter, kijken we of we al geldige tekens hadden gevonden
  'zo ja, wegschrijven
  ElseIf VBA.Left(myString, 1) = Delimiter And QuoteHandlingFlag = 1 And quoteEven = True Then
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If VBA.Len(ParseStringPlus) > 0 Then
      Exit Function
    End If
  ElseIf VBA.Left(myString, 1) = Delimiter And QuoteHandlingFlag = 2 And quoteEven = True Then
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If VBA.Len(ParseStringPlus) > 0 Then
      Exit Function
    End If
  ElseIf VBA.Left(myString, 1) = Delimiter And QuoteHandlingFlag = 0 Then
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If VBA.Len(ParseStringPlus) > 0 Then
      Exit Function
    End If
  ElseIf VBA.Left(myString, 1) = vbCrLf Then
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If VBA.Len(ParseStringPlus) > 0 Then
      Exit Function
    End If
  Else
    'hier gebeurt het werkelijke parsen
    ParseStringPlus = ParseStringPlus & VBA.Left(myString, 1)
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
  End If
Next

End Function

Public Sub TextSnippet(startPos As Long, endPos As Long, myString As String, ByRef LeftSnippet As String, ByRef Snippet As String, ByRef RightSnippet As String)
  'cuts a string in three parts based on two string positions
  LeftSnippet = Left(myString, startPos - 1)
  Snippet = Mid(myString, startPos, endPos - startPos + 1)
  RightSnippet = VBA.Right(myString, Len(myString) - endPos)
End Sub

Public Function BNAString(ID As String, Name As String, x As Double, y As Double) As String
  BNAString = VBA.Chr(34) & ID & VBA.Chr(34) & "," & VBA.Chr(34) & Name & VBA.Chr(34) & ",1," & x & "," & y
End Function

Public Function WagModSTAString(myDate As Double, Prec As Double, EVAP As Double, Qmeas As Double) As String
  Dim TimeStr As String
  Dim myPrec As String, myEvap As String, myQm As String
  
  If Hour(myDate) = 0 Then
    TimeStr = "0"
  Else
    TimeStr = VBA.Trim(Hour(myDate) & "00")
  End If
  
  While VBA.Len(TimeStr) < 4
    TimeStr = " " & TimeStr
  Wend
  
  myPrec = VBA.Format(Prec, "0.000")
  While VBA.Len(myPrec) < 13
    myPrec = " " & myPrec
  Wend

  myEvap = VBA.Format(EVAP, "0.000")
  While VBA.Len(myEvap) < 8
    myEvap = " " & myEvap
  Wend
  
  myQm = VBA.Format(Qmeas, "0.000")
  While VBA.Len(myQm) < 8
    myQm = " " & myQm
  Wend

  WagModSTAString = Year(myDate) & "/" & VBA.Format(Month(myDate), "00") & "/" & VBA.Format(Day(myDate), "00") & " " & TimeStr & " " & myPrec & " " & myEvap & " " & myQm
  
  
End Function

Public Function WalrusDATString(myDate As Double, Prec As Double, EVAP As Double, Qmeas As Double) As String
  Dim TimeStr As String
  Dim myPrec As String, myEvap As String, myQm As String

  WalrusDATString = Year(myDate) & VBA.Format(Month(myDate), "00") & VBA.Format(Day(myDate), "00") & VBA.Format(Hour(myDate), "00") & " " & VBA.Format(Prec, "0.0000") & " " & VBA.Format(EVAP, "0.0000") & " " & VBA.Format(Qmeas, "0.0000") & " 0 0 0 0"
  
End Function

'Binary Conversions
'The Functions in this module are designed to aid in working with BINARY
'numbers. Visual Basic does not include nor allow any representation of a
'number in binary VBA.Format.  Therefore, all of these functions work strictly on
'strings.  All of the parameters passed into them and returned from them are
'strings.
'
'              CONVERSION NEEDED                 FUNCTION
'            ------------------------------------------------------
'              Binary to Hex            BinToHex(BinNum As String)
'              Binary to Octal          BinToOct(BinNum As String)
'              Binary to Decimal        BinToDec(BinNum As String)
'              Hex to Binary            HexToBin(HexNum As String)
'              Octal to Binary          OctToBin(OctNum As String)
'              Decimal to Binary        DecToBin(DecNum As String)
'
'
Function BinToHex(BinNum As String) As String
   Dim BinLen As Integer, i As Integer
   Dim HexNum As Variant
   
   On Error GoTo errorhandler
   BinNum = VBA.Trim(BinNum)
   BinLen = VBA.Len(BinNum)
   For i = BinLen To 1 Step -1
'     Check the string for invalid characters
      If Asc(VBA.Mid(BinNum, i, 1)) < 48 Or _
         Asc(VBA.Mid(BinNum, i, 1)) > 49 Then
         HexNum = ""
         err.Raise 1002, "BinToHex", "Invalid Input"
      End If
'     Calculate HEX value of BinNum
      If VBA.Mid(BinNum, i, 1) And 1 Then
         HexNum = HexNum + 2 ^ Abs(i - BinLen)
      End If
   Next i
'  Return HexNum as String
   BinToHex = Hex(HexNum)
errorhandler:
End Function

Function BinToOct(BinNum As String) As String
   Dim BinLen As Integer, i As Integer
   Dim OctNum As Variant
   
   On Error GoTo errorhandler
   BinNum = VBA.Trim(BinNum)
   BinLen = VBA.Len(BinNum)
   For i = BinLen To 1 Step -1
'     Check the string for invalid characters
      If Asc(VBA.Mid(BinNum, i, 1)) < 48 Or _
         Asc(VBA.Mid(BinNum, i, 1)) > 49 Then
         OctNum = ""
         err.Raise 1002, "BinToOct", "Invalid Input"
      End If
'     Calculate Octal value of BinNum
      If VBA.Mid(BinNum, i, 1) And 1 Then
         OctNum = OctNum + 2 ^ Abs(i - BinLen)
      End If
   Next i
'  Return OctNum as String
   BinToOct = Oct(OctNum)
errorhandler:
End Function

Public Function BinToDec(BinNum As String) As String
   Dim i As Integer
   Dim DecNum As Long
   
   On Error GoTo errorhandler
   BinNum = VBA.Trim(BinNum)
'  Loop thru BinString
   For i = VBA.Len(BinNum) To 1 Step -1
'     Check the string for invalid characters
      If Asc(VBA.Mid(BinNum, i, 1)) < 48 Or _
         Asc(VBA.Mid(BinNum, i, 1)) > 49 Then
         DecNum = ""
         err.Raise 1002, "BinToDec", "Invalid Input"
      End If
'     If bit is 1 then raise 2^LoopCount and add it to DecNum
      If VBA.Mid(BinNum, i, 1) And 1 Then
         DecNum = DecNum + 2 ^ (Len(BinNum) - i)
      End If
   Next i
'  Return DecNum as a String
   BinToDec = DecNum
errorhandler:
End Function

Public Function OctToBin(OctNum As String) As String
   Dim BinNum As String
   Dim lOctNum As Long
   Dim i As Integer
   
   On Error GoTo errorhandler
   OctNum = VBA.Trim(OctNum)
'  Check the string for invalid characters
   For i = 1 To VBA.Len(OctNum)
      If (Asc(VBA.Mid(OctNum, i, 1)) < 48 Or Asc(VBA.Mid(OctNum, i, 1)) > 55) Then
         BinNum = ""
         err.Raise 1008, "OctToBin", "Invalid Input"
      End If
   Next i

   i = 0
   lOctNum = Val("&O" & OctNum)
   
   Do
      If lOctNum And 2 ^ i Then
         BinNum = "1" & BinNum
      Else
         BinNum = "0" & BinNum
      End If
      i = i + 1
   Loop Until 2 ^ i > lOctNum
'  Return BinNum as a String
   OctToBin = BinNum
errorhandler:
End Function

Public Function DecToBin(DecNum As String) As String
   Dim BinNum As String
   Dim lDecNum As Long
   Dim i As Integer
   
   On Error GoTo errorhandler
   DecNum = VBA.Trim(DecNum)
   
'  Check the string for invalid characters
   For i = 1 To VBA.Len(DecNum)
      If Asc(VBA.Mid(DecNum, i, 1)) < 48 Or _
         Asc(VBA.Mid(DecNum, i, 1)) > 57 Then
         BinNum = ""
         err.Raise 1010, "DecToBin", "Invalid Input"
      End If
   Next i
   
   i = 0
   lDecNum = Val(DecNum)
   
   Do
      If lDecNum And 2 ^ i Then
         BinNum = "1" & BinNum
      Else
         BinNum = "0" & BinNum
      End If
      i = i + 1
   Loop Until 2 ^ i > lDecNum
'  Return BinNum as a String
   DecToBin = BinNum
errorhandler:
End Function

Public Function HexToBin(HexNum As String) As String
   Dim BinNum As String
   Dim lHexNum As Long
   Dim i As Integer
   
   On Error GoTo errorhandler
   HexNum = VBA.Str(HexNum)
'  Check the string for invalid characters
   For i = 1 To VBA.Len(HexNum)
      If ((Asc(VBA.Mid(HexNum, i, 1)) < 48) Or _
          (Asc(VBA.Mid(HexNum, i, 1)) > 57 And _
           Asc(UCase(VBA.Mid(HexNum, i, 1))) < 65) Or _
          (Asc(UCase(VBA.Mid(HexNum, i, 1))) > 70)) Then
         BinNum = ""
         err.Raise 1016, "HexToBin", "Invalid Input"
      End If
   Next i
   
   i = 0
   lHexNum = Val("&h" & HexNum)
   Do
      If lHexNum And 2 ^ i Then
         BinNum = "1" & BinNum
      Else
         BinNum = "0" & BinNum
      End If
      i = i + 1
   Loop Until 2 ^ i > lHexNum
'  Return BinNum as a String
   HexToBin = BinNum
errorhandler:
End Function
