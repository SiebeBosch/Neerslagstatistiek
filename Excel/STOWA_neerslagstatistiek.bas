Attribute VB_Name = "STOWA_neerslagstatistiek"

Option Explicit


'-------------------------'
'Auteur: Siebe Bosch      '
'Hydroconsult             '
'Lulofsstraat 55, unit 47 '
'2521 AL Den Haag         '
'siebe@hydroconsult.nl    '
'0617682689               '
'-------------------------'

'versie 1.08
'changelog:
'v1.06: onderscheid aangebracht tussen verandergetal jaarrond en winter; ten behoeve van de neerslagscenario's van 2024

Public Function STOWA2015_2018_WINTER_T(ByVal DuurMinuten As Integer, ByVal Zichtjaar As Integer, ByVal Scenario As String, ByVal Corner As String, ByVal Volume As Double) As Double
    
    'deze functie berekent het winterneerslagvolume conform STOWA 2015/2018 met gegeven duur in minuten en volume in mm
    'we berekenen hem in iteratief door gebruik te maken van de functie STOWA2015_2018_NDJF_V
    Dim T_estimate As Double, P As Double
    Dim V_estimate As Double
    
    'Initialiseer de herhalingstijd simpelweg op T=1. Dit lijkt goed uit te pakken voor alle volumes
    T_estimate = 1
    
    Dim Done As Boolean
    Dim iIter As Integer
    Done = False
    iIter = 0
    
    While Not Done
        'nu gaan we de geschatte herhalingstijd invoeren in de functie met actuele statistiek
        V_estimate = STOWA2015_2018_NDJF_V(DuurMinuten, T_estimate, Zichtjaar, Scenario, Corner)
        iIter = iIter + 1
        If iIter = 1000 Then Done = True
        If Math.Abs(V_estimate - Volume) < 0.001 Then Done = True
        T_estimate = T_estimate * Volume / V_estimate               'pas de geschatte herhalingstijd aan naar rato van de afwijking tussen geschat en opgegeven volume
    Wend
    
    STOWA2015_2018_WINTER_T = T_estimate
'
'
'
'    'deze functie berekent het het herhalingstijd van een gegeven winterneerslag conform STOWA, 2015/2018 met gegeven duur in minuten en volume in mm
'    'we berekenen hem in twee iteraties. In de eerste werken we met een geschatte herhalingstijd < 120 jaar.
'    'in de tweede iteratie gebruiken we de herhalingstijd die werd berekend weer als input
'    Dim locpar As Double, scalepar As Double, shapepar As Double, dispcoef As Double
'    Dim V_2018_Huidig As Double, V_2015_Huidig As Double
'    Dim Multiplier As Double
'
'    Dim p As Double
'    If DuurMinuten > 720 Then  'in STOWA 2015 is lange duur gedefinieerd als >= 2 uur maar in 2018 is korte duur gedefinieerd als de duur t/m 720 min (12 uur)
'        dispcoef = GEVDispCoefBasisstatistiek2015WinterLang(DuurMinuten, Zichtjaar, Scenario, Corner)
'        locpar = GEVLocParBasisstatistiek2015WinterLang(DuurMinuten, Zichtjaar, Scenario, Corner)
'        scalepar = dispcoef * locpar
'        shapepar = GEVShapeParBasisstatistiek2015WinterLang(DuurMinuten, Zichtjaar, Scenario)
'        p = GEVCDF(locpar, scalepar, shapepar, Volume)
'    Else
'        'voor de korte duren zijn geen klimaatscenario's gepubliceerd. Advies is om hier de verhouding klimaat_2015/huidig_2015 toe te passen als multiplier
'        'bereken eerst de kansverdelingsparameters voor het huidige klimaat (2014)
'        Volume = Volume / 1.02
'        dispcoef = GEVDispCoefBasisstatistiek2018WinterKort(DuurMinuten)
'        locpar = GEVLocParBasisstatistiek2018WinterKort(DuurMinuten)
'        scalepar = dispcoef * locpar
'        shapepar = GEVShapeParBasisStatistiek2018WinterKort(DuurMinuten)
'        p = GEVCDF(locpar, scalepar, shapepar, Volume)
'
'        If Zichtjaar = 2014 Then
'            Multiplier = 1
'        Else
'            Multiplier = STOWA2015_WINTER_V(DuurMinuten, 1 / -Math.Log(p), Zichtjaar, Scenario, Corner)
'        End If
'
'
'
'    End If
'
'    If ReturnT Then
'        STOWA2015_2018_WINTER_T = 1 / -Math.Log(p)
'    ElseIf ReturnDispCoef Then
'        STOWA2015_2018_WINTER_T = dispcoef
'    ElseIf ReturnLocPar Then
'        STOWA2015_2018_WINTER_T = locpar
'    ElseIf ReturnScalePar Then
'        STOWA2015_2018_WINTER_T = scalepar
'    ElseIf ReturnShapePar Then
'        STOWA2015_2018_WINTER_T = shapepar
'    End If
    
End Function


Public Function STOWA2015_2018_NDJF_V(ByVal DuurMinuten As Integer, ByVal T As Double, ByVal Zichtjaar As Integer, ByVal Scenario As String, ByVal Corner As String) As Double
    'deze functie berekent de herhalingstijd voor winter-neerslagstatistiek conform STOWA, 2015/2018 met gegeven Herhalingstijd en duur in minuten
    'in de tweede iteratie gebruiken we de herhalingstijd die werd berekend weer als input
    Dim locpar As Double, scalepar As Double, shapepar As Double, dispcoef As Double
    Dim P As Double
    Dim Multiplier As Double
    P = Exp(-1 / T)
        
    If DuurMinuten > 720 Then
        'voor lange duren is de statistiek van alle klimaatscenario's beschikbaar
        dispcoef = GEVDispCoefBasisstatistiek2015WinterLang(DuurMinuten, Zichtjaar, Scenario, Corner)
        locpar = GEVLocParBasisstatistiek2015WinterLang(DuurMinuten, Zichtjaar, Scenario, Corner)
        scalepar = dispcoef * locpar
        shapepar = GEVShapeParBasisstatistiek2015WinterLang(DuurMinuten, Zichtjaar, Scenario)
        STOWA2015_2018_NDJF_V = GEVINVERSE(locpar, scalepar, shapepar, P)
    Else
        'Voor korte duren is alleen de statistiek van huidig klimaat gepubliceerd.
        'Volgens auteur Rudolf Versteeg mag de verhouding klimaat/huidig voor het winterseizoen uit de publicatie van 2015 worden toegepast om toch het klimaateffect te berekenen
        dispcoef = GEVDispCoefBasisstatistiek2018WinterKort(DuurMinuten)
        locpar = GEVLocParBasisstatistiek2018WinterKort(DuurMinuten)
        scalepar = dispcoef * locpar
        shapepar = GEVShapeParBasisStatistiek2018WinterKort(DuurMinuten)
        
        'pas nu de klimaatverandering toe, indien vereist
        If Zichtjaar = 2014 Then
            Multiplier = 1
        Else
            Multiplier = STOWA2015_WINTER_V(DuurMinuten, T, Zichtjaar, Scenario, Corner) / STOWA2015_WINTER_V(DuurMinuten, T, 2014, "", "")
        End If
                
        STOWA2015_2018_NDJF_V = GEVINVERSE(locpar, scalepar, shapepar, P) * 1.02 * Multiplier
        
    End If
End Function

Public Function STOWA2015_WINTER_V(ByVal DuurMinuten As Integer, ByVal T As Double, ByVal Zichtjaar As Integer, ByVal Scenario As String, ByVal Corner As String) As Double
    'deze functie berekent de herhalingstijd voor winter-neerslagstatistiek conform STOWA, 2015/2018 met gegeven Herhalingstijd en duur in minuten
    'in de tweede iteratie gebruiken we de herhalingstijd die werd berekend weer als input
    Dim locpar As Double, scalepar As Double, shapepar As Double, dispcoef As Double
    Dim P As Double
    P = Exp(-1 / T)
    
    If DuurMinuten < 120 Then DuurMinuten = 120 'zoals overeengekomen met Rudolf Versteeg (auteur rapport STOWA) ten behoeve van klimaateffect bij korte duren
    dispcoef = GEVDispCoefBasisstatistiek2015WinterLang(DuurMinuten, Zichtjaar, Scenario, Corner)
    locpar = GEVLocParBasisstatistiek2015WinterLang(DuurMinuten, Zichtjaar, Scenario, Corner)
    scalepar = dispcoef * locpar
    shapepar = GEVShapeParBasisstatistiek2015WinterLang(DuurMinuten, Zichtjaar, Scenario)
    STOWA2015_WINTER_V = GEVINVERSE(locpar, scalepar, shapepar, P)
End Function

Public Function GEVDispCoefBasisstatistiek2015WinterLang(DuurMinuten, Zichtjaar As Integer, Scenario As String, Corner As String) As Double
    'deze functie berekent de dispersiecoefficient voor de GEV-kansverdeling voor lange duur (>= 2 uur) conform STOWA 2015
    If Zichtjaar = 2014 Then
        GEVDispCoefBasisstatistiek2015WinterLang = 0.234
    ElseIf Zichtjaar = 2030 Then
        If Corner = "lower" Then
                        GEVDispCoefBasisstatistiek2015WinterLang = 0.23
        ElseIf Corner = "center" Then
                        GEVDispCoefBasisstatistiek2015WinterLang = 0.233
        ElseIf Corner = "upper" Then
                        GEVDispCoefBasisstatistiek2015WinterLang = 0.236
        End If
    ElseIf Zichtjaar = 2050 Then
        If Scenario = "GL" Then
            If Corner = "lower" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.234
            ElseIf Corner = "center" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.236
            ElseIf Corner = "upper" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.239
            End If
        ElseIf Scenario = "GH" Then
            If Corner = "lower" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.232
            ElseIf Corner = "center" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.235
            ElseIf Corner = "upper" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.237
            End If
        ElseIf Scenario = "WL" Then
            If Corner = "lower" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.235
            ElseIf Corner = "center" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.241
            ElseIf Corner = "upper" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.247
            End If
        ElseIf Scenario = "WH" Then
            If Corner = "lower" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.227
            ElseIf Corner = "center" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.233
            ElseIf Corner = "upper" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.239
            End If
        End If
    ElseIf Zichtjaar = 2085 Then
        If Scenario = "GL" Then
            If Corner = "lower" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.229
            ElseIf Corner = "center" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.233
            ElseIf Corner = "upper" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.237
            End If
        ElseIf Scenario = "GH" Then
            If Corner = "lower" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.228
            ElseIf Corner = "center" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.232
            ElseIf Corner = "upper" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.236
            End If
        ElseIf Scenario = "WL" Then
            If Corner = "lower" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.236
            ElseIf Corner = "center" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.246
            ElseIf Corner = "upper" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.255
            End If
        ElseIf Scenario = "WH" Then
            If Corner = "lower" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.226
            ElseIf Corner = "center" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.236
            ElseIf Corner = "upper" Then
                                GEVDispCoefBasisstatistiek2015WinterLang = 0.245
            End If
        End If
    End If
End Function


Public Function GEVLocParBasisstatistiek2015WinterLang(DuurMinuten, Zichtjaar As Integer, Scenario As String, Corner As String) As Double
    'deze functie berekent de locatieparameter voor de GEV-kansverdeling voor lange duur, winterseizoen conform STOWA 2015
    If Zichtjaar = 2014 Then
        'op aanwizjen van Rudolf de extra decimaal toegevoegd in -0.193
        GEVLocParBasisstatistiek2015WinterLang = (0.67 - 0.0426 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.193)
    ElseIf Zichtjaar = 2030 Then
        If Corner = "lower" Then
            GEVLocParBasisstatistiek2015WinterLang = (0.667 - 0.0435 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.197)
        ElseIf Corner = "center" Then
            GEVLocParBasisstatistiek2015WinterLang = (0.665 - 0.043 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.196)
        ElseIf Corner = "upper" Then
            GEVLocParBasisstatistiek2015WinterLang = (0.666 - 0.0425 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.194)
        End If
    ElseIf Zichtjaar = 2050 Then
        If Scenario = "GL" Then
            If Corner = "lower" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.668 - 0.0431 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.196)
            ElseIf Corner = "center" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.668 - 0.0426 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.194)
            ElseIf Corner = "upper" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.667 - 0.0422 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.193)
            End If
        ElseIf Scenario = "GH" Then
            If Corner = "lower" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.661 - 0.0437 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.2)
            ElseIf Corner = "center" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.661 - 0.0432 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.198)
            ElseIf Corner = "upper" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.66 - 0.0426 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.196)
            End If
        ElseIf Scenario = "WL" Then
            If Corner = "lower" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.671 - 0.0421 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.19)
            ElseIf Corner = "center" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.672 - 0.0411 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.186)
            ElseIf Corner = "upper" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.671 - 0.0402 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.183)
            End If
        ElseIf Scenario = "WH" Then
            If Corner = "lower" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.662 - 0.0431 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.196)
            ElseIf Corner = "center" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.66 - 0.0422 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.193)
            ElseIf Corner = "upper" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.661 - 0.0412 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.189)
            End If
        End If
    ElseIf Zichtjaar = 2085 Then
        If Scenario = "GL" Then
            If Corner = "lower" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.667 - 0.0429 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.195)
            ElseIf Corner = "center" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.666 - 0.0423 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.193)
            ElseIf Corner = "upper" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.667 - 0.0416 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.19)
            End If
        ElseIf Scenario = "GH" Then
            If Corner = "lower" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.651 - 0.0437 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.205)
            ElseIf Corner = "center" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.651 - 0.043 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.202)
            ElseIf Corner = "upper" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.65 - 0.0423 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.2)
            End If
        ElseIf Scenario = "WL" Then
            If Corner = "lower" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.675 - 0.0417 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.185)
            ElseIf Corner = "center" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.674 - 0.0403 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.18)
            ElseIf Corner = "upper" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.675 - 0.0389 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.175)
            End If
        ElseIf Scenario = "WH" Then
            If Corner = "lower" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.653 - 0.043 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.198)
            ElseIf Corner = "center" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.651 - 0.0415 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.193)
            ElseIf Corner = "upper" Then
                GEVLocParBasisstatistiek2015WinterLang = (0.647 - 0.0402 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.19)
            End If
        End If
    End If
End Function


Public Function GEVShapeParBasisstatistiek2015WinterLang(DuurMinuten, Zichtjaar As Integer, Scenario As String) As Double
    'deze functie berekent de vormparameter voor de GEV-kansverdeling voor lange duur conform STOWA 2015
    If DuurMinuten / 60 <= 240 Then
        GEVShapeParBasisstatistiek2015WinterLang = -0.09 + 0.017 * DuurMinuten / 60 / 24
    Else
        GEVShapeParBasisstatistiek2015WinterLang = -0.09 + 0.683 * Math.Log(DuurMinuten / 60 / 24)
    End If
End Function

Public Function GEVDispCoefBasisstatistiek2018WinterKort(DuurMinuten) As Double
    If DuurMinuten <= 91 Then
        GEVDispCoefBasisstatistiek2018WinterKort = 0.41692 - 0.07583 * Log10(CDbl(DuurMinuten))
    Else
        GEVDispCoefBasisstatistiek2018WinterKort = 0.2684
    End If
End Function

Public Function GEVLocParBasisstatistiek2018WinterKort(DuurMinuten) As Double
    GEVLocParBasisstatistiek2018WinterKort = 4.883 - 5.587 * Log10(CDbl(DuurMinuten)) + 3.526 * Log10(CDbl(DuurMinuten)) ^ 2
End Function

Public Function GEVShapeParBasisStatistiek2018WinterKort(DuurMinuten) As Double
    GEVShapeParBasisStatistiek2018WinterKort = -0.294 + 0.1474 * Log10(CDbl(DuurMinuten)) - 0.0192 * Log10(CDbl(DuurMinuten)) ^ 2
End Function

Public Function GEVShapeParBasisStatistiek2018ZomerKort(DuurMinuten) As Double
    GEVShapeParBasisStatistiek2018WinterKort = -0.0336 - 0.264 * Log10(CDbl(DuurMinuten)) + 0.0636 * Log10(CDbl(DuurMinuten)) ^ 2
End Function


Public Function STOWA2019_NDJF_T(ByVal DuurMinuten As Integer, ByVal Volume As Double, ByVal Zichtjaar As Integer, ByVal Scenario As String, ByVal Corner As String) As Double
    'deze functie berekent het winterneerslagvolume conform STOWA, 2019 met gegeven duur in minuten en volume in mm
    'we berekenen hem in iteratief door gebruik te maken van de functie STOWA_JAARROND_V
    Dim T_estimate As Double, P As Double
    Dim V_estimate As Double
    
    'Initialiseer de herhalingstijd op basis van de oude statistiek. We weten dat deze herhalingstijd een overschatting geeft
    T_estimate = STOWA2015_2018_WINTER_T(DuurMinuten, Zichtjaar, Scenario, Corner, Volume)
    
    Dim Done As Boolean
    Dim iIter As Integer
    Done = False
    iIter = 0
    
    While Not Done
        'nu gaan we de geschatte herhalingstijd invoeren in de functie met actuele statistiek
        V_estimate = STOWA2019_NDJF_V(DuurMinuten, T_estimate, Zichtjaar, Scenario, Corner)
        iIter = iIter + 1
        If iIter = 1000 Then Done = True
        If Math.Abs(V_estimate - Volume) < 0.001 Then Done = True
        T_estimate = T_estimate * Volume / V_estimate               'pas de geschatte herhalingstijd aan naar rato van de afwijking tussen geschat en opgegeven volume
    Wend
    
    STOWA2019_NDJF_T = T_estimate

End Function


Public Function STOWA2024_NDJF_V(DuurMinuten As Integer, T As Double, Zichtjaar As Integer, Scenario As String, Corner As String) As Double
    'deze functie berekent de herhalingstijd voor winter-neerslagstatistiek conform STOWA, 2024 met gegeven Herhalingstijd en duur in minuten
    'in de STOWA2024-scenario's zijn géén aanpassingen gedaan aan de kansverdelingsparameters.
    'ook is het scenario 'Huidig' ongemoeid gelaten. De volumes voor de diverse zichtjaren komen tot stand via een multiplier die en functie is van de verwachte temperatuursstijging
    'voor huidig klimaat zijn de statistieken voor 2024 identiek aan die uit 2019
    'let op: voor 'Huidig' hanteren we intern nog altijd het jaartal 2014.
    STOWA2024_NDJF_V = STOWA2019_NDJF_V(DuurMinuten, T, 2014, "", "")
    
    Dim VeranderGetal As Double
    VeranderGetal = getVeranderGetal(Zichtjaar, Scenario, "winter", DuurMinuten / 60)
    
    'het volume is nu eenvoudigweg het verandergetal als multiplier op het volume van huidig 2019.
    STOWA2024_NDJF_V = VeranderGetal * STOWA2024_NDJF_V

End Function


Public Function STOWA2024_NDJF_T(ByVal DuurMinuten As Integer, ByVal Volume As Double, ByVal Zichtjaar As Integer, ByVal Scenario As String, ByVal Corner As String) As Double
    'in de 2024-scenario's zijn de zichtjaren eenvoudigweg multipliers op de volumes van 2019_HUIDIG.
    'scenario Huidig is in de 2024-scenario's identiek aan die van 2019
    'daarom kunnen we vrij eenvoudig terugrekenen als we eerst ons volume terugschalen naar scenario Huidig
    Dim VeranderGetal As Double
    VeranderGetal = getVeranderGetal(Zichtjaar, Scenario, "winter", DuurMinuten / 60)
    
    'corrigeer eerst het volume door te delen door het verandergetal
    Volume = Volume / VeranderGetal
    
    'nu we het volume hebben teruggeschaald naar de equivalent voor scenario 'Huidig'
    'kunnen we eenvoudigweg de functie STOWA2019_NDJF_T aanroepen, met scenario 'Huidig' als referentie
    STOWA2024_NDJF_T = STOWA2019_NDJF_T(DuurMinuten, Volume, 2014, "", "")
    
End Function


Public Function STOWA2024_JAARROND_V(DuurMinuten As Integer, T As Double, Zichtjaar As Integer, Scenario As String, Corner As String) As Double
    'deze functie berekent de herhalingstijd voor winter-neerslagstatistiek conform STOWA, 2024 met gegeven Herhalingstijd en duur in minuten
    'in de STOWA2024-scenario's zijn géén aanpassingen gedaan aan de kansverdelingsparameters.
    'ook is het scenario 'Huidig' ongemoeid gelaten. De volumes voor de diverse zichtjaren komen tot stand via een multiplier die en functie is van de verwachte temperatuursstijging
    'voor huidig klimaat zijn de statistieken voor 2024 identiek aan die uit 2019
    'let op: voor 'Huidig' hanteren we intern nog altijd het jaartal 2014.
    STOWA2024_JAARROND_V = STOWA2019_JAARROND_V(DuurMinuten, T, 2014, "", "")
    
    Dim VeranderGetal As Double
    VeranderGetal = getVeranderGetal(Zichtjaar, Scenario, "jaarrond", DuurMinuten / 60)
    
    'het volume is nu eenvoudigweg het verandergetal als multiplier op het volume van huidig 2019.
    STOWA2024_JAARROND_V = VeranderGetal * STOWA2024_JAARROND_V

End Function


Public Function STOWA2024_JAARROND_T(ByVal DuurMinuten As Integer, ByVal Volume As Double, ByVal Zichtjaar As Integer, ByVal Scenario As String, ByVal Corner As String) As Double
    'in de 2024-scenario's zijn de zichtjaren eenvoudigweg multipliers op de volumes van 2019_HUIDIG.
    'scenario Huidig is in de 2024-scenario's identiek aan die van 2019
    'daarom kunnen we vrij eenvoudig terugrekenen als we eerst ons volume terugschalen naar scenario Huidig
    Dim VeranderGetal As Double
    VeranderGetal = getVeranderGetal(Zichtjaar, Scenario, "jaarrond", DuurMinuten / 60)
    
    'corrigeer eerst het volume door te delen door het verandergetal
    Volume = Volume / VeranderGetal
    
    'nu we het volume hebben teruggeschaald naar de equivalent voor scenario 'Huidig'
    'kunnen we eenvoudigweg de functie STOWA2019_NDJF_T aanroepen, met scenario 'Huidig' als referentie
    STOWA2024_JAARROND_T = STOWA2019_JAARROND_T(DuurMinuten, Volume, 2014, "", "")
    
End Function


Public Function STOWA2019_JAARROND_T(ByVal DuurMinuten As Integer, ByVal Volume As Double, ByVal Zichtjaar As Integer, ByVal Scenario As String, ByVal Corner As String) As Double
    'deze functie berekent het jaarrond neerslagvolume conform STOWA, 2019 met gegeven duur in minuten en volume in mm
    'we berekenen hem in iteratief door gebruik te maken van de functie STOWA_JAARROND_V
    Dim T_estimate As Double, P As Double
    Dim V_estimate As Double
    
    'Initialiseer de herhalingstijd
    T_estimate = 1
    
    Dim Done As Boolean
    Dim iIter As Integer
    Done = False
    iIter = 0
       
    
    While Not Done
        'nu gaan we de geschatte herhalingstijd invoeren in de functie met actuele statistiek
        V_estimate = STOWA2019_JAARROND_V(DuurMinuten, T_estimate, Zichtjaar, Scenario, Corner)
        iIter = iIter + 1
        If iIter = 1000 Then Done = True
        If Math.Abs(V_estimate - Volume) < 0.001 Then Done = True
        T_estimate = T_estimate * Volume / V_estimate
    Wend
    
    STOWA2019_JAARROND_T = T_estimate
    
End Function

Function getVeranderGetal(Zichtjaar As Integer, Scenario As String, Seizoen As String, DuurUren As Double) As Double
    'deze functie berekent het verandergetal voor de klimaatscenario's 2024 als functie van zichtjaar, scenario en duur
    'op zijn beurt roept deze functie weer de functie VeranderGetalFunctie aan, waarin hij de verwachtte temperatuursstijging meegeeft, die afhangt van het zichtjaar en scenario
    '0.6 graden (2033L)
    '0.8 graden (2050L)
    '1.1 graden (2050M)
    '1.5 graden (2050H)
    '0.8 graden (2100L)
    '1.9 graden (2100M)
    '4.0 graden (2100H)
    '0.8 graden (2150L)
    '2.1 graden (2150M)
    '5.5 graden (2150H)
    If Zichtjaar = 2014 Then
                'geen verandering
        getVeranderGetal = 1
    ElseIf Zichtjaar = 2033 Then
        If Scenario = "L" Then
            If Seizoen = "winter" Then
                                getVeranderGetal = VeranderGetalFunctieWinter(0.6, DuurUren)
                        Else
                                getVeranderGetal = verandergetalfunctieJaarrond(0.6, DuurUren)
                        End If
        ElseIf Scenario = "M" Then
                        'is een niet-bestaand scenario
                        getVeranderGetal = 0
        ElseIf Scenario = "H" Then
                        'is een niet-bestaand scenario
                        getVeranderGetal = 0
                Else
                        'is een niet-bestaand scenario
                        getVeranderGetal = 0
        End If
    ElseIf Zichtjaar = 2050 Then
        If Scenario = "L" Then
            If Seizoen = "winter" Then
                getVeranderGetal = VeranderGetalFunctieWinter(0.8, DuurUren)
            Else
                getVeranderGetal = verandergetalfunctieJaarrond(0.8, DuurUren)
            End If
        ElseIf Scenario = "M" Then
            If Seizoen = "winter" Then
                getVeranderGetal = VeranderGetalFunctieWinter(1.1, DuurUren)
            Else
                getVeranderGetal = verandergetalfunctieJaarrond(1.1, DuurUren)
            End If
        ElseIf Scenario = "H" Then
            If Seizoen = "winter" Then
                getVeranderGetal = VeranderGetalFunctieWinter(1.5, DuurUren)
            Else
                getVeranderGetal = verandergetalfunctieJaarrond(1.5, DuurUren)
            End If
        Else
            'is een niet-bestaand scenario
            getVeranderGetal = 0
        End If
    ElseIf Zichtjaar = 2100 Then
        If Scenario = "L" Then
            If Seizoen = "winter" Then
                getVeranderGetal = VeranderGetalFunctieWinter(0.8, DuurUren)
            Else
                getVeranderGetal = verandergetalfunctieJaarrond(0.8, DuurUren)
            End If
        ElseIf Scenario = "M" Then
            If Seizoen = "winter" Then
                getVeranderGetal = VeranderGetalFunctieWinter(1.9, DuurUren)
            Else
                getVeranderGetal = verandergetalfunctieJaarrond(1.9, DuurUren)
            End If
        ElseIf Scenario = "H" Then
            If Seizoen = "winter" Then
                getVeranderGetal = VeranderGetalFunctieWinter(4, DuurUren)
            Else
                getVeranderGetal = verandergetalfunctieJaarrond(4, DuurUren)
            End If
        Else
            'is een niet-bestaand scenario
            getVeranderGetal = 0
        End If
    ElseIf Zichtjaar = 2150 Then
        If Scenario = "L" Then
            'is identiek aan 2050L en 2100L
            If Seizoen = "winter" Then
                getVeranderGetal = VeranderGetalFunctieWinter(0.8, DuurUren)
            Else
                getVeranderGetal = verandergetalfunctieJaarrond(0.8, DuurUren)
            End If
        ElseIf Scenario = "M" Then
            If Seizoen = "winter" Then
                getVeranderGetal = VeranderGetalFunctieWinter(2.1, DuurUren)
            Else
                getVeranderGetal = verandergetalfunctieJaarrond(2.1, DuurUren)
            End If
        ElseIf Scenario = "H" Then
            If Seizoen = "winter" Then
                getVeranderGetal = VeranderGetalFunctieWinter(5.5, DuurUren)
            Else
                getVeranderGetal = verandergetalfunctieJaarrond(5.5, DuurUren)
            End If
        Else
            'is een niet-bestaand scenario
            getVeranderGetal = 0
        End If
    End If
End Function

Function VeranderGetalFunctieWinter(Ts As Double, D As Double, Optional T As Double = 1) As Double
    Dim v As Double
    ' Deze functie berekent het verandergetal wat nodig is voor de klimaatscenario's van 2024
    ' Ts: temperatuurstijging in graden Celsius
    ' D:  duur in uren
    ' T:  terugkeertijd (irrelevante parameter)
    
    If D < 1 / 6 Then
        err.Raise Number:=vbObjectError + 513, _
                  Description:="Gekozen duur " & D & " valt buiten domein (10 minuten t/m 240 uur)"
    ElseIf D <= 24 Then
        v = 1.244
    ElseIf D < 120 Then
        v = Poly(D)
    ElseIf D <= 240 Then
        v = 1.181
    ElseIf D > 240 Then
        err.Raise Number:=vbObjectError + 514, _
                  Description:="Gekozen duur " & D & " valt buiten domein: 10 minuten t/m 10 dagen (240 uur)"
    End If
    
    'vervangen in v1.08
    'VeranderGetalFunctieWinter = 1 + (v - 1) * Ts / 4 ' de factor v is afgeleid voor 4 graden temperatuurstijging
    
    ' de factor v is afgeleid voor 4 graden temperatuurstijging t.o.v. 2005,
    ' maar in 2023 hebben we al 0.4 graden gehad (0.6 graden in 2033, ~0.4 in 2023)
    VeranderGetalFunctieWinter = 1 + (v - 1) * (Ts - 0.4) / (4 - 0.4)

    
End Function

Function verandergetalfunctieJaarrond(Ts As Double, D As Double, Optional T As Double = 1) As Double
    ' Ts: temperatuurstijging in graden Celsius
    ' D:  duur in uren
    ' T:  terugkeertijd (irrelevante parameter)
    
    Dim v As Double
    
    If D < 1 / 6 Then
        err.Raise Number:=vbObjectError + 513, _
                  Description:="Gekozen duur " & D & " valt buiten domein (10 minuten t/m 240 uur)"
    ElseIf D <= 24 Then
        v = 1.234
    ElseIf D < 120 Then
        v = LogPoly(D) ' Assuming LogPoly is a function you have elsewhere
    ElseIf D <= 240 Then
        v = 1.109
    ElseIf D > 240 Then
        err.Raise Number:=vbObjectError + 514, _
                  Description:="Gekozen duur " & D & " valt buiten domein: 10 minuten t/m 10 dagen (240 uur)"
    End If
    
    'vervangen in v1.08
    'verandergetalfunctieJaarrond = 1 + (v - 1) * Ts / 4 ' de factor v is afgeleid voor 4 graden temperatuurstijging

    verandergetalfunctieJaarrond = 1 + (v - 1) * (Ts - 0.4) / (4 - 0.4)
    ' de factor v is afgeleid voor 4 graden temperatuurstijging t.o.v. 2005,
    ' maar in 2023 hebben we al 0.4 graden gehad (0.6 graden in 2033, ~0.4 in 2023)


End Function

Function Poly(D As Double) As Double
    'dit is een door HKV gefitte polynoom aan de multipliers voor verschillende duren, voor het winterseizoen; publicaties 2024
    ' Calculate the polynomial value based on D
    Poly = 0.000005952 * D ^ 2 - 0.001515 * D + 1.277
End Function

Function LogPoly(D As Double) As Double
    'dit is een door HKV gefitter polynoom aan de multipliers voor verschillende duren, voor jaarrond-neerslagstatistiek; publicaties 2024
    ' Logarithmic polynomial calculation as specified
    Dim logD As Double
    logD = Log(D)
    LogPoly = 0.009143 * logD ^ 2 - 0.1508 * logD + 1.621
End Function

Public Function STOWA2019_NDJF_V(DuurMinuten As Integer, T As Double, Zichtjaar As Integer, Scenario As String, Corner As String) As Double
    'deze functie berekent de herhalingstijd voor winter-neerslagstatistiek conform STOWA, 2019 met gegeven Herhalingstijd en duur in minuten
    'in de tweede iteratie gebruiken we de herhalingstijd die werd berekend weer als input
    Dim P As Double
    Dim locpar As Double, scalepar As Double, shapepar As Double, dispcoef As Double
    Dim Volume As Double
    Dim KorteDuurMultiplier1 As Double, KorteDuurMultiplier2 As Double, LangeDuurMultiplier As Double, Multiplier As Double
    
    P = Exp(-1 / T)
    If DuurMinuten > 720 Then
        dispcoef = GEVDispCoefBasisstatistiek2019LangeDuurWinter(DuurMinuten)
        locpar = GEVLocparBasisstatistiek2019LangeDuurWinter(DuurMinuten)
        scalepar = dispcoef * locpar
        shapepar = GEVShapeParBasisstatistiek2019LangeDuurWinter(DuurMinuten)
        Volume = GEVINVERSE(locpar, scalepar, shapepar, P)
    Else
        dispcoef = GEVDispCoefBasisstatistiek2019KorteDuurWinter(DuurMinuten)
        locpar = GEVLocparBasisstatistiek2019KorteDuurWinter(DuurMinuten)
        scalepar = dispcoef * locpar
        shapepar = GEVShapeParBasisstatistiek2019KorteDuurWinter(DuurMinuten)
        Volume = GEVINVERSE(locpar, scalepar, shapepar, P)
    End If
    
    
    'bepaal nu de aanpassingen als gevolg van het onderhavige klimaat
    'let op: voor de winterstatistiek maken we géén onderscheid tussen korte en lange duur bij het berekenen van de klimaatscenario's.
    'informatie hierover ontbreekt in het rapport van STOWA en navraag bij de auteur (Rudolf Versteeg) leert dat voor korte duren
    'gewoon de verhouding STOWA2015_2018_Winter_Klimaat / STOWA2015_2019_Winter_Huidig kan worden toegepast
    'e-mail rudolf versteeg van 30-3-2020 13:47 aan siebe@hydroconsult.nl:
    '"Voor de klimaatscenario’s geldt voor de winter gewoon hetzelfde als voor de lange duren bij jaarstatistiek. Voor de korte duren zijn immers geen verschillende klimaatveranderingsindicatoren voor korte en lange duren gegeven, simpelweg omdat niet wordt verwacht dat die zich verschillend ontwikkelen in de toekomst. Je moet hier dus gewoon aansluiten bij de procentuele veranderingen die uit de statistieken uit 2015 volgen. Dus dezelfde procedure als de procedure voor lange duren van de jaarstatistiek."
    If Zichtjaar <> 2014 Then
        Multiplier = STOWA2019_MULTIPLIER_WINTER(DuurMinuten, T, Zichtjaar, Scenario, Corner)
    Else
        Multiplier = 1
    End If
    Volume = Volume * Multiplier
        
    STOWA2019_NDJF_V = Volume
    
End Function


Public Function GEVDispCoefBasisstatistiek2019LangeDuurWinter(DuurMinuten As Integer) As Double
    GEVDispCoefBasisstatistiek2019LangeDuurWinter = 0.234
End Function


Public Function GEVLocparBasisstatistiek2019LangeDuurWinter(DuurMinuten As Integer) As Double
        GEVLocparBasisstatistiek2019LangeDuurWinter = (0.67 - 0.0426 * Math.Log(DuurMinuten / 60)) ^ (-1 / 0.193)
End Function


Public Function GEVShapeParBasisstatistiek2019LangeDuurWinter(DuurMinuten As Integer) As Double
        GEVShapeParBasisstatistiek2019LangeDuurWinter = -0.09 + 0.017 * DuurMinuten / 60 / 24
End Function


Public Function GEVDispCoefBasisstatistiek2019KorteDuurWinter(DuurMinuten As Integer) As Double
    'deze functie berekent de dispersiecoefficient epsylon voor de GEV-kansverdeling voor korte duur Winter (10 minuten t/m 12 uur) volgens STOWA 2019, deelrapport 1 p12
    'let op: dit is NIET de schaalparameter uit de GEV-verdeling. Daarvoor moet eerst nog met de locatiepar (zeta) worden vermenigvuldigd
    If DuurMinuten <= 91 Then
        GEVDispCoefBasisstatistiek2019KorteDuurWinter = 0.41692 - 0.07583 * Log10(CDbl(DuurMinuten))
    Else
        GEVDispCoefBasisstatistiek2019KorteDuurWinter = 0.2684
    End If
End Function

Public Function GEVLocparBasisstatistiek2019KorteDuurWinter(DuurMinuten As Integer) As Double
    'deze functie berekent de locatieparameter zeta voor de GEV-kansverdeling voor korte duur (10 minuten t/m 12 uur) volgens STOWA 2019, deelrapport 1 p12
    GEVLocparBasisstatistiek2019KorteDuurWinter = 1.07 * 1.02 * (4.883 - 5.587 * Log10(CDbl(DuurMinuten)) + 3.526 * (Log10(CDbl(DuurMinuten))) ^ 2)
End Function


Public Function GEVShapeParBasisstatistiek2019KorteDuurWinter(DuurMinuten As Integer) As Double
    GEVShapeParBasisstatistiek2019KorteDuurWinter = -0.294 + 0.1474 * Log10(CDbl(DuurMinuten)) - 0.0192 * (Log10(CDbl(DuurMinuten))) ^ 2
End Function

Public Function GLOScaleParBasisstatistiek2019KorteDuur(DuurMinuten As Integer) As Double
    'berekent de schaalparameter voor korte duren
    GLOScaleParBasisstatistiek2019KorteDuur = GLODispCoefBasisstatistiek2019KorteDuur(DuurMinuten) * GLOLocparBasisstatistiek2019KorteDuur(DuurMinuten)
End Function

Public Function GEVScaleParBasisstatistiek2019KorteDuurWinter(DuurMinuten As Integer) As Double
    'berekent de schaalparameter voor korte duren voor wintersituaties (NDJF)
    GEVScaleParBasisstatistiek2019KorteDuurWinter = GLODispCoefBasisstatistiek2019KorteDuurWinter(DuurMinuten) * GLOLocparBasisstatistiek2019KorteDuurWinter(DuurMinuten)
End Function

Public Function GLOShapeParBasisstatistiek2019KorteDuur(DuurMinuten As Integer, T_estimate As Double) As Double
    'deze functie berekent de vormparameter voor de GLO-kansverdeling voor korte duur (10 minuten t/m 12 uur) volgens STOWA 2019, deelrapport 1 p12
    'LET OP: De T_estimate <= 120 luidt in het rapport van STOWA (2019) op p12 <= 1. Dit is echter fout. In par 2.3 staat 120 jaar als grenswaarde
    If (DuurMinuten <= 90 Or (DuurMinuten <= 720 And T_estimate <= 120)) Then
        GLOShapeParBasisstatistiek2019KorteDuur = -0.0336 - 0.264 * Log10(CDbl(DuurMinuten)) + 0.0636 * (Log10(CDbl(DuurMinuten))) ^ 2
    Else
        GLOShapeParBasisstatistiek2019KorteDuur = -0.31 - 0.0544 * Log10(CDbl(DuurMinuten)) + 0.0288 * (Log10(CDbl(DuurMinuten))) ^ 2
    End If
End Function


Public Function GEVLocparBasisstatistiek2019LangeDuur(DuurMinuten As Integer) As Double
    'deze functie berekent de locatieparameter zeta voor de GLO-kansverdeling voor korte duur (10 minuten t/m 12 uur) volgens STOWA 2019, deelrapport 1 p13
    '#jaarrondstatistiek voor 12 uur tot 10 dagen. Zie p13 deelrapport I
    'Mu <- 1.02 * (0.239 -0.0250 * log(x))^(1/-0.512) (uit het R-script van Overeem(
    GEVLocparBasisstatistiek2019LangeDuur = 1.02 * (0.239 - 0.025 * Math.Log(DuurMinuten / 60)) ^ (-1 / 0.512)
End Function


Public Function GEVLocParBasisstatistiek2015(DuurMinuten, Zichtjaar As Integer, Scenario As String, Corner As String) As Double
    'deze functie berekent de locatieparameter voor de GEV-kansverdeling voor lange duur (>= 20 uur) conform STOWA 2015
    If Zichtjaar = 2014 Then
        GEVLocParBasisstatistiek2015 = (0.239 - 0.025 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.512)
    ElseIf Zichtjaar = 2030 Then
        If Corner = "lower" Then
            GEVLocParBasisstatistiek2015 = (0.246 - 0.0257 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.503)
        ElseIf Corner = "center" Then
            GEVLocParBasisstatistiek2015 = (0.24 - 0.025 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.506)
        ElseIf Corner = "upper" Then
            GEVLocParBasisstatistiek2015 = (0.235 - 0.0243 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.509)
        Else
            GEVLocParBasisstatistiek2015 = -999
        End If
    ElseIf Zichtjaar = 2050 Then
        If Scenario = "GL" Then
            If Corner = "lower" Then
                GEVLocParBasisstatistiek2015 = (0.247 - 0.0258 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.501)
            ElseIf Corner = "center" Then
                GEVLocParBasisstatistiek2015 = (0.241 - 0.025 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.504)
            ElseIf Corner = "upper" Then
                GEVLocParBasisstatistiek2015 = (0.236 - 0.0243 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.506)
            Else
                GEVLocParBasisstatistiek2015 = -999
            End If
        ElseIf Scenario = "GH" Then
            If Corner = "lower" Then
                GEVLocParBasisstatistiek2015 = (0.269 - 0.0272 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.474)
            ElseIf Corner = "center" Then
                GEVLocParBasisstatistiek2015 = (0.26 - 0.0263 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.479)
            ElseIf Corner = "upper" Then
                GEVLocParBasisstatistiek2015 = (0.252 - 0.0254 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.483)
            Else
                GEVLocParBasisstatistiek2015 = -999
            End If
        ElseIf Scenario = "WL" Then
            If Corner = "lower" Then
                GEVLocParBasisstatistiek2015 = (0.262 - 0.0266 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.48)
            ElseIf Corner = "center" Then
                GEVLocParBasisstatistiek2015 = (0.249 - 0.0252 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.485)
            ElseIf Corner = "upper" Then
                GEVLocParBasisstatistiek2015 = (0.241 - 0.024 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.486)
            Else
                GEVLocParBasisstatistiek2015 = -999
            End If
        ElseIf Scenario = "WH" Then
            If Corner = "lower" Then
                GEVLocParBasisstatistiek2015 = (0.289 - 0.0287 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.451)
            ElseIf Corner = "center" Then
                GEVLocParBasisstatistiek2015 = (0.276 - 0.0271 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.456)
            ElseIf Corner = "upper" Then
                GEVLocParBasisstatistiek2015 = (0.265 - 0.0257 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.459)
            Else
                GEVLocParBasisstatistiek2015 = -999
            End If
        Else
            GEVLocParBasisstatistiek2015 = -999
        End If
    ElseIf Zichtjaar = 2085 Then
            If Scenario = "GL" Then
            If Corner = "lower" Then
                GEVLocParBasisstatistiek2015 = (0.252 - 0.0261 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.494)
            ElseIf Corner = "center" Then
                GEVLocParBasisstatistiek2015 = (0.243 - 0.025 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.498)
            ElseIf Corner = "upper" Then
                GEVLocParBasisstatistiek2015 = (0.235 - 0.0241 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.501)
            Else
                GEVLocParBasisstatistiek2015 = -999
            End If
        ElseIf Scenario = "GH" Then
            If Corner = "lower" Then
                GEVLocParBasisstatistiek2015 = (0.271 - 0.0274 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.471)
            ElseIf Corner = "center" Then
                GEVLocParBasisstatistiek2015 = (0.26 - 0.0262 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.476)
            ElseIf Corner = "upper" Then
                GEVLocParBasisstatistiek2015 = (0.25 - 0.0251 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.481)
            Else
                GEVLocParBasisstatistiek2015 = -999
            End If
        ElseIf Scenario = "WL" Then
            If Corner = "lower" Then
                GEVLocParBasisstatistiek2015 = (0.272 - 0.0272 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.464)
            ElseIf Corner = "center" Then
                GEVLocParBasisstatistiek2015 = (0.248 - 0.0244 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.475)
            ElseIf Corner = "upper" Then
                GEVLocParBasisstatistiek2015 = (0.23 - 0.0223 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.482)
            Else
                GEVLocParBasisstatistiek2015 = -999
            End If
        ElseIf Scenario = "WH" Then
            If Corner = "lower" Then
                GEVLocParBasisstatistiek2015 = (0.286 - 0.0284 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.448)
            ElseIf Corner = "center" Then
                GEVLocParBasisstatistiek2015 = (0.262 - 0.0256 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.458)
            ElseIf Corner = "upper" Then
                GEVLocParBasisstatistiek2015 = (0.247 - 0.0236 * Math.Log(DuurMinuten / 60)) ^ (1 / -0.461)
            Else
                GEVLocParBasisstatistiek2015 = -999
            End If
        Else
            GEVLocParBasisstatistiek2015 = -999
        End If
    Else
        GEVLocParBasisstatistiek2015 = -999
    End If
End Function

Public Function GLODispCoefBasisstatistiek2018(DuurMinuten) As Double
    If DuurMinuten <= 104 Then
        GLODispCoefBasisstatistiek2018 = 0.04704 + 0.1978 * Log10(CDbl(DuurMinuten)) - 0.05729 * Log10(CDbl(DuurMinuten)) ^ 2
    Else
        GLODispCoefBasisstatistiek2018 = 0.2801 - 0.0333 * Log10(CDbl(DuurMinuten))
    End If
End Function

Public Function GLOLocParBasisstatistiek2018(DuurMinuten) As Double
    GLOLocParBasisstatistiek2018 = 7.339 + 0.848 * Log10(CDbl(DuurMinuten)) + 2.844 * Log10(CDbl(DuurMinuten)) ^ 2
End Function

Public Function GLOShapeParBasisStatistiek2018(DuurMinuten) As Double
    GLOShapeParBasisStatistiek2018 = -0.0336 - 0.264 * Log10(CDbl(DuurMinuten)) + 0.0636 * Log10(CDbl(DuurMinuten)) ^ 2
End Function

Public Function GEVDispCoefBasisstatistiek2015(DuurMinuten, Zichtjaar As Integer, Scenario As String, Corner As String) As Double
    'deze functie berekent de dispersiecoefficient voor de GEV-kansverdeling voor lange duur (>= 2 uur) conform STOWA 2015
    If Zichtjaar = 2014 Then
        GEVDispCoefBasisstatistiek2015 = 0.378 - 0.0578 * Math.Log(DuurMinuten / 60) + 0.0054 * Math.Log(DuurMinuten / 60) ^ 2
    ElseIf Zichtjaar = 2030 Then
        If Corner = "lower" Then
            GEVDispCoefBasisstatistiek2015 = 0.377 - 0.0565 * Math.Log(DuurMinuten / 60) + 0.005 * Math.Log(DuurMinuten / 60) ^ 2
        ElseIf Corner = "center" Then
            GEVDispCoefBasisstatistiek2015 = 0.384 - 0.0576 * Math.Log(DuurMinuten / 60) + 0.0051 * Math.Log(DuurMinuten / 60) ^ 2
        ElseIf Corner = "upper" Then
            GEVDispCoefBasisstatistiek2015 = 0.39 - 0.0587 * Math.Log(DuurMinuten / 60) + 0.0052 * Math.Log(DuurMinuten / 60) ^ 2
        Else
            GEVDispCoefBasisstatistiek2015 = -999
        End If
    ElseIf Zichtjaar = 2050 Then
        If Scenario = "GL" Then
            If Corner = "lower" Then
                GEVDispCoefBasisstatistiek2015 = 0.377 - 0.0577 * Math.Log(DuurMinuten / 60) + 0.0053 * Math.Log(DuurMinuten / 60) ^ 2
            ElseIf Corner = "center" Then
                GEVDispCoefBasisstatistiek2015 = 0.384 - 0.0589 * Math.Log(DuurMinuten / 60) + 0.0054 * Math.Log(DuurMinuten / 60) ^ 2
            ElseIf Corner = "upper" Then
                GEVDispCoefBasisstatistiek2015 = 0.391 - 0.06 * Math.Log(DuurMinuten / 60) + 0.0055 * Math.Log(DuurMinuten / 60) ^ 2
            End If
        ElseIf Scenario = "GH" Then
            If Corner = "lower" Then
                GEVDispCoefBasisstatistiek2015 = 0.374 - 0.0563 * Math.Log(DuurMinuten / 60) + 0.0051 * Math.Log(DuurMinuten / 60) ^ 2
            ElseIf Corner = "center" Then
                GEVDispCoefBasisstatistiek2015 = 0.382 - 0.0574 * Math.Log(DuurMinuten / 60) + 0.0051 * Math.Log(DuurMinuten / 60) ^ 2
            ElseIf Corner = "upper" Then
                GEVDispCoefBasisstatistiek2015 = 0.39 - 0.0586 * Math.Log(DuurMinuten / 60) + 0.0052 * Math.Log(DuurMinuten / 60) ^ 2
            End If
        ElseIf Scenario = "WL" Then
            If Corner = "lower" Then
                GEVDispCoefBasisstatistiek2015 = 0.375 - 0.0557 * Math.Log(DuurMinuten / 60) + 0.0049 * Math.Log(DuurMinuten / 60) ^ 2
            ElseIf Corner = "center" Then
                GEVDispCoefBasisstatistiek2015 = 0.386 - 0.0572 * Math.Log(DuurMinuten / 60) + 0.005 * Math.Log(DuurMinuten / 60) ^ 2
            ElseIf Corner = "upper" Then
                GEVDispCoefBasisstatistiek2015 = 0.398 - 0.0591 * Math.Log(DuurMinuten / 60) + 0.0052 * Math.Log(DuurMinuten / 60) ^ 2
            End If
        ElseIf Scenario = "WH" Then
            If Corner = "lower" Then
                GEVDispCoefBasisstatistiek2015 = 0.4 - 0.0698 * Math.Log(DuurMinuten / 60) + 0.0064 * Math.Log(DuurMinuten / 60) ^ 2
            ElseIf Corner = "center" Then
                GEVDispCoefBasisstatistiek2015 = 0.416 - 0.0728 * Math.Log(DuurMinuten / 60) + 0.0066 * Math.Log(DuurMinuten / 60) ^ 2
            ElseIf Corner = "upper" Then
                GEVDispCoefBasisstatistiek2015 = 0.432 - 0.0755 * Math.Log(DuurMinuten / 60) + 0.0069 * Math.Log(DuurMinuten / 60) ^ 2
            End If
        End If
    ElseIf Zichtjaar = 2085 Then
        If Scenario = "GL" Then
            If Corner = "lower" Then
                GEVDispCoefBasisstatistiek2015 = 0.377 - 0.0553 * Math.Log(DuurMinuten / 60) + 0.005 * Math.Log(DuurMinuten / 60) ^ 2
            ElseIf Corner = "center" Then
                GEVDispCoefBasisstatistiek2015 = 0.386 - 0.0563 * Math.Log(DuurMinuten / 60) + 0.0051 * Math.Log(DuurMinuten / 60) ^ 2
            ElseIf Corner = "upper" Then
                GEVDispCoefBasisstatistiek2015 = 0.394 - 0.0572 * Math.Log(DuurMinuten / 60) + 0.0052 * Math.Log(DuurMinuten / 60) ^ 2
            End If
        ElseIf Scenario = "GH" Then
            If Corner = "lower" Then
                GEVDispCoefBasisstatistiek2015 = 0.384 - 0.0559 * Math.Log(DuurMinuten / 60) + 0.0046 * Math.Log(DuurMinuten / 60) ^ 2
            ElseIf Corner = "center" Then
                GEVDispCoefBasisstatistiek2015 = 0.395 - 0.0572 * Math.Log(DuurMinuten / 60) + 0.0047 * Math.Log(DuurMinuten / 60) ^ 2
            ElseIf Corner = "upper" Then
                GEVDispCoefBasisstatistiek2015 = 0.405 - 0.0584 * Math.Log(DuurMinuten / 60) + 0.0047 * Math.Log(DuurMinuten / 60) ^ 2
            End If
        ElseIf Scenario = "WL" Then
            If Corner = "lower" Then
                GEVDispCoefBasisstatistiek2015 = 0.374 - 0.0581 * Math.Log(DuurMinuten / 60) + 0.0053 * Math.Log(DuurMinuten / 60) ^ 2
            ElseIf Corner = "center" Then
                GEVDispCoefBasisstatistiek2015 = 0.398 - 0.0612 * Math.Log(DuurMinuten / 60) + 0.0055 * Math.Log(DuurMinuten / 60) ^ 2
            ElseIf Corner = "upper" Then
                GEVDispCoefBasisstatistiek2015 = 0.423 - 0.0657 * Math.Log(DuurMinuten / 60) + 0.0059 * Math.Log(DuurMinuten / 60) ^ 2
            End If
        ElseIf Scenario = "WH" Then
            If Corner = "lower" Then
                GEVDispCoefBasisstatistiek2015 = 0.391 - 0.0654 * Math.Log(DuurMinuten / 60) + 0.0055 * Math.Log(DuurMinuten / 60) ^ 2
            ElseIf Corner = "center" Then
                GEVDispCoefBasisstatistiek2015 = 0.415 - 0.0681 * Math.Log(DuurMinuten / 60) + 0.0056 * Math.Log(DuurMinuten / 60) ^ 2
            ElseIf Corner = "upper" Then
                GEVDispCoefBasisstatistiek2015 = 0.435 - 0.0702 * Math.Log(DuurMinuten / 60) + 0.0056 * Math.Log(DuurMinuten / 60) ^ 2
            End If
        End If
    End If
End Function

Public Function GEVShapeParBasisstatistiek2015(DuurMinuten, Zichtjaar As Integer, Scenario As String) As Double
    'deze functie berekent de vormparameter voor de GEV-kansverdeling voor lange duur (>= 2 uur) conform STOWA 2015
    If DuurMinuten / 60 <= 240 Then
        GEVShapeParBasisstatistiek2015 = (-0.09 + 0.017 * DuurMinuten / 60 / 24)
    Else
        GEVShapeParBasisstatistiek2015 = 0
    End If
End Function

Public Function GEVDispCoefBasisstatistiek2019LangeDuur(DuurMinuten As Integer) As Double
    'deze functie berekent de dispersiecoefficient epsylon voor de GLO-kansverdeling voor lange duur (> 720 uur en <= 14400 min) volgens STOWA 2019, deelrapport 1 p12
    'let op: dit is NIET de schaalparameter uit de GLO=verdeling. Daarvoor moet eerst nog met de locatiepar (zeta) worden vermenigvuldigd
    'Gamma <- 0.478 -0.0681 * log10(x*60) (uit het R-script van Overeem)
    GEVDispCoefBasisstatistiek2019LangeDuur = 0.478 - 0.0681 * Log10(CDbl(DuurMinuten))
End Function

Public Function GEVScaleParBasisstatistiek2019LangeDuur(DuurMinuten As Integer) As Double
    'berekent de schaalparameter voor lange duren, jaarrond
    GEVScaleParBasisstatistiek2019LangeDuur = GEVLocparBasisstatistiek2019LangeDuur(DuurMinuten) * GEVDispCoefBasisstatistiek2019LangeDuur(DuurMinuten)
End Function

Public Function GEVScaleParBasisstatistiek2019LangeDuurWinter(DuurMinuten As Integer) As Double
    'berekent de schaalparameter voor lange duren, winterseizoen (NDJF)
    GEVScaleParBasisstatistiek2019LangeDuurWinter(DuurMinuten) = GEVLocparBasisstatistiek2019LangeDuurWinter(DuurMinuten) * GEVDispCoefBasisstatistiek2019LangeDuurWinter(DuurMinuten)
End Function

Public Function GEVShapeParBasisstatistiek2019LangeDuur(DuurMinuten As Integer) As Double
    'deze functie berekent de dispersiecoefficient epsylon voor de GLO-kansverdeling voor lange duur (> 720 uur en <= 14400 min) volgens STOWA 2019, deelrapport 1 p12
    'let op: dit is NIET de schaalparameter uit de GLO=verdeling. Daarvoor moet eerst nog met de locatiepar (zeta) worden vermenigvuldigd
    'Theta <- 0.118 -0.266 * log10(x*60) + 0.0586 * (log10(x*60))^2 (uit het R-scrip van Overeem)
        GEVShapeParBasisstatistiek2019LangeDuur = 0.118 - 0.266 * Log10(CDbl(DuurMinuten)) + 0.0586 * (Log10(CDbl(DuurMinuten))) ^ 2
End Function


Public Function STOWA2015_2018_JAARROND_T(ByVal DuurMinuten As Integer, ByVal Zichtjaar As Integer, ByVal Scenario As String, ByVal Corner As String, ByVal Volume As Double) As Double
    'deze functie berekent het jaarrond neerslagvolume conform STOWA, 2015/2018 met gegeven duur in minuten en volume in mm
    'we berekenen hem in twee iteraties. In de eerste werken we met een geschatte herhalingstijd < 120 jaar.
    'in de tweede iteratie gebruiken we de herhalingstijd die werd berekend weer als input
    Dim locpar As Double, scalepar As Double, shapepar As Double, dispcoef As Double
    Dim P As Double
    If DuurMinuten > 720 Then  'in STOWA 2015 is lange duur gedefinieerd als >= 2 uur maar in 2018 is korte duur gedefinieerd als de duur t/m 720 min (12 uur)
        dispcoef = GEVDispCoefBasisstatistiek2015(DuurMinuten, Zichtjaar, Scenario, Corner)
        locpar = GEVLocParBasisstatistiek2015(DuurMinuten, Zichtjaar, Scenario, Corner)
        scalepar = dispcoef * locpar
        shapepar = GEVShapeParBasisstatistiek2015(DuurMinuten, Zichtjaar, Scenario)
        P = GEVCDF(locpar, scalepar, shapepar, Volume)
    Else
        Volume = Volume / 1.02
        dispcoef = GLODispCoefBasisstatistiek2018(DuurMinuten)
        locpar = GLOLocParBasisstatistiek2018(DuurMinuten)
        scalepar = dispcoef * locpar
        shapepar = GLOShapeParBasisStatistiek2018(DuurMinuten)
        P = GLOCDF(locpar, scalepar, shapepar, Volume)
    End If
    STOWA2015_2018_JAARROND_T = 1 / -Math.Log(P)
End Function

Public Function STOWA2015_JAARROND_V(DuurMinuten As Integer, T As Double, Zichtjaar As Integer, Scenario As String, Corner As String) As Double
    'deze functie berekent de herhalingstijd voor jaarrond-neerslagstatistiek conform STOWA, 2015 met gegeven Herhalingstijd en duur in minuten
    'in de tweede iteratie gebruiken we de herhalingstijd die werd berekend weer als input
    Dim locpar As Double, scalepar As Double, shapepar As Double, dispcoef As Double
    Dim P As Double
    P = Exp(-1 / T)
        dispcoef = GEVDispCoefBasisstatistiek2015(DuurMinuten, Zichtjaar, Scenario, Corner)
        locpar = GEVLocParBasisstatistiek2015(DuurMinuten, Zichtjaar, Scenario, Corner)
        scalepar = dispcoef * locpar
        shapepar = GEVShapeParBasisstatistiek2015(DuurMinuten, Zichtjaar, Scenario)
        STOWA2015_JAARROND_V = GEVINVERSE(locpar, scalepar, shapepar, P)
End Function

Public Function STOWA2015_2018_JAARROND_V(DuurMinuten As Integer, T As Double, Zichtjaar As Integer, Scenario As String, Corner As String) As Double
    'deze functie berekent de herhalingstijd voor jaarrond-neerslagstatistiek conform STOWA, 2015/2018 met gegeven Herhalingstijd en duur in minuten
    'in de tweede iteratie gebruiken we de herhalingstijd die werd berekend weer als input
    Dim locpar As Double, scalepar As Double, shapepar As Double, dispcoef As Double
    Dim P As Double
    P = Exp(-1 / T)
    If DuurMinuten > 720 Then
        dispcoef = GEVDispCoefBasisstatistiek2015(DuurMinuten, Zichtjaar, Scenario, Corner)
        locpar = GEVLocParBasisstatistiek2015(DuurMinuten, Zichtjaar, Scenario, Corner)
        scalepar = dispcoef * locpar
        shapepar = GEVShapeParBasisstatistiek2015(DuurMinuten, Zichtjaar, Scenario)
        STOWA2015_2018_JAARROND_V = GEVINVERSE(locpar, scalepar, shapepar, P)
    Else
        dispcoef = GLODispCoefBasisstatistiek2018(DuurMinuten)
        locpar = GLOLocParBasisstatistiek2018(DuurMinuten)
        scalepar = dispcoef * locpar
        shapepar = GLOShapeParBasisStatistiek2018(DuurMinuten)
        STOWA2015_2018_JAARROND_V = GLOINVERSE(locpar, scalepar, shapepar, P) * 1.02
    End If
End Function



Public Function STOWA2019_JAARROND_V(ByVal DuurMinuten As Integer, ByVal T As Double, ByVal Zichtjaar As Integer, ByVal Scenario As String, ByVal Corner As String) As Double
    'deze functie berekent de herhalingstijd voor jaarrond-neerslagstatistiek conform STOWA, 2019 met gegeven Herhalingstijd en duur in minuten
    'in de tweede iteratie gebruiken we de herhalingstijd die werd berekend weer als input
    'voor de multipliers van klimaatscenario's onder korte duren (<= 2 uur) zie STOWA 2019-19 deelrapport 2 tabel 5
    Dim P As Double, Volume As Double
    Dim locpar As Double, scalepar As Double, shapepar As Double, dispcoef As Double
    Dim KorteDuurMultiplier As Double
    Dim LangeDuurMultiplier As Double
    Dim Multiplier As Double
    
    'bepaal eerst het volume op basis van de kansverdelingsparameters voor het huidige klimaat (2014)
    P = Exp(-1 / T)
    If DuurMinuten > 720 Then
        dispcoef = GEVDispCoefBasisstatistiek2019LangeDuur(DuurMinuten)
        locpar = GEVLocparBasisstatistiek2019LangeDuur(DuurMinuten)
        scalepar = dispcoef * locpar
        shapepar = GEVShapeParBasisstatistiek2019LangeDuur(DuurMinuten)
        Volume = GEVINVERSE(locpar, scalepar, shapepar, P)
    Else
        dispcoef = GLODispCoefBasisstatistiek2019KorteDuur(DuurMinuten)
        locpar = GLOLocparBasisstatistiek2019KorteDuur(DuurMinuten)
        scalepar = dispcoef * locpar
        shapepar = GLOShapeParBasisstatistiek2019KorteDuur(DuurMinuten, T)
        Volume = GLOINVERSE(locpar, scalepar, shapepar, P)
    End If
        
    'bepaal nu de aanpassingen als gevolg van het onderhavige klimaat
    'voor duren tot en met 120 minuten zijn multipliers van toepassing. Deze zijn te vinden in rapport STOWA 2019-19, deelrapport 2, tabel 5
    'voor langere duren wordt het relatieve effect van het klimaatscenario uit de 'oude' statistieken (2015) toegepast, dus scenario_oud/huidig_oud
    'Siebe: let op: het kan zijn dat we de langeduurmultiplier toch moeten halen uit een ander scenario dan het onderhavige
    'dit staat beschreven in STOWA2019-19 deelrapport 2, p42 en 43.
    If Zichtjaar <> 2014 Then
        KorteDuurMultiplier = STOWA2019_KORTEDUUR_MULTIPLIER1(Zichtjaar, Scenario, Corner)
        LangeDuurMultiplier = STOWA2019_LANGEDUUR_MULTIPLIER(DuurMinuten, T, Zichtjaar, Scenario, Corner)
    Else
        KorteDuurMultiplier = 1
        LangeDuurMultiplier = 1
    End If
    
    If DuurMinuten <= 120 Then
        'pas alleen de korteduur-multiplier toe
        Multiplier = KorteDuurMultiplier
        Volume = Volume * Multiplier
    ElseIf DuurMinuten < 1440 Then
        'volgende regel code betreft een aanvulling door Rudolf Versteeg (6-4-2020)
        'RV: anders heeft de langeduurmultiplier een waarde bij lager dan bij duur 1440
        LangeDuurMultiplier = STOWA2019_LANGEDUUR_MULTIPLIER(1440, T, Zichtjaar, Scenario, Corner)
        'hier moet worden geïnterpoleerd tussen beide multipliers
        Multiplier = Interpolate(120, KorteDuurMultiplier, 1440, LangeDuurMultiplier, CDbl(DuurMinuten))
        Volume = Volume * Multiplier
    Else
        'pas alleen de langeduur-multiplier toe
        Multiplier = LangeDuurMultiplier
        Volume = Volume * Multiplier
    End If
    
    STOWA2019_JAARROND_V = Volume
    
End Function

Public Function STOWA2019_LANGEDUUR_MULTIPLIER(DuurMinuten As Integer, T As Double, Zichtjaar As Integer, Scenario As String, Corner As String) As Double
    'deze functie berekent de klimaatmultiplier voor een gegeven klimaatscenario, jaarrond
    'de multiplier is gebaseerd op de verhouding klimaat/huidig uit de statistiek van 2015. Dit mag omdat sindsdien de verhoudingen onveranderd zijn gebleven
    STOWA2019_LANGEDUUR_MULTIPLIER = STOWA2015_JAARROND_V(DuurMinuten, T, Zichtjaar, Scenario, Corner) / STOWA2015_JAARROND_V(DuurMinuten, T, 2014, "", "")
End Function

Public Function STOWA2019_MULTIPLIER_WINTER(DuurMinuten As Integer, T As Double, Zichtjaar As Integer, Scenario As String, Corner As String) As Double
    'deze functie berekent de klimaatmultiplier voor een gegeven klimaatscenario, winterseizoen
    'de multiplier is gebaseerd op de verhouding klimaat/huidig uit de statistiek van 2015. Dit mag omdat sindsdien de verhoudingen onveranderd zijn gebleven
    'overigens wordt hierover niets gezegd in het rapport van STOWA uit 2019. Deze informatie is achterhaald in een gesprek met de auteur Rudolf Versteeg op 30-3-2020
    STOWA2019_MULTIPLIER_WINTER = STOWA2015_WINTER_V(DuurMinuten, T, Zichtjaar, Scenario, Corner) / STOWA2015_WINTER_V(DuurMinuten, T, 2014, "", "")
End Function

Public Function STOWA2019_KORTEDUUR_MULTIPLIER2(DuurMinuten As Integer, T As Double, Zichtjaar As Integer, Scenario As String, Corner As String) As Double
    
        If Zichtjaar = 2030 Then
            If Corner = "lower" Then
                STOWA2019_KORTEDUUR_MULTIPLIER2 = STOWA2015_JAARROND_V(DuurMinuten, T, Zichtjaar, Scenario, Corner) / STOWA2015_JAARROND_V(DuurMinuten, T, 2014, "", "")
            ElseIf Corner = "upper" Then
                STOWA2019_KORTEDUUR_MULTIPLIER2 = STOWA2015_JAARROND_V(DuurMinuten, T, Zichtjaar, Scenario, Corner) / STOWA2015_JAARROND_V(DuurMinuten, T, 2014, "", "")
            Else
                STOWA2019_KORTEDUUR_MULTIPLIER2 = 1
            End If
        ElseIf Zichtjaar = 2050 Then
            If Scenario = "GH" And Corner = "lower" Then
                'de equivalent van 2050 GH lower is bij duren <= 2 uur scenario 2050 GL lower
                STOWA2019_KORTEDUUR_MULTIPLIER2 = STOWA2015_JAARROND_V(DuurMinuten, T, Zichtjaar, "GL", "lower") / STOWA2015_JAARROND_V(DuurMinuten, T, 2014, "", "")
            ElseIf Scenario = "WL" And Corner = "upper" Then
                'de equivalent van 2050 WL upper is bij duren <= 2 uur scenario 2050 WH upper
                STOWA2019_KORTEDUUR_MULTIPLIER2 = STOWA2015_JAARROND_V(DuurMinuten, T, Zichtjaar, "WH", "upper") / STOWA2015_JAARROND_V(DuurMinuten, T, 2014, "", "")
            Else
                STOWA2019_KORTEDUUR_MULTIPLIER2 = 1
            End If
        ElseIf Zichtjaar = 2085 Then
            If Scenario = "GH" And Corner = "lower" Then
                'de equivalent van 2085 GH lower is bij duren <= 2 uur scenario 2050 GL lower
                STOWA2019_KORTEDUUR_MULTIPLIER2 = STOWA2015_JAARROND_V(DuurMinuten, T, Zichtjaar, "GL", "lower") / STOWA2015_JAARROND_V(DuurMinuten, T, 2014, "", "")
            ElseIf Scenario = "WL" And Corner = "upper" Then
                'de equivalent van 2085 WL upper is bij duren <= 2 uur scenario 2050 WL upper
                STOWA2019_KORTEDUUR_MULTIPLIER2 = STOWA2015_JAARROND_V(DuurMinuten, T, Zichtjaar, "WL", "upper") / STOWA2015_JAARROND_V(DuurMinuten, T, 2014, "", "")
            Else
                STOWA2019_KORTEDUUR_MULTIPLIER2 = 1
            End If
        End If
End Function

Public Function STOWA2019_KORTEDUUR_MULTIPLIER1(Zichtjaar As Integer, Scenario As String, Corner As String) As Double
        'de multipliers voor korte duur zijn ontleend aan de brochure KNMI '14 klimaatscenarios voor Nederland (KNMI).
        'de oude multipliers (tov zichtjaar 1995) zijn te vinden in tabel op pag 5. Deze moesten worden gecorrigeerd zodat ze uitgedrukt worden
        'tov zichtjaar 2014. Voor zes scenario's is dit al gedaan door KNMI en zijn de waarden gepubliceerd in STOWA 2019-19.
        
        'let op: KNMI heeft dus voor slechts zes generieke scenario's resultaten geleverd: 2030_upper, 2030_lower, 2050_upper, 2050_lower, 2085_upper en 2085_lower
        'in het rapport van STOWA, deelrapport 2 p43 tabel 4, wordt de aansluiting van deze generieke scenario's met de zes best passende scenario's voor lange duur gegeven
        'zo krijgt scenario 2050 GH Lower voor duren > 24 uur de volgende equivalent voor korte duur < 2 uur: 2050 GL Lower
        'en krijgt scenario 2050 WL Upper voor duren > 24 uur de volgende equivalent voor korte duur < 2 uur: 2050 WL Upper
        'en krijgt scenario 2085 GH Lower voor duren > 24 uur de volgende equivalent voor korte duur < 2 uur: 2085 GL Lower
        
        If Zichtjaar = 2030 Then
            If Corner = "lower" Then
                STOWA2019_KORTEDUUR_MULTIPLIER1 = 1.0385  'multiplier tov 2014 was al door stowa gegeven; verfijnd door Rudolf Versteeg op 6-4-2020
            ElseIf Corner = "upper" Then
                STOWA2019_KORTEDUUR_MULTIPLIER1 = 1.077  'multiplier tov 2014 was al door stowa gegeven
            Else
                STOWA2019_KORTEDUUR_MULTIPLIER1 = 1
            End If
        ElseIf Zichtjaar = 2050 Then
            If Scenario = "GH" And Corner = "lower" Then
                'de equivalent van 2050 GH lower is bij duren <= 2 uur scenario 2050 GL lower
                STOWA2019_KORTEDUUR_MULTIPLIER1 = 1.0385  'multiplier tov 2014 was al door STOWA gegeven; verfijnd door Rudolf Versteeg op 6-4-2020
            ElseIf Scenario = "WL" And Corner = "upper" Then
                'de equivalent van 2050 WL upper is bij duren <= 2 uur scenario 2050 WH upper
                STOWA2019_KORTEDUUR_MULTIPLIER1 = 1.2125  'multiplier tov 2014 was al door STOWA gegeven: 1.213; verfijnd door Rudolf Versteeg op 6-4-2020
            Else
                STOWA2019_KORTEDUUR_MULTIPLIER1 = 1
            End If
        ElseIf Zichtjaar = 2085 Then
            If Scenario = "GH" And Corner = "lower" Then
                'de equivalent van 2085 GH lower is bij duren <= 2 uur scenario 2050 GL lower
                STOWA2019_KORTEDUUR_MULTIPLIER1 = 1.064  'multiplier tov 2014 was al door STOWA gegeven
            ElseIf Scenario = "WL" And Corner = "upper" Then
                'de equivalent van 2085 WL upper is bij duren <= 2 uur scenario 2050 WL upper
                STOWA2019_KORTEDUUR_MULTIPLIER1 = 1 + ((3.5 - 0.3) / 3.5 * 45) / 100 'multiplier tov 2014 was al door STOWA gegeven; ; verfijnd door Rudolf Versteeg op 6-4-2020
            Else
                STOWA2019_KORTEDUUR_MULTIPLIER1 = 1
            End If
        End If

End Function

Public Function GLOLocparBasisstatistiek2019KorteDuur(DuurMinuten As Integer) As Double
    'deze functie berekent de locatieparameter zeta voor de GLO-kansverdeling voor korte duur (10 minuten t/m 12 uur) volgens STOWA 2019, deelrapport 1 p12
    GLOLocparBasisstatistiek2019KorteDuur = 1.02 * (7.339 + 0.848 * Log10(CDbl(DuurMinuten)) + 2.844 * (Log10(CDbl(DuurMinuten))) ^ 2)
End Function


Public Function GLODispCoefBasisstatistiek2019KorteDuur(DuurMinuten As Integer) As Double
    'deze functie berekent de dispersiecoefficient epsylon voor de GLO-kansverdeling voor korte duur (10 minuten t/m 12 uur) volgens STOWA 2019, deelrapport 1 p12
    'let op: dit is NIET de schaalparameter uit de GLO=verdeling. Daarvoor moet eerst nog met de locatiepar (zeta) worden vermenigvuldigd
    If DuurMinuten <= 104 Then
        GLODispCoefBasisstatistiek2019KorteDuur = 0.04704 + 0.1978 * Log10(CDbl(DuurMinuten)) - 0.05729 * (Log10(CDbl(DuurMinuten))) ^ 2
    Else
        GLODispCoefBasisstatistiek2019KorteDuur = 0.2801 - 0.0333 * Log10(CDbl(DuurMinuten))
    End If
End Function


