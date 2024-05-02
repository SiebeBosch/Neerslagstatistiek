# -*- coding: utf-8 -*-
"""
Created on Wed Mar 18 14:14:38 2020

@author: siebe bosch, with contributions of robin nicolai

dit script bevat de kansdichtheidsfuncties vooa alle door STOWA gepubliceerde neerslagstatistieken

publicatie-jaren:
2015: deze bevatten uitsluitend statistieken voor lange duren

"""

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import math


def locatieparameter_KNMI_2019_jaar(Duur):
    
    # Invoer: Duur in minuten
    # Invoer:Herhalingstijd in jaar 

    if Duur <= 720 : # Duur is kleiner dan 12u, we gebruiken een GLO-verdeling
        
        # factor 1.02 toegevoegd, zie vergelijking (5)
        
        loc = 1.02*(7.339 + 0.848 * np.log10(Duur) + 2.844 * np.log10(Duur)**2)       
    
    else: # Duur is groter dan 12u, we gebruiken een GEV-verdeling
    
        # factor 1.02 toegevoegd, zie vergelijking (8) 
        # N.B. Dit is een andere factor dan die voor de GLO-verdeling
        
        loc = 1.02*(0.239 - 0.0250 * np.log(Duur/60))**(-1/0.512)
    
    return loc

def dispersiecoefficient_KNMI_2019_jaar(Duur): 
    
    if Duur <= 720 : # Duur is kleiner dan 12u, we gebruiken een GLO-verdeling
        
        if Duur <= 104:
                    
            disp = 0.04704 + 0.1978 * np.log10(Duur) - 0.05729 * np.log10(Duur)**2
            
        else:
            
            disp = 0.2801 - 0.0333 * np.log10(Duur)
            
    else: # Duur is groter dan 12u, we gebruiken een GEV-verdeling
    
        disp = 0.478 - 0.0681 * np.log10(Duur)
        
    return disp

def vormparameter_KNMI_2019_jaar(Duur,T):
    
    if Duur <= 720 : # Duur is kleiner dan 12u, we gebruiken een GLO-verdeling
        
        if Duur <= 90:
                    
            vorm = -0.0336 - 0.264 * np.log10(Duur) + 0.0636 * (np.log10(Duur))**2  
            
        else:
            
            if T <= 120: # in formule (7) STOWA2019 staat 'T<=1'
                
                vorm = -0.0336 - 0.264 * np.log10(Duur) + 0.0636 * (np.log10(Duur))**2  

            else:
                
                vorm = -0.310 - 0.0544 * np.log10(Duur) + 0.0288 * (np.log10(Duur))**2  
       
    else: # Duur is groter dan 12u, we gebruiken een GEV-verdeling
        
        vorm = 0.118 - 0.266 * np.log10(Duur) + 0.0586 * (np.log10(Duur)**2) 
    
    return vorm

def vol_KNMI_2019_jaar(Duur, T):
    
    # Invoer: Duur in minuten
    # Invoer: Herhalingstijd in jaar 
    # Uitvoer: Volume in mm

    Locpar = locatieparameter_KNMI_2019_jaar(Duur)
    Vormpar = vormparameter_KNMI_2019_jaar(Duur,T)
    dispcoeff = dispersiecoefficient_KNMI_2019_jaar(Duur)
    Schaalpar = dispcoeff * Locpar
    
    if Duur <= 720: # Duur is kleiner dan 12u, we gebruiken een GLO-verdeling
    
        Vol_KNMI_2019_jaar = Locpar + (Schaalpar/ Vormpar) * (1 - ((1-np.exp(-1/T))/(np.exp(-1/T))) ** Vormpar)   
        
        if (Duur > 90 )& (120< T)&(T <= 165):
            # Vergelijking (11)
            Vormpar_T120 = vormparameter_KNMI_2019_jaar(Duur,120)
            Vol_KNMI_2019_jaar_T120 = Locpar + (Schaalpar/ Vormpar_T120) * (1 - ((1-np.exp(-1/120))/(np.exp(-1/120))) ** Vormpar_T120)
            Vol_KNMI_2019_jaar = max(Vol_KNMI_2019_jaar, Vol_KNMI_2019_jaar_T120)
        
        #Vol_KNMI_2019_jaar = 1.02*Vol_KNMI_2019_jaar 
        # Deze 1.02 is conform STOWA2019 in de locatieparameter verwerkt.
        
    else: # Duur is groter dan 12u, we gebruiken een GEV-verdeling
        
        Vol_KNMI_2019_jaar = Locpar + (Schaalpar / Vormpar) * (1- (1/T)**Vormpar)
            
    return Vol_KNMI_2019_jaar