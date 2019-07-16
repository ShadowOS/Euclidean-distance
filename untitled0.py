# -*- coding: utf-8 -*-
"""
Created on Thu Jul  4 14:15:09 2019

@author: Karnika
"""

from gurobipy import*
import os
import xlrd
import pandas as pd
import xlsxwriter
from scipy import spatial
import numpy
from sklearn.metrics.pairwise import euclidean_distances
import math



book = xlrd.open_workbook(os.path.join("project.xlsx"))

sh = book.sheet_by_name("Sheet2")
coordinates_of_A = []
xA = {}
yA = {}

i = 1
while True:
    try:
        sp = sh.cell_value(i,0)
        coordinates_of_A.append(sp)
        xA[sp]=sh.cell_value(i,1)
        yA[sp]=sh.cell_value(i,2)
        i = i + 1
        
    except IndexError:
        break
sh = book.sheet_by_name("Sheet1")
coordinates_of_B = []
xB = {}
yB = {}

i = 1
while True:
    try:
        sp = sh.cell_value(i,0)
        coordinates_of_B.append(sp)
        xB[sp]=sh.cell_value(i,1)
        yB[sp]=sh.cell_value(i,2)
        i = i + 1
        
    except IndexError:
        break   
def calculate_dist(x1, x2):
    eudistance = spatial.distance.euclidean(x1, x2)
    return(eudistance)    
    
x1=[]    
x2=[]    

distance={}    
    
for i in coordinates_of_A:
    x1.clear()
    x1.append(xA[i])
    x1.append(yA[i])
    for j in coordinates_of_B:
        x2.clear()
        x2.append(xB[j])
        x2.append(yB[j])        
        distance[i,j]=calculate_dist(x1, x2)
    
workbook=xlsxwriter.Workbook('eproject.xlsx')
worksheet=workbook.add_worksheet('distance_ij')
i=1
k=0
for x in coordinates_of_A:
    j=1
    l=0
    for y in coordinates_of_B:
        e=distance[x,y]
        worksheet.write(i,j,e)
        f=coordinates_of_A[k]
        worksheet.write(i,0,f)
        g=coordinates_of_B[l]
        worksheet.write(0,j,g)
        l+=1
        j+=1    
    i+=1
    k+=1

workbook.close()