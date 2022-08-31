# -*- coding: utf-8 -*-
"""
Created on Tue May 31 12:14:43 2022

@author: baret
"""

import pandas as pd
import xlsxwriter
import numpy as np

def aggr_result(directory, instance_size):
    file_result=directory + '/Resultats.xlsx'
    workbook = xlsxwriter.Workbook(file_result)
    worksheet = workbook.add_worksheet('Result')
    worksheet.write('A1','Instance')
    worksheet.write('B1','nbClients')
    worksheet.write('C1','nbFacilities')
    worksheet.write('D1','Probabilité moyenne')
    worksheet.write('E1','Resolution Time')
    worksheet.write('F1','Niveau 0')
    worksheet.write('G1','Niveau 1')
    worksheet.write('H1','Niveau 2')
    worksheet.write('I1','Niveau 3')

    
    for i in range(instance_size) :
        fichier= (directory + f'/Instances/Instance{i}.xlsx')
        fichierSol= (directory + f'/Resultats/instance_{i}_C_False.xlsx')
        #donnees générales 
            
        Donnees=pd.read_excel(fichier,sheet_name="Donnees", index_col=0 )
        Proba=pd.read_excel(fichier,sheet_name="Proba", index_col=0 )       
        #données facilities 
 #       Capacites=pd.read_excel(fichier,sheet_name="Congestions", index_col=0 )
                
        Do=Donnees.to_numpy()
        P=Proba.to_numpy()
        print(P)     
         # données solutions   
        Resolution_time=pd.read_excel(fichierSol,sheet_name="Sheet1", index_col=0)
        RT=Resolution_time.to_numpy()

        R=RT.size
        if R>0:
            ResT=RT[0,3]
            worksheet.write(i+1,4,ResT)
        nbClients=Do[0]
        nbFacilities=Do[1]
        worksheet.write(i+1,0,i)
        worksheet.write(i+1,1,nbClients)
        worksheet.write(i+1,2,nbFacilities)
        
        
      #  C2=np.resize(C,(1,nbClients))
        
        NS=get_sol(nbFacilities, fichierSol)
        N0=0
        N1=0
        N2=0
        N3=0
        ProbaMoyenne=0
        NbF=int(nbFacilities)
        for k in range(NbF):
            ProbaMoyenne=ProbaMoyenne+P[k,7]
            if NS[k,0]==1:
                N0=N0+1
            elif NS[k,1]==1:
                N1=N1+1
            elif NS[k,2]==1:
                N2=N2+1
            elif NS[k,3]==1:
                N3=N3+1
        
        ProbaMoyenne=ProbaMoyenne/NbF
        worksheet.write(i+1,3,ProbaMoyenne)
        worksheet.write(i+1,5,N0)
        worksheet.write(i+1,6,N1)
        worksheet.write(i+1,7,N2)
        worksheet.write(i+1,8,N3)
    
        PositionsFacilities=pd.read_excel(fichier,sheet_name="Position-facilities", index_col=0 )
        
        positionF=PositionsFacilities.to_numpy()
        PF=positionF.size
        if PF>0 :
            D=calcul_distance(NbF, positionF)
            worksheet.write(i+1,9,D)
        
        PositionsClients=pd.read_excel(fichier,sheet_name="Position-client", index_col=0 )
        positionC=PositionsClients.to_numpy()
        PC=positionC.size
        if PC>0 :
            D=calcul_distance(NbF, positionC)
            worksheet.write(i+1,10,D)
    workbook.close()
    return 
        


def compare_result(directory,instance_size):
    file_result='C:/Users/baret/Documents/Simulateur/Resultats-compare.xlsx'
    workbook = xlsxwriter.Workbook(file_result)
    worksheet = workbook.add_worksheet('Result')
    
    worksheet.write('A1','Instance')
    worksheet.write('B1','nbClients')
    worksheet.write('C1','nbFacilities')
    worksheet.write('D1','Coefficient de congestion')
    worksheet.write('E1','Resolution Time Non relaxé')
    worksheet.write('F1','Resolution Time Relaxé')
    worksheet.write('G1',' Nb Var Diff')
    worksheet.write('H1','Diff Obj')
    
    for i in range(instance_size) :
        fichier= (directory + f'/Instances/Instance{i}.xlsx')
        fichierSolNR= (directory + f'/Resultats/instance_{i}_False.xlsx')
        fichierSolR= (directory + f'/Resultats/instance_{i}_True.xlsx')
        #donnees générales 
            
        Donnees=pd.read_excel(fichier,sheet_name="Donnees", index_col=0 )
        Congestion=pd.read_excel(fichier,sheet_name="Congestions", index_col=0 )       
        #données facilities 
 #       Capacites=pd.read_excel(fichier,sheet_name="Congestions", index_col=0 )
                
        Do=Donnees.to_numpy()
        C=Congestion.to_numpy()
         
        nbClients=Do[0]
        nbFacilities=Do[1]
        worksheet.write(i+1,0,i)
        worksheet.write(i+1,1,nbClients)
        worksheet.write(i+1,2,nbFacilities)
        
         # données solutions   
        Resolution_timeNR=pd.read_excel(fichierSolNR,sheet_name="Sheet1", index_col=0)
        RTNR=Resolution_timeNR.to_numpy()

        R=RTNR.size
        if R>0:
            ResT=RTNR[0,3]
            worksheet.write(i+1,4,ResT)
            ResNR=RTNR[1:,0]
           # print(ResNR)
        else :
            ResNR=np.empty(shape=1)
        Resolution_timeR=pd.read_excel(fichierSolR,sheet_name="Sheet1", index_col=0)
        RTR=Resolution_timeR.to_numpy()

        RR=RTR.size
        if RR>0:
            ResT=RTR[0,3]
            worksheet.write(i+1,5,ResT)
            ResR=RTR[1:,0]
           # print(ResR)
        else :
            ResR=np.empty(shape=1)
        nbVArDiff=0
       # print(ResR.size)
       # print(ResNR.size)
       
        for j in range(min(ResNR.size,ResR.size)):
            if ResNR[j] != ResR[j]:
                nbVArDiff=nbVArDiff+1
        
        DiffScore=0
        if R>0 & RR>0:
            DiffScore=RTR[0,1]-RTNR[0,1]
            
      #  C2=np.resize(C,(1,nbClients))
        ET=np.std(C)

        M=np.mean(C)

        CO=ET/M
        worksheet.write(i+1,3,CO)
        worksheet.write(i+1,6,nbVArDiff)
        worksheet.write(i+1,7,DiffScore)
    workbook.close()
    
    
def compare_capacity(directory1,directory2,instance_size):
    file_result=directory2+'/Resultats-compare.xlsx'
    workbook = xlsxwriter.Workbook(file_result)
    worksheet = workbook.add_worksheet('Result')
        
    worksheet.write('A1','Instance')
    worksheet.write('B1','nbClients')
    worksheet.write('C1','nbFacilities')
    worksheet.write('D1','Coefficient de congestion')
    worksheet.write('E1','Resolution Time 1')
    worksheet.write('F1','Resolution Time 2')
    worksheet.write('G1',' Nb Var Diff')
    worksheet.write('H1','Diff Obj')
        
    for i in range(instance_size) :
        fichier= (directory1 + f'/Instances/Instance{i}.xlsx')
        fichierSol1= (directory1 + f'/Resultats/instance_{i}_C_False.xlsx')
        fichierSol2= (directory2 + f'/Resultats/instance_{i}_C_False.xlsx')
            #donnees générales 
                
        Donnees=pd.read_excel(fichier,sheet_name="Donnees", index_col=0 )
        Congestion=pd.read_excel(fichier,sheet_name="Congestions", index_col=0 )       
            #données facilities 
     #       Capacites=pd.read_excel(fichier,sheet_name="Congestions", index_col=0 )
                    
        Do=Donnees.to_numpy()
        C=Congestion.to_numpy()
             
        nbClients=Do[0]
        nbFacilities=Do[1]
        worksheet.write(i+1,0,i)
        worksheet.write(i+1,1,nbClients)
        worksheet.write(i+1,2,nbFacilities)
            
             # données solutions   
        Resolution_time1=pd.read_excel(fichierSol1,sheet_name="Sheet1", index_col=0)
        RT1=Resolution_time1.to_numpy()

        R=RT1.size
        if R>0:
           ResT=RT1[0,3]
           worksheet.write(i+1,4,ResT)
           Res1=RT1[1:,0]
           #print(Res1)
        else :
            Res1=np.empty(shape=1)
        
        Resolution_time2=pd.read_excel(fichierSol2,sheet_name="Sheet1", index_col=0)
        RT2=Resolution_time2.to_numpy()

        RR=RT2.size
        if RR>0:
            ResT=RT2[0,3]
            worksheet.write(i+1,5,ResT)
            Res2=RT2[1:,0]
            #print(Res2)
        else :
            Res2=np.empty(shape=1)
        nbVArDiff=0
       # print(Res2.size)
       # print(Res1.size)
           
        for j in range(min(Res1.size,Res2.size)):
            if Res1[j] != Res2[j]:
                nbVArDiff=nbVArDiff+1
            
        DiffScore=0
########pb de taille quand pas résolu !!!!!!!!!!!!!!!!!!
        print(i)
        if RT1.size>0 :
            if RT2.size>0 :
                print('rentre')
                if RT1[0,1]>0:
                    DiffScore=(RT1[0,1]-RT2[0,1])/RT1[0,1]
        if DiffScore>0:
           print('positif')
        else :
            DiffScore=0
          #  C2=np.resize(C,(1,nbClients))
        ET=np.std(C)

        M=np.mean(C)

        CO=ET/M
        worksheet.write(i+1,3,CO)
        print( 'CO')
        worksheet.write(i+1,6,nbVArDiff)
        print('nbVar diff')
        print(DiffScore)
        worksheet.write(i+1,7,DiffScore)
        print('Diff score')
    workbook.close()
        
def compare_relaxed(directory1,directory2,instance_size):
    file_result=directory2+'/Resultats-relaxed.xlsx'
    workbook = xlsxwriter.Workbook(file_result)
    worksheet = workbook.add_worksheet('Result')
        
    worksheet.write('A1','Instance')
    worksheet.write('B1','nbClients')
    worksheet.write('C1','nbFacilities')
    worksheet.write('D1','Coefficient de congestion')
    worksheet.write('E1','Resolution Time 1')
    worksheet.write('F1','Resolution Time 2')
    worksheet.write('G1',' Nb Var Diff')
    worksheet.write('H1','Diff Obj')
        
    for i in range(instance_size) :
        fichier= (directory1 + f'/Instances/Instance{i}.xlsx')
        fichierSol1= (directory1 + f'/Resultats/instance_{i}_C_False.xlsx')
        fichierSol2= (directory2 + f'/Resultats/instance_{i}_C_True.xlsx')
            #donnees générales 
                
        Donnees=pd.read_excel(fichier,sheet_name="Donnees", index_col=0 )
        Congestion=pd.read_excel(fichier,sheet_name="Congestions", index_col=0 )       
            #données facilities 
     #       Capacites=pd.read_excel(fichier,sheet_name="Congestions", index_col=0 )
                    
        Do=Donnees.to_numpy()
        C=Congestion.to_numpy()
             
        nbClients=Do[0]
        nbFacilities=Do[1]
        worksheet.write(i+1,0,i)
        worksheet.write(i+1,1,nbClients)
        worksheet.write(i+1,2,nbFacilities)
            
             # données solutions   
        Resolution_timeNR=pd.read_excel(fichierSol1,sheet_name="Sheet1", index_col=0)
        RTNR=Resolution_timeNR.to_numpy()

        R=RTNR.size
        if R>0:
           ResT=RTNR[0,3]
           worksheet.write(i+1,4,ResT)
           ResNR=RTNR[1:,0]
        #   print(ResNR)
        else :
            ResNR=np.empty(shape=1)
        
        Resolution_timeR=pd.read_excel(fichierSol2,sheet_name="Sheet1", index_col=0)
        RTR=Resolution_timeR.to_numpy()

        RR=RTR.size
        if RR>0:
            ResT=RTR[0,3]
            worksheet.write(i+1,5,ResT)
            ResR=RTR[1:,0]
         #   print(ResR)
        else :
            ResR=np.empty(shape=1)
        nbVArDiff=0
      #  print(ResR.size)
       # print(ResNR.size)
           
        for j in range(min(ResNR.size,ResR.size)):
            if ResNR[j] != ResR[j]:
                nbVArDiff=nbVArDiff+1
            
        DiffScore=0
        if R>0 & RR>0:
            DiffScore=RTR[0,1]-RTNR[0,1]
            print('rentrer')
                
          #  C2=np.resize(C,(1,nbClients))
        ET=np.std(C)

        M=np.mean(C)

        CO=ET/M
        worksheet.write(i+1,3,CO)
        worksheet.write(i+1,6,nbVArDiff)
        worksheet.write(i+1,7,DiffScore)
    workbook.close()

def get_sol( nbFacilities,fichierSol):
    # données solutions   
    Solutions=pd.read_excel(fichierSol,sheet_name="Sheet1", index_col=0)
    S=Solutions.to_numpy()
    R=S.size
    NS=np.zeros((int(nbFacilities), 4))
    if R>0:
        S2=S[1:,0]
        for i in range(len(S2)):
            NS[(int(i//4))-1,i%4]=S2[i]
    return NS 

def calcul_distance(nb, position):
    nbF=int(nb)
    D=np.zeros((nbF, nbF))
    for k1 in range (nbF):
        for k2 in range (nbF):
            if k1 != k2 :
                D[k1,k2]=abs(position[k1,0]-position[k2,0])**2+abs(position[k1,1]-position[k2,1])**2
    dist=np.sum(D)/ (nbF**2)
    return dist

    
directory1='C:/Users/baret/Documents/Simulateur/test-non-homogène/Instances-12-120-0,1'   
model='C'  
RT=aggr_result(directory1, 30)
#compare_capacity(directory1,directory2 ,30)
#compare_relaxed(directory1, directory2, 30)