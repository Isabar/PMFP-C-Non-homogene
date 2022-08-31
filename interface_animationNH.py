# -*- coding: utf-8 -*-
"""
Created on Mon May 23 15:03:18 2022

@author: baret
"""

# -*- coding: utf-8 -*-
"""
Created on Mon May 23 14:50:01 2022

@author: baret
"""
import tkinter as tk
from PIL import ImageTk, Image
import pandas as pd 
import numpy as np 
from analyse_instance import *

W=1000
H=600
fenetre =tk.Tk()
canvas = tk.Canvas(fenetre, width=W, height=H,background='white')
photo = (Image.open("hopital.png"))

def creation_clients(Do, De, PC,T): 

    nbClients=Do[0]
    #display clients 
    nbC=int(nbClients)
    for i in range(nbC):
        Ratio=De[i]/max(De)
        k=T[i,0]-1
        #print(Ratio)
        canvas.create_rectangle((W*PC[i,0])/10,(H*PC[i,1])/10,((W*PC[i,0])/10)+10*Ratio,((H*PC[i,1])/10)+10*Ratio)
        #canvas.create_line((W*PC[i,0])/10,(H*PC[i,1])/10,(W*P[k,0])/10,(H*P[k,1])/10)
    return
    
def create_facilities(new_image,Do,C,P):
   
    nbFacilities=Do[1]
    nbF=int(nbFacilities)
    
    #display image of facilities
    for kk in range(nbF):
        canvas.create_image((W*P[kk,0])/10,(H*P[kk,1])/10, image=new_image)
    return 

def display_line(Do, PC,P,T):
    nbClients=Do[0]
    nbC=int(nbClients)
    for i in range(nbC):
        k=T[i,0]-1
        canvas.create_line((W*PC[i,0])/10,(H*PC[i,1])/10,(W*P[k,0])/10,(H*P[k,1])/10)
    return 

def display_line2(Do,PC,P,T):
    nbClients=Do[0]
    nbC=int(nbClients)
    canvas.delete('line')
    for i in range(nbC):
        k2=T[i,1]-1
        k=T[i,0]-1
        canvas.create_line((W*PC[i,0])/10,(H*PC[i,1])/10,(W*P[k,0])/10,(H*P[k,1])/10,fill='white')
        canvas.create_line((W*PC[i,0])/10,(H*PC[i,1])/10,(W*P[k2,0])/10,(H*P[k2,1])/10,dash=(5,1))
    return 


def display_capacity1(Do,P,C,NS):
    nbFacilities=Do[1]
    nbF=int(nbFacilities)
    print(C)
    for k in range(nbF):

        Ratio=C[k]/max(C)
        print(Ratio)
        if NS[k,0]==1:
            print("create l1")
            canvas.create_oval((((W*P[k,0])/10)+20*Ratio[0],((H*P[k,1])/10)+20*Ratio[0]),(((W*P[k,0])/10)-20*Ratio[0],((H*P[k,1])/10)-20*Ratio[0]), width=Ratio[0], fill="green")
        elif NS[k,1]==1:
            print("create l2")
            canvas.create_oval((((W*P[k,0])/10)+20*Ratio[0],((H*P[k,1])/10)+20*Ratio[0]),(((W*P[k,0])/10)-20*Ratio[0],((H*P[k,1])/10)-20*Ratio[0]), width=Ratio[0], fill="blue")
        elif NS[k,2]==1:
            print("create l3")
            canvas.create_oval((((W*P[k,0])/10)+20*Ratio[0],((H*P[k,1])/10)+20*Ratio[0]),(((W*P[k,0])/10)-20*Ratio[0],((H*P[k,1])/10)-20*Ratio[0]), width=Ratio[0], fill="purple")
        elif NS[k,3]==1:
            print("create l4")
            canvas.create_oval((((W*P[k,0])/10)+20*Ratio[0],((H*P[k,1])/10)+20*Ratio[0]),(((W*P[k,0])/10)-20*Ratio[0],((H*P[k,1])/10)-20*Ratio[0]), width=Ratio[0], fill="red")
 

def display_capacity(Do,P,Cmin, C, NS ):
    nbFacilities=Do[1]
    nbF=int(nbFacilities)
    
    for k in range(nbF):

        Ratio=Cmin[k]/(C[k]*4)
        print(Ratio)
        if NS[k,0]==1:
            print("create l1")
            canvas.create_oval((((W*P[k,0])/10)+20*Ratio[0],((H*P[k,1])/10)+20*Ratio[0]),(((W*P[k,0])/10)-20*Ratio[0],((H*P[k,1])/10)-20*Ratio[0]), width=Ratio[0], fill="green")
        elif NS[k,1]==1:
            print("create l2")
            canvas.create_oval((((W*P[k,0])/10)+20*Ratio[0],((H*P[k,1])/10)+20*Ratio[0]),(((W*P[k,0])/10)-20*Ratio[0],((H*P[k,1])/10)-20*Ratio[0]), width=Ratio[0], fill="blue")
        elif NS[k,2]==1:
            print("create l3")
            canvas.create_oval((((W*P[k,0])/10)+20*Ratio[0],((H*P[k,1])/10)+20*Ratio[0]),(((W*P[k,0])/10)-20*Ratio[0],((H*P[k,1])/10)-20*Ratio[0]), width=Ratio[0], fill="purple")
        elif NS[k,3]==1:
            print("create l4")
            canvas.create_oval((((W*P[k,0])/10)+20*Ratio[0],((H*P[k,1])/10)+20*Ratio[0]),(((W*P[k,0])/10)-20*Ratio[0],((H*P[k,1])/10)-20*Ratio[0]), width=Ratio[0], fill="red")
 

def get_data(fichier):
#donnees générales 
    
    Donnees=pd.read_excel(fichier,sheet_name="Donnees", index_col=0 )
       
#données facilities 
    Positions=pd.read_excel(fichier,sheet_name="Position-facilities", index_col=0)
    Capacites=pd.read_excel(fichier,sheet_name="Capacites", index_col=0 )
    
    
    Do=Donnees.to_numpy()
    C=Capacites.to_numpy()
    P=Positions.to_numpy()
       
#données clients 
    PositionC=pd.read_excel(fichier,sheet_name="Position-client", index_col=0)
    PC=PositionC.to_numpy()
    Demandes=pd.read_excel(fichier,sheet_name="Clients", index_col=0)
    Tri = pd.read_excel(fichier,sheet_name="Tri", index_col=0)
    Dee=Demandes.to_numpy()
        
    De=Dee[:,0]
    T=Tri.to_numpy()

    return Do, C,P,PC, De, T

def get_sol( nbFacilities,fichierSol):
    # données solutions   
    Solutions=pd.read_excel(fichierSol,sheet_name="Feuil1", index_col=0)
    S=Solutions.to_numpy()
    NS=np.zeros((int(nbFacilities), 4))
    S2=S[1:,0]
    for i in range(len(S2)):
       NS[(int(i//4))-1,i%4]=S2[i]
    return NS 
    
def display_instance(filename):
    [Do,C,P,PC,De,T]=get_data(filename)
    
    resized_image= photo.resize((60,50), Image.ANTIALIAS)
    new_image= ImageTk.PhotoImage(resized_image)
    canvas.after(1000,create_facilities,new_image,Do, C, P)
    canvas.after(3000,creation_clients,Do, De, PC, T)
    canvas.after(5000,display_line,Do, PC, P, T)
  ##  canvas.after(6000,display_line2,Do, PC, P, T)
    canvas.pack()
    fenetre.mainloop()

def display_results(filename, fichierSol):
    [Do,C,P,PC,De,T]=get_data(filename)
    NS=get_sol(Do[1], fichierSol)
    #Cmin=get_cap_min(12, fichierSol)
    photo = (Image.open("hopital.png"))
    resized_image= photo.resize((60,50), Image.ANTIALIAS)
    new_image= ImageTk.PhotoImage(resized_image)
    canvas.after(1000,create_facilities,new_image,Do, C, P)
    canvas.after(2000,creation_clients,Do, De, PC, T)
    canvas.after(3000,display_line,Do, PC, P, T)
    canvas.after(4000,display_capacity1,Do, P, C, NS)
    #canvas.after(4000, display_capacity2, Do, P, Cmin, C, NS)
    canvas.pack()
    fenetre.mainloop()
    
filename='C:/Users/baret/Documents/Simulateur/test-proba/test-cout/Instances/Instance3.xlsx'
fileSol='C:/Users/baret/Documents/Simulateur/test-proba/test-cout/Resultats/instance_3.xlsx'

display_results(filename, fileSol)
#display_instance(filename)