import  os
from lingo_relationNH import *
from GeneratorNH import *
#from agregation_resultats import *
#from interface_animation import *
import xlsxwriter



def crea_instance(directory,instance_size,minClients,maxClients,minDemand,maxDemand,penality,minFacilities,maxFacilities, nbLevels, cost, probaMax, probaMin):

    for i in range(instance_size):
        nbClients,nbFacilities,Distances,Penalities,Demand,Budget,Ki,Positions,Capacites,ProbaL,ProbaCO, ProbaCOCA,Cost, positionC, positionF, congestion=Generator(minClients,maxClients,minDemand,maxDemand,penality,minFacilities,maxFacilities, nbLevels, cost, probaMax,probaMin)
        filename=directory+'/Instances/Instance'+str(i)+'.xlsx'
        re=excel_write(filename, nbClients,nbFacilities,Budget,Demand,Penalities,Distances,Cost,ProbaL,ProbaCO, ProbaCOCA,Ki,Positions,nbLevels, congestion, Capacites, positionC, positionF);
        #display_instance(filename)
    return 

 

def resolution(instance_debut,instance_fin, directory, model,relaxed,cap,typeProba, budg):
    #3
    instances = [(i) for i in range(instance_debut,instance_fin)]

    for instance in instances:
        create_lingo_ltf_file(directory,instance,model,relaxed,cap,typeProba, budg)
        print('running instance '+str(instance) )
        fileSol= directory+'/Resultats/instance_'+str(instance)+'_'+str(model)+'_'+str(relaxed)+'.xlsx'
       # print(fileSol)
        workbook = xlsxwriter.Workbook(fileSol)

        workbook.close()
        filemodel=directory+'/Modeles/model_'+str(instance)+'_'+str(model)+'_'+str(relaxed)+'.ltf'
        print(filemodel)
        os.system(f'runlingo {filemodel}')

    return 


def resolution_cap(instance_debut,instance_fin, directory, model,relaxed,cap,typeProba, budg):
    #3
    instances = [(i) for i in range(instance_debut,instance_fin)]

    for instance in instances:
        create_lingo_ltf_file(directory,instance,model,relaxed,cap,typeProba,budg)
        print('running instance '+str(instance) )
        fileSol= directory+'/Resultats/instance_'+str(instance)+'_'+str(model)+'_'+str(relaxed)+'.xlsx'
       # print(fileSol)
        workbook = xlsxwriter.Workbook(fileSol)

        workbook.close()
        filemodel=directory+'/Modeles/model_'+str(instance)+'_'+str(model)+'_'+str(relaxed)+'.ltf'
        print(filemodel)
        os.system(f'runlingo {filemodel}')

    return 


def resolution_bi_obj(instance_debut,instance_fin, directory, model,relaxed,cap,typeProba,obj):
    #3
    instances = [(i) for i in range(instance_debut,instance_fin)]

    for instance in instances:
        create_bi_file(directory,instance,model,relaxed,cap,typeProba)
        print('running instance '+str(instance) )
        fileSol= directory+'/Resultats/instance_'+str(instance)+'_'+str(model)+'_'+str(relaxed)+'.xlsx'
       # print(fileSol)
        workbook = xlsxwriter.Workbook(fileSol)

        workbook.close()
        filemodel=directory+'/Modeles/model_'+str(instance)+'_'+str(model)+'_'+str(relaxed)+'.ltf'
        print(filemodel)
        os.system(f'runlingo {filemodel}')

    return 


def modif_proba(instance_debut, instance_fin, directory,probaInit):
    instances = [(i) for i in range(instance_debut,instance_fin)]

    for instance in instances:
        filename=directory+'/Instance'+str(instance)+'.xlsx'
        print(filename)
        write_proba(120, 12, 4, 0.3, filename)

instance_size=30
directory='C:/Users/baret/Documents/Simulateur/test-non-homogène/Instances-12-120-0,1'
#crea_instance(directory,instance_size, 50, 120, 10, 20, 10, 5, 12, 4, 10, 0.4, 0.01)
#ô(9,15,directory, 0.3)


print('test')
relaxed =False
model='C' # C ou CMS ou CML 
cap=4
#directory='C:/Users/baret/Documents/Simulateur/Instances-finales/Instances-24-120-0,3/NR/Budget/Budget-0,4'
resolution_cap(8,25,directory,model,False, cap,'Linear',0.3)
#resolution(instance_size,model,True, cap,'Linear')
#model2='Normal'
#resolution(0,30, model2, False,cap,'Linear')

