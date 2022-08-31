import os

def create_lingo_ltf_file(directory, instance_number,model, relaxed,cap,typeProba, budg):
    
   instance_folder_path = directory
   instance = f'/Instances/Instance{instance_number}.xlsx'
   
   lin_model = open(f'{directory}/Modeles/model_{instance_number}_{model}_{relaxed}.ltf','w')
   lin_model.writelines('set default\nset echoin 1\n\n')

   """
      # First data section
   """

   lin_model.writelines('MODEL:\n\n')
     
   """
      # Sets section
   """
   write_sets(instance_folder_path,instance,lin_model )

   """
      # Data section
   
   """
   lin_model.writelines(f'DATA:\n')
   write_data(instance_folder_path, instance, lin_model,model,typeProba)

   write_results_data(instance_folder_path,instance_number,lin_model,relaxed,model)
   lin_model.writelines(f'ENDDATA\n\n')

   """
      # Objective function 
   """

   write_the_objective_function_init(lin_model)

   write_constraints_init(lin_model, budg)

   write_integrity_constraints(lin_model,relaxed)
   
   if model=='C' or model=='CMS':
       write_capacity_constraints(lin_model,cap)

    
   lin_model.writelines(f'END\n\n')

   """
     # Add command for runlingo
   """

   lin_model.writelines(f'set terseo 1\n')
  # lin_model.writelines(f'set timlim 60\n')
   lin_model.writelines(f'set MULTIS 3 \n ')
   lin_model.writelines(f'go\n')
   lin_model.writelines(f'nonz volume\n')
   lin_model.writelines(f'quit\n')
   
   lin_model.close()

   return 

def write_sets(instance_folder_path, instance,lingo_model):

   lingo_model.writelines(f'SETS:\n')

   lingo_model.writelines(f'Clients: demand, EC, penalty;\n')
   lingo_model.writelines(f'Facilities: Cap;\n')
   lingo_model.writelines(f'Levels;\n')
   lingo_model.writelines(f'SORTED(Clients,Facilities):sort, positions, P, distance;\n')
   lingo_model.writelines(f'LINKS(Facilities, levels): cost, proba,z;\n')

   lingo_model.writelines(f'ENDSETS\n\n')
   return 

def write_data(instance_folder_path, instance, lingo_model, model,typeProba):
 
   lingo_model.writelines(f'Number_clients=@ole(\'{instance_folder_path}{instance}\',\'NbClients\');\n')
   lingo_model.writelines(f'Number_facilities=@ole(\'{instance_folder_path}{instance}\',\'NbFacilities\');\n')  
   lingo_model.writelines(f'Number_levels=@ole(\'{instance_folder_path}{instance}\',\'NbLevels\');\n')  
   lingo_model.writelines( f'Clients=1..Number_clients;\n')
   lingo_model.writelines(f'Facilities=1..Number_facilities; \n')
   lingo_model.writelines(f'Levels=1..Number_levels; \n')
   lingo_model.writelines(f'BUD = @ole(\'{instance_folder_path}{instance}\',\'Budget\');\n')
   lingo_model.writelines(f'demand = @ole(\'{instance_folder_path}{instance}\',\'Demandes\');\n')
   lingo_model.writelines(f'penalty = @ole(\'{instance_folder_path}{instance}\',\'Penalites\');\n')
   lingo_model.writelines(f'distance = @ole(\'{instance_folder_path}{instance}\',\'Distances\');\n')
   lingo_model.writelines(f'cost = @ole(\'{instance_folder_path}{instance}\',\'Cout\');\n')
   
   if typeProba=='Linear':
       lingo_model.writelines(f'proba = @ole(\'{instance_folder_path}{instance}\',\'ProbaL\');\n')
   elif typeProba=='Convex':
       lingo_model.writelines(f'proba = @ole(\'{instance_folder_path}{instance}\',\'ProbaCO\');\n')
   elif typeProba=='Concave':
           lingo_model.writelines(f'proba = @ole(\'{instance_folder_path}{instance}\',\'ProbaCOCA\');\n')
   
   lingo_model.writelines(f'sort = @ole(\'{instance_folder_path}{instance}\',\'Tri\');\n')
   lingo_model.writelines(f'positions = @ole(\'{instance_folder_path}{instance}\',\'Positions\');\n')
 #  if model != 'CMS':
   lingo_model.writelines(f'Cap = @ole(\'{instance_folder_path}{instance}\',\'Capacites\');\n\n')   
   return


def write_results_data(instance_folder_path,instance_number,lingo_model,relaxed,model):

   results_folder_path = instance_folder_path+'/Resultats/'
   results = f'instance_{instance_number}_{model}_{relaxed}.xlsx'
   #results_folder_path = 'test'
   #results = '1'
  

   lingo_model.writelines('!Results;\n')
  
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'B2\')=@WRITE(\'Objectif\');\n')
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'C2\')=@WRITE(H1+H2);\n')

   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'B3\')=@WRITEFOR(LINKS(k,l):z(k,l));\n')
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'D2\')=@WRITE(\' The resolution time is : \',@TIME(),\' seconds\',@NEWLINE(1));\n\n')        
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'C3\')=@WRITEFOR(Facilities(k):Cap(k));\n')
   
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'D3\')=@WRITE(\'H1\');\n')
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'D4\')=@WRITE(H1);\n')
   
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'E3\')=@WRITE(\'H2\');\n')
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'E4\')=@WRITE(H2);\n')
 
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'F3\')=@WRITE(\'Proba1\');\n')
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'F4\')=@WRITE(Proba1);\n')
  
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'G3\')=@WRITE(\'Proba2\');\n')
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'G4\')=@WRITE(Proba2);\n')
   
   return 

    
def write_the_objective_function_init(lingo_model):

   lingo_model.writelines(f'!Objective function;\n')
   lingo_model.writelines(f'[obj] Min=H1+H2;\n\n')
   
   return 


def write_constraints_init(lingo_model, budg):
    lingo_model.writelines(f'H1=(@sum(Clients(i):demand(i)*@sum(SORTED(i,k1)|distance(i,k1) #LT# 5:@prod(SORTED(i,k2)|k2 #LT# k1:@sum(Levels(l1): proba(sort(i,k2),l1)*z(sort(i,k2),l1)))* @sum(Levels(l):(1-proba(sort(i,k1),l))*z(sort(i,k1),l))*distance(i,k1))));\n')
    lingo_model.writelines(f'H2=(Proba1/Proba2)*(@sum(Clients(i):@prod(SORTED(i,k3)|distance(i,k3) #LT# 5:@sum(Levels(l2):proba(sort(i,k3),l2)*z(sort(i,k3),l2)))*penalty(i)*demand(i)));\n')
    lingo_model.writelines(f'Proba1=@sum(Clients(i) : @sum(SORTED(i,k1)|distance(i,k1) #LT# 5:@prod(SORTED(i,k2)|k2 #LT# k1:proba(sort(i,k2),1))* (1-proba(sort(i,k1),1))));\n')
    lingo_model.writelines(f'Proba2=@sum(Clients(i):@prod(SORTED(i,k3)|distance(i,k3) #LT# 5:proba(sort(i,k3),1)));\n')
    lingo_model.writelines(f'BUDGET={budg}*BUD;\n')
 
    lingo_model.writelines(f'@sum(LINKS(j,l): cost(j,l)*z(j,l))<=BUDGET;\n')

    lingo_model.writelines(f'@for(Facilities(j):@sum(Levels(l): z(j,l))=1);\n\n')

    return 

def write_integrity_constraints(lingo_model,relax):

   lingo_model.writelines(f'! Integrity constraints;\n') 
   if relax==True:
       lingo_model.writelines(f'@for(Facilities(j): @for(Levels(l):z(j,l)>=0));\n') 
       lingo_model.writelines(f'@for(Facilities(j): @for(Levels(l):z(j,l)<=1));\n\n') 
   else :
       lingo_model.writelines(f'@for(Facilities(j): @for(Levels(l):@bin(z(j,l))));\n\n') 
       
   return 

def write_capacity_constraints(lingo_model,cap):
       
   lingo_model.writelines(f'! Capacity constraints;\n') 
   lingo_model.writelines(f'@for (Clients(i): @for( Facilities(k): P(i,k)=@prod(SORTED(i,k2)|k2 #LT# positions(i,k):@sum(Levels(l):proba(sort(i,k2),l)*z(sort(i,k2),l)))*@sum(Levels(l2):(1-proba(positions(i,k),l2))*z(positions(i,k),l2))));\n')
   lingo_model.writelines(f'@for(Facilities(j):@sum(Clients(i):P(i,j)*demand(i))<= {cap}*Cap(j));\n\n')
   
   return 
