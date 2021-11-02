

import os
import sys
import pandas as pd
import openpyxl


forces_file = os.path.join(os.getcwd()+ os.sep,"postProcessing" + os.sep ,"forces" +os.sep,"0" +os.sep,"force.dat")

if not os.path.isfile(forces_file):
        print("Forces file not found at "+forces_file)
        print( "Be sure that the case has been run and you have the right directory!")
        print ("Exiting.")
        sys.exit()

time = []
forcex = []
forcey = []
forcez = []

with open(forces_file,"r") as datafile:
        for line in datafile:
                if line[0] == "#":
                        continue
                data=line.split(' ')
                time += [float(data[0])]
                try:
                        result = line[line.find('(')+1:line.find(')')]
                       
                        result=list(result.split(" "))
                        forcex.append(result[0])
                        forcey.append(result[1])
                        forcez.append(result[2])
                        
                except:
                        print("no file")
                        continue
                
datafile.close()

d ={ 'Time' : pd.Series(time,dtype=float),'x force' : pd.Series(forcex,dtype=float),'y force' : pd.Series(forcey,dtype=float),'z force' : pd.Series(forcez,dtype=float)}
df= pd.DataFrame(d)
#print(df)
df.to_excel('force.xlsx', sheet_name='new_sheet_name')
