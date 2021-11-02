

import os
import sys
import pandas as pd
import openpyxl


forces_file = os.path.join(os.getcwd()+ os.sep,"postProcessing" + os.sep ,"forces" +os.sep,"0" +os.sep,"moment.dat")

if not os.path.isfile(forces_file):
        print("Forces file not found at "+forces_file)
        print( "Be sure that the case has been run and you have the right directory!")
        print ("Exiting.")
        sys.exit()

time = []
momentx = []
momenty = []
momentz = []

with open(forces_file,"r") as datafile:
        for line in datafile:
                if line[0] == "#":
                        continue
                data=line.split(' ')
                time += [float(data[0])]
                try:
                        result = line[line.find('(')+1:line.find(')')]
                       
                        result=list(result.split(" "))
                        momentx.append(result[0])
                        momenty.append(result[1])
                        momentz.append(result[2])
                        
                except:
                        print("no file")
                        continue
                
datafile.close()

d ={ 'Time' : pd.Series(time,dtype=float),'x moment' : pd.Series(momentx,dtype=float),'y moment' : pd.Series(momenty,dtype=float),'z moment' : pd.Series(momentz,dtype=float)}
df= pd.DataFrame(d)
#print(df)
df.to_excel('moment.xlsx', sheet_name='new_sheet_name')
