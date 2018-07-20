import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt
import Tkinter
import tkFileDialog

############ Settings ##############
plotxrange=[-1.2,1]
plotyrange=[-3,4]

d = -1 #to look into

area=4.29e-6
############ File Search ############
root = Tkinter.Tk()
root.withdraw() #use to hide tkinter window

#currdir = os.getcwd()
#tempdir = tkFileDialog.askdirectory(parent=root, initialdir='C:\\Users\\E0004621\\Desktop\\Zong Long\\Papers\\Data', title='Please select the data folder') #select directory for data
#tempdir = 'C:\\Users\\E0004621\\Desktop\\Pythontest' #Debugging use
tempdir = 'C:\\Users\\E0004621\\Desktop\\Zong Long\\Papers\\Data\\180719' #Debugging use
os.chdir(tempdir)

filelist = os.listdir(os.getcwd())  # working dir

if os.path.exists('processed_semilog')!=True: # make file
    os.mkdir('processed_semilog')
    
wlist=[] # get working list, only xls files
for i in xrange(len(filelist)):
    ext = os.path.splitext(filelist[i])[1]
    filename = os.path.splitext(filelist[i])[0]
    if ext  in ('.xls','.XLS'):
        wlist.append(os.path.splitext(filelist[i])[0])
print wlist    #is shown on command prompt dialogue

############ Processing data ############
def ana_var(sheet_name):                #analyse variable  
    dev_list = {}
    tol = 1e-1
    while len(dev_list)==0:
        for dev_index in range(1,len(dat2[sheet_name[1]])-4):
            dev = (dat2[sheet_name[1]][dev_index+1]-dat2[sheet_name[1]][dev_index-1])/(dat1['log V1'][dev_index+1]-dat1['log V1'][dev_index-1])
            if np.abs(2-dev)  < tol:
                dev_list[10**dat1['log V1'][dev_index]]=[dev]
        tol*= 2
    for i in dev_list:
        if i in big_dict:
            big_dict[i].append(dev_list[i])
            
        else:
            big_dict[i] =  dev_list[i]
            print sheet_name

            
for file in wlist:
    wb=pd.read_excel(r'%s.xls' %file, None)
    big_dict = {}
    for sheet_index in range(0,1) + range(3,len(wb.keys())):  
        sh=wb.values()[sheet_index]
        
        if sheet_index == 0:
            sheet_name = ['log V1','log j1']            
            dat1=pd.DataFrame({'log V1': np.log10(np.abs(sh.V1))})
            dat2=pd.DataFrame({'log j1': np.log10(np.abs(sh.I1/(10*area)))}) #current density in mA/cm2
            final_array = dat1.join(dat2)

            
        else:
            sheet_name = ['log V%i' % (sheet_index-1),'log j%i' % (sheet_index-1)]            
            dat2=pd.DataFrame({'log j%i' % (sheet_index-1) : np.log10(np.abs(sh.I1/(10*area)))})
            final_array=final_array.join(dat2)
    
        ana_var(sheet_name)

############ Plotting settings ############


############ Export ############
    writer = pd.ExcelWriter('processed_semilog/%s _p.xlsx' %file, engine='xlsxwriter')
    final_array.to_excel(writer, sheet_name = 'Final array')
    writer.save()
    
    plt.savefig('processed_semilog/%s _p.png' %file)
    plt.close()
