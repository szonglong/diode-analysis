import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt
import Tkinter
import tkFileDialog

############ Settings ##############
jcol=2 #Column to take current. First column is 0
vcol=1 #column to take voltage. Usually @ column 

plotxrange=[-8,8]
plotyrange=[10**-5,10**4]

m = 1 #mismatch factor; def = 1
d = -1 #direction of connectors' def = 1, acceptable args = +-1

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

for file in wlist:
    wb=pd.read_excel(r'%s.xls' %file, None)
    ana_array = pd.DataFrame(index=['Jsc (mA/cm2)','Voc (V)','I_inj (mA)','Pmax (W)','Vmpp (V)','Jmpp (mA/cm2)','Rs','PCE (%)','FF (%)'])
    
    for sheet_index in range(0,1) + range(3,len(wb.keys())):  
        sh=wb.values()[sheet_index]
        
        if sheet_index == 0:
            sheet_name = ['p1','j1']            
            dat1=pd.DataFrame({'V1': sh.V1})
            dat2=pd.DataFrame({'I1': sh.I1/(10*area)*d}) #current density in mA/cm2
            final_array = dat1.join(np.abs(dat2))

            power_array = dat1
            quarter_length=int(len(power_array)/4)
            
        else:
            sheet_name = ['p%i' % (sheet_index-1),'j%i' % (sheet_index-1)]            
            dat2=pd.DataFrame({'j%i' % (sheet_index-1) : sh.I1/(10*area)*d})
            final_array=final_array.join(np.abs(dat2))

    #Final array has columns Vs and J1~J8
    final_array.plot(x="V1", figsize=(9,9), logy=True) #Semi-log plot

############ Plotting settings ############
    plt.title('%s'%(file))
    plt.grid(True)
    plt.xlabel('Applied voltage(V)')
    plt.ylabel('Log [Current density](mAcm$^-$$^2$)')
    plt.xlim(plotxrange)
    plt.ylim(plotyrange)
    font={'size':18}
    plt.rc('font',**font)
    plt.legend(loc='lower left', prop={"size":12})


############ Export ############
    writer = pd.ExcelWriter('processed_semilog/%s _p.xlsx' %file, engine='xlsxwriter')
    final_array.to_excel(writer, sheet_name = 'Final array')
    writer.save()
    
    plt.savefig('processed_semilog/%s _p.png' %file)
    plt.close()
