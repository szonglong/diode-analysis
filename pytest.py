import pandas as pd
import os,sys
import numpy as np
import matplotlib.pyplot as plt
import Tkinter
import tkFileDialog

############ Settings ##############
jcol=2 #Column to take current. First column is 0
vcol=1 #column to take voltage. Usually @ column 

plotxrange=[-3,3]
plotyrange=[10**-6,10**4]

m = 1 #mismatch factor
############ File Search ############
root = Tkinter.Tk()
root.withdraw() #use to hide tkinter window

#currdir = os.getcwd()
#tempdir = tkFileDialog.askdirectory(parent=root, initialdir='C:\\Users\\E0004621\\Desktop\\Zong Long\\Papers', title='Please select the data folder') #select directory for data
tempdir = 'C:\\Users\\E0004621\\Desktop\\Pythontest' #Debugging use
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
    ana_array = pd.DataFrame(index=['Pmax','Vmpp','Jmpp','Jsc','Voc','Rs','PCE','FF'])
    quarter_length = 0
    for sheet_index in range(0,1) + range(3,len(wb.keys())):  
        sh=wb.values()[sheet_index]
        
        if sheet_index == 0:
            dat1=pd.DataFrame({'Vs': sh.Vs})
            dat2=pd.DataFrame({'j1': sh.Id})
            dat2 *=1e3/0.0429
            final_array = dat1.join(np.abs(dat2))

            for voc_index in range(len(dat2['j1'])):
                if dat2['j1'][voc_index] < 0:
                    Voc = (dat1['Vs'][voc_index-1]*dat2['j1'][voc_index]-dat1['Vs'][voc_index]*dat2['j1'][voc_index-1])/(dat2['j1'][voc_index]-dat2['j1'][voc_index-1])
                    Rs = abs((dat1['Vs'][voc_index]-dat1['Vs'][voc_index-1])/(dat2['j1'][voc_index]-dat2['j1'][voc_index-1]))
                    break

            power_array = dat1
            power_array['p1']=dat1.Vs*dat2.j1
            quarter_length=len(power_array)/4
            power_array2= power_array.truncate(after=quarter_length)

            max_index=power_array2['p1'].idxmax()
            Pmax = power_array2['p1'][max_index]
            Vmpp = dat1.Vs[max_index]
            Impp = dat2.j1[max_index]
            Isc = dat2.j1[0]
            PCE = power_array2['p1'][max_index]/(1000*0.0429/m)
            FF = Pmax/(Isc*Voc)
            
            ana_array['p1'] = [Pmax,Vmpp,Impp,Isc,Voc,Rs,PCE,FF]
            
            
        else:
            dat2=pd.DataFrame({'j%i' % (sheet_index-1) : sh.Id})
            dat2 *=1e3/0.0429
            final_array=final_array.join(np.abs(dat2))
            
            for voc_index in range(len(dat2['j%i' % (sheet_index-1)])):
                if dat2['j%i' % (sheet_index-1)][voc_index] < 0:
                    Voc = (dat1['Vs'][voc_index-1]*dat2['j%i' % (sheet_index-1)][voc_index]-dat1['Vs'][voc_index]*dat2['j%i' % (sheet_index-1)][voc_index-1])/(dat2['j%i' % (sheet_index-1)][voc_index]-dat2['j%i' % (sheet_index-1)][voc_index-1])
                    Rs = abs((dat1['Vs'][voc_index]-dat1['Vs'][voc_index-1])/(dat2['j%i' % (sheet_index-1)][voc_index]-dat2['j%i' % (sheet_index-1)][voc_index-1]))
                    break
            
            
            power_array['p%i' % (sheet_index-1)]=dat1.Vs*dat2['j%i' % (sheet_index-1)]
            power_array2 = power_array.truncate(after=quarter_length)
            
            max_index=power_array2['p%i' % (sheet_index-1)].idxmax()
            Pmax = power_array2['p%i' % (sheet_index-1)][max_index]
            Vmpp = dat1.Vs[max_index]
            Impp = dat2['j%i' % (sheet_index-1)][max_index]
            Isc = dat2['j%i' % (sheet_index-1)][0]
            PCE = power_array2['p%i' % (sheet_index-1)][max_index]/(1000*0.0429/m)
            FF = Pmax/(Isc*Voc)

            ana_array['p%i' % (sheet_index-1)] = [Pmax,Vmpp,Impp,Isc,Voc,Rs,PCE,FF]

    #Final array has columns Vs and J1~J8
    final_array.plot(x="Vs", figsize=(6,6), logy=True) #Semi-log plot


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
    ana_array.to_excel(writer, sheet_name = 'Analysis array')
    writer.save()
    
    plt.savefig('processed_semilog/%s _p.png' %file)
    plt.close()
