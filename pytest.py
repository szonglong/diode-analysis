import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt
import Tkinter
import tkFileDialog

############ Settings ##############
plotxrange=[-3,3]
plotyrange=[10**-6,10**4]
plotyrange2 = [0,20]

m = 1 #mismatch factor; def = 1
d = -1 #direction of connectors' def = 1, acceptable args = +-1


area=4.29e-6
############ File Search ############
root = Tkinter.Tk()
root.withdraw() #use to hide tkinter window

currdir = os.getcwd()
tempdir = tkFileDialog.askdirectory(parent=root, initialdir='C:\\Users\\E0004621\\Desktop\\Zong Long\\Papers\\Data', title='Please select the data folder') #select directory for data
#tempdir = 'C:\\Users\\E0004621\\Desktop\\Pythontest' #Debugging use
#tempdir = 'C:\\Users\\E0004621\\Desktop\\Zong Long\\Papers\\Data\\180711' #Debugging use
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

    #Finds P_max - only the first quadrant
    power_array[sheet_name[0]]=dat1.Vs*dat2[sheet_name[1]]*(10*area)
    power_array2 = power_array.truncate(after=quarter_length)
    max_index2=power_array2[sheet_name[0]].idxmax()
    Pmax2 = power_array2[sheet_name[0]][max_index2]
    Pmax = Pmax2
    max_index = max_index2

    #Finds Voc
    voc_index = 0
    for vi in range(len(dat2[sheet_name[1]])):
        if dat2[sheet_name[1]][vi] < 0:
            voc_index = vi
            break
    if voc_index == 0:
        Voc=-1 #invalid pixel: check d or pixel is bad
        Rs=9e9   
    else:
        Voc = (dat1['Vs'][voc_index-1]*dat2[sheet_name[1]][voc_index]-dat1['Vs'][voc_index]*dat2[sheet_name[1]][voc_index-1])/(dat2[sheet_name[1]][voc_index]-dat2[sheet_name[1]][voc_index-1])
        Rs = abs((dat1['Vs'][voc_index]-dat1['Vs'][voc_index-1])/(dat2[sheet_name[1]][voc_index]-dat2[sheet_name[1]][voc_index-1]))*1000 #Ohm cm2
    

    Iinj = dat2[sheet_name[1]][quarter_length]*(10*area)*-1000 if len(dat2[sheet_name[1]])>quarter_length else None
    Vmpp = dat1.Vs[max_index]
    Jmpp = dat2[sheet_name[1]][max_index]
    Jsc = dat2[sheet_name[1]][0]
    PCE = Pmax/(1000*area/m)*100
    FF = Jmpp*Vmpp/(Jsc*Voc)*100

    ana_array[sheet_name[0]] = [Jsc,Voc,Iinj,Rs,PCE,FF] #Arranges in an array to put in excel
    
    
    
def ana_var_inv(sheet_name):                #analyse variable  

    trigger_OOI = 1
    #Finds P_max - only the first quadrant
    power_array[sheet_name[0]]=dat1.Vs*dat2[sheet_name[1]]*(10*area)
    power_array2 = power_array.truncate(before=half_length)
#    print(power_array2)
    max_index2=power_array2[sheet_name[0]].idxmax()
    if type(max_index2) != int:
        max_index = 0
        Pmax2 = 0
        Voc = 99
        Rs = 9e9
        trigger_OOI = 0

    else:
        Pmax2 = power_array2[sheet_name[0]][max_index2]
        
        max_index = max_index2
    Pmax = Pmax2

    #Finds Voc via interpolation
    voc_index = 0
    if trigger_OOI != 0:
        for vi in range(len(dat2[sheet_name[1]])):
            if dat2[sheet_name[1]][vi] > 0:
                voc_index = vi
                break
        if voc_index == 0:
            Voc=-1 #invalid pixel: check d or pixel is bad
            Rs=9e9   
        else:
            Voc = (dat1['Vs'][voc_index-1]*dat2[sheet_name[1]][voc_index]-dat1['Vs'][voc_index]*dat2[sheet_name[1]][voc_index-1])/(dat2[sheet_name[1]][voc_index]-dat2[sheet_name[1]][voc_index-1])
            Rs = abs((dat1['Vs'][voc_index]-dat1['Vs'][voc_index-1])/(dat2[sheet_name[1]][voc_index]-dat2[sheet_name[1]][voc_index-1]))*1000 #Ohm cm2
    

    Iinj = dat2[sheet_name[1]][quarter_length]*(10*area)*-1000 if len(dat2[sheet_name[1]])>quarter_length else None
    Vmpp = dat1.Vs[max_index]
    Jmpp = dat2[sheet_name[1]][max_index]
    Jsc = dat2[sheet_name[1]][0]
    PCE = Pmax/(1000*area/m)*100
    FF = Jmpp*Vmpp/(Jsc*Voc)*100

    ana_array[sheet_name[0]] = [Jsc,Voc,Iinj,Rs,PCE,FF] #Arranges in an array to put in excel
    
def smooth(iterable,repeats):
    for t in range(repeats):
        for i in range(len(iterable)):
            if i in [0,1,len(iterable)-2,len(iterable)-1]:
                pass
            else:
               iterable[i] = (0.25*(iterable[i-2]+iterable[i+2])+0.5*(iterable[i+1]+iterable[i-1])+iterable[i])/2.5
    return iterable

for file in wlist:
    wb=pd.read_excel(r'%s.xls' %file, None)
    ana_array = pd.DataFrame(index=['Jsc (mA/cm2)','Voc (V)','I_inj (mA)','Rs (Ohm cm2)','PCE (%)','FF (%)'])
    
    for sheet_index in range(0,1) + range(3,len(wb.keys())):  
        sh=wb.values()[sheet_index]
        
        if sheet_index == 0:
            sheet_name = ['p1','j1']            
            dat1=pd.DataFrame({'Vs': sh.Vs})
            dat2=pd.DataFrame({'j1': smooth(sh.Is,3)/(10*area)*d}) if 'd' in str(file) else pd.DataFrame({'j1': sh.Is/(10*area)*d}) #current density in mA/cm2
            final_array = dat1.join((dat2))
            plot_array = dat1.join(np.abs(dat2))

            power_array = dat1.copy()
            quarter_length=int(len(power_array)/4)
            half_length=int(quarter_length*2)
            
        else:
            sheet_name = ['p%i' % (sheet_index-1),'j%i' % (sheet_index-1)]            
            dat2=pd.DataFrame({'j%i' % (sheet_index-1): smooth(sh.Is,3)/(10*area)*d}) if 'd' in str(file) else pd.DataFrame({'j%i' % (sheet_index-1): sh.Is/(10*area)*d}) #current density in mA/cm2

            final_array=final_array.join(dat2)
            plot_array = plot_array.join(np.abs(dat2))

        if 'bi' in str(file):
            ana_var_inv(sheet_name)
        elif 'b' in str(file):
            ana_var(sheet_name)
        
    #Final array has columns Vs and J1~J8
    fig, ax = plt.subplots()
#    ax2 = ax.twinx()
    plot_array.plot(x="Vs", ax=ax, ylim=plotyrange, figsize=(9,9), logy=True)#, ls='--') #Semi-log plot
#    plot_array.plot(x="Vs", ax=ax2, ls=':', ylim=plotyrange2)
############ Plotting settings ############
    plt.title('%s'%(file))
    plt.grid(True)
    plt.xlabel('Applied voltage(V)')
    ax.set_ylabel('Log [Current density](mAcm$^-$$^2$)')
#    ax2.set_ylabel('Current density(mAcm$^-$$^2$)')
    plt.xlim(plotxrange)
    font={'size':18}
    plt.yticks(np.geomspace(10**-6, 10**4, num=11),(np.geomspace(10**-6, 10**4, num=11).astype(float)))
    plt.rc('font',**font)
    plt.legend(loc='lower left', prop={"size":12})

############ Export ############
    writer = pd.ExcelWriter('processed_semilog/%s _p.xlsx' %file, engine='xlsxwriter')
    final_array.to_excel(writer, sheet_name = 'Final array')
    if 'b' in str(file):
        ana_array.to_excel(writer, sheet_name = 'Analysis array')
    writer.save()
    
    plt.savefig('processed_semilog/%s _p.png' %file)
    plt.close()
