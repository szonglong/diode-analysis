import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt
import xlrd
import Tkinter
import tkFileDialog

############ Settings ##############
plotxrange=[-3,3]
plotyrange=[10**-6,10**4]   

m = 1 #mismatch factor; def = 1
d = -1 # depends on which I is chosen. Default d = -1

area=4.29e-6
############ File Search ############
root = Tkinter.Tk()
root.withdraw() #use to hide tkinter window

currdir = os.getcwd()
tempdir = tkFileDialog.askdirectory(parent=root, initialdir='C:\\Users\\E0004621\\Desktop\\Zong Long\\Papers\\Data\\4. Energy level alignment\\Solar Cell', title='Please select the data folder') #select directory for data
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

############ Function definitions for data processing ############
def ana_var(sheet_name,inv):                #analyse variable  

    max_index, Pmax = find_Pmax(dat1, dat2, sheet_name, inv)
    Voc, Rs = find_voc(dat2, sheet_name, inv)
    if inv == 0:
        if 'od' in str(file):
            Iinj = dat2[sheet_name[1]][half_length]*(10*area)*-1000 if len(dat2[sheet_name[1]])>half_length else None
        else:
            Iinj = dat2[sheet_name[1]][quarter_length]*(10*area)*-1000 if len(dat2[sheet_name[1]])>quarter_length else None
    elif inv == 1:
        Iinj = dat2[sheet_name[1]][half_length+quarter_length]*(10*area)*-1000 if len(dat2[sheet_name[1]])>half_length+quarter_length else None
    Vmpp = dat1.Vs[max_index]
    Jmpp = dat2[sheet_name[1]][max_index]
    Jsc = dat2[sheet_name[1]][0]
    PCE = Pmax/(1000*area/m)*100
    FF = Jmpp*Vmpp/(Jsc*Voc)*100

    ana_array[sheet_name[0]] = [Jsc,Voc,Iinj,Rs,PCE,FF] #Arranges in an array to put in excel

def find_voc(dat2,sheet_name,inv):
    #Finds Voc via interpolation
    voc_index = 0

    for vi in range(len(dat2[sheet_name[1]])):
        if inv == 0:
            if dat2[sheet_name[1]][vi] < 0:
                voc_index = vi
                break
        elif inv == 1:
            if dat2[sheet_name[1]][vi] > 0:
                voc_index = vi
                break
    if voc_index == 0:
        Voc=-99 #invalid pixel: check d or pixel is bad
        Rs=9e9   
    else:
        Voc = (dat1['Vs'][voc_index-1]*dat2[sheet_name[1]][voc_index]-dat1['Vs'][voc_index]*dat2[sheet_name[1]][voc_index-1])/(dat2[sheet_name[1]][voc_index]-dat2[sheet_name[1]][voc_index-1])
        Rs = abs((dat1['Vs'][voc_index]-dat1['Vs'][voc_index-1])/(dat2[sheet_name[1]][voc_index]-dat2[sheet_name[1]][voc_index-1]))*1000 #Ohm cm2
        
    return Voc, Rs
        
def find_Pmax(dat1, dat2, sheet_name, inv):
    #Finds P_max - only the first quadrant
    power_array[sheet_name[0]]=dat1.Vs*dat2[sheet_name[1]]*(10*area)
    if inv == 0:
        power_array2 = power_array.truncate(after=half_length) if 'od' in str(file) else power_array.truncate(after=quarter_length)
    elif inv == 1:
        power_array2 = power_array.truncate(before=half_length, after = half_length+quarter_length)
    max_index=power_array2[sheet_name[0]].idxmax()
    if type(max_index) != int:
        max_index = 0
        Pmax = 0
    else:
        Pmax = power_array2[sheet_name[0]][max_index]    
    return max_index, Pmax

#def smooth(iterable,repeats):           #Smoothing function to remove some jaggedness
#    for t in range(repeats):
#        for i in range(len(iterable)):
#            if i in [0,1,len(iterable)-2,len(iterable)-1]:
#                pass
#            else:
#               iterable[i] = (0.25*(iterable[i-2]+iterable[i+2])+0.5*(iterable[i+1]+iterable[i-1])+iterable[i])/2.5
#    return iterable


############ Run programme ############
for file in wlist:
    wb0 = xlrd.open_workbook(r'%s.xls' %file, logfile=open(os.devnull, 'w'))
    wb=pd.read_excel(wb0, None, engine='xlrd')
    ana_array = pd.DataFrame(index=['Jsc (mA/cm2)','Voc (V)','I_inj (mA)','Rs (Ohm cm2)','PCE (%)','FF (%)'])
    
    for sheet_index in range(0,1) + range(3,len(wb.keys())):  
        sh=wb.values()[sheet_index]

        if sheet_index == 0:
            sheet_name = ['p1','j1']
            dat1=pd.DataFrame({'Vs': sh.Vs})
#            dat2=pd.DataFrame({'j1': smooth(sh.Is,3)/(10*area)*d}) if 'd' in str(file) else pd.DataFrame({'j1': sh.Is/(10*area)*d}) #current density in mA/cm2
            dat2=pd.DataFrame({'j1': (sh.Is)/(10*area)*d})
            final_array = dat1.join((dat2))
            plot_array = dat1.join(np.abs(dat2))

            power_array = dat1.copy()
            quarter_length=int(len(power_array)/4)
            half_length=int(quarter_length*2)
            
        else:
            sheet_name = ['p%i' % (sheet_index-1),'j%i' % (sheet_index-1)]            
#            dat2=pd.DataFrame({'j%i' % (sheet_index-1): smooth(sh.Is,3)/(10*area)*d}) if 'd' in str(file) else pd.DataFrame({'j%i' % (sheet_index-1): sh.Is/(10*area)*d}) #current density in mA/cm2
            dat2=pd.DataFrame({'j%i' % (sheet_index-1): (sh.Is)/(10*area)*d})
            final_array=final_array.join(dat2)
            plot_array = plot_array.join(np.abs(dat2))

        if 'bi' in str(file):
            ana_var(sheet_name,inv=1)
        elif 'b' in str(file):
            ana_var(sheet_name,inv=0)
        
    #Final array has columns Vs and J1~J8
    fig, ax = plt.subplots()

    plot_array.plot(x="Vs", ax=ax, ylim=plotyrange, figsize=(9,9), logy=True)#, ls='--') #Semi-log plot


############ Plotting settings ############
    plt.title('%s'%(file))
    plt.grid(True)
    plt.xlabel('Applied voltage(V)')
    ax.set_ylabel('Log [Current density](mAcm$^-$$^2$)')
#    plotxrange= [0,1.3] if 'od' in str(file) else [-3,3]
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
    print("Completed.. " + str(file))