import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt
import Tkinter
import tkFileDialog

############ Settings ##############
jcol=2 #Column to take current. First column is 0
vcol=1 #column to take voltage. Usually @ column 

plotxrange=[-3,3]
plotyrange=[10**-6,10**4]

m = 1 #mismatch factor; def = 1
d = 1 #direction of connectors' def = 1, acceptable args = +-1

area=42.9e-6
############ File Search ############
root = Tkinter.Tk()
root.withdraw() #use to hide tkinter window

#currdir = os.getcwd()
#tempdir = tkFileDialog.askdirectory(parent=root, initialdir='C:\\Users\\E0004621\\Desktop\\Zong Long\\Papers', title='Please select the data folder') #select directory for data
#tempdir = 'C:\\Users\\E0004621\\Desktop\\Pythontest' #Debugging use
tempdir = 'C:\\Users\\E0004621\\Desktop\\Zong Long\\Papers\\Data\\180711' #Debugging use
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
def ana_var(sheet_name):
    quadrant_counter = 2

    power_array[sheet_name[0]]=dat1.Vs*dat2[sheet_name[1]]*area
    power_array2 = power_array.truncate(after=quarter_length); power_array3 = power_array.truncate(before=quarter_length+1, after=2*quarter_length)
    max_index2=power_array2[sheet_name[0]].idxmax(); max_index3=power_array3[sheet_name[0]].idxmax()
    
    try:
        Pmax2 = power_array2[sheet_name[0]][max_index2]
        Pmax3 = power_array3[sheet_name[0]][max_index3]
        if Pmax2 < Pmax3:
            Pmax = Pmax3
            max_index = max_index3
            quadrant_counter = 3
            print (Pmax3-Pmax2)/Pmax3
        else:
            Pmax = Pmax2
            max_index = max_index2
        
    except TypeError:
        Pmax = Pmax2
        max_index = max_index2

    for voc_index in range(len(dat2[sheet_name[1]])):
        Voc=-1 #invalid pixel: check d or pixel is shorted
        Rs=9e9   
        if quadrant_counter == 2:
            if dat2[sheet_name[1]][voc_index] < 0:
                if voc_index == 0:
                    break
                else:
                    Voc = (dat1['Vs'][voc_index-1]*dat2[sheet_name[1]][voc_index]-dat1['Vs'][voc_index]*dat2[sheet_name[1]][voc_index-1])/(dat2[sheet_name[1]][voc_index]-dat2[sheet_name[1]][voc_index-1])
                    Rs = abs((dat1['Vs'][voc_index]-dat1['Vs'][voc_index-1])/(dat2[sheet_name[1]][voc_index]-dat2[sheet_name[1]][voc_index-1]))*area
                    break
        else:
            voc_index += quarter_length
            if dat2[sheet_name[1]][voc_index] > 0:

                Voc = (dat1['Vs'][voc_index-1]*dat2[sheet_name[1]][voc_index]-dat1['Vs'][voc_index]*dat2[sheet_name[1]][voc_index-1])/(dat2[sheet_name[1]][voc_index]-dat2[sheet_name[1]][voc_index-1])
                Rs = abs((dat1['Vs'][voc_index]-dat1['Vs'][voc_index-1])/(dat2[sheet_name[1]][voc_index]-dat2[sheet_name[1]][voc_index-1]))*area
                break

    Vmpp = dat1.Vs[max_index]
    Jmpp = dat2[sheet_name[1]][max_index]
    try:
        Jsc = dat2[sheet_name[1]][0] if quadrant_counter == 2 else dat2[sheet_name[1]][quarter_length*2]
    except KeyError:
        Jsc = dat2[sheet_name[1]][0]
    PCE = Pmax/(1000*area/m)*100
    FF = Jmpp*Vmpp/(Jsc*Voc)*100

    ana_array[sheet_name[0]] = [Jsc,Voc,Pmax,Vmpp,Jmpp,Rs,PCE,FF]
    
for file in wlist:
    wb=pd.read_excel(r'%s.xls' %file, None)
    ana_array = pd.DataFrame(index=['Jsc','Voc','Pmax','Vmpp','Jmpp','Rs','PCE','FF'])
    
    for sheet_index in range(0,1) + range(3,len(wb.keys())):  
        sh=wb.values()[sheet_index]
        
        if sheet_index == 0:
            sheet_name = ['p1','j1']            
            dat1=pd.DataFrame({'Vs': sh.Vs})
            dat2=pd.DataFrame({'j1': sh.Id/area*d})
            final_array = dat1.join(np.abs(dat2))

            power_array = dat1
            quarter_length=int(len(power_array)/4)
            
        else:
            sheet_name = ['p%i' % (sheet_index-1),'j%i' % (sheet_index-1)]            
            dat2=pd.DataFrame({'j%i' % (sheet_index-1) : sh.Id/area*d})
            final_array=final_array.join(np.abs(dat2))

        if 'b' in str(file):
            ana_var(sheet_name)
        
    #Final array has columns Vs and J1~J8
    final_array.plot(x="Vs", figsize=(9,9), logy=True) #Semi-log plot



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
    if 'b' in str(file):
        ana_array.to_excel(writer, sheet_name = 'Analysis array')
    writer.save()
    
    plt.savefig('processed_semilog/%s _p.png' %file)
    plt.close()
