import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt
#from scipy.signal import savgol_filter
import Tkinter
import tkFileDialog

############ Settings ##############
plotxrange=[0,1.5] #for ideality graph
plotyrange=[0,3]
plotxrange2=[0,1.5]
plotyrange2=[10**-6,10**4]

T=300 #temperature in K
m = 1 #mismatch factor; def = 1
d = 1 #direction of connectors' def = 1, acceptable args = +-1

area=4.29e-6
############ File Search ############
root = Tkinter.Tk()
root.withdraw() #use to hide tkinter window
#
currdir = os.getcwd()
tempdir = tkFileDialog.askdirectory(parent=root, initialdir='C:\\Users\\E0004621\\Desktop\\Zong Long\\Papers\\Data', title='Please select the data folder') #select directory for data
#tempdir = 'C:\\Users\\E0004621\\Desktop\\Pythontest' #Debugging use
#tempdir = 'C:\\Users\\E0004621\\Desktop\\Zong Long\\Papers\\Data\\180711' #Debugging use
os.chdir(tempdir)

filelist = os.listdir(os.getcwd())  # working dir

if os.path.exists('processed_ideality')!=True: # make file
    os.mkdir('processed_ideality')

wlist=[] # get working list, only xls files
for i in xrange(len(filelist)):
    ext = os.path.splitext(filelist[i])[1]
    filename = os.path.splitext(filelist[i])[0]
    if ext  in ('.xls','.XLS') and 'd' in filename:
        wlist.append(os.path.splitext(filelist[i])[0])
print wlist    #is shown on command prompt dialogue

############ Processing data ############
def ana_var(sheet_name):
    j = np.array([0]*len(dat1.Vs),dtype=float)
    for i in range(len(dat2['j%s'%sheet_name])):
        j[i] = dat2['j%s'%sheet_name][i]
    j_hat=smooth(j,4)   #self designed running mean smoothing
    dev = np.gradient(np.array(dat1.Vs),j_hat)
    m0 = np.multiply(np.multiply(j, dev),1/(0.0259*T/300))
    m_hat = smooth(m0,1)
    dev_df['m%s' %sheet_name] = m_hat
    smoothed_df['smoothed j%s' %sheet_name] = np.abs(j)
    return

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

    for sheet_index in range(0,1) + range(3,len(wb.keys())):
        sh=wb.values()[sheet_index]

        if sheet_index == 0:
            sheet_name = '1'
            dat1=pd.DataFrame({'Vs': sh.Vs[0:len(sh.Vs)/4]})
            dev_df = dat1.copy()
            smoothed_df = dev_df.copy()
            dat2=pd.DataFrame({'j1': sh.Id[0:len(sh.Vs)/4]/(10*area)*d}) # ln current density in mA/cm2
            plot_array = dat1.join(np.abs(dat2))

        else:
            sheet_name = '%i' % (sheet_index-1)
            dat2=pd.DataFrame({'j%i' % (sheet_index-1) : sh.Id[0:len(sh.Vs)/4]/(10*area)*d})
            plot_array = plot_array.join(np.abs(dat2))

        m_list=ana_var(sheet_name)
############ Plotting settings ############

    plt.figure(1)    
    fig, ax = plt.subplots()
    ax2 = ax.twinx()
    dev_df.plot(x='Vs',ax=ax, ylim=plotyrange, figsize=(11,11))
    smoothed_df.plot(x='Vs',ax=ax2, ylim=plotyrange2, logy=True, ls='--')
    plot_array.plot(x='Vs',ax=ax2, ylim=plotyrange2, logy=True, ls=':')
    plt.xlim(plotxrange)
    plt.xlabel('Applied voltage(V)')
    ax.set_ylabel('Local Ideality factor')
    ax2.set_ylabel('Current density(mAcm$^-$$^2$)')
    ax.legend(loc=0)
    ax2.legend(loc='lower right')
    font={'size':18}
    plt.rc('font',**font)
    plt.savefig('processed_ideality/%s _m.png' %file)
    plt.close()