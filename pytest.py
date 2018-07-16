import pandas as pd
import os,sys
import numpy as np
import matplotlib.pyplot as plt  #

############ Settings ##############
jcol=2 #Column to take current. First column is 0
vcol=1 #column to take voltage. Usually @ column 

plotxrange=[-3,3]
plotyrange=[10**-6,10**4]

############ File Search ############
filelist = os.listdir(os.getcwd())  # working dir

if os.path.exists('processed_semilog')!=True: # make file
    os.mkdir('processed_semilog')
    
wlist=[] # get working list, only xls files
for i in xrange(len(filelist)):
    ext = os.path.splitext(filelist[i])[1]
    filename = os.path.splitext(filelist[i])[0]
    if ext  in ('.xls','.XLS'):
        wlist.append(os.path.splitext(filelist[i])[0])
print wlist


############ Processing data ############
for file in wlist:
    wb=pd.read_excel(r'%s.xls' %file, None)
    
    for sheet_index in range(0,1) + range(3,len(wb.keys())):  
        sh=wb.values()[sheet_index]
        if sheet_index == 0:
            dat1=pd.DataFrame({'Vs': sh.Vs})
            dat2=pd.DataFrame({'j1': sh.Id})
            dat2 *=1e3/0.0429
            final_array = dat1.join(np.abs(dat2))
            
        else:
            dat2=pd.DataFrame({'j%i' % (sheet_index-1) : sh.Id})
            dat2 *=1e3/0.0429
            final_array=final_array.join(np.abs(dat2))
    
    #Final array has columns Vs and J1~J8
    final_array.plot(x="Vs", figsize=(20,10), logy=True) #Semi-log plot

    plt.title('%s'%(file))
    plt.grid(True)
    plt.xlabel('Applied voltage(V)')
    plt.ylabel('Log [Current density](mAcm$^-$$^2$)')
    plt.xlim(plotxrange)
    plt.ylim(plotyrange)

    final_array.to_excel('processed_semilog/%s _p.xls' %file)
    
    plt.savefig('processed_semilog/%s _p.png' %file)
    plt.close()