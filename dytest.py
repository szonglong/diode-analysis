import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt
import Tkinter
import tkFileDialog

############ Settings ##############
list_tol = 5

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
            dev = (dat2[sheet_name[1]][dev_index+1]-dat2[sheet_name[1]][dev_index-1])/(np.log10(np.abs(dat1['V1'][dev_index+1]))-(np.log10(np.abs(dat1['V1'][dev_index-1]))))
            if np.abs(2-dev)  < tol and dat1['V1'][dev_index] > 0: #note: this version takes the positive sweep only
                dev_list[dat1['V1'][dev_index]]=[dev]
        tol*= 2
    for i in dev_list:
        if i in big_dict:
            big_dict[i].append(dev_list[i])
            
        else:
            big_dict[i] =  dev_list[i]


def ana_sclc():
    plt.figure(2)
    log_array = semilog_array
    log_array.drop(labels='V1', axis='columns', inplace=True)
    log_array['log V1'] = np.log10(np.abs(dat1))
    plog_array = log_array.truncate(before=1, after=len(log_array)/2)
    plog_array.plot(x='log V1', figsize=(9,9))
    plt.xlim(np.log10(min(np.abs(to_plot))),np.log10(max(np.abs(to_plot))))
    plt.show()
            
for file in wlist:
    wb=pd.read_excel(r'%s.xls' %file, None)
    big_dict = {}
    to_plot = []
    for sheet_index in range(0,1) + range(3,len(wb.keys())):  
        sh=wb.values()[sheet_index]
        
        if sheet_index == 0:
            sheet_name = ['V1','log j1']            
            dat1=pd.DataFrame({'V1': sh.V1})
            dat2=pd.DataFrame({'log j1': np.log10(np.abs(sh.I1/(10*area)))}) #log current density in mA/cm2
            semilog_array = dat1.join(dat2)

            
        else:
            sheet_name = ['V%i' % (sheet_index-1),'log j%i' % (sheet_index-1)]            
            dat2=pd.DataFrame({'log j%i' % (sheet_index-1) : np.log10(np.abs(sh.I1/(10*area)))})
            semilog_array=semilog_array.join(dat2)
    
        ana_var(sheet_name)
    for key in big_dict:
        if len(big_dict[key])>=list_tol:
            to_plot.append(key)

############ Basic export ############
    plt.figure(1)
    semilog_array.plot(x='V1', figsize=(9,9))
    plt.title('%s'%(file))
    plt.grid(True)
    plt.xlabel('Applied voltage(V)')
    plt.ylabel('Log [Current density](mAcm$^-$$^2$)')
    plt.xlim(-8,8)
    plt.ylim(-3,4)
    font={'size':18}
    plt.rc('font',**font)
    plt.legend(loc='lower left', prop={'size':12})    
    plt.savefig('processed_semilog/%s _p.png' %file)

    writer = pd.ExcelWriter('processed_semilog/%s _p.xlsx' %file, engine='xlsxwriter')
    semilog_array.to_excel(writer, sheet_name = 'Final array')
############ Plotting settings ############
    to_plot.sort()
    print 'The range of SCLC is : ' + str(to_plot)
    ana_sclc()
#    cont_counter = raw_input('Continue to analyse data? (y/n)')
#    if cont_counter == 'y':
#        log_array = ana_sclc()
#        #run program
#    elif cont_counter == 'n':
#        print 'no'
#    else:
#        print 'poop'
    

############ Export ############

    
    
    writer.save()
    plt.close()
