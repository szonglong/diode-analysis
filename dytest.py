import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt
import Tkinter
import tkFileDialog

############ Settings ##############
list_tol_ref = 5
tol_ref = 1e-1
area=4.29e-6
thickness = 100e-7
epsilon = 3*8.85e-12

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
    dev_list2 =  {}
    tol = tol_ref
    tol2 = tol_ref
    while len(dev_list)==0 or len(dev_list2)==0:
        for dev_index in range(1,len(dat2[sheet_name[1]])-4):
            dev = (dat2[sheet_name[1]][dev_index+1]-dat2[sheet_name[1]][dev_index-1])/(np.log10(np.abs(dat1['V1'][dev_index+1]))-(np.log10(np.abs(dat1['V1'][dev_index-1]))))
            charlie = dat2[sheet_name[1]][dev_index]-dev*np.log10(np.abs(dat1['V1'][dev_index]))
            if np.abs(2-dev)  < tol and dat1['V1'][dev_index] > 0: 
                dev_list[dat1['V1'][dev_index]]=[(dev,charlie,sheet_name[1])]
            if np.abs(2-dev)  < tol and dat1['V1'][dev_index] < 0: 
                dev_list2[dat1['V1'][dev_index]]=[(dev,charlie,sheet_name[1])]
        if len(dev_list)==0:
            tol *= 2
        if len(dev_list2)==0:
            tol2*= 2
    for i in dev_list:
        if i in big_dict:
            big_dict[i].append(dev_list[i])
        else:
            big_dict[i] =  [dev_list[i]]
    for i in dev_list2:
        if i in big_dict2:
            big_dict2[i].append(dev_list2[i])
        else:
            big_dict2[i] =  [dev_list2[i]]
    return


def sclc_plot(to_plot,name):
    plt.figure(1)
    log_array = semilog_array.copy()
    log_array.drop(labels='V1', axis='columns', inplace=True)
    log_array['log V1'] = np.log10(np.abs(dat1))
    min_to_plot= list(semilog_array['V1']).index(min(np.abs(to_plot)))
    max_to_plot=list(semilog_array['V1']).index(max(np.abs(to_plot)))
    plog_array = log_array.truncate(before=min_to_plot, after=max_to_plot)
    
    plog_array.plot(x='log V1', figsize=(9,9))
    plt.savefig('processed_semilog/ _%s.png' %(file+name))
    return 10**plog_array['log V1'][min_to_plot],10**plog_array['log V1'][max_to_plot]

def process_sclc(big_dict,name):
    c_list = []
    list_tol = list_tol_ref
    to_plot = []
    while to_plot == []:    
        for key in big_dict:
            if len(big_dict[key])>=list_tol:
                to_plot.append(key)
                for i in big_dict[key]:
                    c_list.append(i[0][1])
        list_tol -= 1        
    alpha = 10**np.average(c_list)
    (min_to_plot,max_to_plot) = sclc_plot(to_plot,name)
    d = {'SCLC minimum voltage (V)':min_to_plot, 'SCLC maximum voltage (V)':max_to_plot, 'Charge mobility (cm2/V s)':alpha*(8*thickness**3)/(1000*9*epsilon)}
    return d
            
for file in wlist:
    wb=pd.read_excel(r'%s.xls' %file, None)

    big_dict = {}
    big_dict2 = {}

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
    
    ana1=pd.DataFrame(process_sclc(big_dict,"f"),index=['Forward'])
    ana2=pd.DataFrame(process_sclc(big_dict2,"b"),index=['Backward'])
    ana_df=ana1.append(ana2)


############ Basic export ############
    plt.figure(2)
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

############ Export ############
    writer = pd.ExcelWriter('processed_semilog/%s _p.xlsx' %file, engine='xlsxwriter')
    semilog_array.to_excel(writer, sheet_name = 'Final array')
    
    
    ana_df.to_excel(writer, sheet_name = 'Analysis array')
    
    writer.save()
    plt.close()
