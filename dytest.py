import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt
#import Tkinter
#import tkFileDialog
import xlsxwriter

############ Settings ##############
list_tol_ref = 5 #number of pixels required under SCLC
tol_ref = 0.3 #how close the derivative of the log/log is to 2
area=4.29e-6 # in m2 I think
thickness = 98.5*1e-7  #in cm
epsilon = 3*8.85e-12 #n * e0
Vbi = 0.8 #default=0.1

############ File Search ############
#root = Tkinter.Tk()
#root.withdraw() #use to hide tkinter window

#currdir = os.getcwd()
#tempdir = tkFileDialog.askdirectory(parent=root, initialdir='C:\\Users\\E0004621\\Desktop\\Zong Long\\Papers\\Data', title='Please select the data folder') #select directory for data

tempdir = 'C:\\Users\\E0004621\\Desktop\\Pythontest' #Debugging use
#tempdir = 'C:\\Users\\E0004621\\Desktop\\Zong Long\\Papers\\Data\\180719' #Debugging use
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
def ana_var(sheet_name,dev_df):              #updates the 2 dictionaries with useful data , one for forward and one for reverse bias    
    tol = tol_ref
    working_list = []
    dev_list = []                   #init lists
    charlie_list = []
 
    for dev_index in range(1,len(dat2[sheet_name[1]])-2):  #make lists
        dev=(dat2[sheet_name[1]][dev_index+1]-dat2[sheet_name[1]][dev_index-1])/(np.log10(np.abs(dat1['V1'][dev_index+1]))-(np.log10(np.abs(dat1['V1'][dev_index-1])))) #using the central finite difference method to calculate derivative (error O(n2))
        dev_list.append(dev)
        charlie_list.append(dat2[sheet_name[1]][dev_index]-dev*np.log10(np.abs(dat1['V1'][dev_index])))
    
    working_dict=tuple(zip(semilog_array['V1'][1:len(dat2[sheet_name[1]])-4],zip(dev_list,charlie_list,[sheet_name[1]]*len(dev_list))))  #make woeking dict
    dev_df['Derivative %s /  log %s' % (sheet_name[1],sheet_name[0])] = dev_list # to make worksheet of derivatives

    while working_list == []:
        working_list=(list(filter(lambda x:np.abs(2-x)<tol,dev_list)))
        tol*=2

    for (k,v) in working_dict: #is not actually a dictionary but a tuple of size 2 with "key" and "value"
        if v[0] in working_list:
            if k>= -Vbi:
                if k in big_dict:
                    big_dict[k].append(v) #Positive sweep
                else:
                    big_dict[k] = [v]
            else:
                if k in big_dict2:
                    big_dict2[k].append(v)
                else:
                    big_dict2[k] = [v]
    return 

def sclc_plot(to_plot,name): #plots the forward and backward sweeps and returns the ranges
    plt.figure(1)
    plog_array = semilog_array.copy() #deep copy
    plog_array.drop(labels='V1', axis='columns', inplace=True)
    plog_array['V1'] = (np.abs(dat1))
    plog_array.plot(x='V1', figsize=(9,9),logx=True)
#    if len(to_plot) >1:
#        plt.xlim(np.log10(abs(min(to_plot))),np.log10(abs(max(to_plot))))
    plt.ylim(-3,2)
    plt.xlabel('log (V-Vbi)')
    plt.title('%s'%(file+name))
    plt.savefig('processed_semilog/ plotted_log_%s.png' %(file+name))
    plt.close()
    return

def process_sclc(big_dict,name): #calls sclc_plot and returns useful data in pd.df
    c_list = []
    list_tol = list_tol_ref
    to_plot = []
    while to_plot == []:    
        for k,v in big_dict.items():
            if len(v)>=list_tol:
                to_plot.append(k)
                for i in v:
                    c_list.append(i[1])
        list_tol -= 1        
    alpha = 10**np.average(c_list)
    sclc_plot(to_plot,name)
    to_plot=list(map(lambda  x: np.round(x,2),to_plot))
    to_plot.sort()
    d = {'Charge mobility (cm2/V s)':alpha*(8*thickness**3)/(1000*9*epsilon),'Number of pixels under SCLC': list_tol +1,'Data points plotted':[to_plot]}
    return d
big_ana_df = pd.DataFrame()
            
for file in wlist:
    wb=pd.read_excel(r'%s.xls' %file, None)
    big_dict = {}
    big_dict2 = {}

    for sheet_index in range(0,1) + range(3,len(wb.keys())):  
        sh=wb.values()[sheet_index]
        
        if sheet_index == 0:
            sheet_name = ['V1','log j1']            
            dat1=pd.DataFrame({'V1': sh.V1-Vbi})
            dat2=pd.DataFrame({'log j1': np.log10(np.abs(sh.I1/(10*area)))}) #log current density in mA/cm2
            dev_df= pd.DataFrame({'V1':dat1['V1'][1:len(dat2['log j1'])-2]})
            semilog_array = dat1.join(dat2)

            
        else:
            sheet_name = ['V%i' % (sheet_index-1),'log j%i' % (sheet_index-1)]            
            dat2=pd.DataFrame({'log j%i' % (sheet_index-1) : np.log10(np.abs(sh.I1/(10*area)))})
            semilog_array=semilog_array.join(dat2)
        ana_var(sheet_name,dev_df)
    
    ana1=pd.DataFrame(process_sclc(big_dict,"f"),index=['Forward'])
    ana2=pd.DataFrame(process_sclc(big_dict2,"b"),index=['Backward'])
    ana_df=ana1.append(ana2)
    big_ana_df=big_ana_df.append(ana_df)



############ Basic export ############
    plt.figure(2)
    semilog_array.plot(x='V1', figsize=(9,9))
    plt.title('%s'%(file))
    plt.grid(True)
    plt.xlabel('Applied voltage(V)')
    plt.ylabel('Log [Current density](mAcm$^-$$^2$)')
    plt.xlim(-8-Vbi,8+Vbi)
    plt.ylim(-3,4)
    font={'size':18}
    plt.rc('font',**font)
    plt.legend(loc='lower left', prop={'size':12})    
    plt.savefig('processed_semilog/%s _p.png' %file)

############ Export ############
    writer = pd.ExcelWriter('processed_semilog/%s _p.xlsx' %file, engine='xlsxwriter')
    semilog_array.to_excel(writer, sheet_name = 'Final array')

    ana_df.to_excel(writer, sheet_name = 'Analysis array')
    dev_df.to_excel(writer, sheet_name = 'Derivatives')
    workbook  = writer.book
    format3 = workbook.add_format({'bg_color':   '#C6EFCE',
                               'font_color': '#006100'})
    worksheet = writer.sheets['Derivatives']
    worksheet.conditional_format('C3:J322', {'type':'formula','criteria': '=ABS(C3-2)<%f'%tol_ref,'format':format3})

    writer.save()
    plt.close()
