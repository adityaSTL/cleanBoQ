from tkinter import *
from sys import exit
from tkinter import filedialog
import pandas as pd
import numpy as np
import os
from extract import Extract
from create_table1 import Create_table
import openpyxl
import matplotlib.pyplot as plt
from openpyxl_image_loader import SheetImageLoader
import os
import os.path
import tkinter
import warnings
from array import *
import math
warnings.simplefilter(action='ignore', category=FutureWarning)
pd.options.mode.chained_assignment = None


file2=""
file1=""
popupRoot = Tk()

popupRoot.title('Generate BoQ')


def mix_mb(j):
    span_file=j
    columns_ot=['prikey', 'Key', 'Survey Duration', 'User', 'Upload Time','Survey Name', '_scheduled_start', 
                'Version', 'Complete','Survey Notes', 'Location Trigger', 'Instance Name', '_start', '_end',
                '_device', 'instanceid', 'Package_Name', 'Zone_Name', 'District_Name','Mandal_Name', 
                'From_GP', 'To_GP', 'Span_ID', 'Chainage_From','Chainage_To', 'Length', 'n1', 'Location', 
                'From_Location','From_Location_Latitude_Manual', 'From_Location_Longitude_Manual','Offset', 
                'Depth', 'Trench_Side', 'Duct_Laid', 'Method_Execution','Crossing', 'Strata_Type', 'Protection', 
                'Protect_Dwc', 'Protect_Gi','Protect_Rcc', 'Protect_Pcc', 'Rcc_Chamber', 'Rcc_Marker', 'Remark']
    df1_ot=pd.DataFrame(columns=columns_ot)
    column_hdd=['prikey', 'Key', 'Survey Duration', 'User', 'Upload Time',
        'Survey Name', '_scheduled_start', 'Version', 'Complete',
        'Survey Notes', 'Location Trigger', 'Instance Name', '_start', '_end',
        '_device', 'instanceid', 'Package_Name', 'Zone_Name', 'District_Name',
        'Mandal_Name', 'From_GP', 'To_GP', 'Span_ID','Side_of_the_road', 'Chainage_From',
        'Chainage_To', 'Length', 'Start_Point_Location','Start_Point' ,
        'Start_Point_Manual_Latitude', 'Start_Point_Manual_Longitude',
        'Graph_as_built_made', 'Depth', 'Bore_reamer_diameter', 'Ducts'
        , 'End_Point_Location','End_Point' ,
        'End_Point_Manual_Latitude', 'End_Point_Manual_Longitude', 'Remark']

    column_drt=['prikey', 'Key', 'Survey Duration', 'User', 'Upload Time',
        'Survey Name', '_scheduled_start', 'Version', 'Complete',
        'Survey Notes', 'Location Trigger', 'Instance Name', '_start', '_end',
        '_device', 'instanceid', 'Package_Name', 'Zone_Name', 'District_Name',
        'Mandal_Name', 'From_GP', 'To_GP', 'Span_ID', 'Chamber1_Location',
        'Chamber1_Location_Point', 'Chamber1_Manual_Latitude',
        'Chamber1_Manual_Longitude', 'Chamber1_condition',
        'Chamber1_route_marker', 'Chamber2_Location', 'Chamber2_location_Point',
        'Chamber2_Manual_Latitude', 'Chamber2_Manual_Longitude',
        'Chamber2_condition', 'Chambe2_route_marker', 'ch_from', 'ch_to',
        'Length', 'Duct_dam_punct_loc', 'Duct_dam_punct_loc_point',
        'Duct_dam_punct_loc_Manual_Latitude',
        'Duct_dam_punct_loc_Manual_Longitude', 'Duct_dam_punct_loc_ch_from',
        'Duct_dam_punct_loc_ch_to', 'Duct_dam_punct_loc_Length',
        'Duct_miss_loc', 'Duct_miss_loc_pt', 'Duct_miss_loc_Manual_Latitude',
        'Duct_miss_loc_Manual_Longitude', 'Duct_miss_ch_from',
        'Duct_miss_ch_to', 'Duct_miss_ch_Length', 'Remark']
    df1_hdd=pd.DataFrame(columns=column_hdd)
    df1_drt=pd.DataFrame(columns=column_drt)
    # df1_blo=pd.DataFrame(columns=column_blo)
    # import os
    # os.chdir('mandal')
    # for i in os.listdir():
    #j='WRG-ZAF-5706-M-01-GR01-02.xlsx'

    blo=Extract.extract_blo(j)
    ot=Extract.extract_ot(j)
    if len(ot)!=0:
        ot=Create_table.create_ot(ot)
        df1_ot=pd.concat([df1_ot,ot])
    hdd=Extract.extract_hdd(j)
    if len(hdd)!=0:
        hdd=Create_table.create_hdd(hdd)
        df1_hdd=pd.concat([df1_hdd,hdd])
    drt=Extract.extract_drt(j)
    if len(drt)!=0:
        drt=Create_table.create_drt(drt)
        df1_drt=pd.concat([df1_drt,drt])

    df1_ot=df1_ot[['Zone_Name','Mandal_Name','To_GP','Span_ID',
                    'Chainage_From','Chainage_To', 'Length',
                    'Offset','Depth', 'Trench_Side', 'Duct_Laid', 'Method_Execution','Crossing', 
                    'Strata_Type','Protect_Dwc', 'Protect_Gi','Protect_Rcc', 'Protect_Pcc', 'Rcc_Chamber',
                    'Rcc_Marker', 'Remark']]

    df1_hdd['Duct_Laid']=df1_hdd['Length']

    df1_hdd['Method_Execution']='HDD'

    df1_hdd[['Crossing','Offset','Strata_Type','Protect_Dwc', 'Protect_Gi','Protect_Rcc', 'Protect_Pcc', 'Rcc_Chamber',
            'Rcc_Marker']]=''
    df1_hdd.rename(columns={'Side_of_the_road':'Trench_Side'},inplace=True)

    df1_hdd=df1_hdd[['Zone_Name','Mandal_Name','To_GP','Span_ID',
                    'Chainage_From','Chainage_To', 'Length',
                    'Offset','Depth', 'Trench_Side', 'Duct_Laid', 'Method_Execution','Crossing', 
                    'Strata_Type','Protect_Dwc', 'Protect_Gi','Protect_Rcc', 'Protect_Pcc', 'Rcc_Chamber',
                    'Rcc_Marker', 'Remark']]

    # df1['Duct_miss_ch_from']

    df1_drt.reset_index(drop=True,inplace=True)
    df1_drt['Method_Execution']=''
    missing=[]
    for i in range(len(df1_drt)):
        if type(df1_drt.loc[i,'Duct_miss_ch_from'])==int and type(df1_drt.loc[i,'Duct_miss_ch_to'])==int:
            missing.append([df1_drt.loc[i,'Duct_miss_ch_from'],df1_drt.loc[i,'Duct_miss_ch_to'],i])
            df1_drt.loc[i,'Method_Execution']='Missing'
        else:
            df1_drt.loc[i,'Method_Execution']='DRT'


    df1=pd.concat([df1_ot,df1_hdd])
    df1.sort_values(by=['Chainage_From'],inplace=True)
    df1.reset_index(drop=True,inplace=True)
    dr_ind=[]
    for i in range(len(df1)):
        ch1=df1.loc[i,'Chainage_From']
        ch2=df1.loc[i,'Chainage_To']
        for j in missing:
            try:
                if ch1<j[0] or ch2>j[1]:
                    continue
            except:
                print('Some_problem in Chainage')
            try:
                if ch1>=j[0] and ch2<=j[1]:
                    df1.loc[i,'Method_Execution']='Missing'
                    dr_ind.append(j[2])
            except:
                print('Some prob in Chainage')

    df1_drt.drop(dr_ind,inplace=True)
    df1_drt.rename(columns={'ch_from':'Chainage_From','ch_to':'Chainage_To'},inplace=True)
    df1_drt['Duct_Laid']=df1_drt['Length']
    df1_drt[['Offset','Depth','Trench_Side','Crossing','Strata_Type','Protect_Dwc', 'Protect_Gi','Protect_Rcc', 'Protect_Pcc', 'Rcc_Chamber','Rcc_Marker']]=''

    df1_drt=df1_drt[['Zone_Name','Mandal_Name','To_GP','Span_ID',
                    'Chainage_From','Chainage_To', 'Length',
                    'Offset','Depth', 'Trench_Side', 'Duct_Laid', 'Method_Execution','Crossing', 
                    'Strata_Type','Protect_Dwc', 'Protect_Gi','Protect_Rcc', 'Protect_Pcc', 'Rcc_Chamber',
                    'Rcc_Marker', 'Remark']]
    print(len(df1))
    print(len(df1_drt))
    df1=pd.concat([df1,df1_drt])
    print(len(df1))

    df1.sort_values(by=['Chainage_From'],inplace=True)

    # df1.to_csv('xx.csv')
    temp=0
    miss_df=df1[df1['Method_Execution']=='Missing']
    miss_df.reset_index(drop=True,inplace=True)
    miss=[]
    for i in range(len(miss_df)-1):
        if miss_df.loc[i+1,'Chainage_From']==miss_df.loc[i,'Chainage_To']:
            temp+=miss_df.loc[i,'Length']
            print('temp- ',temp)
            
        else:
            if temp>50:
                leng=temp-50
                miss.append([miss_df.loc[i,'Chainage_To']-temp,miss_df.loc[i,'Chainage_To'],leng])
            temp=0
    if temp>50:
        leng=temp-50
        miss.append([miss_df.loc[i,'Chainage_To']-temp,miss_df.loc[i,'Chainage_To'],leng])
        temp=0
    print(miss)
    #df1.to_csv('mix_mb.csv')
    df_miss=df1[df1['Method_Execution']=='Missing'].reset_index()
    miss_lst=[]
    st=0
    for i in range(len(df_miss)):
        if st==0:
            ch_from=df_miss.loc[i,'Chainage_From']
            ch_to=df_miss.loc[i,'Chainage_To']
            st=1
        else:
            if ch_to==df_miss.loc[i,'Chainage_From']:
                ch_to=df_miss.loc[i,'Chainage_To']
            else:
                miss_lst.append((ch_from,ch_to,ch_to-ch_from))
                st=0
        if st==0:
            ch_from=df_miss.loc[i,'Chainage_From']
            ch_to=df_miss.loc[i,'Chainage_To']
            st=1
        if i==len(df_miss)-1:
            miss_lst.append((ch_from,ch_to,ch_to-ch_from))
    # miss_lst[(x,x,x)]
    ### Missing_Chainages sheet from miss_lst###

    values={'Strata_Type':'NORMAL','Depth':1.4}

    df1.fillna(value=values,inplace=True)

    df1['Depth']=df1['Depth'].apply(lambda x: 1.4 if type(x) is str else x)
    df1['Strata_Type']=df1['Strata_Type'].apply(lambda x: 'Normal' if len(x)==0 else x)
    df1['Strata_Type']=df1['Strata_Type'].str.lower()
    df1['Depth']=df1['Depth'].astype(float)
    df1['Protect_Gi']=df1['Protect_Gi'].apply(lambda x: 0 if type(x)!= int else x)
    df1['Protect_Dwc']=df1['Protect_Dwc'].apply(lambda x: 0 if type(x)!= int else x)
    df1['Protect_Rcc']=df1['Protect_Rcc'].apply(lambda x: 0 if type(x)!= int else x)
    df1['Protect_Pcc']=df1['Protect_Pcc'].apply(lambda x: 0 if type(x)!= int else x)
    df1.reset_index(drop=True,inplace=True)
    def method_sum_fun(lst):
        if len(lst)==3:
            if len(lst[-1])==1:
                if lst[-1][0]==1.4:
                    return df1[df1['Depth']>=1.4][df1['Method_Execution']==lst[0]][df1['Strata_Type']==lst[1]]['Length'].sum()
                else:
                    return df1[df1['Depth']<0.6][df1['Method_Execution']==lst[0]][df1['Strata_Type']==lst[1]]['Length'].sum()
            else:
                return df1[(df1['Depth']<lst[-1][0]) & (df1['Depth']>=lst[-1][1])][df1['Method_Execution']==lst[0]][df1['Strata_Type']==lst[1]]['Length'].sum()
        elif len(lst)==1:
            return df1[df1['Method_Execution']==lst[0]]['Length'].sum()

    def gi_sum_fun(lst):
        if len(lst)==3:
            if len(lst[-1])==1:
                if lst[-1][0]==1.4:
                    return df1[df1['Depth']>=1.4]['Protect_Gi'].sum()
                else:
                    return df1[df1['Depth']<0.6]['Protect_Gi'].sum()
            else:
                return df1[(df1['Depth']<lst[-1][0]) & (df1['Depth']>=lst[-1][1])]['Protect_Gi'].sum()
        elif len(lst)==1:
            return df1['Protect_Gi'].sum()
            
    def dwc_sum_fun(lst):
        if len(lst)==3:
            if len(lst[-1])==1:
                if lst[-1][0]==1.4:
                    return df1[df1['Depth']>=1.4]['Protect_Dwc'].sum()
                else:
                    return df1[df1['Depth']<0.6]['Protect_Dwc'].sum()
            else:
                return df1[(df1['Depth']<lst[-1][0]) & (df1['Depth']>=lst[-1][1])]['Protect_Dwc'].sum()
        elif len(lst)==1:
            return df1['Protect_Dwc'].sum()



    final_depth_criteria=[]#7X6
    depth_criteria=[]
    method=['OT','Missing',]
    strata=['normal','hard']
    depth=[[1.4],[1.4,1.3],[1.3,1.2],[1.2,1.1],[1.1,1.0],[1.0,0.6],[0.6]]
    for k in depth:
        for i in method:
            for j in strata:
                depth_criteria.append(method_sum_fun([i,j,k]))
        depth_criteria.append(gi_sum_fun(['xx','yy',k]))
        depth_criteria.append(dwc_sum_fun(['xx','yy',k]))
        final_depth_criteria.append(depth_criteria)
        # print(final_depth_criteria)
        depth_criteria=[]

    method_exec=[]#3X2
    non_bill=0
    for i in miss_lst:
        if i[-1]<50:
            non_bill+=i[-1]
        else:
            non_bill+=50
    method_exec.append([method_sum_fun(['OT']),0])
    method_exec.append([method_sum_fun(['DRT']),0])
    method_exec.append([method_sum_fun(['Missing'])-non_bill,non_bill])

    def check_if_in_missing_fift(ch_from,ch_to,miss_lst):
        for miss_lst_ in miss_lst:
            if (ch_from>=miss_lst_[0] and ch_to<miss_lst_[1]) or (ch_from>=miss_lst_[0] and ch_from<miss_lst_[1]) or (ch_to>=miss_lst_[0] and ch_to<miss_lst_[1]):
                return True
        return False

    protection=[]#3X2
    def check_if_in_missing_fift(ch_from,ch_to,miss_lst):
        for miss_lst_ in miss_lst:
            if (ch_from>=miss_lst_[0] and ch_to<miss_lst_[1]) or (ch_from>=miss_lst_[0] and ch_from<miss_lst_[1]) or (ch_to>=miss_lst_[0] and ch_to<miss_lst_[1]):
                return True
        return False
    dwc_bill=0
    dwc_no_bill=0
    dwc_lst=list(df1[df1['Protect_Dwc']>0]['Protect_Dwc'])
    ch_f_lst=list(df1[df1['Protect_Dwc']>0]['Chainage_From'])
    ch_t_lst=list(df1[df1['Protect_Dwc']>0]['Chainage_To'])
    for i in range(len(ch_f_lst)):
        if check_if_in_missing_fift(ch_f_lst[i],ch_t_lst[i],miss_lst) is True:
            dwc_no_bill+=dwc_lst[i]
        else:
            dwc_bill+=dwc_lst[i]
    gi_bill=0
    gi_no_bill=0
    gi_lst=list(df1[df1['Protect_Gi']>0]['Protect_Gi'])
    ch_f_lst=list(df1[df1['Protect_Gi']>0]['Chainage_From'])
    ch_t_lst=list(df1[df1['Protect_Gi']>0]['Chainage_To'])
    for i in range(len(ch_f_lst)):
        if check_if_in_missing_fift(ch_f_lst[i],ch_t_lst[i],miss_lst) is True:
            gi_no_bill+=gi_lst[i]
        else:
            gi_bill+=gi_lst[i]
    rcc_bill=0
    rcc_no_bill=0
    rcc_lst=list(df1[df1['Protect_Rcc']>0]['Protect_Rcc'])
    ch_f_lst=list(df1[df1['Protect_Rcc']>0]['Chainage_From'])
    ch_t_lst=list(df1[df1['Protect_Rcc']>0]['Chainage_To'])
    for i in range(len(ch_f_lst)):
        if check_if_in_missing_fift(ch_f_lst[i],ch_t_lst[i],miss_lst) is True:
            rcc_no_bill+=rcc_lst[i]
        else:
            rcc_bill+=gi_lst[i]
    protection.append([dwc_bill,dwc_no_bill])
    protection.append([gi_no_bill,gi_bill])
    protection.append([rcc_no_bill,rcc_bill])

    jc=[]#2X2



    cable=[]#4X1
    fdms=[] # 4X1   
    ###
    jc=[]#2X2

    blo=Extract.extract_blo(j)
    column_blo=['prikey', 'Key', 'Survey Duration', 'User', 'Upload Time',
        'Survey Name', '_scheduled_start', 'Version', 'Complete',
        'Survey Notes', 'Location Trigger', 'Instance Name', '_start', '_end',
        '_device', 'instanceid', 'Package_Name', 'Zone_Name', 'District_Name',
        'Mandal_Name', 'From_GP', 'To_GP', 'Span_ID', 'Chainage_From',
        'Chainage_To', 'Length', 'chamber_from', 'chamber_to', 'size_of_ofc',
        'ofc_cable_id', 'cable_len_start', 'cable_len_end',
        'Total_cable_length', 'chamber_a_end_reading',
        'chamber_a_entry_reading', 'chamber_a_end_slack_loop_length',
        'chamber_b_entry_reading', 'chamber_b_end_reading',
        'chamber_b_end_slack_loop_length', 'Remarks']
    df1_blo=pd.DataFrame(columns=column_blo)
    if len(blo)!=0:
        blo,jc=Create_table.create_blo(blo)
        df1_blo=pd.concat([df1_blo,blo])

    drt_ch_from=list(df1[df1['Method_Execution'].isin( ['DRT','Missing'])]['Chainage_From'])
    drt_ch_to=list(df1[df1['Method_Execution'].isin( ['DRT','Missing'])]['Chainage_To'])

    ch_from_lst=list(df1_blo['Chainage_From'])
    ch_to_lst=list(df1_blo['Chainage_To'])
    chmb_from_lst=list(df1_blo['chamber_from'])
    chmb_to_lst=list(df1_blo['chamber_to'])
    jc_bill=0
    jc_no_bill=0

    def check_in_drt(x,lst_from,lst_to):
        if type(x) is int:
            for i in range(len(lst_to)):
                if x>=lst_from[i] and x<lst_to[i]:
                    return True
        return False
    for i in range(len(ch_from_lst)):
        if i==0:
            if check_in_drt(ch_from_lst[i],drt_ch_from,drt_ch_to):
                if 'HH' in chmb_from_lst[i] or 'MH' in chmb_from_lst[i]:
                    jc_no_bill+=1
            else:
                if 'HH' in chmb_from_lst[i] or 'MH' in chmb_from_lst[i]:
                    jc_bill+=1
        if check_in_drt(ch_to_lst[i],drt_ch_from,drt_ch_to):
            if 'HH' in chmb_to_lst[i] or 'MH' in chmb_to_lst[i]:
                jc_no_bill+=1

        else:
            if 'HH' in chmb_to_lst[i] or 'MH' in chmb_to_lst[i]:
                jc_bill+=1

    jc.append([jc_bill,jc_no_bill])
    jc.append([jc_bill,jc_no_bill])

    ####

    print(final_depth_criteria)
    print(method_exec)
    print(protection)


    ############################DATA FILLING###############################################3
    print("Started Writing in excel")
    #j='WRG-ZAF-5706-M-01-GR01-02.xlsx'
    xfile = openpyxl.load_workbook(span_file)
    sheet1=xfile['Summary']

    row=3
    no=48
    for i in final_depth_criteria:
        first=math.floor((no+row-48)/10)
        temp=(chr(no+first))+ chr((no+row-48)%10+48)
        sheet1['B'+temp]=i[0]
        sheet1['C'+temp]=i[1]
        sheet1['D'+temp]=i[2]
        sheet1['E'+temp]=i[3]
        sheet1['F'+temp]=i[4]
        sheet1['G'+temp]=i[5]
        row+=1
        


    row=11
    no=48
    for i in method_exec:
        first=math.floor((no+row-48)/10)
        temp=(chr(no+first))+ chr((no+row-48)%10+48)
        sheet1['C'+temp]=i[0]
        sheet1['D'+temp]=i[1]
        row+=1
    


    row=15
    no=48
    for i in protection:
        first=math.floor((no+row-48)/10)
        temp=(chr(no+first))+ chr((no+row-48)%10+48)
        sheet1['B'+temp]=i[0]
        sheet1['C'+temp]=i[1]
        row+=1
            

    row=19
    no=48
    for i in jc:
        first=math.floor((no+row-48)/10)
        temp=(chr(no+first))+ chr((no+row-48)%10+48)
        sheet1['B'+temp]=i[0]
        sheet1['C'+temp]=i[1]
        row+=1         

    sheet2=xfile['Missing Chainages']

    row=2
    no=48
    for i in miss_lst:
        first=math.floor((no+row-48)/10)
        temp=(chr(no+first))+ chr((no+row-48)%10+48)
        sheet2['B'+temp]=i[0]
        sheet2['C'+temp]=i[0]
        sheet2['D'+temp]=i[0]
        row+=1

    print(miss_lst)

    xfile.save(filename=span_file)

def boq_format():
    global file2
    t=filedialog.askopenfilename(initialdir="C:/")
    file2=file2+t
    print(t,file2)


def mandal_folder(): 
    global file1   
    # file_path = filedialog.askopenfilename()
    t=filedialog.askdirectory()
    file1=file1+t
    print(t,file1)

def boq_checkup():

    counter=0
    top= Toplevel(popupRoot)
    top.geometry("980x550")
    l = Label(top, text = "Here is a list of Errors!!")
    l.config(font =("Courier", 14))
    b2 = Button(top, text = "Exit",command = top.destroy)
    top.title("Error Check-up Window")
    superstring=""

    mandal_file=file1+'/'
    if os.path.exists(mandal_file+"BOQ122.xlsx"):
        #print("123")
        counter+=1
        superstring=superstring+str(counter)+". Remove the already present BoQ122.xlsx file in Mandal Folder"+"\n"
        #os.remove(mandal_file+"BOQ122.xlsx")

    os.chdir(mandal_file)
    
    for info in os.listdir():

        j=mandal_file+info
        
        ###Summary checkup
        summary=[]
        try:
            summary=pd.read_excel(j,sheet_name='Summary')
        
        except:
            if len(summary)==0:
                counter+=1
                superstring = superstring+str(counter)+" .No Summary sheet found. Please make summary sheet in "+info+"\n"

        ###Missing Chainages Checkup
        missing_chainages=[]
        try:
            missing_chainages=pd.read_excel(j,sheet_name='Missing Chainages')

        except:   
            if len(missing_chainages)==0:
                counter+=1
                superstring = superstring+str(counter)+" .No Missing Chainages sheet found. Please make the sheet in "+info+"\n"


        try:
            ot=Extract.extract_ot(j)
            if len(ot.columns)!=18:
                counter+=1
                superstring = superstring+str(counter)+" .Number of columns is different (Duplicate/Extra Columns found) in OT MB of "+info+". Count of columns is "+str(len(ot.columns))+"\n"
                
        except:
            counter+=1
            superstring=superstring+str(counter)+". No OT MB captured in "+info+""". Standard names are "OT","R4_OT","OT MB","R04_T&D"."""+"\n"


        try:
            blo=Extract.extract_blo(j)
            if len(blo.columns)!=18 and len(blo.columns)!=17:
                #print("Working1")
                counter+=1
                superstring = superstring+str(counter)+""". Number of columns is different (Duplicate/Extra Columns found) in Blowing MB of """+info+""". Count of columns is """+str(len(blo.columns))+"\n"
                #print(tem)
        except:
            counter+=1
            superstring=superstring+str(counter)+". No Blowing MB captured in "+info+""". Standard names are "BLOWING","blowing","R10_BLOWING","R10_Blowing"."""+"\n"
        
        try:
            drt=Extract.extract_drt(j)
            j=0
            columns_={}
            for i in drt.columns:
                if j==0:
                    columns_[i]='S_no'
                elif j==1:
                    columns_[i]='Ch1_lat'
                elif j==2:
                    columns_[i]='Ch1_long'
                elif j==3:
                    columns_[i]='Ch1_cond'
                elif j==4:
                    columns_[i]='Ch1_route_marker'
                elif j==5:
                    columns_[i]='Ch2_lat'
                elif j==6:
                    columns_[i]='Ch2_long'
                elif j==7:
                    columns_[i]='Ch2_cond'
                elif j==8:
                    columns_[i]='Ch2_route_marker'
                elif j==9:
                    columns_[i]='Ch_from'
                elif j==10:
                    columns_[i]='Ch_to'
                elif j==11:
                    columns_[i]='Len'
                elif j==12:
                    columns_[i]='Duct_dam_lat'
                elif j==13:
                    columns_[i]='Duct_dam_long'
                elif j==14:
                    columns_[i]='Duct_dam_ch_from'
                elif j==15:
                    columns_[i]='Duct_dam_ch_to'
                elif j==16:
                    columns_[i]='Duct_dam_len'
                elif j==17:
                    columns_[i]='Duct_miss_lat'
                elif j==18:
                    columns_[i]='Duct_miss_long'
                elif j==19:
                    columns_[i]='Duct_miss_ch_from'
                elif j==20:
                    columns_[i]='Duct_miss_ch_to'
                elif j==21:
                    columns_[i]='Duct_miss_len'
                elif j==22:
                    columns_[i]='Remark'
                j+=1
            drt.rename(columns=columns_,inplace=True)
            drt.reset_index(drop=True,inplace=True)
            for i in range(len(drt)):
                if type(drt.loc[i,'Duct_dam_ch_from'])==int and type(drt.loc[i,'Duct_dam_ch_to'])==int and type(drt.loc[i,'Ch_from'])==int and type(drt.loc[i,'Ch_to'])==int:
                    if drt.loc[i,'Duct_dam_ch_from']!=drt.loc[i,'Ch_from'] and drt.loc[i,'Duct_dam_ch_to']!=drt.loc[i,'Ch_to']:
                        counter+=1
                        superstring=superstring+str(counter)+". Please correct chainage error in DRT MB of "+info+"\n"

            for i in range(len(drt)):
                if type(drt.loc[i,'Duct_miss_ch_from'])==int and type(drt.loc[i,'Duct_miss_ch_to'])==int and type(drt.loc[i,'Ch_from'])==int and type(drt.loc[i,'Ch_to'])==int:
                    if drt.loc[i,'Duct_miss_ch_from']!=drt.loc[i,'Ch_from'] and drt.loc[i,'Duct_miss_ch_to']!=drt.loc[i,'Ch_to']:
                        counter+=1
                        superstring=superstring+str(counter)+". Please correct chainage error in DRT MB of "+info+"\n"


            if len(drt.columns)!=23:
                #print("Working2")
                counter+=1
                superstring = superstring+str(counter)+". Number of columns is different (Duplicate/Extra Columns found) in DRT MB of "+info+". Count of columns is "+str(len(drt.columns))+"\n"
                
        except:
            counter+=1
            superstring=superstring+str(counter)+". No DRT MB captured in "+info+""". Standard names are "DRT","R09_DRT","R09-DRT","R9_DRT"."""+"\n"
        
        

    text_box = Text(
    top,
    height=50,
    width=200
    )
    text_box.pack(expand=True)
    text_box.insert('end', superstring)
    text_box.config(state='disabled')

    #tata="Raadsf aesr faaser asd aw fesf aesf /n aser asef asef "
    #l1 = Label(top, text = tata)
    #print("superstring",type(superstring),len(superstring),superstring,"1")
    #l1.config(font =("Courier", 7))
    #l1.pack()
    l.pack()
    b2.pack()

def boq12():
    xfile = openpyxl.load_workbook(file2)
    sheet1=xfile['GP-Wise BOQ']
    mandal_file=file1+'/'
    print("Grabbed mandal location")
    no=0  # Later used in feeding the data
    #sorted_mandal=os.chdir(mandal_file)
    #sorted_mandal=sorted(sorted_mandal)
    for info in sorted(os.listdir(mandal_file)):

        j=mandal_file+info
        print(j)

        blo=Extract.extract_blo(j)
        if len(blo)!=0:
            (blo,joint_closer)=Create_table.create_blo(blo)
            #joint_closer.to_csv('Joint_closure.csv')

        drt=Extract.extract_drt(j)
        drt.to_csv("drt1.csv")
        if len(drt)!=0:
            drt=Create_table.create_drt(drt)
            drt.reset_index(drop=True,inplace=True)

        ot=Extract.extract_ot(j)
        if len(ot)!=0:
            ot=Create_table.create_ot(ot)   

        hdd=Extract.extract_hdd(j)
        if len(hdd)!=0:
            hdd=Create_table.create_hdd(hdd)  
        """
        try:
            mix_mb(j)
        except:
            print("Can't do calculation in summary sheet. Some Error")    
        """

        drt.to_csv("drt2.csv")
        s=0 #no idea!
        fin=0 #Missing section sum
        fin1=0  #Damaged section sum
        fin50=0
        fin5=0
        fintot=0
        temp=0  #Temp no imp value
        ####DRT Calculation starts!!!!
        ##Case 1: SUM From chainage
        try:
            sum=0
            drt=drt.reset_index(drop=True)
            for i in range(len(drt)):
                ch1=drt.loc[i,'ch_from']
                ch2=drt.loc[i,'ch_to']
                if i!=(len(drt)-1):
                    ch3=drt.loc[i+1,'ch_from']
                    if(ch3!=ch2):
                        sum+=ch3-ch2
            
            if np.isnan(sum):
                sum=0

            print("sum after chaining",sum)
            
        ##Case 2: SUM From missing section
            ##For missing case -50 agreement  
            temp=0
            s=0  
            for i in range(len(drt)):
                ch1=drt.loc[i,'Duct_miss_ch_from']
                ch2=drt.loc[i,'Duct_miss_ch_to']
                if i!=(len(drt)-1):
                    ch3=drt.loc[i+1,'Duct_miss_ch_from']

                temp=(drt.loc[i,'Duct_miss_ch_Length'])

                if ch3==ch2:
                    s+=temp  

                elif ch3!=ch2:
                    #print("got it")
                    s=np.nansum([s+temp])
                    #print(temp,"s is",s,"done")
                    #print(s)
                    if(s>=50):
                        fin5+=1
                        fin50+=s
                        s-=50
                        fin+=s
                        s=0
                    elif s<50:
                        fintot+=s
                        s=0
            ##Case 3: SUM From damaged section
            ##For damaged case -50 Agreement
            temp=0
            s=0
            for i in range(len(drt)):
                ch1=drt.loc[i,'Duct_dam_ch_from']
                ch2=drt.loc[i,'Duct_dam_ch_to']
                if i!=(len(drt)-1):
                    ch3=drt.loc[i+1,'Duct_dam_ch_from']

                temp=(drt.loc[i,'Duct_dam_ch_Length'])

                if ch3==ch2:
                    s+=temp  

                elif ch3!=ch2:
                    #print("got it")
                    s=np.nansum([s+temp])
                    #print(temp,"s is",s,"done")
                    #print(s)
                    if(s>=50):
                        s-=50
                        fin1+=s
                        s=0
                    elif s<50:
                        s=0
                #print(temp," ",s," ",fin," ",ch1," ",ch2," ",ch3)       

            if np.isnan(sum):
                sum=0  


        except:
            print("Something Wrong in missing calculation")
        #print("Sum after missing")
        if np.isnan(fin):
            fin=0   

        if np.isnan(fin1):
           fin1=0           
        
        sum+=fin1
        print("Present value of sum after Chainaing && (mis+dam) case",sum)

        print("Missing case",fin,"Dam case",fin1)                

        try:
            blo=blo.reset_index(drop=True)
            start_drt=drt['ch_from'].iloc[0]
            sum+=(start_drt-0)
            print("Sum before starting 0 is ",start_drt-0)
            print("Present value of sum after Chaining && starting case and after (mis+dam) case",sum)

            #####################Calculation Last Blow########################################
            ##Case 4: SUM From lat blowing and DRT reading mismatch section
            last_blow=blo['Chainage_To'].iloc[-1]
            print("Doing this calc. of last_blowing",last_blow)
            att=pd.isnull(last_blow)
            #print(att)
            atp=not(isinstance(last_blow, int))
            #print(atp)
            atc=att+atp
            #print(atc)
            #np.isnan(last_blow) or
            if (atc):
                last_blow=blo['Chainage_To'].iloc[-2]
                print("Last blow for -2 case",last_blow)
            print("Final case",last_blow)
            if np.isnan(last_blow):
                last_blow=0    
            ##############################Last BLow Calculation################################

            ###################Last DRT Calculation#####################           
            last_drt=drt['ch_to'].iloc[-1]
            att1=pd.isnull(last_drt)
            #print(att)
            atp1=not(isinstance(last_drt, int))
            #print(atp)
            atc1=att1+atp1
            #print(atc1)
            #np.isnan(last_blow) or
            if (atc1):
                last_drt=drt['Ch_to'].iloc[-2]
                print("Last drt for -2 case",last_drt)
            
            ##########Last DRT CAlculation########################

            if np.isnan(last_drt):
                last_drt=0           
            print("Doing last cal of last drt",last_drt)
            sum+=last_blow-last_drt
            print("Final Sum is ",sum)    
            
        
        except:
            print('Something wrong in blowing gap calculation or drt file')    
        
        ##DRT Calculation over!!!!!!
        ######################################################################################################################################
        ######################################################################################################################################
            

        ###Feeeding the data starts!!!    
        #no=0
        var='F'
        if no<=10:
            var=chr(ord(var)+2*no)
        elif no>10:
            var='A'+chr(ord(var)+2*no-26)
        try:
            sheet1[var+'3']=info[:-5]
            #sheet1[var1+'4']='STATE'
            #sheet1[var+'4']='BBNL'
            sheet1[var+'11']=(blo.loc[blo['size_of_ofc'].isin(['288F','288 F','288']),'Total_cable_length'].sum())/1000
            sheet1[var+'14']=(blo.loc[blo['size_of_ofc'].isin(['144F','144 F','144']),'Total_cable_length'].sum())/1000
            sheet1[var+'17']=(blo.loc[blo['size_of_ofc'].isin(['96F','96 F','96']),'Total_cable_length'].sum())/1000
            sheet1[var+'18']=(blo.loc[blo['size_of_ofc'].isin(['48F','48 F','48']),'Total_cable_length'].sum())/1000
   

        except:
            print("No such Blowing files..")

        try:                
            sheet1[var+'29']=(joint_closer.loc[joint_closer['cha_loop'].isin(['288F','288 F','288']),'chb_end'].sum())
            sheet1[var+'32']=(joint_closer.loc[joint_closer['cha_loop'].isin(['144F','144 F','144']),'chb_end'].sum())
            sheet1[var+'35']=(joint_closer.loc[joint_closer['cha_loop'].isin(['96F','96 F','96']),'chb_end'].sum())
            sheet1[var+'36']=(joint_closer.loc[joint_closer['cha_loop'].isin(['48F','48 F','48']),'chb_end'].sum())

            
        except:
            print("No such Joint Closure files..")  

        try:
            if len(drt)!=0 and sum==0:
                sheet1[var+'22']=sum/1000
                #sheet1[var+'139']=fin/1000
                sheet1[var+'137']=(fin50 + fintot)/1000
                sheet1[var+'138']=(fin5*50)/1000
                
                a=(drt['Duct_miss_ch_Length']==0)
                b=~(drt['Duct_miss_ch_Length'].astype(str).str.isdigit())
                y=a+b
                print("Dit length only miss",drt.loc[y,'Length'].sum()/1000)
                c=(drt['Duct_dam_punct_loc_Length']==0)
                d=~(drt['Duct_dam_punct_loc_Length'].astype(str).str.isdigit())
                #print(c+d)
                z=c+d
                #tim=(drt.loc[z,'Length'])
                #print(tim.loc['Length'].sum())
                #print(drt.loc[z,'Length'].sum())
                print("Dit length only dam",drt.loc[z,'Length'].sum()/1000)

                e=y*z
                print("Dit length",drt.loc[e,'Length'].sum()/1000)
                sheet1[var+'98']=drt.loc[e,'Length'].sum()/1000
                #sheet1[var+'21']=sum/1000
                print("Duct laid",sum)

            elif len(drt)==0 and sum==0:
                print("No DRT File case triggered,dit=0, duct laid=sum(blowing)")
                sum=(blo['Length'].sum())
                print("No drt case+no sum case",sum)
                sheet1[var+'99']=0
                sheet1[var+'22']=sum/1000
                sheet1[var+'138']=(fin5*50)/1000
                sheet1[var+'137']=(fin50 + fintot)/1000
                #sheet1[var+'138']=(fin5*50)/1000

            else:    
                a=(drt['Duct_miss_ch_Length']==0)
                b=~(drt['Duct_miss_ch_Length'].astype(str).str.isdigit())
                y=a+b
                #print(a+b)
                print("Dit length only miss",drt.loc[y,'Length'].sum()/1000)
                c=(drt['Duct_dam_punct_loc_Length']==0)
                d=~(drt['Duct_dam_punct_loc_Length'].astype(str).str.isdigit())
                #print(c+d)
                z=c+d
                #tim=(drt.loc[z,'Length'])
                #print(tim.loc['Length'].sum())
                #print(drt.loc[z,'Length'].sum())
                print("Dit length only dam",drt.loc[z,'Length'].sum()/1000)

                e=y*z
                print("Dit length",drt.loc[e,'Length'].sum()/1000)
                sheet1[var+'99']=drt.loc[e,'Length'].sum()/1000
                sheet1[var+'22']=sum/1000
                sheet1[var+'138']=(fin5*50)/1000
                sheet1[var+'137']=(fin50 + fintot)/1000
                #sheet1[var+'138']=(fin5*50)/1000
                print("Duct laid",sum)
        except:
            print("No such DRT files..")


        try:     
            sheet1[var+'23']=(ot['Protect_Dwc'].sum())/1000
            sheet1[var+'24']=(ot['Protect_Gi'].sum())/1000
            sheet1[var+'25']=(ot['Rcc_Marker'].sum())
            sheet1[var+'26']=(ot['Rcc_Chamber'].sum())
   

        except:
            print("No such OT files..")    

        xfile.save(filename="BOQ122.xlsx")
        ##print(duct_laid,dwc_pipe,gi_pipe,route_indicators,joint_chambers)
        no+=1
    popupRoot.destroy()

popuplabel = Label(popupRoot, text = 'BoQer Code V2.1',font = ("Times New Roman", 11)).grid(row=0,column=1)
popupButton = Label(popupRoot, text = 'STL',bg='white',fg='blue',font = ("Times New Roman", 13,'bold'), anchor="e",justify=RIGHT).grid(row=0,column=2)
popuplabel = Label(popupRoot, text = ' ',font = ("Times New Roman", 12)).grid(row=1,column=1)
popuplabel = Label(popupRoot, text = 'Select the BoQ Format',font = ("Times New Roman", 12)).grid(row=2,column=1)
popupButton = Button(popupRoot, text = 'BoQ Format', font = ("Times New Roman", 12), command = boq_format,width = 20).grid(row=2,column=2)
popuplabel = Label(popupRoot, font = ("Times New Roman", 12),text = 'Select Mandal Folder').grid(row=3,column=1)
popupButton = Button(popupRoot, text = 'Mandal Folder', font = ("Times New Roman", 12), command = mandal_folder,width = 20).grid(row=3,column=2)
popuplabel = Label(popupRoot, text = ' ',font = ("Times New Roman", 12)).grid(row=4,column=1)
popupButton = Button(popupRoot, text = 'Run BoQ Script', bg='light green',font = ("Times New Roman", 15), command = boq12,width = 12).grid(row=5,column=2)
popupButton = Button(popupRoot, text = 'Error Checkup', bg='orange',font = ("Times New Roman", 15), command = boq_checkup,width = 12).grid(row=5,column=1)
popupRoot.geometry('350x200')
#print("1")

#print("printing")
#print(file2)
#print(file1)
#print("1")

#popupRoot.destroy()
mainloop()
