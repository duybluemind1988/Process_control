#Clear all
# add from spyder

import streamlit as st
import seaborn as sns; sns.set()
import pandas as pd
import numpy as np
import copy

import plotly.graph_objects as go
from plotly.subplots import make_subplots
#import statistics as st
import warnings
warnings.filterwarnings("ignore")
import glob
#import sys,os
import xlrd

#-----------------Design_layout main side-----------------#
#trial_path='//Vn01w2k16v18/data/Copyroom/Test_software/Data/Control plan/Control plan set up 3000/Set up control limit'
trial_path='/mnt/01D6B57CFBE4DB20/1.Linux/Data/Process_control/Control plan set up VPU/Set up limit'
#trial_path='//Vn01w2k16v18/data/Copyroom/Test_software/Data/Control plan/Control plan 3000'
#trial_path='/media/ad/01D6B57CFBE4DB20/1.Linux/Data/Process_control/Control plan 3000'
#trial_path='/media/ad/01D6B57CFBE4DB20/1.Linux/Data/Process_control/Control plan 2600'
#trial_path='/media/ad/01D6B57CFBE4DB20/1.Linux/Data/Process_control/Control plan E series 1'
st.markdown('<style>h1{color: green;}</style>', unsafe_allow_html=True)
st.title('Process quality control')

st.subheader('Created by: DNN')
st.header("1. Set up control limit for control chart")

input_folder= st.text_input("1. Please input folder name for data analysis (MUST)",trial_path,key="setup_data")
input_folder=input_folder +'/'
#path='//Vn01w2k16v18/data/Copyroom/Test_software/Data/Membrane 3000S/'
#st.write('path input: '+path)

st.write('Accept any public folders like copy room or QA folder, currently not allow private folder')
save_folder=st.text_input("2. Please input folder name for saving (option)",'//Vn01w2k16v18/data/Copyroom/Test_software/Data/Save')
save_folder=save_folder+'/'

#-----------------Design_layout left side-----------------#
st.sidebar.title("Help")
#st.sidebar.info("2. Number of Sample: all sample need to be measured each inspection")
st.sidebar.info("1. Folder name: Folder contains all excel file related to specific product")
st.sidebar.info("2. Name file: product name or item config for saving. Ex: Membrane_33AA079")

st.sidebar.title("About author")
st.sidebar.info("if you have questionns, please contact to DNN@sonion.com")
st.sidebar.title("Software")
st.sidebar.info("This web app was written by Python program language (offline mode). Please use\
                this app inside Sonion only")
#---------------sort file by created_time-------------------------# 

#path=path+'/'
#print(path)
all_files_xlsx=glob.glob(input_folder + '*.xlsx')
all_files_xlsm=glob.glob(input_folder + '*.xlsm')
all_files_combine=all_files_xlsx+all_files_xlsm
#sort file in directory by reverse:
all_files_combine = sorted(all_files_combine, reverse = False)
st.text('number of files: '+str(len(all_files_combine)))
st.text(all_files_combine)
# Get master sheet dataframe:
@st.cache(suppress_st_warning=True)
def master_sheet_data_func(all_files_combine):
    path_name=all_files_combine[-1] # get latest file
    xls = xlrd.open_workbook(path_name, on_demand=True)
    sheet_names=xls.sheet_names()
    master_sheet_name=sheet_names[0] 
    master_sheet=pd.read_excel(xls, master_sheet_name)
    return master_sheet
master_sheet_data=master_sheet_data_func(all_files_combine)
#---------------process data-------------------------# 
#@st.cache(suppress_st_warning=True,allow_output_mutation=True)
@st.cache(suppress_st_warning=True,allow_output_mutation=True)
def create_sheet_dict(all_files_combine):
  all_process_week={}
  for path_name in all_files_combine: # Đọc 2 file đầu tiên thôi
    week_name=path_name[-16:-5]
    print(week_name,'-'*20)
    #print(path_name)
    xls = xlrd.open_workbook(path_name, on_demand=True)
    sheet_names=xls.sheet_names()
    #print(xls.sheet_names())
    sheet_dict={}
    xls = pd.ExcelFile(path_name) 
    #for name in sheet_names: # read each sheet in excel file
    #    sheet_dict[name] = pd.read_excel(xls, name)
    sheet_all={} # most important (reset sheet_all to empty)
    sheet_error=[]
    for name_sheet in sheet_names[1:]: # Đọc tat ca tru sheet 1 (chua name process)
      print(name_sheet)
      #if name_sheet!='83748-Check Golden samples': continue # debug each name sheet
      sheet_dict[name_sheet] = pd.read_excel(xls, name_sheet)
      sheet=sheet_dict[name_sheet]
      # Tim begin col và end col
      row_contain_indicate_feature_column = 21
      values_col=sheet.iloc[row_contain_indicate_feature_column,:]
      values_col.reset_index(drop=True,inplace=True)  # Date, MSVN, DIM A,B,C....
      begin_col=values_col.first_valid_index()
      end_col=values_col.last_valid_index()
      if end_col==begin_col: # for loop did not allow same value
          end_col=end_col+1
      df_dict={} # add all value, USL, LSL, UCL... in each process sheet
      for name in sheet.columns[begin_col:(end_col+1)]: 
          row_contain_feature_column=22
          dim_name=sheet[name][row_contain_feature_column]
          df=pd.DataFrame()
          df_dict[dim_name]={}
          # Add time value (hour) to data:
          column_hour = 9
          row_start_value = 25
          try:
              df['Hour']=sheet[sheet.columns[column_hour]][row_start_value:]
              #print(df['Date'])
              df['Hour']=df['Hour'].apply(lambda x: x.strftime("%Y %m %d %H")) # group theo hour, gần như trùng với tần suất lấy mẫu đo control plan
              df['Hour']=pd.to_datetime(df['Hour'])
              #print(df['Hour'])
          except Exception as e: 
              print(e)
              continue
          df['Value']=sheet[name][row_start_value:]
          #print(df["Value"])
          row_contain_usl = 23
          row_contain_lsl = 24  
          if isinstance(sheet[name][row_contain_usl],str): 
              continue #check string USL 25/11
          if isinstance(sheet[name][row_contain_usl],str): 
              continue #check string LSL 25/11
          df['USL']=sheet[name][row_contain_usl] # max
          df['LSL']=sheet[name][row_contain_lsl] # min
          #print(df['USL'],df['LSL'])
          df.dropna(subset=['Value'],inplace=True)
          # Convert to numeric
          df=df[pd.to_numeric(df['Value'], errors='coerce').notnull()] #25/11: loai bo duy nhật 1 cột Value có not numeric value truoc khi convert
          # Convert all column data to float (except first column-date)
          df[df.columns[1:]]=df[df.columns[1:]].astype('float32')
          #print(df)
          df_dict[dim_name]=df.reset_index(drop=True)

      sheet_all[name_sheet]=df_dict # add each process name to shee_all dict each week
    all_process_week[week_name]=sheet_all
    #print(week_name,sheet_error)
  return all_process_week

all_process_week=create_sheet_dict(all_files_combine) # Run function above
st.text('number of week: '+str(len(all_process_week)))

#-----------Concat all week to one base week-------------------------#

#@st.cache(suppress_st_warning=True,allow_output_mutation=True)
@st.cache(suppress_st_warning=True,allow_output_mutation=True)
def concat_baseweek(master_sheet_data,all_process_week):
    base_week={}
    #Shallow copy: (wrong, reference to wk 37)
    #base_week=all_process_week[list(all_process_week.keys())[-1]].copy() # must have copy()
    #Deep copy:
    base_week=copy.deepcopy(all_process_week[list(all_process_week.keys())[-1]])
    #print(len(base_week))
    # concat all process in each weeks based on keys value from base weeks:
    for week_name in list(all_process_week.keys())[:-1]: # ko tính base week nên trừ 1: 
      #print(week_name)      
      other_week=all_process_week[week_name]
      for process_name in other_week.keys(): # dict all process
        for process_name_base in base_week.keys(): # dict all process
            if process_name_base==process_name: # Neu process name trung voi base week thi concat
              for dim_name in base_week[process_name_base].keys(): # concat dim trong process base week voi cac week khac nhau
                  for other_week_dim_name in other_week[process_name_base].keys():     
                      if dim_name==other_week_dim_name: # Neu dim name trong other week trung voi dim name trong base week thi concat
                          if len(base_week[process_name_base][dim_name]) == 0: #20/11 convert dict to df for append
                             base_week[process_name_base][dim_name] = pd.DataFrame.from_dict(base_week[process_name_base][dim_name])
                          if len(other_week[process_name_base][dim_name]) == 0: #20/11 convert dict to df for append
                            other_week[process_name_base][dim_name] = pd.DataFrame.from_dict(other_week[process_name_base][dim_name])
                          #try:   
                          #base_week[process_name_base][dim_name]=pd.concat([base_week[process_name_base][dim_name],other_week[process_name_base][dim_name]])
                          base_week[process_name_base][dim_name]=base_week[process_name_base][dim_name].append(other_week[process_name_base][dim_name]) 
                          # sử dụng hàm append có thể concat empty dataframe tất cả các week
                          #except:continue
    # 15/12/2020 Change sheet name to process name in baseweek data:  
    all_sheet_name=copy.deepcopy(list(base_week.keys()))
    sheet_name_column='Unnamed: 1'
    process_name_column="Unnamed: 4"
    def isNaN(string):
        return string != string
    for name_sheet in all_sheet_name:
        #print(name_sheet)
        try:
            process_name=master_sheet_data.loc[master_sheet_data[sheet_name_column] == name_sheet][process_name_column].values[0]
        except:
            process_name=name_sheet
        #print(process_name)
   
        if process_name!=name_sheet and isNaN(process_name)!= True: # chi thay doi ten dict neu 2 name khac nhau va process name != nan
            base_week[process_name] = base_week[name_sheet]
            del base_week[name_sheet]   
        
    # Check len all dim in base week for remove zero dim and zero process
    process_name_list=[]
    dim_name_list=[]
    dim_len_list=[]
    for process_name in base_week.keys(): # all process 
      #print(process_name)
      for dim_name in base_week[process_name].keys(): # all dim in each process
        dim_len=len(base_week[process_name][dim_name])
        #print(dim_name)
        #print()
        process_name_list.append(process_name)
        dim_name_list.append(dim_name)
        dim_len_list.append(dim_len)
    dim_and_len_df =pd.DataFrame(list(zip(process_name_list,dim_name_list,dim_len_list)),
                             columns=['Process_Name','Dim_name','Len'])
    dim_and_len_df.sort_values(by='Len')
    # Remove all dim with len = 0
    for i in range(len(dim_and_len_df)):
      if dim_and_len_df.loc[i].Len==0:
        Process_Name=dim_and_len_df.loc[i].Process_Name
        Dim_name=dim_and_len_df.loc[i].Dim_name
        #print(Process_Name,Dim_name) # bat buoc se add vo sau khi he debug
        base_week[Process_Name].pop(Dim_name,None)
    # Remove all process with no dim:
    all_base_week_process_list=list(base_week.keys())
    for process_name in all_base_week_process_list:
      if len(base_week[process_name]) == 0 :
        base_week.pop(process_name,None)
    return base_week

final_data=concat_baseweek(master_sheet_data,all_process_week)    
#print(base_week['39682'])
st.text('number of process in line: '+str(len(final_data)))

#-----------Calculate process indicator base on 25 lates subgroup sample-------------------
@st.cache(allow_output_mutation=True) # Mo them vao sau nay
def process_performance(df):
  constants={
    2:1.128,3:1.693,4:2.059,5:2.326,6:2.534,7:2.704,8:2.847, 9: 2.970,10: 3.078,
    11: 3.173,12: 3.258,13: 3.336,14: 3.407,15: 3.472,16: 3.532,17:3.588,18:3.640,
    19:3.689,20:3.735,
    }  
  #print('dim: ',name)
  # Calculate sigma
  sigma=np.std(df.Value)
  n=df.Hour.value_counts()[0] # check sub group num
  # Method 1: Collect only 25 group sample for calculation # Sẽ có 33 dim name / 14 process
  #num_sample=n*25
  #df_temp=df[-num_sample:]
  #df_temp=df_temp.reset_index(drop=True)
  # Method 2: Not collect 25 group sample, use full #  có 33 dim name / 14 process
  df_temp=df.reset_index(drop=True) # Tai sao phải reset index ?

  # Calculae usl, lsl
  usl=df_temp.USL[0]
  lsl=df_temp.LSL[0]
  m=df_temp.Value.mean() 
    
  if sigma ==0 :
      Mean=df_temp['Value'].mean()
      UCL=LCL=Mean
      Cp=float("NaN")
      Cpk=float("NaN")
      Pp=float("NaN")
      Ppk=float("NaN")
  if sigma !=0:
      #UCL, LCL, Mean
      k=3 
      UCL=df_temp['Value'].mean() + sigma*k
      LCL=df_temp['Value'].mean() - sigma*k
      Mean=df_temp['Value'].mean()
      #Ppk
      #print("sigma",sigma)
      Pp = float(usl - lsl) / (6*sigma)
      Ppu = float(usl - m) / (3*sigma)
      Ppl = float(m - lsl) / (3*sigma)
      Ppk = np.min([Ppu, Ppl])
      #print('Pp:{:.2f} , Ppk: {:.2f}'.format(Pp,Ppk))
      #Cpk
      temp=df_temp.groupby('Hour').agg({'Value':['min','max']})
      temp['Range']=temp['Value','max']-temp['Value','min']
      Range=temp['Range'].mean()

      if n <= 20:
        sigma_within = Range/constants[n]
      else:
        sigma_within = Range/constants[20]

      Cp = float(usl - lsl) / (6*sigma_within)
      Cpu = float(usl - m) / (3*sigma_within)
      Cpl = float(m - lsl) / (3*sigma_within)
      Cpk = np.min([Cpu, Cpl])
      #print('Cp:{:.2f} , Cpk:{:.2f}'.format(Cp,Cpk))
      if np.isnan(usl):
        Cpk=Cpl
        Ppk=Ppl
      elif np.isnan(lsl):
        Cpk=Cpu
        Ppk=Ppu
      else:
        Cpk = np.min([Cpu, Cpl])
        Ppk = np.min([Ppu, Ppl])
  #print(Ppk,Cpk)
  Cp=round(Cp,2)
  Cpk=round(Cpk,2)
  Pp=round(Pp,2)
  Ppk=round(Ppk,2)
  UCL=round(UCL,2)
  LCL=round(LCL,2)
  Mean=round(Mean,2)
  return Cp,Cpk,Pp,Ppk,UCL,LCL,Mean

#-----------create_process_indicator-------------------
@st.cache(suppress_st_warning=True,allow_output_mutation=True)
def create_process_indicator(base_week): 

    process_indicator_dict={}
    process_indicator_df=pd.DataFrame(columns=['Process_name','Dim_name','Cp','Pp','Cpk','Ppk','UCL','LCL','Mean'])
    #process_indicator_df.columns=['Dim','Cp','Cpk','Pb','Ppk']
    i=0
    for process_name in list(base_week.keys()):
      #print(process_name)
      df_dict=base_week[process_name]
      for dim_name in list(df_dict.keys()): #also group
        #print(dim_name)
        df=df_dict[dim_name]
        #print(dim_name)
        try: # debug purpose
          Cp,Cpk,Pp,Ppk,UCL,LCL,Mean=process_performance(df) 
          df["UCL"]=UCL
          df["LCL"]=LCL
          df["Mean"]=Mean
        except:
            continue # sigma = 0 phai continue de tranh dung chuong trinh
            #print(process_name)
            #print(dim_name)
            #print(df)
        #process_indicator_dict[process_name]=[dim_name,Cp,Pp,Cpk,Ppk]
        process_indicator_df.loc[i]=process_name,dim_name,Cp, Pp, Cpk, Ppk,UCL,LCL,Mean
        print(process_indicator_df)
        i+=1

    #process_indicator_df=process_indicator_df.sort_values(by='Ppk').reset_index(drop=True)

    #conver process indicator dict to list name:
    name_list_dict={} # key: process name, value: list all dim in process name

    a=process_indicator_df
    for process_name in list(base_week.keys()):
      dim_infor_string_list=[]
      for dim_name in list(base_week[process_name].keys()):
        try:  
            dim_infor_string=str(dim_name) + ' Cp: ' + str(a[(a.Process_name==process_name) & (a.Dim_name==dim_name)]['Cp'].values[0]) + ' Pp: '+ \
            str(a[(a.Process_name==process_name) & (a.Dim_name==dim_name)]['Pp'].values[0]) +' Cpk: ' \
            + str(a[(a.Process_name==process_name) & (a.Dim_name==dim_name)]['Cpk'].values[0]) +' Ppk: '+ \
            str(a[(a.Process_name==process_name) & (a.Dim_name==dim_name)]['Ppk'].values[0]) 
    
            dim_infor_string_list.append(dim_infor_string)
        except: # if no process indicator (sigma=0, no Cp, Cpk...) so need continue (he qua cua try except tren)
            continue

      name_list_dict[process_name]=dim_infor_string_list
      #name_list.append(dim_name)
    return  base_week,process_indicator_df,name_list_dict

final_data,process_indicator_df,name_list_dict= create_process_indicator(final_data)

#------------------------Show process indicator-----------------------------#
st.subheader("Process indicator: Cp, Pb, Cpk, Ppk")
limit=st.number_input("Please input lower limit of Cpk and Ppk",value=1.33,key="input1")
limit=float(limit)
#limit = 1.33  # sigma: 4, Yield: 99.99%   
#@st.cache(allow_output_mutation=True) # Moi them vao
def hightlight_price(row):
    ret = ["" for _ in row.index]
    if row.Cpk < limit or row.Ppk < limit:
      ret[row.index.get_loc("Process_name")] = "background-color: yellow"
      ret[row.index.get_loc("Dim_name")] = "background-color: yellow"
    if row.Cpk < limit:    
      ret[row.index.get_loc("Cpk")] = "background-color: yellow"
    if row.Ppk < limit:  
      ret[row.index.get_loc("Ppk")] = "background-color: yellow"
    return ret

@st.cache(suppress_st_warning=True,allow_output_mutation=True)
def highlight_process_indicator(process_indicator_df): 
    return process_indicator_df.style.apply(hightlight_price, axis=1)

process_indicator_df_highlight=highlight_process_indicator(process_indicator_df)

st.text('Highlight yellow for dim below lower limit:')
st.dataframe(process_indicator_df_highlight)
#-----------------------Select process to show--------------------------------#
st.subheader('Select process to show:')
process_list=list(final_data.keys())
process_select = st.multiselect('Select process to show: ', process_list)

#st.write('You selected:', process_select)

#----------------------------------------
st.subheader('(Optional:) Please input start date and end date for data analysis')
st.write('Skip this step or leave it blank if you need full time range')
start_date= st.text_input("Please input start Date (Ex: 2018, 2018/07, 2018/07/31)")
end_date= st.text_input("Please input end Date (Ex:  2019, 2019/08, 2018/08/20)")

#------------------------Control Chart group by hour--------------------------------------#
st.header("Control chart group by hour")

#@st.cache(suppress_st_warning=True,allow_output_mutation=True)
@st.cache(suppress_st_warning=True,allow_output_mutation=True)
def line_chart_by_hour(final_data,process_select,name_list_dict):
    fig_all=[]
    for process_name in process_select:
      df_dict=final_data[process_name] # process name
      i=1
      #Layout
      fig = make_subplots(          # Dim name
          rows=len(df_dict), cols=1,
          #shared_xaxes=True, # share same axis
          #vertical_spacing=0.05, # adjust spacing between charts
          #column_widths=[0.8, 0.2],
          subplot_titles=(name_list_dict[process_name]) # dict with key is process name and value is list of dim (contain name, cp, cpk...)
      )
      for name in list(df_dict.keys()): #also group
        df=df_dict[name].copy()
        df=df.sort_values(by=['Hour'])
        for a in df.columns[1:]:
          df[a] = df[a].round(decimals=3)
        df=df.reset_index(drop=True)
        # Draw control chart
        df_group=df.groupby('Hour').mean()
        #df=df.set_index('Date')
        if start_date != '' and end_date != '':
            df_group=df_group[start_date:end_date]
        #Control chart 1 
        fig.append_trace(go.Scatter(
                                x=df_group.index, y=df_group['Value'],mode='lines+markers',
                                name='mean ', 
                                line=dict( color='#4280F5')
                                ),row=i, col=1)
        #USL, LSL
        fig.append_trace(go.Scatter(x=df_group.index, y=df_group['USL'],name='USL ', line=dict( color='#FF5733'),mode='lines'),row=i, col=1)
        fig.append_trace(go.Scatter(x=df_group.index, y=df_group['LSL'],name='LSL ',line=dict( color='#FF5733'),mode='lines'),row=i, col=1)
        #fig.append_trace(go.Scatter(x=df_group.index, y=df_group['Mean'],name='Nominal ',line=dict( color='#FF5733'),mode='lines'),row=i, col=1)
        # UCL, LCL
        fig.append_trace(go.Scatter(x=df_group.index, y=df_group['UCL'],name='UCL ', line=dict( color='#33C2FF'),mode='lines'),row=i, col=1)
        fig.append_trace(go.Scatter(x=df_group.index, y=df_group['LCL'],name='LCL ', line=dict( color='#33C2FF'),mode='lines'),row=i, col=1)
        fig.append_trace(go.Scatter(x=df_group.index, y=df_group['Mean'],name='Mean ', line=dict( color='#33C2FF'),mode='lines'),row=i, col=1)
        i=i+1

      if len(df_dict)>1:
        fig.update_layout(height=200*len(df_dict), width=1200, title_text='Process: '+process_name)
      else:
        fig.update_layout(height=300, width=1200, title_text='Process: '+process_name)
      #fig update each process (contain a lot of dim inside)
      #fig.show()
      fig_all.append(fig)  
    return fig_all # fig nay chi la 1 process cuoi cung thoi, lam sao return toan bo fig ?


if st.checkbox('Please tick to this box if you need to show control chart'):
    fig_line_chart=line_chart_by_hour(final_data,process_select,name_list_dict)    
    for fig in fig_line_chart:
        st.plotly_chart(fig)
        
#------------------------Box plot Chart by hour--------------------------------------#

st.header("Box chart by hour")
@st.cache(suppress_st_warning=True,allow_output_mutation=True)
def box_chart(final_data,process_select,name_list_dict,type_chart):
    fig_all=[]
    for process_name in process_select:
      df_dict=final_data[process_name] # process name
      i=1
      #Layout
      fig = make_subplots(          # Dim name
          rows=len(df_dict), cols=1,
          #shared_xaxes=True, # share same axis
          #vertical_spacing=0.05, # adjust spacing between charts
          #column_widths=[0.8, 0.2],
          subplot_titles=list(df_dict.keys()) # dict with key is process name and value is list of dim (contain name, cp, cpk...)
      )
      for name in list(df_dict.keys()): #also group
        df=df_dict[name].copy()
        df=df.sort_values(by=['Hour'])
        if type_chart== 'Hour':
            for a in df.columns[1:]:
              df[a] = df[a].round(decimals=3)
            df=df.reset_index(drop=True)
            # Draw control chart
            df=df.set_index('Hour')
        if type_chart == 'Day':
            df['Day']=df['Hour'].apply(lambda x: x.strftime("%Y %m %d"))
            df['Day']=pd.to_datetime(df['Day'])
            #df['Day']=df['Hour'].dt.strftime("%Y %m %d")
            cols = df.columns.tolist()
            cols = cols[-1:] + cols[:-1]
            df=df[cols]
            for a in df.columns[2:]:
                df[a] = df[a].round(decimals=3)
            df=df.reset_index(drop=True)
            # Draw control chart
            df=df.set_index('Day')
        if type_chart == 'Week':
            df['Week']=df['Hour'].dt.strftime('%Y-w%U')
            #df['Week']=pd.to_datetime(df['Week'])
            cols = df.columns.tolist()
            cols = cols[-1:] + cols[:-1]
            df=df[cols]
            for a in df.columns[2:]:
              df[a] = df[a].round(decimals=3)
            df=df.reset_index(drop=True)
            # Draw control chart
            df=df.set_index('Week')
            
        if start_date != '' and end_date != '':
            df=df[start_date:end_date]
        #Control chart 1 
        fig.append_trace(go.Box(
                                x=df.index, y=df['Value'],name='value ', 
                                line=dict( color='#4280F5'),boxpoints='all'
                                ),row=i, col=1)
        #USL, LSL
        fig.append_trace(go.Scatter(x=df.index, y=df['USL'],name='USL ', line=dict( color='#FF5733'),mode='lines'),row=i, col=1)
        fig.append_trace(go.Scatter(x=df.index, y=df['LSL'],name='LSL ',line=dict( color='#FF5733'),mode='lines'),row=i, col=1)
        #fig.append_trace(go.Scatter(x=df['Datef'], y=df['Nominal'],name='Nominal '+name,line=dict( color='#FF5733')),row=i, col=1)
        # UCL, LCL
        #fig.append_trace(go.Scatter(x=df_group.index, y=df_group['UCL'],name='UCL ', line=dict( color='#33C2FF'),mode='lines'),row=i, col=1)
        #fig.append_trace(go.Scatter(x=df_group.index, y=df_group['LCL'],name='LCL ', line=dict( color='#33C2FF'),mode='lines'),row=i, col=1)
        #fig.append_trace(go.Scatter(x=df_group.index, y=df_group['Mean'],name='Mean ', line=dict( color='#33C2FF'),mode='lines'),row=i, col=1)
        i=i+1

      if len(df_dict)>1:
        fig.update_layout(height=200*len(df_dict), width=1200, title_text='Process: '+process_name)
      else:
        fig.update_layout(height=300, width=1200, title_text='Process: '+process_name)
      #fig update each process (contain a lot of dim inside)
      #fig.show()
      fig_all.append(fig)  
    return fig_all # fig nay chi la 1 process cuoi cung thoi, lam sao return toan bo fig ?

if st.checkbox('Show box chart by hour'):
    fig_box_chart_hour=box_chart(final_data,process_select,name_list_dict,'Hour')    
    for fig in fig_box_chart_hour:
        st.plotly_chart(fig)
    
#------------------------Box plot Chart by day--------------------------------------#
st.header("Box chart by day")

if st.checkbox('Show box chart by day'):
    fig_box_chart_day=box_chart(final_data,process_select,name_list_dict,'Day')    
    for fig in fig_box_chart_day:
        st.plotly_chart(fig)
    
#------------------------Box plot Chart by week--------------------------------------#
st.header("Box chart by week")

if st.checkbox('Show box chart by week'):
    fig_box_chart_week=box_chart(final_data,process_select,name_list_dict,'Week')    
    for fig in fig_box_chart_week:
        st.plotly_chart(fig)

#---------------------------DataFrame----------------------------
st.header("Data Frame")
if st.checkbox('Show dataframe'):
#if st.button("Show dataframe"):
    # selectbox
    option1 = st.selectbox(
    'Which process to show ?',list(final_data.keys()))
    option2 = st.selectbox(
    'Which dim to show ?',list(final_data[option1].keys()))
    st.dataframe(final_data[option1][option2])

if st.checkbox('Save dataframe'):
    for name in final_data[option1].keys():
        final_data[option1][name].to_csv(save_folder+name+'.csv')
        st.write('Path file: ',save_folder+name+'.csv')

#-----------------PART 2----------------#
#---------------------------------------#
#---------------------------------------#
#---------------------------------------#
#---------------------------------------#
st.header("2. Applied control limit to new data")

#-----------------Design_layout main side-----------------#
trial_path2='//Vn01w2k16v18/data/Copyroom/Test_software/Data/Control plan/Control plan set up 3000/Applied to new data'
#trial_path='//Vn01w2k16v18/data/Copyroom/Test_software/Data/Control plan/Control plan 3000'
#trial_path='/media/ad/01D6B57CFBE4DB20/1.Linux/Data/Process_control/Control plan 3000'
#trial_path='/media/ad/01D6B57CFBE4DB20/1.Linux/Data/Process_control/Control plan 2600'
#trial_path='/media/ad/01D6B57CFBE4DB20/1.Linux/Data/Process_control/Control plan E series 1'

input_folder2= st.text_input("1. Please input folder name for data analysis (MUST)",trial_path2,key="applied_data")
input_folder2=input_folder2 +'/'
#path='//Vn01w2k16v18/data/Copyroom/Test_software/Data/Membrane 3000S/'
#st.write('path input: '+path)

st.write('Accept any public folders like copy room or QA folder, currently not allow private folder')
save_folder2=st.text_input("2. Please input folder name for saving (option)",'//Vn01w2k16v18/data/Copyroom/Test_software/Data/Save',key="applied_save")
save_folder2=save_folder2+'/'

#---------------sort file by created_time-------------------------# 

all_files_xlsx2=glob.glob(input_folder2 + '*.xlsx')
all_files_xlsm2=glob.glob(input_folder2 + '*.xlsm')
all_files_combine2=all_files_xlsx2+all_files_xlsm2
#sort file in directory by reverse:
all_files_combine2 = sorted(all_files_combine2, reverse = False)
st.text('number of files: '+str(len(all_files_combine2)))
st.text(all_files_combine2)
# Get master sheet dataframe:
#@st.cache(suppress_st_warning=True)
master_sheet_data_2=master_sheet_data_func(all_files_combine2)

#---------------process data-------------------------# 
#@st.cache(suppress_st_warning=True,allow_output_mutation=True)
all_process_week2=create_sheet_dict(all_files_combine2) # Run function above
st.text('number of week: '+str(len(all_process_week2)))

#-----------Concat all week to one base week-------------------------#

#@st.cache(suppress_st_warning=True,allow_output_mutation=True)
#@st.cache(suppress_st_warning=True,allow_output_mutation=True)
final_data_2=concat_baseweek(master_sheet_data_2,all_process_week2)    
#print(base_week['39682'])
st.text('number of process in line: '+str(len(final_data_2)))

#-----------Calculate process indicator base on 25 lates subgroup sample-------------------
@st.cache(allow_output_mutation=True) # Mo them vao sau nay
def process_performance2(df):
  constants={
    2:1.128,3:1.693,4:2.059,5:2.326,6:2.534,7:2.704,8:2.847, 9: 2.970,10: 3.078,
    11: 3.173,12: 3.258,13: 3.336,14: 3.407,15: 3.472,16: 3.532,17:3.588,18:3.640,
    19:3.689,20:3.735,
    }  
  #print('dim: ',name)
  # Calculate sigma
  sigma=np.std(df.Value)
  n=df.Hour.value_counts()[0] # check sub group num
  # Method 1: Collect only 25 group sample for calculation # Sẽ có 33 dim name / 14 process
  #num_sample=n*25
  #df_temp=df[-num_sample:]
  #df_temp=df_temp.reset_index(drop=True)
  # Method 2: Not collect 25 group sample, use full #  có 33 dim name / 14 process
  df_temp=df.reset_index(drop=True) # Tai sao phải reset index ?

  # Calculae usl, lsl
  usl=df_temp.USL[0]
  lsl=df_temp.LSL[0]
  m=df_temp.Value.mean() 
    
  if sigma ==0 :
      #Mean=df_temp['Value'].mean()
      #UCL=LCL=Mean
      Cp=float("NaN")
      Cpk=float("NaN")
      Pp=float("NaN")
      Ppk=float("NaN")
  if sigma !=0:
      #UCL, LCL, Mean
      k=3 
      #UCL=df_temp['Value'].mean() + sigma*k
      #LCL=df_temp['Value'].mean() - sigma*k
      #Mean=df_temp['Value'].mean()
      #Ppk
      print("sigma",sigma)
      Pp = float(usl - lsl) / (6*sigma)
      Ppu = float(usl - m) / (3*sigma)
      Ppl = float(m - lsl) / (3*sigma)
      Ppk = np.min([Ppu, Ppl])
      #print('Pp:{:.2f} , Ppk: {:.2f}'.format(Pp,Ppk))
      #Cpk
      temp=df_temp.groupby('Hour').agg({'Value':['min','max']})
      temp['Range']=temp['Value','max']-temp['Value','min']
      Range=temp['Range'].mean()

      if n <= 20:
        sigma_within = Range/constants[n]
      else:
        sigma_within = Range/constants[20]

      Cp = float(usl - lsl) / (6*sigma_within)
      Cpu = float(usl - m) / (3*sigma_within)
      Cpl = float(m - lsl) / (3*sigma_within)
      Cpk = np.min([Cpu, Cpl])
      #print('Cp:{:.2f} , Cpk:{:.2f}'.format(Cp,Cpk))
      if np.isnan(usl):
        Cpk=Cpl
        Ppk=Ppl
      elif np.isnan(lsl):
        Cpk=Cpu
        Ppk=Ppu
      else:
        Cpk = np.min([Cpu, Cpl])
        Ppk = np.min([Ppu, Ppl])
  #print(Ppk,Cpk)
  Cp=round(Cp,2)
  Cpk=round(Cpk,2)
  Pp=round(Pp,2)
  Ppk=round(Ppk,2)
  #UCL=round(UCL,2)
  #LCL=round(LCL,2)
  #Mean=round(Mean,2)
  return Cp,Cpk,Pp,Ppk

#-----------create_process_indicator-------------------
@st.cache(suppress_st_warning=True,allow_output_mutation=True)
def create_process_indicator2(base_week,process_indicator_old_df): 

  process_indicator_dict={}
  process_indicator_df=pd.DataFrame(columns=['Process_name','Dim_name','Cp','Pp','Cpk','Ppk','UCL','LCL','Mean'])
  i=0
  for process_name in list(base_week.keys()):
    #print(process_name)
    df_dict=base_week[process_name]
    for dim_name in list(df_dict.keys()): #also group
      #print(dim_name)
      df=df_dict[dim_name]
      #print(dim_name)
      try: # debug purpose
        Cp,Cpk,Pp,Ppk=process_performance2(df) 
        UCL=process_indicator_old_df[(process_indicator_old_df["Process_name"]==process_name) &(process_indicator_old_df["Dim_name"]==dim_name)]["UCL"].values[0]
        LCL=process_indicator_old_df[(process_indicator_old_df["Process_name"]==process_name) &(process_indicator_old_df["Dim_name"]==dim_name)]["LCL"].values[0]
        Mean=process_indicator_old_df[(process_indicator_old_df["Process_name"]==process_name) &(process_indicator_old_df["Dim_name"]==dim_name)]["Mean"].values[0]
        df["UCL"]= UCL
        df["LCL"]= LCL
        df["Mean"]= Mean
      except:
        continue # sigma = 0 phai continue de tranh dung chuong trinh
        #print(process_name)
        #print(dim_name)
        #print(df)
      #process_indicator_dict[process_name]=[dim_name,Cp,Pp,Cpk,Ppk]
      process_indicator_df.loc[i]=process_name,dim_name,Cp, Pp, Cpk, Ppk,UCL,LCL,Mean
      i+=1

  #process_indicator_df=process_indicator_df.sort_values(by='Ppk').reset_index(drop=True)

  #conver process indicator dict to list name:
  name_list_dict={} # key: process name, value: list all dim in process name

  a=process_indicator_df
  for process_name in list(base_week.keys()):
    dim_infor_string_list=[]
    for dim_name in list(base_week[process_name].keys()):
      try:  
          dim_infor_string=str(dim_name) + ' Cp: ' + str(a[(a.Process_name==process_name) & (a.Dim_name==dim_name)]['Cp'].values[0]) + ' Pp: '+ \
          str(a[(a.Process_name==process_name) & (a.Dim_name==dim_name)]['Pp'].values[0]) +' Cpk: ' \
          + str(a[(a.Process_name==process_name) & (a.Dim_name==dim_name)]['Cpk'].values[0]) +' Ppk: '+ \
          str(a[(a.Process_name==process_name) & (a.Dim_name==dim_name)]['Ppk'].values[0]) 
  
          dim_infor_string_list.append(dim_infor_string)
      except: # if no process indicator (sigma=0, no Cp, Cpk...) so need continue (he qua cua try except tren)
          continue

    name_list_dict[process_name]=dim_infor_string_list
    #name_list.append(dim_name)
  return  base_week,process_indicator_df,name_list_dict

final_data_2,process_indicator_df_2,name_list_dict_2= create_process_indicator2(final_data_2,process_indicator_df)

#------------------------Show process indicator-----------------------------#
st.subheader("Process indicator: Cp, Pb, Cpk, Ppk")
limit=st.number_input("Please input lower limit of Cpk and Ppk",value=1.33,key="input2")
limit=float(limit)
#limit = 1.33  # sigma: 4, Yield: 99.99%   
#@st.cache(allow_output_mutation=True) # Moi them vao

process_indicator_df_highlight2=highlight_process_indicator(process_indicator_df_2)

st.text('Highlight yellow for dim below lower limit:')
st.dataframe(process_indicator_df_highlight2)
#-----------------------Select process to show--------------------------------#
st.subheader('Select process to show:')
process_list2=list(final_data_2.keys())
process_select2 = st.multiselect('Select process to show: ', process_list2)

#st.write('You selected:', process_select)

#----------------------------------------
st.subheader('(Optional:) Please input start date and end date for data analysis')
st.write('Skip this step or leave it blank if you need full time range')
start_date= st.text_input("Please input start Date (Ex: 2018, 2018/07, 2018/07/31)",key="start_2")
end_date= st.text_input("Please input end Date (Ex:  2019, 2019/08, 2018/08/20)",key="end_2")

#------------------------Control Chart group by hour--------------------------------------#
st.header("Control chart group by hour")

#@st.cache(suppress_st_warning=True,allow_output_mutation=True)
#@st.cache(suppress_st_warning=True,allow_output_mutation=True)
if st.checkbox('Please tick to this box if you need to show control chart',key="box_2"):
    fig_line_chart2=line_chart_by_hour(final_data_2,process_select2,name_list_dict_2)    
    for fig in fig_line_chart2:
        st.plotly_chart(fig)
        
#------------------------Box plot Chart by hour--------------------------------------#

st.header("Box chart by hour")
if st.checkbox('Show box chart by hour',key="box_3"):
    fig_box_chart_hour2=box_chart(final_data_2,process_select2,name_list_dict_2,'Hour')    
    for fig in fig_box_chart_hour2:
        st.plotly_chart(fig)
    
#------------------------Box plot Chart by day--------------------------------------#
st.header("Box chart by day")

if st.checkbox('Show box chart by day',key="box_4"):
    fig_box_chart_day2=box_chart(final_data_2,process_select2,name_list_dict_2,'Day')     
    for fig in fig_box_chart_day2:
        st.plotly_chart(fig)
    
#------------------------Box plot Chart by week--------------------------------------#
st.header("Box chart by week")

if st.checkbox('Show box chart by week',key="box_5"):
    fig_box_chart_week2=box_chart(final_data_2,process_select2,name_list_dict_2,'Week')    
    for fig in fig_box_chart_week2:
        st.plotly_chart(fig)

#---------------------------DataFrame----------------------------
st.header("Data Frame")
if st.checkbox('Show dataframe',key="box_6"):
#if st.button("Show dataframe"):
    # selectbox
    option1 = st.selectbox(
    'Which process to show ?',list(final_data_2.keys()))
    option2 = st.selectbox(
    'Which dim to show ?',list(final_data_2[option1].keys()))
    st.dataframe(final_data_2[option1][option2])

if st.checkbox('Save dataframe',key="box_7"):
    for name in final_data_2[option1].keys():
        final_data_2[option1][name].to_csv(path3+name+'.csv')
        st.write('Path file: ',path3+name+'.csv')