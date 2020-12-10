# add something from vs code dnn
# add something 2
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
st.markdown('<style>h1{color: green;}</style>', unsafe_allow_html=True)
st.title('Process quality control')

st.subheader('Created by: DNN')
st.header("Information")

text2= st.text_input("1. Please input folder name for data analysis (MUST)",'//Vn01w2k16v18/data/Copyroom/Test_software/Data/Control plan/Control plan 3000')
path=text2 +'/'
#path='//Vn01w2k16v18/data/Copyroom/Test_software/Data/Membrane 3000S/'
#st.write('path input: '+path)

st.write('Accept any public folders like copy room or QA folder, currently not allow private folder')
text3=st.text_input("2. Please input folder name for saving (option)",'//Vn01w2k16v18/data/Copyroom/Test_software/Data/Save')
path3=text3+'/'

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

#path='//Vn01w2k16v18/data/Copyroom/Test_software/Data/Control plan/Control plan 3000'
#path='//Vn01w2k16v18/data/Copyroom/Test_software/Data/Control plan/Control plan 2600'
#path='//Vn01w2k16v18/data/Copyroom/Test_software/Data/Control plan/Control plan E series'
path=path+'/'
#print(path)
all_files1=glob.glob(path + '*.xlsx')
all_files2=glob.glob(path + '*.xlsm')
all_files=all_files1+all_files2
#sort file in directory by reverse:
all_files = sorted(all_files, reverse = False)
st.text('number of files: '+str(len(all_files)))
st.text(all_files)

#---------------process data-------------------------# 
#@st.cache(suppress_st_warning=True,allow_output_mutation=True)
@st.cache(suppress_st_warning=True,allow_output_mutation=True)
def create_sheet_dict(all_files):
    all_process_week={}
    for path_name in all_files: # read each excel file
      week_name=path_name[-16:-5]
      print(week_name,'-'*20)
      xls = xlrd.open_workbook(path_name, on_demand=True)
      sheet_names=xls.sheet_names()
      #print(xls.sheet_names())
      sheet_dict={}
      xls = pd.ExcelFile(path_name) 
      for name in sheet_names: # read each sheet in excel file
        sheet_dict[name] = pd.read_excel(xls, name)
        #sheet_dict[name].to_excel(name+'.xlsx')
      #break
      # Read all process in one week
      sheet_all={} # most important (reset sheet_all to empty)
      sheet_error=[]
      for name_sheet in sheet_names[1:]: # read each sheet name (process)
        #print(name_sheet)
        sheet=sheet_dict[name_sheet]

        # find begin and end col
        values_col=sheet.iloc[22,:]
        values_col.reset_index(drop=True,inplace=True)  # Date, MSVN, DIM A,B,C....
        #print(values_col)
        #begin_col=values_col[values_col=='Kích thước\nDimension'].index
        begin_col=values_col[values_col.str.contains('Dime')==True].index #sửa lại ngày 2/11/2021, begin col la col bat dau lay du lieu
        begin_col=begin_col+1
        end_col=values_col[values_col.str.contains('MSNV|OP')==True].index # end coil la col dung lai sau khi ket thuc lay du lieu
        #if (not begin_col.values or not end_col.values): #check empty value in begin and end col (khong co Dime or MSVN/OP, thuong la cac cong doang can do khoi luong)
        #    continue
        if (not begin_col.values): continue #check empty value in begin col only (not end coil due to end coil always have MSNV or OP)
        #print(begin_col,end_col)

        df_dict={} # add all value, USL, LSL, UCL... in each process sheet
        #try:
        for name in sheet.columns[begin_col[0]:end_col[-1]]: # all dim in each process (dim a, b, c,d ....)
          dim_name=sheet[name][22]
          #print(dim_name)
          df=pd.DataFrame()
          df_dict[dim_name]= {} # them vao ngay 20/11, lưu tất cả dim name ke cả khi value = None

          #tolerance_dict[sheet[name][22]]=[sheet[name][24],sheet[name][23]]  
          try:  
              df['Hour']=sheet[sheet.columns[9]][25:]
              #print(df['Date'])
              df['Hour']=df['Hour'].apply(lambda x: x.strftime("%Y %m %d %H")) # group theo hour, gần như trùng với tần suất lấy mẫu đo control plan
              df['Hour']=pd.to_datetime(df['Hour'])
          except:
            continue
          df['Value']=sheet[name][25:]
          #print(df['Value'])
          #if np.std(df.Value) == 0: # chuyển qua dim khac nếu các giá trị là giống nhau
          #  continue
          #print('susessful')
          if isinstance(sheet[name][23],str): continue #check string USL 25/11
          if isinstance(sheet[name][23],str): continue #check string LSL 25/11
          df['USL']=sheet[name][23] # max
          df['LSL']=sheet[name][24] # min
          #print(df['USL'],df['LSL'])  
          df.dropna(subset=['Value'],inplace=True)
          df=df[pd.to_numeric(df['Value'], errors='coerce').notnull()] #25/11: loai bo duy nhật 1 cột Value có not numeric value truoc khi convert
          #df=df[df.columns[1:]].apply(pd.to_numeric, errors='coerce').dropna() # loai bo tat ca not numeric value truoc khi convert
          df[df.columns[1:]]=df[df.columns[1:]].astype('float32')

          df_dict[dim_name]=df.reset_index(drop=True)

        sheet_all[name_sheet]=df_dict # add each process name to shee_all dict each week
        #sheet_all chi chứa name_sheet (process), khong chứa dim name nếu như không có dữ liệu, lưu ý thông tin này khi concat dữ liệu dưới đây
      all_process_week[week_name]=sheet_all
      print(week_name,sheet_error)
    return all_process_week

all_process_week2=create_sheet_dict(all_files) # Run function above
st.text('number of week: '+str(len(all_process_week2)))

#-----------Concat all week to one base week-------------------------#

#@st.cache(suppress_st_warning=True,allow_output_mutation=True)
@st.cache(suppress_st_warning=True,allow_output_mutation=True)
def concat_baseweek(all_process_week):
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
                  if len(base_week[process_name_base][dim_name]) == 0: #20/11 convert dict to df for append
                     base_week[process_name_base][dim_name] = pd.DataFrame.from_dict(base_week[process_name_base][dim_name])
                  if len(other_week[process_name_base][dim_name]) == 0: #20/11 convert dict to df for append
                    other_week[process_name_base][dim_name] = pd.DataFrame.from_dict(other_week[process_name_base][dim_name])
                  #try:   
                  #base_week[process_name_base][dim_name]=pd.concat([base_week[process_name_base][dim_name],other_week[process_name_base][dim_name]])
                  base_week[process_name_base][dim_name]=base_week[process_name_base][dim_name].append(other_week[process_name_base][dim_name]) 
                  # sử dụng hàm append có thể concat empty dataframe tất cả các week
                  #except:continue
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

final_data=concat_baseweek(all_process_week2)    
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
  n=df.Hour.value_counts()[0]
  num_sample=n*25
  df_temp=df[-num_sample:]
  df_temp=df_temp.reset_index(drop=True)
  usl=df_temp.USL[0]
  lsl=df_temp.LSL[0]
  m=df_temp.Value.mean() 

  #Ppk
  sigma=np.std(df.Value)
  Pp = float(usl - lsl) / (6*sigma)
  Ppu = float(usl - m) / (3*sigma)
  Ppl = float(m - lsl) / (3*sigma)
  Ppk = np.min([Ppu, Ppl])
  #print('Pp:{:.2f} , Ppk: {:.2f}'.format(Pp,Ppk))

  #UCL, LCL, Mean
  k=3
  df['UCL']=df_temp['Value'].mean() + sigma*k
  df['LCL']=df_temp['Value'].mean() - sigma*k
  df['Mean']=df_temp['Value'].mean()
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
  return Cp,Cpk,Pp,Ppk

#-----------create_process_indicator-------------------
@st.cache(suppress_st_warning=True,allow_output_mutation=True)
def create_process_indicator(base_week): 

    process_indicator_dict={}
    process_indicator_df=pd.DataFrame(columns=['Process_name','Dim_name','Cp','Pp','Cpk','Ppk'])
    #process_indicator_df.columns=['Dim','Cp','Cpk','Pb','Ppk']
    i=0
    for process_name in list(base_week.keys()):
      #print(process_name)
      df_dict=base_week[process_name]
      for dim_name in list(df_dict.keys()): #also group
        #print(dim_name)
        df=df_dict[dim_name]
        #print(dim_name)
        #try: # object column cannot be calculated process indicator (OK/Not OK, all value have the same...) How to remove object colum in the beginning ?
        Cp,Cpk,Pp,Ppk=process_performance(df) 
        #except: continue
        #process_indicator_dict[process_name]=[dim_name,Cp,Pp,Cpk,Ppk]
        process_indicator_df.loc[i]=process_name,dim_name,Cp, Pp, Cpk, Ppk
        i+=1

    #process_indicator_df=process_indicator_df.sort_values(by='Ppk').reset_index(drop=True)

    #conver process indicator dict to list name:
    name_list_dict={} # key: process name, value: list all dim in process name

    a=process_indicator_df
    for process_name in list(base_week.keys()):
      dim_infor_string_list=[]
      for dim_name in list(base_week[process_name].keys()):
        dim_infor_string=str(dim_name) + ' Cp: ' + str(a[(a.Process_name==process_name) & (a.Dim_name==dim_name)]['Cp'].values[0]) + ' Pp: '+ \
        str(a[(a.Process_name==process_name) & (a.Dim_name==dim_name)]['Pp'].values[0]) +' Cpk: ' \
        + str(a[(a.Process_name==process_name) & (a.Dim_name==dim_name)]['Cpk'].values[0]) +' Ppk: '+ \
        str(a[(a.Process_name==process_name) & (a.Dim_name==dim_name)]['Ppk'].values[0]) 

        dim_infor_string_list.append(dim_infor_string)

      name_list_dict[process_name]=dim_infor_string_list
      #name_list.append(dim_name)
    return  process_indicator_df,name_list_dict

process_indicator_df,name_list_dict= create_process_indicator(final_data)

#------------------------Show process indicator-----------------------------#
st.subheader("Process indicator: Cp, Pb, Cpk, Ppk")
limit=st.number_input("Please input lower limit of Cpk and Ppk",value=1.33)
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

process_indicator_df=highlight_process_indicator(process_indicator_df)

st.text('Highlight yellow for dim below lower limit:')
st.dataframe(process_indicator_df)
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

#------------------------Box plot Chart by hour--------------------------------------#

st.header("Box chart by hour")
@st.cache(suppress_st_warning=True,allow_output_mutation=True)
def box_chart(process_select):
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
        for a in df.columns[1:]:
          df[a] = df[a].round(decimals=3)
        df=df.reset_index(drop=True)
        # Draw control chart
        df=df.set_index('Hour')
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
    fig_all_2=box_chart(process_select)    
    for fig in fig_all_2:
        st.plotly_chart(fig)
    
#------------------------Box plot Chart by day--------------------------------------#
st.header("Box chart by day")
@st.cache(suppress_st_warning=True,allow_output_mutation=True)
def box_chart_2(process_select):
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

if st.checkbox('Show box chart by day'):
    fig_all_2=box_chart_2(process_select)    
    for fig in fig_all_2:
        st.plotly_chart(fig)
    
#------------------------Box plot Chart by week--------------------------------------#
st.header("Box chart by week")
@st.cache(suppress_st_warning=True,allow_output_mutation=True)
def box_chart_2(process_select):
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
        df['Week']=df['Hour'].dt.strftime('%Y-%U')
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

if st.checkbox('Show box chart by week'):
    fig_all_2=box_chart_2(process_select)    
    for fig in fig_all_2:
        st.plotly_chart(fig)
#------------------------Line Chart group by shift--------------------------------------#
st.header("Control chart group by hour")

#@st.cache(suppress_st_warning=True,allow_output_mutation=True)
@st.cache(suppress_st_warning=True,allow_output_mutation=True)
def line_chart(process_select):
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
        #fig.append_trace(go.Scatter(x=df_group.index, y=df_group['Nominal'],name='Nominal ',line=dict( color='#FF5733'),mode='lines'),row=i, col=1)
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


if st.checkbox('Please tick to this box if you need to show control chart'):
    fig_all=line_chart(process_select)    
    for fig in fig_all:
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
        final_data[option1][name].to_csv(path3+name+'.csv')
        st.write('Path file: ',path3+name+'.csv')