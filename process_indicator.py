#-----------Calculate process indicator base on 25 lates subgroup sample-------------------
@st.cache(allow_output_mutation=True) # Mo them vao sau nay
def process_performance(df):
  constants={
    2:1.128,3:1.693,4:2.059,5:2.326,6:2.534,7:2.704,8:2.847, 9: 2.970,10: 3.078,
    11: 3.173,12: 3.258,13: 3.336,14: 3.407,15: 3.472,16: 3.532,17:3.588,18:3.640,
    19:3.689,20:3.735,
    }  
  #print('dim: ',name)
  n=df.Date.value_counts()[0]
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
  
  temp=df_temp.groupby('Date').agg({'Value':['min','max']})
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
@st.cache(allow_output_mutation=True)
def create_process_indicator(base_week): 

    #process_indicator_dict={}
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

process_indicator_df,name_list_dict= create_process_indicator(base_week)

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

#@st.cache(allow_output_mutation=True)
#def highlight_process_indicator(process_indicator_df): 
#   return process_indicator_df.style.apply(hightlight_price, axis=1)

process_indicator_df=highlight_process_indicator(process_indicator_df)

#st.text('Highlight yellow for dim below lower limit:')
#st.dataframe(process_indicator_df)# -*- coding: utf-8 -*-

