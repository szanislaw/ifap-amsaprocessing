import tkinter as tk
from tkinter import filedialog
from collections import OrderedDict
import pandas as pd
import os
import glob
import re
import configparser
import math
import xlsxwriter
import shutil

def open_file():    
    global data

    file_path = filedialog.askopenfilename()
    file_label.config(text=file_path)
    data = pd.read_excel(file_path)

def open_file1():    
    global DMC_Code

    DMC_file_path = filedialog.askopenfilename()
    file_label_1.config(text=DMC_file_path)
    DMC_Code = pd.read_csv(DMC_file_path)
    
def split_file():
    global data
    global DMC_Code

    data[['id', 'DMC', 'temp1']] = data['chip_id'].str.split('=', expand=True)
    data[['DMC', 'Temp2']] = data['DMC'].str.split(',', expand=True)     

    DMC_Code = DMC_Code['DMC'].tolist()

    data = data.loc[data['DMC'].isin(DMC_Code)]

#for Akari
    data_columns = {
        '3001': '20Hz',
        '3002': '35Hz',
        '3003': '80Hz',
        '3004': '300Hz',
        '3005': '900Hz',
        '3006': '1000Hz',
        '3007': '1100Hz',
        '3008': '3000Hz',
        '3009': '8000Hz',
        '3010': '10000Hz',
    }

    data_columns1 = {
        '4001': '75Hz',
        '4002': '1000Hz',
        '4003': '3000Hz',
        '4004': '10000Hz',
    }
    
    data_columns2 = {
        '5001': '94dBSPL',
        '5002': '100dBSPL',
        '5003': '106dBSPL',
        '5004': '112dBSPL',
        '5005': '118dBSPL',
        '5006': '124dBSPL',
        '5007': '127dBSPL',
        '5008': '130dBSPL',
    }

    with pd.ExcelWriter('AMSA.xlsx') as writer:
        for col, sheet in data_columns.items():
            data1 = data.loc[:,['site_num','DMC',f'{col};Output_{sheet}_94dBSPL_NPM;value;p']]
            data1.rename(columns={f'{col};Output_{sheet}_94dBSPL_NPM;value;p': 'SENS'}, inplace=True)

            data1.to_excel(writer, sheet_name=sheet, index=False)

            for col, sheet in data_columns1.items():
                data2 = data.loc[:,['site_num','DMC',f'{col};Phase_{sheet}_94dBSPL_NPM;value;p']]
                data2.rename(columns={f'{col};Phase_{sheet}_94dBSPL_NPM;value;p': 'PHASE'}, inplace=True)

                data2.to_excel(writer, sheet_name='Phase_'+sheet, index=False)

        for col, sheet in data_columns2.items():
            data3 = data.loc[:,['site_num','DMC',f'{col};THD_1000Hz_{sheet}_NPM;value;p']]
            data3.rename(columns={f'{col};THD_1000Hz_{sheet}_NPM;value;p': 'THD'}, inplace=True)

            data3.to_excel(writer, sheet_name='THD_'+sheet, index=False)

    tk.messagebox.showinfo("Success", "File processing completed!")

def calibration_setup_file():

    global file_path_ini

    file_path_ini = filedialog.askopenfilename()
    inifile_label.config(text=file_path_ini)
    data = pd.read_csv(file_path_ini)

def run_calib_setup_file():
    fpath=file_path_ini
    
    def inifile(file):
        global ini_data

        fpath=file
        
        config = configparser.ConfigParser()
        config.read(fpath)

        data = []
        for section in config.sections():
            for key in config[section]:
                value = config[section][key]
                data.append([section, key, value])

        ini_data = pd.DataFrame(data, columns=['Freq', 'Test', 'Value'])
        ini_data = ini_data.loc[(ini_data['Test'] == 'calibrationtarget') | (ini_data['Test'] == 'goldenphase'), :]

        split_data=ini_data['Freq'].str.replace('Name','Freq').str.split('=',expand=True)

        ini_data = ini_data.drop(columns='Freq',axis=1)
        split_data = split_data.drop(columns=0,axis=1)

        ini_data=pd.concat([ini_data,split_data],axis=1)
        ini_data=ini_data.rename(columns={1:'Freq'})
        
        ini_data = ini_data[ini_data['Value'] != '#'].reset_index(drop=True)
        ini_data['Value'] = ini_data['Value'].astype(float)
    
    fpath = "Acoustic_Chambers_Calibration_Data/CalibrationSetupFile.csv"

    calib_setup = pd.read_csv(fpath, header=None)
    column_list = calib_setup.iloc[0, :].tolist()
    calib_setup.columns = column_list

    column_list = [round(num) for num in column_list[1:] if not math.isnan(num)]
    folder_path = os.path.dirname(fpath)
    regex = re.compile(".*MicrophoneCharacterization.*")
    writer = pd.ExcelWriter('New_ini_Target.xlsx', engine='xlsxwriter')
    
    for value in column_list:
        search_pattern = os.path.join(folder_path, f"*{value}*")
        folders = glob.glob(search_pattern)
        if len(folders) > 0:
            for folder in folders:
                print(f"Found folder for value {value}: {folder}")
                files = glob.glob(os.path.join(folder, "*.ini"))
                if len(files) > 0:
                    for file in files:
                        match = regex.search(file)
                        if match:
                            inifile(file)
                            sheetname=f'{value}'
                            ini_data.to_excel(writer, sheet_name=sheetname, index=False)
                else:       
                    print(f"No files found in folder {folder}")
        else:
            print(f"No folders found for value {value}")
    writer.close()
    
    AMSA_file_path='AMSA.xlsx'
    data=pd.read_excel(AMSA_file_path,sheet_name=None)

    writer = pd.ExcelWriter('AMSA1.xlsx', engine='xlsxwriter')

    for sheet_name,df in data.items():
        if 'SENS' in df.columns:
            df = df.groupby('site_num')['SENS'].mean().reset_index()
        elif df.columns.str.contains('PHASE').any():
            df = df.groupby('site_num')[df.columns[df.columns.str.contains('PHASE')]].mean().reset_index()
        elif df.columns.str.contains('THD').any():
            df = df.groupby('site_num')[df.columns[df.columns.str.contains('THD')]].mean().reset_index()

        sheet_name = sheet_name.replace('Hz', '')
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        
    writer.close()
        
    amsa_doc = pd.read_excel('AMSA1.xlsx', sheet_name=None)
    new_ini_doc = pd.read_excel('New_ini_Target.xlsx', sheet_name=None)
    writer = pd.ExcelWriter('AMSA2.xlsx', engine='xlsxwriter')
    ordered_sheets = OrderedDict()

    for sheet_name, ini_sheet in new_ini_doc.items():
        print(sheet_name)
        if sheet_name != '94':
            continue
        
        freq_values = ini_sheet['Freq'].unique()

        for freq in freq_values:
            value = ini_sheet.loc[ini_sheet['Freq'] == freq, 'Value'].values[0]  # Freq Amplitude Target in inifile value
            value2=ini_sheet.loc[(ini_sheet['Freq'] == freq) & (ini_sheet['Test']=='goldenphase'), 'Value'].values[0]  #Phase Target in inifile value
            freq=int(freq)
            Phase=f'Phase_{freq}'

            processed = False

            for sheet_name2, amsa_sheet in amsa_doc.items():
                if str(freq)==sheet_name2 and not processed:
                    amsa_sheet.loc[:, 'New_Target'] = value
                    sheet_name3 = '{}'.format(freq)
                    amsa_sheet.to_excel(writer, sheet_name=sheet_name3, index=False)
                    ordered_sheets[sheet_name3] = amsa_sheet
                    #marked sheet is processed 
                    processed = False

            for sheet_name2, amsa_sheet in amsa_doc.items():
                if str(Phase)==sheet_name2:
                    amsa_sheet.loc[:, 'Phase_New_Target'] = value2
                
                    sheet_name4 = '{}'.format(Phase)
                    amsa_sheet.to_excel(writer, sheet_name=sheet_name4, index=False)
                    ordered_sheets[sheet_name4] = amsa_sheet
                    processed = True

            for sheet_name2, amsa_sheet in amsa_doc.items():
                if 'THD' in sheet_name2:
                    amsa_sheet.to_excel(writer, sheet_name=sheet_name2, index=False)
                    ordered_sheets[sheet_name2] = amsa_sheet
                else:
                    continue
    writer.close()
    
    fpath="Acoustic_Chambers_Calibration_Data/CalibrationSetupFile.csv"
    folder_path = os.path.dirname(os.path.dirname(fpath))

    calib_setup = pd.read_csv(fpath,header=None)

    column_list=list(calib_setup.loc[calib_setup[0]==1000].iloc[0,:])
    calib_setup.columns=column_list

    column_list =  [x for x in column_list if not math.isnan(x)]
    column_list=[round(num) for num in column_list]
    column_list=column_list[1::]
    
    ACC_setup_file_Path = ("Acoustic_Chambers_Calibration_Data/CalibrationSetupFile.csv")

    dir_path = os.path.dirname(ACC_setup_file_Path)

    Old_ACC_Folder = "Acoustic_Chambers_Calibration_Data"
    New_ACC_Folder=shutil.copytree(Old_ACC_Folder, Old_ACC_Folder+'_new')
    New_ACC_Path = os.path.join(os.path.dirname(Old_ACC_Folder), 'Acoustic_Chambers_Calibration_Data_New')
    
    regex = re.compile(".*CalibSpeakersVoltageRMS.*")   # search the file in ACC folder (Speaker voltage file) 
    regex1 = re.compile(".*SystemPhase_94dBSPL.*")      # search the file in ACC folder (SystemPhase file) 

    writer1 = pd.ExcelWriter('output_spk.xlsx', engine='xlsxwriter')
    writer2 = pd.ExcelWriter('output_Phase.xlsx', engine='xlsxwriter')

    for value in column_list:
        search_pattern = os.path.join(New_ACC_Folder, f"*{value}*")
        folders = glob.glob(search_pattern)
        if len(folders) > 0:
            for folder in folders:
                print(f"Found folder for value {value}: {folder}")
                
                files = glob.glob(os.path.join(folder, "*.csv"))
                if len(files) > 0:
                    for file in files:
                        match = regex.search(file)
                        if match and 'Real' not in file:
                            print(f"Found file: {file}")
                            
                            data=pd.read_csv(file)
                            sheet_name = str(value)
                            data.to_excel(writer1,sheet_name=sheet_name,index=False) 

                    for file in files:
                        match = regex1.search(file)
                        if match:
                            print(f"Found file: {file}")
                            data = pd.read_csv(file)
                            sheet_name = str(value)

                            data.to_excel(writer2,sheet_name=sheet_name,index=False) 
                else:
                    print(f"No files found in folder {folder}")
        else:
            print(f"No folders found for value {value}")
            
    writer1.close()
    writer2.close()
    
    #Phase compensation and ACC file system phase compensation and saved

    folder_path = "output_Phase.xlsx"
    data = pd.read_excel(folder_path, sheet_name=None)

    AMSA2_Path = 'AMSA2.xlsx'
    AMSA2 = pd.read_excel(AMSA2_Path, sheet_name=None)

    if '94' in data.keys():
        data_94 = data['94'] 
        print(data_94)

        for sheet_name, sheet_data in AMSA2.items():
            if 'Phase' in sheet_name:
                match = re.search(r'\d+', sheet_name)
                
                if match:
                    numeric_part = match.group()
                                        
                    if numeric_part in data_94.columns.values:
                        data_94_column = data_94[numeric_part]

                        sheet_data['SystemPhase'] = data_94_column
                        sheet_data['deltaPhase']=sheet_data['PHASE']-sheet_data['Phase_New_Target']
                        sheet_data['New_SystemPhase']=sheet_data['SystemPhase']+sheet_data['deltaPhase']
                        sheet_data[numeric_part]=sheet_data['New_SystemPhase']
                        
                        for root, dirs, files in os.walk(New_ACC_Path):
                            for file in files:

                                if file=='SystemPhase_94dBSPL.csv':
                                    Spk94DBSPL_file_path=os.path.join(root,file)
                                    Spk94DBSPL_data=pd.read_csv(Spk94DBSPL_file_path)

                                    if numeric_part in Spk94DBSPL_data.columns:
                                        Spk94DBSPL_data.loc[:, numeric_part] = sheet_data[numeric_part]
                                        Spk94DBSPL_data.to_csv(Spk94DBSPL_file_path, index=False)
                                    else:
                                        continue

    with pd.ExcelWriter(AMSA2_Path, engine='openpyxl') as writer:
        for sheet_name, sheet_data in AMSA2.items():
            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)
            
    #amplitude compensation and ACC file Speaker voltage / RefMic MIC Amplitude compensation and saved

    folder_path_1 = "output_spk.xlsx"
    data_1 = pd.read_excel(folder_path_1, sheet_name=None)

    AMSA2_Path = 'AMSA2.xlsx'
    AMSA2 = pd.read_excel(AMSA2_Path, sheet_name=None)

    if '94' in data_1.keys():
        data_94_1 = data_1['94'] 
        print(data_94_1)


        for sheet_name, sheet_data in AMSA2.items():
            if sheet_name in data_94_1.columns.values:            
                data_94_1_column = data_94_1[sheet_name]
        
                sheet_data['spkVol'] = data_94_1_column
                sheet_data['deltaSens']=sheet_data['New_Target']-sheet_data['SENS']
                sheet_data['vol_ratio']=10**(sheet_data['deltaSens']/20)
                sheet_data['New_spkVol']=sheet_data['vol_ratio']*sheet_data['spkVol']
                sheet_data[sheet_name]=sheet_data['vol_ratio']*sheet_data['spkVol']

                for root, dirs, files in os.walk(New_ACC_Path):
                    for file in files:
                        if file=='CalibSpeakersVoltageRMS_94dBSPL.csv':
                            Spk94DBSPL_file_path=os.path.join(root,file)
                            Spk94DBSPL_data=pd.read_csv(Spk94DBSPL_file_path)

                            if sheet_name in Spk94DBSPL_data.columns:
                                Spk94DBSPL_data.loc[:, sheet_name] = sheet_data[sheet_name]
                                Spk94DBSPL_data.to_csv(Spk94DBSPL_file_path, index=False)

                        if file=='ReferenceMicAmplitude_94dBSPL.csv':
                            Ref_output_file=os.path.join(root,file)
                            Ref_output_data=pd.read_csv(Ref_output_file)
                        
                            if sheet_name in Ref_output_data.columns:

                                Ref_output_data.loc[:,sheet_name]=Ref_output_data.loc[:,sheet_name]+sheet_data['deltaSens']
                                Ref_output_data.to_csv(Ref_output_file,index=False)

    with pd.ExcelWriter(AMSA2_Path, engine='openpyxl') as writer:
        for sheet_name, sheet_data in AMSA2.items():
            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)
            
    AMSA2_Path = 'AMSA2.xlsx'
    AMSA2 = pd.read_excel(AMSA2_Path, sheet_name=None)

    for sheet_name, sheet_data in AMSA2.items():
        if sheet_name=='1000':
            data_94_THD = sheet_data['New_spkVol']

    for sheet_name, sheet_data in AMSA2.items():
        if 'THD' in sheet_name:
            match = re.search(r'\d+', sheet_name)
            if match:
                numeric_part = int(match.group())

                sheet_data['94dBspl'] = data_94_THD
                sheet_data['Delta_SPL'] = numeric_part - 94
                sheet_data['Sens']=10**(sheet_data['Delta_SPL']/20)*sheet_data['94dBspl']
                sheet_data[numeric_part]=sheet_data['Sens']

                for root, dirs, files in os.walk(New_ACC_Path):
                    for file in files:
                        if file=='CalibSpeakersVoltageRMS_'+str(numeric_part)+'dBSPL.csv' and file!='CalibSpeakersVoltageRMS_94dBSPL.csv' : 
                            Spk94DBSPL_file_path=os.path.join(root,file)
                            Spk94DBSPL_data=pd.read_csv(Spk94DBSPL_file_path)

                            if '1000' in Spk94DBSPL_data.columns:
                                Spk94DBSPL_data.loc[:, '1000'] = sheet_data[numeric_part]
                                Spk94DBSPL_data.to_csv(Spk94DBSPL_file_path, index=False)
                                
    with pd.ExcelWriter(AMSA2_Path, engine='openpyxl') as writer:
        for sheet_name, sheet_data in AMSA2.items():
            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)
            
    #Upon successful completion of the above steps, the new AMSA2.xlsx file is generated, and the new ACC folder is generated. The new ACC folder is renamed to the original ACC folder name, and the original ACC folder is renamed to the original ACC folder name + _old. The new ACC folder is copied to the original ACC folder name, and the original ACC folder is deleted. The new AMSA2.xlsx file is copied to the original AMSA2.xlsx file name, and the original AMSA2.xlsx file is deleted.
    tk.messagebox.showinfo("Success", "File processing completed!")
    
def window_size():
    global window

    window = tk.Tk()
    window.wm_state('zoomed')
    window.config(background='sky blue')

    project_option()
    AMSA_Icon()
    DMC_Code_loading()
    Calib_setup_file_loading()
    
    window.mainloop()

def project_option():
    
    var=tk.StringVar(window)
    var.set('please choose project ')

    option_menu=tk.OptionMenu(window,var, 'project : Akari', 'Project : Squid', 'Project : Kassandra',)
    option_menu.config(font=("Calibri", 13))
    option_menu.place(relx=0.3,rely=0.05,relwidth=0.2, relheight=0.1)

def AMSA_Icon():
    global file_label

    frame = tk.Frame(bg="skyblue", bd=0, highlightthickness=1, highlightbackground="gray")
    frame.place(relx=0.04, rely=0.18, relwidth=0.8, relheight=0.2)

    file_label = tk.Label(text=" ", background='sky blue')
    file_label.place(relx=0.2, rely=0.2)

    button_font=("Calibri", 10)
    open_button = tk.Button(text='please input GD AMSA file\n请载入GD AMSA 文件', command=open_file, font=button_font)
    open_button.place(relx=0.05, rely=0.2)
    
    edit_button = tk.Button(text='processing AMSA file\n处理AMSA文件', command=split_file, font=button_font)
    edit_button.place(relx=0.7, rely=0.25)


def DMC_Code_loading():
    global file_label_1

    file_label_1 = tk.Label(text=" ", background='sky blue')
    file_label_1.place(relx=0.2, rely=0.3)

    button_font=("Calibri", 10)
    open_button = tk.Button(text='please input DMC code file\n请载入DMC code 文件', command=open_file1, font=button_font)
    open_button.place(relx=0.05, rely=0.3)


def Calib_setup_file_loading():
    global inifile_label
    
    frame = tk.Frame(bg="skyblue", bd=0, highlightthickness=1, highlightbackground="gray")
    frame.place(relx=0.04, rely=0.38, relwidth=0.8, relheight=0.1)
    
    inifile_label=tk.Label(text="",background='sky blue')
    inifile_label.place(relx=0.2,rely=0.4)

    button_font=("Calibri", 10) 
    open_button=  tk.Button(text='please input Calib_setup. file\n请输入 Calib_setup 文件',command=calibration_setup_file,font=button_font)
    open_button.place(relx=0.05, rely=0.4)

    edit_button = tk.Button(text='processing Calib_setup file\n处理setup文件', command=run_calib_setup_file,font=button_font)
    edit_button.place(relx=0.7, rely=0.4)

window_size()
