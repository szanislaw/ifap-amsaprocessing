import tkinter as tk
from tkinter import filedialog
import pandas as pd


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
    
    DMC_Code=DMC_Code['DMC'].tolist()
  
    data = data.loc[data['DMC'].isin(DMC_Code)]
    
    data_columns = {     
        #Akari AMSA data amplitude / phase / THD
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

    data_columns1 ={
        '4001': '75Hz',
        '4002': '1000Hz',
        '4003': '3000Hz',
        '4004': '10000Hz',
        
    }

    data_columns2={
        '5001': '94dBSPL',
        '5002': '100dBSPL',
        '5003': '106dBSPL',
        '5004': '112dBSPL',
        '5005': '118dBSPL',
        '5006': '124dBSPL',
        '5007': '127dBSPL',
        '5008': '130dBSPL',
    }
    
    with pd.ExcelWriter('AMSA1.xlsx') as writer:
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
    inifile_lable.config(text=file_path_ini)
    data = pd.read_csv(file_path_ini)

def run_calib_setup_file():
    fpath=file_path_ini
    # %run A101_ini_target_generat.ipynb - hi GPT, this will be changed to include the functions in other ipynb files 
    # %run A201_AMSA1_good.ipynb - hi GPT, this will be changed to include the functions in other ipynb files 
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
    
    # New_ACC_Folder_Generation()
    # AMSA_file_1()
    window.mainloop()


def project_option():
    
    var=tk.StringVar(window)
    var.set('please choose project ')

    option_menu=tk.OptionMenu(window,var,
                              'project : Akari',
                              'Project : Squid',
                              'Project : Kassandra',
                            )
    option_menu.config(font=("Calibri", 13))
    option_menu.place(relx=0.3,rely=0.05,relwidth=0.2, relheight=0.1)

def AMSA_Icon():
    global file_label

    #create a frame for different step
    frame = tk.Frame(bg="skyblue", bd=0, highlightthickness=1, highlightbackground="gray")
    frame.place(relx=0.04, rely=0.18, relwidth=0.8, relheight=0.2)

    # create a label
    file_label = tk.Label(text=" ", background='sky blue')
    file_label.place(relx=0.2, rely=0.2)

    # create a button 
   
    button_font=("Calibri", 10)
    open_button = tk.Button(text='please input GD AMSA file\n请载入GD AMSA 文件', command=open_file, font=button_font)
    open_button.place(relx=0.05, rely=0.2)

    # create a button for edit
    
    edit_button = tk.Button(text='processing AMSA file\n处理AMSA文件', command=split_file, font=button_font)
    edit_button.place(relx=0.7, rely=0.25)

def DMC_Code_loading():
    global file_label_1

    # create a label
    file_label_1 = tk.Label(text=" ", background='sky blue')
    file_label_1.place(relx=0.2, rely=0.3)

    # create a button    
    button_font=("Calibri", 10)
    open_button = tk.Button(text='please input DMC code file\n请载入DMC code 文件', command=open_file1, font=button_font)
    open_button.place(relx=0.05, rely=0.3)


def Calib_setup_file_loading():
    global inifile_lable

    #create a frame for different step
    frame = tk.Frame(bg="skyblue", bd=0, highlightthickness=1, highlightbackground="gray")
    frame.place(relx=0.04, rely=0.38, relwidth=0.8, relheight=0.1)

    #creat a lable 
    inifile_lable=tk.Label(text="",background='sky blue')
    inifile_lable.place(relx=0.2,rely=0.4)

    #create a button for inifile loading
    button_font=("Calibri", 10) 
    open_button=  tk.Button(text='please input Calib_setup. file\n请输入 Calib_setup 文件',command=calibration_setup_file,font=button_font)
    open_button.place(relx=0.05, rely=0.4)

    # create a button for edit
    # Generate New_ini_Target file.
    edit_button = tk.Button(text='processing Calib_setup file\n处理setup文件', command=run_calib_setup_file,font=button_font)
    edit_button.place(relx=0.7, rely=0.4)

window_size()



 

