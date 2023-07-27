import os
import glob
import re
import configparser
import pandas as pd
import numpy as np
import xlwings as xl
import math
import shutil
import matplotlib.pyplot as plt
from collections import OrderedDict


fpath = "CalibrationSetupFile.csv"

folder_path = os.path.dirname(os.path.dirname(fpath))

calib_setup = pd.read_csv(fpath, header=None)
column_list = list(calib_setup.loc[calib_setup[0] == 1000].iloc[0, :])
calib_setup.columns = column_list
column_list = [x for x in column_list if not math.isnan(x)]
column_list = [round(num) for num in column_list[1::]]


def inifile(file):
    global ini_data
    fpath = file
    config = configparser.ConfigParser()
    config.read(fpath)

    data = []
    for section in config.sections():
        for key in config[section]:
            value = config[section][key]
            data.append([section, key, value])

    ini_data = pd.DataFrame(data, columns=['Freq', 'Test', 'Value'])
    ini_data = ini_data.loc[(ini_data['Test'] == 'calibrationtarget') | (ini_data['Test'] == 'goldenphase'), :]
    split_data = ini_data['Freq'].str.replace('Name', 'Freq').str.split('=', expand=True)
    ini_data = ini_data.drop(columns='Freq', axis=1)
    split_data = split_data.drop(columns=0, axis=1)
    ini_data = pd.concat([ini_data, split_data], axis=1)
    ini_data = ini_data.rename(columns={1: 'Freq'})
    ini_data = ini_data[ini_data['Value'] != '#'].reset_index(drop=True)
    ini_data['Value'] = ini_data['Value'].astype(float)
    print(ini_data)


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
                        sheetname = f'{value}'
                        ini_data.to_excel(writer, sheet_name=sheetname, index=False)
            else:
                print(f"No files found in folder {folder}")
    else:
        print(f"No folders found for value {value}")

writer.close()


AMSA_file_path = 'AMSA.xlsx'
data = pd.read_excel(AMSA_file_path, sheet_name=None)

writer = pd.ExcelWriter('AMSA1.xlsx', engine='xlsxwriter')

for sheet_name, df in data.items():
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
        value = ini_sheet.loc[ini_sheet['Freq'] == freq, 'Value'].values[0] 
        value2 = ini_sheet.loc[(ini_sheet['Freq'] == freq) & (ini_sheet['Test'] == 'goldenphase'), 'Value'].values[0]  
        freq = int(freq)
        Phase = f'Phase_{freq}'

        processed = False

        for sheet_name2, amsa_sheet in amsa_doc.items():
            if str(freq) == sheet_name2 and not processed:
                amsa_sheet.loc[:, 'New_Target'] = value
                sheet_name3 = '{}'.format(freq)
                amsa_sheet.to_excel(writer, sheet_name=sheet_name3, index=False)
                ordered_sheets[sheet_name3] = amsa_sheet
                processed = False

        for sheet_name2, amsa_sheet in amsa_doc.items():
            if Phase in sheet_name2:
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

ACC_setup_file_Path = "CalibrationSetupFile.csv"

dir_path = os.path.dirname(ACC_setup_file_Path)

print(dir_path)

Old_ACC_Folder = "Acoustic_Chambers_Calibration_Data"
New_ACC_Folder = shutil.copytree(Old_ACC_Folder, Old_ACC_Folder + '_new')
New_ACC_Path = os.path.join(os.path.dirname(Old_ACC_Folder), 'Acoustic_Chambers_Calibration_Data_New')

regex = re.compile(".*CalibSpeakersVoltageRMS.*") 
regex1 = re.compile(".*SystemPhase_94dBSPL.*")

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
                        data = pd.read_csv(file)
                        sheet_name = str(value)
                        data.to_excel(writer1, sheet_name=sheet_name, index=False)

                for file in files:
                    match = regex1.search(file)
                    if match:
                        print(f"Found file: {file}")
                        data = pd.read_csv(file)
                        sheet_name = str(value)
                        data.to_excel(writer2, sheet_name=sheet_name, index=False)

            else:
                print(f"No files found in folder {folder}")
    else:
        print(f"No folders found for value {value}")

writer1.close()
writer2.close()

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
                    sheet_data['deltaPhase'] = sheet_data['PHASE'] - sheet_data['Phase_New_Target']
                    sheet_data['New_SystemPhase'] = sheet_data['SystemPhase'] + sheet_data['deltaPhase']
                    sheet_data[numeric_part] = sheet_data['New_SystemPhase']

                    for root, dirs, files in os.walk(New_ACC_Path):
                        for file in files:
                            if file == 'SystemPhase_94dBSPL.csv':
                                Spk94DBSPL_file_path = os.path.join(root, file)
                                Spk94DBSPL_data = pd.read_csv(Spk94DBSPL_file_path)
                                if numeric_part in Spk94DBSPL_data.columns:
                                    print('yes')
                                    Spk94DBSPL_data.loc[:, numeric_part] = sheet_data[numeric_part]
                                    Spk94DBSPL_data.to_csv(Spk94DBSPL_file_path, index=False)
                                else:
                                    continue

with pd.ExcelWriter(AMSA2_Path, engine='openpyxl') as writer:
    for sheet_name, sheet_data in AMSA2.items():
        sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)

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
            sheet_data['deltaSens'] = sheet_data['New_Target'] - sheet_data['SENS']
            sheet_data['vol_ratio'] = 10 ** (sheet_data['deltaSens'] / 20)
            sheet_data['New_spkVol'] = sheet_data['vol_ratio'] * sheet_data['spkVol']
            sheet_data[sheet_name] = sheet_data['vol_ratio'] * sheet_data['spkVol']

            for root, dirs, files in os.walk(New_ACC_Path):
                for file in files:
                    if file == 'CalibSpeakersVoltageRMS_94dBSPL.csv':
                        Spk94DBSPL_file_path = os.path.join(root, file)
                        Spk94DBSPL_data = pd.read_csv(Spk94DBSPL_file_path)
                        print(Spk94DBSPL_data)

                        if sheet_name in Spk94DBSPL_data.columns:
                            print('yes')
                            Spk94DBSPL_data.loc[:, sheet_name] = sheet_data[sheet_name]
                            Spk94DBSPL_data.to_csv(Spk94DBSPL_file_path, index=False)

                    if file == 'ReferenceMicAmplitude_94dBSPL.csv':
                        Ref_output_file = os.path.join(root, file)
                        Ref_output_data = pd.read_csv(Ref_output_file)
                        if sheet_name in Ref_output_data.columns:
                            Ref_output_data.loc[:, sheet_name] = Ref_output_data.loc[:, sheet_name] + sheet_data['deltaSens']
                            Ref_output_data.to_csv(Ref_output_file, index=False)

with pd.ExcelWriter(AMSA2_Path, engine='openpyxl') as writer:
    for sheet_name, sheet_data in AMSA2.items():
        sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)

AMSA2_Path = 'AMSA2.xlsx'
AMSA2 = pd.read_excel(AMSA2_Path, sheet_name=None)

for sheet_name, sheet_data in AMSA2.items():
    if sheet_name == '1000':
        data_94_THD = sheet_data['New_spkVol']

for sheet_name, sheet_data in AMSA2.items():
    if 'THD' in sheet_name:
        match = re.search(r'\d+', sheet_name)
        if match:
            numeric_part = int(match.group())
            sheet_data['94dBspl'] = data_94_THD
            sheet_data['Delta_SPL'] = numeric_part - 94
            sheet_data['Sens'] = 10 ** (sheet_data['Delta_SPL'] / 20) * sheet_data['94dBspl']
            sheet_data[numeric_part] = sheet_data['Sens']

            for root, dirs, files in os.walk(New_ACC_Path):
                for file in files:
                    if file == f'CalibSpeakersVoltageRMS_{numeric_part}dBSPL.csv' and file != 'CalibSpeakersVoltageRMS_94dBSPL.csv':
                        Spk94DBSPL_file_path = os.path.join(root, file)
                        Spk94DBSPL_data = pd.read_csv(Spk94DBSPL_file_path)
                        print(numeric_part)
                        print(Spk94DBSPL_data)

                        if '1000' in Spk94DBSPL_data.columns:
                            print('yes')
                            Spk94DBSPL_data.loc[:, '1000'] = sheet_data[numeric_part]
                            Spk94DBSPL_data.to_csv(Spk94DBSPL_file_path, index=False)
                    elif file == 'CalibSpeakersVoltageRMS_94dBSPL.csv':
                        continue

with pd.ExcelWriter(AMSA2_Path, engine='openpyxl') as writer:
    for sheet_name, sheet_data in AMSA2.items():
        sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)
