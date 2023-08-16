import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from collections import OrderedDict
import pandas as pd
import os
import glob
import re
import configparser
import math
import xlsxwriter
import shutil
from PIL import Image, ImageTk

logo_path = r"iVBORw0KGgoAAAANSUhEUgAAAGQAAAAtCAYAAABYtc7wAAAAAXNSR0IArs4c6QAAAMBlWElmTU0AKgAAAAgABwESAAMAAAABAAEAAAEaAAUAAAABAAAAYgEbAAUAAAABAAAAagEoAAMAAAABAAIAAAExAAIAAAAPAAAAcgEyAAIAAAAUAAAAgodpAAQAAAABAAAAlgAAAAAAAABIAAAAAQAAAEgAAAABUGl4ZWxtYXRvciAzLjkAADIwMjA6MDU6MTQgMTQ6MDU6OTUAAAOgAQADAAAAAQABAACgAgAEAAAAAQAAAGSgAwAEAAAAAQAAAC0AAAAAbrCj3QAAAAlwSFlzAAALEwAACxMBAJqcGAAABCNpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IlhNUCBDb3JlIDUuNC4wIj4KICAgPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4KICAgICAgPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIKICAgICAgICAgICAgeG1sbnM6ZGM9Imh0dHA6Ly9wdXJsLm9yZy9kYy9lbGVtZW50cy8xLjEvIgogICAgICAgICAgICB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iCiAgICAgICAgICAgIHhtbG5zOmV4aWY9Imh0dHA6Ly9ucy5hZG9iZS5jb20vZXhpZi8xLjAvIgogICAgICAgICAgICB4bWxuczp0aWZmPSJodHRwOi8vbnMuYWRvYmUuY29tL3RpZmYvMS4wLyI+CiAgICAgICAgIDxkYzpzdWJqZWN0PgogICAgICAgICAgICA8cmRmOkJhZy8+CiAgICAgICAgIDwvZGM6c3ViamVjdD4KICAgICAgICAgPHhtcDpNb2RpZnlEYXRlPjIwMjAtMDUtMTRUMTQ6MDU6OTU8L3htcDpNb2RpZnlEYXRlPgogICAgICAgICA8eG1wOkNyZWF0b3JUb29sPlBpeGVsbWF0b3IgMy45PC94bXA6Q3JlYXRvclRvb2w+CiAgICAgICAgIDxleGlmOlBpeGVsWERpbWVuc2lvbj4xMDA8L2V4aWY6UGl4ZWxYRGltZW5zaW9uPgogICAgICAgICA8ZXhpZjpQaXhlbFlEaW1lbnNpb24+NDU8L2V4aWY6UGl4ZWxZRGltZW5zaW9uPgogICAgICAgICA8ZXhpZjpDb2xvclNwYWNlPjE8L2V4aWY6Q29sb3JTcGFjZT4KICAgICAgICAgPHRpZmY6Q29tcHJlc3Npb24+MDwvdGlmZjpDb21wcmVzc2lvbj4KICAgICAgICAgPHRpZmY6WFJlc29sdXRpb24+NzI8L3RpZmY6WFJlc29sdXRpb24+CiAgICAgICAgIDx0aWZmOk9yaWVudGF0aW9uPjE8L3RpZmY6T3JpZW50YXRpb24+CiAgICAgICAgIDx0aWZmOlJlc29sdXRpb25Vbml0PjI8L3RpZmY6UmVzb2x1dGlvblVuaXQ+CiAgICAgICAgIDx0aWZmOllSZXNvbHV0aW9uPjcyPC90aWZmOllSZXNvbHV0aW9uPgogICAgICA8L3JkZjpEZXNjcmlwdGlvbj4KICAgPC9yZGY6UkRGPgo8L3g6eG1wbWV0YT4KF1T4aQAAExFJREFUeAHtWwl0VUWarqp773sve0IWkpCNrWWRgYaGoXVUkNV21GFLS7cOKhJ0PB6ZGWZOnx5HkXO0T6s9Ot3SYEDapRcnENoFRJCmo4DaI8xg9xFUUCIkIftilvfeXarm++/LTV7CYojhACMFL3VvLX9V/X/9a9Xl7CJLDSNuSPyyTQ0yNJUkhZ2oJEsyNJGimEpWiiUxpmK4YrZSwuZcNXOmGqXizVzJJl0X1R1f2o0FE1gbLyuzL7Kl9Wk6ep9aDWAjxZj4fNjMBJ/J8gRnw4DwoYqrEYzxEYKx/FDYSjF0RvNSQvEWLli1ctgJxtkX7k/y447gNbrkbUypjpDNwlrAb8YE2y2bxUk9PqBYGcMwl2bi53va1YNnx4W5VYDdO1VxPokrNVYxMRL4TovjmuEDxpulLcEBFZjLfhDhgMOdw1zxjyzDqf2gPL21kG1yzvc8Lxb454UgxzKnFRiaMRkiZRZnfIpkarSfC18M1xie2ZcS0oSzw3h5H4goA4ccZKzjSG7Fe8GLBTEXah5fiyDpy0uv5dzIt+2OI40bbn2/NmfGXIeJ+yEvJuucp+vAOg0g8QsqJwx5fwASa4dk4k3GEv+SW7HpG0+A3oTvH0GK9hvprPwpJvR7uKZrzLHMDsFXf/H7/9xvavqbBsQQAQ4qaYII/8OVeE1wuXVwhX2Ys0tT2fZG3Pl67xdB0u/e/D2m69sYNLL7AwHAFbYu1VV/2fbMT2xNH4r332DSr2WdsA5eJkLfyXdOVtbRwbMzhjuNKcmGGAudwJTdaVmCMNDPusWc0QmmeV9dwFeXW7Gjse/TuNzSw0CfCILdzqtyZ92Kxqsbrbj9Lf7A+vSQiVIwGJwDBuJAbDmg0OFA/b5PPODnmo+4/w1/qxUaqWx4IMFQeeNvbvvyXGFc6u2/UmRVZk+/QmniCZ2JmzRoBkvJFpv7Zl75t8sWw+J/gGmGxmzLhOh6rLZ4wWogBBQ69zT43tKpUolfAM5o9Oaci5NSygfqixdsS72nZIrGfN/mtrmzZkPhsXOHfun0OCtBKrNnLRaCPQklnR2Gd0aYThA663CsrZmVvoUZ9941hUttqOD84+p18/67v8tOKSpJ0pR4Txj+0Uo6YDgNoOCZONZ1jKtWuIhlXPcnMjt4SFf+66qKb6rv71gXe7/Tiqz9bJKROWTQIxpn/4oFaCEQg+wmAyKqXTqbdS5+zNh2s27t9j2op9/XShpX14IILjGYtA9y6TzhMJlVn71wb0bl5nt4ICZRmkGwjT7GdIL5GOybQ5ATOd+NESz+F37Ol5pQ1uRDkD+B/RoMK/XwexUp/+F5zml3b14sDH0ic2BXMf+D2Lkd/aKMEt/iugGHxeXCX9UUL/itB0cu27KLm+1HuDDylbR2ch/rt47yYF7MObRxd1JsGvy5+LUBLpaGO4lhgDcgqhoszhbnVux6wiMGW6UEdPlqbsSuBIT5pi1JzvQrwX0MRDoqSCjeFA2kfv38TzVTXq24/Ou6rNR5db8sbIuu///23ENkVQ7RVwUYX0IiivQFcQY4pMlRvDCv4q3dp1l8UNlhahvuXTfo/tIcPahbtRtuqWGLSnypSfpwSDynfv28I9DZBD46dcWqwIkG2muTUprEgeLlZFer6l8V1iUuLXEG1bZkNboxr+7+6f9QkunAG23csKCC+qWlsREK0ceGT1OPsrLpp434Dr59RxyLab3SUTIAv/ZEzdoFn0dPpsdz0bNGGksbxxRPwKaoaShe+HGP+s4X0oO6w9Lrnis8SkVZ95XmS8tJdOzW8vqNS1tP1+d0ZV1KvTJ31s2aYqWINekkpqgCyto0pbo9v3JXySmdwSHpVaUHoWzHgSifKtv4Tv3GW1oJQczmG7ke+DbKS5im3mRKe4hJNRZKQHIpy7gmimrWza8lmGnLSm8WhrEWCjzb3QUMkV1Na2COc6Que/6tgz97KcaJidkm9MBw5ZiH61T9jax4uZV079YUvx1awwz/NRjnT0wXTzOHwcpTE92ZK+cDzYhZenLNjV9Ez33w8k0zJNMfh3ycCDGI5k4DROWLvjj5bxVPFfYI5WQs3zIeO+dpNLqaC92AwQEDQ23jMryidsMPajy46UWlj8IPW4T52eDx2wRXK2GZzFBKxYP7P3M4W9Gwbv7pNrQHoit3Rdaxgmk4a1A/BQFcYlAtOX6OVGtOS4yu7qc+qJBfIJo7FTs9E0T9IZjtVRBnKnREAlonMX/sLUrJf/F6CqGyue4DMYhpsBzdlyv88RPwLNgqLq3kNIhCPoEJLQdlBMNNgTbDxK65Gi8oZ7OZLd/gvpjpXPPRmUki88XOsK3gI5HWkb8Zy18eL5n2MsI9E2k8IBCWgkhlRuAfw+3ioei2g+4uzQFOSoHoaS4xyPkVWgI3/Lcq4Xsxu+j1WK89RbCZplMEewS07ZvMiFvMNF8G5h0Lt2AcNuHPsVHjvfZny12C+G3f3we4NoqUOCUoEgYdckwo+7GzdT5dnWa1tGBSx10FLbRU7BCpwu0vIra4Fe0VHEjCxVVs2h9dcamkfF+ZHbuxaMwf+soKbWWhjkdxGPIMwVfhNsxGWQQPvSMTRHnNS3PaFZ2RgOJMCBCKx6pQewnGKcEADolSEOwacjYJDjpDzhkPgqPTwI0dAHo7fJ3vwlDYEdkM/L60u0quiLSFaamxfwbywZU2XCPreSB9gZL2PhgwtGlmW9xe7LWVXB2lckwEKldPV2b7H6UZWoeObRgLAlqN0hxfntf+bLkAdwQgpu6K2FORpqTIkV7Irio7Z/MyJyZoYrFBLBYTkUHB9UV16xctCenaHa54QCVSoKDAPYRidesLDyrHeQU7DPQQEJPidzXrbn6wft2iMmp41iQZZDONw0wmVFHd+oXf98fKO4B8WGVgLMViGsJtrsGQUrQpF4ifDSRBWsmy2mcX/rr22fkfciZ+BkRLIDlBaGI6jZd8x+9JYsyjTcCVc9xWcgXabhGOtRL9XX0J7foD0lnu/DiPRBRoTGn/PI8VzKlfv/BeTGA7OAwbA7KameDcr04iYPv/CjgabbkigzYVaWiQVcq3vrr7qS0ONKVIrJm8O6xfNYzOOridWiVYBuQzWVAYAQcl0T1xcNVlXGBzYAV9S+AQjAP+UU5TrOTuOBVPLQqBEA2RlQA/muXuLo2L8ZhTogvZPX+JjMGF+gyIdy03kGo8lfoMewxmmeeuAYdmTcWFLVSuiZY/gyAniaOwtnFpCc5gKqcXF3Ek1hR740Dxdyy3nPEamt+5JOgMe5oBgYfYh9uPTjBgfbQHDOPYuQA6Q1t+tGUMIZj4mRLky/lInFtA49kgA1F53DDApK5oGZmxfPNtTHLuODKD4jTETlDG2QQDT98CtwIVQKbi1R7cquyiUHplaTVKC0DwNEQW0lFX5dVT3mtDnRs10J8E9wTZyR3dgIVCeK9LXneXX8JPXECpAj+ki4S2yP1BwsACIQ5DlYAhyF0Ogv4axHG6SQn8F+pa9cOgVRGLvLvK0RrUVTdADzp4zeghPzAmph1rhzntli7TboDGu3BglOzcrYRJpwEmLMQnLdVNhAKKDp3ofO/OoJG7Xh55BJ3Hen1AXCLnwCZc3mBHfJiJ0zkusQVMXn9IOLDn2f8O7HAXEBpX7TQ6TF4ITvsZyxRPq5iObuSizqcbEdlPxoIgTFCUgndaaXj9aAzng0TkHWoEqq8ZpQOadBiIux3BfoSZ9dS0jM/HSM8N6GgXEBjURAX8DkgmqBop45ufX3BGZEohyzkuYkTEFk/zpg2HxwfXwFXkoEeT0vQ6r26gcsH00AFY2uV01uEl8kcQ6b2+YsjMqV7ZJZ9L68/QH+2uhST4tQwhke41KT5o+StjvHfo8kNgjWbXRGZq/NhVJT6qsxL14WCLTFL2iLl92tAQHHiC5B/f2wQr9GUKrXuJBCN8kQDMyof3T5oUNXGvxaWX1zaJcoigPa7lxMXkdJb2eGrR5lGpy16enF60+TlN2Tuylpbm08oa1i6qAg7QFsYXF6NqqrT7EavKk5qzEo5fjKt6uHydbSrEsenAJtdGZ9L3y6BSJ8lD9xJFe/1MzM2qTXnAKzs1p63i9unuGGmEd/zHdurVp7M8ih2pASlcanpKc683Kty6Xg3I/qAi1w7x2rr5qfPaVEim1KMweztcRAttBRb/ARzXPQjt3Ml9sTnSgONIifwkTfyMbtPgGSOoJ3Fi+iEU0BLqq6zwUQR1usW5wtaNzM/t3vXHKz8VD11Nej+4BMmp3F6Bh4ewsB6YopivpvjqyuyZt/TuGHmnkIZE7Ax5dCIBrBxguWc5UIe4Ccp573KB8yhEzggWQqrRoKLHgbjx/JlIMYd/g36gSK/x8e6WY7yoVFc8fy/iaMsQQThJUQGEUeLx8yO80SrN9sdNKUu85vVr578Nf2wF3r9EG7QNJLtet2N/inUtiQ4uQj9h/m5sB4onyirj3J0faGVygcuZfUhdHnJxxVsb7xoyc1Kc0O7BpTa3KxEEAY0YLvgLx7NnLsmr2vVqF0wE/uTdW36IgHAc7PdgY2s4cjiFnajuLSmSdjARJ4uhiqRDbqihKvtAaNDJCd/ndjAOq2stL3i7i90RsvidbVnvwmtWIZ841jUGHpqaWFtGgjNXctNgumxla4tAlOVuE8HslY5tP6YpPVzVHOx04IARrWR5ZPxAqKVeRsIanUBxRv/b7KKSvZZU12MD5GIuNbi4t7exeP6h6HHpGaH2tZl3lrzt+PlMcEs6wjFHuBPaEU0MagcG3ejGxLCnYCAdpTJKjq4/yS3r1yCIo8WlfBIpPf1fCvAKU78Sm7Y7VWXfFMtE6AWcFi4MQWR5JKVzETx3YMf8U07lH57t7nH5aSAwcCJn1khcJJwYdLR3exCEgNPlaKmrNT7Bl5jgPo8srhVG4lqxDdwO/Tires+AWxgDsbhLCQYdl2ss7gabK0N3Yl/Prnq9p2PkLYYuOWTlJD+ocY38Ex+u/rhVRD06J4GsPQym+fesil1biEZev8t53zEArrgOuvRvIC735J3Y+Y7Xk3B8xnTcPUVUT+KsZGTkjD2CezKR6fsBAHwVEa+f5lTuev+MQC5X9MBAVe6Ma6Tk8xAXrkT07HmoAESmu9NZCULNyvPnZPmk8xBE1Z0UUiHCEFk8bgkphfMP9YqSat2QSmfP5Xu83cj1nj5OuzohyR87F6GvhTCX2pUt1uSf3HnAq4/Ov5IgXmPy2oUQP4IhdCP8Fd0jDNnN9NEN3oln9sFqf8lh9rbcit2VXt9vYq7gUFdWp0zSBSvEIdi12MXHTCbX4OZO2dnw0WeCeEAqhsyagXPw+4H8OeCYAIVZyDwmQN5nCHACqsFRu2HzvwKX6d1vCnEah81M6giridC930NIeQ5Qkg4X8w9SsnXZVbv2ejg8W37OBPGAnci5fgqOW5eAFn9HV00JEJ06kkAji4z0DBEKv2rwzn74t285TPtTe1z7oVGf7OvztRhvvIsxb4Lv0G5poyEVJmF+MyDKr8IpSjryI9iw/+Vw8SK+AujyS/qyhn4TxANeM3TGYMcUs3GFoBCa5Wof5ylUh89kXeLQIQOFZCiHb2PCEa+CN34Q1sWHhnL2m0IdtgxRP/zzXe4xqQf3Ysrp3gHmk+wP+4bAJRiH53FAHO5qqQnYfOkwelgHsyEV+Dbo003lqv2dq/r5ed7XJkg04mpyrx9uKm26xtRclE9GgCePOIWMACIQ5aRziED0LyLuZBAVx7HLyiHejkmpjkDufo6JVdpKNOt2qDkzNbOFf7Spy7MHiAFN5A/4jNhE0xKJ2DDJuP+UpQktD2bkUByXD8OcczH9XFyTykAkw71A2IboEOJAH0Mk7cbKtiuuQzR//W9iBpQg0ViqyZgxOKTzcQZn12FhUzDpK1GfBb2D/URBKIgzrJISFgoxF0lUQj8QCxm+P+eyBTsPF9Rws0OoBhCyEbCawXGtjmLtEIUdaBhG4MwSktuwL9x4F77rxeftEsfkGA5xUsVFDA6dYgErDkhMxJDJGAbfv7NkwEwAzEQ8JwAhybhK63I0bRyKUpAB046DI7Q/hj8HMf93bM726U7gEDlzKB+wdN4I0nuGJzNvSMc5PS4PyElY1DjY4VegzQjsviQgJJaQQCsmHUQ5/Sh5E/TynmXRpd1t3Y7448Hw3qNz6hndG3NwqymOBze4HQIXR7z8KIoPg4j4TpLtjzOs8pTysjMebEXD7+9z9Jz6C6Nf/RRbJZqG7U1os0SukE6BFCwPV1nzsfgCICIHsjgVCI3HLw4DxGJX+jzxRwMSsiOki7zQu5fcReFPZHERVNMzERuGB4UdgnhvRR+6/oOLfewExv0Me+IY3k/Ylv1ZbFLC8bRPXmtDXTRoVJ/fRPO8KJMacYO/zg6mmJaWKIVIRAw3HiHzeESxMx3mpEKoJCPun4T3WOgfHKIpHBy5NADGOe6AqaCGC2zwitsQaW+2HXZS01WdZcuOgBCttjKadNtpyqzZ6Z61X5RIuDypC4+B/wMIrEtXUadYaQAAAABJRU5ErkJggg=="

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

            data3.to_excel(writer, sheet_name='THD_'+ sheet, index=False)

    messagebox.showinfo("ifap-amsaprocessing", "AMSA File Processing completed!")

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

    column_list = list(calib_setup.loc[calib_setup[0]==1000].iloc[0,:])
    calib_setup.columns=column_list

    column_list =  [x for x in column_list if not math.isnan(x)]
    column_list = [round(num) for num in column_list]
    column_list = column_list[1::]
    
    ACC_setup_file_Path = ("Acoustic_Chambers_Calibration_Data/CalibrationSetupFile.csv")

    dir_path = os.path.dirname(ACC_setup_file_Path)

    Old_ACC_Folder = "Acoustic_Chambers_Calibration_Data"
    New_ACC_Folder = shutil.copytree(Old_ACC_Folder, Old_ACC_Folder+'_new')
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
                            data = pd.read_csv(file)
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
                    messagebox.showinfo("Error", "No files found in folder {folder}")
        else:
            messagebox.showinfo("Error", "No folders found for value {value}")
            
    writer1.close()
    writer2.close()
    
    #Phase Compensation and ACC file (SystemPhase / SystemPhase_94dBSPL.csv)
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
            
    #Amplitude compensation and ACC file Speaker voltage / RefMic MIC Amplitude compensation and saved
    folder_path_1 = "output_spk.xlsx"
    data_1 = pd.read_excel(folder_path_1, sheet_name=None)

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
    messagebox.showinfo("ifap-amsaprocessing", "Calibration Setup File Processing completed!")
    
def mainWindow():
    global window

    window = tk.Tk()
    window.wm_state('zoomed')

    projectOption()
    processAMSA()
    loadDMCCode()
    loadCalibSetupFile()
    
    logo_image = r"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"
    logo_photo = tk.PhotoImage(data=logo_image)
    logo_label = tk.Label(window, image=logo_photo,)
    logo_label.image = logo_photo  
    logo_label.place(relx=0.75, rely=0.05) 
    
    window.mainloop()   

def projectOption():
    var = tk.StringVar(window)
    var.set('Please select a project')

    frame_width = 0.2
    frame_height = 0.15

    screen_width = window.winfo_screenwidth()

    frame_x = (screen_width - frame_width * screen_width) / 2
    frame = tk.Frame(bd=0, highlightthickness=1, highlightbackground="black", highlightcolor="black")
    frame.place(relx=frame_x/screen_width, rely=0.18, relwidth=frame_width, relheight=frame_height)

    logo_image = tk.PhotoImage(data=logo_path)
    logo_label = tk.Label(frame, image=logo_image)
    logo_label.image = logo_image
    logo_label.pack(side="top", pady=5)

    project_combobox = ttk.Combobox(frame, textvariable=var, values=['Akari', 'Squid', 'Kassandra'], font=("Calibri", 12))
    project_combobox.set('Please select a project')
    project_combobox.pack(side="bottom", padx=10, pady=5, fill="x")

def processAMSA():
    global file_label

    frame = tk.Frame(bd=0, highlightthickness=1, highlightbackground="black", highlightcolor="black")
    frame.place(relx=0.04, rely=0.35, relwidth=0.9, relheight=0.15)

    file_label = tk.Label(text=" ")
    file_label.place(relx=0.2, rely=0.2)

    button_font=("Calibri", 12)
    open_button = tk.Button(text='Upload Raw Data (GD AMSA) file (.xlsx)', command=open_file, font=button_font)
    open_button.place(relx=0.05, rely=0.2)

    edit_button = tk.Button(text='Process Raw Data and GD Selection', command=split_file, font=button_font)
    edit_button.place(relx=0.8, rely=0.25)

def loadDMCCode():
    global file_label_1

    file_label_1 = tk.Label(text=" ",)
    file_label_1.place(relx=0.2, rely=0.5)

    button_font=("Calibri", 12)
    open_button = tk.Button(text='Upload GD DMC Code (.xlsx)', command=open_file1, font=button_font)
    open_button.place(relx=0.05, rely=0.5)

def loadCalibSetupFile():
    global inifile_label

    frame = tk.Frame(bd=0, highlightthickness=1, highlightbackground="black", highlightcolor="black")
    frame.place(relx=0.04, rely=0.65, relwidth=0.9, relheight=0.15)

    inifile_label=tk.Label(text="")
    inifile_label.place(relx=0.2,rely=0.4)

    button_font=("Calibri", 12)
    open_button=  tk.Button(text='Upload CalibrationSetupFile (.csv)', command=calibration_setup_file, font=button_font)
    open_button.place(relx=0.05, rely=0.4)

    edit_button = tk.Button(text='Process CalibrationSetupFile', command=run_calib_setup_file, font=button_font)
    edit_button.place(relx=0.8, rely=0.45)

mainWindow()
