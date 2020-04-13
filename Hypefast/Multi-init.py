from GoogleAPI import Create_Service
import pandas as pd
import os
import glob

#Global Variable
main = os.getcwd()
CLIENT_SECRET_FILE = '/Users/sbahri/Documents/Hypefast/Brand Report Performance/credentials.json'
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
path = os.getcwd()
output_structured = glob.glob(path + '/Output/*/*.xls')
output = []
googlesheet = pd.DataFrame({'Brand': ['Motiviga', 'Monomom', 'Nona', 'Soleram', 'Sparse_label', 'Wearsidd', 'Wearstatusquo', 'Sabine'],
                            'SheetID': ['12wDobW_XjcjZ6v6ydMKe7cfCDVyisH14j3vvjtdPM0s', 
                                        '1y4lAIPjcsDhXsSDruoNtQ51cCjb3VmiFbCMkMwgUC1I',
                                        '1TYmZ01MbfgNE0rhq0I-hmXjZYBq6P6DN2-icWhcOPMI',
                                        '1bAcvfLlOg_8wiiqM78MwTJX60Km6AU4644RLk9MNkoA',
                                        '16-ft6bhopp7Rv6wLiq1motynrBnwkBo5EeGzNHn-UuI',
                                        '1JeGQxmDuybaX2rlOeNZd2u9i_cqqpQQS4WmBLGlyQpQ',
                                        '1iX4VRH5lpKN4MBmanJeHvgU-SYqZTIHnpcznB0GVq_Q',
                                        '12QyEDnjc2s5V548v-SFEmDYoGISqzX_L6RMsGwAxLZQ']})

service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)

df = pd.DataFrame()
df_2 = pd.DataFrame()
df_list = []
value = []

#adjust SheetID with the folder structure
for i in range(len(output_structured)):
    name = os.path.dirname(output_structured[i]).split('/')[-1]
    output.append(name)
output = pd.DataFrame({'Brand': output})
output = pd.merge(output, googlesheet, on = 'Brand')
Brand = list(output['SheetID'])

#data cleaning function
def cleaning(data):
    clean = data.T.reset_index().T.values.tolist()
    return clean

#send data function   
def send_data(brand, val):
    response = service.spreadsheets().values().append(
    spreadsheetId = brand,
    valueInputOption = 'RAW',
    insertDataOption = 'INSERT_ROWS',
        range='A1',
        body = dict(
            majorDimension='ROWS',
            values = val)
    ).execute()   

#load data
for i in range(len(output_structured)):
    df = pd.read_excel(output_structured[i])
    df_list.append(df)
    df_2 = cleaning(df_list[i])
    value.append(df_2)
    for i, a in zip(range(len(value)), Brand):
        send_data(a, value[i])