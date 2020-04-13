import pandas as pd
import os
import glob
import locale


#global variable
important_columns = ['Total Pembayaran', 
                     'Waktu Pesanan Selesai', 
                     'Waktu Pesanan Dibuat', 
                     'Jumlah Produk di Pesan',
                    'Ongkos Kirim Dibayar oleh Pembeli',
                    'No. Pesanan', 'Voucher Ditanggung Shopee',
                    'Diskon Dari Penjual']

str_cols = ['Total Pembayaran',
            'Ongkos Kirim Dibayar oleh Pembeli',
            'Voucher Ditanggung Shopee',
                    'Diskon Dari Penjual']

Input = ['/Input']
Output = ['/Output']
path = os.getcwd()
w = glob.glob(path + '/Input/*/*.xls')


locale.setlocale(locale.LC_ALL, '')
locale._override_localeconv = {'mon_thousands_sep': ''}


#data cleaning function
def data_cleaning(df):
    #choose important column
    
    daily_report = df[important_columns].dropna()
    date = df['Waktu Pesanan Selesai'].str.split(" ", n = 1, expand = True).dropna()
    date.columns = ['Tanggal Pesanan Selesai', 'Jam Pesanan Selesai']
    daily_report = pd.concat([daily_report, date], axis = 1)
    daily_report = daily_report.drop(['Waktu Pesanan Selesai'], axis = 1)

    #waktu pesanan dibuat
    date = df['Waktu Pesanan Dibuat'].str.split(" ", n = 1, expand = True).dropna()
    date.columns = ['Tanggal Pesanan Dibuat', 'Jam Pesanan Dibuat']
    daily_report = pd.concat([daily_report, date], axis = 1)
    daily_report = daily_report.drop(['Waktu Pesanan Dibuat'], axis = 1)

    #remove Rp
    daily_report[str_cols] = daily_report[str_cols].replace('[Rp .]', '', regex=True)
    daily_report[str_cols] = daily_report[str_cols].replace(' ', '', regex=False)
    
    return daily_report


# Initialize empty list and dataframe
data = pd.DataFrame()
df = []
daily_report = []

#path
for i in range(len(w)):
    filename = w[i]
    data = pd.read_excel(filename)
    df.append(data)
    data = data_cleaning(df[i])
    daily_report.append(data)
    brand = os.path.dirname(w[i]).split('/')[-1]
    date = w[i][-23:-4]
    new_name = '/Users/sbahri/Documents/Hypefast/Brand Report Performance/Output/{}/{} - Output - {}.xls'.format(brand, brand, date)
    daily_report[i].to_excel(new_name, index = False)
