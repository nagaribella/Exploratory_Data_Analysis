# mengimpor pustaka tkinker
import tkinter
# mengimpor Pandas sebagai dasar pustaka untuk Data Analysis
import pandas as pd
# mengimpor matplotlib sebagai dasar pustaka untuk visualisasi Data
import matplotlib.pyplot as plt
# mengimpor modul rcParams untuk menggunakan fitur Auto Layout pada penyimpanan gambar
from matplotlib import rcParams
# mengimpor pustaka Numpy untuk membentuk objek N-dimensional
import numpy as np
# mengimpor modul color dan colormaps dari pustaka matplotlib untuk 
# pemilihan colormap pada visualisasi data
from matplotlib import colors
from matplotlib import cm
# mengimpor datetime sebagai dasar library untuk pengolahan data yang berbentuk tanggal
import datetime
# mengimpor library seaborn sebagai tambahan estetika pada plot
import seaborn as sns
plt.style.use('seaborn')
# mengimpor library matplotlib.ticker untuk mengganti yticks dalam bentuk persen
import matplotlib.ticker as mtick
# mengimpor os untuk memberi akses kepada windows
import os
# mengimpor tinter.ttk dan PIL untuk memasukkan gambar
import tkinter.ttk as ttk
import PIL.ImageTk
import PIL.Image

# PROGRAM UNTUK MENGETAHUI BARANG APA SAJA YANG TIDAK DIBELI TETAPI MASUK MATERIAL LIST
def barang_yang_tidak_dibeli(text, nama_path_arr):
    # memberi nama inputan yang telah dimasukkan melalui GUI
    Nama_File = str(text)
    file_name = os.path.basename(Nama_File)

    # membaca dan mendefinisikan data MR, INDENT, PO, BBM, MATERIAL_LIST yang berformat xlsx
    df_MR = pd.read_excel("{}".format(Nama_File), sheet_name="MR", usecols='A,B,C,F')
    df_INDENT= pd.read_excel("{}".format(Nama_File), sheet_name="INDENT", usecols='A,B,C,G,I')
    df_PO = pd.read_excel("{}".format(Nama_File), sheet_name="PURCHASE_ORDER", usecols='A,C,D,E,G')
    df_BBM = pd.read_excel("{}".format(Nama_File), sheet_name="BBM", usecols='D,E,F,I,A')
    df_Material_List = pd.read_excel("{}".format(Nama_File), sheet_name="MATERIAL_LIST", usecols='B,J')
    
    # membentuk series Material List untuk mengekstrak kolom Nama Barang
    Series_Material_List = df_Material_List["NAMA_BARANG"]

    # me-merge/menjoinkan dengan key 'nama barang' pada dataframe MR, INDENT, PO, dan BBM 
    # lalu didefinisikan dataframe baru yang bernama df_merge 
    df_merge = pd.merge(df_MR, df_INDENT, how='inner', on='NAMA_BARANG')
    df_merge = pd.merge(df_merge, df_PO, how='inner', on='NAMA_BARANG')
    df_merge = pd.merge(df_merge, df_BBM, how='inner', on='NAMA_BARANG')

    # menghilangkan kolom yang berisi sama/duplicate pada df_merge
    df_merge = df_merge.T.drop_duplicates().T
    # meng-update kolom Qty_Received dengan menjumlahkan row yang memiliki nama barang yang sama
    df_merge['Qty_Received'] = df_merge.groupby(['NAMA_BARANG'])['Qty_Received'].transform('sum')
    df_merge.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    # membentuk Seriesnamabarang dengan mengambil dari dataframe df_merge
    Seriesnamabarang = df_merge["NAMA_BARANG"]
    
    # membentuk List Barang
    List_Barang = []
    # menambahkan beberapa nama barang yang berada dalam df_Material_List ke dalam List Barang
    # jika barang tersebut berada dalam Seriesnamabarang
    for nama in Seriesnamabarang:
        df_Material_List[df_Material_List['NAMA_BARANG'] == nama]
        List_Barang.append(nama)
    # membentuk List_Barang_Material
    List_Barang_Material = []
    # menambahkan beberapa nama barang yang berada dalam df_Material_List ke dalam List Barang Material
    # jika barang tersebut berada dalam Series_Material_List
    for nama in Series_Material_List:
        df_Material_List[df_Material_List['NAMA_BARANG'] == nama]
        List_Barang_Material.append(nama)
    # membuang beberapa nama barang dalam list List_Barang_Material
    for barang in List_Barang:
        List_Barang_Material.remove(barang)
    # membuat dictionary Dict_barang
    Dict_Barang = {}
    # memasukkan key da value dalam Dict_Barang
    for barang in List_Barang_Material:
        Dict_Barang[barang] = int(df_Material_List[df_Material_List['NAMA_BARANG'] == barang]["TOTAL ORDER"])
    # membuat file excel
    df_barang = pd.DataFrame(Dict_Barang.items(), columns = ['Nama Barang', 'Kuantitas'])
    df_barang.to_excel(nama_path_arr[0] + 'Bahan Baku yang Diambil dari Stok.xlsx')
    # mendefinisikan keys dan values
    keys = Dict_Barang.keys()
    vals = Dict_Barang.values()
    # membuat figure
    plt.figure()
    # membuat bar plot dari keys dan vals yang sudah didefinisikan
    plot = plt.bar(keys, vals, color='grey')
    # menambah nilai data pada atas bar
    for value in plot:
        height = value.get_height()
        plt.text(value.get_x() + value.get_width()/2.,
                1.002*height,'%d' % int(height), ha='center', va='bottom')
    # menambah title dan nama axis plot
    plt.title('Grafik Bahan Baku yang Diambil dari Stok untuk Style {}'.format(os.path.splitext(file_name)[0]))
    plt.xlabel('Nama Barang')
    plt.ylabel('Kuantitas')
    # merotasi xticks
    plt.xticks(rotation = 90)
    # membuat autolayout jika gambar terpotong
    rcParams.update({'figure.autolayout': True})
    # menyimpan gambar
    plt.savefig(nama_path_arr[0] + 'Grafik Bahan Baku yang Diambil dari Stok.png', bbox_inches = "tight", dpi=1000)

#PROGRAM UNTUK MENGETAHUI BARANG APA SAJA YANG DIBELI UNTUK STOK GUDANG (TIDAK TERPAKAI)
def barang_yang_tidak_dipakai(text, nama_path_arr):
    #memberi nama inputan yang telah dimasukkan melalui GUI
    Nama_File = str(text)
    file_name = os.path.basename(Nama_File)

    #membaca dan mendefinisikan data Material_List berformat xlsx
    df_Material_List = pd.read_excel("{}".format(Nama_File), sheet_name="MATERIAL_LIST", usecols='B,J,L')
    #membentuk dataframe baru berupa barang yang memiliki kegunaan sebagai stock
    df_Barang_STOCK = df_Material_List[df_Material_List['KEGUNAAN'] == "STOCK"]
    #membentuk Series Seriesnamabarang
    SeriesNamaBarang = df_Barang_STOCK["NAMA_BARANG"]
    #membuat dictionary Dict_Stock
    Dict_Stok = {}
    #memasukkan keys dan values ke dalam Dict_Stock
    for nama in SeriesNamaBarang:
        Dict_Stok[nama] = int(df_Material_List[df_Material_List['KEGUNAAN'] == "STOCK"]["TOTAL ORDER"][df_Material_List["NAMA_BARANG"]==nama])
    # membuat file excel
    df_barang = pd.DataFrame(Dict_Stok.items(), columns = ['Nama Barang', 'Kuantitas'])
    df_barang.to_excel(nama_path_arr[0] + 'Bahan Baku yang Dibeli untuk Stok Gudang.xlsx')
    #membuat figure
    plt.figure()
    #membuat bar plot dari keys dan values Dictionary yang telah dibuat
    plot = plt.bar(Dict_Stok.keys(), Dict_Stok.values(), color='grey')
    # Add the data value on head of the bar
    for value in plot:
        height = value.get_height()
        plt.text(value.get_x() + value.get_width()/2.,
                1.002*height,'%d' % int(height), ha='center', va='bottom')

    #menambah title dan nama axis plot
    plt.title('Grafik Bahan Baku yang Dibeli untuk Stok Gudang Style {}'.format(os.path.splitext(file_name)[0]))
    plt.xlabel('Nama Barang')
    plt.ylabel('Kuantitas')
    #merotasi xticks
    plt.xticks(rotation = 90)
    #membuat autolayout jika gambar terpotong
    rcParams.update({'figure.autolayout': True})
    #menyimpan gambar
    plt.savefig(nama_path_arr[0] + 'Grafik Bahan Baku yang Dibeli untuk Stok Gudang.png', bbox_inches = "tight", dpi=1000)

#PROGRAM PERBANDINGAN FREKUENSI KEDATANGAN BARANG
def Perbandingan_Frekuensi_Kedatangan_Barang(text, nama_path_arr):
    #memberi nama inputan yang telah dimasukkan melalui GUI
    Nama_File = str(text)
    file_name = os.path.basename(Nama_File) 

    #membaca dan mendefinisikan data BBM yang berformat xlsx
    df_BBM = pd.read_excel("{}".format(Nama_File), sheet_name="BBM", usecols='D,E,F,I,A')
    #membentuk list DELIV, BARANG dan List yang berisi nama barang yang terduplikasi 
    BARANG = df_BBM["NAMA_BARANG"]
    LIST_Duplicate = []
    #menambahkan nama barang ke dalam List Duplicate apabila nama barang keluar lebih dari 1 kali dalam list BARANG
    for macam_barang in BARANG:
        if BARANG.value_counts()[macam_barang] > 1:
            LIST_Duplicate.append(macam_barang)
        else:
            continue
    ##mebuat List baru yang
    list_barang = list(set(df_BBM['NAMA_BARANG']))
    #menambah colormaps yang akan digunakan
    colormap = plt.cm.nipy_spectral
    #membreakdown color apa saja yang akan digunakan dalam grafik
    colors = [colormap(i) for i in np.linspace(0,1,len(list_barang))]
    #membentuk Dict Barang
    dict_barang = {}
    #menambah key dan value ke dalam Dict Barang
    for barang in list_barang:
        dict_barang[barang] = df_BBM[df_BBM['NAMA_BARANG'] == barang]
    # membuat figure
    plt.figure()
    #membuat scatter plot dengan x, y = DElivery_Date dan Qty_Received
    for i, nama_barang in enumerate(list_barang):
        x = dict_barang[nama_barang]['Delivery_Date']
        y = dict_barang[nama_barang]['Qty_Received']
        plt.scatter(x,y, label=nama_barang, c=colors[i])
    #menambahkan legend
    plt.legend()
    #menambahkan title dan nama axis dalam plot
    plt.title('Grafik Perbandingan Frekuensi dan Waktu Kedatangan Bahan Baku Style {}'.format(os.path.splitext(file_name)[0]))
    plt.xlabel('Tanggal')
    plt.ylabel('Kuantitas')
    #merotasi xticks
    plt.xticks(rotation = 15)
    #menyimpan gambar
    plt.savefig(nama_path_arr[0] + 'Grafik Perbandingan Frekuensi dan Waktu Kedatangan Bahan Baku.png', dpi=1000)
    
#PROGRAM HARI KEDATANGAN
def hari_kedatangan(text, nama_path_arr):
    #memberi nama inputan yang telah dimasukkan melalui GUI
    Nama_File = str(text)
    file_name = os.path.basename(Nama_File)

    #membaca dan mendefinisikan data BBM dan PO yang berformat xlsx
    df_BBM = pd.read_excel("{}".format(Nama_File), sheet_name="BBM", usecols='A,B,D')
    df_PO= pd.read_excel("{}".format(Nama_File), sheet_name="PURCHASE_ORDER", usecols='A,L,O,P')

    #me-merge/menjoinkan dengan key 'nama barang' pada dataframe PO dan BBM 
    #lalu didefinisikan dataframe baru yang bernama df_merge
    df_merge = pd.merge(df_PO, df_BBM, how='inner', on='KODE_PO')
    df_merge = df_merge.T.drop_duplicates().T
    #menghilangkan kolom yang berisi sama/duplicate pada df_merge
    df_merge.drop_duplicates(subset ="KODE_PO", 
                        keep = 'first', inplace = True)
    #menambah kolom Difference dengan mengurangkan tanggal Delivery dan Tanggal dibuatkannya PO
    df_merge['Difference'] = df_merge['Delivery_Date'].sub(df_merge['PO_Date'], axis=0)
    #menambah kolom batas_waktu_pengiriman
    df_merge['batas_waktu_pengiriman'] = df_merge.apply(lambda x: x['PO_Date'] + pd.offsets.DateOffset(days=x['CREDIT PAYMENT(DAYS)']), 1)
    #membuat list days(jumlah hari) dan difference(selisih hari)
    days = []
    difference = []
    #menambah data dalam List Days
    for day in df_merge['CREDIT PAYMENT(DAYS)']:
    #menambah data dalam list difference
        days.append(day)
    for day in df_merge['Difference']:
        difference.append(day.days)
    #membuat dataframe baru yang berisi Batas Waktu Pengiriman dan Waktu Pengiriman
    df = pd.DataFrame({'Batas Waktu Pengiriman (Hari)': days,
                    'Waktu Pengiriman (Hari)': difference}, index=df_merge["Nama_Vendor_x"])
    # membuat file excel
    df.to_excel(nama_path_arr[0] + 'Perbandingan Waktu Pengiriman dan Batas Waktu Pengiriman Bahan Baku.xlsx')
    #membuat bar plot
    df.plot.bar(rot=90)
    #menambah title dan nama axis plot
    plt.title('Grafik Perbandingan Waktu Pengiriman dan Batas Waktu Pengiriman Bahan Baku Style {}'.format(os.path.splitext(file_name)[0]))
    plt.xlabel('Nama Vendor')
    plt.ylabel('Waktu (Hari)')
    #membuat autolayout jika gambar terpotong
    rcParams.update({'figure.autolayout': True})
    #menyimpan gambar
    plt.savefig(nama_path_arr[0] + 'Grafik Perbandingan Waktu Pengiriman dan Batas Waktu Pengiriman Bahan Baku.png', bbox_inches = "tight", dpi=1000)
    
#PROGRAM PENGECEKAN INTEGRASI DATA APAKAH SESUAI ATAU TIDAK
def integrasi_data(text, nama_path_arr):
    #memberi nama inputan yang telah dimasukkan melalui GUI
    Nama_File = str(text)
    file_name = os.path.basename(Nama_File)

    #membaca dan mendefinisikan data BBM, PO, INDENT dan MR yang berformat xlsx
    df_MR = pd.read_excel("{}".format(Nama_File), sheet_name="MR", usecols='A,B,C,F')
    df_INDENT= pd.read_excel("{}".format(Nama_File), sheet_name="INDENT", usecols='A,B,C,G,I')
    df_PO = pd.read_excel("{}".format(Nama_File), sheet_name="PURCHASE_ORDER", usecols='A,C,D,E,G')
    df_BBM = pd.read_excel("{}".format(Nama_File), sheet_name="BBM", usecols='D,E,F,I,A')
    #merubah kolom quantity pada df_merge, membuang baris yang terduplikasi
    df_MR['Quantity'] = df_MR.groupby(['NAMA_BARANG'])['Quantity'].transform('sum')
    df_MR.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    #me-merge/menjoinkan dengan key 'nama barang' pada dataframe MR, INDENT, PO, dan BBM 
    #lalu didefinisikan dataframe baru yang bernama df_merge
    df_merge = pd.merge(df_MR, df_INDENT, how='outer', on='NAMA_BARANG')
    df_merge = pd.merge(df_merge, df_PO, how='outer', on='NAMA_BARANG')
    df_merge = pd.merge(df_merge, df_BBM, how='outer', on='NAMA_BARANG')
    #menghilangkan baris dengan nama barang yang sama
    df_merge.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    #menghilangkan kolom yang berisi sama/duplicate pada df_merge
    df_merge = df_merge.T.drop_duplicates().T
    df_merge['Quantity_y'] = df_merge['Quantity_y'].fillna(0)
    df_merge['Qty_Received'] = df_merge['Qty_Received'].fillna(0)
    df_merge['Diorder'] = df_merge['Diorder'].fillna(0)
    #membentuk List MR, INDENT, PO, BBM 
    List_MR = []
    List_INDENT = []
    List_PO = []
    List_BBM = []
    #membentuk Series baru MR, INDENT, PO, BBM dan activity
    MR = df_merge["Quantity_x"]
    INDENT = df_merge["Diorder"]
    PO = df_merge["Quantity_y"]
    BBM = df_merge["Qty_Received"]
    activity = df_merge["NAMA_BARANG"]
    #menambahkan value pada list yang telah dibuat sebelumnya
    for qty in MR:
        List_MR.append(qty)
    for qty in INDENT:
        List_INDENT.append(qty)
    for qty in PO:
        List_PO.append(qty)
    for qty in BBM:
        List_BBM.append(qty)
    #membuat dataframe baru
    df = pd.DataFrame({'MR': List_MR,
                    'INDENT': List_INDENT,
                    'PO' : List_PO,
                    'BBM' : List_BBM}, index=activity)
    # membuat file excel
    df.to_excel(nama_path_arr[0] + 'Perbandingan Kuantitas yang Ada di MR, BBM, PO dan INDEN.xlsx')
    # membuat plot/grafik
    df.plot.bar(rot=90)
    #menambah title dan nama axis dari plot
    plt.title('Grafik Perbandingan Kuantitas yang Ada di MR, BBM, PO dan INDEN Style {}'.format(os.path.splitext(file_name)[0]))
    plt.xlabel('Nama Barang')
    plt.ylabel('Kuantitas')
    #membuat autolayout jika gambar terpotong
    rcParams.update({'figure.autolayout': True})
    #menyimpan gambar
    plt.savefig(nama_path_arr[0] + 'Grafik Perbandingan Kuantitas yang Ada di MR, BBM, PO dan INDEN.png', bbox_inches = "tight", dpi=1000)

#PROGRAM PENGECEKAN PERBANDINGAN INDENT
def perbandingan_indent(text, nama_path_arr):
    #memberi nama inputan yang telah dimasukkan melalui GUI
    Nama_File = str(text)
    file_name = os.path.basename(Nama_File)

    #membaca dan mendefinisikan data INDENT yang berformat xlsx
    df_INDENT = pd.read_excel("{}".format(Nama_File), sheet_name="INDENT", usecols='C, E, F, G')
    #membuat List Kebutuhan, stock dan Diorder
    Kebutuhan = []
    Stock = []
    Diorder = []
    #menambah nilai ke dalam list kebutuhan, stock dan diorder dengan mengambil dari dataframe df_INDENT
    for keb in df_INDENT['Kebutuhan']:
        Kebutuhan.append(keb)
    for stk in df_INDENT['Stock']:
        Stock.append(stk)
    for order in df_INDENT['Diorder']:
        Diorder.append(order)
    #membuat dataframe baru
    df = pd.DataFrame({'Kebutuhan': Kebutuhan,
                    'Stock': Stock,
                    'Diorder': Diorder}, index=df_INDENT["NAMA_BARANG"])
    # membuat file excel
    df.to_excel(nama_path_arr[0] + 'Perbandingan Antara Kebutuhan, Stok, dan Diorder pada Inden.xlsx')
    #membuat plot bar
    df.plot.bar(rot=90)
    #menambahkan title dan axis names dalam plot
    plt.title('Grafik Perbandingan Antara Kebutuhan, Stok, dan Diorder pada Inden Style {}'.format(os.path.splitext(file_name)[0]))
    plt.xlabel('Nama Barang')
    plt.ylabel('Kuantitas')
    #membuat gambar auto layout agar tidak terpotong
    rcParams.update({'figure.autolayout': True})
    #menyimpan gambar
    plt.savefig(nama_path_arr[0] + 'Grafik Perbandingan Antara Kebutuhan, Stok, dan Diorder pada Inden.png', bbox_inches = "tight", dpi=1000)

#PROGRAM PERBANDINGAN BARANG YANG ADA PADA MATERIAL LIST, MR, BARANG YANG DIGUNAKAN DAN SISA
def perb_barang_MLMRdigunakansisa(text, nama_path_arr):
    #memberi nama inputan yang telah dimasukkan melalui GUI
    Nama_File = str(text)
    file_name = os.path.basename(Nama_File)
    
    #membaca dan mendefinisikan data BBK dan PO yang berformat xlsx
    df_BBK = pd.read_excel("{}".format(Nama_File), sheet_name="BBK", usecols='C, E, F')
    df_MR = pd.read_excel("{}".format(Nama_File), sheet_name="MR", usecols='A,B,C,F')
    df_Material_List = pd.read_excel("{}".format(Nama_File), sheet_name="MATERIAL_LIST", usecols='A,B,J')
    #mengupdate kolom baru 'Quantity' pada df_BBK dan df_MR dengan menjumlahkan nama barang yang sama
    df_BBK['Quantity'] = df_BBK.groupby(['NAMA_BARANG'])['Quantity'].transform('sum')
    df_BBK.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    df_MR['Quantity'] = df_MR.groupby(['NAMA_BARANG'])['Quantity'].transform('sum')
    df_MR.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    #me-merge/menjoinkan dengan key 'nama barang' pada dataframe PO dan BBK
    #lalu didefinisikan dataframe baru yang bernama df_merge
    df_merge = pd.merge(df_MR, df_BBK, how='outer', on='NAMA_BARANG')
    df_merge = pd.merge(df_merge, df_Material_List, how='outer', on='NAMA_BARANG')
    #menghilangkan kolom yang berisi value yang sama/terduplikasi
    df_merge = df_merge.T.drop_duplicates().T
    #mengisi nilai yang hilang pada kolom Quantity_y dan TOTAL ORDER dengan 0
    df_merge['Quantity_y'] = df_merge['Quantity_y'].fillna(0)
    df_merge['TOTAL ORDER'] = df_merge['TOTAL ORDER'].fillna(0)
    #membentuk kolom baru 'Sisa' pada dataframe df_merge
    df_merge['Sisa'] = df_merge["Quantity_x"] - df_merge["Quantity_y"]
    #membuat List Dibeli, Dipakai, dan Sisa
    MR = []
    BBK = []
    Sisa = []
    MATERIAL_LIST = []
    #menambahkan nilai ke dalam list Dibeli, Dipakai, dan Sisa dengan mengambil dari df_merge
    for mr in df_merge['Quantity_x']:
        MR.append(mr)
    for pakai in df_merge['Quantity_y']:
        BBK.append(pakai)
    for sisa in df_merge['Sisa']:
        Sisa.append(sisa)
    for ml in df_merge['TOTAL ORDER']:
        MATERIAL_LIST.append(ml)
    #membuat dataframe baru
    df = pd.DataFrame({'Material List': MATERIAL_LIST,
                    'MR': MR,
                    'BBK': BBK,
                    'Sisa' : Sisa}, index=df_merge["NAMA_BARANG"])
    # membuat file excel
    df.to_excel(nama_path_arr[0] + 'Perbandingan Bahan Baku yang ada di Daftar Material, Material Requisition, BBK, dan Sisa Penggunaan.xlsx')
    #membuat plot bar
    df.plot.bar(rot=90)
    #menambahkan title dan axis name dalam plot
    plt.title('Grafik Perbandingan Bahan Baku yang ada di Daftar Material, Material Requisition, BBK, dan Sisa Penggunaan Style {}'.format(os.path.splitext(file_name)[0]))
    plt.xlabel('Nama Barang')
    plt.ylabel('Kuantitas')
    #membuat gambar auto layout agar tidak terpotong
    rcParams.update({'figure.autolayout': True})
    #menyimpan gambar
    plt.savefig(nama_path_arr[0] + 'Grafik Perbandingan Bahan Baku yang ada di ML dan MR, Penggunaan, dan Sisa.png', bbox_inches = "tight", dpi=1000)

#PROGRAM DISTRIBUSI BAHAN BAKU
def distribusi_bahan_baku(text, nama_path_arr):
    #memberi nama inputan yang telah dimasukkan melalui GUI
    Nama_File = str(text)
    file_name = os.path.basename(Nama_File)
    
    #membaca dan mendefinisikan data MR, INDENT, PO, BBM, MATERIAL_LIST yang berformat xlsx
    df_MR = pd.read_excel("{}".format(Nama_File), sheet_name="MR", usecols='A,B,C,F')
    df_INDENT= pd.read_excel("{}".format(Nama_File), sheet_name="INDENT", usecols='A,B,C,G,I')
    df_PO = pd.read_excel("{}".format(Nama_File), sheet_name="PURCHASE_ORDER", usecols='A,C,D,E,G')
    df_BBM = pd.read_excel("{}".format(Nama_File), sheet_name="BBM", usecols='D,E,F,I,A')
    df_Material_List = pd.read_excel("{}".format(Nama_File), sheet_name="MATERIAL_LIST", usecols='B,J')
    #membentuk series Material List untuk mengekstrak kolom Nama Barang
    Series_Material_List = df_Material_List["NAMA_BARANG"]
    #me-merge/menjoinkan dengan key 'nama barang' pada dataframe MR, INDENT, PO, dan BBM 
    #lalu didefinisikan dataframe baru yang bernama df_merge
    df_merge = pd.merge(df_MR, df_INDENT, how='inner', on='NAMA_BARANG')
    df_merge = pd.merge(df_merge, df_PO, how='inner', on='NAMA_BARANG')
    df_merge = pd.merge(df_merge, df_BBM, how='inner', on='NAMA_BARANG')
    #mwnghilangkan kolom yang berisi value yang sama/terduplikasi
    df_merge = df_merge.T.drop_duplicates().T
    #meng-update kolom Qty_Received dengan menjumlahkan row yang memiliki nama barang yang sama
    df_merge['Qty_Received'] = df_merge.groupby(['NAMA_BARANG'])['Qty_Received'].transform('sum')
    df_merge.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    #membentuk Seriesnamabarang dengan mengambil dari dataframe df_merge
    Seriesnamabarang = df_merge["NAMA_BARANG"]
    #membentuk List Barang
    List_Barang = []
    #menambahkan beberapa nama barang yang berada dalam df_Material_List ke dalam List Barang
    #jika barang tersebut berada dalam Seriesnamabarang
    for nama in Seriesnamabarang:
        df_Material_List[df_Material_List['NAMA_BARANG'] == nama]
        List_Barang.append(nama)
    #membentuk List_Barang_Material
    List_Barang_Material = []
    #menambahkan beberapa nama barang yang berada dalam df_Material_List ke dalam List_Barang_Material
    #jika barang tersebut berada dalam Series_Material_List
    for nama in Series_Material_List:
        df_Material_List[df_Material_List['NAMA_BARANG'] == nama]
        List_Barang_Material.append(nama)
    #membuang beberapa nama barang dalam list List_Barang_Material
    for barang in List_Barang:
        List_Barang_Material.remove(barang)
    #membuat list baru yang berisi panjang dari List_Barang dan List_Barang_Material
    List_Check = [len(List_Barang), len(List_Barang_Material)]
    #membuat variabel labels
    labels = "Dari Pembelian", "Dari Stock Gudang"
    #membuat variabel my colors
    my_colors = ['lightsteelblue','silver']
    #membuat variable my_explode untude meng-explode tampilan pie chart
    my_explode = (0, 0.2)

    #membuat figure
    plt.figure()
    #membuat pie chart
    plt.pie(List_Check,labels=labels,autopct='%1.2f%%', startangle=15, shadow = True, colors=my_colors, explode=my_explode)
    plt.title('Distribusi Bahan Baku Style {}'.format(os.path.splitext(file_name)[0]))
    plt.axis('equal')
    #menyimpan gambar
    plt.savefig(nama_path_arr[0] + 'Distribusi Bahan Baku.png')
    #membuat autolayout jika gambar terpotong
    rcParams.update({'figure.autolayout': True})

#PROGRAM DISTRIBUSI BAHAN BAKU TERPAKAI
def distribusi_bahan_baku_terpakai(text, nama_path_arr):
    #memberi nama inputan yang telah dimasukkan melalui GUI
    Nama_File = str(text)
    file_name = os.path.basename(Nama_File)
    
    #membaca dan mendefinisikan data MATERIAL_LIST yang berformat xlsx
    df_Material_List = df_Material_List = pd.read_excel("{}".format(Nama_File), sheet_name="MATERIAL_LIST", usecols='B,J,L')
    #membuat dataframe baru mengnai barang dipakai dan barang yang digunakan sebagai stock
    df_Dipakai = df_Material_List[df_Material_List['KEGUNAAN'] == "DIPAKAI"]
    df_Barang_STOCK = df_Material_List[df_Material_List['KEGUNAAN'] == "STOCK"]
    #membuat series SeriesNamaBarangStok dan SeriesNamaBarangDipakai
    SeriesNamaBarangStok = df_Barang_STOCK["NAMA_BARANG"]
    SeriesNamaBarangDipakai = df_Dipakai["NAMA_BARANG"]
    #membuat dictionary Dict_Stock
    Dict_Stok = {}
    #menambah value dan keys kedalam Dict_Sock
    for nama in SeriesNamaBarangStok:
        Dict_Stok[nama] = int(df_Material_List[df_Material_List['KEGUNAAN'] == "STOCK"]["TOTAL ORDER"][df_Material_List["NAMA_BARANG"]==nama])
    #membuat dictionary Dict_dipakai
    Dict_dipakai = {}
    #menambah value dan keys kedalam Dict_dipakai
    for nama in SeriesNamaBarangDipakai:
        Dict_dipakai[nama] = int(df_Material_List[df_Material_List['KEGUNAAN'] == "DIPAKAI"]["TOTAL ORDER"][df_Material_List["NAMA_BARANG"]==nama])
    #membuat list baru yang berisi panjang dari keys Dict_dipakai dan keys Dict_Stok.keys
    List_Check = [len(Dict_dipakai.keys()), len(Dict_Stok.keys())]
    #membuat variabel labels
    labels = "Dipakai", "Stock/Tidak Dipakai"
    #membuat variabel my colors
    my_colors = ['lightsteelblue','silver']
    #membuat variable my_explode untude meng-explode tampilan pie chart
    my_explode = (0, 0.2)

    #membuat figure
    plt.figure()
    #membuat pie chart
    plt.pie(List_Check,labels=labels,autopct='%1.2f%%', startangle=15, shadow = True, colors=my_colors, explode=my_explode)
    plt.title('Distribusi Kegunaan dari Bahan Baku yang Dibeli Style {}'.format(os.path.splitext(file_name)[0]))
    plt.axis('equal')
    #menyimpan gambar
    plt.savefig(nama_path_arr[0] + 'Distribusi Kegunaan dari Bahan Baku yang Dibeli.png')
    #membuat autolayout jika gambar terpotong
    rcParams.update({'figure.autolayout': True})

#PROGRAM DISTRIBUSI SHIPPING MARK
def distribusi_shipping_mark(text, nama_path_arr):
    #memberi nama inputan yang telah dimasukkan melalui GUI
    Nama_File = str(text)
    file_name = os.path.basename(Nama_File)

    #membaca dan mendefinisikan data MATERIAL_LIST yang berformat xlsx
    df_PO = pd.read_excel("{}".format(Nama_File), sheet_name="PURCHASE_ORDER", usecols='Q')
    #membuat Series Shipping Mark
    Series_Shipping = df_PO["Shipping_Mark"]
    #membuat List Shipping
    List_Shipping = []
    #memberi value kedalam List Shipping
    for barang in Series_Shipping:
        List_Shipping.append(barang)
    #membuat list baru yang berisi panjang dari List_Shipping Land dan List_Shipping SEA dan AIR
    List_Check = [List_Shipping.count("LAND"), List_Shipping.count("SEA"), List_Shipping.count("AIR")]
    #menambahkan labels
    labels = "LAND", "SEA", "AIR"
    #membuat variabel my colors
    my_colors = ['lightsteelblue','silver', 'lightblue']
    #membuat variable my_explode untude meng-explode tampilan pie chart
    my_explode = (0, 0.2, 0)
    #membuat figure
    plt.figure()
    #membuat pie chart
    plt.pie(List_Check,labels=labels,autopct='%1.2f%%', startangle=15, shadow = True, 
            colors=my_colors, explode=my_explode)
    #menambah title grafik
    plt.title('Distribusi Tanda Pengiriman Bahan Baku Style {}'.format(os.path.splitext(file_name)[0]))
    plt.axis('equal')
    #menyimpan gambar
    plt.savefig(nama_path_arr[0] + 'Distribusi Tanda Pengiriman Bahan Baku.png')
    #membuat autolayout jika gambar terpotong
    rcParams.update({'figure.autolayout': True})

#PROGRAM DISTRIBUSI VENDOR
def distribusi_vendor(text, nama_path_arr):
    #memberi nama inputan yang telah dimasukkan melalui GUI
    Nama_File = str(text)
    file_name = os.path.basename(Nama_File)
    
    #membaca dan mendefinisikan data PO yang berformat xlsx
    df_PO = pd.read_excel("{}".format(Nama_File), sheet_name="PURCHASE_ORDER", usecols='P')
    #membuat Series Credit
    Series_Credit = df_PO["CREDIT PAYMENT(DAYS)"]
    #membuat List_Credit_Payment
    List_Credit_Payment = []
    #menambahkan value kedalam List_Credit_Payment
    for barang in Series_Credit:
        List_Credit_Payment.append(barang)
    #membuat list baru yang berisi panjang dari List_Credit_Payment 60 hari(lama) 
    # dan List_Credit_Payment 45 hari(baru)
    List_Check = [List_Credit_Payment.count(60), List_Credit_Payment.count(45)]
    #menambahkan labels
    labels = "Vendor Lama", "Vendor Baru"
    #membuat variabel my colors
    my_colors = ['lightsteelblue','silver']
    #membuat variable my_explode untude meng-explode tampilan pie chart
    my_explode = (0, 0.2)

    #membuat figure
    plt.figure()
    #membuat pie chart
    plt.pie(List_Check,labels=labels,autopct='%1.2f%%', startangle=15, shadow = True, 
            colors=my_colors, explode=my_explode)
    #menambahkan title grafik
    plt.title('Distribusi Vendor Style {}'.format(os.path.splitext(file_name)[0]))
    plt.axis('equal')
    #menyimpan gambar
    plt.savefig(nama_path_arr[0] + 'Distribusi Vendor.png')
    #membuat autolayout jika gambar terpotong
    rcParams.update({'figure.autolayout': True})

#PROGRAM NAMA YANG TIDAK ADA DALAM MATERIAL LIST NAMUN TERDAPAT DALAM BBK
def nama_yang_tidak_ada_di_material_list(text, nama_path_arr):
    #memberi nama inputan yang telah dimasukkan melalui GUI
    Nama_File = str(text)
    file_name = os.path.basename(Nama_File)

    #membaca dan mendefinisikan data BBK dan MATERIAL_LIST yang berformat xlsx
    df_BBK = pd.read_excel("{}".format(Nama_File), sheet_name="BBK", usecols='C, E, F')
    df_Material_List = pd.read_excel("{}".format(Nama_File), sheet_name="MATERIAL_LIST", usecols='A, B, J')
    #menambahkan kolom baru dalam dataframe BBK 'Quantity'
    df_BBK['Quantity'] = df_BBK.groupby(['NAMA_BARANG'])['Quantity'].transform('sum')
    #menghapus baris yang sama/terduplikasi
    df_BBK.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    #me-merge/menjoinkan dengan key 'nama barang' pada dataframe Material List dan BBK 
    #lalu didefinisikan dataframe baru yang bernama df_merge
    df_merge = pd.merge(df_Material_List, df_BBK, how='inner', on='NAMA_BARANG')
    #menhapus kolom yang memiliki value yang sama/terduplikasi
    df_merge = df_merge.T.drop_duplicates().T
    #membuat list nama barang dan nama barang yang ada di BBK (nama_bbk)
    Kode_Barang = []
    Kode_BBK = []
    #menambahkan value ke dalam List Kode_Barang
    for nama in df_Material_List['KODE_PRODUK']:
        Kode_Barang.append(nama)
    #menambahkan value ke dalam List Kode_BBK
    for nama in df_BBK['KODE_PRODUK']:
        Kode_BBK.append(nama)
    #membuat list baru mengenai nama yang tidak ada di material list
    Kode_yang_tidak_ada = list(set(Kode_BBK) - set(Kode_Barang))

    #membuat Dictionary Dict_Kode
    Dict_Kode = {}
    for nama in Kode_yang_tidak_ada:
        Dict_Kode[nama] = int(df_BBK[df_BBK['KODE_PRODUK'] == nama]["Quantity"])
    #mendefinsikan keys dan values dari Dict_Nama
    keys = Dict_Kode.keys()
    vals = Dict_Kode.values()
    #membuat list baru 2-D lalu memasukkan value nya dengan keys dan value dict_kode
    list_baru = []
    list_baru.append(list(keys))
    list_baru.append(list(vals))
    #mengganti axis pada list baru
    list_baru = np.swapaxes(np.array(list_baru),0,1)
    df = pd.DataFrame(list_baru, columns = ['Kode Barang', 'Kuantitas'])
    #membuat plot
    fig, ax = plt.subplots()
    #menyembunyikan axis
    fig.patch.set_visible(False)
    ax.axis('off')
    ax.axis('tight')
    #membuat tabel
    ax.table(cellText=df.values, colLabels=df.columns, loc='center')
    #membuat tabel agar tight layout
    fig.tight_layout() 
    #menambahkan judul dan axis name pada plot
    plt.title('Tabel Bahan Baku yang Tidak Terdaftar di Material List tetapi Muncul di Barang Bukti Keluar Style {}'.format(os.path.splitext(file_name)[0]))
    #menyimpan gambar
    plt.savefig(nama_path_arr[0] + 'Bahan Baku yang Tidak Terdaftar di Material List tetapi Muncul di Barang Bukti Keluar.png', bbox_inches = "tight", dpi=1000)

#PROGRAM DISTRIBUSI PERBEDAAN ALLOWANCE AKTUAL DAN ALLOWANCE SETELAH DIPERBAIKI
def allowance(text, nama_path_arr):
    #memberi nama inputan yang telah dimasukkan melalui GUI
    Nama_File = str(text)
    file_name = os.path.basename(Nama_File)

    #membaca dan mendefinisikan data BBK, PO, dan MATERIAL_LIST yang berformat xlsx
    df_BBK = pd.read_excel("{}".format(Nama_File), sheet_name="BBK", usecols='C, E')
    df_PO = pd.read_excel("{}".format(Nama_File), sheet_name="PURCHASE_ORDER", usecols='E, G')
    df_Material_List = pd.read_excel("{}".format(Nama_File), sheet_name="MATERIAL_LIST", usecols='B,C,D,E,F,J')
    ##menambahkan kolom baru dalam dataframe BBK 'Quantity'
    df_BBK['Quantity'] = df_BBK.groupby(['NAMA_BARANG'])['Quantity'].transform('sum')
    #menghapus baris yang sama/terduplikasi
    df_BBK.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    #me-merge/menjoinkan dengan key 'nama barang' pada dataframe Material List, BBK dan PO
    #lalu didefinisikan dataframe baru yang bernama df_merge
    df_merge = pd.merge(df_Material_List, df_PO, how='inner', on='NAMA_BARANG')
    df_merge = pd.merge(df_merge, df_BBK, how='inner', on='NAMA_BARANG')
    #membuang kolom yang memuat value yang sama/terduplikasi
    df_merge = df_merge.T.drop_duplicates().T
    #menambah kolom baru pada dataframe df_merge
    df_merge['FIXED_ALLOWANCE'] = (df_merge["Quantity_y"] / df_merge["QTY_ORDER"]) - df_merge["CONSP"]
    #membuang beberapa baris apabila memiliki nilai fixed alowance <= 0
    df_merge = df_merge.drop(df_merge[(df_merge.FIXED_ALLOWANCE <= 0)].index)
    #membuat list Allowance Lama dan Allowance baru
    AL = []
    AB = []
    #menambahkan nilai kedalam List AL dan AB
    for allowance in df_merge['LOSS']:
        AL.append(allowance)
    for allowance in df_merge['FIXED_ALLOWANCE']:
        AB.append(allowance)
    #membuat dataframe baru
    df = pd.DataFrame({'Allowance Aktual': AL,
                    'Allowance Seharusnya': AB}, 
                    index=df_merge['NAMA_BARANG'])
    # membuat file excel
    df.to_excel(nama_path_arr[0] + 'Perbandingan kelonggaran Dalam Pembelian Bahan Baku.xlsx')
    #membuat bar plot dengan rotasi axis 90 derajat
    df.plot.bar(rot=90)
    #menambahkan title dan nama axis ke dalam plot
    plt.title('Grafik Perbandingan kelonggaran Dalam Pembelian Bahan Baku Style {}'.format(os.path.splitext(file_name)[0]))
    plt.xlabel('Nama Barang')
    plt.ylabel('Allowance')
    #membuat autolayout jika gambar terpotong
    rcParams.update({'figure.autolayout': True})
    #menyimpan gambar
    plt.savefig(nama_path_arr[0] + 'Grafik Perbandingan kelonggaran Dalam Pembelian Bahan Baku.png', bbox_inches = "tight", dpi=1000)
    
#PROGRAM UNTUK MENAMPILKAN BARANG YANG ADA PADA MR TETAPI TIDAK TERPAKAI (TIDAK ADA DI BBK)
def adadimr_tapitidakadadibbk(text, nama_path_arr):
    #memberi nama inputan yang telah dimasukkan melalui GUI
    Nama_File = str(text)
    file_name = os.path.basename(Nama_File)

    #membaca dan mendefinisikan data MR dan BBK berformat xlsx
    df_MR = pd.read_excel("{}".format(Nama_File), sheet_name="MR", usecols='A, B, C, F')
    df_BBK = pd.read_excel("{}".format(Nama_File), sheet_name="BBK", usecols='C, E')
    #menambahkan kolom 'Quantity' pada dataframe MR dan BBK, lalu menghapus baris yang sama
    df_BBK['Quantity'] = df_BBK.groupby(['NAMA_BARANG'])['Quantity'].transform('sum')
    df_BBK.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    df_MR['Quantity'] = df_MR.groupby(['NAMA_BARANG'])['Quantity'].transform('sum')
    df_MR.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    #me-merge/menjoinkan dengan key 'nama barang' pada dataframe MR dan BBK 
    #lalu didefinisikan dataframe baru yang bernama df_merge
    df_merge = pd.merge(df_BBK, df_MR, how='inner', on='NAMA_BARANG')
    df_merge.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    #menghapus kolom yang sama
    df_merge = df_merge.T.drop_duplicates().T
    #menginisiasikan series MR dan MERGE
    MR = df_MR['NAMA_BARANG']
    MERGE = df_merge['NAMA_BARANG']
    #membuat list MR dan Merge
    List_MR = []
    List_Merge = []
    #memasukkan values kepada list MR dan merge berupa nama barang
    for nama in MR:
        List_MR.append(nama)
    for nama in MERGE:
        List_Merge.append(nama)
    #me-remove beberapa nama barang pada list MR
    List_MR = [barang for barang in List_MR if barang not in List_Merge]
    #membuat dictionary yang berisikan nama barang dan nilainya
    Dict_Nama = {}
    for nama in List_MR:
        Dict_Nama[nama] = list(df_MR[df_MR['NAMA_BARANG'] == nama]["Quantity"])
    #menginisiasikan keys dan values pada dictionary
    keys = Dict_Nama.keys()
    vals = np.array(list(Dict_Nama.values()))
    vals = np.squeeze(vals, axis=1)
    # membuat file excel
    pd.DataFrame(Dict_Nama).to_excel(nama_path_arr[0] + 'Barang yang Terdaftar di MR tapi tidak Terdaftar di BBK.xlsx')
    #membuat figure
    plt.figure()
    #membuat plot
    plot = plt.bar(keys, vals, color='grey')
    #membuat angka pada bar
    for value in plot:
        height = value.get_height()
        plt.text(value.get_x() + value.get_width()/2.,
                1.002*height,'%d' % int(height), ha='center', va='bottom')
    #menambahkan judul dan nama axis pada plot bar
    plt.title('Grafik Barang yang Terdaftar di MR tapi tidak Terdaftar di BBK Style {}'.format(os.path.splitext(file_name)[0]))
    plt.xlabel('Nama Barang')
    plt.ylabel('Kuantitas')
    #merotasi x labels 90 derajat
    plt.xticks(rotation = 90)
    #membuat autu layout pada plot bar
    rcParams.update({'figure.autolayout': True})
    #menyimpan gambar
    plt.savefig(nama_path_arr[0] + 'Grafik Barang yang Terdaftar di MR tapi tidak Terdaftar di BBK.png', bbox_inches = "tight", dpi=1000)

#PROGRAM UNTUK MENAMPILKAN KEUNTUNGAN KOTOR ATAS PEMAKAIAN BARANG YANG TERCATAT PADA BBK
def keuntungan_kotor(text, nama_path_arr):
    #memberi nama inputan yang telah dimasukkan melalui GUI
    Nama_File = str(text)
    file_name = os.path.basename(Nama_File)

    #membaca dan mendefinisikan data MR dan BBK berformat xlsx
    df_BBK = pd.read_excel("{}".format(Nama_File), sheet_name="BBK", usecols='C, E, F')
    df_PO = pd.read_excel("{}".format(Nama_File), sheet_name="PURCHASE_ORDER", usecols='D, E, G, H, I')
    df_Harga = pd.read_excel("{}".format(Nama_File), sheet_name="HARGA", usecols='E')
    #mengupdate kolom Quantity BBK dan mengapus baris dengan nama barang yang sama
    df_BBK['Quantity'] = df_BBK.groupby(['NAMA_BARANG'])['Quantity'].transform('sum')
    df_BBK.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    #me-merge/menjoinkan dengan key 'nama barang' pada dataframe PO dan BBK 
    #lalu didefinisikan dataframe baru yang bernama df_merge
    df_merge = df_merge = pd.merge(df_PO, df_BBK, how='inner', on='NAMA_BARANG')
    df_merge = df_merge.T.drop_duplicates().T
    #menghitung total pengeluaran atas pemakaian barang
    df_merge['Total_y'] = df_merge["Quantity_y"] * df_merge["Unit_Price"]
    #membuat dictionary Total
    Dict_Total = {'Peembelian Bahan Baku': df_merge["Total_y"].sum(),
                    'Penjualan': df_Harga["AMOUNT_USD"].sum()*14000,
                    'Keuntungan Kotor': df_Harga["AMOUNT_USD"].sum()*14000-df_merge["Total_y"].sum()}
    # membuat file excel
    df = pd.DataFrame(Dict_Total.items(), columns = ['Nilai', 'Jumlah (IDR)'])
    df.to_excel(nama_path_arr[0] + 'Evaluasi Pasca Produksi.xlsx')
    #menginisiasikan keys dan values pada dictionary
    keys = Dict_Total.keys()
    vals = Dict_Total.values()
    #membuat figure
    plt.figure()
    #membuat plot
    plot = plt.bar(keys, vals, color='grey')
    #membuat angka pada bar
    for value in plot:
        height = value.get_height()
        plt.text(value.get_x() + value.get_width()/2.,
                1.002*height,'{:,.2f}'.format(int(height)), ha='center', va='bottom')
    #menambahkan judul dan nama axis pada plot bar
    plt.title('Evaluasi Pasca Produksi Style {}'.format(os.path.splitext(file_name)[0]))
    plt.xlabel('Nilai')
    plt.ylabel('Jumlah (IDR)')
    #merotasi x labels 90 derajat
    plt.xticks(rotation = 90)
    #membuat autu layout pada plot bar
    rcParams.update({'figure.autolayout': True})
    #menghapus yticks
    plt.yticks([])
    #menyimpan gambar
    plt.savefig(nama_path_arr[0] + 'Evaluasi Pasca Produksi.png', bbox_inches = "tight", dpi=1000)

#PROGRAM UNTUK MENAMPILKAN PERSENTASE PENGGUNAAN BARANG TERHADAP MR
def persentase_penggunaan(text, nama_path_arr):
    #memberi nama inputan yang telah dimasukkan melalui GUI
    Nama_File = str(text)
    file_name = os.path.basename(Nama_File)

    #membaca dan mendefinisikan data MR dan BBK berformat xlsx
    df_MR = pd.read_excel("{}".format(Nama_File), sheet_name="MR", usecols='A, B, C, F')
    df_BBK = pd.read_excel("{}".format(Nama_File), sheet_name="BBK", usecols='C, E')
    #mengupdate kolom Quantity ada MR dan BBK dan mengapus baris dengan nama barang yang sama
    df_BBK['Quantity'] = df_BBK.groupby(['NAMA_BARANG'])['Quantity'].transform('sum')
    df_BBK.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    df_MR['Quantity'] = df_MR.groupby(['NAMA_BARANG'])['Quantity'].transform('sum')
    df_MR.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    #membuat data frame baru df_merge dari BBK dan MR lalu menghapus baris yang sama/terduplikasi               
    df_merge = pd.merge(df_MR, df_BBK, how='outer', on='NAMA_BARANG')
    df_merge.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    #menghapus kolom pada df_merge yang sama
    df_merge = df_merge.T.drop_duplicates().T
    #mengisi nilai yang hilang pada kolom 'Quantity_y' df_merge menjadi 0
    df_merge['Quantity_y'].fillna(0, inplace = True)
    #membuat kolom 'usage' dari df_merge
    df_merge['Usage'] = (df_merge['Quantity_y']/df_merge['Quantity_x']) * 100
    #membuat series Nama_Barang
    Nama_Barang = df_merge['NAMA_BARANG']
    #membuat list nama dan usage lalu mengisi listnya dengan data yang ada di df_merge
    List_nama = []
    for nama in Nama_Barang:
        List_nama.append(nama)
    List_usage = []
    for usage in df_merge['Usage']:
        List_usage.append(usage)
    #membuat dictionary nama yang berisi nama barang dan usage
    Dict_Nama = {'Usage (%)' : List_usage}
    #membuat dataframe baru dari Dict_Nama
    df_baru = pd.DataFrame(Dict_Nama, index = df_merge['NAMA_BARANG'])
    # membuat file excel
    df_baru.to_excel(nama_path_arr[0] + 'Persentase Penggunaan Bahan Baku terhadap Material Requisition.xlsx')
    #membuat plot bar
    ax = df_baru.plot(kind='bar')
    #mengubah yticks menjadi bentuk persen
    ax.yaxis.set_major_formatter(mtick.PercentFormatter())
    #menambah title dan nama axis pada plot bar
    plt.title('Grafik Persentase Penggunaan Bahan Baku terhadap Material Requisition Style {}'.format(os.path.splitext(file_name)[0]))
    plt.xlabel('Nama Barang')
    plt.ylabel('Persentase')
    #merotasi x labels 90 derajat
    plt.xticks(rotation = 90)
    #membuat gambar auto layout
    rcParams.update({'figure.autolayout': True})
    #menyimpan gambar
    plt.savefig(nama_path_arr[0] + 'Grafik Persentase Penggunaan Bahan Baku terhadap Material Requisition.png', bbox_inches = "tight", dpi=1000)

#PROGRAM UNTUK PERSENTASE PEMAKAIAN BAHAN BAKU Pada BARANG BUKTI MASUK Terhadap MATERIAL LIST
def persentase_penggunaan2(text, nama_path_arr):
    #memberi nama inputan yang telah dimasukkan melalui GUI
    Nama_File = str(text)
    file_name = os.path.basename(Nama_File)

    #membaca dan mendefinisikan data ML dan BBK berformat xlsx
    df_ML = pd.read_excel("{}".format(Nama_File), sheet_name="MATERIAL_LIST", usecols='A, B, J')
    df_BBK = pd.read_excel("{}".format(Nama_File), sheet_name="BBK", usecols='C, E')
    #mengupdate kolom Quantity ada ML dan BBK dan mengapus baris dengan nama barang yang sama
    df_BBK['Quantity'] = df_BBK.groupby(['NAMA_BARANG'])['Quantity'].transform('sum')
    df_BBK.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    #membuat data frame baru df_merge dari BBK dan ML lalu menghapus baris yang sama/terduplikasi               
    df_merge = pd.merge(df_ML, df_BBK, how='left', on='NAMA_BARANG')
    df_merge.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    #menghapus kolom pada df_merge yang sama
    df_merge = df_merge.T.drop_duplicates().T
    #mengisi nilai yang hilang pada kolom 'Quantity' df_merge menjadi 0
    df_merge['Quantity'].fillna(0, inplace = True)
    #membuat kolom 'usage' dari df_merge
    df_merge['Usage'] = (df_merge['Quantity']/df_merge['TOTAL ORDER']) * 100
    #membuat series Nama_Barang
    Nama_Barang = df_merge['NAMA_BARANG']
    #membuat list nama dan usage lalu mengisi listnya dengan data yang ada di df_merge
    List_nama = []
    for nama in Nama_Barang:
        List_nama.append(nama)
    List_usage = []
    for usage in df_merge['Usage']:
        List_usage.append(usage)
    #membuat dictionary nama yang berisi nama barang dan usage
    Dict_Nama = {'Usage (%)' : List_usage}
    #membuat dataframe baru dari Dict_Nama
    df_baru = pd.DataFrame(Dict_Nama, index = df_merge['NAMA_BARANG'])
    # membuat file excel
    df_baru.to_excel(nama_path_arr[0] + 'Persentase Penggunaan Bahan Baku Pada Barang Bukti Keluar Terhadap Material List.xlsx')
    #membuat plot bar
    ax = df_baru.plot(kind='bar')
    #mengubah yticks menjadi bentuk persen
    ax.yaxis.set_major_formatter(mtick.PercentFormatter())
    #menambah title dan nama axis pada plot bar
    plt.title('Grafik Persentase Penggunaan Bahan Baku Pada Barang Bukti Keluar Terhadap Material List Style {}'.format(os.path.splitext(file_name)[0]))
    plt.xlabel('Nama Barang')
    plt.ylabel('Persentase')
    #merotasi x labels 90 derajat
    plt.xticks(rotation = 90)
    #membuat gambar auto layout
    rcParams.update({'figure.autolayout': True})
    #menyimpan gambar
    plt.savefig(nama_path_arr[0] + 'Grafik Persentase Penggunaan Bahan Baku Pada BBK Terhadap Material List.png', bbox_inches = "tight", dpi=1000)

#PROGRAM UNTUK MENAMPILKAN BARANG YANG ADA DI MR NAMUN TIDAK ADA DI MATERIAL LIST
def tidak_ada_di_ml_tap_mr(text, nama_path_arr):
    #memberi nama inputan yang telah dimasukkan melalui GUI
    Nama_File = str(text)
    file_name = os.path.basename(Nama_File)

    #membaca dan mendefinisikan data MR dan Material_List berformat xlsx
    df_MR = pd.read_excel("{}".format(Nama_File), sheet_name="MR", usecols='A,B,C,F')
    df_Material_List= pd.read_excel("{}".format(Nama_File), sheet_name="MATERIAL_LIST", usecols='B')
    #membuat List MR dan ML
    list_MR = []
    list_ML = []
    #memasukkan nilai ke dalam List ML dan MR
    for nama in df_MR['NAMA_BARANG']:
        list_MR.append(nama)
    for nama in df_Material_List['NAMA_BARANG']:
        list_ML.append(nama)
    #membuang beberapa nama pada list MR
    list_MR = [nama for nama in list_MR if nama not in list_ML]
    #membuat dictionary barang
    dict_barang = {}
    #menambah keys dan values ke dalam dictionary barang
    for barang in list_MR:
        dict_barang[barang] = float(df_MR[df_MR['NAMA_BARANG'] == barang]['Quantity'].values)
    # membuat file excel
    df_barang = pd.DataFrame(dict_barang.items(), columns = ['Nama Barang', 'Kuantitas'])
    df_barang.to_excel(nama_path_arr[0] + 'Bahan Baku yang Tidak Terdaftar di Material List tetapi Terdaftar di MR.xlsx')
    #membuat figure
    plt.figure()
    #membuat plot bar
    plot = plt.bar(list(dict_barang.keys()), dict_barang.values(), color='grey')
    #menambah nilai data pada atas bar
    for value in plot:
        height = value.get_height()
        plt.text(value.get_x() + value.get_width()/2.,
                1.002*height,'%d' % int(height), ha='center', va='bottom')
    #menambahkan title dan axis name pada plot bar
    plt.title('Grafik Bahan Baku yang Tidak Terdaftar di Material List tetapi Terdaftar di MR Style {}'.format(os.path.splitext(file_name)[0]))
    plt.xlabel('Nama Barang')
    plt.ylabel('Kuantitas')
    #merotasi x labels 90 derajat
    plt.xticks(rotation = 90)
    #membuat gambar auto layout
    rcParams.update({'figure.autolayout': True})
    #menyimpan gambar
    plt.savefig(nama_path_arr[0] + 'Grafik Bahan Baku yang Tidak Terdaftar di Material List tetapi Terdaftar di MR.png', bbox_inches = "tight", dpi=1000)

#PROGRAM UNTUK MENAMPILKAN BARANG YANG TIDAK MUNCUL DI BBK TETAPI ADA DI MR
def Tidak_keluar_di_bbk_tetapi_ada_di_mr(text, nama_path_arr):
    #memberi nama inputan yang telah dimasukkan melalui GUI
    Nama_File = str(text)
    file_name = os.path.basename(Nama_File)

    #membaca dan mendefinisikan data MR dan BBK berformat xlsx
    df_MR = pd.read_excel("{}".format(Nama_File), sheet_name="MR", usecols='A, B, C, F')
    df_BBK = pd.read_excel("{}".format(Nama_File), sheet_name="BBK", usecols='C, E')
    #mengupdate kolom Quantity ada MR dan BBK dan mengapus baris dengan nama barang yang sama
    df_BBK['Quantity'] = df_BBK.groupby(['NAMA_BARANG'])['Quantity'].transform('sum')
    df_BBK.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    df_MR['Quantity'] = df_MR.groupby(['NAMA_BARANG'])['Quantity'].transform('sum')
    df_MR.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    #membuat data frame baru df_merge dari BBK dan MR lalu menghapus baris yang sama/terduplikasi                
    df_merge = pd.merge(df_MR, df_BBK, how='outer', on='NAMA_BARANG')
    df_merge.drop_duplicates(subset ="NAMA_BARANG", 
                        keep = 'first', inplace = True)
    #menghapus kolom yang terduplikasi
    df_merge = df_merge.T.drop_duplicates().T
    df_merge['Quantity_y'].fillna(0, inplace = True)
    df_merge['Sisa'] = df_merge['Quantity_x']-df_merge['Quantity_y']
    #membuat series nama barang 
    Nama_Barang = df_merge['NAMA_BARANG']
    #membuat list nama dan list sisa lalu memasukkan value nya ke dalam list
    List_nama = []
    for nama in Nama_Barang:
        List_nama.append(nama)
    List_sisa = []
    for sisa in df_merge['Sisa']:
        List_sisa.append(sisa)
    #membuat dictionary nama
    Dict_Nama = {'Tidak Diberikan Oleh Gudang' : List_sisa}
    #membuat dataframe baru dari dictionary nama
    df_baru = pd.DataFrame(Dict_Nama, index = df_merge['NAMA_BARANG'])
    # membuat file excel
    df_baru.to_excel(nama_path_arr[0] + 'Bahan Baku yang Tidak Diberikan oleh Gudang Berdasarkan Material Requisition.xlsx')
    #membuat plot bar dari dictionary nama
    df_baru.plot(kind='bar', color='grey')
    #menambah judul dan axis name pada plot
    plt.title('Grafik Bahan Baku yang Tidak Diberikan oleh Gudang Berdasarkan Material Requisition Style {}'.format(os.path.splitext(file_name)[0]))
    plt.xlabel('Nama Barang')
    plt.ylabel('Kuantitas')
    #merotasi x labels 90 derajat
    plt.xticks(rotation = 90)
    #membuat gambar auto layout
    rcParams.update({'figure.autolayout': True})
    #menyimpan gambar
    plt.savefig(nama_path_arr[0] + 'Grafik Bahan Baku yang Tidak Diberikan oleh Gudang Berdasarkan Material Requisition.png', bbox_inches = "tight", dpi=1000)

def main():
    # membuat window
    mainform = tkinter.Tk()
    mainform.grid_columnconfigure((0, 1, 2), weight=1)

    #mengubah title bar
    mainform.wm_title("Analisis Data Eksploratif")

    #mengubah wana form
    mainform["background"] = "#B2EBE0"

    #Label 1
    lbl = tkinter.Label(mainform)
    lbl['text'] = '\n\nSelamat Datang di Aplikasi Sederhana\nAnalisis Data Eksploratif\n\nPT. Sari Warna Asli Garment\n'
    lbl["background"] = "#B2EBE0"
    lbl.configure(font='Helvetica 18 bold')
    lbl.grid(row=0, column=0, columnspan=3)

    #image logo sari warna
    def importimg(file_name):
        return PIL.ImageTk.PhotoImage(PIL.Image.open(file_name))
    bg_img = importimg('logo sari warna resized.png')
    bg = ttk.Label(mainform, image=bg_img)
    bg["background"] = "#B2EBE0"
    bg.grid(column=0, row=0)

    #Label 2
    lbl2 = tkinter.Label(mainform)
    lbl2['text'] = 'Pilih File yang Ingin Dianalisis (.xlsx)'
    lbl2.config(font=('Arial', 12, 'bold'))
    lbl2["background"] = "#B2EBE0"
    lbl2.grid(row=1, column=0, columnspan=1)

    # === FILE DIALOG ===
    # membuat array nama file
    nama_file_arr = []
    nama_path_arr = ['']
    nama_file = ''

    # Fungsi untuk membuka file dialog
    def openFile():
        filename = tkinter.filedialog.askopenfilename()
        nama_file_arr.append(filename)
        pathlabel.config(text=filename)
    
    def saveLocation():
        location_path = tkinter.filedialog.askdirectory()
        location_path += '\\' if os.name == 'nt' else '/'
        nama_path_arr.insert(0, location_path)
        pathlabel2.config(text=location_path)
    
    # Tombol untuk membuka file dialog
    open_file = tkinter.Button(mainform, command=openFile)
    open_file['text'] = 'Pilih Berkas'
    open_file.grid(row=3, column=0, columnspan=1)
    
    # label 4
    pathlabel = tkinter.Label(mainform)
    pathlabel.grid(row=4, column=0, columnspan=1)
    pathlabel["background"] = "#B2EBE0"

    # Tombol save
    save_loc = tkinter.Button(mainform, command=saveLocation)
    save_loc['text'] = 'Lokasi Simpan'
    save_loc.grid(row=7, column=0, columnspan=1)

    # label 5
    pathlabel2 = tkinter.Label(mainform)
    pathlabel2.grid(row=8, column=0, columnspan=1)
    pathlabel2["background"] = "#B2EBE0"

    #Label 3
    lbl3 = tkinter.Label(mainform)
    lbl3['text'] = 'Pilih Grafik Atau Tabel yang Ingin Ditampilkan :'
    lbl3.config(font=('Arial', 12, 'bold'))
    lbl3["background"] = "#B2EBE0"
    lbl3.grid(row=1, column=1, columnspan=2, pady = 20, sticky=tkinter.W)

    # === CHECK BUTTON ===
    # Variable
    barang_yang_tidak_dibeli_flag = tkinter.IntVar()
    barang_yang_tidak_dipakai_flag = tkinter.IntVar()
    Perbandingan_Frekuensi_Kedatangan_Barang_flag = tkinter.IntVar()
    hari_kedatangan_flag = tkinter.IntVar()
    integrasi_data_flag = tkinter.IntVar()
    perbandingan_indent_flag = tkinter.IntVar()
    perb_barang_MLMRdigunakansisa_flag = tkinter.IntVar()
    distribusi_bahan_baku_flag = tkinter.IntVar()
    distribusi_bahan_baku_terpakai_flag = tkinter.IntVar()
    distribusi_shipping_mark_flag = tkinter.IntVar()
    distribusi_vendor_flag = tkinter.IntVar()
    nama_yang_tidak_ada_di_material_list_flag = tkinter.IntVar()
    allowance_flag = tkinter.IntVar()
    adadimr_tapitidakadadibbk_flag = tkinter.IntVar()
    keuntungan_kotor_flag = tkinter.IntVar()
    persentase_penggunaan_flag = tkinter.IntVar()
    tidak_ada_di_ml_tap_mr_flag = tkinter.IntVar()
    Tidak_keluar_di_bbk_tetapi_ada_di_mr_flag = tkinter.IntVar()
    persentase_penggunaan2_flag = tkinter.IntVar()

    # Check Button
    barang_yang_tidak_dibeli_check = tkinter.Checkbutton(mainform, text = "Grafik Bahan Baku yang Diambil dari Stok", variable = barang_yang_tidak_dibeli_flag, onvalue = 1, offvalue = 0)
    barang_yang_tidak_dibeli_check["background"] = "#B2EBE0"
    barang_yang_tidak_dibeli_check.grid(row=2, column=1, columnspan=1, sticky=tkinter.W)

    barang_yang_tidak_dipakai_check = tkinter.Checkbutton(mainform, text = "Grafik Bahan Baku yang Dibeli untuk Stok Gudang", variable = barang_yang_tidak_dipakai_flag, onvalue = 1, offvalue = 0)
    barang_yang_tidak_dipakai_check["background"] = "#B2EBE0"
    barang_yang_tidak_dipakai_check.grid(row=3, column=1, columnspan=1, sticky=tkinter.W)

    Perbandingan_Frekuensi_Kedatangan_Barang_check = tkinter.Checkbutton(mainform, text = "Grafik Perbandingan Frekuensi dan Waktu Kedatangan Bahan Baku", variable = Perbandingan_Frekuensi_Kedatangan_Barang_flag, onvalue = 1, offvalue = 0)
    Perbandingan_Frekuensi_Kedatangan_Barang_check["background"] = "#B2EBE0"
    Perbandingan_Frekuensi_Kedatangan_Barang_check.grid(row=4, column=1, columnspan=1, sticky=tkinter.W)

    hari_kedatangan_check = tkinter.Checkbutton(mainform, text = "Grafik Perbandingan Waktu Pengiriman dan Batas Waktu Pengiriman Bahan Baku", variable = hari_kedatangan_flag, onvalue = 1, offvalue = 0)
    hari_kedatangan_check["background"] = "#B2EBE0"
    hari_kedatangan_check.grid(row=5, column=1, columnspan=1, sticky=tkinter.W)

    integrasi_data_check = tkinter.Checkbutton(mainform, text = "Grafik Perbandingan Kuantitas yang Ada di MR, BBM, PO dan Inden", variable = integrasi_data_flag, onvalue = 1, offvalue = 0)
    integrasi_data_check["background"] = "#B2EBE0"
    integrasi_data_check.grid(row=6, column=1, columnspan=1, sticky=tkinter.W)

    perbandingan_indent_check = tkinter.Checkbutton(mainform, text = "Grafik Perbandingan Antara Kebutuhan, Stok, dan Diorder pada Inden", variable = perbandingan_indent_flag, onvalue = 1, offvalue = 0)
    perbandingan_indent_check["background"] = "#B2EBE0"
    perbandingan_indent_check.grid(row=7, column=1, columnspan=1, sticky=tkinter.W)

    perb_barang_MLMRdigunakansisa_check = tkinter.Checkbutton(mainform, text = "Grafik Perbandingan Bahan Baku yang ada di Daftar Material, MR, BBK, dan Sisa Penggunaan", variable = perb_barang_MLMRdigunakansisa_flag, onvalue = 1, offvalue = 0)
    perb_barang_MLMRdigunakansisa_check["background"] = "#B2EBE0"
    perb_barang_MLMRdigunakansisa_check.grid(row=8, column=1, columnspan=1, sticky=tkinter.W)

    distribusi_bahan_baku_check = tkinter.Checkbutton(mainform, text = "Distribusi Bahan Baku", variable = distribusi_bahan_baku_flag, onvalue = 1, offvalue = 0)
    distribusi_bahan_baku_check["background"] = "#B2EBE0"
    distribusi_bahan_baku_check.grid(row=9, column=1, columnspan=1, sticky=tkinter.W)

    distribusi_bahan_baku_terpakai_check = tkinter.Checkbutton(mainform, text = "Distribusi Kegunaan dari Bahan Baku yang Dibeli", variable = distribusi_bahan_baku_terpakai_flag, onvalue = 1, offvalue = 0)
    distribusi_bahan_baku_terpakai_check["background"] = "#B2EBE0"
    distribusi_bahan_baku_terpakai_check.grid(row=10, column=1, columnspan=1, sticky=tkinter.W)

    distribusi_shipping_mark_check = tkinter.Checkbutton(mainform, text = "Distribusi Tanda Pengiriman", variable = distribusi_shipping_mark_flag, onvalue = 1, offvalue = 0)
    distribusi_shipping_mark_check["background"] = "#B2EBE0"
    distribusi_shipping_mark_check.grid(row=2, column=2, columnspan=1, sticky=tkinter.W)

    distribusi_vendor_check = tkinter.Checkbutton(mainform, text = "Distribusi Vendor", variable = distribusi_vendor_flag, onvalue = 1, offvalue = 0)
    distribusi_vendor_check["background"] = "#B2EBE0"
    distribusi_vendor_check.grid(row=3, column=2, columnspan=1, sticky=tkinter.W)

    nama_yang_tidak_ada_di_material_list_check = tkinter.Checkbutton(mainform, text = "Tabel Bahan Baku yang Tidak Terdaftar di Material List tetapi Terdaftar di BBK", variable = nama_yang_tidak_ada_di_material_list_flag, onvalue = 1, offvalue = 0)
    nama_yang_tidak_ada_di_material_list_check["background"] = "#B2EBE0"
    nama_yang_tidak_ada_di_material_list_check.grid(row=4, column=2, columnspan=1, sticky=tkinter.W)

    allowance_check = tkinter.Checkbutton(mainform, text = "Grafik Perbandingan Kelonggaran dalam Pembelian Bahan Baku", variable = allowance_flag, onvalue = 1, offvalue = 0)
    allowance_check["background"] = "#B2EBE0"
    allowance_check.grid(row=5, column=2, columnspan=1, sticky=tkinter.W)

    adadimr_tapitidakadadibbk_check = tkinter.Checkbutton(mainform, text = "Grafik Bahan Baku yang Terdaftar di MR tetapi tidak Terdaftar di BBK", variable = adadimr_tapitidakadadibbk_flag, onvalue = 1, offvalue = 0)
    adadimr_tapitidakadadibbk_check["background"] = "#B2EBE0"
    adadimr_tapitidakadadibbk_check.grid(row=6, column=2, columnspan=1, sticky=tkinter.W)

    keuntungan_kotor_check = tkinter.Checkbutton(mainform, text = "Grafik Evaluasi Pasca Produksi", variable = keuntungan_kotor_flag, onvalue = 1, offvalue = 0)
    keuntungan_kotor_check["background"] = "#B2EBE0"
    keuntungan_kotor_check.grid(row=7, column=2, columnspan=1, sticky=tkinter.W)

    persentase_penggunaan_check = tkinter.Checkbutton(mainform, text = "Grafik Persentase Penggunaan Bahan Baku Terhadap Material Requisition", variable = persentase_penggunaan_flag, onvalue = 1, offvalue = 0)
    persentase_penggunaan_check["background"] = "#B2EBE0"
    persentase_penggunaan_check.grid(row=8, column=2, columnspan=1, sticky=tkinter.W)

    tidak_ada_di_ml_tap_mr_check = tkinter.Checkbutton(mainform, text = "Grafik Bahan Baku yang Tidak Terdaftar di Material List tetapi Terdaftar di MR", variable = tidak_ada_di_ml_tap_mr_flag, onvalue = 1, offvalue = 0)
    tidak_ada_di_ml_tap_mr_check["background"] = "#B2EBE0"
    tidak_ada_di_ml_tap_mr_check.grid(row=9, column=2, columnspan=1, sticky=tkinter.W)

    Tidak_keluar_di_bbk_tetapi_ada_di_mr_check = tkinter.Checkbutton(mainform, text = "Grafik Bahan Baku yang Tidak Diberikan oleh Gudang Berdasarkan Material Requisition", variable = Tidak_keluar_di_bbk_tetapi_ada_di_mr_flag, onvalue = 1, offvalue = 0)
    Tidak_keluar_di_bbk_tetapi_ada_di_mr_check["background"] = "#B2EBE0"
    Tidak_keluar_di_bbk_tetapi_ada_di_mr_check.grid(row=10, column=2, columnspan=1, sticky=tkinter.W)

    persentase_penggunaan2_check = tkinter.Checkbutton(mainform, text = "Grafik Persentase Penggunaan Bahan Baku Pada Barang Bukti Keluar Terhadap Material List", variable = persentase_penggunaan2_flag, onvalue = 1, offvalue = 0)
    persentase_penggunaan2_check["background"] = "#B2EBE0"
    persentase_penggunaan2_check.grid(row=11, column=1, columnspan=1, sticky=tkinter.W)

    #membuat list mengenai mark grafik dan tabel
    LIST_BUTTON =[barang_yang_tidak_dibeli_check, barang_yang_tidak_dipakai_check, Perbandingan_Frekuensi_Kedatangan_Barang_check,
                  hari_kedatangan_check, integrasi_data_check, perbandingan_indent_check, perb_barang_MLMRdigunakansisa_check,
                  Perbandingan_Frekuensi_Kedatangan_Barang_check, distribusi_bahan_baku_check, distribusi_bahan_baku_terpakai_check,
                  distribusi_vendor_check, distribusi_shipping_mark_check, nama_yang_tidak_ada_di_material_list_check,
                  allowance_check, adadimr_tapitidakadadibbk_check, keuntungan_kotor_check, persentase_penggunaan_check,
                  tidak_ada_di_ml_tap_mr_check, Tidak_keluar_di_bbk_tetapi_ada_di_mr_check, persentase_penggunaan2_check]
    #membuat fungsi select all dan deselect all
    def checkall():
        for cb in LIST_BUTTON:
            cb.select()
    def deselect_all():
        for cb in LIST_BUTTON:
            cb.deselect()
    #membuat tombol select all dan deselect all
    tkinter.Button(mainform, text="Pilih Semua", command=checkall).grid(row=12, column=1, columnspan=2, pady= 10)
    tkinter.Button(mainform, text="Hapus Semua", command=deselect_all).grid(row=14, column=1, columnspan=2)
    

    #Button (memunculkan semua plot)
    def buttonClick():
        # nama_file = txt.get()
        nama_file = nama_file_arr[0]
        nama_path = nama_path_arr
        if barang_yang_tidak_dibeli_flag.get() == 1:
            barang_yang_tidak_dibeli(nama_file,nama_path)
        if barang_yang_tidak_dipakai_flag.get() == 1:
            barang_yang_tidak_dipakai(nama_file,nama_path)
        if Perbandingan_Frekuensi_Kedatangan_Barang_flag.get() == 1:
            Perbandingan_Frekuensi_Kedatangan_Barang(nama_file,nama_path)
        if hari_kedatangan_flag.get() == 1:
            hari_kedatangan(nama_file,nama_path)
        if integrasi_data_flag.get() == 1:
            integrasi_data(nama_file,nama_path)
        if perbandingan_indent_flag.get() == 1:
            perbandingan_indent(nama_file,nama_path)
        if perb_barang_MLMRdigunakansisa_flag.get() == 1:
            perb_barang_MLMRdigunakansisa(nama_file,nama_path)
        if distribusi_bahan_baku_flag.get() == 1:
            distribusi_bahan_baku(nama_file,nama_path)
        if distribusi_bahan_baku_terpakai_flag.get() == 1:
            distribusi_bahan_baku_terpakai(nama_file,nama_path)
        if distribusi_shipping_mark_flag.get() == 1:
            distribusi_shipping_mark(nama_file,nama_path)
        if distribusi_vendor_flag.get() == 1:
            distribusi_vendor(nama_file,nama_path)
        if nama_yang_tidak_ada_di_material_list_flag.get() == 1:
            nama_yang_tidak_ada_di_material_list(nama_file,nama_path)
        if allowance_flag.get() == 1:
            allowance(nama_file,nama_path)
        if adadimr_tapitidakadadibbk_flag.get() == 1:
            adadimr_tapitidakadadibbk(nama_file,nama_path)
        if keuntungan_kotor_flag.get() == 1:
            keuntungan_kotor(nama_file,nama_path)
        if persentase_penggunaan_flag.get() == 1:
            persentase_penggunaan(nama_file,nama_path)
        if tidak_ada_di_ml_tap_mr_flag.get() == 1:
            tidak_ada_di_ml_tap_mr(nama_file,nama_path)
        if Tidak_keluar_di_bbk_tetapi_ada_di_mr_flag.get() == 1:
            Tidak_keluar_di_bbk_tetapi_ada_di_mr(nama_file,nama_path)
        if persentase_penggunaan2_flag.get() == 1:
            persentase_penggunaan2(nama_file,nama_path)
        plt.show()
    #membuat tombol SHOW
    btn = tkinter.Button(mainform, command=buttonClick, height=3, width=60)
    btn['text'] = 'Tampilkan'
    #membuat ukuran 
    btn.grid(row=22, column=1, columnspan=2, pady=50, sticky=tkinter.S)
    
    #membuat layar menjadi fullscreen
    mainform.attributes("-fullscreen", True)
    mainform.bind("<F11>", lambda event: mainform.attributes("-fullscreen",
                                        not mainform.attributes("-fullscreen")))
    mainform.bind("<Escape>", lambda event: mainform.attributes("-fullscreen", False))

    # Button for closing 
    exit_button = tkinter.Button(mainform, text="Keluar", command=mainform.destroy, width =20)
    exit_button['background'] = '#FA8072'
    exit_button.grid(row=10, column=0, columnspan=1) 
    
    
    # Show window
    mainform.mainloop()
#trigger program ini
if __name__ == "__main__":
    main()
