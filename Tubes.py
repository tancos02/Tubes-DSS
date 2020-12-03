# Library yang digunakan
import pandas as pd
import PySimpleGUI as sg 
from sklearn.linear_model import LogisticRegression
import xlsxwriter

# Fungsi untuk memprediksi jika sumber data dari sebuah file excel(dataset)
def predictDataset(clf) :
    filename = sg.popup_get_file('Enter or Search the file')
    if(filename is not None) :
        value = filename
        df = pd.read_excel(value)
        data = df[["Jumlah Penduduk", "Jumlah Tenaga Kerja Medis", "Jumlah Wanita Hamil", "Jumlah Lansia", "Jumlah Anak-anak", "Jumlah Pasien Penyakit Kronis", "Jumlah Tentara"]]
        result = clf.predict(data)
        label_result = [0 for i in range(len(result))]
        for i in range(len(result)) :
            if (result[i] == 0) :
                label_result[i] = "Tidak divaksin"
            else :
                label_result[i] = "Divaksin"
        df['Result'] = label_result
        fname = "result.xlsx"
        workbook = xlsxwriter.Workbook(fname)
        worksheet = workbook.add_worksheet()
        workbook.close()
        writer = pd.ExcelWriter(fname)
        df.to_excel(writer, "Sheet1")
        writer.save()
        hasil = "Hasil telah disimpan di " + fname
    if(filename is not None) :
        layout = [[sg.Text(hasil)],
                  [sg.OK()]
                 ]
        window = sg.Window('Vaccine Prediction').Layout(layout)
        button, values = window.Read()
        if(button == 'OK') :
            window.close()

# Fungsi untuk memprediksi jika sumber data dari input user
def predictData(clf) :
    layout = [
        [sg.Text('Jumlah Penduduk', size =(30, 1)), sg.Input()],
        [sg.Text('Jumlah Tenaga Kerja Medis', size =(30, 1)), sg.Input()],
        [sg.Text('Jumlah Wanita Hamil', size =(30, 1)), sg.Input()],
        [sg.Text('Jumlah Lansia', size =(30, 1)), sg.Input()],
        [sg.Text('Jumlah Anak-anak', size =(30, 1)), sg.Input()],
        [sg.Text('Jumlah Pasien Penyakit Kronis', size =(30, 1)), sg.Input()],
        [sg.Text('Jumlah Tentara', size =(30, 1)), sg.Input()],
        [sg.Text('Nama Daerah', size =(30, 1)), sg.Input()],
        [sg.Submit(), sg.Cancel()]
    ]
    window = sg.Window('Vaccine Prediction').Layout(layout)
    button, values = window.Read()
    if(button == "Submit") :
        data = pd.DataFrame({"Jumlah Penduduk" : [int(values[0])],
                             "Jumlah Tenaga Kerja Medis" : [int(values[1])],
                             "Jumlah Wanita Hamil" : [int(values[2])],
                             "Jumlah Lansia" : [int(values[3])],
                             "Jumlah Anak-anak" : [int(values[4])],
                             "Jumlah Pasien Penyakit Kronis" : [int(values[5])],
                             "Jumlah Tentara" : [int(values[6])]})
        result = clf.predict(data)
        if(result[0] == 0) :
            hasil = values[7] + " tidak perlu divaksin"
        else :
            hasil = values[7] + " perlu divaksin"
    window.close()
    if(button == 'Submit') :
        layout = [[sg.Text(hasil)],
                  [sg.OK()]
                 ]
        window = sg.Window('Vaccine Prediction').Layout(layout)
        button, values = window.Read()
        if(button == 'OK') :
            window.close()

# Training machine learning menggunakan logistic regression
df = pd.read_excel("DSS Dataset.xlsx",sheet_name = 1)
data = df[["Jumlah Penduduk", "Jumlah Tenaga Kerja Medis", "Jumlah Wanita Hamil", "Jumlah Lansia", "Jumlah Anak-anak", "Jumlah Pasien Penyakit Kronis", "Jumlah Tentara"]]
target = df["Label"]
clf = LogisticRegression(random_state=0).fit(data, target)

# Main program
button = " "
layout = [[sg.Text('Selamat datang di aplikasi prediksi vaksin')],
          [sg.Text('Silahkan pilih fitur yang ingin digunakan')],
          [sg.Button('Prediksi 1 daerah'), sg.Button('Prediksi dataset'), sg.Quit()]
         ]
window = sg.Window('Vaccine Prediction').Layout(layout)
while(button != "Quit") :
    button, values = window.Read()
    if(button == 'Prediksi 1 daerah') :
        predictData(clf)
    elif(button == 'Prediksi dataset') :
        predictDataset(clf)
window.close()