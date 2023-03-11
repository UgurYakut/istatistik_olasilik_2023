# İstatistik ve olasılık dersi ödevi 
# Ödev içeriği: Rastgele ham veri üreterek üretilen veri üzerinde frekans tablosu oluşturma.
#
# Author: Ugur Yakut
import pandas as pd 
import numpy as np 
import math
import jpype
import asposecells
jpype.startJVM()
from asposecells.api import *
excel_file=Workbook(FileFormatType.XLSX) # excel çalışma dosyası açılıyor 
data_size=1000 # veri miktarı
upper_bound=100 # üst sınır
lower_bound=25 # alt sınır 
def freq_calculate(arr):
    freq_counter=1
    temp_list=[] # listenin veri türü data,freq,data1,freq1,data2,freq2 şeklindedir
    for i in range(len(arr)):
            if(i==len(arr)-1):
                temp_list.append(arr[i])
                temp_list.append(freq_counter)
                break
            if(arr[i]==arr[i+1]):
                freq_counter+=1
            else:
                temp_list.append(arr[i])
                temp_list.append(freq_counter)
                freq_counter=1
    return temp_list
        
def create_random_array():
        return np.random.randint(lower_bound,upper_bound,data_size)

def freq_by_class(arr1,c):
    boundary=arr1[0]+c
    temp_freq=0
    all_temp_freq=0
    readable_arr=[]
    value_arr=[]
    for i in range(0,len(arr1),2):
        if(i==len(arr1)-2):
            text=str(boundary-c)+"-"+str(arr1[i])
            readable_arr.append(text)
            temp_freq+=arr1[i+1]  
            readable_arr.append(temp_freq)
            mean_bound=(arr1[i]+boundary-c)/2
            readable_arr.append(mean_bound)
            all_temp_freq+=temp_freq
            print("merhaba:",arr1[i])
            temp_freq=0
            boundary=arr1[i]+c
        elif(arr1[i]<boundary):
            temp_freq+=arr1[i+1]
        else:
            text=str(boundary-c)+"-"+str(arr1[i])
            readable_arr.append(text)
            readable_arr.append(temp_freq)
            mean_bound=(arr1[i]+boundary-c)/2
            readable_arr.append(mean_bound)
            all_temp_freq+=temp_freq
            temp_freq=0
            temp_freq+=arr1[i+1]  
            boundary=arr1[i]+c
    all_temp_freq+=temp_freq
    print("alltempfreq:",all_temp_freq)
    return readable_arr
             
     


def relative_freq_calc(arr):
    accurate=0.0
   
    temp_arr=[]
    for i in range(1,len(arr),3):
       
        new_relative_value=float(float(arr[i])/float(data_size))
        temp_arr.append(new_relative_value)
        accurate+=new_relative_value
      
    print(accurate) # saglamasını yapıyorum göreli frekansın 1.0 yerine çok küçükte olsa hata payı çıkıyor anlamadım
    return temp_arr
def calculate_number_of_classes():
     k=1+3.3*math.log(data_size,10) # sınıf sayısı fonksiyonu
     k=round(k) # k sayısı yuvarlanarak tam sayı oluyor
     return k
def calculate_class_bound():
    dg=upper_bound-lower_bound # veri genişliği -- data band
    c=float(dg)/float(k)
    c=round(c) 
    print(c)
    return c 

def table_func(arr1,arr2):
    array=[]
    x=0
    for i in range(0,len(arr1),3): # iki farklı dizi birleştirildi
           
            array.append(x+1)
            array.append(arr1[i])

            array.append(arr1[i+1])
            array.append(arr2[x])
            x+=1
            array.append(arr1[i+2])
    print(array)   
    return array
def python_to_excel(list):
    counter=1
    for i in range(0,len(list),5): # liste halinde olan dizi excel tablosuna aktarıldı
        
        excel_file.getWorksheets().get(0).getCells().get("A"+str(counter)).putValue(str(list[i]))
        excel_file.getWorksheets().get(0).getCells().get("B"+str(counter)).putValue(str(list[i+1]))
        excel_file.getWorksheets().get(0).getCells().get("C"+str(counter)).putValue(str(list[i+2]))
        excel_file.getWorksheets().get(0).getCells().get("D"+str(counter)).putValue(str(list[i+3]))
        excel_file.getWorksheets().get(0).getCells().get("E"+str(counter)).putValue(str(list[i+4]))
        counter+=1
    excel_file.save("freq_table.xlsx")
random_array= create_random_array() # ham veri üretildi
     
sorted_array=np.sort(random_array) # sıralanmış veri
print(sorted_array)
print("len sorted array :",len(sorted_array)) 
k=calculate_number_of_classes() # sınıf sayısı 
print(k)
c=calculate_class_bound() # sınıf aralığı
freq_array=freq_calculate(sorted_array) # frekans dizisi oluşturuluyor
new_arr=freq_array
print(freq_array)
print("len freq array :",len(freq_array))
classed_freq=freq_by_class(freq_array,c) # sınıflandırılmış frekans dizisi oluşturuluyor
print(classed_freq) #goreli sınıflandırılmıs
print()
relative_freq_array=relative_freq_calc(classed_freq) #goreli
print(relative_freq_array)
print()
table=table_func(classed_freq,relative_freq_array) # tablo tek bir liste haline dönüştürüldü
python_to_excel(table) #tablo excel dosyasına aktarıldı

