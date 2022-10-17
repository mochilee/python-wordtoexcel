# -*- coding: utf-8 -*-
"""
Created on Sat Oct 15 22:16:51 2022

@author: tony
"""
import docx
import pandas as pd
def word_clip(filename):
    check = False
    text = []
    text2 = []
    text3 = []
    text4 = []
    text5 = []
    text6 = []
    text7 = []
    text8 = []
    text9 = []
    text10 = []
    try:
        doc = docx.Document(filename)
        '''
        legth  = len(filename)
        print(filename[legth - 1])
        if filename[legth - 1] != "x":
           check = True
          '''
    except:
        check = True
        
    if check == False:      
        temp = ""
        for para in doc.paragraphs:
            text.append(para.text)
        for i in range(len(text)):
            #print(i[0])
            if len(text[i]) > 1:   
                if text[i][0] == "知" :
                    text2.append(text[i])
                elif text[i][0] == "難" :       
                    text3.append(text[i][4:])
                elif text[i][0] == "出" and text[i][1] == "處" :       
                    text4.append(text[i][3:])
                elif text[i][1] == "A" : 
                    text5.append(text[i][3:])
                    text9.append(temp)
                elif text[i][1] == "B" : 
                    text6.append(text[i][3:])
                elif text[i][1] == "C" : 
                    text7.append(text[i][3:])
                elif text[i][1] == "D" : 
                    text8.append(text[i][3:])
                elif text[i][0] == "題" and text[i][1] == "目" and text[i][2] == "編": 
                    text10.append(text[i][5:])
                temp = text[i]
                    
                    
                    
        
        col1 = "知識點"
        col2 = "難易度"
        col3 = "出處"
        col4 = "選項(A)"
        col5 = "選項(B)"
        col6 = "選項(C)"
        col7 = "選項(D)"
        col8 = "題目"
        col9 = "題目編號"
        
        #data = pd.DataFrame({col1:text2,col2:text3,col3:text4,col4:text5,col5:text6,col6:text7,col7:text8,col8:text9})
        data = pd.DataFrame({col1:text2,col2:text3,col3:text4,col9:text10})
        data.to_excel('sample_data.xlsx', sheet_name='sheet1', index=False)
        data = pd.DataFrame({col4:text5,col5:text6,col6:text7,col7:text8,col8:text9})
        data.to_excel('sample_data2.xlsx', sheet_name='sheet1', index=False)
        txt = "success"
        nq = len(text9)
    else:
        txt = "error"
        nq = 0
    return txt,nq