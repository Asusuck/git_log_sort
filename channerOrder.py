#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Mar  5 22:41:51 2018

@author: hsuanminyin
"""

import openpyxl


def wt_xlsx(j,text1, text2):
    workbook = openpyxl.load_workbook('channerOrder.xlsx')
    sheet = workbook.get_sheet_by_name('sheet1')
    sheet['A' + str(j)] = text1
    sheet['B' + str(j)] = text2
    workbook.save('channerOrder.xlsx')
    

def is_note_head(lt, i):
    if 'commit' in lt[i] and 'Author:' in lt[i+1] and 'Date:' in lt[i+2]:
        return True
    return False

def main():
    testList=[]
    fin=open('release.txt')
    for line in fin:
        testList.append(line.strip())
        
    i=0
    count=0
    note_flag=[]   
    while True:
        if is_note_head(testList,i):
            note_flag.append(i)
            count+=1                    
        i+=1
        if i>=len(testList):
            break
    
    release_note={}
    for j in range(len(note_flag)):
        note_text=[]
        text=''
        note_text.append(testList[note_flag[j]])
        note_text.append(testList[note_flag[j]+1])
        note_text.append(testList[note_flag[j]+2])
        '''
        print(testList[note_flag[j]])
        print(testList[note_flag[j]+1])
        print(testList[note_flag[j]+2])
        print(testList[note_flag[j]+3])
        '''
        if j != len(note_flag)-1:
            for k in range(note_flag[j]+4,note_flag[j+1]):
                if testList[k]=='' or '------' in testList[k]:
                    continue
                text+=testList[k]
                '''
                print(testList[k], end='')
                print('\n')
                '''
        else:
            for k in range(note_flag[j]+4,len(testList)):
                text+=testList[k]
        note_text.append(text)
        release_note[str(k)+': '+testList[note_flag[j]+1][8:]]=note_text
        wt_xlsx(j+2,testList[note_flag[j]+1][8:],str(note_text))
    '''    
    print(release_note)
    print(len(release_note))
    print(len(note_flag))
    '''
    
        
if __name__=="__main__":
    main()


