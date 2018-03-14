 # -*- coding: utf-8 -*-
import openpyxl
import datetime


"""
此函數將所有的git log, 按照 commit, git note, 時間, 順序寫入xlsx檔中.
j: 為not_flag中第幾項git log, 並且為寫入xlsx檔中第幾行(須注意覆蓋問題).
text1: 為第j項git log中的commit.
git_note: 為第j項git log中的note, 輸入時是list格式, 在此函數內會合併成string.
dat_str: 為第j項git log中的時間資料, 輸入時是list格式, 在此函數內會轉成date格式, 
目的是為了貼上xlsx時就符合excel資料欄位中的時間格式, 可以直接做排序.
 
"""
def wt_xlsx(j, text1, git_note, date_str):
    
    #打開同層資料夾下名為commitLog的xlsx檔(要貼上git log的檔案)
    workbook = openpyxl.load_workbook('git_log.xlsx')
    sheet = workbook.get_sheet_by_name('sheet1') #要貼上git_log的excel分頁
    
    #將git_log list 合併成可讀的string-text2
    text2=''
    for text in git_note:
        text2+=text
    
    #將date_str轉換成
    text3 = datetime.datetime.strptime(date_str, "%b %d %H:%M:%S %Y %z")
    
    #指定的欄位寫入, git_log資料
    sheet['A' + str(j)] = text1
    sheet['B' + str(j)] = text2
    sheet['C' + str(j)] = text3
    
    #儲存xlsl
    workbook.save('git_log.xlsx')

"""
判斷是不是git log的開頭(非merge).
判斷標準為, 第一列字串有包含commit, 第二列字串包含Authur:, 第三列包含Date:.
"""
def is_note_head(lt, i):
    if 'commit' in lt[i] and 'Author:' in lt[i + 1] and 'Date:' in lt[i + 2]:
        return True
    return False

"""
從note_flag中紀錄所有git log(非merge)的開頭位置(testList中的位置), 開始往下寫入前三列進
入note_text中. note_text為list資料型, 第一個值為commit, 第二個值為作者, 第三個為日期.
而從第四列開始往下到下一個note_flag所在的位置前一列, 是git note, 會一列一列合併進text合併完
成後加入note_text中第四個值.
最後呼叫wt_xlsx函數將note_text資料按照順序寫進xlsx檔中.
"""
def add_note(testList, note_flag):
    for j in range(len(note_flag)):
        note_text = []
        text = ''
        note_text.append(testList[note_flag[j]])
        note_text.append(testList[note_flag[j] + 1])
        note_text.append(testList[note_flag[j] + 2])

        #如果不是最後一筆git log, 請執行這一段
        if j != len(note_flag) - 1:
            #從第四列開始直到下一個note_flag都是git note
            for k in range(note_flag[j] + 4, note_flag[j + 1]):
                #如果到下一個git log前中間如果遇到 merge的 git log開頭, 迴圈停止
                if is_note_head2(testList, k):
                    break
                #排除----或‘’字串
                if testList[k] == '' or '------' in testList[k]:
                    continue
                text += testList[k]
        #如果是最後一筆git log, 請執行這一段
        else:
            #從第四列開始直到全部git log的最後一行, 都是git note
            for k in range(note_flag[j] + 4, len(testList)):
                #如果到下一個git log前中間如果遇到 merge的 git log開頭, 迴圈停止
                if is_note_head2(testList, k):
                    break
                text += testList[k]
        note_text.append(text)
        wt_xlsx(j + 2, testList[note_flag[j] + 1][8:], note_text, testList[note_flag[j] + 2][12:-1])

"""
與is_note_head函數相同, 判斷是不是git log的開頭(merge).
判斷標準為, 第一列字串有包含commit, 第二列包含Merge:, 第三列包含Authur:, 第三列包含Date:
"""
def is_note_head2(lt, i):
    if 'commit' in lt[i] and 'Merge:' in lt[i + 1] and  'Author:' in lt[i + 2] and 'Date:' in lt[i + 3]:
        return True
    return False

"""
與add_note函數相同作法, 只是add_note2會先寫入前四列進note_text, 第五列開始往下到下一個
note_flag都是git_note. 
最後呼叫wt_xlsx函數將note_text資料按照順序寫進xlsx檔中, 但是需注意merge的git log寫入的
順序必須接在非merge的後面.
"""

def add_note2(testList, note_flag):
    for j in range(len(note_flag)):
        note_text = []
        text = ''
        note_text.append(testList[note_flag[j]])
        note_text.append(testList[note_flag[j] + 1])
        note_text.append(testList[note_flag[j] + 2])
        note_text.append(testList[note_flag[j] + 3])

        if j != len(note_flag) - 1:
            for k in range(note_flag[j] + 5, note_flag[j + 1]):
                if is_note_head(testList, k):
                    break
                if testList[k] == '' or '------' in testList[k]:
                    continue
                text += testList[k]

        else:
            for k in range(note_flag[j] + 5, len(testList)):
                if is_note_head(testList, k):
                    break
                text += testList[k]

        note_text.append(text)
        #這裡要特別注意, j要加上一個數字, 這個數字為非merge 的git log總數, 此例子為136
        wt_xlsx(j + 136, testList[note_flag[j] + 2][8:], note_text, testList[note_flag[j] + 3][12:-1])


def main():
    fin=open('commit.txt')
    testList=fin.readlines()

    i = 0
    note_flag1 = []
    note_flag2 = []
    while True:
        if is_note_head(testList, i):
            note_flag1.append(i)

        if is_note_head2(testList, i):
            note_flag2.append(i)

        i += 1

        if i >= len(testList):
            break

    #先寫入非merge 的git note
    add_note(testList, note_flag1)
    #再寫入merge 的git note
    add_note2(testList, note_flag2)
    
    #檢查用
    print(len(note_flag1))
    print(len(note_flag2))


if __name__ == "__main__":
    main()

