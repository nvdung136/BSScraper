import requests as rq
from bs4 import BeautifulSoup as BS
import os
import sys
import pandas as pd
from openpyxl import load_workbook

 #######################################

def  Extract_Elements(Table_data):
    data_for_frame = []
    for element in Table_data:
        sub_data = []
        for sub_element in element:
            try:
                if (sub_element.get_text() != "\n"):#and(sub_element.attrs == {'class': ['b_r_c'], 'align': 'right', 'style': 'width:15%;padding:4px;'}):
                    sub_data.append(sub_element.get_text())
            except:
                continue
        data_for_frame.append(sub_data)
    return data_for_frame

def Request_n_Parse(fedURL):
    Getpage = rq.get(fedURL)
    parse = BS(Getpage.text, 'html.parser')
    return parse


def SelectInfo(): #Select the stock/period/ BS or P&L to retrieve the data (by input)
    print('Select stock(Symbol) to scrap')
    Sym = input()
    Sym = Sym.upper()
    #Can add a hash map function to find if Sym available or not
    print('Input 1 for Balance sheet \nInput 2 for Income Statement\nInput 3 for Cash Flow')
    Type = input()
    if (Type == '1'):
        SType = 'BSheet'
    elif (Type == '2'):
        SType = 'IncSta'
    elif(Type == '3'):
        SType = 'CashFlow'
    else:
        print('Error input')
    print('Start to scrap:')
    Year = input()

    ListArg = [Sym,SType,Year]
    return ListArg

def GenerateURL(ListArg): #To create the URL to feed into Request_n_Parse function
    BaseUrl = 'https://s.cafef.vn/bao-cao-tai-chinh/'
    scrapeURL =  BaseUrl + ListArg[0] + '/' + ListArg[1] + '/' + ListArg[2]+ '/4/0/0/bao-cao-tai-chinh-cong-ty-co-phan-sua-viet-nam.chn'
    return scrapeURL

def ExtractHTML(HTMLFile):
    Table = HTMLFile.find("table",{"id": "tableContent"})
    DataRows = Table.find_all("tr",{"class": ["r_item", "r_item_a"]})
    return DataRows
#def PartSwitch(): #The modify part to switch between BS and P&L

def Excel_writing(dataframe,ArgList): #Using available data to write into excel file
    if (ArgList[1]=='BSheet'):
        SheetName = 'BS_'+ ArgList[2]
    elif (ArgList[1]=='IncSta'):
        SheetName = 'Inc_' + ArgList[2]
    elif (ArgList[1]=='CashFlow'):
        SheetName = 'CF_' + ArgList[2]

    ExcelPath = os.path.join(ArgList[0]) + '.xlsx'   
    if(os.path.exists(ExcelPath)):
        book = load_workbook(ExcelPath)
        writer = pd.ExcelWriter(ExcelPath,engine='openpyxl') 
        writer.book = book
        dataframe.to_excel(writer,SheetName)
        writer.close()
    else:
        dataframe.to_excel(ArgList[0] + '.xlsx',SheetName)    
    
def ContinueScrap(ArgList,Indicator): #Update loop status whether to continue scrap or not ? if continue then which year to scrap
    print('Done scarping '+ ArgList[0] + ' ' + ArgList[1] + ' for ' + ArgList[2] + ' fiscal year' )
    print('Continue or not ? - press y for continue')
    Indicator = input()
    if (Indicator == 'y'):
        print ('Scrap for the year:')
        ArgList[2] = input()
    return ArgList, Indicator



def main():
    ArgList = SelectInfo()
    LoopIndicator = 'y'
    while (LoopIndicator == 'y'):   #create a loop to scrap multiple years 
        URL = GenerateURL(ArgList)
        Parse_HTLM = Request_n_Parse(URL)
        HTML_data = ExtractHTML(Parse_HTLM)

        data = []
        list_header = ['Items', ArgList[2]+ '_Q1',ArgList[2]+ '_Q2',ArgList[2]+ '_Q3',ArgList[2]+ '_Q4','null']
        data = Extract_Elements(HTML_data)
        dataframe = pd.DataFrame(data = data, columns= list_header)
        
        Excel_writing(dataframe,ArgList)
        ArgList, LoopIndicator = ContinueScrap(ArgList,LoopIndicator)

if __name__ == '__main__':
    main()
