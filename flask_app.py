
# A very simple Flask Hello World app for you to get started with...

from flask import Flask, request
#from rhino3dm import *
from openpyxl import load_workbook


#wb = load_workbook('C:/Users/danae/Desktop/Stats IC Workbook.xlsx')

wb = load_workbook('/home/tesszz28/brambles/Stats IC Workbook.xlsx')


app = Flask(__name__) #name of the module currently running
wsn = wb.sheetnames


@app.route('/')
def home():
    return 'Hello from UNSW CODE2120! Please type /sheets and input paramters'


@app.route('/sheets')
def sheetname_display():
    args = request.args
    sheetnumber = args.get("sheetnumber")
    sheetname = args.get("sheetname")
    cell = args.get("cell")
    #if sheetnumber in range(len(wb.sheetnames)): #and sheetname is None:
    if None not in (sheetnumber,sheetname):
        resp = 'You can only use one of the parameters sheetnumber & sheetname at a time'
    elif sheetnumber is not None:
        resp = 'Sheet Name: ' + wsn[int(sheetnumber)]
        if cell is not None:
            ws = wb[resp]
            resp = ws[cell].value
    elif sheetname is not None:
        resp = 'Sheet Number: ' + str(wsn.index(str(sheetname)))
        if cell is not None:
            ws = wb[sheetname]
            resp = ws[cell].value
    else:
        resp = 'Please define parameters. You must pick either sheetname or sheetnumber before inputting a cell value'
    return resp

def main():
    return app.run()

if __name__ == '__main__':
    main()