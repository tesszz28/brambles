
# A very simple Flask Hello World app for you to get started with...

from flask import Flask
#from rhino3dm import *
from openpyxl import load_workbook
wb = load_workbook('C:/Users/danae/Documents/Studies and Whatever Folder/Uni Years/Year 3/T2 2022/CODE2120/Stats IC Workbook.xlsx')

app = Flask(__name__)

@app.route('/')
#def hello_world():
#    return 'Hello from Danae from UNSW CODE!'
def sheetname_display():
    return print(wb.sheetnames)

#@app.route('/page1')
#def hello_page1():
#    return 'Welcome to my first non-home page, fully coded by yours truly!'
