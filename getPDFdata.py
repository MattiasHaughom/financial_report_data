import PyPDF2
import tabula
from tabula import read_pdf
import re
import os
import pandas as pd

directory = r"C:\Users\pj20\Downloads"
os.chdir(directory)


report = max(os.listdir(directory), key = os.path.getctime)   


doc = PyPDF2.PdfFileReader(report)

pages = doc.getNumPages()

# Seach for terms in the document
incomeStmPages = []
incomeCompPages= []
opProfitPages= []
netProfitPages= []
profitPeriodPages= []
profitPages = []

income_statement = 'income statement'
comp_income = 'comprehensive income'
op_profit = 'operating profit'
net_profit = 'net profit'
profit_period = 'profit for the period'
profit = 'profit'

for i in range(pages):
    current_page = doc.getPage(i)
    text = current_page.extractText()
    if re.findall(income_statement, text.lower()):
        incomeStmPages.append(i)
    if re.findall(comp_income, text.lower()):
        incomeCompPages.append(i)
    if re.findall(op_profit, text.lower()):
        opProfitPages.append(i)
    if re.findall(net_profit, text.lower()):
        netProfitPages.append(i)
    if re.findall(profit_period, text.lower()):
        profitPeriodPages.append(i)
    if re.findall(profit, text.lower()):
        profitPages.append(i)


if len(incomeStmPages) == 1:
    result = incomeStmPages[0]
    pass
elif (len(incomeStmPages) > 1 and (len(opProfitPages) > 1 or len(opProfitPages) > 1 or len(opProfitPages) > 1) and len(profitPages>1)):
    profit = [opProfitPages,netProfitPages,profitPages]
    tmp = []
    for i in profit:
        tmp.append([x for x in incomeStmPages if x in i])
    c1set = frozenset(tmp[0])
    c2 = tmp[1:]
    result = [n for lst in c2 for n in lst if n in c1set]
    result = [i + 1 for i in result]
    # check comprehensive income instead
elif len(incomeStmPages) == 0 and len(incomeCompPages) > 0:
    tmp = []
    for i in profit:
        tmp.append([x for x in incomeStmPages if x in i])
    c1set = frozenset(tmp[0])
    c2 = tmp[1:]
    result = [n for lst in c2 for n in lst if n in c1set]





if len(result) > 1:
    tables = tabula.read_pdf(report,pages=result,multiple_tables=True,guess=True, stream=True)
    table = {}
    for number,i in enumerate(tables):
        table[number] = i
else:
    tables = tabula.read_pdf(report,pages=result[0],multiple_tables=True,guess=True, stream=True)
    table = pd.concat(tables)



