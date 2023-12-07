import openpyxl as vb
from openpyxl import Workbook, load_workbook
from datetime import datetime, timedelta
import xlsxwriter


# 读取数据类，将input数据读到对应的对象中
class inputData:
    def __init__(self, Product, Step_Up, Step_Up_Period, Rider_Charge, Initial_Premium, First_Withdrawal_Age, Annuity_Commencement_Date,
                 Last_Death_Age, Mortality, Withdrawal_Rate, Fixed_Allocation_Funds_Automatic_Rebalancing_Target, MandE, Fund_Fees,
                 Risk_Free_Rate, Volatility, MAW_Age1, MAW_Age2, MAW_Age3, MAW_Age4, MAW_Rate1, MAW_Rate2, MAW_Rate3, MAW_Rate4):
        self.Product = Product
        self.Step_Up = Step_Up
        self.Step_Up_Period = Step_Up_Period
        self.Rider_Charge = Rider_Charge
        self.Initial_Premium = Initial_Premium
        self.First_Withdrawal_Age = First_Withdrawal_Age
        self.Annuity_Commencement_Date = Annuity_Commencement_Date
        self.Last_Death_Age = Last_Death_Age
        self.Mortality = Mortality
        self.Withdrawal_Rate = Withdrawal_Rate
        self.Fixed_Allocation_Funds_Automatic_Rebalancing_Target = Fixed_Allocation_Funds_Automatic_Rebalancing_Target
        self.MandE = MandE
        self.Fund_Fees = Fund_Fees
        self.Risk_Free_Rate = Risk_Free_Rate
        self.Volatility = Volatility
        self.MAW_Age1 = MAW_Age1
        self.MAW_Age2 = MAW_Age2
        self.MAW_Age3 = MAW_Age3
        self.MAW_Age4 = MAW_Age4
        self.MAW_Rate1 = MAW_Rate1
        self.MAW_Rate2 = MAW_Rate2
        self.MAW_Rate3 = MAW_Rate3
        self.MAW_Rate4 = MAW_Rate4

    # 显示所有数据
    def printAllInputData(self):
        print(self.Product)
        print(self.Step_Up)
        print(self.Step_Up_Period)
        print(self.Rider_Charge)
        print(self.Initial_Premium)
        print(self.First_Withdrawal_Age)
        print(self.Annuity_Commencement_Date)
        print(self.Last_Death_Age)
        print(self.Mortality)
        print(self.Withdrawal_Rate)
        print(self.Fixed_Allocation_Funds_Automatic_Rebalancing_Target)
        print(self.MandE)
        print(self.Fund_Fees)
        print(self.Risk_Free_Rate)
        print(self.Volatility)
        print(self.MAW_Age1)
        print(self.MAW_Age2)
        print(self.MAW_Age3)
        print(self.MAW_Age4)
        print(self.MAW_Rate1)
        print(self.MAW_Rate2)
        print(self.MAW_Rate3)
        print(self.MAW_Rate4)


# 输出数据的类，将数据通过相应的对象输出到excel中
class policy:
    # A
    Year = []
    # B
    Anniversary = []
    # C
    Age = []
    # D
    Contribution = []
    # E
    AV_Pre_Fee = []
    # F
    Fund1_Pre_Fee = []
    # G
    Fund2_Pre_Fee = []
    # H
    MandE_Fund_Fees = []
    # I
    AV_Pre_Withdrawal = []
    # L
    Withdrawal_Amount = []
    # M
    AV_Post_Withdrawal = []
    # P
    Rider_Charge = []
    # Q
    AV_Post_Charges = []
    # T
    Death_Payment = []
    # U
    AV_Post_Death_Claims = []
    # V
    Fund1_Post_Death_Claims = []
    # W
    Fund2_Post_Death_Claims = []
    # X
    Fund1_Post_Rebalance = []
    # Y
    Fund2_Post_Rebalance = []
    # Z
    ROP_Death_Base = []

    # AI
    Eligible_Step_Up = []
    # AJ
    Growth_Phase = []
    # AK
    Withdrawal_Phase = []
    # AL
    Automatic_Periodic_Benefit_Status = []
    # AM
    Last_Death = []

    Fund1_Return=[]
    Fund2_Return=[]
    Rebalance_Indicator=[]
    DF=[]
    qx=[]
    Death_Claims=[]
    Withdrawal_Claims=[]

    def outputToExcel(self):
        pass


wb = vb.load_workbook('Test.xlsx')
ws = wb["Main"]

# print(list(ws.rows))
# why data1
data1 = ws['A']
# print(data1[1].value)
# print(ws['A2'].value)

# 显示表头
titles = []
for item in list(ws.rows)[0]:
# 将excel中数据，按行读取出来，并转换成列表，然后遍历列表中的第一行的每一列,添加到空列表中，当字典中的key
    titles.append(item.value)

# 读取数据，将数据保存至iData对象中
inData = inputData(ws['B2'].value, ws['B3'].value, ws['B4'].value, ws['B5'].value, ws['B7'].value, ws['B8'].value, ws['B9'].value,
                  ws['B10'].value, ws['B11'].value, ws['B12'].value, ws['B13'].value, ws['E3'].value, ws['E4'].value, ws['E6'].value,
                  ws['E7'].value, ws['G4'].value, ws['G5'].value, ws['G6'].value, ws['G7'].value, ws['H4'].value, ws['H5'].value,
                  ws['H6'].value, ws['H7'].value)
inData.printAllInputData()

# 定义对象
outData = policy()

# 给Year赋值
outData.Year = list(range(41))

# 给Anniversary赋值
start_date = datetime(2016, 8, 1)
end_date = datetime(2056, 8, 1)
current_date = start_date
while current_date <= end_date:
    outData.Anniversary.append(current_date.strftime("%Y/%m/%d"))
    current_date = current_date.replace(year=current_date.year+1)

# 给Age赋值
outData.Age = list(range(60, 101))
print(outData.Year)
print(outData.Anniversary)
print(outData.Age)

# 给D赋值
for item in range(0,41):
    outData.Contribution.append(0)



# 给AJ赋值
# =IF(AND(C19<=Age.FirstWD,C19<=Age.AnnuityComm,C19<Age.Death),1,0)
outData.Growth_Phase.append('')
for item in range(1, 41):
    if outData.Age[item] <= inData.First_Withdrawal_Age and outData.Age[item] <= inData.Annuity_Commencement_Date and outData.Age[item] <= inData.Last_Death_Age:
        outData.Growth_Phase.append(1)
    else:
        outData.Growth_Phase.append(0)

# 给AI赋值，需要已知AJ19
# =IF(AND(A19<=StepUp.Yr,AJ19=1),1,0)
outData.Eligible_Step_Up.append('')
for item in range(1, 41):
    if outData.Year[item] <= inData.Step_Up_Period and outData.Growth_Phase[item] == 1:
        outData.Eligible_Step_Up.append(1)
    else:
        outData.Eligible_Step_Up.append(0)

# 给AM赋值
outData.Last_Death.append('')
for item in range(1, 41):
    if outData.Age[item] == inData.Last_Death_Age:
        outData.Last_Death.append(1)
    else:
        outData.Last_Death.append(0)

# 给AK赋值，需要U
#Assume U are all positive for now, will modify later
# =IF(AND(OR(C19>Age.FirstWD,C19>Age.AnnuityComm),U18>0,C19<Age.Death),1,0)
outData.Withdrawal_Phase.append('')
for item in range(1,41):
    if (outData.Age[item] > inData.First_Withdrawal_Age or outData.Age[item] > inData.Annuity_Commencement_Date) and outData.Age[item]<inData.Last_Death_Age:
        outData.Withdrawal_Phase.append(1)
    else:
        outData.Withdrawal_Phase.append(0)
# Assign value to AL #Assume U are all positive for now, will modify later
#=IF(C23>=Age.Death,0,IF(AND(AK22=1,U22=0),1,AL22))
outData.Automatic_Periodic_Benefit_Status.append('')
for item in range(1,41): #item starting with 1
    if(outData.Age[item]>inData.Last_Death_Age):
        outData.Automatic_Periodic_Benefit_Status.append(0)
    #elif(outData.Withdrawal_Phase[item]==1 and outData.AV_Post_Death_Claims[item]==0):
    elif(outData.Withdrawal_Phase[item-1]==1):
        outData.Automatic_Periodic_Benefit_Status.append(1)
    else: #need to modify here
        outData.Automatic_Periodic_Benefit_Status.append(1)
        #outData.Automatic_Periodic_Benefit_Status[item]=outData.Automatic_Periodic_Benefit_Status[item-1]


outData.Fund1_Return.append('')
for item in range(1,41): #item starting with 1
    outData.Fund1_Return.append(inData.Risk_Free_Rate)


# those prints can be hidden
print(outData.Eligible_Step_Up)
print(outData.Growth_Phase)
print(outData.Withdrawal_Phase)
print(outData.Last_Death)
print(outData.Automatic_Periodic_Benefit_Status)


# Workbook() takes one, non-optional, argument
# which is the filename that we want to create.
workbook = xlsxwriter.Workbook('TestOut.xlsx')
worksheet = workbook.add_worksheet("cashFlowOutput")
workbook.close()


# 输出到excel表中
wb2 = vb.load_workbook('TestOut.xlsx')
wsOutput = wb2["cashFlowOutput"]
wsOutput.cell(1, 1).value = "Year"
wsOutput.cell(1, 2).value = "Anniversary"
wsOutput.cell(1, 3).value = "Age"

wsOutput.cell(1, 35).value = "Eligible_Step_UP"
wsOutput.cell(1, 36).value = "Growth_Phase"
wsOutput.cell(1, 37).value = "WithDrawal_Phase"
wsOutput.cell(1, 38).value = "Automatic Periodic Benefit Status"
wsOutput.cell(1, 39).value = "Last_Death"

wsOutput.cell(1, 41).value = "Fund1 Return"
wsOutput.cell(1, 42).value = "Fund2 Return"
wsOutput.cell(1, 43).value = "Rebalance indicator"
wsOutput.cell(1, 44).value = "DF"

for item in range(0, 41):
    wsOutput.cell(item + 2, 1).value = outData.Year[item]
    wsOutput.cell(item + 2, 2).value = outData.Anniversary[item]
    wsOutput.cell(item + 2, 3).value = outData.Age[item]

    # AI、AJ、AK,AM
    wsOutput.cell(item + 2, 35).value = outData.Eligible_Step_Up[item]
    wsOutput.cell(item + 2, 36).value = outData.Growth_Phase[item]
    wsOutput.cell(item + 2, 37).value = outData.Withdrawal_Phase[item]
    wsOutput.cell(item + 2, 38).value = outData.Automatic_Periodic_Benefit_Status[item]

    wsOutput.cell(item + 2, 39).value = outData.Last_Death[item]

    #AO AP AQ Fund1 Return Fund2 Return Rebalance Indicator
    wsOutput.cell(item + 2, 41).value = outData.Fund1_Return[item]
    wsOutput.cell(item + 2, 41).number_format='0.00%'


wb2.save('TestOut.xlsx')
wb2.close()


# wb.close()

