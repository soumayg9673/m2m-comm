## Interface for Employee Attendance
# Column D is used for marking attendance
# On-Site  '1'
# Off-Site '0'

from tkinter import *
import openpyxl
import xlrd

emp_atd = Tk()
emp_atd.title("XYZ PRIVATE LIMITED") # Windows Title
emp_atd_l1 = Label(emp_atd, text = "Welcome to XYZ PRIVATE LIMITED", font = ("Arial", 20)) # Welcome Title
emp_atd_l1.grid(row = 0, column = 0, columnspan = 3)

emp_atd_l2 = Label(emp_atd, text = "Click to Mark Attendace", font = ('bold'))
emp_atd_l2.grid(row = 1, column = 0, columnspan=3)

loc = ("Data.xlsx") # Employee Data Reading
wb = xlrd.open_workbook(loc) 
sheetr = wb.sheet_by_index(0)
sheetr.cell_value(0, 0)

xfile = openpyxl.load_workbook(loc) # Employee Data Writing and Saving File
sheetw = xfile.get_sheet_by_name('Emp_Data')

emp_id =[]
for i in range(1, sheetr.nrows):
    emp_id.append(sheetr.cell_value(i, 0))

emp_name =[]
for i in range(1, sheetr.nrows):
    emp_name.append(sheetr.cell_value(i, 1))

emp_fo = []
temp_ons = 'green'
temp_ofs = 'red'
for i in range(1, sheetr.nrows):
    if sheetr.cell_value(i,3) == 1.0:
        emp_fo.append(temp_ons)
    else:
        emp_fo.append(temp_ofs)

def CEO_mgt():
    if emp_fo[0] == 'green':
        CEO_mgt_bt.configure(bg = 'red')
        emp_fo[0] = 'red'
        sheetw['D2'] = 0.0
        xfile.save(loc)
    else:
        CEO_mgt_bt.configure(bg = 'green')
        emp_fo[0] = 'green'
        sheetw['D2'] = 1.0
        xfile.save(loc)

def CFO_fn():
    if emp_fo[1] == 'green':
        CFO_fn_bt.configure(bg = 'red')
        emp_fo[1] = 'red'
        sheetw['D3'] = 0.0
        xfile.save(loc)
    else:
        CFO_fn_bt.configure(bg = 'green')
        emp_fo[1] = 'green'
        sheetw['D3'] = 1.0
        xfile.save(loc)

def COM_mgt():
    if emp_fo[2] == 'green':
        COM_mgt_bt.configure(bg = 'red')
        emp_fo[2] = 'red'
        sheetw['D4'] = 0.0
        xfile.save(loc)
    else:
        COM_mgt_bt.configure(bg = 'green')
        emp_fo[2] = 'green'
        sheetw['D4'] = 1.0
        xfile.save(loc)

def SMIW_spi():
    if emp_fo[3] == 'green':
        SMIW_spi_bt.configure(bg = 'red')
        emp_fo[3] = 'red'
        sheetw['D5'] = 0.0
        xfile.save(loc)
    else:
        SMIW_spi_bt.configure(bg = 'green')
        emp_fo[3] = 'green'
        sheetw['D5'] = 1.0
        xfile.save(loc)

def GMS_it():
    if emp_fo[4] == 'green':
        GMS_it_bt.configure(bg = 'red')
        emp_fo[4] = 'red'
        sheetw['D6'] = 0.0
        xfile.save(loc)
    else:
        GMS_it_bt.configure(bg = 'green')
        emp_fo[4] = 'green'
        sheetw['D6'] = 1.0
        xfile.save(loc)

def CAO_mgt():
    if emp_fo[5] == 'green':
        CAO_mgt_bt.configure(bg = 'red')
        emp_fo[5] = 'red'
        sheetw['D7'] = 0.0
        xfile.save(loc)
    else:
        CAO_mgt_bt.configure(bg = 'green')
        emp_fo[5] = 'green'
        sheetw['D7'] = 1.0
        xfile.save(loc)

def OM_sp():
    if emp_fo[6] == 'green':
        OM_sp_bt.configure(bg = 'red')
        emp_fo[6] = 'red'
        sheetw['D8'] = 0.0
        xfile.save(loc)
    else:
        OM_sp_bt.configure(bg = 'green')
        emp_fo[6] = 'green'
        sheetw['D8'] = 1.0
        xfile.save(loc)

def DM_de():
    if emp_fo[7] == 'green':
        DM_de_bt.configure(bg = 'red')
        emp_fo[7] = 'red'
        sheetw['D9'] = 0.0
        xfile.save(loc)
    else:
        DM_de_bt.configure(bg = 'green')
        emp_fo[7] = 'green'
        sheetw['D9'] = 1.0
        xfile.save(loc)
        
def QM_qi1():
    if emp_fo[8] == 'green':
        QM_qi1_bt.configure(bg = 'red')
        emp_fo[8] = 'red'
        sheetw['D10'] = 0.0
        xfile.save(loc)
    else:
        QM_qi1_bt.configure(bg = 'green')
        emp_fo[8] = 'green'
        sheetw['D10'] = 1.0
        xfile.save(loc)

def QM_qi2():
    if emp_fo[9] == 'green':
        QM_qi2_bt.configure(bg = 'red')
        emp_fo[9] = 'red'
        sheetw['D11'] = 0.0
        xfile.save(loc)
    else:
        QM_qi2_bt.configure(bg = 'green')
        emp_fo[9] = 'green'
        sheetw['D11'] = 1.0
        xfile.save(loc)

def PS_clthm():
    if emp_fo[10] == 'green':
        PS_clthm_bt.configure(bg = 'red')
        emp_fo[10] = 'red'
        sheetw['D12'] = 0.0
        xfile.save(loc)
    else:
        PS_clthm_bt.configure(bg = 'green')
        emp_fo[10] = 'green'
        sheetw['D12'] = 1.0
        xfile.save(loc)

def PW_3clm1():
    if emp_fo[11] == 'green':
        PW_3clm1_bt.configure(bg = 'red')
        emp_fo[11] = 'red'
        sheetw['D13'] = 0.0
        xfile.save(loc)
    else:
        PW_3clm1_bt.configure(bg = 'green')
        emp_fo[11] = 'green'
        sheetw['D13'] = 1.0
        xfile.save(loc)

def PW_3clm2():
    if emp_fo[12] == 'green':
        PW_3clm2_bt.configure(bg = 'red')
        emp_fo[12] = 'red'
        sheetw['D14'] = 0.0
        xfile.save(loc)
    else:
        PW_3clm2_bt.configure(bg = 'green')
        emp_fo[12] = 'green'
        sheetw['D14'] = 1.0
        xfile.save(loc)

def PW_3clm3():
    if emp_fo[13] == 'green':
        PW_3clm3_bt.configure(bg = 'red')
        emp_fo[13] = 'red'
        sheetw['D15'] = 0.0
        xfile.save(loc)
    else:
        PW_3clm3_bt.configure(bg = 'green')
        emp_fo[13] = 'green'
        sheetw['D15'] = 1.0
        xfile.save(loc)

def PW_5clm1():
    if emp_fo[14] == 'green':
        PW_5clm1_bt.configure(bg = 'red')
        emp_fo[14] = 'red'
        sheetw['D16'] = 0.0
        xfile.save(loc)
    else:
        PW_5clm1_bt.configure(bg = 'green')
        emp_fo[14] = 'green'
        sheetw['D16'] = 1.0
        xfile.save(loc)

def PW_5clm2():
    if emp_fo[15] == 'green':
        PW_5clm2_bt.configure(bg = 'red')
        emp_fo[15] = 'red'
        sheetw['D17'] = 0.0
        xfile.save(loc)
    else:
        PW_5clm2_bt.configure(bg = 'green')
        emp_fo[15] = 'green'
        sheetw['D17'] = 1.0
        xfile.save(loc)

def PS_ccm():
    if emp_fo[16] == 'green':
        PS_ccm_bt.configure(bg = 'red')
        emp_fo[16] = 'red'
        sheetw['D18'] = 0.0
        xfile.save(loc)
    else:
        PS_ccm_bt.configure(bg = 'green')
        emp_fo[16] = 'green'
        sheetw['D18'] = 1.0
        xfile.save(loc)

def PW_cpm():
    if emp_fo[17] == 'green':
        PW_cpm_bt.configure(bg = 'red')
        emp_fo[17] = 'red'
        sheetw['D19'] = 0.0
        xfile.save(loc)
    else:
        PW_cpm_bt.configure(bg = 'green')
        emp_fo[17] = 'green'
        sheetw['D19'] = 1.0
        xfile.save(loc)

def PW_clm1():
    if emp_fo[18] == 'green':
        PW_clm1_bt.configure(bg = 'red')
        emp_fo[18] = 'red'
        sheetw['D20'] = 0.0
        xfile.save(loc)
    else:
        PW_clm1_bt.configure(bg = 'green')
        emp_fo[18] = 'green'
        sheetw['D20'] = 1.0
        xfile.save(loc)

def PW_clm2():
    if emp_fo[19] == 'green':
        PW_clm2_bt.configure(bg = 'red')
        emp_fo[19] = 'red'
        sheetw['D21'] = 0.0
        xfile.save(loc)
    else:
        PW_clm2_bt.configure(bg = 'green')
        emp_fo[19] = 'green'
        sheetw['D21'] = 1.0
        xfile.save(loc)

def PS_cmm():
    if emp_fo[20] == 'green':
        PS_cmm_bt.configure(bg = 'red')
        emp_fo[20] = 'red'
        sheetw['D22'] = 0.0
        xfile.save(loc)
    else:
        PS_cmm_bt.configure(bg = 'green')
        emp_fo[20] = 'green'
        sheetw['D22'] = 1.0
        xfile.save(loc)

def PW_3cmm1():
    if emp_fo[21] == 'green':
        PW_3cmm1_bt.configure(bg = 'red')
        emp_fo[21] = 'red'
        sheetw['D23'] = 0.0
        xfile.save(loc)
    else:
        PW_3cmm1_bt.configure(bg = 'green')
        emp_fo[21] = 'green'
        sheetw['D23'] = 1.0
        xfile.save(loc)

def PW_3cmm2():
    if emp_fo[22] == 'green':
        PW_3cmm2_bt.configure(bg = 'red')
        emp_fo[22] = 'red'
        sheetw['D24'] = 0.0
        xfile.save(loc)
    else:
        PW_3cmm2_bt.configure(bg = 'green')
        emp_fo[22] = 'green'
        sheetw['D24'] = 1.0
        xfile.save(loc)

def PW_3cmm3():
    if emp_fo[23] == 'green':
        PW_3cmm3_bt.configure(bg = 'red')
        emp_fo[23] = 'red'
        sheetw['D25'] = 0.0
        xfile.save(loc)
    else:
        PW_3cmm3_bt.configure(bg = 'green')
        emp_fo[23] = 'green'
        sheetw['D25'] = 1.0
        xfile.save(loc)

def PW_5cmm1():
    if emp_fo[24] == 'green':
        PW_5cmm1_bt.configure(bg = 'red')
        emp_fo[24] = 'red'
        sheetw['D26'] = 0.0
        xfile.save(loc)
    else:
        PW_5cmm1_bt.configure(bg = 'green')
        emp_fo[24] = 'green'
        sheetw['D26'] = 1.0
        xfile.save(loc)

def PW_5cmm2():
    if emp_fo[25] == 'green':
        PW_5cmm2_bt.configure(bg = 'red')
        emp_fo[25] = 'red'
        sheetw['D27'] = 0.0
        xfile.save(loc)
    else:
        PW_5cmm2_bt.configure(bg = 'green')
        emp_fo[25] = 'green'
        sheetw['D27'] = 1.0
        xfile.save(loc)

def PS_3dp():
    if emp_fo[26] == 'green':
        PS_3dp_bt.configure(bg = 'red')
        emp_fo[26] = 'red'
        sheetw['D28'] = 0.0
        xfile.save(loc)
    else:
        PS_3dp_bt.configure(bg = 'green')
        emp_fo[26] = 'green'
        sheetw['D28'] = 1.0
        xfile.save(loc)

def PW_3dp1():
    if emp_fo[27] == 'green':
        PW_3dp1_bt.configure(bg = 'red')
        emp_fo[27] = 'red'
        sheetw['D29'] = 0.0
        xfile.save(loc)
    else:
        PW_3dp1_bt.configure(bg = 'green')
        emp_fo[27] = 'green'
        sheetw['D29'] = 1.0
        xfile
        
        xfile.save(loc)

def PW_3dp2():
    if emp_fo[28] == 'green':
        PW_3dp2_bt.configure(bg = 'red')
        emp_fo[28] = 'red'
        sheetw['D30'] = 0.0
        xfile.save(loc)
    else:
        PW_3dp2_bt.configure(bg = 'green')
        emp_fo[28] = 'green'
        sheetw['D30'] = 1.0
        xfile.save(loc)

def PW_3dp3():
    if emp_fo[29] == 'green':
        PW_3dp3_bt.configure(bg = 'red')
        emp_fo[29] = 'red'
        sheetw['D31'] = 0.0
        xfile.save(loc)
    else:
        PW_3dp3_bt.configure(bg = 'green')
        emp_fo[29] = 'green'
        sheetw['D31'] = 1.0
        xfile.save(loc)

def PW_3dp4():
    if emp_fo[30] == 'green':
        PW_3dp4_bt.configure(bg = 'red')
        emp_fo[30] = 'red'
        sheetw['D32'] = 0.0
        xfile.save(loc)
    else:
        PW_3dp4_bt.configure(bg = 'green')
        emp_fo[30] = 'green'
        sheetw['D32'] = 1.0
        xfile.save(loc)

def PS_inv():
    if emp_fo[31] == 'green':
        PS_inv_bt.configure(bg = 'red')
        emp_fo[31] = 'red'
        sheetw['D33'] = 0.0
        xfile.save(loc)
    else:
        PS_inv_bt.configure(bg = 'green')
        emp_fo[31] = 'green'
        sheetw['D33'] = 1.0
        xfile.save(loc)

def PW_inv1():
    if emp_fo[32] == 'green':
        PW_inv1_bt.configure(bg = 'red')
        emp_fo[32] = 'red'
        sheetw['D34'] = 0.0
        xfile.save(loc)
    else:
        PW_inv1_bt.configure(bg = 'green')
        emp_fo[32] = 'green'
        sheetw['D34'] = 1.0
        xfile.save(loc)

def PW_inv2():
    if emp_fo[33] == 'green':
        PW_inv2_bt.configure(bg = 'red')
        emp_fo[33] = 'red'
        sheetw['D35'] = 0.0
        xfile.save(loc)
    else:
        PW_inv2_bt.configure(bg = 'green')
        emp_fo[33] = 'green'
        sheetw['D35'] = 1.0
        xfile.save(loc)

temp_d = 0
CEO_mgt_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = CEO_mgt)
CEO_mgt_bt.grid(row = 2, column = 0)

temp_d += 1
CFO_fn_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = CFO_fn)
CFO_fn_bt.grid(row = 2, column = 1)

temp_d += 1
COM_mgt_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command =  COM_mgt)
COM_mgt_bt.grid(row = 3, column = 0)

temp_d += 1
SMIW_spi_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = SMIW_spi)
SMIW_spi_bt.grid(row = 3, column = 1)

temp_d += 1
GMS_it_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = GMS_it)
GMS_it_bt.grid(row = 4, column = 0)

temp_d += 1
CAO_mgt_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = CAO_mgt)
CAO_mgt_bt.grid(row = 4, column = 1)

temp_d += 1
OM_sp_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = OM_sp)
OM_sp_bt.grid(row = 5, column = 0)

temp_d += 1
DM_de_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = DM_de)
DM_de_bt.grid(row = 5, column = 1)

temp_d += 1
QM_qi1_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = QM_qi1)
QM_qi1_bt.grid(row = 6, column = 0)

temp_d += 1
QM_qi2_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = QM_qi2)
QM_qi2_bt.grid(row = 6, column = 1)

temp_d += 1
PS_clthm_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PS_clthm)
PS_clthm_bt.grid(row = 7, column = 0)

temp_d += 1
PW_3clm1_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PW_3clm1)
PW_3clm1_bt.grid(row = 7, column = 1)

temp_d += 1
PW_3clm2_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PW_3clm2)
PW_3clm2_bt.grid(row = 8, column = 0)

temp_d += 1
PW_3clm3_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PW_3clm3)
PW_3clm3_bt.grid(row = 8, column = 1)

temp_d += 1
PW_5clm1_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PW_5clm1)
PW_5clm1_bt.grid(row = 9, column = 0)

temp_d += 1
PW_5clm2_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PW_5clm2)
PW_5clm2_bt.grid(row = 9, column = 1)

temp_d += 1
PS_ccm_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PS_ccm)
PS_ccm_bt.grid(row = 10, column = 0)

temp_d += 1
PW_cpm_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PW_cpm)
PW_cpm_bt.grid(row = 10, column = 1)

temp_d += 1
PW_clm1_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PW_clm1)
PW_clm1_bt.grid(row = 11, column = 0)

temp_d += 1
PW_clm2_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PW_clm2)
PW_clm2_bt.grid(row = 11, column = 1)

temp_d += 1
PS_cmm_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PS_cmm)
PS_cmm_bt.grid(row = 12, column = 0)

temp_d += 1
PW_3cmm1_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PW_3cmm1)
PW_3cmm1_bt.grid(row = 12, column = 1)

temp_d += 1
PW_3cmm2_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PW_3cmm2)
PW_3cmm2_bt.grid(row = 13, column = 0)

temp_d += 1
PW_3cmm3_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PW_3cmm3)
PW_3cmm3_bt.grid(row = 13, column = 1)

temp_d += 1
PW_5cmm1_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PW_5cmm1)
PW_5cmm1_bt.grid(row = 14, column = 0)

temp_d += 1
PW_5cmm2_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PW_5cmm2)
PW_5cmm2_bt.grid(row = 14, column = 1)

temp_d += 1
PS_3dp_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PS_3dp)
PS_3dp_bt.grid(row = 15, column = 0)

temp_d += 1
PW_3dp1_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PW_3dp1)
PW_3dp1_bt.grid(row = 15, column = 1)

temp_d += 1
PW_3dp2_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PW_3dp2)
PW_3dp2_bt.grid(row = 16, column = 0)

temp_d += 1
PW_3dp3_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PW_3dp3)
PW_3dp3_bt.grid(row = 16, column = 1)

temp_d += 1
PW_3dp4_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PW_3dp4)
PW_3dp4_bt.grid(row = 17, column = 0)

temp_d += 1
PS_inv_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PS_inv)
PS_inv_bt.grid(row = 17, column = 1)

temp_d += 1
PW_inv1_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PW_inv1)
PW_inv1_bt.grid(row = 18, column = 0)    

temp_d += 1
PW_inv2_bt = Button(emp_atd, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d], command = PW_inv2)
PW_inv2_bt.grid(row = 18, column = 1) 

emp_atd.mainloop()
