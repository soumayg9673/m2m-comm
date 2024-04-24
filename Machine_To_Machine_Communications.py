from tkinter import * # GUI Design
import xlrd # Reading an excel file using Python
import openpyxl # Writing/Update an excel file using Python
import threading

main_win = Tk() # Employee Login Window
main_win.geometry('1366x720') # Window Size
main_win.title("XYZ PRIVATE LIMITED") # Windows Title
main_win_l1 = Label(main_win, text = "Welcome to XYZ PRIVATE LIMITED", font = ("Arial", 20)) # Welcome Title
main_win_l1.pack()
main_win_l2 = Label(main_win, text = "LOGIN TO ACCESS PORTAL", font = ("Arial",12))
main_win_l2.pack()




def Login():

    if et_user.get() ==  "":
       if et_pass.get() == "":
            ety_up_lb = Label(main_win, text = "Please Enter Username and Password!", fg = 'red').pack()

    elif et_user.get() == "":
        ety_user_lb = Label(main_win, text = "Please Enter Username!", fg = 'red').pack()

    elif et_pass.get() == "":
        ety_pass_lb = Label(main_win, text = "Please Enter Password!", fg = 'red').pack()
    
    else:
     
        loc = ("Data.xlsx") 
        wb = xlrd.open_workbook(loc) 
        sr_empd = wb.sheet_by_index(0)
        sr_empd.cell_value(0, 0)
        sr_mchd = wb.sheet_by_index(1)
        sr_mchd.cell_value(0, 0)

        xfile = openpyxl.load_workbook(loc) # Employee Data Writing and Saving File
        wr_empd = xfile.get_sheet_by_name('Emp_Data')
        wr_mchd = xfile.get_sheet_by_name('Machine_Data')

        temp_user = 0 # To check whether Username exist in Employee Data
        for i in range(1,sr_empd.nrows): # Verifying Username, Password, Employee is On-Site

            temp_user += 1
            if et_user.get() == sr_empd.cell_value(i, 0): # Username Found

                if et_pass.get() == sr_empd.cell_value(i,2): # Password Matched with Username

                    if sr_empd.cell_value(i,3) == 1.0: # Employee On-Site or Off-Site

                        main_win.destroy()

                        def log_out_npw():
                            emp_win.destroy()
                            
                        def log_out_pw():
                            emp_mob.destroy()

                        if sr_empd.cell_value(i,4) != "Production Worker": # Production Worker will have Tablets
                            emp_win = Tk() # Employee 
                            emp_win.geometry('1366x720')
                            emp_win.title("XYZ PRIVATE LIMITED") # Windows Title
                            emp_win_n = Label(emp_win, text = "Employee Name: "+ sr_empd.cell_value(i,1))
                            emp_win_n.grid (row = 0, column = 0)
                            emp_win_ID = Label(emp_win, text = "Employee ID: "+ sr_empd.cell_value(i,0))
                            emp_win_ID.grid (row = 1, column = 0)
                            emp_win_D = Label(emp_win, text = sr_empd.cell_value(i,4))
                            emp_win_D.grid(row = 2, column = 0)
                            emp_win_dpt = Label(emp_win, text = "Department: " + sr_empd.cell_value(i,5))
                            emp_win_dpt.grid(row = 3, column = 0)
                            emp_win_lgt = Button(emp_win, text = "Logut", width = 35, command = log_out_npw)
                            emp_win_lgt.grid(row = 4, column = 0)
                            emp_win_l1 = Label(emp_win, text = "  ")
                            emp_win_l1.grid(row = 5, column = 0)
                            emp_win_l2 = Label(emp_win, text = " ")
                            emp_win_l2.grid (row = 7, column = 0)

                        else:
                            emp_mob = Tk() # Employee 
                            emp_mob.geometry('720x720')
                            emp_mob.title("XYZ PRIVATE LIMITED") # Windows Title
                            emp_mob_n = Label(emp_mob, text = "Employee Name: "+ sr_empd.cell_value(i,1), width = 30)
                            emp_mob_n.grid (row = 0, column = 0)
                            emp_mob_ID = Label(emp_mob, text = "Employee ID: "+ sr_empd.cell_value(i,0))
                            emp_mob_ID.grid (row = 1, column = 0)
                            emp_mob_D = Label(emp_mob, text = sr_empd.cell_value(i,4))
                            emp_mob_D.grid(row = 2, column = 0)
                            emp_mob_dpt = Label (emp_mob, text = "Department: " + sr_empd.cell_value(i,5))
                            emp_mob_dpt.grid(row = 3, column = 0)
                            emp_mob_lgt = Button(emp_mob, text = "Logut", width = 35, command = log_out_pw)
                            emp_mob_lgt.grid(row = 5, column = 0)
                            emp_mob_l1 = Label(emp_mob, text = "  ")
                            emp_mob_l1.grid(row = 6, column = 0)
                            

                        emp_id =[] # Reading Employee ID
                        for temp_e in range(1, sr_empd.nrows):
                            emp_id.append(sr_empd.cell_value(temp_e, 0))

                        emp_name =[] #Reading Employee Name
                        for temp_e in range(1, sr_empd.nrows):
                            emp_name.append(sr_empd.cell_value(temp_e, 1))

                        emp_fo = [] # Reading Employee Attendance
                        def emp_attd_udt():
                            loc = ("D:\Book1.xlsx") 
                            wb_emp_attd_udt = xlrd.open_workbook(loc) 
                            sr_emp_attd_udt = wb_emp_attd_udt.sheet_by_index(0)
                            sr_emp_attd_udt.cell_value(0, 0)
                            temp_ons = 'green'
                            temp_ofs = 'red'
                            emp_fo.clear()
                            for temp_e in range(1, sr_emp_attd_udt.nrows):
                                if sr_emp_attd_udt.cell_value(temp_e,3) == 1.0:
                                    emp_fo.append(temp_ons)
                                else:
                                    emp_fo.append(temp_ofs)
                        emp_attd_udt()

                        def employee_attendance():
                            emp_attd_udt()
    
                            def employee_attendance_hide(): # Hide Employee Attendance from CEO Login
                                emp_win_hr_bt.configure(command = employee_attendance)
                                CEO_mgt_bt.destroy()
                                CFO_fn_bt.destroy()
                                COM_mgt_bt.destroy()
                                SMIW_spi_bt.destroy()
                                GMS_it_bt.destroy()
                                CAO_mgt_bt.destroy()
                                OM_sp_bt.destroy()
                                DM_de_bt.destroy()
                                QM_qi1_bt.destroy()
                                QM_qi2_bt.destroy()
                                PS_clthm_bt.destroy()
                                PW_3clm1_bt.destroy()
                                PW_3clm2_bt.destroy()
                                PW_3clm3_bt.destroy()
                                PW_5clm1_bt.destroy()
                                PW_5clm2_bt.destroy()
                                PS_ccm_bt.destroy()
                                PW_cpm_bt.destroy()
                                PW_clm1_bt.destroy()
                                PW_clm2_bt.destroy()
                                PS_cmm_bt.destroy()
                                PW_3cmm1_bt.destroy()
                                PW_3cmm2_bt.destroy()
                                PW_3cmm3_bt.destroy()
                                PW_5cmm1_bt.destroy()
                                PW_5cmm2_bt.destroy()
                                PS_3dp_bt.destroy()
                                PW_3dp1_bt.destroy()
                                PW_3dp2_bt.destroy()
                                PW_3dp3_bt.destroy()
                                PW_3dp4_bt.destroy()
                                PS_inv_bt.destroy()
                                PW_inv1_bt.destroy()
                                PW_inv2_bt.destroy()

                            emp_win_hr_bt.configure(command = employee_attendance_hide)

                            temp_d = 0
                            CEO_mgt_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            CEO_mgt_bt.grid(row = 8, column = 0)

                            temp_d += 1
                            CFO_fn_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            CFO_fn_bt.grid(row = 8, column = 1)

                            temp_d += 1
                            COM_mgt_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            COM_mgt_bt.grid(row = 8, column = 2)

                            temp_d += 1
                            SMIW_spi_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            SMIW_spi_bt.grid(row = 8, column = 3)

                            temp_d += 1
                            GMS_it_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            GMS_it_bt.grid(row = 8, column = 4)

                            temp_d += 1
                            CAO_mgt_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            CAO_mgt_bt.grid(row = 9, column = 0)

                            temp_d += 1
                            OM_sp_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            OM_sp_bt.grid(row = 9, column = 1)

                            temp_d += 1
                            DM_de_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            DM_de_bt.grid(row = 9, column = 2)

                            temp_d += 1
                            QM_qi1_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            QM_qi1_bt.grid(row = 9, column = 3)

                            temp_d += 1
                            QM_qi2_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            QM_qi2_bt.grid(row = 9, column = 4)

                            temp_d += 1
                            PS_clthm_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PS_clthm_bt.grid(row = 10, column = 0)

                            temp_d += 1
                            PW_3clm1_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PW_3clm1_bt.grid(row = 10, column = 1)

                            temp_d += 1
                            PW_3clm2_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PW_3clm2_bt.grid(row = 10, column = 2)

                            temp_d += 1
                            PW_3clm3_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PW_3clm3_bt.grid(row = 10, column = 3)

                            temp_d += 1
                            PW_5clm1_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PW_5clm1_bt.grid(row = 10, column = 4)

                            temp_d += 1
                            PW_5clm2_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PW_5clm2_bt.grid(row = 11, column = 0)

                            temp_d += 1
                            PS_ccm_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PS_ccm_bt.grid(row = 11, column = 1)

                            temp_d += 1
                            PW_cpm_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PW_cpm_bt.grid(row = 11, column = 2)

                            temp_d += 1
                            PW_clm1_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PW_clm1_bt.grid(row = 11, column = 3)

                            temp_d += 1
                            PW_clm2_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PW_clm2_bt.grid(row = 11, column = 4)

                            temp_d += 1
                            PS_cmm_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PS_cmm_bt.grid(row = 12, column = 0)

                            temp_d += 1
                            PW_3cmm1_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PW_3cmm1_bt.grid(row = 12, column = 1)

                            temp_d += 1
                            PW_3cmm2_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PW_3cmm2_bt.grid(row = 12, column = 2)

                            temp_d += 1
                            PW_3cmm3_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PW_3cmm3_bt.grid(row = 12, column = 3)

                            temp_d += 1
                            PW_5cmm1_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PW_5cmm1_bt.grid(row = 12, column = 4)

                            temp_d += 1
                            PW_5cmm2_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PW_5cmm2_bt.grid(row = 13, column = 0)

                            temp_d += 1
                            PS_3dp_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PS_3dp_bt.grid(row = 13, column = 1)

                            temp_d += 1
                            PW_3dp1_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PW_3dp1_bt.grid(row = 13, column = 2)

                            temp_d += 1
                            PW_3dp2_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PW_3dp2_bt.grid(row = 13, column = 3)

                            temp_d += 1
                            PW_3dp3_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PW_3dp3_bt.grid(row = 13, column = 4)

                            temp_d += 1
                            PW_3dp4_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PW_3dp4_bt.grid(row = 14, column = 0)

                            temp_d += 1
                            PS_inv_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PS_inv_bt.grid(row = 14, column = 1)

                            temp_d += 1
                            PW_inv1_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PW_inv1_bt.grid(row = 14, column = 2)    

                            temp_d += 1
                            PW_inv2_bt = Button(emp_win, text = emp_id[temp_d] + " " + emp_name[temp_d], width = 30, bg = emp_fo[temp_d])
                            PW_inv2_bt.grid(row = 14, column = 3) 

                        mchd_hfl_upd = []
                        mchd_hdp_upd = []
                        mchd_ckp_upd = []
                        mchd_lbl_upd = []
                        mchd_chp_upd = []
                        mchd_wds_upd = [] 
                        mchd_res_upd = []
                        def CNC_LM_upt(): #CNC Lathe Machine & CNC Milling Machine
                            loc = ("D:\Book1.xlsx") 
                            wb_mchd_udt = xlrd.open_workbook(loc) 
                            sr_mchd_udt = wb_mchd_udt.sheet_by_index(1)
                            sr_mchd_udt.cell_value(0, 0)
                            temp_ons = 'green'
                            temp_ofs = 'red'
                            mchd_hfl_upd.clear()
                            mchd_hdp_upd.clear()
                            mchd_ckp_upd.clear()
                            mchd_lbl_upd.clear()
                            mchd_chp_upd.clear()
                            mchd_wds_upd.clear()
                            mchd_res_upd.clear()

                            for temp_e in range(1, sr_mchd_udt.nrows):
                                if sr_mchd_udt.cell_value(temp_e,4) == 1.0:
                                    mchd_hfl_upd.append(temp_ons)
                                else:
                                    mchd_hfl_upd.append(temp_ofs)

                            for temp_e in range(1, sr_mchd_udt.nrows):
                                    if sr_mchd_udt.cell_value(temp_e,5) == 1.0:
                                        mchd_hdp_upd.append(temp_ons)
                                    else:
                                        mchd_hdp_upd.append(temp_ofs)

                            for temp_e in range(1, sr_mchd_udt.nrows):
                                if sr_mchd_udt.cell_value(temp_e,6) == 1.0:
                                    mchd_ckp_upd.append(temp_ons)
                                else:
                                    mchd_ckp_upd.append(temp_ofs)

                            for temp_e in range(1, sr_mchd_udt.nrows):
                                if sr_mchd_udt.cell_value(temp_e,7) == 1.0:
                                    mchd_lbl_upd.append(temp_ons)
                                else:
                                    mchd_lbl_upd.append(temp_ofs)

                            for temp_e in range(1, sr_mchd_udt.nrows):
                                if sr_mchd_udt.cell_value(temp_e,8) == 1.0:
                                    mchd_chp_upd.append(temp_ons)
                                else:
                                    mchd_chp_upd.append(temp_ofs)

                            for temp_e in range(1, sr_mchd_udt.nrows):
                                if sr_mchd_udt.cell_value(temp_e,9) == 1.0:
                                    mchd_wds_upd.append(temp_ons)
                                else:
                                    mchd_wds_upd.append(temp_ofs)

                            for temp_e in range(1, sr_mchd_udt.nrows):
                                if sr_mchd_udt.cell_value(temp_e,4) == 1.0:
                                    if sr_mchd_udt.cell_value(temp_e,5) == 1.0:
                                        if sr_mchd_udt.cell_value(temp_e,6) == 1.0:
                                            if sr_mchd_udt.cell_value(temp_e,7) == 1.0:
                                                if sr_mchd_udt.cell_value(temp_e,8) == 1.0:
                                                    if sr_mchd_udt.cell_value(temp_e,9) == 1.0:
                                                        mchd_res_upd.append('yellow')
                                                        col = 'K'
                                                        row = temp_e+1
                                                        col += str(row)
                                                        wr_mchd[col] = 1.0
                                                        xfile.save(loc)
                                                    elif sr_mchd_udt.cell_value(temp_e,9) == 2.0:
                                                        mchd_res_upd.append('green')
                                                    else:
                                                        mchd_res_upd.append('red')
                                                        col = 'K'
                                                        row = temp_e+1
                                                        col += str(row)
                                                        wr_mchd[col] = 0.0
                                                        xfile.save(loc)
                                                else:
                                                    mchd_res_upd.append('red')
                                                    col = 'K'
                                                    row = temp_e+1
                                                    col += str(row)
                                                    wr_mchd[col] = 0.0
                                                    xfile.save(loc)
                                            else:
                                                mchd_res_upd.append('red')
                                                col = 'K'
                                                row = temp_e+1
                                                col += str(row)
                                                wr_mchd[col] = 0.0
                                                xfile.save(loc)
                                        else:
                                            mchd_res_upd.append('red')
                                            col = 'K'
                                            row = temp_e+1
                                            col += str(row)
                                            wr_mchd[col] = 0.0
                                            xfile.save(loc)
                                    else:
                                        mchd_res_upd.append('red')
                                        col = 'K'
                                        row = temp_e+1
                                        col += str(row)
                                        wr_mchd[col] = 0.0
                                        xfile.save(loc)
                                else:
                                    mchd_res_upd.append('red')
                                    col = 'K'
                                    row = temp_e+1
                                    col += str(row)
                                    wr_mchd[col] = 0.0
                                    xfile.save(loc)
                            

                        if sr_empd.cell_value(i,4) == "Chief Executive Officer":  
                            emp_win_hr_bt = Button(emp_win, text = "Human Resource", width = 35, command = employee_attendance)
                            emp_win_hr_bt.grid (row = 6, column = 0)
                            emp_win_fn_bt = Button(emp_win, text = "Finance", width = 35)
                            emp_win_fn_bt.grid (row = 6, column = 1)
                            emp_win_inv_bt = Button(emp_win, text = "Inventory", width = 35)
                            emp_win_inv_bt.grid (row = 6, column = 2)
                            emp_win_sf_bt = Button(emp_win, text = "Shop-Floor", width = 35)
                            emp_win_sf_bt.grid (row = 6, column = 3)
                            emp_win_qr_bt = Button(emp_win, text = "Quality Reports", width = 35)
                            emp_win_qr_bt.grid (row = 6, column = 4)
                            
                        elif sr_empd.cell_value(i,4) == "Chief Finance Officer":
                            emp_win_D = Label(emp_win, text = "Chief Finance Officer")
                            emp_win_D.grid(row = 2, column = 0)
                                                  
                        elif sr_empd.cell_value(i,4) == "Chief Operations Manager":
                            emp_win_D = Label(emp_win, text = "Chief Operations Manager")
                            emp_win_D.grid(row = 2, column = 0)
                        
                        elif sr_empd.cell_value(i,4) == "Senior Manager Inventory & Warehousing":
                            emp_win_D = Label(emp_win, text = "Senior Manager Inventory & Warehousing")
                            emp_win_D.grid(row = 2, column = 0)
                      
                        elif sr_empd.cell_value(i,4) == "General Manager System":
                            emp_win_D = Label(emp_win, text = "General Manager System")
                            emp_win_D.grid(row = 2, column = 0)
                            
                        elif sr_empd.cell_value(i,4) == "Chief Administrative Officer":
                            emp_win_hr_bt = Button(emp_win, text = "Human Resource", width = 35, command = employee_attendance)
                            emp_win_hr_bt.grid (row = 7, column = 0)

                        elif sr_empd.cell_value(i,4) == "Operations Manager":
                            emp_win_D = Label(emp_win, text = "Operations Manager")
                            emp_win_D.grid(row = 2, column = 0)
                           
                        elif sr_empd.cell_value(i,4) == "Designing Manager":
                            emp_win_D = Label(emp_win, text = "Designing Manager")
                            emp_win_D.grid(row = 2, column = 0)
                            
                        elif sr_empd.cell_value(i,4) == "Quality Manager":
                            emp_win_D = Label(emp_win, text = "Quality Manager")
                            emp_win_D.grid(row = 2, column = 0)
                            
                        elif sr_empd.cell_value(i,4) == "Production Supervisor":                                                 
                                                      
                            def CNC_LM_PS():  
                                CNC_LM_upt()
                                if sr_empd.cell_value(i,5) == "CNC Lathe Machine":
                                    temp_ms = 1
                                else:
                                    temp_ms = 6                      

                                emp_win_mchdst_clm1_bt = Button(emp_win, text = sr_mchd.cell_value(temp_ms,1) + '\n' + sr_mchd.cell_value(temp_ms,2), 
                                                                width = 30,pady = 20, bg = mchd_res_upd[temp_ms-1])
                                emp_win_mchdst_clm1_bt.grid(row = 8,column = 1)
                                emp_win_mchdst_clm1_lb1 = Label(emp_win, text = " ")
                                emp_win_mchdst_clm1_lb1.grid(row = 9, column = 1)
                                emp_win_hflr_clm1_bt = Button(emp_win, text = "HYDRAULIC FLUID LEVEL", width=25, bg =  mchd_hfl_upd[temp_ms-1])
                                emp_win_hflr_clm1_bt.grid(row = 10, column = 1)
                                emp_win_hdpr_clm1_bt = Button(emp_win, text = "HYDRAULIC PRESSURE", width=25, bg = mchd_hdp_upd[temp_ms-1])
                                emp_win_hdpr_clm1_bt.grid(row = 11, column = 1)
                                emp_win_ckpr_clm1_bt = Button(emp_win, text = "CHUCK PRESSURE", width=25, bg = mchd_ckp_upd[temp_ms-1])
                                emp_win_ckpr_clm1_bt.grid(row = 12, column = 1)
                                emp_win_lblr_clm1_bt = Button(emp_win, text = "LUBE LEVEL", width=25, bg = mchd_lbl_upd[temp_ms-1])
                                emp_win_lblr_clm1_bt.grid(row = 13, column = 1)
                                emp_win_chpr_clm1_bt = Button(emp_win, text = "CHIPS", width=25, bg = mchd_chp_upd[temp_ms-1])
                                emp_win_chpr_clm1_bt.grid(row = 14, column = 1)
                                emp_win_wdsr_clm1_bt = Button(emp_win, text = "WIPE DOWN ALL SURFACES", width=25, bg = mchd_wds_upd[temp_ms-1])
                                emp_win_wdsr_clm1_bt.grid(row = 15, column = 1)
                                emp_win_mchdst_clm1_lb2 = Label(emp_win, text = " ")
                                emp_win_mchdst_clm1_lb2.grid(row = 16, column = 1)
                                if mchd_res_upd[temp_ms-1] == 'yellow':
                                    def CNC_LM_rdy1_stt():
                                        def CNC_LM_rdy1_stp():
                                            loc = 'D:\Book1.xlsx'
                                            wr_mchd['I2'] = 0.0
                                            wr_mchd['J2'] = 0.0
                                            wr_mchd['K2'] = 0.0
                                            xfile.save(loc)
                                            emp_win_mchdst_rdy1_bt.destroy()
                                            CNC_LM_PS()

                                        loc = 'D:\Book1.xlsx'
                                        emp_win_mchdst_rdy1_bt.configure(text = "STOP", bg = 'green', command = CNC_LM_rdy1_stp)
                                        emp_win_mchdst_clm1_bt.configure(bg = 'green')
                                        wr_mchd['K2'] = 2.0
                                        xfile.save(loc)
                              
                                        
                                    emp_win_mchdst_rdy1_bt = Button(emp_win, text = "START", width = 25, pady = 10, bg = 'red', command = CNC_LM_rdy1_stt)
                                    emp_win_mchdst_rdy1_bt.grid(row = 17, column = 1)


                                temp_ms += 1
                                emp_win_mchdst_clm2_bt = Button(emp_win, text = sr_mchd.cell_value(temp_ms,1) + '\n' + sr_mchd.cell_value(temp_ms,2),
                                                               width = 30,pady = 20, bg = mchd_res_upd[temp_ms-1])
                                emp_win_mchdst_clm2_bt.grid(row = 8,column = 2)
                                emp_win_mchdst_clm2_lb1 = Label(emp_win, text = " ")
                                emp_win_mchdst_clm2_lb1.grid(row = 9, column = 2)
                                emp_win_hflr_clm2_bt = Button(emp_win, text = "HYDRAULIC FLUID LEVEL", width=25, bg =  mchd_hfl_upd[temp_ms-1])
                                emp_win_hflr_clm2_bt.grid(row = 10, column = 2)
                                emp_win_hdpr_clm2_bt = Button(emp_win, text = "HYDRAULIC PRESSURE", width=25, bg = mchd_hdp_upd[temp_ms-1])
                                emp_win_hdpr_clm2_bt.grid(row = 11, column = 2)
                                emp_win_ckpr_clm2_bt = Button(emp_win, text = "CHUCK PRESSURE", width=25, bg = mchd_ckp_upd[temp_ms-1])
                                emp_win_ckpr_clm2_bt.grid(row = 12, column = 2)
                                emp_win_lblr_clm2_bt = Button(emp_win, text = "LUBE LEVEL", width=25, bg = mchd_lbl_upd[temp_ms-1])
                                emp_win_lblr_clm2_bt.grid(row = 13, column = 2)
                                emp_win_chpr_clm2_bt = Button(emp_win, text = "CHIPS", width=25, bg = mchd_chp_upd[temp_ms-1])
                                emp_win_chpr_clm2_bt.grid(row = 14, column = 2)
                                emp_win_wdsr_clm2_bt = Button(emp_win, text = "WIPE DOWN ALL SURFACES", width=25, bg = mchd_wds_upd[temp_ms-1])
                                emp_win_wdsr_clm2_bt.grid(row = 15, column = 2)
                                emp_win_mchdst_clm2_lb2 = Label(emp_win, text = " ")
                                emp_win_mchdst_clm2_lb2.grid(row = 16, column = 2)
                                if mchd_res_upd[temp_ms-1] == 'yellow':
                                    def CNC_LM_rdy2_stt():
                                        def CNC_LM_rdy2_stp():
                                            loc = 'D:\Book1.xlsx'
                                            wr_mchd['I3'] = 0.0
                                            wr_mchd['J3'] = 0.0
                                            wr_mchd['K3'] = 0.0
                                            xfile.save(loc)
                                            emp_win_mchdst_rdy2_bt.destroy()
                                            CNC_LM_PS()

                                        loc = 'D:\Book1.xlsx'
                                        emp_win_mchdst_rdy2_bt.configure(text = "STOP", bg = 'green', command = CNC_LM_rdy2_stp)
                                        emp_win_mchdst_clm2_bt.configure(bg = 'green')
                                        wr_mchd['K3'] = 2.0
                                        xfile.save(loc)
                                        
                                    emp_win_mchdst_rdy2_bt = Button(emp_win, text = "START", width = 25, pady = 10, bg = 'red', command = CNC_LM_rdy2_stt)
                                    emp_win_mchdst_rdy2_bt.grid(row = 17, column = 2)

                                temp_ms += 1
                                emp_win_mchdst_clm3_bt = Button(emp_win, text = sr_mchd.cell_value(temp_ms,1) + '\n' + sr_mchd.cell_value(temp_ms,2),
                                                               width = 30,pady = 20, bg = mchd_res_upd[temp_ms-1])
                                emp_win_mchdst_clm3_bt.grid(row = 8,column = 3)
                                emp_win_mchdst_clm3_lb1 = Label(emp_win, text = " ")
                                emp_win_mchdst_clm3_lb1.grid(row = 9, column = 3)
                                emp_win_hflr_clm3_bt = Button(emp_win, text = "HYDRAULIC FLUID LEVEL", width=25, bg =  mchd_hfl_upd[temp_ms-1])
                                emp_win_hflr_clm3_bt.grid(row = 10, column = 3)
                                emp_win_hdpr_clm3_bt = Button(emp_win, text = "HYDRAULIC PRESSURE", width=25, bg = mchd_hdp_upd[temp_ms-1])
                                emp_win_hdpr_clm3_bt.grid(row = 11, column = 3)
                                emp_win_ckpr_clm3_bt = Button(emp_win, text = "CHUCK PRESSURE", width=25, bg = mchd_ckp_upd[temp_ms-1])
                                emp_win_ckpr_clm3_bt.grid(row = 12, column = 3)
                                emp_win_lblr_clm3_bt = Button(emp_win, text = "LUBE LEVEL", width=25, bg = mchd_lbl_upd[temp_ms-1])
                                emp_win_lblr_clm3_bt.grid(row = 13, column = 3)
                                emp_win_chpr_clm3_bt = Button(emp_win, text = "CHIPS", width=25, bg = mchd_chp_upd[temp_ms-1])
                                emp_win_chpr_clm3_bt.grid(row = 14, column = 3)
                                emp_win_wdsr_clm3_bt = Button(emp_win, text = "WIPE DOWN ALL SURFACES", width=25, bg = mchd_wds_upd[temp_ms-1])
                                emp_win_wdsr_clm3_bt.grid(row = 15, column = 3)
                                emp_win_mchdst_clm3_lb2 = Label(emp_win, text = " ")
                                emp_win_mchdst_clm3_lb2.grid(row = 16, column =3)
                                if mchd_res_upd[temp_ms-1] == 'yellow':
                                    def CNC_LM_rdy3_stt():
                                        def CNC_LM_rdy3_stp():
                                            loc = 'D:\Book1.xlsx'
                                            wr_mchd['I4'] = 0.0
                                            wr_mchd['J4'] = 0.0
                                            wr_mchd['K4'] = 0.0
                                            xfile.save(loc)
                                            emp_win_mchdst_rdy3_bt.destroy()
                                            CNC_LM_PS()

                                        loc = 'D:\Book1.xlsx'
                                        emp_win_mchdst_rdy3_bt.configure(text = "STOP", bg = 'green', command = CNC_LM_rdy3_stp)
                                        emp_win_mchdst_clm3_bt.configure(bg = 'green')
                                        wr_mchd['K4'] = 2.0
                                        xfile.save(loc)
                                        
                                    emp_win_mchdst_rdy3_bt = Button(emp_win, text = "START", width = 25, pady = 10, bg = 'red', command = CNC_LM_rdy3_stt)
                                    emp_win_mchdst_rdy3_bt.grid(row = 17, column = 3)

                                temp_ms += 1
                                emp_win_mchdst_clm4_bt = Button(emp_win, text = sr_mchd.cell_value(temp_ms,1) + '\n' + sr_mchd.cell_value(temp_ms,2),
                                                               width = 30,pady = 20, bg = mchd_res_upd[temp_ms-1])
                                emp_win_mchdst_clm4_bt.grid(row = 8,column = 4)
                                emp_win_mchdst_clm4_lb1 = Label(emp_win, text = " ")
                                emp_win_mchdst_clm4_lb1.grid(row = 9, column = 4)
                                emp_win_hflr_clm4_bt = Button(emp_win, text = "HYDRAULIC FLUID LEVEL", width=25, bg =  mchd_hfl_upd[temp_ms-1])
                                emp_win_hflr_clm4_bt.grid(row = 10, column = 4)
                                emp_win_hdpr_clm4_bt = Button(emp_win, text = "HYDRAULIC PRESSURE", width=25, bg = mchd_hdp_upd[temp_ms-1])
                                emp_win_hdpr_clm4_bt.grid(row = 11, column = 4)
                                emp_win_ckpr_clm4_bt = Button(emp_win, text = "CHUCK PRESSURE", width=25, bg = mchd_ckp_upd[temp_ms-1])
                                emp_win_ckpr_clm4_bt.grid(row = 12, column = 4)
                                emp_win_lblr_clm4_bt = Button(emp_win, text = "LUBE LEVEL", width=25, bg = mchd_lbl_upd[temp_ms-1])
                                emp_win_lblr_clm4_bt.grid(row = 13, column = 4)
                                emp_win_chpr_clm4_bt = Button(emp_win, text = "CHIPS", width=25, bg = mchd_chp_upd[temp_ms-1])
                                emp_win_chpr_clm4_bt.grid(row = 14, column = 4)
                                emp_win_wdsr_clm4_bt = Button(emp_win, text = "WIPE DOWN ALL SURFACES", width=25, bg = mchd_wds_upd[temp_ms-1])
                                emp_win_wdsr_clm4_bt.grid(row = 15, column = 4)
                                emp_win_mchdst_clm4_lb2 = Label(emp_win, text = " ")
                                emp_win_mchdst_clm4_lb2.grid(row = 16, column = 4)
                                if mchd_res_upd[temp_ms-1] == 'yellow':
                                    def CNC_LM_rdy4_stt():
                                        def CNC_LM_rdy4_stp():
                                            loc = 'D:\Book1.xlsx'
                                            wr_mchd['I5'] = 0.0
                                            wr_mchd['J5'] = 0.0
                                            wr_mchd['K5'] = 0.0
                                            xfile.save(loc)
                                            emp_win_mchdst_rdy4_bt.destroy()
                                            CNC_LM_PS()

                                        loc = 'D:\Book1.xlsx'
                                        emp_win_mchdst_rdy4_bt.configure(text = "STOP", bg = 'green', command = CNC_LM_rdy4_stp)
                                        emp_win_mchdst_clm4_bt.configure(bg = 'green')
                                        wr_mchd['K5'] = 2.0
                                        xfile.save(loc)
                                        
                                    emp_win_mchdst_rdy4_bt = Button(emp_win, text = "START", width = 25, pady = 10, bg = 'red', command = CNC_LM_rdy4_stt)
                                    emp_win_mchdst_rdy4_bt.grid(row = 17, column = 4)

                                temp_ms += 1
                                emp_win_mchdst_clm5_bt = Button(emp_win, text = sr_mchd.cell_value(temp_ms,1) + '\n' + sr_mchd.cell_value(temp_ms,2),
                                                               width = 30,pady = 20, bg = mchd_res_upd[temp_ms-1])
                                emp_win_mchdst_clm5_bt.grid(row = 8,column = 5)
                                emp_win_mchdst_clm5_lb1 = Label(emp_win, text = " ")
                                emp_win_mchdst_clm5_lb1.grid(row = 9, column = 5)
                                emp_win_hflr_clm5_bt = Button(emp_win, text = "HYDRAULIC FLUID LEVEL", width=25, bg =  mchd_hfl_upd[temp_ms-1])
                                emp_win_hflr_clm5_bt.grid(row = 10, column = 5)
                                emp_win_hdpr_clm5_bt = Button(emp_win, text = "HYDRAULIC PRESSURE", width=25, bg = mchd_hdp_upd[temp_ms-1])
                                emp_win_hdpr_clm5_bt.grid(row = 11, column = 5)
                                emp_win_ckpr_clm5_bt = Button(emp_win, text = "CHUCK PRESSURE", width=25, bg = mchd_ckp_upd[temp_ms-1])
                                emp_win_ckpr_clm5_bt.grid(row = 12, column = 5)
                                emp_win_lblr_clm5_bt = Button(emp_win, text = "LUBE LEVEL", width=25, bg = mchd_lbl_upd[temp_ms-1])
                                emp_win_lblr_clm5_bt.grid(row = 13, column = 5)
                                emp_win_chpr_clm5_bt = Button(emp_win, text = "CHIPS", width=25, bg = mchd_chp_upd[temp_ms-1])
                                emp_win_chpr_clm5_bt.grid(row = 14, column = 5)
                                emp_win_wdsr_clm5_bt = Button(emp_win, text = "WIPE DOWN ALL SURFACES", width=25, bg = mchd_wds_upd[temp_ms-1])
                                emp_win_wdsr_clm5_bt.grid(row = 15, column = 5)
                                emp_win_mchdst_clm5_lb2 = Label(emp_win, text = " ")
                                emp_win_mchdst_clm5_lb2.grid(row = 16, column = 5)
                                if mchd_res_upd[temp_ms-1] == 'yellow':
                                    def CNC_LM_rdy5_stt():
                                        def CNC_LM_rdy5_stp():
                                            loc = 'D:\Book1.xlsx'
                                            wr_mchd['I6'] = 0.0
                                            wr_mchd['J6'] = 0.0
                                            wr_mchd['K6'] = 0.0
                                            xfile.save(loc)
                                            emp_win_mchdst_rdy5_bt.destroy()
                                            CNC_LM_PS()

                                        loc = 'D:\Book1.xlsx'
                                        emp_win_mchdst_rdy5_bt.configure(text = "STOP", bg = 'green', command = CNC_LM_rdy5_stp)
                                        emp_win_mchdst_clm5_bt.configure(bg = 'green')
                                        wr_mchd['K6'] = 2.0
                                        xfile.save(loc)
                                        
                                    emp_win_mchdst_rdy5_bt = Button(emp_win, text = "START", width = 25, pady = 10, bg = 'red', command = CNC_LM_rdy5_stt)
                                    emp_win_mchdst_rdy5_bt.grid(row = 17, column = 5)

                               
                            if sr_empd.cell_value(i,5) == "CNC Lathe Machine":
                                emp_win_mchdst_bt = Button(emp_win, text = "CNC Lathe Machine Status", width = 35, command = CNC_LM_PS)
                                emp_win_mchdst_bt.grid(row = 6, column = 0)
                             
                            elif sr_empd.cell_value(i,5) == "CNC Milling Machine":
                                emp_win_mchdst_bt = Button(emp_win, text = "CNC Milling Machine Status", width = 35, command = CNC_LM_PS)
                                emp_win_mchdst_bt.grid(row = 6, column = 0)


                            elif sr_empd.cell_value(i,5) == "CNC Cutting Machine":
                                emp_win_D = Label(emp_win, text = "Department: CNC Cutting Machine")
                                emp_win_D.grid(row = 3, column = 0)
                                                            
                            elif sr_empd.cell_value(i,5) == "3D Printer":
                                emp_win_D = Label(emp_win, text = "Department: 3D Printer")
                                emp_win_D.grid(row = 3, column = 0)

                            elif sr_empd.cell_value(i,5) ==  "Inventory":
                                emp_win_D = Label(emp_win, text = "Department: Inventory")
                                emp_win_D.grid(row = 3, column = 0)
                                                                     
                      

                        elif sr_empd.cell_value(i,4) == "Production Worker":
                                                                     
                            for temp_ms in range(1, sr_mchd.nrows):
                                                                
                                def CNC_LM(): #CNC Lathe Machine & CNC Milling Machine

                                    CNC_LM_upt()
                                   
                                    def CNC_LM_hfl_udt():
                                        CNC_LM_upt()
                                        col = 'E'
                                        row = temp_ms+1
                                        col += str(row)
                                        if mchd_hfl_upd[temp_ms-1] == 'green':
                                            mchd_hfl_upd[temp_ms-1] = 'red'
                                            emp_mob_hflr_bt.configure(bg =  mchd_hfl_upd[temp_ms-1])
                                            wr_mchd[col] = 0.0
                                            xfile.save(loc)
                                        else:
                                            mchd_hfl_upd[temp_ms-1] = 'green'
                                            emp_mob_hflr_bt.configure(bg = mchd_hfl_upd[temp_ms-1])
                                            wr_mchd[col] = 1.0
                                            xfile.save(loc)

                                    def CNC_LM_hdp_udt():
                                        CNC_LM_upt()
                                        col = 'F'
                                        row = temp_ms+1
                                        col += str(row)
                                        if mchd_hdp_upd[temp_ms-1] == 'green':
                                            mchd_hdp_upd[temp_ms-1] = 'red'
                                            emp_mob_hdpr_bt.configure(bg =  mchd_hdp_upd[temp_ms-1])
                                            wr_mchd[col] = 0.0
                                            xfile.save(loc)
                                        else:
                                            mchd_hdp_upd[temp_ms-1] = 'green'
                                            emp_mob_hdpr_bt.configure(bg = mchd_hdp_upd[temp_ms-1])
                                            wr_mchd[col] = 1.0
                                            xfile.save(loc)

                                    def CNC_LM_ckp_udt():
                                        CNC_LM_upt()
                                        col = 'G'
                                        row = temp_ms+1
                                        col += str(row)
                                        if mchd_ckp_upd[temp_ms-1] == 'green':
                                            mchd_ckp_upd[temp_ms-1] = 'red'
                                            emp_mob_ckpr_bt.configure(bg =  mchd_ckp_upd[temp_ms-1])
                                            wr_mchd[col] = 0.0
                                            xfile.save(loc)
                                        else:
                                            mchd_ckp_upd[temp_ms-1] = 'green'
                                            emp_mob_ckpr_bt.configure(bg = mchd_ckp_upd[temp_ms-1])
                                            wr_mchd[col] = 1.0
                                            xfile.save(loc)

                                    def CNC_LM_lbl_udt():
                                        CNC_LM_upt()
                                        col = 'H'
                                        row = temp_ms+1
                                        col += str(row)
                                        if mchd_lbl_upd[temp_ms-1] == 'green':
                                            mchd_lbl_upd[temp_ms-1] = 'red'
                                            emp_mob_lblr_bt.configure(bg =  mchd_lbl_upd[temp_ms-1])
                                            wr_mchd[col] = 0.0
                                            xfile.save(loc)
                                        else:
                                            mchd_lbl_upd[temp_ms-1] = 'green'
                                            emp_mob_lblr_bt.configure(bg = mchd_lbl_upd[temp_ms-1])
                                            wr_mchd[col] = 1.0
                                            xfile.save(loc)

                                    def CNC_LM_chp_udt():
                                        CNC_LM_upt()
                                        col = 'I'
                                        row = temp_ms+1
                                        col += str(row)
                                        if mchd_chp_upd[temp_ms-1] == 'green':
                                            mchd_chp_upd[temp_ms-1] = 'red'
                                            emp_mob_chpr_bt.configure(bg =  mchd_chp_upd[temp_ms-1])
                                            wr_mchd[col] = 0.0
                                            xfile.save(loc)
                                        else:
                                            mchd_chp_upd[temp_ms-1] = 'green'
                                            emp_mob_chpr_bt.configure(bg = mchd_chp_upd[temp_ms-1])
                                            wr_mchd[col] = 1.0
                                            xfile.save(loc)

                                    def CNC_LM_wds_udt():
                                        CNC_LM_upt()
                                        col = 'J'
                                        row = temp_ms+1
                                        col += str(row)
                                        if mchd_wds_upd[temp_ms-1] == 'green':
                                            mchd_wds_upd[temp_ms-1] = 'red'
                                            emp_mob_wdsr_bt.configure(bg =  mchd_wds_upd[temp_ms-1])
                                            wr_mchd[col] = 0.0
                                            xfile.save(loc)
                                        else:
                                            mchd_wds_upd[temp_ms-1] = 'green'
                                            emp_mob_wdsr_bt.configure(bg = mchd_wds_upd[temp_ms-1])
                                            wr_mchd[col] = 1.0
                                            xfile.save(loc)

                                    emp_mob_mid_bt = Button(emp_mob, text = "MACHINE ID", width=35)
                                    emp_mob_mid_bt.grid(row = 7, column = 0)
                                    emp_mob_hfl_bt = Button(emp_mob, text = "HYDRAULIC FLUID LEVEL", width=35)
                                    emp_mob_hfl_bt.grid(row = 8, column = 0)
                                    emp_mob_hdp_bt = Button(emp_mob, text = "HYDRAULIC PRESSURE", width=35)
                                    emp_mob_hdp_bt.grid(row = 9, column = 0)
                                    emp_mob_ckp_bt = Button(emp_mob, text = "CHUCK PRESSURE", width=35)
                                    emp_mob_ckp_bt.grid(row = 10, column = 0)
                                    emp_mob_lbl_bt = Button(emp_mob, text = "LUBE LEVEL", width=35)
                                    emp_mob_lbl_bt.grid(row = 11, column = 0)
                                    emp_mob_chp_bt = Button(emp_mob, text = "CHIPS", width=35)
                                    emp_mob_chp_bt.grid(row = 12, column = 0)
                                    emp_mob_wds_bt = Button(emp_mob, text = "WIPE DOWN ALL SURFACES", width=35)
                                    emp_mob_wds_bt.grid(row = 13, column = 0)
                                    emp_mob_l2 = Label(emp_mob, text = " ")
                                    emp_mob_l2.grid(row = 7, column = 1)

                                    emp_mob_hflr_bt = Button(emp_mob, text = "READY", width=15, bg =  mchd_hfl_upd[temp_ms-1], command = CNC_LM_hfl_udt)
                                    emp_mob_hflr_bt.grid(row = 8, column = 2)
                                    emp_mob_hdpr_bt = Button(emp_mob, text = "READY", width=15, bg = mchd_hdp_upd[temp_ms-1], command = CNC_LM_hdp_udt)
                                    emp_mob_hdpr_bt.grid(row = 9, column = 2)
                                    emp_mob_ckpr_bt = Button(emp_mob, text = "READY", width=15, bg = mchd_ckp_upd[temp_ms-1], command = CNC_LM_ckp_udt)
                                    emp_mob_ckpr_bt.grid(row = 10, column = 2)
                                    emp_mob_lblr_bt = Button(emp_mob, text = "READY", width=15, bg = mchd_lbl_upd[temp_ms-1], command = CNC_LM_lbl_udt)
                                    emp_mob_lblr_bt.grid(row = 11, column = 2)
                                    emp_mob_chpr_bt = Button(emp_mob, text = "READY", width=15, bg = mchd_chp_upd[temp_ms-1], command = CNC_LM_chp_udt)
                                    emp_mob_chpr_bt.grid(row = 12, column = 2)
                                    emp_mob_wdsr_bt = Button(emp_mob, text = "READY", width=15, bg = mchd_wds_upd[temp_ms-1], command = CNC_LM_wds_udt)
                                    emp_mob_wdsr_bt.grid(row = 13, column = 2)

                                if sr_empd.cell_value(i,0) == sr_mchd.cell_value(temp_ms,0): # Employee ID verifying with Machine_Data Sheet

                                    if sr_mchd.cell_value(temp_ms,3) == "CNC Lathe Machine":
                                        emp_mob_mch = Label(emp_mob, text = "Machine: " + sr_mchd.cell_value(temp_ms,2))
                                        emp_mob_mch.grid(row = 4, column = 0)
                                        CNC_LM()
                                        emp_mob_mchid = Label(emp_mob, text = sr_mchd.cell_value(temp_ms,1), font = ('bold'))
                                        emp_mob_mchid.grid(row = 7, column = 2)
                                        break

                                    elif sr_mchd.cell_value(temp_ms,3) == "CNC Milling Machine":
                                        emp_mob_mch = Label(emp_mob, text = "Machine: " + sr_mchd.cell_value(temp_ms,2))
                                        emp_mob_mch.grid(row = 4, column = 0)
                                        CNC_LM()
                                        emp_mob_mchid = Label(emp_mob, text = sr_mchd.cell_value(temp_ms,1), font = ('bold'))
                                        emp_mob_mchid.grid(row = 7, column = 2)
                                        break
                                        
                                    elif sr_mchd.cell_value(j,3) == "CNC Cutting Machine":
                                        emp_mob_mch = Label(emp_mob, text = "Machine: " + sr_mchd.cell_value(temp_ms,2))
                                        emp_mob_mch.grid(row = 4, column = 0)
                                        emp_mob_gg = Button(emp_mob, text = "MACHINE ID", width=35)
                                        emp_mob_gg.grid(row = 7, column = 0)
                                        emp_mob_mchid = Label(emp_mob, text = sr_mchd.cell_value(temp_ms,1), font = ('bold'))
                                        emp_mob_mchid.grid(row = 7, column = 2)
                                        break
                                                                                                                                            
                                    elif sr_mchd.cell_value(temp_ms,3) == "3D Printer":
                                        emp_mob_mch = Label(emp_mob, text = "Machine: " + sr_mchd.cell_value(temp_ms,2))
                                        emp_mob_mch.grid(row = 4, column = 0)
                                        emp_mob_gg = Button(emp_mob, text = "MACHINE ID", width=35)
                                        emp_mob_gg.grid(row = 7, column = 0)
                                        emp_mob_mchid = Label(emp_mob, text = sr_mchd.cell_value(temp_ms,1), font = ('bold'))
                                        emp_mob_mchid.grid(row = 7, column = 2)
                                        break

                                else:
                                    if sr_empd.cell_value(i,5) ==  "Inventory":
                                        emp_mob_dpt = Label (emp_mob, text = "Department: " + sr_empd.cell_value(i,5))
                                        emp_mob_dpt.grid(row = 3, column = 0)
        
                        emp_mob.mainloop()   
                        emp_win.mainloop()
                       
                    else:
                        ety_emp_fo = Label(main_win, text = "Employee Off-Site!", fg= 'red', font = ('bold')).pack()
                        
                    break 

                else:
                    ety_pass_lb = Label(main_win, text = "Please Enter Correct Password!", fg = 'red').pack()
                    break

 
        if temp_user == sr_empd.nrows-1:
                ety_user_lb = Label(main_win, text = "Please Enter Correct Username!", fg = 'red').pack()
    

def HP():  # Hide Password
    shd_pass.configure(text = "Show Password", command = SP)
    et_pass.configure(show='*')

def SP():  # Show Password
    shd_pass.configure(text="Hide Password", command = HP)
    et_pass.configure(show = '')

lb_user = Label(main_win, text = "USERNAME*")
et_user = Entry(main_win, width = 50)
lb_pass = Label(main_win, text = "PASSWORD*")
et_pass = Entry(main_win, width = 50, show='*')
shd_pass = Button(main_win, text = "Show Password", command = SP)
login_bt = Button(main_win, text = "LOGIN", padx = 30, font = ('bold'), command = Login)
main_lbl = Label(main_win,text="    ")
fgt_pass = Button(main_win, text = "Forgot Password", padx = 12)

lb_user.pack()
et_user.pack()
lb_pass.pack()
et_pass.pack()
shd_pass.pack()
login_bt.pack()
main_lbl.pack()
fgt_pass.pack()

main_win.mainloop()
