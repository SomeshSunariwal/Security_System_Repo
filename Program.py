import RPi.GPIO as a
import xlrd,xlwt,time
from xlutils.copy import copy
import smtplib as b
import datetime

a.setmode(a.BOARD)
a.setwarnings(0)
############################################### IR Secton #########################################################
a.setup(37,a.IN)
def ir():
    return(a.input(37))
############################################### Data_Base Section #################################################
file_loc = "Data_Base.xls"
Book1 = xlrd.open_workbook(file_loc)
Book2 = Book1.sheet_by_index(0)
Book3 = copy(Book1)
Book4 = Book3.get_sheet(0)
Book4.write(0,0,"Password")
Book3.save("Data_Base.xls")

################################################  Mail Section  ###################################################

def mail(msg):
    server = b.SMTP("smtp.gmail.com",587)
    server.starttls()
    server.login("Someshsunariwal@gmail.com","Password")
    server.sendmail("Someshsunariwal@gmail.com","Someshsunariwal0@gmail.com",msg)
    server.close()

################################################  lcd_Section   ###################################################

rs,en,d0,d1,d2,d3,d4,d5,d6,d7 = 3,5,7,11,13,15,19,21,23,29
list1 = [0X01,0X02,0X06,0X38,0X0C]
list2 = [d0,d1,d2,d3,d4,d5,d6,d7,rs,en]
list3 = [0X01,0X02,0X04,0X08,0X10,0X20,0X40,0X80]

for i in list2:
    a.setup(list2,a.OUT)

def pin(data):
    for n in range(0,8):
        if(data&(list3[n]) == list3[n]):
           a.output(list2[n],1)
        else:
           a.output(list2[n],0)

def lcd_init():
    for p in list1:
        lcd_data(p)

def lcd_data(disp):
    if type(disp) is int:
        x = 0
        pin(disp)
    else:
        val = ord(disp)
        pin(val)
        x=1
    a.output(rs,x)
    a.output(en,1)
    time.sleep(0.01)
    a.output(en,0)
    time.sleep(0.01)

def lcd_string(name,):
    for j in name:
        lcd_data(j)
        
lcd_init()
lcd_data(0X85)
lcd_string("WELCOME")
lcd_data(0xC0)
lcd_string("   SOMESH CODE")
time.sleep(4)
lcd_data(0X01)
####################################################  Keypad_section  ##################################################
Row = [8,10,12,16]
Col = [18,22,24,26]

for k in range(4):
    a.setup(Col[k],a.OUT)
    
for l in range(4):
    a.setup(Row[l], a.IN ,pull_up_down = a.PUD_UP)

def Matrix_call():
    Matrix = [['D','C','B','A'],
              ['#',9,6,3],
              [0,8,5,2],
              ['*',7,4,1]]

    for k in range(4):
        a.output(Col[k],1)

    while(True):
        for m in range(0,4):
            time.sleep(.01)
            a.output(Col[m],0)
            for o in range(0,4):
                time.sleep(.01)
                if(a.input(Row[o]) == 0):
                    time.sleep(.01)
                    input_store = Matrix[o][m]
                    return(input_store)
            a.output(Col[m],1)


##############################################           Main Program           ###################################

def lcd_msg(msg222,msg333 = ""):
    lcd_data(0X80)
    lcd_string(msg222)
    lcd_data(0XC0)
    lcd_string(msg333)

############ Time Section #############

def lcd_time():
    time_sec = datetime.datetime.today()
    time_sec1 = time_sec.time()
    time_sec2 = time_sec.date()
    lcd_msg("     TIME",str(time_sec1))
    time.sleep(2)
    lcd_data(0X01)
    lcd_msg("     DATE",str(time_sec2))
    time.sleep(2)
    lcd_data(0X01)

def wrong():
    lcd_msg("ALERT!  ALERT!","ALERT!  ALERT!")
    time.sleep(5)
    lcd_data(0X01)
    lcd_msg("MAIL SEND")
    time.sleep(5)

list11,list22,list33,list44,list55,list66 = [],[],[],[],[],[]

def list_call():
    for temp1 in range(6):
        int2 = Matrix_call()
        list22.append(int2)
        print(int2)
        lcd_string(str(int2))
        time.sleep(.2)
    return(list22)

try:
    while(True):
        if(ir() == 0):
            print("A. OPEN DOOR\nB. SETTING")
            lcd_data(0X01)
            lcd_msg("A. OPEN DOOR","B. SETTING")
            Book1 = xlrd.open_workbook(file_loc)
            Book2 = Book1.sheet_by_index(0)
            Book3 = copy(Book1)
            Book4 = Book3.get_sheet(0)
            Book3.save("Data_Base.xls")
            R_v = Book2.nrows
            R_v = R_v - 1
            print(R_v)                                                     ### DELETE THIS

            if(Matrix_call() == "A"):                                                  ##### First Input
                lcd_data(0X01)
                print("ENTER 6 DIGIT CODE")
                lcd_msg("ENTER 6 DIGIT CODE")
                data1 = Book2.cell_value(R_v,0)                          ### 6 Digit code from row
                print(data1)
            
                for round1 in range(1,4):
                    lcd_data(0XC0)
                    list11 = list_call()
                    print(list11)
                        
                    if(str(list11) == data1):
                        print("DOOR UNLOCKED")
                        lcd_data(0X01)
                        lcd_msg("WELCOME HOME","DOOR UNLOCKED")
                        time.sleep(5)
                        list11.clear()
                        break
                    else:
                        list11.clear()
                        lcd_data(0XC0)
                        lcd_data(0X01)
                        lcd_msg("WRONG PASSCODE","ATTEMP : "+str(round1))
                        print("Wrong Passcode")
                        time.sleep(2)
                        lcd_data(0X01)
                        lcd_msg("ENTER 6 DIGIT CODE")
                        if(round1 == 3):
                            lcd_data(0X01)
                            wrong()
                        #   mail("Wrong Person want to enter")
                            print("Mail Sent")                                ###### Test Mail Send
                            break
                         
            elif(Matrix_call() == "B"):                                                ####  Second Input
                print("A. CHANGE PASSCODE\nB. DEFAULT PASSCODE")
                lcd_msg("A.CHANGE PASSCODE","B.DEFAULT PASSCODE")
                input2 = Matrix_call()
                flag1,flag2 = 0,0
                if(input2 == "A"):
                    lcd_data(0X01)
                    print("Enter old Password")
                    lcd_msg("ENTER OLD PASSCODE")
                    data2 = Book2.cell_value(R_v,0)     ### 6 Digit code from row
                    print(data2)

                    for round2 in range(1,4):
                        lcd_data(0XC0)
                        if(flag1 == 1):
                            break
                        list22 = list_call()
                        print(list22)

                        if(str(list22) == data2):
                            list22.clear()
                            for round3 in range(1,4):
                                if(flag2 == 1):
                                    break
                                print("ENTER NEW CODE :")
                                lcd_data(0X01)
                                lcd_msg("ENTER NEW CODE :")
                                
                                list33 = list_call()
                                print(list33)
                                print("ENTER AGAIN :")

                                lcd_data(0X01)
                                lcd_msg("ENTER AGAIN :")
                                
                                list44 = list_call()
                                list55 = list44
                                list66 = list44
                                
                                print(list44)
                                    
                                if(list33 == list44):
                                    list33.clear()
                                    list44.clear()
                                    flag1 = 1
                                    for check in range(1,R_v):
                                        z = Book2.cell_value(check,0)
                                        print(z)
                                        print(list55)
                                        if(str(list55) == z):
                                            list55.clear()
                                            list66.clear()
                                            print("Pass")
                                            lcd_data(0X01)
                                            lcd_msg("ALREADY USED","  BEFORE")
                                            time.sleep(2)
                                        else:
                                            R_v1 = R_v + 1
                                            print(R_v)
                                            print(list66)
                                            Book4.write(R_v1,0,str(list66))
                                            list66.clear()
                                            list55.clear()
                                            Book3.save("Data_Base.xls")
                                            lcd_data(0X01)
                                            lcd_msg("CHANGED ","SUCSSES-FULL")
                                            time.sleep(3)
                                            flag2 = 1
                                            break
                                
                                else:
                                    list33.clear()
                                    list44.clear()
                                    list66.clear()
                                    list55.clear()
                                    print("NOT MATCH")
                                    lcd_data(0X01)
                                    lcd_msg("  NOT MATCH  ","     RETRY")
                                    time.sleep(2)
                                    lcd_data(0x01)
                                    if(round3 == 3):
                                        break
                        
                        else:
                            list22.clear()
                            print("Enter Wrong old Passcode")                  ##### print Statement
                            lcd_data(0XC0)
                            lcd_data(0X01)
                            lcd_msg("WRONG PASSCODE","ATTEMP :"+str(round2))                   ###### Wrong Passcode lcd
                            time.sleep(1)
                            lcd_data(0X01)
                            lcd_msg("ENTER OLD PASSCODE")
                            if(round2 == 3):
                                lcd_data(0X01)
                                wrong()
                             #  mail(" ALERT! SOMEONE WANT TO CHANGE YOUR PASSCODE")
                                break
                                
                elif(input2 == "B"):
                    data1 = Book2.cell_value(1,0)
                    lcd_data(0X01)
                    print("Enter old Password")
                    lcd_msg("ENTER OLD PASSCODE")
                    data2 = Book2.cell_value(R_v,0)     ### 6 Digit code from row
                    print(data2)

                    for round2 in range(1,4):
                        lcd_data(0XC0)
                        
                        list22 = list_call()
                        print(list22)

                        if(str(list22) == data2):
                            list22.clear()
                            algo= [1, 1, 1, 1, 1, 1]
                            R_v1 = R_v + 1
                            print(R_v)
                            Book4.write(R_v1,0,str(algo))
                            Book3.save("Data_Base.xls")
                            lcd_data(0X01)
                            lcd_msg("CHANGED ","SUCSSES-FULL")
                            time.sleep(3)
                            break
                        else:
                            list22.clear()
                            print("Enter Wrong old Passcode")                  ##### print Statement
                            lcd_data(0X01)
                            lcd_msg("WRONG PASSCODE","ATTEMP :"+str(round2))                   ###### Wrong Passcode lcd
                            time.sleep(2)
                            lcd_data(0X01)
                            lcd_msg("ENTER OLD PASSCODE")
                            if(round2 == 3):
                                lcd_data(0X01)
                                wrong()
                             #  mail(" ALERT! SOMEONE WANT TO CHANGE YOUR PASSCODE")
                                break
        elif(ir() == 1):
            lcd_data(0x01)
            lcd_time()
finally:
    lcd_data(0X01)
    a.cleanup()
