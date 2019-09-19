# -*- coding: utf-8 -*-
"""
Created on Thu Sep 12 11:54:04 2019

@author: malte
"""

#import numpy as np

import subprocess
import os 
import time

desktop = XSCRIPTCONTEXT.getDesktop()
model = desktop.getCurrentComponent()
#check whether there's already an opened document. Otherwise, create a new one
if not hasattr(model, "Sheets"):
    model = desktop.loadComponentFromURL("private:factory/scalc","_blank", 0, () )
    #get the XText interface

###CHECKSETTINGS

#Prints Output "string" in "Console"
def print_out(string):
    ActiveSheet = model.Sheets.getByIndex(0)    
    Console = ActiveSheet.getCellRangeByName("A33")
    OldLines = Console.String 
    Console.String = OldLines + string + "\n"
    return 1

#Check weather file exists
def check_def_case():
    ActiveSheet = model.Sheets.getByIndex(0)
    defTemplatekCasePath = ActiveSheet.getCellRangeByName("B2").String
    return os.path.isfile(defTemplatekCasePath)

#check weather desired working directory exists
def check_WD():
    ActiveSheet = model.Sheets.getByIndex(0)
    dir_path = ActiveSheet.getCellRangeByName("B3").String
    return os.path.isdir(dir_path)

#check weather setting inputs make sense
def check_ints():
    ActiveSheet = model.Sheets.getByIndex(0)
    Cell = ActiveSheet.getCellRangeByName("B9")
    ram = Cell.Value
    Cell = ActiveSheet.getCellRangeByName("B12")
    wall = Cell.Value
    Cell = ActiveSheet.getCellRangeByName("B15")
    nod = Cell.Value
    Cell = ActiveSheet.getCellRangeByName("B15")
    ppn = Cell.Value

    try:
        ram = int(ram)
        wall = int(wall)
        nod = int(nod)
        ppn = int(ppn)

    except ValueError:
        return False
    
    return True

#check if desired module input makes sense
def check_module():
    ActiveSheet = model.Sheets.getByIndex(0)
    mod = ActiveSheet.getCellRangeByName("B24").String
    mod_split = mod.split("/")
    if len(mod_split) != 2:
        return False
    if mod_split[0] != "ANSYS":
        return False
    ver_split = mod_split[1].split(".")
    if len(ver_split) != 2:
        return False
    
    valid_ords = [48,57]
    ord_list = []
    for item in range(valid_ords[0],valid_ords[1]+1,1):
        ord_list.append(item)
    for num in ver_split[0]:
        num = ord(num)
        if num not in ord_list:
            return False
    for num in ver_split[1]:
        num = ord(num)
        if num not in ord_list:
            return False
    return True

#check jobchainname
def check_chainname():
    ActiveSheet = model.Sheets.getByIndex(0)
    name = ActiveSheet.getCellRangeByName("B6").String
    if any([False if x != "_" and x != "-" else x.isalnum() for x in name]): #keine Sonderzeichen außer Unterstrich und Bindestrich erlaubt
        print_out("fehlerhafte Eingabe Chainname")
        return False
    return True

#check if emailinput makes sense
def check_mail():
    ActiveSheet = model.Sheets.getByIndex(0)
    mail = ActiveSheet.getCellRangeByName("B27").String
    
    if mail == "choose@yourmail.com":
        return False
    
    #Valid ASCII chars for an emailadress
    valid_ord_ascii = [[48,57],[65,90],[97,122]]
    also_valid_ord = [45,46,95]
    
    list_ords = []
    for blocks in valid_ord_ascii:
        ord_list = range(blocks[0],blocks[1]+1,1)
        for num in ord_list:
            list_ords.append(num)
    for num in also_valid_ord:
        list_ords.append(num)
    
    #Split -- >check for @
    check_split = mail.split("@")    
    if len(check_split) != 2:
        return False
    else:
        check_1 = list(check_split[0])
        check_2 = list(check_split[1])
        
        #check if valid chars
        for char in check_1:
            if ord(char) not in list_ords:
                return False
        for char in check_2:
            if ord(char) not in list_ords:
                return False
        
        #must contain another dot
        if len(check_split[1].split(".")) != 2:
            return False
        #ending must be 2-3 chars long
        if len(check_split[1].split(".")[1]) < 2 or len(check_split[1].split(".")[1]) > 3:
            return False
    return True


#sets bg-color of input-cells with bad values orange
def set_err_colors(case,WD,mail,ints,mod,shname):
    ActiveSheet = model.Sheets.getByIndex(0)
    ErrCol = 0xff9000      

    Cell = ActiveSheet.getCellRangeByName("B2")
    if case:
        Cell.CellBackColor=-1
    else:
        Cell.CellBackColor=ErrCol 

    Cell = ActiveSheet.getCellRangeByName("B3")
    if WD:
        Cell.CellBackColor=-1
    else:
        Cell.CellBackColor=ErrCol 

    Cell = ActiveSheet.getCellRangeByName("B27:B29")
    if mail:
        Cell.CellBackColor=-1
    else:
        Cell.CellBackColor=ErrCol 

    Cell = ActiveSheet.getCellRangeByName("B9:B20")
    if ints:
        Cell.CellBackColor=-1
    else:
        Cell.CellBackColor=ErrCol

    Cell = ActiveSheet.getCellRangeByName("B24:B27")
    if mod:
        Cell.CellBackColor=-1
    else:
        Cell.CellBackColor=ErrCol
    
    Cell = ActiveSheet.getCellRangeByName("B6:B8")
    if shname:
        Cell.CellBackColor=-1
    else:
        Cell.CellBackColor=ErrCol
        
#checks inputs        
def check_settings():
    case_bool = check_def_case()
    WD_bool = check_WD()
    mail_bool = check_mail()
    ints_bool = check_ints()
    modul_bool = check_module()
    shname_bool = check_chainname()

    if case_bool == True and WD_bool == True and mail_bool == True and ints_bool == True and modul_bool == True and shname_bool == True :
        set_err_colors(case_bool,WD_bool,mail_bool,ints_bool,modul_bool,shname_bool)
        print_out("JobChainGen is ready")
        return True
    else:
        print_out("Input Error")#hide Button
        set_err_colors(case_bool,WD_bool,mail_bool,ints_bool,modul_bool,shname_bool)
        return  False
    
###END CHECKSETTINGS

###CHECKERROROUTPUT
        

###ENDCHECKERROROUT
    
###

###
    
def make_Jobchain():

    if check_settings() == False:
        return False

    print_out("Starting JobChainGen Python-Script...")
    
    ActiveSheet = model.Sheets.getByIndex(2)
    
    CCLExpressionLineRawTXT = ActiveSheet.getCellRangeByName("B1").String
    CCLRawTXT = ActiveSheet.getCellRangeByName("B2").String
    SubmissionRawTXT = ActiveSheet.getCellRangeByName("B5").String
    SubmissionCFXStart = ActiveSheet.getCellRangeByName("B6").String
    SubmissionCFXInitial = ActiveSheet.getCellRangeByName("B7").String
    SubIndependendRAWTXT = ActiveSheet.getCellRangeByName("B4").String
    SubDependendRAWTXT = ActiveSheet.getCellRangeByName("B3").String
    
    ActiveSheet = model.Sheets.getByIndex(0)
    
    AnsysModule = ActiveSheet.getCellRangeByName("B24").String
    DEFCasePath = ActiveSheet.getCellRangeByName("B2").String
    JobChainName = ActiveSheet.getCellRangeByName("B6").String
    Partitions = ActiveSheet.getCellRangeByName("B21").String
    
    dir_path = ActiveSheet.getCellRangeByName("B3").String
    DEFCaseName = os.path.basename(DEFCasePath)
    #caseNameBaseString = DEFCaseName.split(".")[0]
    
    ram = ActiveSheet.getCellRangeByName("B9").String                         
    walltime = ActiveSheet.getCellRangeByName("B12").String                     
    nodes = ActiveSheet.getCellRangeByName("B15").String
    ppn = ActiveSheet.getCellRangeByName("B18").String
    mail = ActiveSheet.getCellRangeByName("B27").String
    
    ActiveSheet = model.Sheets.getByIndex(1)
    
    #Sheet as Cursor
    cursor = ActiveSheet.createCursor()
    cursor.gotoStartOfUsedArea(True)
    cursor.gotoEndOfUsedArea(True)
    BoundaryConditionsSheet = ActiveSheet.getCellRangeByName(cursor.AbsoluteName).getDataArray()
    
    os.chdir(dir_path)
    
    print_out("Reading Boundary Conditions...")
    Expressions = []
    Units = []
    BoundaryConditions = []
    
    for idx, line in enumerate(BoundaryConditionsSheet):
        BC_line = []
        for item in line:
            data = str(item)
            if idx == 0:
                Expressions.append(data)
            elif idx == 1:
                Units.append(data)
            else:
                BC_line.append(data.replace(",","."))
                
        if idx > 1:
            BoundaryConditions.append(BC_line)
    
    simnames = [JobChainName + "_" + i[2] for i in BoundaryConditions]
    sh_names = [i+".sh" for i in simnames]
    ccl_names = [i+".ccl" for i in simnames]
    def_names = [i+".def" for i in simnames]
    dependencies = [i[0] for i in BoundaryConditions]
    maxiterrations = [int(float(i[1])) for i in BoundaryConditions]
    #BoundaryConditions = [i for i in BoundaryConditions[-3:]

    Expressions = Expressions[3:] #Ersten Columns aussortieren
    Units = Units[3:]
    BoundaryConditions = [i[3:] for i in BoundaryConditions]

    print_out("creating ccl-files and submission scripts...")
    
    jobchainTXT = ""
    for i in range(len(BoundaryConditions)):
        SIMBC = BoundaryConditions[i]
        	
        BC_Lines = ""
        for y in range(len(Expressions)):
        		BC_Line_to_Write = CCLExpressionLineRawTXT
        		BC_Line_to_Write = BC_Line_to_Write.replace("<<EXPRESSIONNAME>>",Expressions[y])
        		BC_Line_to_Write = BC_Line_to_Write.replace("<<EXPRESSIONVALUE>>",SIMBC[y])
        		BC_Line_to_Write = BC_Line_to_Write.replace("<<EXPRESSIONUNIT>>",Units[y])
        		BC_Lines += BC_Line_to_Write + "\n"
        
        CCLtxt = CCLRawTXT
        CCLtxt = CCLtxt.replace("<<MINITER>>","10")
        CCLtxt = CCLtxt.replace("<<MAXITER>>",str(maxiterrations[i]))
        CCLtxt = CCLtxt.replace("<<MAXITER>>",str(maxiterrations[i]))
        CCLtxt = CCLtxt.replace("<<EXPRESSIONLINES>>",str(BC_Lines))
                
        with open(ccl_names[i],"w") as cclfobj:
            cclfobj.write(CCLtxt)
	
        QSUBTXT = SubmissionRawTXT.replace("PARTITIONS",Partitions)
        QSUBTXT = QSUBTXT.replace("SIMNAME",simnames[i])
        QSUBTXT = QSUBTXT.replace("NODES",nodes)
        QSUBTXT = QSUBTXT.replace("PPN",ppn)
        QSUBTXT = QSUBTXT.replace("WALL",walltime+":00:00")
        QSUBTXT = QSUBTXT.replace("MEM",ram)
        QSUBTXT = QSUBTXT.replace("MAIL",mail)
        QSUBTXT = QSUBTXT.replace("ANSYSMODULE",AnsysModule)
        QSUBTXT = QSUBTXT.replace("ORIGFILE",DEFCaseName)
        
        if dependencies[i] == "1.0":
            QSUBTXT = QSUBTXT.replace("CFXSTARTLINE",SubmissionCFXStart+SubmissionCFXInitial)
            QSUBTXT = QSUBTXT.replace("RES_NAME",simnames[i-1]+"_001.res")
            jobchainTXT += SubDependendRAWTXT.replace("BASHNAME",sh_names[i])
        elif dependencies[i] == "0.0":
            QSUBTXT = QSUBTXT.replace("CFXSTARTLINE",SubmissionCFXStart)
            jobchainTXT += SubIndependendRAWTXT.replace("BASHNAME",sh_names[i])
            
        QSUBTXT = QSUBTXT.replace("CCLNAME",ccl_names[i])
        QSUBTXT = QSUBTXT.replace("DEFNAME",def_names[i])
        
        with open(sh_names[i],"wb") as shfobj:
            shfobj.write(QSUBTXT.encode('ascii'))
    
    print_out("creating sh-JobChainScript")
    with open(JobChainName+".sh","wb") as shfobj:
            shfobj.write(jobchainTXT.encode('ascii'))
    
    print_out("JobChainGen finished")
    
    print_out("")
    print_out("Start the JobChain with sh " + JobChainName +".sh")

    return True
