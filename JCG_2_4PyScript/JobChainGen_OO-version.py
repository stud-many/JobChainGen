#import numpy as np
import subprocess
import os 
#import sys
#import uno
import time

desktop = XSCRIPTCONTEXT.getDesktop()
model = desktop.getCurrentComponent()
#check whether there's already an opened document. Otherwise, create a new one
if not hasattr(model, "Sheets"):
    model = desktop.loadComponentFromURL("private:factory/scalc","_blank", 0, () )
    #get the XText interface


###CHECKSETTINGS

def print_out(string):
    ActiveSheet = model.Sheets.getByIndex(0)    
    Console = ActiveSheet.getCellRangeByName("A29")
    OldLines = Console.String 
    Console.String = OldLines + string + "\n"
    return 1
        
def check_cfx_exe():
    ActiveSheet = model.Sheets.getByIndex(0)
    cfx5pre_path_exe = os.path.join(ActiveSheet.getCellRangeByName("B1").String,'bin','cfx5pre.exe')
    return os.path.isfile(cfx5pre_path_exe)

def check_cfx_case():
    ActiveSheet = model.Sheets.getByIndex(0)
    cfxBlankCasePath = ActiveSheet.getCellRangeByName("B2").String
    return os.path.isfile(cfxBlankCasePath)

def check_WD():
    ActiveSheet = model.Sheets.getByIndex(0)
    dir_path = ActiveSheet.getCellRangeByName("B3").String
    return os.path.isdir(dir_path)

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

def check_modul():
    ActiveSheet = model.Sheets.getByIndex(0)
    mod = ActiveSheet.getCellRangeByName("B21").String
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

def check_sh():
    ActiveSheet = model.Sheets.getByIndex(0)
    sh = ActiveSheet.getCellRangeByName("B6").String
    if len(sh.split(".")) != 2:
        return False
    
    sh_ending = sh.split(".")[1] 
    if sh_ending != "sh":
        return False
    return True

def check_mail():
    ActiveSheet = model.Sheets.getByIndex(0)
    mail = ActiveSheet.getCellRangeByName("B24").String
    
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

def set_err_colors(cfx,case,WD,mail,ints,mod,shname):
    ActiveSheet = model.Sheets.getByIndex(0)
    ErrCol = 0xff9000
    
    Cell = ActiveSheet.getCellRangeByName("B1")
    if cfx:
        Cell.CellBackColor=-1
    else:
        Cell.CellBackColor=ErrCol        

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

    Cell = ActiveSheet.getCellRangeByName("B24:B26")
    if mail:
        Cell.CellBackColor=-1
    else:
        Cell.CellBackColor=ErrCol 

    Cell = ActiveSheet.getCellRangeByName("B9:B20")
    if ints:
        Cell.CellBackColor=-1
    else:
        Cell.CellBackColor=ErrCol

    Cell = ActiveSheet.getCellRangeByName("B21:B23")
    if mod:
        Cell.CellBackColor=-1
    else:
        Cell.CellBackColor=ErrCol
    
    Cell = ActiveSheet.getCellRangeByName("B6:B8")
    if shname:
        Cell.CellBackColor=-1
    else:
        Cell.CellBackColor=ErrCol
        
        
def check_settings():
    cfx_bool = check_cfx_exe()
    case_bool = check_cfx_case()
    WD_bool = check_WD()
    mail_bool = check_mail()
    ints_bool = check_ints()
    modul_bool = check_modul()
    shname_bool = check_sh()

    if cfx_bool == True and case_bool == True and WD_bool == True and mail_bool == True and ints_bool == True and modul_bool == True and shname_bool == True :
        set_err_colors(cfx_bool,case_bool,WD_bool,mail_bool,ints_bool,modul_bool,shname_bool)
        print_out("JobChainGen is ready")
        return True
    else:
        print_out("Input Error")#hide Button
        set_err_colors(cfx_bool,case_bool,WD_bool,mail_bool,ints_bool,modul_bool,shname_bool)
        return  False
    
###END CHECKSETTINGS

###CHECKERROROUTPUT
        #Wenn neues Errorlog vorhaden, dann wird diese als string zurückgegeben
def find_errlog():
    ActiveSheet = model.Sheets.getByIndex(0)
    WD = os.path.join(ActiveSheet.getCellRangeByName("B3").String)
    
    now = time.time()
    allowed = 20 # Datei darf X Sekunden alt sein, nicht älter
    files = [f for f in os.listdir(WD) if os.path.isfile(os.path.join(WD,f))]
    for f in files:
        if f.split(".")[1] == "log":
            if "cfxpre_engine_error" in f.split(".")[0]:
                datatime = os.path.getmtime(os.path.join(WD,f))
                if now < datatime + allowed:
                    with open(os.path.join(WD,f),"r") as fobj:
                        error = fobj.read()
                        print_out("Aktuelle CFX-PRE Fehlerausgabe gefunden.")
                        print_out(error)
                        return True
    return False    
###ENDCHECKERROROUT
    
###

def Pre_ExpressionsCheckString (expressions_list):
    ActiveSheet = model.Sheets.getByIndex(2)
    rawtxt = ActiveSheet.getCellRangeByName("B7").String
    rawobj = ActiveSheet.getCellRangeByName("B8").String
    print_out("Checking Expressions in CFX-Pre...")
    txt = rawtxt
    
    for idx, name in enumerate(expressions_list):
        if idx < len(expressions_list)-1:
            txt = txt.replace("OBJ",rawobj.replace("EXPRESSIONNAME",name)+"OBJ")
        else:   
            txt = txt.replace("OBJ",rawobj.replace("EXPRESSIONNAME",name))
    return txt

def Pre_ReadExpressionOut(expressions_list):
    with open("used_expressions.ccl",'r') as file:
        data = file.read()
        count = 0
        should = len(expressions_list)
        
        for name in expressions_list:
            if name in data:
                count += 1
        if count == should:
            return True
        else:
            return False

###
    
def make_Jobchain():

    if check_settings() == False:
        return False

    print_out("Starting JobChainGen Python-Script...")
    
    ActiveSheet = model.Sheets.getByIndex(2)
    
    SubmissionRawTXT = ActiveSheet.getCellRangeByName("B6").String
    PrePreambleRawTXT = ActiveSheet.getCellRangeByName("B2").String
    PreExpressionChangeRAWTXT = ActiveSheet.getCellRangeByName("B1").String
    PreWriteDefRAWTXT = ActiveSheet.getCellRangeByName("B3").String
    SubIndependendRAWTXT = ActiveSheet.getCellRangeByName("B5").String
    SubDependendRAWTXT = ActiveSheet.getCellRangeByName("B4").String
    
    ActiveSheet = model.Sheets.getByIndex(0)
    
    cfxPreScript = os.path.join(ActiveSheet.getCellRangeByName("B3").String,"CFX_Pre_JobChain.pre")
    AnsysModule = ActiveSheet.getCellRangeByName("B21").String
    cfx5pre_path_exe = os.path.join(ActiveSheet.getCellRangeByName("B1").String,'bin','cfx5pre.exe')
    cfxBlankCasePath = ActiveSheet.getCellRangeByName("B2").String
    JobChainSH = ActiveSheet.getCellRangeByName("B6").String
    
    dir_path = ActiveSheet.getCellRangeByName("B3").String
    cfxBlankCaseName = os.path.basename(cfxBlankCasePath)
    caseNameBaseString = cfxBlankCaseName.split(".")[0]
    
    ram = ActiveSheet.getCellRangeByName("B9").String                         
    walltime = ActiveSheet.getCellRangeByName("B12").String                     
    nodes = ActiveSheet.getCellRangeByName("B15").String
    ppn = ActiveSheet.getCellRangeByName("B18").String
    mail = ActiveSheet.getCellRangeByName("B24").String
    
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
    
    name_startfrom = []

    PrePreamble = PrePreambleRawTXT
    PrePreamble +="\n"
    PrePreamble = PrePreamble.replace("LOADCASE",os.path.join(cfxBlankCasePath))
        
    PrePreamble += Pre_ExpressionsCheckString(Expressions[1:])
    PrePreamble += "\n"
    
    PreExpressionChange = PreExpressionChangeRAWTXT
    PreExpressionChange += "\n"
    
    PreWriteDef = PreWriteDefRAWTXT
    PreWriteDef +="\n"

    f = open(cfxPreScript, "w")
    f.write(PrePreamble)

    JBChain = open(os.path.join(dir_path,JobChainSH) , "w")

    for idx,item in enumerate(BoundaryConditions):
        sim_name = caseNameBaseString+'_'+str(idx)
        sim_name = sim_name
        def_name = sim_name+".def"
        SubBash_name = sim_name+".sh"
    
        BashF = open(os.path.join(dir_path,SubBash_name), "wb")    
   
        if item[0] == '0.0':
            res_name = 0
        else:
            res_name = name_startfrom[idx-1][0][:-4]+"_001.res"

        name_startfrom.append([def_name,res_name])
        TXT = SubmissionRawTXT.replace("SIMNAME",sim_name)
        TXT = TXT.replace("MEM",ram)
        TXT = TXT.replace("NODES",nodes)
        TXT = TXT.replace("ANSYSMODULE",AnsysModule)
        TXT = TXT.replace("DEFNAME",def_name)
        TXT = TXT.replace("PPN",ppn)
        TXT = TXT.replace("WALL",walltime+":00:00")
        TXT = TXT.replace("MAIL",mail)
        
        if item[0] == '0.0':
            TXT = TXT[:-30]
        else:
            TXT = TXT.replace("RES_NAME",res_name)
    
        BashF.write(TXT.encode('ascii'))  
        BashF.close()
    
        if item[0] == '0.0':
            JBCTXT = SubIndependendRAWTXT
            JBCTXT = JBCTXT.replace("BASHNAME",SubBash_name)
        else:
            JBCTXT = SubDependendRAWTXT
            JBCTXT = JBCTXT.replace("BASHNAME",SubBash_name)

        JBChain.write(JBCTXT)
    
        for idx2,BC in enumerate(item):
            BCasNP = BC
            if idx2 !=0:
                exp_changeWrite = PreExpressionChange
                exp_changeWrite = exp_changeWrite.replace("EXPNAME",Expressions[idx2])
                exp_changeWrite = exp_changeWrite.replace("EXPVAL",str(BCasNP))
                exp_changeWrite = exp_changeWrite.replace("EXPUNIT",Units[idx2])

                f.write(exp_changeWrite)
        def_Write = PreWriteDef
        def_Write = def_Write.replace("SAVEDEF",os.path.join(dir_path,def_name))
        f.write(def_Write)
    f.close()
    JBChain.close()

    filename = cfxPreScript
    args = cfx5pre_path_exe + " -batch " + filename 
    subprocess.call(args, shell=False)
    
    if find_errlog() == False:# Prüfe auf CFX-PRE-Fehler
        print_out("Script ended successfull")
    if Pre_ReadExpressionOut(Expressions[1:]):
        print_out("used Expression names are okay")
    else:
        print_out("Error detected in Expressions. Please doublecheck for correct settings of the CFX-Casefile")
        print_out("Check also for correct spelling (CaseSensitive!)")
    return True
