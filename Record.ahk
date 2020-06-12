#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;Force singleton - Only one instance of script
#SingleInstance force

;Set Globals
global Verison = 1.02
global CenterValue := 130
global MyDateTime, Cash, EBT, Credit, Taxable, Tax, MP, Reg
global TaxRate := 8.75 ;Default Sacramento County Tax in 2019
global FontSize := 12 ;Default Font size is 12
global Total := 0.00
global File := "Database\Data"
global SettingFile := "Database\Setting"
global DatabaseSource = 
global DatabaseBackup = 
global AutoLogLocation = 

;Build all Gui but show only the Main Gui at launch
CreateMainGui()
CreateSettingGui()
LoadSettingFile()
Gui, Main:Default
Gui, Main:Show, AutoSize 
GuiControl, Main:Text, Reg, 
GuiControl, Main:Focus, Reg

CreateMainGui(){
    Gui, Main:New,, Main
    Gui, Main:Font, s%FontSize%
    Gui, Add, DateTime, w330 vMyDateTime, LongDate
    Gui, Add, Text,x%CenterValue%, Cash:
    Gui, Add, Edit, vCash w100 gUpdateTotalAndMP Right, 0.00
    Gui, Add, Text,, EBT:
    Gui, Add, Edit, vEBT w100 gUpdateTotal Right, 0.00
    Gui, Add, Text,, Credit:
    Gui, Add, Edit, vCredit w100 gUpdateTotal Right, 0.00
    Gui, Add, Text,, Taxable:
    Gui, Add, Edit, vTaxable w100 gUpdateTax Right, 0.00
    Gui, Add, Text,, Tax:
    Gui, Add, Edit, vTax w100 ReadOnly Right, 0.00
    Gui, Add, Text,, Reg:
    Gui, Add, Edit, vReg w100 gUpdateMP Right, 0.00
    Gui, Add, Text,, Total: ;Total us calculated automatically
    Gui, Add, Edit, vTotal w100 ReadOnly Right, 0.00
    Gui, Add, Text,, MP: ;MP us calculated automatically
    Gui, Add, Edit, vMP w100 ReadOnly Right, 0.00
    Gui, Add, Button, gSaveFile w100 r3 xs Section, Save
    Gui, Add, Button, gLoadFile w100 r3 ys, Load
    Gui, Add, Button, gAutolog w100 r3 ys, AutoLog
    Gui, Add, Button, gOpenSetting w100 r3 xs Section, Setting
    Gui, Add, Button, gCopyFolder w100 r3 ys, BackUp Database
    Gui, Add, Button, gExitFunc w100 r3 ys, Quit
    return
}

CreateSettingGui(){
    Gui, Setting:New,, Settings
    Gui, Setting:Font, s%FontSize%
    Gui, Add, Text,, Tax Rate:
    Gui, Add, Edit, vTaxRate w100 Right, 8.75
    Gui, Add, Text,, CRE.NET Location: 
    Gui, Add, Edit, vDatabaseSource w200 Section, %DatabaseSource%
    Gui, Add, Button, gPromptSource ys, Set
    Gui, Add, Text, xs Section, CRE.NET Backup: 
    Gui, Add, Edit, vDatabaseBackup w200 xs Section, %DatabaseBackup%
    Gui, Add, Button, gPromptTarget ys, Set
    Gui, Add, Text, xs Section, AutoLog Location: 
    Gui, Add, Edit, vAutoLogLocation w200 xs Section, %AutoLogLocation%
    Gui, Add, Button, gPromptAutoLog ys, Set
    Gui, Add, Button, gSaveSettingButton w100 r3 xs Section, Save Settings
    Gui, Add, Button, gCloseSetting w100 r3 ys, Cancel
    Gui, Add, Text,xs, Verison:%Verison%
    return
}

;Change to float will cause 6 leading zeros
UpdateValueToDollar(value){
    value += 0.0 
    return SubStr(value,1, Max(StrLen(value)-4,1) )
}
;Debug Tool: Used for reading elements
^=::
    MouseGetPos, , , id, control
    WinGetTitle, title, ahk_id %id%
    WinGetClass, class, ahk_id %id%
    ToolTip, ahk_id %id%`nahk_class %class%`n%title%`nControl: %control%
return

;NumPad Enter will work like tab
NumpadEnter::
    if WinActive("Main")
        Send, {Tab}
    else
        Send, {NumpadEnter}
return

;Calculate Cash Updates both Total and MP
UpdateTotalAndMP(){
    UpdateTotal()
    UpdateMP()
}

;Calculate Total based on Cash, EBT, and Credit
UpdateTotal(){
    Gui, Submit, NoHide
    newTotal := Cash + EBT + Credit
    newTotal := UpdateValueToDollar(newTotal)
    GuiControl,, Total, %newTotal%
}

;Update MP based on Reg and Cash
UpdateMP(){
    Gui, Submit, NoHide
    newMP := Reg - Cash
    newMP := UpdateValueToDollar(newMP)
    GuiControl,, MP, %newMP%
}

;Calculate Reg based on MP and Cash
CalculateReg(){
    Gui, Submit, NoHide
    newReg := MP + Cash
    newReg := UpdateValueToDollar(newReg)
    GuiControl,, Reg, %newReg%
}

UpdateTax(){
    Gui, Submit, NoHide
    newTax := Round(Taxable * TaxRate * 0.01, 2)
    newTax := UpdateValueToDollar(newTax)
    GuiControl,, Tax, %newTax%
}

;Saves database based on locations set in settings
CopyFolder(){
    LoadSettingFile()
    FileAdd := SubStr(MyDateTime, 1, 8) ;Exclude Time
    MsgBox, 4, , Source Folder: "%DatabaseSource%"`nTarget Folder: "%DatabaseBackup%".`nConfirm?
    IfMsgBox, No
        return
    if DatabaseSource =
        return
    if DatabaseBackup =
        return
    SplitPath, DatabaseSource, DatabaseSourceName  ; Extract only the folder name from its full path.
    FileCopyDir, %DatabaseSource%, %DatabaseBackup%\%DatabaseSourceName%-%FileAdd% ;Tack on current Date
    if ErrorLevel
        MsgBox The folder could not be copied, perhaps because a folder of that name already exists in "%DatabaseBackup%".
        return
    MSGBOX, Backup Successful
}
;Source location in settings
PromptSource(){
    Gui, Setting:+Disabled
    Gui, Setting:-AlwaysOnTop
    FileSelectFolder, SourceFolder,, 3, Select Source Folder
    if SourceFolder !=
        GuiControl,, DatabaseSource, %SourceFolder%
    Gui, Setting:-Disabled
    Gui, Setting:+AlwaysOnTop
}
;Target Backup location in settings
PromptTarget(){
    Gui, Setting:+Disabled
    Gui, Setting:-AlwaysOnTop
    FileSelectFolder, TargetFolder,, 3, Select Target Folder
    if TargetFolder !=
        GuiControl,, DatabaseBackup, %TargetFolder%
    Gui, Setting:-Disabled
    Gui, Setting:+AlwaysOnTop
}

PromptAutoLog(){
    Gui, Setting:+Disabled
    Gui, Setting:-AlwaysOnTop
    FileSelectFile, TargetFile,, 3, Select Target File
    if TargetFile !=
        GuiControl,, AutoLogLocation, %TargetFile%
    Gui, Setting:-Disabled
    Gui, Setting:+AlwaysOnTop
}

;Save into a txt in an hashtable-like format
;So it can be easily parsed with other languages
SaveFile(){
    Gui, Submit, NoHide
    FileAdd := SubStr(MyDateTime, 1, 8) ;Exclude Time
    FileDelete, %File%%FileAdd%.txt
    if Total = 0
    {
        MsgBox, Entry Error. Save Failed.
        return
    }
    FileAppend, 
    (
{Date:%FileAdd%,
Cash:%Cash%,
EBT:%EBT%,
Credit:%Credit%,
Taxable:%Taxable%,
Tax:%Tax%,
Total:%Total%,
MP:%MP%}
    ), %File%%FileAdd%.txt
    if (ErrorLevel > 0){
        MSGBOX, Save Failed
        return
    }
    MSGBOX, Save Successful
    ExitFunc()
}
;Load the hashtable-like format
LoadFile(){
    Gui, Submit, NoHide
    FileAdd := SubStr(MyDateTime, 1, 8)
    FileRead, OutputVar, %File%%FileAdd%.txt
    if (ErrorLevel > 0){
        MSGBOX, Load Failed
        return
    }
    ;Have to carefully parse and match data into the correct spot
    Loop, parse, OutputVar, `{`}, %A_Space%%A_Tab%
    {
        Loop, parse, A_LoopField, `,`:
        {
            if(A_Index==2){
                MyDateTime := A_LoopField
                GuiControl,, MyDateTime, %A_LoopField%
            }
            if(A_Index==4){
                Cash := A_LoopField
                GuiControl,, Cash,%A_LoopField%
            }
            if(A_Index==6){
                EBT := A_LoopField
                GuiControl,, EBT,%A_LoopField%
            }
            if(A_Index==8){
                Credit := A_LoopField
                GuiControl,, Credit,%A_LoopField%
            }
            if(A_Index==10){
                Taxable := A_LoopField
                GuiControl,, Taxable,%A_LoopField%
            }
            if(A_Index==12){
                Tax := A_LoopField
                GuiControl,, Tax,%A_LoopField%
            }
            if(A_Index==14){
                Total := A_LoopField
                GuiControl,, Total,%A_LoopField%
            }
            if(A_Index==16){
                MP := A_LoopField
                GuiControl,, MP,%A_LoopField%
            }
            ;MsgBox, 4, , File number %A_Index% is %A_LoopField%.`n`nContinue?
            ;IfMsgBox, No, break
        }
    }
    CalculateReg() ;Reg is not saved so have to be calculated
    MSGBOX, Load Successful
    OutputVar = ;Free Memory
}

Autolog(){
    run, %AutoLogLocation%
    sleep, 10000
    send, user{enter}
    sleep, 5000
    send, password{enter}
    sleep, 2000
    ExitFunc()
    return
}

SaveSettingButton(){
    SaveSettingFile()
    CloseSetting()
}

OpenSetting(){
    Gui, Main:+Disabled
    Gui, Setting:+AlwaysOnTop
    Gui, Setting:Show, AutoSize
    LoadSettingFile()
    return
}

CloseSetting(){
    Gui, Main:-Disabled
    Gui, Setting:-AlwaysOnTop
    Gui, Setting:Hide
    return
}

LoadSettingFile(){
    Gui, Submit, NoHide
    FileRead, OutputVar, %SettingFile%.txt
    if (ErrorLevel > 0){
        if (ErrorLevel == 1){
            SaveSettingFile()
        }else{
            MSGBOX, Load Failed
            return
        }
    }
    Loop, parse, OutputVar, `{`}, %A_Space%%A_Tab%
    {
        Loop, parse, A_LoopField, `,`=
        {
            if(A_Index==2){
                TaxRate := A_LoopField
                GuiControl,, TaxRate, %A_LoopField%
            }
            if(A_Index==4){
                DatabaseSource := A_LoopField
                GuiControl,, DatabaseSource, %A_LoopField%
            }
            if(A_Index==6){
                DatabaseBackup := A_LoopField
                GuiControl,, DatabaseBackup, %A_LoopField%
            }
            if(A_Index==8){
                AutoLogLocation := A_LoopField
                GuiControl,, AutoLogLocation, %A_LoopField%
            }
        }
    }
}

SaveSettingFile(){
    Gui, Submit, NoHide
    FileDelete, %SettingFile%.txt
    FileAppend, 
    (
{TaxRate=%TaxRate%,
SourceFolder=%DatabaseSource%,
TargetFolder=%DatabaseBackup%,
AutoLogLocation=%AutoLogLocation%}
    ), %SettingFile%.txt
    if (ErrorLevel > 0){
        MSGBOX, Save Failed
        return
    }
}

;Exit function
ExitFunc(){
    msgbox, 4, Quit, Exit Record?
    ifmsgbox Yes
        Exitapp
    else
    return
}

;AUTO-LOG
^l::
Autolog()
return

SettingGuiClose:
CloseSetting()
return

^Q::
MainGuiClose:
ExitFunc()
return
