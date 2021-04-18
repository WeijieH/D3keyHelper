; =================================================================
;                   暗黑3 “老沙”按键助手
; Designed by Oldsand
; 转载请注明原作者
; 
; 
; 查看最新更新：https://github.com/WeijieH/D3keyHelper
; 欢迎提交bug，PR
; =================================================================


#SingleInstance Force
#IfWinActive, ahk_class D3 Main Window Class
#NoEnv
#InstallKeybdHook
SetWorkingDir %A_ScriptDir%
SetBatchLines -1
CoordMode, Pixel, Client

VERSION:=20210417
TITLE:=Format("暗黑3技能连点器 v{:d}   by Oldsand", VERSION)
D3W:=3440
D3H:=1440

ReadCfgFile("d3oldsand.ini", tabs, hotkeys, actions, intervals, others, generals)
Gui -MaximizeBox -MinimizeBox +Owner
tabslen:= ObjCount(StrSplit(tabs, "`|"))
currentProfile:=1
vRunning:=False
vPausing:=False
profileKeybinding:={}
keysOnHold:={}
Gui Font, s11
Gui Add, Tab3, x5 y5 w790 h400 gSetTabFocus, %tabs%
Gui Font
Loop, parse, tabs, `|
{
    currentTab := A_Index
    yFirstLine:=85
    y:=yFirstLine
    Gui Tab, %currentTab%
    Loop, 6
    {
        ac:=actions[currentTab][A_Index]
        if A_Index <= 4
        {
            Gui Add, Hotkey, x45 y%y% w75 vskillset%currentTab%s%A_Index%hotkey, % hotkeys[currentTab][A_Index]
        }
        Else
        {
            Gui Add, Edit, x45 y%y% w75 vskillset%currentTab%s%A_Index%hotkey +Disabled, % hotkeys[currentTab][A_Index]
        }
        Gui Add, DropDownList, x200 y%y% w80 AltSubmit Choose%ac% vskillset%currentTab%s%A_Index%dropdown, 禁用||按住不放||连点||保持Buff
        Gui Add, Edit, vskillset%currentTab%s%A_Index%edit x360 y%y% w61 h21 Number
        Gui Add, Updown, vskillset%currentTab%s%A_Index%updown Range20-30000, % intervals[currentTab][A_Index]
        y+=50
    }
    Gui Add, Text, x490 y%yFirstLine%, 快速切换：
    yd:=yFirstLine-5
    pfmd:=others[currentTab].profilemethod
    Gui Add, DropDownList, x+5 y%yd% w90 AltSubmit Choose%pfmd% vskillset%currentTab%profilekeybindingdropdown gSetProfileKeybinding, 无||鼠标中键||滚轮向上||滚轮向下||侧键1||侧键2||键盘按键
    Gui Add, Hotkey, x+15 w85 vskillset%currentTab%profilekeybindinghkbox gSetProfileKeybinding +Disabled, % others[currentTab].profilehotkey
    Gui Add, Text, x490 y+25, 走位辅助：
    pfmv:=others[currentTab].movingmethod
    pflm:=others[currentTab].lazymode
    Gui Add, DropDownList, x+5 y120 w70 AltSubmit Choose%pfmv% vskillset%currentTab%movingdropdown gSetMovingHelper, 无||强制站立||强制走位
    Gui Add, Text, vskillset%currentTab%movingtext x+15 y125 +Disabled, 间隔：
    Gui Add, Edit, vskillset%currentTab%movingedit x+5 y120 w61 h21 Number +Disabled
    Gui Add, Updown, vskillset%currentTab%movingupdown Range0-3000 +Disabled, % others[currentTab].movinginterval
    Gui Add, Text, x490 y+25, 宏启动方式：
    Gui Add, DropDownList, x+5 y160 w90 AltSubmit Choose%pflm% vskillset%currentTab%profilestartmodedropdown, 懒人模式||仅按下时
    if (currentTab>1)
    {
        Gui Font, s20
        Gui Add, Text, center x490 y250 w270, 辅助功能见主设置
        Gui Font, s10
        Gui Add, Text, center x490 y330 w250, 提交bug，检查更新:
        Gui Add, Link, x500 y+5 w250, <a href="https://github.com/WeijieH/D3keyHelper">https://github.com/WeijieH/D3keyHelper</a>
        Gui Font
    }
    Else
    {
        smartpause:=generals.enablesmartpause
        enablegamblehelper:=generals.enablegamblehelper
        gambleHK:=generals.gamblehelperhk
        Gui Add, CheckBox, x490 y245 vextragambleckbox gSetGambleHelper Checked%enablegamblehelper%, 赌博助手:
        Gui Add, Hotkey, vextragamblehk x+5 y242 w50 gSetGambleHelper, % gambleHK
        Gui Add, Text, vextragambletext x+5 y245, 发送右键次数
        Gui Add, Edit, vextragambleedit x+5 y242 w60 Number
        Gui Add, Updown, vextragambleupdown Range2-30, % generals.gamblehelpertimes
        Gui Add, CheckBox, x490 y280 vextrasmartpause Checked%smartpause%, 智能暂停
        Gui Add, CheckBox, x+20 y280 vextramore1 +Disabled, Coming Soon
        Gui Add, CheckBox, x+20 y280 vextramore2 +Disabled, Coming Soon
        Gui Font, s12
        Gui Add, Text, center x490 y330 w250, 提交bug，检查更新:
        Gui Add, Link, y+5 w250, <a href="https://github.com/WeijieH/D3keyHelper">https://github.com/WeijieH/D3keyHelper</a>
        Gui Font
    }
    Gui Add, GroupBox, x20 y50 w125 h340, 技能
    Gui Add, GroupBox, x+30 y50 w130 h340, 策略
    Gui Add, GroupBox, x+30 y50 w110 h340, 执行间隔（毫秒）
    Gui Add, GroupBox, x+30 y50 w300 h150, 额外设置
    Gui Add, GroupBox, y+10 w300 h100, 辅助功能
}
Gui Tab

startRunHK:=generals.starthotkey
startmethod:=generals.startmethod
Gui Font, cRed s10
Gui Add, Text, x470 y8, 宏启动暂停快捷键：
Gui Font
Gui Add, DropDownList, x+5 y5 w90 AltSubmit Choose%startmethod% vStartRunDropdown gSetStartRun, 鼠标右键||鼠标中键||滚轮向上||滚轮向下||侧键1||侧键2||键盘按键
Gui Add, Hotkey, x+5 y5 w70 vStartRunHKinput gSetStartRun, %startRunHK%

Menu, Tray, NoStandard
Menu, Tray, Add, 设置
Menu, Tray, Add, 退出
Menu, Tray, Default, 设置
Menu, Tray, Click, 1
Menu, Tray, Tip, %TITLE%
Menu, Tray, Icon, , , 1

Gosub, SetStartRun
Gosub, SetProfileKeybinding
Gosub, SetMovingHelper
Gosub, SetGambleHelper
SetTimer, safeGuard, 300
Gui Show, w800 h410, %TITLE%
Return


; =================================== User Functions =====================================
ReadCfgFile(cfgFileName, ByRef tabs, ByRef hotkeys, ByRef actions, ByRef intervals, ByRef others, ByRef generals){
    global VERSION
    if FileExist(cfgFileName)
    {
        generals:={}
        IniRead, ver, %cfgFileName%, General, version
        if (VERSION != ver)
        {
            MsgBox, 配置文件版本不匹配，如有错误请删除配置文件并手动配置。
        }
        IniRead, enablegamblehelper, %cfgFileName%, General, enablegamblehelper
        IniRead, gamblehelperhk, %cfgFileName%, General, gamblehelperhk
        IniRead, gamblehelpertimes, %cfgFileName%, General, gamblehelpertimes
        IniRead, enablesmartpause, %cfgFileName%, General, enablesmartpause
        IniRead, startmethod, %cfgFileName%, General, startmethod
        IniRead, starthotkey, %cfgFileName%, General, starthotkey
        generals:={"enablegamblehelper":enablegamblehelper ,"gamblehelpertimes":gamblehelpertimes, "gamblehelperhk":gamblehelperhk, "startmethod":startmethod, "starthotkey":starthotkey, "enablesmartpause":enablesmartpause}

        IniRead, tabs, %cfgFileName%
        tabs:=StrReplace(StrReplace(tabs, "`n", "`|"), "General|", "")
        hotkeys:=[]
        actions:=[]
        intervals:=[]
        others:=[]
        Loop, parse, tabs, `|
        {
            cSection:=A_LoopField
            thk:=[]
            tac:=[]
            tiv:=[]
            tos:={}
            Loop, 6
            {
                IniRead, hk, %cfgFileName%, %cSection%, skill_%A_Index%
                IniRead, ac, %cfgFileName%, %cSection%, action_%A_Index%
                IniRead, iv, %cfgFileName%, %cSection%, interval_%A_Index%
                thk.Push(hk)
                tac.Push(ac)
                tiv.Push(iv)
            }
            hotkeys.Push(thk)
            actions.Push(tac)
            intervals.Push(tiv)
            IniRead, pfmd, %cfgFileName%, %cSection%, profilehkmethod
            IniRead, pfhk, %cfgFileName%, %cSection%, profilehkkey
            IniRead, pfmv, %cfgFileName%, %cSection%, movingmethod
            IniRead, pfmi, %cfgFileName%, %cSection%, movinginterval
            IniRead, pflm, %cfgFileName%, %cSection%, lazymode
            tos:={"profilemethod":pfmd, "profilehotkey":pfhk, "movingmethod":pfmv, "movinginterval":pfmi, "lazymode":pflm}
            others.Push(tos)
        }

    }
    Else
    {
        tabs=配置1|配置2|配置3|配置4
        hotkeys:=[]
        actions:=[]
        intervals:=[]
        others:=[]
        Loop, parse, tabs, `|
        {
            hotkeys.Push(["1","2","3","4","左键","右键"])
            actions.Push([1,1,1,1,1,1])
            intervals.Push([300,300,300,300,300,300])
            others.Push({"profilemethod":1, "profilehotkey":"", "movingmethod":1, "movinginterval":100, "lazymode":1})
        }
        generals:={"enablegamblehelper":1 ,"gamblehelpertimes":15, "gamblehelperhk":"F4", "startmethod":7, "starthotkey":"F2", "enablesmartpause":1}
    }
    Return
}

SaveCfgFile(cfgFileName, tabs, VERSION){
    FileDelete, %cfgFileName%

    GuiControlGet, extragambleckbox
    GuiControlGet, extragamblehk
    GuiControlGet, extragambleedit
    GuiControlGet, extrasmartpause
    
    IniWrite, %VERSION%, %cfgFileName%, General, version
    IniWrite, %extragambleckbox%, %cfgFileName%, General, enablegamblehelper
    IniWrite, %extragamblehk%, %cfgFileName%, General, gamblehelperhk
    IniWrite, %extragambleedit%, %cfgFileName%, General, gamblehelpertimes
    IniWrite, %extrasmartpause%, %cfgFileName%, General, enablesmartpause

    GuiControlGet, StartRunDropdown
    GuiControlGet, StartRunHKInput
    IniWrite, %StartRunDropdown%, %cfgFileName%, General, startmethod
    IniWrite, %StartRunHKInput%, %cfgFileName%, General, starthotkey

    Loop, parse, tabs, `|
    {
        cSection:=A_Index
        nSction:=A_LoopField
        Loop, 6
        {
            GuiControlGet, skillset%cSection%s%A_Index%hotkey
            GuiControlGet, skillset%cSection%s%A_Index%dropdown
            GuiControlGet, skillset%cSection%s%A_Index%updown
            hk:=skillset%cSection%s%A_Index%hotkey
            ac:=skillset%cSection%s%A_Index%dropdown
            iv:=skillset%cSection%s%A_Index%updown
            IniWrite, %hk%, %cfgFileName%, %nSction%, skill_%A_Index%
            IniWrite, %ac%, %cfgFileName%, %nSction%, action_%A_Index%
            IniWrite, %iv%, %cfgFileName%, %nSction%, interval_%A_Index%
        }
        GuiControlGet, skillset%cSection%profilekeybindingdropdown
        GuiControlGet, skillset%cSection%profilekeybindinghkbox
        pfhkdd:=skillset%cSection%profilekeybindingdropdown
        pfhkbx:=skillset%cSection%profilekeybindinghkbox
        IniWrite, %pfhkdd%, %cfgFileName%, %nSction%, profilehkmethod
        IniWrite, %pfhkbx%, %cfgFileName%, %nSction%, profilehkkey
        GuiControlGet, skillset%cSection%movingdropdown
        GuiControlGet, skillset%cSection%movingupdown
        movingmethod:=skillset%cSection%movingdropdown
        movinginterval:=skillset%cSection%movingupdown
        IniWrite, %movingmethod%, %cfgFileName%, %nSction%, movingmethod
        IniWrite, %movinginterval%, %cfgFileName%, %nSction%, movinginterval
        GuiControlGet, skillset%cSection%profilestartmodedropdown
        lazymode:=skillset%cSection%profilestartmodedropdown
        IniWrite, %lazymode%, %cfgFileName%, %nSction%, lazymode
    }
    Return
}

getSkillButtonPos(buttonID, ww, wh){
    if (ww<2000)
    {
        x1:=(ww/2+67.2*buttonID-391.53+2)*wh/1080
        x2:=(ww/2+67.2*buttonID-391.53+4)*wh/1080
    }
    Else
    {
        x1:=(ww/2+90.031*buttonID-523.26+2)*wh/1440
        x2:=(ww/2+90.031*buttonID-523.26+4)*wh/1440
    }
    y:=0.9222*wh-0.4304
    Return [Round(x1), Round(x2), Round(y)]
}

splitRGB(vthiscolor){
    vblue := (vthiscolor & 0xFF)
    vgreen := ((vthiscolor & 0xFF00) >> 8)
    vred := ((vthiscolor & 0xFF0000) >> 16)
    Return [vred, vgreen, vblue]
}


; =====================================Subroutines===================================

SetTabFocus:
    ControlFocus %A_GuiControl%
Return

SetProfileKeybinding:
    Gui, Submit, NoHide
    mouseKeyArray:=["", "MButton", "WheelUp", "WheelDown", "XButton1", "XButton2", ""]
    global profileKeybinding, tabslen
    Loop, %tabslen%
    {
        currentPage:=A_Index
        if (skillset%currentPage%profilekeybindingdropdown = 7)
        {
            GuiControl, Enable, skillset%currentPage%profilekeybindinghkbox
        }
        Else
        { 
            GuiControl, Disable, skillset%currentPage%profilekeybindinghkbox
        }
        Gui, Submit, NoHide

        for key, value in profileKeybinding.Clone()
        {
            if (value = currentPage)
            {
                Hotkey, ~%key%, SwitchProfile, Off
                Hotkey, ~+%key%, SwitchProfile, Off
                profileKeybinding.Pop(key)
            }
        }
        voption:=skillset%currentPage%profilekeybindingdropdown
        if voption in 2,3,4,5,6
        {
            ckey:=mouseKeyArray[voption]
            Hotkey, ~%ckey%, SwitchProfile, on
            Hotkey, ~+%ckey%, SwitchProfile, on
            profileKeybinding[ckey]:=currentPage
        }
        else if (voption = 7)
        {
            ckey:=skillset%currentPage%profilekeybindinghkbox
            if (ckey)
            {
                Hotkey, ~%ckey%, SwitchProfile, on
                Hotkey, ~+%ckey%, SwitchProfile, on
                profileKeybinding[ckey]:=currentPage
            } 
        }
    }
Return

SwitchProfile:
    global profileKeybinding, currentProfile, vRunning
    currentHK:=StrReplace(StrReplace(A_ThisHotkey, "+"), "~")
    if (currentProfile!=profileKeybinding[currentHK])
    {
        currentProfile:=profileKeybinding[currentHK]
        Gosub, StopMarco
    }
Return

SetStartRun:
    Gui, Submit, NoHide
    global startRunHK, tabslen
    startRunMouseKeyArray:=["RButton", "MButton", "WheelUp", "WheelDown", "XButton1", "XButton2", ""]
    if (StartRunDropdown = 7)
    {
        GuiControl, Enable, StartRunHKinput
        newstartRunHK=%StartRunHKinput%
        Loop, %tabslen%
        {
            GuiControl, Enable, skillset%A_Index%s6dropdown
            GuiControl, Enable, skillset%A_Index%s6updown
            GuiControl, Enable, skillset%A_Index%s6edit
        }
    }
    Else
    {
        if (StartRunDropdown = 1)
        {
            Loop, %tabslen%
            {
                GuiControl, choose, skillset%A_Index%s6dropdown, 1
                GuiControl, Disable, skillset%A_Index%s6dropdown
                GuiControl, Disable, skillset%A_Index%s6updown
                GuiControl, Disable, skillset%A_Index%s6edit
            }
        }
        Else
        {
            Loop, %tabslen%
            {
                GuiControl, Enable, skillset%A_Index%s6dropdown
                GuiControl, Enable, skillset%A_Index%s6updown
                GuiControl, Enable, skillset%A_Index%s6edit
            }
        }
        GuiControl, Disable, StartRunHKinput
        newstartRunHK:=startRunMouseKeyArray[StartRunDropdown]
    }
    Try
    {
        Hotkey, ~%startRunHK%, MainMacro, off
        Hotkey, ~+%startRunHK%, MainMacro, off
        Hotkey, ~%newstartRunHK%, MainMacro, on
        Hotkey, ~+%newstartRunHK%, MainMacro, on
        startRunHK = %newstartRunHK%
    }
Return

SetMovingHelper:
    Gui, Submit, NoHide
    Loop, %tabslen%{
        npage:=A_Index
        if (skillset%npage%movingdropdown = 3)
        {
            GuiControl, Enable, skillset%npage%movingtext
            GuiControl, Enable, skillset%npage%movingedit
            GuiControl, Enable, skillset%npage%movingupdown
        }
        Else
        { 
            GuiControl, Disable, skillset%npage%movingtext
            GuiControl, Disable, skillset%npage%movingedit
            GuiControl, Disable, skillset%npage%movingupdown
        }
    }
Return


MainMacro:
    global vRunning, startRunHK, currentProfile, D3W, D3W
    WinGetPos, , , D3W, D3H, A
    ; c:=getSkillButtonPos(1, D3W, D3H)
    GuiControlGet, skillset%currentProfile%profilestartmodedropdown
    lazy:=skillset%currentProfile%profilestartmodedropdown
    switch lazy
    {
        case 1:
             if not vRunning
            {
                Gosub, RunMarco
            }
            Else
            {
                Gosub, StopMarco
            } 
        case 2:
            Gosub, RunMarco
            KeyWait, %startRunHK%
            Gosub, StopMarco
    }
Return

RunMarco:
    global currentProfile, vRunning, vPausing, keysOnHold
    Gui, Submit, NoHide
    Loop, 6
    {
        GuiControlGet, skillset%currentProfile%s%A_Index%dropdown
        GuiControlGet, skillset%currentProfile%s%A_Index%hotkey
        Switch skillset%currentProfile%s%A_Index%dropdown
        {
        Case 2:
            k:=skillset%currentProfile%s%A_Index%hotkey
            Switch A_Index
            {
                case 5:
                    k:="LButton"
                case 6:
                    k:="RButton"
            }
            send {%k% Down}
            keysOnHold[k]:=1
        Case 3, 4:
            GuiControlGet, skillset%currentProfile%s%A_Index%updown
            SetTimer, spamSkillKey%A_Index%, % skillset%currentProfile%s%A_Index%updown
        Default:
            SetTimer, spamSkillKey%A_Index%, off
        }
    }
    GuiControlGet, skillset%currentProfile%movingdropdown
    Switch skillset%currentProfile%movingdropdown
    {
        case 2:
            send {LShift Down}
            keysOnHold["LShift"]:=1
        case 3:
            GuiControlGet, skillset%currentProfile%movingedit
            if (skillset%currentProfile%movingedit<20)
            {
                send {e Down}
                keysOnHold["e"]:=1
            }
            Else
            {
                SetTimer, forceMoving, % skillset%currentProfile%movingedit
            }
    }
    vRunning:=True 
    vPausing:=False
Return

StopMarco:
    global keysOnHold, vRunning, vPausing
    Loop, 6
    {
        SetTimer, spamSkillKey%A_Index%, off
    }
    SetTimer, forceMoving, off
    for key, value in keysOnHold.Clone(){
        if GetKeyState(key)
        {
            send {%key% up}
            keysOnHold.Pop(key)
        }
    }
    vRunning:=False
    vPausing:=False
Return

spamSkillKey1:
    global currentProfile, vPausing, D3W, D3H
    if vPausing
    {
        Return
    }
    GuiControlGet, skillset%currentProfile%s1hotkey
    GuiControlGet, skillset%currentProfile%s1dropdown
    k:=skillset%currentProfile%s1hotkey
    switch skillset%currentProfile%s1dropdown
    {
        case 3:
            send %k%
        case 4:
            magicXY:=getSkillButtonPos(1, D3W, D3H)
            PixelGetColor, cright, magicXY[2], magicXY[3], rgb
            PixelGetColor, cleft, magicXY[1], magicXY[3], rgb
            crgbl:=splitRGB(cleft)
            crgbr:=splitRGB(cright)
            If !(crgbl[2]>crgbl[1] and crgbl[1]>crgbl[3] and crgbr[2]>crgbr[1] and crgbr[1]>crgbr[3] and crgbr[3]>7)
            {
                send %k%
            }
    }
Return

spamSkillKey2:
    global currentProfile, vPausing, D3W, D3H
    if vPausing
    {
        Return
    }
    GuiControlGet, skillset%currentProfile%s2hotkey
    GuiControlGet, skillset%currentProfile%s2dropdown
    k:=skillset%currentProfile%s2hotkey
    switch skillset%currentProfile%s2dropdown
    {
        case 3:
            send %k%
        case 4:
            magicXY:=getSkillButtonPos(2, D3W, D3H)
            PixelGetColor, cright, magicXY[2], magicXY[3], rgb
            PixelGetColor, cleft, magicXY[1], magicXY[3], rgb
            crgbl:=splitRGB(cleft)
            crgbr:=splitRGB(cright)
            If !(crgbl[2]>crgbl[1] and crgbl[1]>crgbl[3] and crgbr[2]>crgbr[1] and crgbr[1]>crgbr[3] and crgbr[3]>7)
            {
                send %k%
            }
    }
Return

spamSkillKey3:
    global currentProfile, vPausing, D3W, D3H
    if vPausing
    {
        Return
    }
    GuiControlGet, skillset%currentProfile%s3hotkey
    GuiControlGet, skillset%currentProfile%s3dropdown
    k:=skillset%currentProfile%s3hotkey
    switch skillset%currentProfile%s3dropdown
    {
        case 3:
            send %k%
        case 4:
            magicXY:=getSkillButtonPos(3, D3W, D3H)
            PixelGetColor, cright, magicXY[2], magicXY[3], rgb
            PixelGetColor, cleft, magicXY[1], magicXY[3], rgb
            crgbl:=splitRGB(cleft)
            crgbr:=splitRGB(cright)
            If !(crgbl[2]>crgbl[1] and crgbl[1]>crgbl[3] and crgbr[2]>crgbr[1] and crgbr[1]>crgbr[3] and crgbr[3]>7)
            {
                send %k%
            }
    }
Return

spamSkillKey4:
    global currentProfile, vPausing, D3W, D3H
    if vPausing
    {
        Return
    }
    GuiControlGet, skillset%currentProfile%s4hotkey
    GuiControlGet, skillset%currentProfile%s4dropdown
    k:=skillset%currentProfile%s4hotkey
    switch skillset%currentProfile%s4dropdown
    {
        case 3:
            send %k%
        case 4:
            magicXY:=getSkillButtonPos(4, D3W, D3H)
            PixelGetColor, cright, magicXY[2], magicXY[3], rgb
            PixelGetColor, cleft, magicXY[1], magicXY[3], rgb
            crgbl:=splitRGB(cleft)
            crgbr:=splitRGB(cright)
            If !(crgbl[2]>crgbl[1] and crgbl[1]>crgbl[3] and crgbr[2]>crgbr[1] and crgbr[1]>crgbr[3] and crgbr[3]>7)
            {
                send %k%
            }
    }
Return

spamSkillKey5:
    global currentProfile, vPausing, D3W, D3H
    if vPausing
    {
        Return
    }
    GuiControlGet, skillset%currentProfile%s5dropdown
    switch skillset%currentProfile%s5dropdown
    {
        case 3:
            Click
        case 4:
            magicXY:=getSkillButtonPos(5, D3W, D3H)
            PixelGetColor, cright, magicXY[2], magicXY[3], rgb
            PixelGetColor, cleft, magicXY[1], magicXY[3], rgb
            crgbl:=splitRGB(cleft)
            crgbr:=splitRGB(cright)
            If !(crgbl[2]>crgbl[1] and crgbl[1]>crgbl[3] and crgbr[2]>crgbr[1] and crgbr[1]>crgbr[3] and crgbr[3]>7)
            {
                Click
            }
    }
Return

spamSkillKey6:
    global currentProfile, vPausing, D3W, D3H
    if vPausing
    {
        Return
    }
    GuiControlGet, skillset%currentProfile%s6hotkey
    switch skillset%currentProfile%s6dropdown
    {
        case 3:
            Click Right
        case 4:
            magicXY:=getSkillButtonPos(6, D3W, D3H)
            PixelGetColor, cright, magicXY[2], magicXY[3], rgb
            PixelGetColor, cleft, magicXY[1], magicXY[3], rgb
            crgbl:=splitRGB(cleft)
            crgbr:=splitRGB(cright)
            If !(crgbl[2]>crgbl[1] and crgbl[1]>crgbl[3] and crgbr[2]>crgbr[1] and crgbr[1]>crgbr[3] and crgbr[3]>7)
            {
                Click Right
            }
    }
Return

forceMoving:
    global vPausing
    if vPausing
    {
        Return
    }
    send e
Return

gambleHelper:
    GuiControlGet, extragambleedit
    Send {RButton %extragambleedit%}
Return

safeGuard:
    global vRunning, tabslen
    If !WinActive("ahk_class D3 Main Window Class")
    {
        Gosub, StopMarco
    }
    Else
    {
        if vRunning
        {
            Loop, %tabslen%
            {
                currentTab:=A_Index
                Loop, 4
                {
                    GuiControl, Disable, skillset%currentTab%s%A_Index%hotkey
                }
            }
        }
        Else
        {
            Loop, %tabslen%
            {
                currentTab:=A_Index
                Loop, 4
                {
                    GuiControl, Enable, skillset%currentTab%s%A_Index%hotkey
                }
            }
        }
    }   
Return

SetGambleHelper:
    global gambleHK
    Gui, Submit, NoHide
    GuiControlGet, extragambleckbox
    GuiControlGet, extragamblehk
    Try
    {
        Hotkey, ~%gambleHK%, gambleHelper, off
    }
    If extragambleckbox
    {
        GuiControl, Enable, extragamblehk
        GuiControl, Enable, extragambletext
        GuiControl, Enable, extragambleedit
        GuiControl, Enable, extragambleupdown
        Try
        {
            Hotkey, ~%extragamblehk%, gambleHelper, on
            gambleHK:=extragamblehk
        }
    }
    Else
    {
        GuiControl, Disable, extragamblehk
        GuiControl, Disable, extragambletext
        GuiControl, Disable, extragambleedit
        GuiControl, Disable, extragambleupdown
    }
Return

; ========================================= Hotkeys =======================================
~Enter::
~T::
~M::
~+Enter::
~+T::
~+M::
    GuiControlGet, extrasmartpause
    if extrasmartpause
    {
        Gosub, StopMarco
    }
Return

~Tab::
    global vPausing, keysOnHold
    GuiControlGet, extrasmartpause
    if extrasmartpause
    {
        vPausing:=!vPausing
        if vPausing
        {
            for key, value in keysOnHold
            {
                send {%key% up}
            }
        }
        Else
        {
            for key, value in keysOnHold
            {
                send {%key% Down}
            }
        }
    }
Return

; ===================================== System Functions ==================================
GuiEscape:
GuiClose:
    Gui, Submit
    SaveCfgFile("d3oldsand.ini", tabs, VERSION)
Return

设置:
    Gui, Show,, %TIELE%
Return

退出:
    SaveCfgFile("d3oldsand.ini", tabs, VERSION)
ExitApp
