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

VERSION:=210421
TITLE:=Format("暗黑3技能连点器 v1.1.{:d}   by Oldsand", VERSION)
MainWindowW:=850
MainWindowH:=500

currentProfile:=ReadCfgFile("d3oldsand.ini", tabs, hotkeys, actions, intervals, ivdelays, others, generals)
Gui -MaximizeBox -MinimizeBox +Owner
tabsarray:=StrSplit(tabs, "`|")
tabslen:= ObjCount(tabsarray)
vRunning:=False
vPausing:=False
profileKeybinding:={}
keysOnHold:={}
DblClickTime:=DllCall("GetDoubleClickTime", "UInt")

tabw:=MainWindowW-10
tabh:=MainWindowH-30
extraSettingGroupy:=310
helperSettingGroupx:=510
helperSettingGroupy:=40
extraSettingLine1y:=extraSettingGroupy+30
extraSettingLine2y:=extraSettingLine1y+30
extraSettingLine3y:=extraSettingLine2y+30
extraSettingLine4y:=extraSettingLine3y+30
extraSettingLine1yo:=extraSettingLine1y-3
extraSettingLine2yo:=extraSettingLine2y-3
extraSettingLine3yo:=extraSettingLine3y-3
extraSettingLine4yo:=extraSettingLine4y-3
helperSettingLinex:=helperSettingGroupx+20
helperSettingLine1y:=helperSettingGroupy+35
helperSettingLine2y:=helperSettingLine1y+35
helperSettingLine3y:=helperSettingLine2y+35
helperSettingLine4y:=helperSettingLine3y+35
helperSettingLine5y:=helperSettingLine4y+35
helperSettingLine1yo:=helperSettingLine1y-3
helperSettingLine2yo:=helperSettingLine2y-3
helperSettingLine3yo:=helperSettingLine3y-3
helperSettingLine4yo:=helperSettingLine4y-3
helperSettingLine5yo:=helperSettingLine5y-3

Gui Font, s11
Gui Add, Tab3, x5 y5 w%tabw% h%tabh% vActiveTab gSetTabFocus AltSubmit, %tabs%
Gui Font
Loop, parse, tabs, `|
{
    currentTab := A_Index
    yFirstLine:=90
    y:=yFirstLine
    Gui Tab, %currentTab%
    Gui Add, Hotkey, x0 y0 w0 w0
    skillLabels:=["技能一：", "技能二：", "技能三：", "技能四：", "左键技能：", "右键技能："]
    yl:=yFirstLine-30
    Gui Add, Text, x100 w60 y%yl% center, 快捷键
    Gui Add, Text, x+10 w80 y%yl% center, 策略
    Gui Add, Text, x+25 w100 y%yl% center, 执行间隔（毫秒）
    Gui Add, Text, x+10 w100 y%yl% center, 随机延迟（毫秒）
    Loop, 6
    {
        yl:=y+2
        Gui Add, Text, x40 w60 y%yl% center, % skillLabels[A_Index]
        ac:=actions[currentTab][A_Index]
        switch A_Index
        {
            case 1,2,3,4:
                Gui Add, Hotkey, x100 y%y% w60 vskillset%currentTab%s%A_Index%hotkey, % hotkeys[currentTab][A_Index]
            case 5:
                Gui Add, Edit, x100 y%y% w60 vskillset%currentTab%s%A_Index%hotkey +Disabled, LButton
            case 6:
                Gui Add, Edit, x100 y%y% w60 vskillset%currentTab%s%A_Index%hotkey +Disabled, RButton
        }
        Gui Add, DropDownList, x+10 y%y% w80 AltSubmit Choose%ac% gSetSkillsetDropdown vskillset%currentTab%s%A_Index%dropdown, 禁用||按住不放||连点||保持Buff
        Gui Add, Edit, vskillset%currentTab%s%A_Index%edit x+20 y%y% w100 Number
        Gui Add, Updown, vskillset%currentTab%s%A_Index%updown Range20-30000, % intervals[currentTab][A_Index]
        Gui Add, Edit, vskillset%currentTab%s%A_Index%delayedit x+25 y%y% w70 Number
        Gui Add, Updown, vskillset%currentTab%s%A_Index%delayupdown Range0-3000, % ivdelays[currentTab][A_Index]
        y+=35
    }
    Gui Add, Text, x40 y%extraSettingLine1y%, 快速切换至本配置：
    pfmd:=others[currentTab].profilemethod
    Gui Add, DropDownList, x+5 y%extraSettingLine1yo% w90 AltSubmit Choose%pfmd% vskillset%currentTab%profilekeybindingdropdown gSetProfileKeybinding, 无||鼠标中键||滚轮向上||滚轮向下||侧键1||侧键2||键盘按键
    Gui Add, Hotkey, x+15 w100 vskillset%currentTab%profilekeybindinghkbox gSetProfileKeybinding, % others[currentTab].profilehotkey
    
    Gui Add, Text, x40 y%extraSettingLine2y%, 走位辅助：
    pfmv:=others[currentTab].movingmethod
    pflm:=others[currentTab].lazymode
    Gui Add, DropDownList, x+5 y%extraSettingLine2yo% w85 AltSubmit Choose%pfmv% vskillset%currentTab%movingdropdown gSetMovingHelper, 无||强制站立||强制走位
    Gui Add, Text, vskillset%currentTab%movingtext x+10 y%extraSettingLine2y%, 移动间隔（毫秒）：
    Gui Add, Edit, vskillset%currentTab%movingedit x+5 y%extraSettingLine2yo% w60 Number
    Gui Add, Updown, vskillset%currentTab%movingupdown Range0-3000, % others[currentTab].movinginterval
    
    Gui Add, Text, x40 y%extraSettingLine3y%, 宏启动方式：
    Gui Add, DropDownList, x+5 y%extraSettingLine3yo% w90 AltSubmit Choose%pflm% vskillset%currentTab%profilestartmodedropdown, 懒人模式||仅按下时

    pfqp:=others[currentTab].enablequickpause
    pfqpm1:=others[currentTab].quickpausemethod1
    pfqpm2:=others[currentTab].quickpausemethod2
    Gui Add, Checkbox, x40 y%extraSettingLine4y% Checked%pfqp% vskillset%currentTab%clickpauseckbox gSetQuickPause, 快速暂停：
    Gui Add, DropDownList, x+0 y%extraSettingLine4yo% w50 AltSubmit Choose%pfqpm1% vskillset%currentTab%clickpausedropdown1 gSetQuickPause, 双击||单击
    Gui Add, DropDownList, x+5 y%extraSettingLine4yo% w100 AltSubmit Choose%pfqpm2% vskillset%currentTab%clickpausedropdown2 gSetQuickPause, 鼠标左键||鼠标右键||鼠标中键||侧键1||侧键2
    Gui Add, Text, x+5 y%extraSettingLine4y% vskillset%currentTab%clickpausetext1, 则暂停压键
    Gui Add, Edit, vskillset%currentTab%clickpauseedit x+5 y%extraSettingLine4yo% w60 Number
    Gui Add, Updown, vskillset%currentTab%clickpauseupdown Range500-5000, % others[currentTab].quickpausedelay
    Gui Add, Text, x+5 y%extraSettingLine4y% vskillset%currentTab%clickpausetext2, 毫秒
    
    if (currentTab>1)
    {
        Gui Font, s20
        Gui Add, Text, center x540 y240 w270, 辅助功能见主设置
        Gui Font
    }
    Else
    {
        oldsandhelperhk:=generals.oldsandhelperhk
        oldsandhelpermethod:=generals.oldsandhelpermethod
        smartpause:=generals.enablesmartpause
        enablegamblehelper:=generals.enablegamblehelper
        enablesalvagehelper:=generals.enablesalvagehelper
        salvagehelpermethod:=generals.salvagehelpermethod
        playsound:=generals.enablesoundplay
        Gui Font, cRed s10
        Gui Add, Text, x%helperSettingLinex% y%helperSettingLine1y%, 助手宏启动快捷键：
        Gui Font
        Gui Add, DropDownList, x+0 y%helperSettingLine1yo% w70 AltSubmit Choose%oldsandhelpermethod% vhelperKeybindingdropdown gSetHelperKeybinding, 无||鼠标中键||滚轮向上||滚轮向下||侧键1||侧键2||键盘按键
        Gui Add, Hotkey, x+5 w70 vhelperKeybindingHK gSetHelperKeybinding, %oldsandhelperhk%
        Gui Add, CheckBox, x%helperSettingLinex% y%helperSettingLine2y% vextragambleckbox gSetGambleHelper Checked%enablegamblehelper%, 血岩赌博助手：
        Gui Add, Text, vextragambletext x+5 y%helperSettingLine2y%, 发送右键次数
        Gui Add, Edit, vextragambleedit x+5 y%helperSettingLine2yo% w60 Number
        Gui Add, Updown, vextragambleupdown Range2-30, % generals.gamblehelpertimes
        Gui Add, CheckBox, x%helperSettingLinex% y%helperSettingLine3y% vextraSalvageHelperCkbox gSetSalvageHelper Checked%enablesalvagehelper%, 铁匠分解助手：
        Gui Add, DropDownList, x+5 y%helperSettingLine3yo% w150 AltSubmit vextraSalvageHelperDropdown Choose%salvagehelpermethod%, 快速分解||Coming Soon||Coming Soon
        Gui Add, CheckBox, x%helperSettingLinex% y%helperSettingLine4y% vextramore3 +Disabled, 魔盒重铸助手（Coming Soon）
        Gui Add, CheckBox, x%helperSettingLinex% y%helperSettingLine5y% vextramore4 +Disabled, 魔盒升级助手（Coming Soon）
        Gui Add, CheckBox, x%helperSettingLinex% y+40 vextraSoundonProfileSwitch Checked%playsound%, 使用快捷键切换配置成功时播放声音
        Gui Add, CheckBox, x%helperSettingLinex% y+20 vextrasmartpause Checked%smartpause%, 智能暂停
        Gui Add, CheckBox, y+20 vextramore2 +Disabled, Coming Soon
        
    }
    Gui Add, GroupBox, x20 y%helperSettingGroupy% w475 h260, 按键宏设置
    Gui Add, GroupBox, x%helperSettingGroupx% y%helperSettingGroupy% w320 h420, 辅助功能
    Gui Add, GroupBox, x20 y%extraSettingGroupy% w475 h150, 额外设置
}
Gui Tab
GuiControl , Choose, ActiveTab, % currentProfile

startRunHK:=generals.starthotkey
startmethod:=generals.startmethod
Gui Font, cRed s10
Gui Add, Text, x530 y8, 战斗宏启动快捷键：
Gui Font
Gui Add, DropDownList, x+5 y5 w90 AltSubmit Choose%startmethod% vStartRunDropdown gSetStartRun, 鼠标右键||鼠标中键||滚轮向上||滚轮向下||侧键1||侧键2||键盘按键
Gui Add, Hotkey, x+5 y5 w70 vStartRunHKinput gSetStartRun, %startRunHK%
skillsetText:=tabsarray[currentProfile]
ybottomtext:=MainWindowH-20
Gui Add, Text, x10 y%ybottomtext%, 当前激活配置:
Gui Font, cRed s10
Gui Add, Text, x+5 w350 vStatuesSkillsetText, % skillsetText
Gui Font
Gui Add, Link, x530 y%ybottomtext%, 提交bug，检查更新: <a href="https://github.com/WeijieH/D3keyHelper">https://github.com/WeijieH/D3keyHelper</a>

Menu, Tray, NoStandard
Menu, Tray, Add, 设置
Menu, Tray, Add, 退出
Menu, Tray, Default, 设置
Menu, Tray, Click, 1
Menu, Tray, Tip, %TITLE%
Menu, Tray, Icon, , , 1

Gosub, SetSkillsetDropdown
Gosub, SetStartRun
Gosub, SetProfileKeybinding
Gosub, SetMovingHelper
Gosub, SetGambleHelper
Gosub, SetSalvageHelper
Gosub, SetHelperKeybinding
Gosub, SetQuickPause
SetTimer, safeGuard, 300
Gui Show, w%MainWindowW% h%MainWindowH%, %TITLE%
Return


; =================================== User Functions =====================================
ReadCfgFile(cfgFileName, ByRef tabs, ByRef hotkeys, ByRef actions, ByRef intervals, ByRef ivdelays, ByRef others, ByRef generals){
    local
    global VERSION
    if FileExist(cfgFileName)
    {
        generals:={}
        IniRead, ver, %cfgFileName%, General, version
        if (VERSION != ver)
        {
            MsgBox, 配置文件版本不匹配，如有错误请删除配置文件并手动配置。
        }
        IniRead, currentProfile, %cfgFileName%, General, activatedprofile, 1
        IniRead, oldsandhelperhk, %cfgFileName%, General, oldsandhelperhk, F5
        IniRead, oldsandhelpermethod, %cfgFileName%, General, oldsandhelpermethod, 7
        IniRead, enablegamblehelper, %cfgFileName%, General, enablegamblehelper, 1
        IniRead, gamblehelpertimes, %cfgFileName%, General, gamblehelpertimes, 15
        IniRead, enablesalvagehelper, %cfgFileName%, General, enablesalvagehelper, 0
        IniRead, salvagehelpermethod, %cfgFileName%, General, salvagehelpermethod, 1
        IniRead, enablesmartpause, %cfgFileName%, General, enablesmartpause, 0
        IniRead, enablesoundplay, %cfgFileName%, General, enablesoundplay, 1
        IniRead, startmethod, %cfgFileName%, General, startmethod, 7
        IniRead, starthotkey, %cfgFileName%, General, starthotkey, F2
        generals:={"oldsandhelpermethod":oldsandhelpermethod, "oldsandhelperhk":oldsandhelperhk
        , "enablesalvagehelper":enablesalvagehelper, "salvagehelpermethod":salvagehelpermethod
        , "enablegamblehelper":enablegamblehelper, "gamblehelpertimes":gamblehelpertimes
        , "startmethod":startmethod, "starthotkey":starthotkey
        , "enablesmartpause":enablesmartpause, "enablesoundplay":enablesoundplay}

        IniRead, tabs, %cfgFileName%
        tabs:=StrReplace(StrReplace(tabs, "`n", "`|"), "General|", "")
        hotkeys:=[]
        actions:=[]
        intervals:=[]
        ivdelays:=[]
        others:=[]
        Loop, parse, tabs, `|
        {
            cSection:=A_LoopField
            thk:=[]
            tac:=[]
            tiv:=[]
            tdy:=[]
            tos:={}
            Loop, 6
            {
                IniRead, hk, %cfgFileName%, %cSection%, skill_%A_Index%
                IniRead, ac, %cfgFileName%, %cSection%, action_%A_Index%, 1
                IniRead, iv, %cfgFileName%, %cSection%, interval_%A_Index%, 300
                IniRead, dy, %cfgFileName%, %cSection%, delay_%A_Index%, 10
                thk.Push(hk)
                tac.Push(ac)
                tiv.Push(iv)
                tdy.Push(dy)
            }
            hotkeys.Push(thk)
            actions.Push(tac)
            intervals.Push(tiv)
            ivdelays.Push(tdy)
            IniRead, pfmd, %cfgFileName%, %cSection%, profilehkmethod, 1
            IniRead, pfhk, %cfgFileName%, %cSection%, profilehkkey
            IniRead, pfmv, %cfgFileName%, %cSection%, movingmethod, 1
            IniRead, pfmi, %cfgFileName%, %cSection%, movinginterval, 100
            IniRead, pflm, %cfgFileName%, %cSection%, lazymode, 1
            IniRead, pfqp, %cfgFileName%, %cSection%, enablequickpause, 0
            IniRead, pfqpm1, %cfgFileName%, %cSection%, quickpausemethod1, 1
            IniRead, pfqpm2, %cfgFileName%, %cSection%, quickpausemethod2, 1
            IniRead, pfqpdy, %cfgFileName%, %cSection%, quickpausedelay, 1500
            tos:={"profilemethod":pfmd, "profilehotkey":pfhk, "movingmethod":pfmv, "movinginterval":pfmi, "lazymode":pflm
            , "enablequickpause":pfqp, "quickpausemethod1":pfqpm1, "quickpausemethod2":pfqpm2, "quickpausedelay":pfqpdy}
            others.Push(tos)
        }

    }
    Else
    {
        tabs=配置1|配置2|配置3|配置4
        currentProfile:=1
        hotkeys:=[]
        actions:=[]
        intervals:=[]
        ivdelays:=[]
        others:=[]
        Loop, parse, tabs, `|
        {
            hotkeys.Push(["1","2","3","4"])
            actions.Push([1,1,1,1,1,1])
            intervals.Push([300,300,300,300,300,300])
            ivdelays.Push([10,10,10,10,10,10])
            others.Push({"profilemethod":1, "profilehotkey":"", "movingmethod":1, "movinginterval":100, "lazymode":1
            , "enablequickpause":0, "quickpausemethod1":1, "quickpausemethod2":1, "quickpausedelay":1500})
        }
        generals:={"enablegamblehelper":1 ,"gamblehelpertimes":15, "oldsandhelperhk":"F5"
        , "startmethod":7, "starthotkey":"F2", "enablesmartpause":1, "salvagehelpermethod":1
        , "oldsandhelpermethod":7, "enablesalvagehelper":0, "enablesoundplay":1}
    }
    Return currentProfile
}

SaveCfgFile(cfgFileName, tabs, currentProfile, VERSION){
    createOrTruncateFile(cfgFileName)

    GuiControlGet, extragambleckbox
    GuiControlGet, helperKeybindingdropdown
    GuiControlGet, helperKeybindingHK
    GuiControlGet, extragambleedit
    GuiControlGet, extrasmartpause
    GuiControlGet, extraSalvageHelperCkbox
    GuiControlGet, extraSalvageHelperDropdown
    GuiControlGet, extraSoundonProfileSwitch

    IniWrite, %VERSION%, %cfgFileName%, General, version
    IniWrite, %currentProfile%, %cfgFileName%, General, activatedprofile
    IniWrite, %extragambleckbox%, %cfgFileName%, General, enablegamblehelper
    IniWrite, %extragambleedit%, %cfgFileName%, General, gamblehelpertimes
    IniWrite, %extrasmartpause%, %cfgFileName%, General, enablesmartpause
    IniWrite, %extraSalvageHelperCkbox%, %cfgFileName%, General, enablesalvagehelper
    IniWrite, %extraSalvageHelperDropdown%, %cfgFileName%, General, salvagehelpermethod
    IniWrite, %extraSoundonProfileSwitch%, %cfgFileName%, General, enablesoundplay
    IniWrite, %helperKeybindingHK%, %cfgFileName%, General, oldsandhelperhk
    IniWrite, %helperKeybindingdropdown%, %cfgFileName%, General, oldsandhelpermethod
    

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
            GuiControlGet, skillset%cSection%s%A_Index%delayupdown
            IniWrite, % skillset%cSection%s%A_Index%dropdown, %cfgFileName%, %nSction%, action_%A_Index%
            IniWrite, % skillset%cSection%s%A_Index%updown, %cfgFileName%, %nSction%, interval_%A_Index%
            IniWrite, % skillset%cSection%s%A_Index%delayupdown, %cfgFileName%, %nSction%, delay_%A_Index%
            if (A_Index < 5)
            {
                IniWrite, % skillset%cSection%s%A_Index%hotkey, %cfgFileName%, %nSction%, skill_%A_Index%
            }
        }
        GuiControlGet, skillset%cSection%profilekeybindingdropdown
        GuiControlGet, skillset%cSection%profilekeybindinghkbox
        IniWrite, % skillset%cSection%profilekeybindingdropdown, %cfgFileName%, %nSction%, profilehkmethod
        IniWrite, % skillset%cSection%profilekeybindinghkbox, %cfgFileName%, %nSction%, profilehkkey
        GuiControlGet, skillset%cSection%movingdropdown
        GuiControlGet, skillset%cSection%movingupdown
        IniWrite, % skillset%cSection%movingdropdown, %cfgFileName%, %nSction%, movingmethod
        IniWrite, % skillset%cSection%movingupdown, %cfgFileName%, %nSction%, movinginterval
        GuiControlGet, skillset%cSection%profilestartmodedropdown
        IniWrite, % skillset%cSection%profilestartmodedropdown, %cfgFileName%, %nSction%, lazymode
        GuiControlGet, skillset%cSection%clickpauseckbox
        IniWrite, % skillset%cSection%clickpauseckbox, %cfgFileName%, %nSction%, enablequickpause
        GuiControlGet, skillset%cSection%clickpausedropdown1
        IniWrite, % skillset%cSection%clickpausedropdown1, %cfgFileName%, %nSction%, quickpausemethod1
        GuiControlGet, skillset%cSection%clickpausedropdown2
        IniWrite, % skillset%cSection%clickpausedropdown2, %cfgFileName%, %nSction%, quickpausemethod2
        GuiControlGet, skillset%cSection%clickpauseupdown
        IniWrite, % skillset%cSection%clickpauseupdown, %cfgFileName%, %nSction%, quickpausedelay
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

skillKey(currentProfile, nskill, D3W, D3H){
    local
    global vPausing
    GuiControlGet, skillset%currentProfile%s%nskill%hotkey
    GuiControlGet, skillset%currentProfile%s%nskill%dropdown
    GuiControlGet, skillset%currentProfile%s%nskill%delayupdown
    switch nskill
    {
        case 1,2,3,4:
            k:=skillset%currentProfile%s%nskill%hotkey
        case 5:
            k:="LButton"
        case 6:
            k:="RButton"
    }
    switch skillset%currentProfile%s%nskill%dropdown
    {
        case 3:
            if (skillset%currentProfile%s%nskill%delayupdown>1)
            {
                Random, delay, 1, skillset%currentProfile%s%nskill%delayupdown
                Sleep, delay
            }
            if !vPausing
            {
                send {%k%}
            }
        case 4:
            magicXY:=getSkillButtonPos(nskill, D3W, D3H)
            PixelGetColor, cright, magicXY[2], magicXY[3], rgb
            PixelGetColor, cleft, magicXY[1], magicXY[3], rgb
            crgbl:=splitRGB(cleft)
            crgbr:=splitRGB(cright)
            If (!vPausing and !(crgbl[2]>crgbl[1] and crgbl[1]>crgbl[3] and crgbr[2]>crgbr[1] and crgbr[1]>crgbr[3] and crgbr[3]>7))
            {
                switch nskill
                {
                    case 5:
                        if GetKeyState("LShift")
                        {
                            send {%k%}
                        }
                        Else
                        {
                            send {LShift down}{%k% down}
                            send {LShift up}{%k% up}
                        }
                    Default:
                        send {%k%}
                }
            }
    }
    Return
}

createOrTruncateFile(FileName){
    if (FileName = "")
    {
        return
    }
    file:=FileOpen(FileName, "w", "UTF-16")
    if !IsObject(file)
    {
        MsgBox 无法创建或写入文件："%FileName%"
        return
    }
    file.Write("; ===============================================`r`n")
    file.Write("; 欢迎来到“老沙”D3按键宏的配置文件。`r`n")
    file.Write("; 每个非General区块都对应一套按键配置，可以自由增删。`r`n")
    file.Write("; ===============================================`r`n")
    file.Close()
}

oldsandHelper(){
    WinGetPos, , , D3W, D3H, A
    MouseGetPos, xpos, ypos
    GuiControlGet, extragambleckbox
    GuiControlGet, extraSalvageHelperCkbox
    GuiControlGet, extraSalvageHelperDropdown
    if (xpos<D3W/2)
    {
        if extragambleckbox
        {
            Gosub, gambleHelper
        }
    }
    Else
    {
        if (extraSalvageHelperCkbox and extraSalvageHelperDropdown=1)
        {
            Gosub, SalvageHelper
        }
    }
    Return
}

clickPauseMarco(keysOnHold, pausetime, vRunning)
{
    if vRunning
    {
        for key, value in keysOnHold
        {
            if GetKeyState(key)
            {
                send {%key% up}
            }
        }
        SetTimer, clickResumeMarco, off
        SetTimer, clickResumeMarco, -%pausetime%
    }
    Return
}

clickResumeMarco()
{
    local
    global keysOnHold, vRunning
    for key, value in keysOnHold
    {
        if (vRunning and !GetKeyState(key))
        {
            send {%key% down}
        }
    }
    Return
}

; =====================================Subroutines===================================
spamSkillKey1:
spamSkillKey2:
spamSkillKey3:
spamSkillKey4:
spamSkillKey5:
spamSkillKey6:
    if !vPausing
    {
        nkey:=SubStr(A_ThisLabel, 0, 1)
        skillKey(currentProfile, nkey, D3W, D3H)
    }
Return

SetTabFocus:
    Gui, Submit, NoHide
    GuiControl, , StatuesSkillsetText, % tabsarray[ActiveTab]
    currentProfile:=ActiveTab
    Gosub, SetQuickPause
Return

SetQuickPause:
    Gui, Submit, NoHide
    GuiControlGet, skillset%currentProfile%clickpauseckbox
    GuiControlGet, skillset%currentProfile%clickpausedropdown2
    mousePauseKeyArray:=["LButton", "RButton", "MButton", "XButton1", "XButton2"]
    currentQuickPauseHK:=mousePauseKeyArray[skillset%currentProfile%clickpausedropdown2]
    if skillset%currentProfile%clickpauseckbox
    {
        GuiControl, Enable, skillset%currentProfile%clickpausedropdown1
        GuiControl, Enable, skillset%currentProfile%clickpausedropdown2
        GuiControl, Enable, skillset%currentProfile%clickpausetext1
        GuiControl, Enable, skillset%currentProfile%clickpauseedit
        GuiControl, Enable, skillset%currentProfile%clickpauseupdown
        GuiControl, Enable, skillset%currentProfile%clickpausetext2
        Try {
            Hotkey, ~%quickPauseHK%, quickPause, off
            Hotkey, ~+%quickPauseHK%, quickPause, off
        } 
        Hotkey, ~%currentQuickPauseHK%, quickPause, on
        Hotkey, ~+%currentQuickPauseHK%, quickPause, on
        quickPauseHK:=currentQuickPauseHK
    }
    Else
    {
        GuiControl, Disable, skillset%currentProfile%clickpausedropdown1
        GuiControl, Disable, skillset%currentProfile%clickpausedropdown2
        GuiControl, Disable, skillset%currentProfile%clickpausetext1
        GuiControl, Disable, skillset%currentProfile%clickpauseedit
        GuiControl, Disable, skillset%currentProfile%clickpauseupdown
        GuiControl, Disable, skillset%currentProfile%clickpausetext2
        Hotkey, ~%currentQuickPauseHK%, quickPause, off
        Hotkey, ~+%currentQuickPauseHK%, quickPause, off
        quickPauseHK:=""
    }
Return

SetHelperKeybinding:
    Gui, Submit, NoHide
    mouseKeyArray:=["", "MButton", "WheelUp", "WheelDown", "XButton1", "XButton2", ""]
    GuiControlGet, HelperKeybindingdropdown
    GuiControlGet, HelperKeybindingHK
    Try
    {
        Hotkey, ~%oldsandHelperHK%, oldsandHelper, off
    }
    switch HelperKeybindingdropdown
    {
        case 1:
            GuiControl, Disable, HelperKeybindingHK
            newoldsandHelperHK:=""
        case 2,3,4,5,6:
            GuiControl, Disable, HelperKeybindingHK
            newoldsandHelperHK:=mouseKeyArray[HelperKeybindingdropdown]
        case 7:
            GuiControl, Enable, HelperKeybindingHK
            newoldsandHelperHK:=HelperKeybindingHK
    }
    Try
    {
        Hotkey, ~%oldsandHelperHK%, oldsandHelper, off
        Hotkey, ~%newoldsandHelperHK%, oldsandHelper, on
        oldsandHelperHK:=newoldsandHelperHK
    }
Return

SetProfileKeybinding:
    Gui, Submit, NoHide
    mouseKeyArray:=["", "MButton", "WheelUp", "WheelDown", "XButton1", "XButton2", ""]
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
    currentHK:=StrReplace(StrReplace(A_ThisHotkey, "+"), "~")
    if (currentProfile!=profileKeybinding[currentHK])
    {
        currentProfile:=profileKeybinding[currentHK]
        GuiControl, , StatuesSkillsetText, % tabsarray[currentProfile]
        GuiControl , Choose, ActiveTab, % tabsarray[currentProfile]
        Gosub, StopMarco
        GuiControlGet, extraSoundonProfileSwitch
        if extraSoundonProfileSwitch
        {
            SoundBeep, 750, 250
        }
        Gosub, SetQuickPause
    }
Return

SetStartRun:
    Gui, Submit, NoHide
    startRunMouseKeyArray:=["RButton", "MButton", "WheelUp", "WheelDown", "XButton1", "XButton2", ""]
    if (StartRunDropdown = 7)
    {
        GuiControl, Enable, StartRunHKinput
        newstartRunHK=%StartRunHKinput%
        Loop, %tabslen%
        {
            GuiControl, Enable, skillset%A_Index%s6dropdown
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
                GuiControl, Disable, skillset%A_Index%s6delayedit
                GuiControl, Disable, skillset%A_Index%s6delayupdown
            }
        }
        Else
        {
            Loop, %tabslen%
            {
                GuiControl, Enable, skillset%A_Index%s6dropdown
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

SetSkillsetDropdown:
    Gui, Submit, NoHide
    Loop, %tabslen%{
        npage:=A_Index
        Loop, 6
        {
            switch skillset%npage%s%A_Index%dropdown
            {
                case 1,2:
                    GuiControl, Disable, skillset%npage%s%A_Index%edit
                    GuiControl, Disable, skillset%npage%s%A_Index%updown
                    GuiControl, Disable, skillset%npage%s%A_Index%delayedit
                    GuiControl, Disable, skillset%npage%s%A_Index%delayupdown
                case 3:
                    GuiControl, Enable, skillset%npage%s%A_Index%edit
                    GuiControl, Enable, skillset%npage%s%A_Index%updown
                    GuiControl, Enable, skillset%npage%s%A_Index%delayedit
                    GuiControl, Enable, skillset%npage%s%A_Index%delayupdown
                case 4:
                    GuiControl, Enable, skillset%npage%s%A_Index%edit
                    GuiControl, Enable, skillset%npage%s%A_Index%updown
                    GuiControl, Disable, skillset%npage%s%A_Index%delayedit
                    GuiControl, Disable, skillset%npage%s%A_Index%delayupdown
            }
        }
    }
Return

MainMacro:
    WinGetPos, , , D3W, D3H, A
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

quickPause:
    GuiControlGet, skillset%currentProfile%clickpausedropdown1
    GuiControlGet, skillset%currentProfile%clickpauseupdown
    switch skillset%currentProfile%clickpausedropdown1
    {
        case 1:
            If (A_PriorHotkey=A_ThisHotkey and A_TimeSincePriorHotkey < DblClickTime)
            {
                clickPauseMarco(keysOnHold, skillset%currentProfile%clickpauseupdown, vRunning)
            }
        case 2:
            clickPauseMarco(keysOnHold, skillset%currentProfile%clickpauseupdown, vRunning)
    }
Return

forceMoving:
    if !vPausing
    {
        send e
    }
Return

gambleHelper:
    GuiControlGet, extragambleedit
    Send {RButton %extragambleedit%}
Return

SalvageHelper:
    Click
    sleep 100
    send {enter}
Return

safeGuard:
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
    Gui, Submit, NoHide
    GuiControlGet, extragambleckbox
    If extragambleckbox
    {
        GuiControl, Enable, extragambletext
        GuiControl, Enable, extragambleedit
    }
    Else
    {
        GuiControl, Disable, extragambletext
        GuiControl, Disable, extragambleedit
    }
Return

SetSalvageHelper:
    Gui, Submit, NoHide
    GuiControlGet, extraSalvageHelperCkbox
    GuiControlGet, extraSalvageHelperHK
    If extraSalvageHelperCkbox
    {
        GuiControl, Enable, extraSalvageHelperDropdown
    }
    Else
    {
        GuiControl, Disable, extraSalvageHelperDropdown
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
    SaveCfgFile("d3oldsand.ini", tabs, currentProfile, VERSION)
Return

设置:
    Gui, Show,, %TIELE%
Return

退出:
    SaveCfgFile("d3oldsand.ini", tabs, currentProfile, VERSION)
ExitApp
