; =================================================================
;                  暗黑3 “老沙”按键助手  (MIT License)
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
#UseHook
SetWorkingDir %A_ScriptDir%
SetBatchLines -1
CoordMode, Pixel, Client
CoordMode, Mouse, Client
Process, Priority, , High

VERSION:=210504
TITLE:=Format("暗黑3技能连点器 v1.2.{:d}   by Oldsand", VERSION)
MainWindowW:=850
MainWindowH:=500
tabw:=MainWindowW-347
tabh:=MainWindowH-30
helperSettingGroupx:=515

currentProfile:=ReadCfgFile("d3oldsand.ini", tabs, hotkeys, actions, intervals, ivdelays, others, generals)
SendMode, % generals.sendmode
tabsarray:=StrSplit(tabs, "`|")
tabslen:= ObjCount(tabsarray)
vRunning:=False
vPausing:=False
helperDelay:=100
mouseDelay:=2
helperRunning:=False
helperBreak:=False
helperNonEmpty:=[]
safezone:={}
Loop, Parse, % generals.safezone, CSV
{
    safezone[A_LoopField]:=1
}
profileKeybinding:={}
keysOnHold:={}
gameGamma:=(generals.gamegamma>=0.5 and generals.gamegamma<=1.5)? generals.gamegamma:1
buffpercent:=(generals.buffpercent>=0 and generals.buffpercent<=1)? generals.buffpercent:0.05
DblClickTime:=DllCall("GetDoubleClickTime", "UInt")

Gui -MaximizeBox -MinimizeBox +Owner +DPIScale +LastFound
Gui, Margin, 5, 5
Gui Font, s11
Gui Add, Tab3, xm ym w%tabw% h%tabh% vActiveTab gSetTabFocus AltSubmit, %tabs%
Gui Font
Loop, parse, tabs, `|
{
    currentTab := A_Index
    Gui Tab, %currentTab%
    Gui Add, Hotkey, x0 y0 w0 w0
    Gui Add, GroupBox, xm+10 ym+30 w480 h260 section, 按键宏设置
    skillLabels:=["技能一：", "技能二：", "技能三：", "技能四：", "左键技能：", "右键技能："]
    Gui Add, Text, xs+85 ys+20 w60 center section, 快捷键
    Gui Add, Text, x+10 w80 center, 策略
    Gui Add, Text, x+30 w100 center, 执行间隔（毫秒）
    Gui Add, Text, x+10 w100 center, 随机延迟（毫秒）
    Loop, 6
    {
        Gui Add, Text, xs-65 w60 yp+36 center, % skillLabels[A_Index]
        ac:=actions[currentTab][A_Index]
        switch A_Index
        {
            case 1,2,3,4:
                Gui Add, Hotkey, x+5 yp-2 w60 vskillset%currentTab%s%A_Index%hotkey, % hotkeys[currentTab][A_Index]
            case 5:
                Gui Add, Edit, x+5 yp-2 w60 vskillset%currentTab%s%A_Index%hotkey +Disabled, LButton
            case 6:
                Gui Add, Edit, x+5 yp-2 w60 vskillset%currentTab%s%A_Index%hotkey +Disabled, RButton
        }
        Gui Add, DropDownList, x+10 w85 AltSubmit Choose%ac% gSetSkillsetDropdown vskillset%currentTab%s%A_Index%dropdown, 禁用||按住不放||连点||保持Buff
        Gui Add, Edit, vskillset%currentTab%s%A_Index%edit x+20 w100 Number
        Gui Add, Updown, vskillset%currentTab%s%A_Index%updown Range20-30000, % intervals[currentTab][A_Index]
        Gui Add, Edit, vskillset%currentTab%s%A_Index%delayedit hwndskillset%currentTab%s%A_Index%delayeditID x+25 w70 Number
        Gui Add, Updown, vskillset%currentTab%s%A_Index%delayupdown Range0-3000, % ivdelays[currentTab][A_Index]
        AddToolTip(skillset%currentTab%s%A_Index%delayeditID, "这里填入随机延迟的最大值，设为0可以关闭随即延迟")
    }

    Gui Add, GroupBox, xm+10 yp+45 w480 h160 section, 额外设置
    Gui Add, Text, xs+20 ys+30, 快速切换至本配置：
    pfmd:=others[currentTab].profilemethod
    Gui Add, DropDownList, x+5 yp-2 w90 AltSubmit Choose%pfmd% vskillset%currentTab%profilekeybindingdropdown gSetProfileKeybinding, 无||鼠标中键||滚轮向上||滚轮向下||侧键1||侧键2||键盘按键
    Gui Add, Hotkey, x+15 w100 vskillset%currentTab%profilekeybindinghkbox gSetProfileKeybinding, % others[currentTab].profilehotkey
    
    Gui Add, Text, xs+20 yp+35, 走位辅助：
    pfmv:=others[currentTab].movingmethod
    pflm:=others[currentTab].lazymode
    Gui Add, DropDownList, x+5 yp-2 w130 AltSubmit Choose%pfmv% vskillset%currentTab%movingdropdown gSetMovingHelper, 无||强制站立||强制走位（按住不放）||强制走位（连点）
    Gui Add, Text, vskillset%currentTab%movingtext x+10 yp+2, 间隔（毫秒）：
    Gui Add, Edit, vskillset%currentTab%movingedit x+5 yp-2 w60 Number
    Gui Add, Updown, vskillset%currentTab%movingupdown Range20-3000, % others[currentTab].movinginterval
    
    pfusq:=others[currentTab].useskillqueue
    Gui Add, Text, xs+20 yp+35, 宏启动方式：
    Gui Add, DropDownList, x+5 yp-2 w90 AltSubmit Choose%pflm% vskillset%currentTab%profilestartmodedropdown, 懒人模式||仅按下时
    Gui Add, Checkbox, x+10 yp+2 Checked%pfusq% hwnduseskillqueueckbox%currentTab%ID vskillset%currentTab%useskillqueueckbox gSetSkillQueue, 使用单线程按键队列（毫秒）：
    AddToolTip(useskillqueueckbox%currentTab%ID, "开启后按键不会被立刻按下而是存储至一个按键队列中`n连点会使技能加入队列头部，保持buff会使技能加入队列尾部")
    Gui Add, Edit, vskillset%currentTab%useskillqueueedit hwnduseskillqueueedit%currentTab%ID x+0 yp-2 w50 Number
    Gui Add, Updown, vskillset%currentTab%useskillqueueupdown Range30-1000, % others[currentTab].useskillqueueinterval
    AddToolTip(useskillqueueedit%currentTab%ID, "按键队列中的连点按键会以此间隔一一发送至游戏窗口")

    pfqp:=others[currentTab].enablequickpause
    pfqpm1:=others[currentTab].quickpausemethod1
    pfqpm2:=others[currentTab].quickpausemethod2
    Gui Add, Checkbox, xs+20 yp+35 Checked%pfqp% vskillset%currentTab%clickpauseckbox gSetQuickPause, 快速暂停：
    Gui Add, DropDownList, x+0 yp-2 w50 AltSubmit Choose%pfqpm1% vskillset%currentTab%clickpausedropdown1 gSetQuickPause, 双击||单击
    Gui Add, DropDownList, x+5 yp w100 AltSubmit Choose%pfqpm2% vskillset%currentTab%clickpausedropdown2 gSetQuickPause, 鼠标左键||鼠标右键||鼠标中键||侧键1||侧键2
    Gui Add, Text, x+5 yp+2 vskillset%currentTab%clickpausetext1, 则暂停压键
    Gui Add, Edit, vskillset%currentTab%clickpauseedit x+5 yp-2 w60 Number
    Gui Add, Updown, vskillset%currentTab%clickpauseupdown Range500-5000, % others[currentTab].quickpausedelay
    Gui Add, Text, x+5 yp+2 vskillset%currentTab%clickpausetext2, 毫秒
}
Gui Tab
GuiControl , Choose, ActiveTab, % currentProfile

Gui Add, GroupBox, x%helperSettingGroupx% ym+30 w327 h440 section, 辅助功能
oldsandhelperhk:=generals.oldsandhelperhk
oldsandhelpermethod:=generals.oldsandhelpermethod
smartpause:=generals.enablesmartpause
enablegamblehelper:=generals.enablegamblehelper
enablesalvagehelper:=generals.enablesalvagehelper
salvagehelpermethod:=generals.salvagehelpermethod
playsound:=generals.enablesoundplay
usecustomstanding:=generals.customstanding
usecustommoving:=generals.custommoving
helperspeed:=generals.helperspeed
Gui Font, cRed s10
Gui Add, Text, xs+20 ys+30, 助手宏启动快捷键：
Gui Font
Gui Add, DropDownList, x+0 yp-2 w75 AltSubmit Choose%oldsandhelpermethod% vhelperKeybindingdropdown gSetHelperKeybinding, 无||鼠标中键||滚轮向上||滚轮向下||侧键1||侧键2||键盘按键
Gui Add, Hotkey, x+5 w70 vhelperKeybindingHK gSetHelperKeybinding, %oldsandhelperhk%

Gui Add, Text, xs+20 yp+40, 助手宏动画速度：
Gui Add, DropDownList, x+5 yp-2 w90 AltSubmit Choose%helperspeed% vhelperAnimationSpeedDropdown, 非常快||快速||中等||慢速
Gui Add, Text, x+20 yp+2 w80 hwndhelperSafeZoneTextID vhelperSafeZoneText gdummyFunction
AddToolTip(helperSafeZoneTextID, "修改配置文件中Generals区块下的safezone值来设置安全格")

Gui Add, CheckBox, xs+20 yp+35 vextragambleckbox gSetGambleHelper Checked%enablegamblehelper%, 血岩赌博助手：
Gui Add, Text, vextragambletext x+5 yp, 发送右键次数
Gui Add, Edit, vextragambleedit x+10 yp-2 w60 Number
Gui Add, Updown, vextragambleupdown Range2-30, % generals.gamblehelpertimes

Gui Add, CheckBox, xs+20 yp+37 hwndextraSalvageHelperCkboxID vextraSalvageHelperCkbox gSetSalvageHelper Checked%enablesalvagehelper%, 铁匠分解助手：
Gui Add, DropDownList, x+5 yp-3 w150 AltSubmit vextraSalvageHelperDropdown gSetSalvageHelper Choose%salvagehelpermethod%, 快速分解||一键分解||智能分解||智能分解（只留太古）
AddToolTip(extraSalvageHelperCkboxID, "快速分解：按下快捷键即等同于点击鼠标左键+回车`n一键分解：一键分解背包内所有非安全格的装备`n智能分解：同一键分解，但会跳过远古，太古`n智能分解（只留太古）：只保留太古装备")

Gui Add, CheckBox, xs+20 yp+37 vextramore3 +Disabled, 魔盒重铸助手（Coming Soon）
Gui Add, CheckBox, xs+20 yp+35 vextramore4 +Disabled, 魔盒升级助手（Coming Soon）

Gui Add, CheckBox, xs+20 yp+60 vextraSoundonProfileSwitch Checked%playsound%, 使用快捷键切换配置成功时播放声音
Gui Add, CheckBox, xs+20 yp+35 hwndextraSmartPauseID vextraSmartPause Checked%smartpause%, 智能暂停
AddToolTip(extraSmartPauseID, "开启后，游戏中按tab键可以暂停宏`n回车键，M键，T键会停止宏")
Gui Add, CheckBox, xs+20 yp+35 vextraCustomStanding gSetCustomStanding Checked%usecustomstanding%, 使用自定义强制站立按键：
Gui Add, Hotkey, x+5 yp-2 w70 vextraCustomStandingHK gSetCustomStanding, % generals.customstandinghk

Gui Add, CheckBox, xs+20 yp+35 vextraCustomMoving gSetCustomMoving Checked%usecustommoving%, 使用自定义强制移动按键：
Gui Add, Hotkey, x+5 yp-2 w70 Limit14 vextraCustomMovingHK gSetCustomMoving, % generals.custommovinghk
Gui Add, CheckBox, xs+20 yp+35 vextramore2 +Disabled, Coming Soon

startRunHK:=generals.starthotkey
startmethod:=generals.startmethod
Gui Font, cRed s10
Gui Add, Text, x530 ym+3, 战斗宏启动快捷键：
Gui Font
Gui Add, DropDownList, x+5 yp-3 w90 AltSubmit Choose%startmethod% vStartRunDropdown gSetStartRun, 鼠标右键||鼠标中键||滚轮向上||滚轮向下||侧键1||侧键2||键盘按键
Gui Add, Hotkey, x+5 yp w70 vStartRunHKinput gSetStartRun, %startRunHK%

ybottomtext:=MainWindowH-20
Gui Add, Text, x10 y%ybottomtext%, 当前激活配置:
Gui Font, cRed s11
Gui Add, Text, x+5 yp w350 vStatuesSkillsetText, % tabsarray[currentProfile]
Gui Add, Text, x465 yp hwndCurrentmodeTextID gdummyFunction, % A_SendMode
Gui Font
Gui Add, Text, x380 yp hwndSendmodeTextID gdummyFunction, 按键发送模式:
AddToolTip(SendmodeTextID, "修改配置文件General区块下的sendmode值来设置按键发送模式")
AddToolTip(CurrentmodeTextID, "Event：默认模式，最佳兼容性`nInput：推荐模式，最佳速度但在旧操作系统上可能无效")
Gui Add, Link, x520 yp, 提交bug，检查更新: <a href="https://github.com/WeijieH/D3keyHelper">https://github.com/WeijieH/D3keyHelper</a>

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
Gosub, SetHelperKeybinding
Gosub, SetQuickPause
SetGambleHelper()
SetSalvageHelper()
SetCustomStanding()
SetCustomMoving()
SetSkillQueue()
SetTimer, safeGuard, 300
Gui Show, w%MainWindowW% h%MainWindowH%, %TITLE%
Return


; =================================== User Functions =====================================
/*
读取配置文件，无配置文件则返回默认设置
参数：
    cfgFileName：文件名
    tabs：ByRef String，存储由竖线“|”分隔的配置名，用于初始化Tab控件
    hotkeys：ByRef Array，存储配置的技能快捷键
    actions：ByRef Array，存储配置的技能策略选择
    intervals：ByRef Array，存储配置的技能施放间隔
    ivdelays Array，存储配置的技能施放间隔延迟
    others：ByRef Array，存储额外配置
    generals：ByRef Array，存储一些通用配置
返回：
    上次退出时激活的配置编号，用于初始化Tab控件
*/
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
        IniRead, custommoving, %cfgFileName%, General, custommoving, 0
        IniRead, custommovinghk, %cfgFileName%, General, custommovinghk, e
        IniRead, customstanding, %cfgFileName%, General, customstanding, 0
        IniRead, customstandinghk, %cfgFileName%, General, customstandinghk, LShift
        IniRead, safezone, %cfgFileName%, General, safezone, "61,62,63"
        IniRead, helperspeed, %cfgFileName%, General, helperspeed, 3
        IniRead, gamegamma, %cfgFileName%, General, gamegamma, 1.000000
        IniRead, sendmode, %cfgFileName%, General, sendmode, "Event"
        IniRead, buffpercent, %cfgFileName%, General, buffpercent, 0.050000
        generals:={"oldsandhelpermethod":oldsandhelpermethod, "oldsandhelperhk":oldsandhelperhk
        , "enablesalvagehelper":enablesalvagehelper, "salvagehelpermethod":salvagehelpermethod
        , "enablegamblehelper":enablegamblehelper, "gamblehelpertimes":gamblehelpertimes
        , "startmethod":startmethod, "starthotkey":starthotkey
        , "enablesmartpause":enablesmartpause, "enablesoundplay":enablesoundplay
        , "custommoving":custommoving, "custommovinghk":custommovinghk, "customstanding":customstanding, "customstandinghk":customstandinghk
        , "safezone":safezone, "helperspeed":helperspeed, "gamegamma":gamegamma, "sendmode":sendmode, "buffpercent":buffpercent}

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
            IniRead, pfusq, %cfgFileName%, %cSection%, useskillqueue, 0
            IniRead, pfusqiv, %cfgFileName%, %cSection%, useskillqueueinterval, 200
            tos:={"profilemethod":pfmd, "profilehotkey":pfhk, "movingmethod":pfmv, "movinginterval":pfmi, "lazymode":pflm
            , "enablequickpause":pfqp, "quickpausemethod1":pfqpm1, "quickpausemethod2":pfqpm2, "quickpausedelay":pfqpdy
            , "useskillqueue":pfusq, "useskillqueueinterval":pfusqiv}
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
            , "enablequickpause":0, "quickpausemethod1":1, "quickpausemethod2":1, "quickpausedelay":1500
            , "useskillqueue":0, "useskillqueueinterval":200})
        }
        generals:={"enablegamblehelper":1 ,"gamblehelpertimes":15, "oldsandhelperhk":"F5"
        , "startmethod":7, "starthotkey":"F2", "enablesmartpause":1, "salvagehelpermethod":1
        , "oldsandhelpermethod":7, "enablesalvagehelper":0, "enablesoundplay":1
        , "custommoving":0, "custommovinghk":"e", "customstanding":0, "customstandinghk":"LShift"
        , "safezone":"61,62,63", "helperspeed":3, "gamegamma":1.000000, "sendmode":"Event"
        , "buffpercent":0.050000}
    }
    Return currentProfile
}

/*
保存配置文件
参数：
    cfgFileName：文件名
    tabs：String，由竖线“|”分隔的配置名
    currentProfile：int， 当前激活的配置页面编号
    safezone： Array，安全区域的配置int
    VERSION：int，版本
返回：
    无
*/
SaveCfgFile(cfgFileName, tabs, currentProfile, safezone, VERSION){
    createOrTruncateFile(cfgFileName)

    GuiControlGet, extragambleckbox
    GuiControlGet, helperKeybindingdropdown
    GuiControlGet, helperKeybindingHK
    GuiControlGet, extragambleedit
    GuiControlGet, extraSmartPause
    GuiControlGet, extraSalvageHelperCkbox
    GuiControlGet, extraSalvageHelperDropdown
    GuiControlGet, extraSoundonProfileSwitch
    GuiControlGet, extraCustomMoving
    GuiControlGet, extraCustomMovingHK
    GuiControlGet, extraCustomStanding
    GuiControlGet, extraCustomStandingHK
    GuiControlGet, helperAnimationSpeedDropdown

    IniWrite, %VERSION%, %cfgFileName%, General, version
    IniWrite, %currentProfile%, %cfgFileName%, General, activatedprofile
    IniWrite, %extragambleckbox%, %cfgFileName%, General, enablegamblehelper
    IniWrite, %extragambleedit%, %cfgFileName%, General, gamblehelpertimes
    IniWrite, %extraSmartPause%, %cfgFileName%, General, enablesmartpause
    IniWrite, %extraSalvageHelperCkbox%, %cfgFileName%, General, enablesalvagehelper
    IniWrite, %extraSalvageHelperDropdown%, %cfgFileName%, General, salvagehelpermethod
    IniWrite, %extraSoundonProfileSwitch%, %cfgFileName%, General, enablesoundplay
    IniWrite, %helperKeybindingHK%, %cfgFileName%, General, oldsandhelperhk
    IniWrite, %helperKeybindingdropdown%, %cfgFileName%, General, oldsandhelpermethod
    IniWrite, %extraCustomMoving%, %cfgFileName%, General, custommoving
    IniWrite, %extraCustomMovingHK%, %cfgFileName%, General, custommovinghk
    IniWrite, %extraCustomStanding%, %cfgFileName%, General, customstanding
    IniWrite, %extraCustomStandingHK%, %cfgFileName%, General, customstandinghk
    IniWrite, %helperAnimationSpeedDropdown%, %cfgFileName%, General, helperspeed
    safezone:=keyJoin(",", safezone)
    IniWrite, %safezone%, %cfgFileName%, General, safezone
    global gameGamma, buffpercent
    IniWrite, %gameGamma%, %cfgFileName%, General, gamegamma
    IniWrite, %A_SendMode%, %cfgFileName%, General, sendmode
    IniWrite, %buffpercent%, %cfgFileName%, General, buffpercent
    

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
        GuiControlGet, skillset%cSection%useskillqueueckbox
        IniWrite, % skillset%cSection%useskillqueueckbox, %cfgFileName%, %nSction%, useskillqueue
        GuiControlGet, skillset%cSection%useskillqueueedit
        IniWrite, % skillset%cSection%useskillqueueedit, %cfgFileName%, %nSction%, useskillqueueinterval
    }
    Return
}

/*
计算当前分辨率下技能buff条最左边像素的坐标
参数：
    D3W：int，窗口区域的宽度
    D3H：int，窗口区域的高度
    buttonID：int，按钮的ID，最左为1，最右（鼠标右键）为6
    percent： float，从左计算，取样点在Buff条上位置的百分比
返回：
    [x坐标，y坐标]
*/
getSkillButtonBuffPos(D3W, D3H, buttonID, percent){
    x:=[1288, 1377, 1465, 1554, 1647, 1734]
    w:=63
    y:=1328*D3H/1440
    Return [Round(D3W/2-(3440/2-x[buttonID]-percent*w)*D3H/1440), Round(y)]
}

/*
将16进制的颜色标签转化为RGB array。 FFFFFF -> [255, 255, 255]
当游戏gamma不为1时，会尝试进行gamma修正。
参数：
    vthiscolor：16进制的RGB颜色标签，PixelGetColor直出
返回：
    [R，G，B]
*/
splitRGB(vthiscolor){
    local
    global gameGamma
    vblue:=(vthiscolor & 0xFF)
    vgreen:=((vthiscolor & 0xFF00) >> 8)
    vred:=((vthiscolor & 0xFF0000) >> 16)
    if (Abs(gameGamma-1)>0.05)
    {
        vblue:=((vblue / 255) ** (1.75*gameGamma-0.75)) * 255
        vgreen:=((vgreen / 255) ** (1.9*gameGamma-0.9)) * 255
        vred:=((vred / 255) ** (1.9*gameGamma-0.9)) * 255
    }
    Return [vred, vgreen, vblue]
}

/*
负责发送技能按键
参数：
    currentProfile：int，当前激活的配置编号
    nskill: int, 技能按钮编号 1-6
    D3W：int，窗口区域的宽度
    D3H：int，窗口区域的高度
    forceStandingKey：强制站立按键
    useSkillQueue：Bool，是否使用技能列表
返回：
    无
*/
skillKey(currentProfile, nskill, D3W, D3H, forceStandingKey, useSkillQueue){
    local
    global vPausing, skillQueue, buffpercent
    GuiControlGet, skillset%currentProfile%s%nskill%hotkey
    GuiControlGet, skillset%currentProfile%s%nskill%dropdown
    GuiControlGet, skillset%currentProfile%s%nskill%delayupdown
    k:=skillset%currentProfile%s%nskill%hotkey
    switch skillset%currentProfile%s%nskill%dropdown
    {
        ; 连点
        case 3:
            if (skillset%currentProfile%s%nskill%delayupdown>1)
            {
                Random, delay, 1, skillset%currentProfile%s%nskill%delayupdown
                Sleep, delay
            }
            if !vPausing
            {
                if useSkillQueue
                {
                    ; 当技能列表大于100时什么都不做，防止占用过多内存
                    if (skillQueue.Count() < 100){
                        ; 按键加入技能列表头部
                        ; [k, 3] k是具体按键，3代表因为连点加入
                        skillQueue.InsertAt(1, [k, 3])
                    }
                }
                Else
                {
                    Send {Blind}{%k%}
                }
            }
        ; 保持buff
        case 4:
            ; 获得对应按键buff条最左侧坐标
            magicXY:=getSkillButtonBuffPos(D3W, D3H, nskill, buffpercent)
            PixelGetColor, cpixel, magicXY[1], magicXY[2], rgb
            crgb:=splitRGB(cpixel)
            ; 具体判断是否需要补buff
            If (!vPausing and crgb[1]+crgb[2]+crgb[3] < 220)
            {
                switch nskill
                {
                    case 5:
                        ; 判断按键是否是左键
                        if useSkillQueue
                        {
                            if (skillQueue.Count() < 100){
                                ; 4代表因为补buff加入
                                skillQueue.Push([k, 4])
                            }
                        }
                        Else
                        {
                            ; 判断是否需要强制站立再点击左键
                            if GetKeyState(forceStandingKey)
                            {
                                Send {Blind}{%k%}
                            }
                            Else
                            {
                                Send {Blind}{%forceStandingKey% down}{%k% down}
                                Send {Blind}{%k% up}{%forceStandingKey% up}
                            }
                        }
                    Default:
                        if useSkillQueue
                        {
                            if (skillQueue.Count() < 100){
                                skillQueue.Push([k, 4])
                            }
                        }
                        Else
                        {
                            Send {Blind}{%k%}
                        }
                }
            }
    }
    Return
}

/*
清空配置文件，并写入默认的文件头
参数：
    FileName：配置文件名
返回：
    无
*/
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

/*
负责开启助手宏
参数：
    无
返回：
    无
*/
oldsandHelper(){
    local
    global helperRunning, helperBreak, helperDelay, mouseDelay
    if helperRunning{
        ; 防止过快连按
        ; 宏在执行中再按可以打断
        helperBreak:=True
        helperRunning:=False
        Sleep, 200
        Return
    }
    helperRunning:=True
    helperBreak:=False
    ; 获得当前游戏分辨率
    VarSetCapacity(rect, 16)
    DllCall("GetClientRect", "ptr", WinExist("A"), "ptr", &rect)
    D3W:=NumGet(rect, 8, "int")
    D3H:=NumGet(rect, 12, "int")
    GuiControlGet, extragambleckbox
    GuiControlGet, extraSalvageHelperCkbox
    GuiControlGet, extraSalvageHelperDropdown
    GuiControlGet, helperAnimationSpeedDropdown
    MouseGetPos, xpos, ypos ; 当前鼠标位置，用于宏结束后返回
    ; 载入预设动画速度
    switch helperAnimationSpeedDropdown
    {
        case 1:
            mouseDelay:=0
            helperDelay:=50
        case 2:
            mouseDelay:=2
            helperDelay:=100
        case 3:
            mouseDelay:=5
            helperDelay:=150
        case 4:
            mouseDelay:=10
            helperDelay:=200
    }
    SetDefaultMouseSpeed, mouseDelay
    ; 当鼠标在左侧
    if (xpos<680*D3H/1440)
    {
        if (extragambleckbox and isGambleOpen(D3W, D3H))
        {
            SetTimer, gambleHelper, -1
            Return
        }
    }
    Else if(xpos>D3W-(3440-2740)*D3H/1440)
    {
        ; 当鼠标在右侧
        if (extraSalvageHelperCkbox and extraSalvageHelperDropdown=1)
        {
            ; 快速分解
            quickSalvageHelper(D3W, D3H, helperDelay)
            helperRunning:=False
            Return
        }
    }
    ; 一键分解
    if (extraSalvageHelperCkbox and extraSalvageHelperDropdown>1)
    {
        ; 判断分解页面是否打开
        r:=isSalvagePageOpen(D3W, D3H)
        switch r[1]
        {
            ; 铁匠页面打开且分解页面打开
            case 2:
                salvageIconXY:=getSalvageIconXY(D3W, D3H, "center")
                MouseMove, salvageIconXY[1][1], salvageIconXY[1][2]
                ; 判断拆解按钮是否已经按下
                if (r[2][3]<10 and r[2][1]+r[2][2]>400)
                {
                    if helperBreak
                    {
                        helperRunning:=False
                        Return
                    }
                    ; 分解按钮已经按下，右键取消然后重新获得颜色信息
                    Click, Right
                    Sleep, helperDelay
                    p:=getSalvageIconXY(D3W, D3H, "edge")
                    PixelGetColor, cpixel, p[2][1], p[2][2], rgb
                    r[3]:=splitRGB(cpixel)
                    PixelGetColor, cpixel, p[3][1], p[3][2], rgb
                    r[4]:=splitRGB(cpixel)
                    PixelGetColor, cpixel, p[4][1], p[4][2], rgb
                    r[5]:=splitRGB(cpixel)
                }
                if (r[5][1]>50)
                {
                    ; 一键分解黄
                    if helperBreak
                    {
                        helperRunning:=False
                        Return
                    }
                    MouseMove, salvageIconXY[4][1], salvageIconXY[4][2]
                    Click
                    Sleep, helperDelay
                    Send {Enter}
                }
                if (r[4][3]>70)
                {
                    ; 一键分解蓝
                    if helperBreak
                    {
                        helperRunning:=False
                        Return
                    }
                    MouseMove, salvageIconXY[3][1], salvageIconXY[3][2]
                    Click
                    Sleep, helperDelay
                    Send {Enter}
                }
                if (r[3][1]>65)
                {
                    ; 一键分解白/灰
                    if helperBreak
                    {
                        helperRunning:=False
                        Return
                    }
                    MouseMove, salvageIconXY[2][1], salvageIconXY[2][2]
                    Click
                    Sleep, helperDelay
                    Send {Enter}
                }
                ; 点击分解按钮
                MouseMove, salvageIconXY[1][1], salvageIconXY[1][2]
                Sleep, helperDelay*0.5
                Click
                if helperBreak
                {
                    helperRunning:=False
                    Return
                }
                Sleep, helperDelay*0.5
                ; 执行一键分解
                fn:=Func("oneButtonSalvageHelper").Bind(D3W, D3H, xpos, ypos)
                SetTimer, %fn%, -1
                Return
            case 1:
                ; 铁匠页面打开但是不在分解页面
                helperRunning:=False
                Return
            Default:
                ; 铁匠页面未打卡 
                helperRunning:=False
                Return
        }
    }
    Return
}

/*
负责一键赌博（连按右键）
参数：
    无
返回：
    无
*/
gambleHelper(){
    local
    global helperDelay, helperBreak, helperRunning
    GuiControlGet, extragambleedit
    Loop, %extragambleedit%
    {
        if helperBreak{
            Break
        }
        Send {RButton}
        sleep helperDelay*0.5
    }
    helperRunning:=False
    Return
}

/*
负责快速分解（左键+回车）
参数：
    D3W：int，窗口区域的宽度
    D3H：int，窗口区域的高度
    helperDelay：按键延迟
返回：
    无
*/
quickSalvageHelper(D3W, D3H, helperDelay){
    Click
    Sleep, helperDelay
    if isDialogBoXOnScreen(D3W, D3H){
        Send {Enter}
    }
    Return
}

/*
负责一键分解
参数：
    D3W：int，窗口区域的宽度
    D3H：int，窗口区域的高度
    xpos：之前鼠标x坐标
    ypos：之前鼠标y坐标
返回：
    无
*/
oneButtonSalvageHelper(D3W, D3H, xpos, ypos){
    local
    global helperBreak, helperRunning, helperDelay, safezone, helperNonEmpty, mouseDelay
    GuiControlGet, extraSalvageHelperDropdown
    helperNonEmpty:=[]
    helperSkip:={}
    i:=0    ; 格子编号
    q:=0    ; 当前格子装备品质，1：普通传奇，2：远古传奇，3：太古传奇
    SetDefaultMouseSpeed, mouseDelay
    ; 开启一单独线程查找空格子
    fn1:=Func("listNonEmptyInventorySpaceIDs").Bind(D3W, D3H)
    SetTimer, %fn1%, -1
    Loop
    {
        ; 如果找到了空格子
        if helperNonEmpty.Count()>0 {
            i:=helperNonEmpty.RemoveAt(1)
        }
        ; 如果找完了
        if (helperBreak or i<0) {
            helperRunning:=False
            Click, Right
            MouseMove, xpos, ypos
            Return
        }
        Else if (i>0 and !helperSkip.HasKey(i)) {
            ; 得到空格子坐标
            m:=getInventorySpaceXY(D3W, D3H, i)
            MouseMove, m[1], m[2]
            ; 智能分解判断
            if (extraSalvageHelperDropdown > 2)
            {
                Sleep, Min(helperDelay*2, 300)  ; 等待边框显示完毕
                ; 获取三个位于边框上的点颜色
                PixelGetColor, cpixel, Round(m[3]-1-10*D3H/1440), m[2], RGB
                c1:=splitRGB(cpixel)
                PixelGetColor, cpixel, Round(m[3]-10*D3H/1440), m[2], RGB
                c2:=splitRGB(cpixel)
                PixelGetColor, cpixel, Round(m[3]+1-10*D3H/1440), m[2], RGB
                c3:=splitRGB(cpixel)
                c:=[Max(c1[1],c2[1],c3[1]),Max(c1[2],c2[2],c3[2]),Max(c1[3],c2[3],c3[3])]
                if (c[1]>100 or c[3]<20) {
                    ; 装备是太古或者远古
                    q:=(c[2]<35) ? 4:3
                } else {
                    ; 装备是普通传奇
                    q:=2
                }
            }
            if (q>=extraSalvageHelperDropdown) {
                Continue
            }
            Click
            Sleep, helperDelay  ; 等待对话框显示完毕
            if isDialogBoXOnScreen(D3W, D3H)
            {
                Send {Enter}
                if (i<=50)
                {
                    ; 如果不是最后一行，判断下方格子是否变为空格
                    Sleep, Min(Round(helperDelay*3), 300) ; 等待装备消失动画显示完毕
                    newID:=i+10
                    if (isInventorySpaceEmpty(D3W, D3H, newID, [[0.65625,0.714285714], [0.375,0.365079365]])){
                        helperSkip[newID]:=1
                    }
                }
            }
            Continue
        }
        Sleep, Round(helperDelay*0.5)
    }
}

/*
枚举所有空格子
参数：
    D3W：int，窗口区域的宽度
    D3H：int，窗口区域的高度
返回：
    无
*/
listNonEmptyInventorySpaceIDs(D3W, D3H){
    local
    global safezone, helperNonEmpty
    e:=[[0.65625,0.714285714], [0.375,0.365079365]]
    Loop, 60
    {
        ; 跳过安全区域，将找到的空格子压入列表中
        if (!safezone.HasKey(A_Index) and !isInventorySpaceEmpty(D3W, D3H, A_Index, e)){
            helperNonEmpty.Push(A_Index)
        }
    }
    ; 如果找完了则压入-1
    helperNonEmpty.Push(-1)
}

/*
负责快速暂停
参数：
    keysOnHold：object，当前所有压下的按键
    pausetime: int, 暂停的时间
    vRunning：Bool，战斗宏是否在运行
返回：
    无
*/
clickPauseMarco(keysOnHold, pausetime, vRunning){
    if vRunning
    {
        for key, value in keysOnHold
        {
            if GetKeyState(key)
            {
                Send {%key% up}
            }
        }
        ; 自动恢复
        SetTimer, clickResumeMarco, off
        SetTimer, clickResumeMarco, -%pausetime%
    }
    Return
}

/*
负责快速暂停恢复
参数：
    无
返回：
    无
*/
clickResumeMarco(){
    local
    global keysOnHold, vRunning
    ; 重新压下所有压键
    for key, value in keysOnHold
    {
        if (vRunning and !GetKeyState(key))
        {
            Send {%key% down}
        }
    }
    Return
}

/*
设置自定义强制站立按键相关的控件动画
参数：
    无
返回：
    无
*/
SetCustomStanding(){
    GuiControlGet, extraCustomStanding
    if extraCustomStanding
    {
        GuiControl, Enable, extraCustomStandingHK
        GuiControlGet, extraCustomStandingHK
        if !extraCustomStandingHK
        {
            GuiControl,, extraCustomStandingHK, LShift
        }
    }
    Else
    {
        GuiControl, Disable, extraCustomStandingHK
    }
    Return
}

/*
设置自定义强制移动按键相关的控件动画
参数：
    无
返回：
    无
*/
SetCustomMoving(){
    GuiControlGet, extraCustomMoving
    if extraCustomMoving
    {
        GuiControl, Enable, extraCustomMovingHK
        GuiControlGet, extraCustomMovingHK
        if !extraCustomMovingHK
        {
            GuiControl,, extraCustomMovingHK, e
        }
    }
    Else
    {
        GuiControl, Disable, extraCustomMovingHK
    }
    Return
}

/*
设置赌博助手相关的控件动画
参数：
    无
返回：
    无
*/
SetGambleHelper(){
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
}

/*
设置分解助手相关的控件动画
参数：
    无
返回：
    无
*/
SetSalvageHelper(){
    local
    global safezone
    Gui, Submit, NoHide
    GuiControlGet, extraSalvageHelperCkbox
    GuiControlGet, extraSalvageHelperHK
    GuiControlGet, extraSalvageHelperDropdown
    If extraSalvageHelperCkbox
    {
        GuiControl, Enable, extraSalvageHelperDropdown
        switch extraSalvageHelperDropdown
        {
            case 1:
                GuiControl, hide, helperSafeZoneText
            case 2,3,4:
                ; 如果是一键分解，检查安全区域设置
                hasSafeZone:=False
                Loop, 60
                {
                    if safezone.HasKey(A_Index)
                    {
                        hasSafeZone:=True
                        Break
                    }
                }
                if hasSafeZone
                {
                    GuiControl, +c348017, helperSafeZoneText
                    GuiControl,, helperSafeZoneText, 安全格已设置
                }
                Else
                {
                    GuiControl, +cFF0000, helperSafeZoneText
                    GuiControl,, helperSafeZoneText, 安全格未设置
                }
                GuiControl, show, helperSafeZoneText
        }
    }
    Else
    {
        GuiControl, Disable, extraSalvageHelperDropdown
        GuiControl, Hide, helperSafeZoneText
    }
    Return
}

/*
设置技能队列相关的控件动画
参数：
    无
返回：
    无
*/
SetSkillQueue(){
    local
    global tabslen
    Loop, %tabslen%
    {
        GuiControlGet, skillset%A_Index%useskillqueueckbox
        if skillset%A_Index%useskillqueueckbox
        {
            GuiControl, Enable, skillset%A_Index%useskillqueueedit
            GuiControl, Enable, skillset%A_Index%useskillqueueupdown
        }
        Else
        {
            GuiControl, Disable, skillset%A_Index%useskillqueueedit
            GuiControl, Disable, skillset%A_Index%useskillqueueupdown
        }
    }
    Return
}

/*
设置发送技能队列按键
参数：
    inv：int，技能队列延迟
返回：
    无
*/
spamSkillQueue(inv){
    local
    global skillQueue, forceStandingKey, keysOnHold
    while (skillQueue.Count() > 0)
    {
        ; 取出排在第一的按键
        _k:=skillQueue.RemoveAt(1)
        k:=_k[1]
        switch _k[1]
        {
            case "LButton":
                switch _k[2]
                {
                    case 4:
                        ; 如果是左键保持buff
                        if GetKeyState(forceStandingKey)
                        {
                            Send {Blind}{%k%}
                        }
                        Else
                        {
                            Send {Blind}{%forceStandingKey% down}{%k% down}
                            Send {Blind}{%k% up}{%forceStandingKey% up}
                        }
                    Default:
                        Send {Blind}{%k%}
                }
            Default:
                ; 如果是连点，按键前后停止一段时间
                if (_k[2] = 3){
                    sleep Round(inv*0.5)
                }
                Send {Blind}{%k%}
                if (_k[2] = 3){
                    sleep Round(inv*0.5)
                    Break
                }
        }
    }
    Return
}

/*
判断屏幕上是否有对话框
参数：
    D3W：int，窗口区域的宽度
    D3H：int，窗口区域的高度
返回：
    bool
*/
isDialogBoXOnScreen(D3W, D3H){
    ; 2点取色判断
    point1:=[D3W/2-(3440/2-1655)*D3H/1440, 500*D3H/1440]
    point2:=[D3W/2+(3440/2-1800)*D3H/1440, 500*D3H/1440]
    PixelGetColor, cpixel, Round(point1[1]), Round(point1[2]), rgb
    c1:=splitRGB(cpixel)
    PixelGetColor, cpixel, Round(point2[1]), Round(point2[2]), rgb
    c2:=splitRGB(cpixel)
    if (c1[1]>c1[2] and c1[2]>c1[3] and c1[3]<5 and c1[2]<15 and c1[1]>25 and c2[1]>c2[2] and c2[2]>c2[3] and c2[3]<5 and c2[2]<15 and c2[1]>25)
    {
        Return True
    }
    Else
    {
        Return False
    }
}

/*
判断屏幕左侧或者右侧是否有关闭按钮（红X）
参数：
    D3W：int，窗口区域的宽度
    D3H：int，窗口区域的高度B
    position：string，“left”或者“right”
返回：
    Bool
*/
isRedXonScreen(D3W, D3H, position){
    _centerWhiteL:=[680, 24]
    _centerWhiteR:=[3417, 24]
    _XsizeInside:=29
    _XsizeOutside:=35
    switch position
    {
        case "left":
            centerPoint:=[Round(_centerWhiteL[1]*D3H/1440), Round(_centerWhiteL[2]*D3H/1440)]
            upPoint:=[Round(_centerWhiteL[1]*D3H/1440), Round((_centerWhiteL[2]-_XsizeInside/3)*D3H/1440)]
            leftPoint:=[Round((_centerWhiteL[1]-_XsizeOutside/2)*D3H/1440), Round(_centerWhiteL[2]*D3H/1440)]
        case "right":
            centerPoint:=[Round(D3W-((3440-_centerWhiteR[1])*D3H/1440)), Round(_centerWhiteR[2]*D3H/1440)]
            upPoint:=[Round(D3W-((3440-_centerWhiteR[1])*D3H/1440)), Round((_centerWhiteR[2]-_XsizeInside/3)*D3H/1440)]
            leftPoint:=[Round(D3W-((3440-_centerWhiteR[1]-_XsizeOutside/2)*D3H/1440)), Round(_centerWhiteR[2]*D3H/1440)]
    }
    ; 3点取色判断
    PixelGetColor, cpixel, centerPoint[1], centerPoint[2], rgb
    centerrgb:=splitRGB(cpixel)
    PixelGetColor, cpixel, upPoint[1], upPoint[2], rgb
    uprgb:=splitRGB(cpixel)
    PixelGetColor, cpixel, leftPoint[1], leftPoint[2], rgb
    leftrgb:=splitRGB(cpixel)
    if (centerrgb[1]+centerrgb[2]>370 and uprgb[3]<5 and uprgb[1]>40 and leftrgb[1]>leftrgb[2] and leftrgb[1]>leftrgb[3])
    {
        Return True
    }
    Else
    {
        Return False
    }
}

/*
获取背包格子的坐标
参数：
    D3W：int，窗口区域的宽度
    D3H：int，窗口区域的高度
    ID：int，格子的编号
返回：
    [格子中心x，格子中心y，格子左上角x，格子左上角y]
*/
getInventorySpaceXY(D3W, D3H, ID){
    _firstSpaceUL:=[2753, 747]
    _spaceSizeInnerW:=64
    _spaceSizeInnerH:=63
    _spaceSizeW:=67
    _spaceSizeH:=66
    targetColumn:=Mod(ID-1,10)
    targetRow:=Floor((ID-1)/10)
    Return [Round(D3W-((3440-_firstSpaceUL[1]-_spaceSizeW*targetColumn-_spaceSizeInnerW/2)*D3H/1440)), Round((_firstSpaceUL[2]+targetRow*_spaceSizeH+_spaceSizeInnerH/2)*D3H/1440)
    , Round(D3W-((3440-_firstSpaceUL[1]-_spaceSizeW*targetColumn)*D3H/1440)), Round((_firstSpaceUL[2]+targetRow*_spaceSizeH)*D3H/1440)]
}

/*
判断铁匠/分解页面是否开启
参数：
    D3W：int，窗口区域的宽度
    D3H：int，窗口区域的高度
返回：
    [0]：如果没有开启
    [1]: 如果铁匠页面开启但拆解页面没开启
    [2, 大拆解按钮边缘坐标rgb, 白色解按钮边缘坐标rgb, 蓝色解按钮边缘坐标rgb, 黄色解按钮边缘坐标rgb]：如果铁匠开启且同时在拆解页面
*/
isSalvagePageOpen(D3W, D3H){
    point1:=[Round(321*D3H/1440),Round(86*D3H/1440)]
    point2:=[Round(351*D3H/1440),Round(107*D3H/1440)]
    point3:=[Round(388*D3H/1440),Round(86*D3H/1440)]
    point4:=[Round(673*D3H/1440),Round(1040*D3H/1440)]
    PixelGetColor, cpixel, point1[1], point1[2], rgb
    c1:=splitRGB(cpixel)
    PixelGetColor, cpixel, point2[1], point2[2], rgb
    c2:=splitRGB(cpixel)
    PixelGetColor, cpixel, point3[1], point3[2], rgb
    c3:=splitRGB(cpixel)
    PixelGetColor, cpixel, point4[1], point4[2], rgb
    c4:=splitRGB(cpixel)
    if (c1[3]>c1[2] and c1[2]>c1[1] and c1[3]>110 and c3[3]>c3[2] and c3[2]>c3[1] and c3[3]>110 and c2[1]+c2[2]>350 and c4[1]>50 and c4[2]<15 and c4[3]<15 and isRedXonScreen(D3W, D3H, "left")){
        p:=getSalvageIconXY(D3W, D3H, "edge")
        PixelGetColor, cpixel, p[1][1], p[1][2], rgb
        cLeg:=splitRGB(cpixel)
        PixelGetColor, cpixel, p[2][1], p[2][2], rgb
        cWhite:=splitRGB(cpixel)
        PixelGetColor, cpixel, p[3][1], p[3][2], rgb
        cBlue:=splitRGB(cpixel)
        PixelGetColor, cpixel, p[4][1], p[4][2], rgb
        cRare:=splitRGB(cpixel)
        if (cBlue[3]>cBlue[2] and cBlue[2]>cBlue[1] and cRare[3]<20 and cRare[1]>cRare[2] and cRare[2]>cRare[3]) {
            Return [2, cLeg, cWhite, cBlue, cRare]
        } Else {
            Return [1]
        }
    }
    Else{
        Return [0]
    }
}

/*
获得拆解页面4个按钮的坐标
参数：
    D3W：int，窗口区域的宽度
    D3H：int，窗口区域的高度
    c: string，“center”-中心坐标，“edge”-边缘颜色带内坐标
返回：
    [大拆解按钮坐标xy，白色解按钮坐标xy，蓝色解按钮坐标xy，黄色解按钮坐标xy]
*/
getSalvageIconXY(D3W, D3H, c){
    switch c
    {
        case "center":
            centerLeg:=[Round(221*D3H/1440),Round(388*D3H/1440)]
            centerWhite:=[Round(335*D3H/1440),Round(388*D3H/1440)]
            centerBlue:=[Round(424*D3H/1440),Round(388*D3H/1440)]
            centerRare:=[Round(514*D3H/1440),Round(388*D3H/1440)]
            Return [centerLeg, centerWhite, centerBlue, centerRare]
        case "edge":
            edgeColorLeg:=[Round(203*D3H/1440),Round(337*D3H/1440)]
            edgeColorWhite:=[Round(335*D3H/1440),Round(371*D3H/1440)]
            edgeColorBlue:=[Round(424*D3H/1440),Round(371*D3H/1440)]
            edgeColorRare:=[Round(514*D3H/1440),Round(371*D3H/1440)]
            Return [edgeColorLeg, edgeColorWhite, edgeColorBlue, edgeColorRare]
    }
}

/*
判断赌博页面是否开启
参数：
    D3W：int，窗口区域的宽度
    D3H：int，窗口区域的高度
返回：
    Bool
*/
isGambleOpen(D3W, D3H){
    point1:=[Round(320*D3H/1440),Round(96*D3H/1440)]
    point2:=[Round(351*D3H/1440),Round(100*D3H/1440)]
    point4:=[Round(194*D3H/1440),Round(67*D3H/1440)]
    point5:=[Round(147*D3H/1440),Round(94*D3H/1440)]
    PixelGetColor, cpixel, point1[1], point1[2], rgb
    c1:=splitRGB(cpixel)
    PixelGetColor, cpixel, point2[1], point2[2], rgb
    c2:=splitRGB(cpixel)
    PixelGetColor, cpixel, point4[1], point4[2], rgb
    c4:=splitRGB(cpixel)
    PixelGetColor, cpixel, point5[1], point5[2], rgb
    c5:=splitRGB(cpixel)
    if (c1[3]>c1[1] and c1[1]>c1[2] and c1[3]>130 and c2[1]+c2[2]>330 and c4[1]+c4[2]+c4[3]+c5[1]+c5[2]+c5[3]<10){
        Return True
    }
    Else{
        Return False
    }
}

/*
判断格子是否为空
参数：
    D3W：int，窗口区域的宽度
    D3H：int，窗口区域的高度
    ID：int，格子编号
    ckpoints：object，要检查的位置的xy百分比list
返回：
    Bool
*/
isInventorySpaceEmpty(D3W, D3H, ID, ckpoints){
    _spaceSizeInnerW:=64
    _spaceSizeInnerH:=63
    m:=getInventorySpaceXY(D3W, D3H, ID)
    for i, p in ckpoints
    {
        xy:=[Round(m[3]+_spaceSizeInnerW*ckpoints[A_Index][1]*D3H/1440), Round(m[4]+_spaceSizeInnerH*ckpoints[A_Index][2]*D3H/1440)]
        PixelGetColor, cpixel, xy[1], xy[2], rgb
        c:=splitRGB(cpixel)
        if !(c[1]<22 and c[2]<20 and c[3]<15 and c[1]>c[3] and c[2]>c[3])
        {
            Return False
        }
    }
    Return True
}

/*
转化Object的所有key为字符串
参数：
    sep：分隔符
    dict：输入的字典
返回：
    String
*/
keyJoin(sep, dict){
    for key,value in dict
        str .= key . sep
    return SubStr(str, 1, -StrLen(sep))
}

/*
一个空方程，用于绑定Text控件的gLabel从而使tooltip可以工作
参数：
    无
返回：
    无
*/
dummyFunction(){
    Return
}

/*
为控件添加tooltip
修改自：https://gist.github.com/andreberg/55d003569f0564cd8695
参数：
    con：控件的hwnd
    text：tooltip字符串
    duration: tooltip的持续时间
    Modify：为1则修改一个已创建的tooltip
返回：
    无
*/
AddToolTip(con, text, duration=30000, Modify=0){
    Static TThwnd, GuiHwnd
    TInfo =
    UInt := "UInt"
    Ptr := (A_PtrSize ? "Ptr" : UInt)
    PtrSize := (A_PtrSize ? A_PtrSize : 4)
    Str := "Str"
    WM_USER := 0x400
    TTM_ADDTOOL := (A_IsUnicode ? WM_USER+50 : WM_USER+4)
    TTM_UPDATETIPTEXT := (A_IsUnicode ? WM_USER+57 : WM_USER+12)
    TTM_SETMAXTIPWIDTH := WM_USER+24
    TTM_SETDELAYTIME := WM_USER+3
    TTF_IDISHWND := 1
    TTF_CENTERTIP := 2
    TTF_RTLREADING := 4
    TTF_SUBCLASS := 16
    TTF_TRACK := 0x0020
    TTF_ABSOLUTE := 0x0080
    TTF_TRANSPARENT := 0x0100
    TTF_PARSELINKS := 0x1000
    TTF_AUTOPOP := 2
    If (!TThwnd) {
        Gui, +LastFound
        GuiHwnd := WinExist()
        TThwnd := DllCall("CreateWindowEx"
                    ,UInt,0
                    ,Str,"tooltips_class32"
                    ,UInt,0
                    ,UInt,2147483648
                    ,UInt,-2147483648
                    ,UInt,-2147483648
                    ,UInt,-2147483648
                    ,UInt,-2147483648
                    ,UInt,GuiHwnd
                    ,UInt,0
                    ,UInt,0
                    ,UInt,0)
    }
    cbSize := 6*4+6*PtrSize
    uFlags := TTF_IDISHWND|TTF_SUBCLASS|TTF_PARSELINKS
    VarSetCapacity(TInfo, cbSize, 0)
    NumPut(cbSize, TInfo)
    NumPut(uFlags, TInfo, 4)
    NumPut(GuiHwnd, TInfo, 8)
    NumPut(con, TInfo, 8+PtrSize)
    NumPut(&text, TInfo, 6*4+3*PtrSize)
    NumPut(0,TInfo, 6*4+6*PtrSize)
    DetectHiddenWindows, On
    If (!Modify) {
        DllCall("SendMessage", Ptr, TThwnd, UInt, TTM_ADDTOOL, Ptr, 0, Ptr, &TInfo, Ptr) 
        DllCall("SendMessage", Ptr, TThwnd, UInt, TTM_SETMAXTIPWIDTH, Ptr, 0, Ptr, A_ScreenWidth) 
        DllCall("SendMessage", Ptr, TThwnd, UInt, TTM_SETDELAYTIME, Ptr, TTF_AUTOPOP, Ptr, duration)
    }
    DllCall("SendMessage", Ptr, TThwnd, UInt, TTM_UPDATETIPTEXT, Ptr, 0, Ptr, &TInfo, Ptr)
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
        skillKey(currentProfile, nkey, D3W, D3H, forceStandingKey, skillset%currentProfile%useskillqueueckbox)
    }
Return

; 将currentProfile值关联到当前激活的tab
SetTabFocus:
    Gui, Submit, NoHide
    GuiControl, , StatuesSkillsetText, % tabsarray[ActiveTab]
    currentProfile:=ActiveTab
    Gosub, SetQuickPause
Return

; 设置快速暂停相关的快捷键和控件动画
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
            Hotkey, ~*%quickPauseHK%, quickPause, off
        } 
        Hotkey, ~*%currentQuickPauseHK%, quickPause, on
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
        Hotkey, ~*%currentQuickPauseHK%, quickPause, off
        quickPauseHK:=""
    }
Return

; 设置助手宏相关的控件动画
SetHelperKeybinding:
    Gui, Submit, NoHide
    mouseKeyArray:=["", "MButton", "WheelUp", "WheelDown", "XButton1", "XButton2", ""]
    GuiControlGet, HelperKeybindingdropdown
    GuiControlGet, HelperKeybindingHK
    Try
    {
        Hotkey, ~*%oldsandHelperHK%, oldsandHelper, off
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
        Hotkey, ~*%oldsandHelperHK%, oldsandHelper, off
        Hotkey, ~*%newoldsandHelperHK%, oldsandHelper, on
        oldsandHelperHK:=newoldsandHelperHK
    }
Return

; 设置配置的快速切换功能，以及相关控件动画
SetProfileKeybinding:
    Gui, Submit, NoHide
    mouseKeyArray:=["", "MButton", "WheelUp", "WheelDown", "XButton1", "XButton2", ""]
    Loop, %tabslen%
    {
        currentPage:=A_Index
        for key, value in profileKeybinding.Clone()
        {
            if (value = currentPage)
            {
                Hotkey, ~*%key%, SwitchProfile, Off
                profileKeybinding.Delete(key)
            }
        }
        switch skillset%currentPage%profilekeybindingdropdown
        {
            case 1:
                GuiControl, Disable, skillset%currentPage%profilekeybindinghkbox
            case 2,3,4,5,6:
                GuiControl, Disable, skillset%currentPage%profilekeybindinghkbox
                ckey:=mouseKeyArray[skillset%currentPage%profilekeybindingdropdown]
                Hotkey, ~*%ckey%, SwitchProfile, on
                profileKeybinding[ckey]:=currentPage
            case 7:
                GuiControl, Enable, skillset%currentPage%profilekeybindinghkbox
                ckey:=skillset%currentPage%profilekeybindinghkbox
                if (ckey)
                {
                    Hotkey, ~*%ckey%, SwitchProfile, on
                    profileKeybinding[ckey]:=currentPage
                } 
        }
    }
Return

; 处理配置快速切换逻辑
SwitchProfile:
    ;移除快捷键的modifier 
    currentHK:=StrReplace(StrReplace(A_ThisHotkey, "*"), "~")
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

; 设置开启战斗宏快捷键和相关控件动画
SetStartRun:
    Gui, Submit, NoHide
    startRunMouseKeyArray:=["RButton", "MButton", "WheelUp", "WheelDown", "XButton1", "XButton2", ""]
    if (StartRunDropdown = 7)
    {
        GuiControl, Enable, StartRunHKinput
        newstartRunHK:=StartRunHKinput
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
        Hotkey, ~*%startRunHK%, MainMacro, off
        Hotkey, ~*%newstartRunHK%, MainMacro, on
        startRunHK:=newstartRunHK
    }
Return

; 设置强制移动相关控件动画
SetMovingHelper:
    Gui, Submit, NoHide
    Loop, %tabslen%{
        npage:=A_Index
        if (skillset%npage%movingdropdown = 4)
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

; 设置按键宏策略控件动画
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

; 处理战斗宏的执行逻辑
MainMacro:
    VarSetCapacity(rect, 16)
    DllCall("GetClientRect", "ptr", WinExist("A"), "ptr", &rect)
    D3W:=NumGet(rect, 8, "int")
    D3H:=NumGet(rect, 12, "int")
    GuiControlGet, skillset%currentProfile%profilestartmodedropdown
    switch skillset%currentProfile%profilestartmodedropdown
    {
        ; 懒人模式
        case 1:
            if !vRunning
            {
                Gosub, RunMarco
            }
            Else
            {
                Gosub, StopMarco
            } 
        case 2:
        ; 仅按下时
            Gosub, RunMarco
            KeyWait, %startRunHK%
            Gosub, StopMarco
    }
Return

; 开启战斗宏
RunMarco:
    Gui, Submit, NoHide
    GuiControlGet, extraCustomStanding
    GuiControlGet, extraCustomStandingHK
    forceStandingKey:=extraCustomStanding? extraCustomStandingHK:"LShift"
    GuiControlGet, extraCustomMoving
    GuiControlGet, extraCustomMovingHK
    forceMovingKey:=extraCustomMoving? extraCustomMovingHK:"e"
    skillQueue:=[]
    ; 处理技能按键
    Loop, 6
    {
        GuiControlGet, skillset%currentProfile%s%A_Index%dropdown
        GuiControlGet, skillset%currentProfile%s%A_Index%hotkey
        Switch skillset%currentProfile%s%A_Index%dropdown
        {
        Case 2:
            k:=skillset%currentProfile%s%A_Index%hotkey
            Send {%k% Down}
            keysOnHold[k]:=1
        Case 3, 4:
            GuiControlGet, skillset%currentProfile%s%A_Index%updown
            SetTimer, spamSkillKey%A_Index%, % skillset%currentProfile%s%A_Index%updown
        Default:
            SetTimer, spamSkillKey%A_Index%, off
        }
    }
    ; 处理位移按键
    GuiControlGet, skillset%currentProfile%movingdropdown
    Switch skillset%currentProfile%movingdropdown
    {
        case 2:
            Send {%extraCustomStandingHK% Down}
            keysOnHold[extraCustomStandingHK]:=1
        case 3:
            Send {%extraCustomMovingHK% Down}
            keysOnHold[extraCustomMovingHK]:=1
        case 4:
            GuiControlGet, skillset%currentProfile%movingedit
            SetTimer, forceMoving, % skillset%currentProfile%movingedit

    }
    ; 处理按键队列
    if skillset%currentProfile%useskillqueueckbox{
        GuiControlGet, skillset%currentProfile%useskillqueueupdown
        sqfunc:=Func("spamSkillQueue").Bind(skillset%currentProfile%useskillqueueupdown)
        SetTimer, %sqfunc%, % skillset%currentProfile%useskillqueueupdown
    }
    vRunning:=True 
    vPausing:=False
Return

; 停止战斗宏
StopMarco:
    if IsObject(sqfunc){
        SetTimer, %sqfunc%, off
    }
    skillQueue:=[]
    Loop, 6
    {
        SetTimer, spamSkillKey%A_Index%, off
    }
    SetTimer, forceMoving, off
    for key, value in keysOnHold{
        if GetKeyState(key){
            Send {%key% up}
        }
    }
    keysOnHold:={}
    vRunning:=False
    vPausing:=False
Return

; 处理快速暂停按键
quickPause:
    GuiControlGet, skillset%currentProfile%clickpausedropdown1
    GuiControlGet, skillset%currentProfile%clickpauseupdown
    switch skillset%currentProfile%clickpausedropdown1
    {
        case 1:
            ; 双击
            If (A_PriorHotkey=A_ThisHotkey and A_TimeSincePriorHotkey < DblClickTime)
            {
                clickPauseMarco(keysOnHold, skillset%currentProfile%clickpauseupdown, vRunning)
            }
        case 2:
            clickPauseMarco(keysOnHold, skillset%currentProfile%clickpauseupdown, vRunning)
    }
Return

; 发送强制移动按键
forceMoving:
    if !vPausing
    {
        Send {%forceMovingKey%}
    }
Return

; 安全宏
safeGuard:
    ; 暗黑三不是焦点时停止宏
    If !WinActive("ahk_class D3 Main Window Class")
    {
        Gosub, StopMarco
    }
    ; 如果宏在运行，关闭一些控件防止误输入
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
Return
; ========================================= Hotkeys =======================================
~*Enter::
~*T::
~*M::
    if extraSmartPause
    {
        Gosub, StopMarco
    }
Return

~*Tab::
    if extraSmartPause
    {
        vPausing:=!vPausing
        if vPausing
        {
            for key, value in keysOnHold{
                if GetKeyState(key){
                    Send {%key% up}
                }
            }
        }
        Else
        {
            for key, value in keysOnHold{
                if !GetKeyState(key){
                    Send {%key% down}
                }
            }
        }
    }
Return

; ===================================== System Functions ==================================
GuiEscape:
GuiClose:
    Gui, Submit
    SaveCfgFile("d3oldsand.ini", tabs, currentProfile, safezone, VERSION)
Return

设置:
    Gui, Show,, %TIELE%
Return

退出:
    SaveCfgFile("d3oldsand.ini", tabs, currentProfile, safezone, VERSION)
ExitApp
