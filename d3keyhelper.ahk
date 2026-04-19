; =================================================================
;                  Diablo 3 "Oldsand" Key Helper (MIT License)
; Designed by Oldsand
; Перевод на русский: совместная работа
; 
; 
; Последняя версия: https://github.com/WeijieH/D3keyHelper
; Сообщения об ошибках и предложения приветствуются
; =================================================================

;@Ahk2Exe-IgnoreBegin
AHK_MIN_VERSION:="1.1.33.00"
if (A_AhkVersion < AHK_MIN_VERSION)
    MsgBox, 0x40, Если возникли ошибки - обновите AHK!, % Format("Данный помощник разработан на основе AHK v{:s}.`nВаша версия AHK: v{:s}.", AHK_MIN_VERSION, A_AhkVersion)

#SingleInstance Force
#NoEnv
#InstallKeybdHook
#InstallMouseHook
SetWorkingDir %A_ScriptDir%
SetBatchLines -1
Thread, interrupt, 0
CoordMode, Pixel, Client
CoordMode, Mouse, Client
Process, Priority, , High

VERSION:=260403
MainWindowW:=900
MainWindowH:=570
CompactWindowW:=551
TitleBarHight:=25
;@Ahk2Exe-Obey U_Y, U_Y := A_YYYY
;@Ahk2Exe-Obey U_M, U_M := A_MM
;@Ahk2Exe-Obey U_D, U_D := A_DD
;@Ahk2Exe-SetFileVersion 1.4.%U_Y%.%U_M%%U_D%
;@Ahk2Exe-SetLanguage 0x0804
;@Ahk2Exe-SetDescription Автокликер умений для Diablo 3
;@Ahk2Exe-SetProductName D3keyHelper
;@Ahk2Exe-SetCopyright Oldsand
;@Ahk2Exe-Bin Unicode 64-bit.bin
; ======================================== Глобальные переменные из конфига ===================================================
currentProfile:=ReadCfgFile("d3oldsand.ini", tabs, combats, others, generals)
SendMode, % generals.sendmode
tabsarray:=StrSplit(tabs, "`|")
tabslen:=ObjCount(tabsarray)
safezone:={}
isCompact:= generals.compactmode
runOnStart:= generals.runonstart
d3only:= generals.d3only
maxreforge:= (generals.maxreforge)?generals.maxreforge:10
TitleString:=(d3only)? "Diablo 3 — Автокликер умений":"Автокликер мыши и клавиатуры"
TITLE:=Format(TitleString " v1.4.{:d}   by Oldsand", VERSION)
helperMouseSpeed:= generals.helpermousespeed
helperAnimationDelay:= generals.helperanimationdelay
gameResolution:= InStr(generals.gameresolution, "x")? generals.gameresolution:"Auto"
hBMPButtonLeft_Normal := isCompact? hBMPButtonExpand_Normal:hBMPButtonBack_Normal
hBMPButtonLeft_Hover := isCompact? hBMPButtonExpand_Hover:hBMPButtonBack_Hover
hBMPButtonLeft_Pressed := isCompact? hBMPButtonExpand_Pressed:hBMPButtonBack_Pressed
Loop, Parse, % generals.safezone, CSV
{
    safezone[A_LoopField]:=1
}
#If WinActive((d3only)?"ahk_class D3 Main Window Class":"A")
gameGamma:=(generals.gamegamma>=0.5 and generals.gamegamma<=1.5)? generals.gamegamma:1
buffpercent:=(generals.buffpercent>=0 and generals.buffpercent<=1)? generals.buffpercent:0.05
; ==============================================================================================================
GuiCreate()
SetTrayMenu()
StartUp()
showMainWindow(isCompact? CompactWindowW:MainWindowW, MainWindowH)

OnExit("OnUnload")
Return

; =================================== Пользовательские функции =====================================
/*
Выполняется при загрузке программы
Параметры:
    нет
Возврат:
    нет
*/
OnLoad(){
    Global
    Static Init := OnLoad() ; Выполняется перед всеми остальными командами

    ; ============================================ Глобальные переменные ===========================================================
    vRunning:=False
    vPausing:=False
    vFront:=True
    helperDelay:=100
    mouseDelay:=2
    helperRunning:=False
    helperBreak:=False
    profileKeybinding:={}
    keysOnHold:={}
    lastpotion:=[]
    DblClickTime:=DllCall("GetDoubleClickTime", "UInt")
    RightButtonState:=0
    LeftButtonState:=0
    _CloseButtonNormal := "iVBORw0KGgoAAAANSUhEUgAAAB4AAAAZCAYAAAAmNZ4aAAAAM0lEQVRIiWMYBaNgFIyCUUAsYCSkrnLe2v/khGZ7UjBes5lGo2gUjIJRMApGAVbAwMAAAMjYBAQ0LnL/AAAAAElFTkSuQmCC"
    _CloseButtonHover := "iVBORw0KGgoAAAANSUhEUgAAAB4AAAAZCAYAAAAmNZ4aAAAARklEQVRIiWN8Iaj8n2EAANNAWMowavGoxaMWj1pMCWAhpFfi/V2yjH8hqIxXfvD6mJDLyQWjqXrU4lGLRy0etZg4wMDAAACGJAZtrV+pPwAAAABJRU5ErkJggg=="
    _CloseButtonPressed := "iVBORw0KGgoAAAANSUhEUgAAAB4AAAAZCAYAAAAmNZ4aAAAARklEQVRIiWO85B32n2EAANNAWMowavGoxaMWj1pMCWAhpNftwjGyjN9lYIVXfvD6mJDLyQWjqXrU4lGLRy0etZg4wMDAAACzuwbMPgoPPgAAAABJRU5ErkJggg=="
    _BackButtonNormal := "iVBORw0KGgoAAAANSUhEUgAAAB4AAAAZCAYAAAAmNZ4aAAAAr0lEQVRIie2UwQ3CMAwA7Sa0ZhJGYAOWoCswDSuUJdiAFViEtIkw8gMJEKiOan4+KY9Iji+2nIDjOFagJs/xfHnb853345gGItoA4vUz/rDbzuZsagsQaUq3IYQA36RaqsRPaYwRVm2r6tYv1GJLKWjF1lIhaoJkkJgZcs6yWHFk9nIqcddRb12xqtXY4Ilo3ZdSIE+TpmIb8T/kVc/JUl79gbzKZdqXyB3HWQ4APACzI1jSHwESAQAAAABJRU5ErkJggg=="
    _BackButtonHover := "iVBORw0KGgoAAAANSUhEUgAAAB4AAAAZCAYAAAAmNZ4aAAAAn0lEQVRIiWOUbL31n2EAANNAWMowavGoxYPO4j/XdzK8qFFn+Pf+Ef0sBln6ekkuAzMLCwOToBx9LIZZysLKyiDacJVsS0mymJqWEm0xtS0FAaLKalBC+v+f+CJdsvUWQTUsxBgkEj2J6j4mKqhZNN0ZRGMmM/z5/ZvhdYM2/SymheUkZSdqWk5yAYJsOSi1kwtGWyCjFo9aTB3AwMAAAPFsSKyupuluAAAAAElFTkSuQmCC"
    _BackButtonPressed := "iVBORw0KGgoAAAANSUhEUgAAAB4AAAAZCAYAAAAmNZ4aAAAAn0lEQVRIiWM0nnz/P8MAAKaBsJRh1OJRiwedxZ8v7WA4l6fE8OvtI/pZDLL01uxMBmYWFgY2YTn6WAyzlIWVlUG/7xbZlpJkMTUtJdpialsKAkSV1aCE9P8/8UW68eT7BNWwEGOQaso0qvuYqKDm1fNgUEudzvDn92+Gi0Vq9LOYFpaTlJ2oaTnJBQiy5aDUTi4YbYGMWjxqMXUAAwMDALPRSRXM0WlaAAAAAElFTkSuQmCC"
    _ExpandButtonNormal := "iVBORw0KGgoAAAANSUhEUgAAAB4AAAAZCAYAAAAmNZ4aAAAAVUlEQVRIiWMYBaOAVoCRkLmV89b+J8fu9qRgvGYzDVSUEvTx//9keZiBkRG/0QPm45FnMQshBVXz15EXyQTSz2iqphsYTdUYYNil6lEwCoYZYGBgAACe2A+sakz0agAAAABJRU5ErkJggg=="
    _ExpandButtonHover := "iVBORw0KGgoAAAANSUhEUgAAAB4AAAAZCAYAAAAmNZ4aAAAASklEQVRIiWOUbL31n2EAANNAWMowajE9AQshu55Xq5HlHMnWW0PUx4RcTi4YzU50A6OpGgOMpuohb/FoqsYAo6l61OKhZTEDAwMAZw0QHWhren8AAAAASUVORK5CYII="
    _ExpandButtonPressed := "iVBORw0KGgoAAAANSUhEUgAAAB4AAAAZCAYAAAAmNZ4aAAAASklEQVRIiWM0nnz/P8MAAKaBsJRh1GJ6AhZCdp3NVSTLOcaT7w9RHxNyOblgNDvRDYymagwwmqqHvMWjqRoDjKbqUYuHlsUMDAwASU0Q0fLg6gAAAAAASUVORK5CYII="
    DllCall("LoadLibrary", "Str", "Crypt32.dll")
    DllCall("LoadLibrary", "Str", "Shlwapi.dll")
    DllCall("LoadLibrary", "Str", "Gdiplus.dll")
    VarSetCapacity(GdiplusStartupInput, (A_PtrSize = 8 ? 24 : 16), 0) ; GdiplusStartupInput structure
    NumPut(1, GdiplusStartupInput, 0, "UInt") ; GdiplusVersion
    DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &GdiplusStartupInput, "Ptr", 0) ; Initialize GDI+

    hBMPButtonClose_Normal := GdipCreateFromBase64(_CloseButtonNormal)
    hBMPButtonClose_Hover := GdipCreateFromBase64(_CloseButtonHover)
    hBMPButtonClose_Pressed := GdipCreateFromBase64(_CloseButtonPressed)
    hBMPButtonBack_Normal := GdipCreateFromBase64(_BackButtonNormal)
    hBMPButtonBack_Hover := GdipCreateFromBase64(_BackButtonHover)
    hBMPButtonBack_Pressed := GdipCreateFromBase64(_BackButtonPressed)
    hBMPButtonExpand_Normal := GdipCreateFromBase64(_ExpandButtonNormal)
    hBMPButtonExpand_Hover := GdipCreateFromBase64(_ExpandButtonHover)
    hBMPButtonExpand_Pressed := GdipCreateFromBase64(_ExpandButtonPressed)
}

/*
Очистка при выходе из программы
Параметры:
    нет
Возврат:
    нет
*/
OnUnload(ExitReason, ExitCode){
    Global ; Assume-Global mode
    ; Освобождение ресурсов GDI+
    DllCall("GdiplusShutdown", "Ptr", pToken)
    DllCall("DeregisterShellHookWindow", "Ptr", A_ScriptHwnd)
    if (hHookMouse){
        DllCall("UnhookWindowsHookEx", "Uint", hHookMouse)
    }
}

/*
Создание графического интерфейса
Параметры:
    нет
Возвращаемое значение:
    нет
*/
GuiCreate(){
    Global
    local tabw:=MainWindowW-357
    local tabh:=MainWindowH-35-TitleBarHight
    local helperSettingGroupx:=MainWindowW-345

    Gui Font, s11, Segoe UI
    Gui -MaximizeBox -MinimizeBox +Owner +DPIScale +LastFound -Caption -Border
    Gui, Margin, 5, % TitleBarHight+10
    ; SS_BITMAP:=0x40
    ; SS_REALSIZECONTROL:=0x0E
    ; 0x4E=SS_BITMAP|SS_REALSIZECONTROL
    Gui, Add, Picture, % "x1 y1 w" MainWindowW-2 " h" TitleBarHight " +0x4E hwndTitlebarID vTitlebar"
    Gui, Add, Picture, % "x1 y+0 w" MainWindowW-2 " h1 +0x4E hwndTitlebarLineID"
    Gui, Add, Picture, % "x0 y0 w" MainWindowW " h1 +0x4E hwndBorderTopID vBorderTop"
    Gui, Add, Picture, % "x0 y" MainWindowH-1 " w" MainWindowW " h1 +0x4E hwndBorderBottomID vBorderBottom"
    Gui, Add, Picture, % "x0 y1 w1 h" MainWindowH-2 " +0x4E hwndBorderLeftID vBorderLeft"
    Gui, Add, Picture, % "x" MainWindowW-1 " y1 w1 h" MainWindowH-2 " +0x4E hwndBorderRightID vBorderRight"
    Gui, Add, Text, % "x1 y1 h" TitleBarHight " hwndTitleBarTextID vTitleBarText +BackgroundTrans +0x200", %TITLE%
    Gui, Add, Picture, % "x" MainWindowW-31 " y1 w-1 h" TitleBarHight " hwndUIRightButtonID vUIRightButton gdummyFunction +BackgroundTrans", % "HBITMAP:*" hBMPButtonClose_Normal
    AddToolTip(UIRightButtonID, "ЛКМ: сохранить настройки и свернуть окно в правый нижний угол`nПКМ: сохранить настройки и выйти из программы")
Gui, Add, Picture, % "x" 1 " y1 w-1 h" TitleBarHight " hwndUILeftButtonID vUILeftButton gdummyFunction +BackgroundTrans", % "HBITMAP:*" hBMPButtonLeft_Normal
AddToolTip(UILeftButtonID, "Нажмите для переключения между полным и компактным макетом")
    GuiControlGet, TitleBarSize, Pos , TitleBarText
    Gui Add, Tab3, xm ym w%tabw% h%tabh% vActiveTab gSetTabFocus AltSubmit, %tabs%
    Gui Font, s9, Segoe UI
    local skillLabels:=["Умение 1:", "Умение 2:", "Умение 3:", "Умение 4:", "Левый навык:", "Правый навык:"]
    Loop, parse, tabs, `|
    {
        local currentTab:=A_Index
        Gui Tab, %currentTab%
        Gui Add, Hotkey, x0 y0 w0 w0
        
               Gui Add, GroupBox, xm+10 ym+40 w520 h260 section, Настройки макросов клавиш
        Gui Add, Text, xs+90 ys+20 w60 center section, Клавиша
        Gui Add, Text, x+10 w80 center, Режим
        Gui Add, Text, x+15 w110 center, Интервал (мс)
        Gui Add, Text, x+5 w90 center, Задержка (мс)
        Gui Add, Text, x+0 center, Случ. задержка
        Loop, 6
        {
            Gui Add, Text, xs-75 w70 yp+34 center, % skillLabels[A_Index]
            local ac:=combats[currentTab][A_Index]["action"]
            local rd:=combats[currentTab][A_Index]["random"]
            switch A_Index
            {
                case 1,2,3,4:
                    Gui Add, Hotkey, x+5 yp-2 w60 vskillset%currentTab%s%A_Index%hotkey, % combats[currentTab][A_Index]["hotkey"]
                case 5:
                    Gui Add, Edit, x+5 yp-2 w60 vskillset%currentTab%s%A_Index%hotkey +Disabled, LButton
                case 6:
                    Gui Add, Edit, x+5 yp-2 w60 vskillset%currentTab%s%A_Index%hotkey +Disabled, RButton
            }
            Gui Add, DropDownList, x+10 w80 AltSubmit Choose%ac% gSetSkillsetDropdown vskillset%currentTab%s%A_Index%dropdown, Отключено||Удерживать||Автоповтор||Поддержание баффа||По нажатию
            Gui Add, Edit, vskillset%currentTab%s%A_Index%edit x+20 w90 Number
            Gui Add, Updown, vskillset%currentTab%s%A_Index%updown gSetSkillQueueWarning Range20-60000, % combats[currentTab][A_Index]["interval"]
            Gui Add, Edit, vskillset%currentTab%s%A_Index%delayedit hwndskillset%currentTab%s%A_Index%delayeditID x+25 w70
            Gui Add, Updown, vskillset%currentTab%s%A_Index%delayupdown Range-30000-30000, % combats[currentTab][A_Index]["delay"]
            AddToolTip(skillset%currentTab%s%A_Index%delayeditID, "Положительное число — задержка выполнения, отрицательное — упреждение. Установите 0 для отключения задержки")
            Gui Add, Checkbox, x+35 yp+2 Checked%rd% vskillset%currentTab%s%A_Index%randomckbox hwndskillset%currentTab%s%A_Index%randomckboxID
            AddToolTip(skillset%currentTab%s%A_Index%randomckboxID, "При включении реальная задержка при каждом срабатывании будет случайным числом от 0 до заданного значения")
        }
        Gui Add, GroupBox, xm+10 yp+45 w520 h192 section, Дополнительные настройки
        Gui Add, Text, xs+20 ys+27, Быстрое переключение на эту конфигурацию:
        Gui Add, DropDownList, % "x+5 yp-3 w90 AltSubmit Choose" others[currentTab].profilemethod " vskillset" currentTab "profilekeybindingdropdown gSetProfileKeybinding", Нет||Средняя кнопка мыши||Колесо вверх||Колесо вниз||Боковая кнопка 1||Боковая кнопка 2||Клавиша клавиатуры
        Gui Add, Hotkey, x+15 w100 vskillset%currentTab%profilekeybindinghkbox gSetProfileKeybinding, % others[currentTab].profilehotkey
        Gui Add, Checkbox, % "x+15 yp+3 Checked" others[currentTab].autostartmarco " vskillset" currentTab "autostartmarcockbox hwndskillset" currentTab "autostartmarcockboxID", Автозапуск макроса при переключении
        AddToolTip(skillset%currentTab%autostartmarcockboxID, "При включении боевой макрос, запущенный в ленивом режиме, может бесшовно переключаться во время работы")

        Gui Add, Text, xs+20 yp+35, Режим запуска макроса：
                Gui Add, DropDownList, % "x+5 yp-3 w90 AltSubmit Choose" others[currentTab].lazymode " hwndprofileStartModeDropdown" currentTab "ID vskillset" currentTab "profilestartmodedropdown gSetStartMode", Ленивый режим||Только при удержании||Только одно нажатие
        AddToolTip(profileStartModeDropdown%currentTab%ID, "Ленивый режим: нажмите горячую клавишу для запуска макроса, повторное нажатие останавливает`nТолько при удержании: макрос активен, пока зажата горячая клавиша`nТолько одно нажатие: при нажатии горячей клавиши однократно прожимаются все «зажатые» умения")
        Gui Add, Checkbox, % "x+20 yp+3 Checked" others[currentTab].useskillqueue " hwnduseskillqueueckbox" currentTab "ID vskillset" currentTab "useskillqueueckbox gSetSkillQueue", Использовать очередь клавиш (мс):
        AddToolTip(useskillqueueckbox%currentTab%ID, "При включении клавиши не отправляются мгновенно, а помещаются в очередь`nАвтоповтор помещает умение в начало очереди, поддержание баффа — в конец`nПри автоповторе автоматически зажимается кнопка принудительного стояния")
        Gui Add, Edit, vskillset%currentTab%useskillqueueedit hwnduseskillqueueedit%currentTab%ID x+0 yp-3 w50 Number
        Gui Add, Updown, vskillset%currentTab%useskillqueueupdown gSetSkillQueueWarning Range50-1000, % others[currentTab].useskillqueueinterval
        AddToolTip(useskillqueueedit%currentTab%ID, "Клавиши автоповтора из очереди будут отправляться с этим интервалом в окно игры")
        Gui Add, Text, x+8  yp+3 vskillset%currentTab%skillqueuewarningtext hwndskillset%currentTab%skillqueuewarningtextID gdummyFunction +cRed +Hidden, % "Внимание!"
        AddToolTip(skillset%currentTab%skillqueuewarningtextID, "Некорректные настройки очереди клавиш")

        Gui Add, Checkbox, % "xs+20 yp+35 Checked" others[currentTab].enablequickpause " vskillset" currentTab "clickpauseckbox gSetQuickPause", Быстрая пауза:
        Gui Add, DropDownList, % "x+0 yp-3 w50 AltSubmit Choose" others[currentTab].quickpausemethod1 " vskillset" currentTab "clickpausedropdown1 gSetQuickPause", Двойной клик||Одинарный клик||Удержание
        Gui Add, DropDownList, % "x+5 yp w75 AltSubmit Choose" others[currentTab].quickpausemethod2 " vskillset" currentTab "clickpausedropdown2 gSetQuickPause", Левая кнопка мыши||Правая кнопка мыши||Средняя кнопка мыши||Боковая кнопка 1||Боковая кнопка 2
        Gui Add, Text, x+5 yp+3 vskillset%currentTab%clickpausetext1, тогда
        Gui Add, DropDownList, % "x+5 yp-3 w140 AltSubmit Choose" others[currentTab].quickpausemethod3 " vskillset" currentTab "clickpausedropdown3", Приостановить макрос клавиш||Приостановить макрос и зажать ЛКМ
        Gui Add, Edit, vskillset%currentTab%clickpauseedit x+5 yp w60 Number
        Gui Add, Updown, vskillset%currentTab%clickpauseupdown Range500-5000, % others[currentTab].quickpausedelay
        Gui Add, Text, x+5 yp+3 vskillset%currentTab%clickpausetext2, мс

        Gui Add, Text, xs+20 yp+35, Помощь в движении:
        Gui Add, DropDownList, % "x+5 yp-3 w150 AltSubmit Choose" pfmv:=others[currentTab].movingmethod " vskillset" currentTab "movingdropdown gSetMovingHelper", Отключено||Принудительно стоять||Принудительное движение (зажать)||Принудительное движение (автоповтор)
        Gui Add, Text, vskillset%currentTab%movingtext x+10 yp+3, Интервал (мс):
        Gui Add, Edit, vskillset%currentTab%movingedit x+5 yp-3 w60 Number
        Gui Add, Updown, vskillset%currentTab%movingupdown Range20-3000, % others[currentTab].movinginterval

        Gui Add, Text, xs+20 yp+35, Помощь с зельем:
        Gui Add, DropDownList, % "x+5 yp-3 w120 AltSubmit Choose" pfpo:=others[currentTab].potionmethod "hwndpotionDropdown" currentTab "ID vskillset" currentTab "potiondropdown gSetMovingHelper", Отключено||Периодический автоповтор||Поддержание КД зелья
        AddToolTip(potionDropdown%currentTab%ID, "Периодический автоповтор: нажимать клавишу зелья с заданным интервалом`nПоддержание КД зелья: нажимать клавишу зелья сразу после окончания кулдауна, чтобы зелье как можно скорее снова ушло на перезарядку")
        Gui Add, Text, vskillset%currentTab%potiontext x+10 yp+3, Интервал (мс):
        Gui Add, Edit, vskillset%currentTab%potionedit x+5 yp-3 w60 Number
        Gui Add, Updown, vskillset%currentTab%potionupdown Range200-30000, % others[currentTab].potioninterval
    }
    Gui Tab
    GuiControl, Choose, ActiveTab, % currentProfile

    Gui Add, GroupBox, x%helperSettingGroupx% ym+40 w338 h470 section, Вспомогательные функции
    oldsandhelperhk:=generals.oldsandhelperhk
    Gui Font,s10
    Gui Add, Text, xs+20 ys+30 +cRed, Горячая клавиша помощника:
    Gui Font,s9
    Gui Add, DropDownList, % "x+0 yp-3 w75 vhelperKeybindingdropdown gSetHelperKeybinding AltSubmit Choose" generals.oldsandhelpermethod, Нет||Средняя кнопка мыши||Колесо вверх||Колесо вниз||Боковая кнопка 1||Боковая кнопка 2||Клавиша клавиатуры
    Gui Add, Hotkey, x+5 w70 vhelperKeybindingHK gSetHelperKeybinding, %oldsandhelperhk%

    Gui Add, Text, xs+20 yp+40 hwndhelperSpeedTextID gdummyFunction, Скорость анимации помощника:
    AddToolTip(helperSpeedTextID, "При высоком пинге снижение скорости анимации уменьшает вероятность ошибок макроса")
        Gui Add, DropDownList, % "x+5 yp-3 w90 vhelperAnimationSpeedDropdown hwndhelperAnimationSpeedDropdownID AltSubmit Choose" generals.helperspeed, Очень быстро||Быстро||Средне||Медленно||Вручную
    AddToolTip(helperAnimationSpeedDropdownID, "Очень быстро: скорость мыши 0, задержка анимации 50`nБыстро: скорость мыши 1, задержка анимации 100`nСредне: скорость мыши 2, задержка анимации 150`nМедленно: скорость мыши 3, задержка анимации 200`nВручную: использовать значения из файла конфигурации")

    Gui Add, Text, x+20 yp+4 w80 hwndhelperSafeZoneTextID vhelperSafeZoneText gdummyFunction
    AddToolTip(helperSafeZoneTextID, "Измените значение safezone в секции Generals конфигурационного файла для настройки защищённых ячеек`nФормат: номера ячеек, разделённые запятой`nЛевая верхняя ячейка — 1, правая верхняя — 10, левая нижняя — 51, правая нижняя — 60")

    Gui Add, CheckBox, % "xs+20 yp+35 hwndextraGambleHelperCKboxID vextraGambleHelperCKbox gSetGambleHelper Checked" generals.enablegamblehelper, Помощник Кадалы (гэмбл)：
    AddToolTip(extraGambleHelperCKboxID, "При активации помощника горячей клавишей автоматически нажимается правая кнопка мыши")
Gui Add, Text, vextraGambleHelperText x+5 yp, Правых кликов
    Gui Add, Edit, vextraGambleHelperEdit x+10 yp-4 w60 Number
    Gui Add, Updown, vextraGambleHelperUpdown Range2-60, % generals.gamblehelpertimes

    Gui Add, CheckBox, % "xs+20 yp+40 hwndextraLootHelperCkboxID vextraLootHelperCkbox gSetLootHelper Checked" generals.enableloothelper, Помощник быстрого сбора:
    AddToolTip(extraLootHelperCkboxID, "При сборе предметов нажатие горячей клавиши помощника автоматически кликает левой кнопкой мыши")
    Gui Add, Text, vextraLootHelperText x+5 yp, Левых кликов
    Gui Add, Edit, vextraLootHelperEdit x+10 yp-4 w60 Number
    Gui Add, Updown, vextraLootHelperUpdown Range2-99, % generals.loothelpertimes

    Gui Add, CheckBox, % "xs+20 yp+40 hwndextraSalvageHelperCkboxID vextraSalvageHelperCkbox gSetSalvageHelper Checked" generals.enablesalvagehelper, Помощник кузнеца (разбор):
        Gui Add, DropDownList, % "x+5 yp-4 w180 AltSubmit hwndextraSalvageHelperDropdownID vextraSalvageHelperDropdown gSetSalvageHelper Choose" generals.salvagehelpermethod, Быстрое разбирание||Разобрать всё||Умное разбирание||Умное разбирание (оставлять древние, святые, эфирные)||Умное разбирание (оставлять только первозданные)
    AddToolTip(extraSalvageHelperCkboxID, "При разборе предметов нажатие горячей клавиши помощника автоматически выполняет выбранную стратегию")
    AddToolTip(extraSalvageHelperDropdownID, "Быстрое разбирание: нажатие горячей клавиши эквивалентно клику левой кнопкой мыши + Enter`nРазобрать всё: разобрать всё снаряжение в инвентаре, кроме защищённых ячеек`nУмное разбирание: как «Разобрать всё», но пропускает древние, святые и первозданные предметы`nУмное разбирание (оставлять древние, святые, эфирные): сохраняет только древние, святые и эфирные предметы`nУмное разбирание (оставлять только первозданные): сохраняет только первозданные предметы")

        Gui Add, CheckBox, % "xs+20 yp+40 hwndextraReforgeHelperCkboxID vextraReforgeHelperCkbox gSetReforgeHelper Checked" generals.enablereforgehelper, Помощник перековки в Кубе:
    Gui Add, DropDownList, % "x+5 yp-4 w180 AltSubmit hwndextraReforgeHelperDropdownID vextraReforgeHelperDropdown Choose" generals.reforgehelpermethod, Перековать один раз||Перековка до древнего/первозданного||Перековка до первозданного
    AddToolTip(extraReforgeHelperCkboxID, "Когда Куб открыт на странице перековки, нажатие горячей клавиши помощника запускает выбранную стратегию`n***Максимальное число попыток можно изменить в конфиге через переменную maxreforge***")
    local strMaxReforge1:= "Бесконечно перековывать предмет под курсором, пока он не станет древним или первозданным (максимум попыток: " maxreforge ")"
    local strMaxReforge2:= "Бесконечно перековывать предмет под курсором, пока он не станет первозданным (максимум попыток: " maxreforge ")"
    AddToolTip(extraReforgeHelperDropdownID, "Перековать один раз: одна перековка предмета под курсором`nПерековка до древнего/первозданного: " strMaxReforge1 "`nПерековка до первозданного: " strMaxReforge2 "`n***Повторное нажатие горячей клавиши во время работы прерывает макрос!***")

        Gui Add, CheckBox, % "xs+20 yp+40 hwndextraUpgradeHelperCkboxID vextraUpgradeHelperCkbox gSetSalvageHelper Checked" generals.enableupgradehelper, Помощник улучшения в Кубе
    AddToolTip(extraUpgradeHelperCkboxID, "Когда Куб открыт на странице улучшения, нажатие горячей клавиши помощника автоматически улучшает все редкие (жёлтые) предметы в инвентаре, кроме защищённых ячеек")

    Gui Add, CheckBox, % "x+20 yp+0 hwndextraConvertHelperCkboxID vextraConvertHelperCkbox gSetSalvageHelper Checked" generals.enableconverthelper, Помощник конвертации в Кубе
    AddToolTip(extraConvertHelperCkboxID, "Когда Куб открыт на странице конвертации материалов, нажатие горячей клавиши помощника автоматически использует все предметы в инвентаре (кроме защищённых ячеек) для конвертации")

    Gui Add, CheckBox, % "xs+20 yp+36 hwndextraAbandonHelperCkboxID vextraAbandonHelperCkbox gSetSalvageHelper Checked" generals.enableabandonhelper, Помощник быстрой выброски/складирования
    AddToolTip(extraAbandonHelperCkboxID, "При открытом инвентаре и нахождении курсора в области инвентаря, нажатие горячей клавиши помощника автоматически выбрасывает все предметы из незащищённых ячеек`nЕсли открыт сундук и курсор находится в ячейках сундука, макрос перемещает все предметы из незащищённых ячеек инвентаря в сундук")

        Gui Add, CheckBox, % "xs+20 yp+55 vextraSoundonProfileSwitch Checked" generals.enablesoundplay, Звуковой сигнал при смене конфигурации
    Gui Add, CheckBox, % "x+20 yp+0 hwndextraSmartPauseID vextraSmartPause Checked" generals.enablesmartpause, Умная пауза
    AddToolTip(extraSmartPauseID, "При включении макрос ставится на паузу клавишей Tab`nКлавиши Enter, M, T останавливают макрос")

    Gui Add, CheckBox, % "xs+20 yp+35 vextraCustomStanding gSetCustomStanding Checked" generals.customstanding, Использовать свою клавишу принудительного стояния:
    Gui Add, Hotkey, x+5 yp-3 w70 vextraCustomStandingHK gSetCustomStanding, % generals.customstandinghk

    Gui Add, CheckBox, % "xs+20 yp+35 vextraCustomMoving gSetCustomMoving Checked" generals.custommoving, Использовать свою клавишу принудительного движения:
    Gui Add, Hotkey, x+5 yp-3 w70 Limit14 vextraCustomMovingHK gSetCustomMoving, % generals.custommovinghk

    Gui Add, CheckBox, % "xs+20 yp+35 vextraCustompotion gSetCustomPotion Checked" generals.custompotion, Использовать свою клавишу зелья:
    Gui Add, Hotkey, x+5 yp-3 w70 Limit14 vextraCustompotionHK gSetCustomPotion, % generals.custompotionhk

    startRunHK:=generals.starthotkey
    Gui Font, s10
    Gui Add, Text, x570 ym+3 +cRed, Горячая клавиша боевого макроса:
    Gui Font, s9
    Gui Add, DropDownList, % "x+5 yp-3 w90 vStartRunDropdown gSetStartRun AltSubmit Choose" generals.startmethod, Правая кнопка мыши||Средняя кнопка мыши||Колесо вверх||Колесо вниз||Боковая кнопка 1||Боковая кнопка 2||Клавиша клавиатуры
    Gui Add, Hotkey, x+5 yp w70 vStartRunHKinput gSetStartRun, %startRunHK%

    Gui Add, Text, % "x10 y" MainWindowH-20 " section", Активная конфигурация:
    Gui Font, s11
    Gui Add, Text, x+5 ys-4 w300 +cRed vStatuesSkillsetText, % tabsarray[currentProfile]
    Gui Add, Text, x505 yp +cRed hwndCurrentmodeTextID gdummyFunction, % A_SendMode
    Gui Font, s9
        Gui Add, Text, xp-95 ys hwndSendmodeTextID gdummyFunction, Режим отправки клавиш:
    AddToolTip(SendmodeTextID, "Измените значение sendmode в секции General конфигурационного файла для смены режима отправки клавиш")
    AddToolTip(CurrentmodeTextID, "Event: режим по умолчанию, лучшая совместимость`nInput: рекомендуемый режим, максимальная скорость, но может блокироваться некоторыми антивирусами")
    Gui Add, Link, x570 ys hwndAboutLinkID, Проект с открытым исходным кодом: <a href="https://github.com/WeijieH/D3keyHelper">https://github.com/WeijieH/D3keyHelper</a>
    AddToolTip(AboutLinkID, "Не забудьте поставить звёздочку ╰(*°▽°*)╯")
    Return
}

/*
Инициализация после создания GUI
Параметры:
    нет
Возврат:
    нет
*/
StartUp(){
    Global
    Gosub, SetSkillsetDropdown
    Gosub, SetStartRun
    Gosub, SetProfileKeybinding
    Gosub, SetMovingHelper
    Gosub, SetHelperKeybinding
    Gosub, SetQuickPause
    SetGambleHelper()
    SetLootHelper()
    SetReforgeHelper()
    SetSalvageHelper()
    SetCustomStanding()
    SetCustomMoving()
    SetCustomPotion()
    SetSkillQueue()
    SetStartMode()

    DllCall("RegisterShellHookWindow", "Ptr", A_ScriptHwnd)
    hHookMouse:=0
    OnMessage(DllCall("RegisterWindowMessage", "Str", "SHELLHOOK"), "Watchdog")
    Watchdog(4, 0)
}

/*
Настройка меню иконки в трее
Параметры:
    нет
Возврат:
    нет
*/
SetTrayMenu(){
    Global
    Menu, Tray, NoStandard
    Menu, Tray, Add, Настройки, GuiShowMainWindow
    Menu, Tray, Add, Выход, GuiExit
    Menu, Tray, Default, Настройки
    Menu, Tray, Click, 1
    Menu, Tray, Tip, %TITLE%
    Menu, Tray, Icon, , , 1
}

/*
Чтение конфигурационного файла, если файла нет — возврат значений по умолчанию
Параметры:
    cfgFileName: имя файла
    tabs: ByRef String, имена конфигураций через "|", для инициализации Tab
    combats: ByRef Array, настройки боевого макроса
    others: ByRef Array, дополнительные настройки
    generals: ByRef Array, общие настройки
Возврат:
    Номер последней активной конфигурации для инициализации Tab
*/
ReadCfgFile(cfgFileName, ByRef tabs, ByRef combats, ByRef others, ByRef generals){
    local
    Global VERSION
    if FileExist(cfgFileName)
    {
        generals:={}
        IniRead, ver, %cfgFileName%, General, version
        if (VERSION != ver)
        {
            MsgBox, Версия файла конфигурации не совпадает. При ошибках удалите файл d3oldsand.ini и настройте заново.
        }
        IniRead, currentProfile, %cfgFileName%, General, activatedprofile, 1
        IniRead, oldsandhelperhk, %cfgFileName%, General, oldsandhelperhk, F5
        IniRead, oldsandhelpermethod, %cfgFileName%, General, oldsandhelpermethod, 7
        IniRead, enablegamblehelper, %cfgFileName%, General, enablegamblehelper, 1
        IniRead, gamblehelpertimes, %cfgFileName%, General, gamblehelpertimes, 15
        IniRead, enablesalvagehelper, %cfgFileName%, General, enablesalvagehelper, 0
        IniRead, salvagehelpermethod, %cfgFileName%, General, salvagehelpermethod, 1
        IniRead, reforgehelpermethod, %cfgFileName%, General, reforgehelpermethod, 1
        IniRead, enablereforgehelper, %cfgFileName%, General, enablereforgehelper, 0
        IniRead, enableconverthelper, %cfgFileName%, General, enableconverthelper, 0
        IniRead, enableupgradehelper, %cfgFileName%, General, enableupgradehelper, 0
        IniRead, enablesmartpause, %cfgFileName%, General, enablesmartpause, 0
        IniRead, enablesoundplay, %cfgFileName%, General, enablesoundplay, 1
        IniRead, enableabandonhelper, %cfgFileName%, General, enableabandonhelper, 0
        IniRead, startmethod, %cfgFileName%, General, startmethod, 7
        IniRead, starthotkey, %cfgFileName%, General, starthotkey, F2
        IniRead, custommoving, %cfgFileName%, General, custommoving, 0
        IniRead, custommovinghk, %cfgFileName%, General, custommovinghk, e
        IniRead, customstanding, %cfgFileName%, General, customstanding, 0
        IniRead, customstandinghk, %cfgFileName%, General, customstandinghk, LShift
        IniRead, custompotion, %cfgFileName%, General, custompotion, 0
        IniRead, custompotionhk, %cfgFileName%, General, custompotionhk, q
        IniRead, safezone, %cfgFileName%, General, safezone, "61,62,63"
        IniRead, helperspeed, %cfgFileName%, General, helperspeed, 3
        IniRead, gamegamma, %cfgFileName%, General, gamegamma, 1.000000
        IniRead, sendmode, %cfgFileName%, General, sendmode, "Event"
        IniRead, buffpercent, %cfgFileName%, General, buffpercent, 0.050000
        IniRead, compactmode, %cfgFileName%, General, compactmode, 0
        IniRead, runonstart, %cfgFileName%, General, runonstart, 1
        IniRead, gameresolution, %cfgFileName%, General, gameresolution, "Auto"
        IniRead, enableloothelper, %cfgFileName%, General, enableloothelper, 0
        IniRead, loothelpertimes, %cfgFileName%, General, loothelpertimes, 30
        IniRead, helpermousespeed, %cfgFileName%, General, helpermousespeed, 2
        IniRead, helperanimationdelay, %cfgFileName%, General, helperanimationdelay, 150
        IniRead, d3only, %cfgFileName%, General, d3only, 1
        IniRead, maxreforge, %cfgFileName%, General, maxreforge, 10
        generals:={"oldsandhelpermethod":oldsandhelpermethod, "oldsandhelperhk":oldsandhelperhk, "maxreforge":maxreforge
        , "enablesalvagehelper":enablesalvagehelper, "salvagehelpermethod":salvagehelpermethod, "reforgehelpermethod":reforgehelpermethod
        , "d3only":d3only, "enablereforgehelper":enablereforgehelper, "runonstart":runonstart, "gameresolution":gameresolution
        , "enablegamblehelper":enablegamblehelper, "gamblehelpertimes":gamblehelpertimes, "helpermousespeed":helpermousespeed
        , "startmethod":startmethod, "starthotkey":starthotkey, "enableupgradehelper":enableupgradehelper, "helperanimationdelay":helperanimationdelay
        , "enablesmartpause":enablesmartpause, "enablesoundplay":enablesoundplay, "enableconverthelper":enableconverthelper, "enableabandonhelper":enableabandonhelper
        , "custommoving":custommoving, "custommovinghk":custommovinghk, "customstanding":customstanding, "customstandinghk":customstandinghk
        , "custompotion":custompotion, "custompotionhk":custompotionhk
        , "safezone":safezone, "helperspeed":helperspeed, "gamegamma":gamegamma, "sendmode":sendmode, "buffpercent":buffpercent
        , "enableloothelper":enableloothelper, "loothelpertimes":loothelpertimes, "compactmode":compactmode}

        IniRead, tabs, %cfgFileName%
        tabs:=StrReplace(StrReplace(tabs, "`n", "`|"), "General|", "")
        combats:=[]
        others:=[]
        Loop, parse, tabs, `|
        {
            cSection:=A_LoopField
            trow:=[]
            tos:={}
            Loop, 6
            {
                IniRead, hk, %cfgFileName%, %cSection%, skill_%A_Index%, %A_Index%
                IniRead, ac, %cfgFileName%, %cSection%, action_%A_Index%, 1
                IniRead, iv, %cfgFileName%, %cSection%, interval_%A_Index%, 300
                IniRead, dy, %cfgFileName%, %cSection%, delay_%A_Index%, 10
                IniRead, rd, %cfgFileName%, %cSection%, random_%A_Index%, 1
                IniRead, pr, %cfgFileName%, %cSection%, priority_%A_Index%, 1
                IniRead, rp, %cfgFileName%, %cSection%, repeat_%A_Index%, 1
                IniRead, rpiv, %cfgFileName%, %cSection%, repeatinterval_%A_Index%, 30
                IniRead, tgbt, %cfgFileName%, %cSection%, triggerbutton_%A_Index%, LButton
                trow.Push({"hotkey":hk, "action":ac, "interval":iv, "delay":dy, "random":rd, "priority":pr, "repeat":rp, "repeatinterval":rpiv, "triggerbutton": tgbt})
            }
            combats.Push(trow)
            IniRead, pfmd, %cfgFileName%, %cSection%, profilehkmethod, 1
            IniRead, pfhk, %cfgFileName%, %cSection%, profilehkkey
            IniRead, pfmv, %cfgFileName%, %cSection%, movingmethod, 1
            IniRead, pfmi, %cfgFileName%, %cSection%, movinginterval, 100
            IniRead, pfpo, %cfgFileName%, %cSection%, potionmethod, 1
            IniRead, pfpi, %cfgFileName%, %cSection%, potioninterval, 500
            IniRead, pflm, %cfgFileName%, %cSection%, lazymode, 1
            IniRead, pfqp, %cfgFileName%, %cSection%, enablequickpause, 0
            IniRead, pfqpm1, %cfgFileName%, %cSection%, quickpausemethod1, 1
            IniRead, pfqpm2, %cfgFileName%, %cSection%, quickpausemethod2, 1
            IniRead, pfqpm3, %cfgFileName%, %cSection%, quickpausemethod3, 1
            IniRead, pfqpdy, %cfgFileName%, %cSection%, quickpausedelay, 1500
            IniRead, pfusq, %cfgFileName%, %cSection%, useskillqueue, 0
            IniRead, pfusqiv, %cfgFileName%, %cSection%, useskillqueueinterval, 200
            IniRead, pfasm, %cfgFileName%, %cSection%, autostartmarco, 0
            tos:={"profilemethod":pfmd, "profilehotkey":pfhk, "movingmethod":pfmv, "movinginterval":pfmi
            , "potionmethod":pfpo, "potioninterval":pfpi, "lazymode":pflm
            , "enablequickpause":pfqp, "quickpausemethod1":pfqpm1, "quickpausemethod2":pfqpm2, "quickpausemethod3":pfqpm3
            , "quickpausedelay":pfqpdy, "useskillqueue":pfusq, "useskillqueueinterval":pfusqiv, "autostartmarco":pfasm}
            others.Push(tos)
        }

    }
    Else
    {
        tabs=Конфигурация 1|Конфигурация 2|Конфигурация 3|Конфигурация 4
        currentProfile:=1
        combats:=[]
        others:=[]
        hks:="1,2,3,4,LButton,RButton"
        Loop, parse, tabs, `|
        {
            crow:=[]
            loop, parse, hks, CSV
            {
                crow.Push({"hotkey":A_LoopField, "action":1, "interval":300, "delay":10, "random": 1, "priority":1, "repeat":1, "repeatinterval":30, "triggerbutton": LButton})
            }
            combats.Push(crow)
            others.Push({"profilemethod":1, "profilehotkey":"", "movingmethod":1, "movinginterval":100
            , "potionmethod":1, "potioninterval":500, "lazymode":1
            , "enablequickpause":0, "quickpausemethod1":1, "quickpausemethod2":1, "quickpausemethod3":1, "quickpausedelay":1500
            , "useskillqueue":0, "useskillqueueinterval":200, "autostartmarco":0})
        }
        generals:={"enablegamblehelper":1 ,"gamblehelpertimes":15, "oldsandhelperhk":"F5", "d3only":1, "maxreforge":10
        , "startmethod":7, "starthotkey":"F2", "enablesmartpause":1, "salvagehelpermethod":1, "reforgehelpermethod":1
        , "oldsandhelpermethod":7, "enablesalvagehelper":0, "enablesoundplay":1, "enableconverthelper":0
        , "enablereforgehelper":0, "enableupgradehelper":0, "enableabandonhelper":0, "runonstart":1
        , "custommoving":0, "custommovinghk":"e", "customstanding":0, "customstandinghk":"LShift"
        , "custompotion":0, "custompotionhk":"q", "helpermousespeed":2
        , "safezone":"61,62,63", "helperspeed":3, "gamegamma":1.000000, "sendmode":"Event", "helperanimationdelay":150
        , "buffpercent":0.050000, "enableloothelper":0, "loothelpertimes":30, "compactmode":0, "gameresolution":"Auto"}
    }
    Return currentProfile
}

/*
Сохранение конфигурационного файла
Параметры:
    cfgFileName: имя файла
    tabs: String, имена конфигураций через "|"
    currentProfile: int, номер текущей активной конфигурации
    safezone: Array, защищённые ячейки
    VERSION: int, версия
Возврат:
    нет
*/
SaveCfgFile(cfgFileName, tabs, currentProfile, safezone, VERSION){
    createOrTruncateFile(cfgFileName)

    GuiControlGet, extraGambleHelperCKbox
    GuiControlGet, extraGambleHelperUpdown
    GuiControlGet, helperKeybindingdropdown
    GuiControlGet, helperKeybindingHK  
    GuiControlGet, extraLootHelperCkbox
    GuiControlGet, extraLootHelperUpdown
    GuiControlGet, extraSmartPause
    GuiControlGet, extraSalvageHelperCkbox
    GuiControlGet, extraSalvageHelperDropdown
    GuiControlGet, extraReforgeHelperCkbox
    GuiControlGet, extraReforgeHelperDropdown
    GuiControlGet, extraConvertHelperCkbox
    GuiControlGet, extraUpgradeHelperCkbox
    GuiControlGet, extraSoundonProfileSwitch
    GuiControlGet, extraAbandonHelperCkbox
    GuiControlGet, extraCustomMoving
    GuiControlGet, extraCustomMovingHK
    GuiControlGet, extraCustomStanding
    GuiControlGet, extraCustomStandingHK
    GuiControlGet, extraCustomPotion
    GuiControlGet, extraCustomPotionHK
    GuiControlGet, helperAnimationSpeedDropdown

    IniWrite, %VERSION%, %cfgFileName%, General, version
    IniWrite, %currentProfile%, %cfgFileName%, General, activatedprofile
    IniWrite, %extraGambleHelperCKbox%, %cfgFileName%, General, enablegamblehelper
    IniWrite, %extraGambleHelperUpdown%, %cfgFileName%, General, gamblehelpertimes
    IniWrite, %extraSmartPause%, %cfgFileName%, General, enablesmartpause
    IniWrite, %extraSalvageHelperCkbox%, %cfgFileName%, General, enablesalvagehelper
    IniWrite, %extraSalvageHelperDropdown%, %cfgFileName%, General, salvagehelpermethod
    IniWrite, %extraReforgeHelperCkbox%, %cfgFileName%, General, enablereforgehelper
    IniWrite, %extraReforgeHelperDropdown%, %cfgFileName%, General, reforgehelpermethod
    Global maxreforge
    IniWrite, %maxreforge%, %cfgFileName%, General, maxreforge
    IniWrite, %extraUpgradeHelperCkbox%, %cfgFileName%, General, enableupgradehelper
    IniWrite, %extraConvertHelperCkbox%, %cfgFileName%, General, enableconverthelper
    IniWrite, %extraAbandonHelperCkbox%, %cfgFileName%, General, enableabandonhelper
    IniWrite, %extraLootHelperCkbox%, %cfgFileName%, General, enableloothelper
    IniWrite, %extraLootHelperUpdown%, %cfgFileName%, General, loothelpertimes
    IniWrite, %extraSoundonProfileSwitch%, %cfgFileName%, General, enablesoundplay
    IniWrite, %helperKeybindingHK%, %cfgFileName%, General, oldsandhelperhk
    IniWrite, %helperKeybindingdropdown%, %cfgFileName%, General, oldsandhelpermethod
    IniWrite, %extraCustomMoving%, %cfgFileName%, General, custommoving
    IniWrite, %extraCustomMovingHK%, %cfgFileName%, General, custommovinghk
    IniWrite, %extraCustomStanding%, %cfgFileName%, General, customstanding
    IniWrite, %extraCustomStandingHK%, %cfgFileName%, General, customstandinghk
    IniWrite, %extraCustomPotion%, %cfgFileName%, General, custompotion
    IniWrite, %extraCustomPotionHK%, %cfgFileName%, General, custompotionhk
    IniWrite, %helperAnimationSpeedDropdown%, %cfgFileName%, General, helperspeed
    safezone:=keyJoin(",", safezone)
    IniWrite, %safezone%, %cfgFileName%, General, safezone
    Global gameGamma, buffpercent, isCompact, runOnStart, gameResolution, helperAnimationDelay, helperMouseSpeed, d3only
    IniWrite, %d3only%, %cfgFileName%, General, d3only
    IniWrite, %gameGamma%, %cfgFileName%, General, gamegamma
    IniWrite, %A_SendMode%, %cfgFileName%, General, sendmode
    IniWrite, %buffpercent%, %cfgFileName%, General, buffpercent
    IniWrite, %isCompact%, %cfgFileName%, General, compactmode
    IniWrite, %runOnStart%, %cfgFileName%, General, runonstart
    IniWrite, %gameResolution%, %cfgFileName%, General, gameresolution
    IniWrite, %helperAnimationDelay%, %cfgFileName%, General, helperanimationdelay
    IniWrite, %helperMouseSpeed%, %cfgFileName%, General, helpermousespeed
    
    GuiControlGet, StartRunDropdown
    GuiControlGet, StartRunHKInput
    IniWrite, %StartRunDropdown%, %cfgFileName%, General, startmethod
    IniWrite, %StartRunHKInput%, %cfgFileName%, General, starthotkey
    global combats
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
            GuiControlGet, skillset%cSection%s%A_Index%randomckbox
            pr:=combats[cSection][A_Index]["priority"]
            rp:=combats[cSection][A_Index]["repeat"]
            rpiv:=combats[cSection][A_Index]["repeatinterval"]
            tgbt:=combats[cSection][A_Index]["triggerbutton"]
            IniWrite, % skillset%cSection%s%A_Index%dropdown, %cfgFileName%, %nSction%, action_%A_Index%
            IniWrite, % skillset%cSection%s%A_Index%updown, %cfgFileName%, %nSction%, interval_%A_Index%
            IniWrite, % skillset%cSection%s%A_Index%delayupdown, %cfgFileName%, %nSction%, delay_%A_Index%
            IniWrite, % skillset%cSection%s%A_Index%randomckbox, %cfgFileName%, %nSction%, random_%A_Index%
            IniWrite, % pr, %cfgFileName%, %nSction%, priority_%A_Index%
            IniWrite, % rp, %cfgFileName%, %nSction%, repeat_%A_Index%
            IniWrite, % rpiv, %cfgFileName%, %nSction%, repeatinterval_%A_Index%
            IniWrite, % tgbt, %cfgFileName%, %nSction%, triggerbutton_%A_Index%
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
        GuiControlGet, skillset%cSection%potiondropdown
        GuiControlGet, skillset%cSection%potionupdown
        IniWrite, % skillset%cSection%potiondropdown, %cfgFileName%, %nSction%, potionmethod
        IniWrite, % skillset%cSection%potionupdown, %cfgFileName%, %nSction%, potioninterval
        GuiControlGet, skillset%cSection%profilestartmodedropdown
        IniWrite, % skillset%cSection%profilestartmodedropdown, %cfgFileName%, %nSction%, lazymode
        GuiControlGet, skillset%cSection%clickpauseckbox
        IniWrite, % skillset%cSection%clickpauseckbox, %cfgFileName%, %nSction%, enablequickpause
        GuiControlGet, skillset%cSection%clickpausedropdown1
        IniWrite, % skillset%cSection%clickpausedropdown1, %cfgFileName%, %nSction%, quickpausemethod1
        GuiControlGet, skillset%cSection%clickpausedropdown2
        IniWrite, % skillset%cSection%clickpausedropdown2, %cfgFileName%, %nSction%, quickpausemethod2
        GuiControlGet, skillset%cSection%clickpausedropdown3
        IniWrite, % skillset%cSection%clickpausedropdown3, %cfgFileName%, %nSction%, quickpausemethod3
        GuiControlGet, skillset%cSection%clickpauseupdown
        IniWrite, % skillset%cSection%clickpauseupdown, %cfgFileName%, %nSction%, quickpausedelay
        GuiControlGet, skillset%cSection%useskillqueueckbox
        IniWrite, % skillset%cSection%useskillqueueckbox, %cfgFileName%, %nSction%, useskillqueue
        GuiControlGet, skillset%cSection%useskillqueueedit
        IniWrite, % skillset%cSection%useskillqueueedit, %cfgFileName%, %nSction%, useskillqueueinterval
        GuiControlGet, skillset%cSection%autostartmarcockbox
        IniWrite, % skillset%cSection%autostartmarcockbox, %cfgFileName%, %nSction%, autostartmarco
    }
    Return
}

/*
Возвращает координаты левого края полоски баффа скилла для текущего разрешения
Параметры:
    D3W: int, ширина окна
    D3H: int, высота окна
    buttonID: int, ID кнопки (1 - самый левый, 6 - правый клик)
    percent: float, процент от ширины бафф-полоски (слева)
Возврат:
    [x, y]
*/
getSkillButtonBuffPos(D3W, D3H, buttonID, percent){
    static x:=[1288, 1377, 1465, 1554, 1647, 1734]
    static w:=63
    y:=1328*D3H/1440
    Return [Round(D3W/2-(3440/2-x[buttonID]-percent*w)*D3H/1440), Round(y)]
}

/*
Разбирает HEX-цвет в RGB массив. FFFFFF -> [255, 255, 255].
Если gameGamma != 1, пытается скорректировать гамму.
Параметры:
    vthiscolor: HEX-цвет от PixelGetColor
Возврат:
    [R, G, B]
*/
splitRGB(vthiscolor){
    local
    Global gameGamma
    vblue:=(vthiscolor & 0xFF)
    vgreen:=((vthiscolor & 0xFF00) >> 8)
    vred:=((vthiscolor & 0xFF0000) >> 16)
    if (Abs(gameGamma-1)>=0.01)
    {
        vblue:=((vblue / 255) ** (1.75*gameGamma-0.75)) * 255
        vgreen:=((vgreen / 255) ** (1.9*gameGamma-0.9)) * 255
        vred:=((vred / 255) ** (1.9*gameGamma-0.9)) * 255
    }
    Return [vred, vgreen, vblue]
}

/*
Отправляет нажатие клавиши умения
Параметры:
    currentProfile: int, активная конфигурация
    nskill: int, номер умения 1-6
    D3W, D3H: int, размеры окна
    forceStandingKey: клавиша принудительного стояния
    useSkillQueue: Bool, использовать очередь
Возврат:
    нет
*/
skillKey(currentProfile, nskill, D3W, D3H, forceStandingKey, useSkillQueue){
    local
    Global vPausing, vRunning, skillQueue, buffpercent, gameX, gameY, syncTimer, syncDelay
    global combats
    GuiControlGet, skillset%currentProfile%s%nskill%hotkey
    GuiControlGet, skillset%currentProfile%s%nskill%delayupdown
    GuiControlGet, skillset%currentProfile%s%nskill%randomckbox
    GuiControlGet, skillset%currentProfile%s%nskill%updown
    Loop, 6
    {
        GuiControlGet, skillset%currentProfile%s%A_Index%dropdown
        ; Проверяем стратегию других кнопок
        if (A_Index = nskill){
            Continue
        }
        ; Если другая кнопка в режиме "Поддержание баффа" и приоритет выше
        if (skillset%currentProfile%s%A_Index%dropdown = 4 and combats[currentProfile][A_Index]["priority"]>combats[currentProfile][nskill]["priority"])
        {
            ; Проверяем, активен ли бафф
            magicXY:=getSkillButtonBuffPos(D3W, D3H, A_Index, buffpercent)
            crgb:=getPixelRGB(magicXY)
            ; Если активен, выходим
            if (crgb[2]>=95) {
                Return
            }
        }
    }
    k:=skillset%currentProfile%s%nskill%hotkey
    switch skillset%currentProfile%s%nskill%dropdown
    {
        ; Автоповтор, по нажатию
        case 3,5:
            if !(vPausing) and vRunning
            {
                if (abs(skillset%currentProfile%s%nskill%delayupdown)>20)
                {
                    if (skillset%currentProfile%s%nskill%randomckbox)
                    {
                        Random, delay, 10, abs(skillset%currentProfile%s%nskill%delayupdown)
                    }
                    Else
                    {
                        delay:=abs(skillset%currentProfile%s%nskill%delayupdown)
                    }
                    syncDelay[nskill]:=delay
                    if (skillset%currentProfile%s%nskill%delayupdown<0)
                    {
                        syncDelay[nskill]:=skillset%currentProfile%s%nskill%updown - delay
                    }
                    syncTimer[nskill]:=A_TickCount
                    while (A_TickCount - syncTimer[nskill] <= syncDelay[nskill])
                    {
                        sleep 10
                    }
                }
                ; Повтор нажатий
                Loop,% combats[currentProfile][nskill]["repeat"]
                {
                    if useSkillQueue
                    {
                        if (skillQueue.Count() < 1000){
                            ; Добавляем в начало очереди
                            ; [k, 3] -> 3 означает добавление из-за автоповтора
                            skillQueue.InsertAt(1, [k, 3])
                        }
                    }
                    Else
                    {
                        Send {Blind}{%k%}
                    }
                    if (combats[currentProfile][nskill]["repeat"] > 1){
                        sleep, % combats[currentProfile][nskill]["repeatinterval"]
                    }

                }
            }
        ; Поддержание баффа
        case 4:
            if !(vPausing) and vRunning
            {
                ; Координаты левого края баффа
                magicXY:=getSkillButtonBuffPos(D3W, D3H, nskill, buffpercent)
                crgb:=getPixelRGB(magicXY)
                ; Нужно ли обновлять бафф
                if (crgb[2]<95)
                {
                    switch nskill
                    {
                        case 5:
                            ; Левая кнопка мыши
                            if useSkillQueue
                            {
                                if (skillQueue.Count() < 1000){
                                    ; 4 - добавление из-за баффа
                                    skillQueue.Push([k, 4])
                                }
                            }
                            Else
                            {
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
                                if (skillQueue.Count() < 1000){
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
    }
    Return
}

/*
Создаёт или очищает конфигурационный файл и записывает заголовок
Параметры:
    FileName: имя файла
Возврат:
    нет
*/
createOrTruncateFile(FileName){
    if (FileName = "")
    {
        return
    }
    file:=FileOpen(FileName, "w", "UTF-16")
    if !IsObject(file)
    {
        MsgBox Не удалось создать или записать файл: "%FileName%"
        return
    }
    file.Write("; ===============================================`r`n")
    file.Write("; Добро пожаловать в конфигурационный файл D3KeyHelper (Oldsand).`r`n")
    file.Write("; Каждая секция, кроме General, соответствует отдельной конфигурации клавиш.`r`n")
    file.Write("; Вы можете добавлять или удалять их по своему усмотрению.`r`n")
    file.Write("; ===============================================`r`n")
    file.Close()
}

/*
Запускает макрос помощника
Параметры:
    нет
Возврат:
    нет
*/
oldsandHelper(){
    local
    Global helperRunning, helperBreak, helperDelay, mouseDelay, vRunning, helperAnimationDelay, helperMouseSpeed, gameX, gameY
    if helperRunning{
        helperBreak:=True
        helperRunning:=False
        Sleep, 200
        Return
    }
    if (vRunning or !getGameResulution(D3W, D3H)){
        Return
    }
    gameXY:=getGameXYonScreen(0,0)
    gameX:=gameXY[1]
    gameY:=gameXY[2]
    helperRunning:=True
    helperBreak:=False
    GuiControlGet, extraGambleHelperCKbox
    GuiControlGet, extraLootHelperCkbox
    GuiControlGet, extraSalvageHelperCkbox
    GuiControlGet, extraReforgeHelperCkbox
    GuiControlGet, extraUpgradeHelperCkbox
    GuiControlGet, extraConvertHelperCkbox
    GuiControlGet, extraAbandonHelperCkbox
    GuiControlGet, extraSalvageHelperDropdown
    GuiControlGet, helperAnimationSpeedDropdown
    MouseGetPos, xpos, ypos ; текущая позиция мыши для возврата
    switch helperAnimationSpeedDropdown
    {
        case 1:
            mouseDelay:=0
            helperDelay:=50
        case 2:
            mouseDelay:=1
            helperDelay:=100
        case 3:
            mouseDelay:=2
            helperDelay:=150
        case 4:
            mouseDelay:=3
            helperDelay:=200
        Default:
            mouseDelay:=helperMouseSpeed
            helperDelay:=helperAnimationDelay
    }
    SetDefaultMouseSpeed, mouseDelay
    mousePosition:=-1
    if (xpos>D3W-(3440-2740)*D3H/1440 and ypos>730*D3H/1440 and ypos<1150*D3H/1440)
    {
        mousePosition:=1
    }
    else if (xpos>65*D3H/1440 and xpos<640*D3H/1440 and ypos>275*D3H/1440 and ypos<1150*D3H/1440)
    {
        mousePosition:=2
    }
    if (xpos<680*D3H/1440)
    {
        if (extraGambleHelperCKbox and isGambleOpen(D3W, D3H))
        {
            SetTimer, gambleHelper, -1
            Return
        }
    }

    if (extraSalvageHelperCkbox)
    {
        r:=isSalvagePageOpen(D3W, D3H)
        switch r[1]
        {
            case 2:
                if(extraSalvageHelperDropdown=1)
                {
                    if(mousePosition = 1)
                    {
                        quickSalvageHelper(D3W, D3H, helperDelay)
                        helperRunning:=False
                    }
                }
                Else
                {
                    salvageIconXY:=getSalvageIconXY(D3W, D3H, "center")
                    MouseMove, salvageIconXY[1][1], salvageIconXY[1][2]
                    if (r[2][3]<10 and r[2][1]+r[2][2]>400)
                    {
                        if helperBreak
                        {
                            helperRunning:=False
                            Return
                        }
                        Click, Right
                        Sleep, helperDelay
                        p:=getSalvageIconXY(D3W, D3H, "edge")
                        r[3]:=getPixelRGB(p[2])
                        r[4]:=getPixelRGB(p[3])
                        r[5]:=getPixelRGB(p[4])
                    }
                    _wait:=-1
                    for i, _c in [r[5][1]>50, r[4][3]>65, r[3][1]>65]
                    {
                        if _c
                        {
                            if helperBreak
                            {
                                helperRunning:=False
                                Return
                            }
                            _wait:=-helperDelay-50
                            MouseMove, salvageIconXY[5-i][1], salvageIconXY[5-i][2]
                            Click
                            Sleep, helperDelay
                            Send {Enter}
                        }
                    }
                    MouseMove, salvageIconXY[1][1], salvageIconXY[1][2]
                    Sleep, helperDelay//2
                    Click
                    if helperBreak
                    {
                        helperRunning:=False
                        Return
                    }
                    Sleep, helperDelay//2
                    fn:=Func("oneButtonSalvageHelper").Bind(D3W, D3H, xpos, ypos)
                    SetTimer, %fn%, %_wait%
                }
                Return
            case 1:
                helperRunning:=False
                Return
            Default:
        }
    }
    if (extraReforgeHelperCkbox or extraUpgradeHelperCkbox or extraConvertHelperCkbox)
    {
        switch isKanaiCubeOpen(D3W, D3H)
        {
            case 1:
                helperRunning:=False
                Return
            case 2:
                if (extraReforgeHelperCkbox and mousePosition=1)
                {
                    fn:=Func("oneButtonReforgeHelper").Bind(D3W, D3H, xpos, ypos)
                    SetTimer, %fn%, -1
                    Return
                }
            case 3:
                if extraUpgradeHelperCkbox
                {
                    fn:=Func("oneButtonUpgradeConvertHelper").Bind(D3W, D3H, xpos, ypos)
                    SetTimer, %fn%, -1
                    Return
                }
            case 4:
                if extraConvertHelperCkbox
                {
                    fn:=Func("oneButtonUpgradeConvertHelper").Bind(D3W, D3H, xpos, ypos)
                    SetTimer, %fn%, -1
                    Return
                }
            Default:
        }
    }
    if (extraAbandonHelperCkbox and mousePosition>0 and isInventoryOpen(D3W, D3H))
    {
        fn:=Func("oneButtonAbandonHelper").Bind(D3W, D3H, xpos, ypos, mousePosition)
        SetTimer, %fn%, -1
        Return
    }
    if (extraLootHelperCkbox)
    {
        fn:=Func("lootHelper").Bind(D3W, D3H, helperDelay)
        SetTimer, %fn%, -1
        Return
    }
}

/*
oneButtonReforgeHelper - рефорж в Кубе
Параметры:
    D3W, D3H: размеры окна
    xpos, ypos: исходная позиция мыши
*/
oneButtonReforgeHelper(D3W, D3H, xpos, ypos){
    local
    Global helperRunning, helperBreak, helperDelay, mouseDelay, maxreforge
    GuiControlGet, extraReforgeHelperDropdown
    SetDefaultMouseSpeed, mouseDelay
    kanai:=getKanaiCubeButtonPos(D3W, D3H)
    box_1_1:=getInventorySpaceXY(D3W, D3H, 1, "kanai")
    box_1_2:=getInventorySpaceXY(D3W, D3H, 2, "kanai")
    Loop, %maxreforge% {
        q:=0
        Click, Right
        Sleep, helperDelay//4
        MouseMove, kanai[2][1], kanai[2][2]
        Click
        Sleep, helperDelay//4
        MouseMove, kanai[1][1], kanai[1][2]
        Click
        Sleep, helperDelay//4
        MouseMove, kanai[3][1], kanai[3][2]
        Click
        Sleep, helperDelay//4
        MouseMove, kanai[4][1], kanai[4][2]
        Click
        if (extraReforgeHelperDropdown > 1 and not helperBreak)
        {
            MouseMove, xpos, ypos
            Click, Right
            MouseMove, box_1_1[1], box_1_1[2]
            Sleep, helperDelay//2
            c_t:=[-255,-255,-255]
            StartTime1:=A_TickCount
            while (A_TickCount-StartTime1<=helperDelay)
            {
                c:=getPixelsRGB(box_1_2[3]+1, box_1_2[2], 3, 1, "Max", False)
                if (c_t[1]=c[1] and c_t[2]=c[2] and c_t[3]=c[3]){
                    Break
                }
                c_t:=c
                Sleep, 20
            }
            if ((c[1]>=70 or c[3]<=20) and Max(Abs(c[1]-c[2]), Abs(c[1]-c[3]), Abs(c[3]-c[2]))>20 and (c[1]+c[2]+c[3]<460)) {
                q:=(c[2]<35) ? 5:3
            } else if (c[3]>100 and c[3]>c[2] and c[2]>c[1]) {
                q:=4
            } else {
                q:=2
            }
            MouseMove, xpos, ypos
            if (q > extraReforgeHelperDropdown)
            {
                Break
            }
        }
        Else
        {
            Break
        }
    }
    MouseMove, xpos, ypos
    helperRunning:=False
    Return
}

/*
oneButtonUpgradeConvertHelper - апгрейд или конвертация в Кубе
Параметры:
    D3W, D3H: размеры окна
    xpos, ypos: исходная позиция мыши
*/
oneButtonUpgradeConvertHelper(D3W, D3H, xpos, ypos)
{
    local
    Global helperBreak, helperRunning, helperDelay, helperBagZone, mouseDelay
    helperBagZone:=make1DArray(60, -1)
    k:=getKanaiCubeButtonPos(D3W, D3H)
    fn1:=Func("scanInventorySpaceGDIP").Bind(D3W, D3H)
    SetTimer, %fn1%, -1

    SetDefaultMouseSpeed, mouseDelay
    i:=1
    w:=0
    while (i<=60)
    {
        w++
        if (helperBreak or w>200) {
            Break
        }
        switch helperBagZone[i]
        {
            case -1:
                Sleep, 20
            case 10:
                pLargeItem:=False
                m:=getInventorySpaceXY(D3W, D3H, i, "bag")
                m2:=getInventorySpaceXY(D3W, D3H, i+10, "bag")
                MouseMove, m[1], m[2]
                Click, Right
                if (i<=50 and (helperBagZone[i+10]=-1 or helperBagZone[i+10]=10))
                {
                    pLargeItem:=True
                    cd_before:=getPixelRGB(m2)
                }
                Sleep, helperDelay
                MouseMove, k[2][1], k[2][2]
                Click
                Sleep, helperDelay+50
                MouseMove, k[1][1], k[1][2]
                Click
                Sleep, helperDelay+50
                MouseMove, k[4][1], k[4][2]
                Click
                Sleep, helperDelay+50
                MouseMove, k[3][1], k[3][2]
                Click
                Sleep, helperDelay+50
                if (pLargeItem)
                {
                    cd_after:=getPixelRGB(m2)
                    if !isArraysEqual(cd_before, cd_after, 3)
                    {
                        helperBagZone[i+10]:=5
                    }
                }
                i++
            Default:
                i++
        }
    }
    helperRunning:=False
    MouseMove, xpos, ypos
    Return
}

/*
gambleHelper - клики ПКМ у Кадалы
*/
gambleHelper(){
    local
    Global helperDelay, helperBreak, helperRunning
    GuiControlGet, extraGambleHelperEdit
    Loop, %extraGambleHelperEdit%
    {
        if helperBreak{
            Break
        }
        Click, Right
        Sleep, helperDelay//4
    }
    helperRunning:=False
    Return
}

/*
lootHelper - сбор лута (клики ЛКМ)
*/
lootHelper(D3W, D3H, helperDelay){
    local
    Global helperBreak, helperRunning
    MouseGetPos, xpos, ypos
    if (Abs(xpos - D3W/2)<600*1440/D3H and Abs(ypos - D3H/2)<500*1440/D3H)
    {
        GuiControlGet, extraLootHelperEdit
        Loop, %extraLootHelperEdit%
        {
            if helperBreak{
                Break
            }
            Click
            Sleep, helperDelay//2
        }
    }
    Else
    {
        Click
    }
    helperRunning:=False
    Return
}

/*
quickSalvageHelper - быстрое разбирание (ЛКМ + Enter)
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
oneButtonSalvageHelper - умное разбирание инвентаря
*/
oneButtonSalvageHelper(D3W, D3H, xpos, ypos){
    local
    static _spaceSizeInnerH:=63
    static _spaceSizeInnerW:=64
    Global helperBreak, helperRunning, helperDelay, helperBagZone, mouseDelay, cInventorySpace
    helperBagZone:=make1DArray(60, -1)
    fn1:=Func("scanInventorySpaceGDIP").Bind(D3W, D3H)
    SetTimer, %fn1%, -1

    q:=0
    i:=1
    w:=0
    SetDefaultMouseSpeed, mouseDelay
    GuiControlGet, extraSalvageHelperDropdown
    while (i<=60)
    {
        w++
        if (helperBreak or w>200) {
            Break
        }
        switch helperBagZone[i]
        {
            case -1:
                Sleep, 20
            case 10:
                m:=getInventorySpaceXY(D3W, D3H, i, "bag")
                MouseMove, m[1], m[2]
                if (extraSalvageHelperDropdown > 2)
                {
                    c_t:=[-255,-255,-255]
                    StartTime1:=A_TickCount
                    while (A_TickCount-StartTime1<=helperDelay)
                    {
                        c:=getPixelsRGB(Round(m[3]-1-10*D3H/1440), m[2], 3, 1, "Max", False)
                        if (c_t[1]=c[1] and c_t[2]=c[2] and c_t[3]=c[3]){
                            Break
                        }
                        c_t:=c
                        Sleep, 20
                    }
                    if ((c[1]>=70 or c[3]<=20) and Max(Abs(c[1]-c[2]), Abs(c[1]-c[3]), Abs(c[3]-c[2]))>20 and (c[1]+c[2]+c[3]<410)) {
                        q:=(c[2]<35) ? 5:3
                    } else if (c[3]>100 and c[3]>c[2] and c[2]>c[1]) {
                        q:=4
                    } else if (c[1]<50 and c[2]>c[3] and c[3]>c[1]) {
                        q:=4
                    } else {
                        q:=2
                    }
                }
                if (i<=50 and (helperBagZone[i+10]=10 or helperBagZone[i+10]=-1))
                {
                    md:=getInventorySpaceXY(D3W, D3H, i+10, "bag")
                    c_b:=cInventorySpace[i+10]
                    c_t:=[-255,-255,-255]
                    StartTime1:=A_TickCount
                    while (A_TickCount-StartTime1<=helperDelay)
                    {
                        c_a:=getPixelRGB([Round(md[3]+_spaceSizeInnerW*0.08*D3H/1440), Round(md[4]+_spaceSizeInnerH*0.7*D3H/1440)])
                        if (c_t[1]=c_a[1] and c_t[2]=c_a[2] and c_t[3]=c_a[3]){
                            Break
                        }
                        c_t:=c_a
                    }
                    if !(c_b[1]=c_a[1] and c_b[2]=c_a[2] and c_b[3]=c_a[3]){
                        helperBagZone[i+10]:=5
                    }
                }
                if (q>=extraSalvageHelperDropdown) {
                    i++
                    Continue
                }
                Click
                StartTime1:=A_TickCount
                while (A_TickCount-StartTime1<=helperDelay)
                {
                    if isDialogBoXOnScreen(D3W, D3H)
                    {
                        Sleep, helperDelay//4
                        Send {Enter}
                        StartTime2:=A_TickCount
                        while (A_TickCount-StartTime2<=2*helperDelay)
                        {
                            if isInventorySpaceEmpty(D3W, D3H, i, "", "bag")
                            {
                                if ((helperBagZone[i+10]=10 or helperBagZone[i+10]=-1) and isInventorySpaceEmpty(D3W, D3H, i+10, "", "bag"))
                                {
                                    helperBagZone[i+10]:=5
                                }
                                Break
                            }
                        }
                        Break
                    }
                }
                i++
            Default:
                i++
        }
    }
    helperRunning:=False
    Click, Right
    MouseMove, xpos, ypos
    Return
}

/*
oneButtonAbandonHelper - выброс предметов или складирование в сундук
*/
oneButtonAbandonHelper(D3W, D3H, xpos, ypos, mousePosition){
    local
    Global helperBreak, helperRunning, helperDelay, helperBagZone, mouseDelay, forceStandingKey
    helperBagZone:=make1DArray(60, -1)
    fn1:=Func("scanInventorySpaceGDIP").Bind(D3W, D3H)
    SetTimer, %fn1%, -1
    SetDefaultMouseSpeed, mouseDelay
    stashOpen:=-1
    i:=1
    w:=0
    while (i<=60)
    {
        w++
        if (helperBreak or w>200) {
            Break
        }
        switch helperBagZone[i]
        {
            case -1:
                Sleep, 20
            case 10:
                m:=getInventorySpaceXY(D3W, D3H, i, "bag")
                MouseMove, m[1], m[2]
                if (stashOpen=-1)
                {
                    Sleep, helperDelay//2
                    stashOpen:=isStashOpen(D3W, D3H)
                    if (stashOpen=0 and mousePosition!=1)
                    {
                        Break
                    }
                }
                if (mousePosition=1)
                {
                    Click
                    Sleep, helperDelay//2
                    MouseMove, D3W//2, D3H//2
                    if GetKeyState(forceStandingKey)
                    {
                        Click
                    }
                    Else
                    {
                        Send {%forceStandingKey% down}
                        Click
                        Send {%forceStandingKey% up}
                    }
                }
                Else
                {
                    Click, Right
                    Sleep, helperDelay//2
                }
                if (i<=50 and (helperBagZone[i+10]=10 or helperBagZone[i+10]=-1))
                {
                    StartTime2:=A_TickCount
                    while (A_TickCount-StartTime2<=helperDelay)
                    {
                        if isInventorySpaceEmpty(D3W, D3H, i+10, "", "bag")
                        {
                            helperBagZone[i+10]:=5
                            Break
                        }
                    }
                }
                i++
            Default:
                i++
        }
    }
    helperRunning:=False
    MouseMove, xpos, ypos
    Return
}

/*
potionHelper - авто-зелье
*/
potionHelper(action){
    local
    Global vPausing, potionKey, gameX, gameY, lastpotion, D3W, D3H
    static _x := 1822
    static _y := 1340
    static _w := 66
    if !vPausing
    {
        switch action
        {
            case 2:
                Send {%potionKey%}
            case 3:
                currentpotion:=getPixelsRGB(Round(D3W/2-(3440/2-1822)*D3H/1440), Round(_y*D3H/1440), Round(_w*D3H/1440), Round(_w*D3H/1440), "", True, gameX, gameY)
                if (lastpotion and isArraysEqual(lastpotion, currentpotion[1], 0)) {
                    Send {%potionKey%}
                }
                lastpotion:=currentpotion[1]
        }
    }
    Return
}

/*
scanInventorySpaceGDIP - сканирование ячеек инвентаря через GDI+
*/
scanInventorySpaceGDIP(D3W, D3H){
    local
    static _spaceSizeInnerW:=64
    static _spaceSizeInnerH:=63
    sxy:=getGameXYonScreen(0, 0)
    pInventoryBitmap:=Gdip_BitmapFromScreen(Format("{}|{}|{}|{}", sxy[1], sxy[2], D3W, D3H))
    Gdip_LockBits(pInventoryBitmap, 0, 0, Gdip_GetImageWidth(pInventoryBitmap), Gdip_GetImageHeight(pInventoryBitmap), Stride, Scan0, BitmapData)
    static _e:=[[0.65625,0.71429], [0.375,0.36508], [0.725,0.251]]
    Global safezone, helperBagZone, cInventorySpace
    cInventorySpace:={}
    Loop, 60
    {
        m:=getInventorySpaceXY(D3W, D3H, A_Index, "bag")
        cInventorySpace[A_Index]:=splitRGB(Gdip_GetLockBitPixel(Scan0, Round(m[3]+_spaceSizeInnerW*0.08*D3H/1440), Round(m[4]+_spaceSizeInnerH*0.7*D3H/1440), Stride))
        if safezone.HasKey(A_Index)
        {
            helperBagZone[A_Index]:=0
        }
        Else
        {
            if (helperBagZone[A_Index]!=-1){
                continue
            }
            r:=1
            for i, p in _e
            {
                xy:=[Round(m[3]+_spaceSizeInnerW*_e[i][1]*D3H/1440), Round(m[4]+_spaceSizeInnerH*_e[i][2]*D3H/1440)]
                c:=splitRGB(Gdip_GetLockBitPixel(Scan0, xy[1], xy[2], Stride))
                if !(c[1]<22 and c[2]<20 and c[3]<15 and c[1]>c[3] and c[2]>c[3])
                {
                    r:=10
                    Break
                }
            }
            helperBagZone[A_Index]:=r
        }
    }
    Gdip_UnlockBits(pInventoryBitmap, BitmapData)
    Gdip_DisposeImage(pInventoryBitmap)
    Return
}

/*
clickPauseMarco - быстрая пауза
*/
clickPauseMarco(pausetime, pauseAction){
    local
    Global vRunning, forceStandingKey, keysOnHold, quickPauseHK
    if vRunning
    {
        Gosub, StopMarco
        if (pausetime>0)
        {
            SetTimer, RunMarco, off
            SetTimer, RunMarco, -%pausetime%
            if (pauseAction=2)
            {
                startTime:=A_TickCount
                while (A_TickCount-startTime<pausetime)
                {
                    if GetKeyState(forceStandingKey)
                    {
                        Send {%forceStandingKey% up}
                        Click
                        Send {%forceStandingKey% down}
                    }
                    Else
                    {
                        Click
                    }
                    Sleep, 50
                }
            }
        }
        Else
        {
            Loop, 1000
            {
                if (pauseAction=2)
                {
                    if GetKeyState(forceStandingKey)
                    {
                        Send {%forceStandingKey% up}
                        Click
                        Send {%forceStandingKey% down}
                    }
                    Else
                    {
                        Click
                    }
                }
                Sleep, 50
                if !GetKeyState(quickPauseHK, "P")
                {
                    Break
                }
            }
            SetTimer, RunMarco, -1
        }
    }
    Return
}

SetStartMode(){
    local
    Global currentProfile
    GuiControlGet, skillset%currentProfile%profilestartmodedropdown
    switch skillset%currentProfile%profilestartmodedropdown
    {
        case 2:
            GuiControl, , skillset%currentProfile%clickpauseckbox, 0
            GuiControl, Disable, skillset%currentProfile%clickpauseckbox
            GuiControl, Enable, skillset%currentProfile%useskillqueueckbox
            GuiControl, Enable, skillset%currentProfile%movingdropdown
            GuiControl, Enable, skillset%currentProfile%potiondropdown
        case 3:
            GuiControl, , skillset%currentProfile%useskillqueueckbox, 0
            GuiControl, Choose, skillset%currentProfile%movingdropdown, 1
            GuiControl, Choose, skillset%currentProfile%potiondropdown, 1
            GuiControl, , skillset%currentProfile%clickpauseckbox, 0
            GuiControl, Disable, skillset%currentProfile%useskillqueueckbox
            GuiControl, Disable, skillset%currentProfile%movingdropdown
            GuiControl, Disable, skillset%currentProfile%potiondropdown
            GuiControl, Disable, skillset%currentProfile%clickpauseckbox
            WinSet, Redraw,, A
        Default:
            GuiControl, Enable, skillset%currentProfile%useskillqueueckbox
            GuiControl, Enable, skillset%currentProfile%movingdropdown
            GuiControl, Enable, skillset%currentProfile%clickpauseckbox
    }
    Gosub, SetQuickPause
    Gosub, SetMovingHelper
    Return
}

SetSkillQueueWarning(){
    local
    Global currentProfile
    GuiControlGet, skillset%currentProfile%useskillqueueckbox
    if skillset%currentProfile%useskillqueueckbox
    {
        GuiControlGet, skillset%currentProfile%useskillqueueupdown
        _out:=1000/skillset%currentProfile%useskillqueueupdown
        _in:=0
        Loop, 6
        {
            GuiControlGet, skillset%currentProfile%s%A_Index%dropdown
            if (skillset%currentProfile%s%A_Index%dropdown==3)
            {
                GuiControlGet, skillset%currentProfile%s%A_Index%updown
                _in+=1000/skillset%currentProfile%s%A_Index%updown
            }
        }
        if (_in>_out)
        {
            GuiControl, Show, skillset%currentProfile%skillqueuewarningtext
            _s:=Format("Текущая конфигурация добавляет в очередь {:.2f} нажатий в секунду, но извлекает только {:.2f}", _in, _out)
            GuiControlGet, _hwnd, Hwnd, skillset%currentProfile%skillqueuewarningtext
            AddToolTip(_hwnd, _s "`nРекомендуется для баффов использовать режим «Поддержание баффа», а не «Автоповтор».`nЛибо увеличьте интервал автоповтора, либо уменьшите интервал отправки очереди.", 30000, True)
        }
        Else
        {
            GuiControl, Hide, skillset%currentProfile%skillqueuewarningtext
        }
    }
    Else
    {
        GuiControl, Hide, skillset%currentProfile%skillqueuewarningtext
    }
    Return
}

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

SetCustomPotion(){
    GuiControlGet, extraCustomPotion
    if extraCustomPotion
    {
        GuiControl, Enable, extraCustomPotionHK
        GuiControlGet, extraCustomPotionHK
        if !extraCustomPotionHK
        {
            GuiControl,, extraCustomPotionHK, q
        }
    }
    Else
    {
        GuiControl, Disable, extraCustomPotionHK
    }
    Return
}

SetGambleHelper(){
    GuiControlGet, extraGambleHelperCKbox
    If extraGambleHelperCKbox
    {
        GuiControl, Enable, extraGambleHelperText
        GuiControl, Enable, extraGambleHelperEdit
    }
    Else
    {
        GuiControl, Disable, extraGambleHelperText
        GuiControl, Disable, extraGambleHelperEdit
    }
    Return
}

SetLootHelper(){
    GuiControlGet, extraLootHelperCkbox
    If extraLootHelperCkbox
    {
        GuiControl, Enable, extraLootHelperText
        GuiControl, Enable, extraLootHelperEdit
    }
    Else
    {
        GuiControl, Disable, extraLootHelperText
        GuiControl, Disable, extraLootHelperEdit
    }
    Return
}

SetReforgeHelper(){
    GuiControlGet, extraReforgeHelperCkbox
    If extraReforgeHelperCkbox
    {
        GuiControl, Enable, extraReforgeHelperDropdown
    }
    Else
    {
        GuiControl, Disable, extraReforgeHelperDropdown
    }
    Return
}

SetSalvageHelper(){
    local
    Global safezone
    Gui, Submit, NoHide
    GuiControlGet, extraSalvageHelperCkbox
    GuiControlGet, extraSalvageHelperDropdown
    GuiControlGet, extraUpgradeHelperCkbox
    GuiControlGet, extraConvertHelperCkbox
    GuiControlGet, extraAbandonHelperCkbox
    If extraSalvageHelperCkbox or extraUpgradeHelperCkbox or extraConvertHelperCkbox or extraAbandonHelperCkbox
    {
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
            GuiControl,, helperSafeZoneText, Защищённые ячейки заданы
        }
        Else
        {
            GuiControl, +cFF0000, helperSafeZoneText
            GuiControl,, helperSafeZoneText, Защищённые ячейки НЕ заданы
        }

        GuiControl, Enable, extraSalvageHelperDropdown
        switch extraSalvageHelperDropdown
        {
            case 1:
                if extraUpgradeHelperCkbox or extraConvertHelperCkbox or extraAbandonHelperCkbox
                {
                    GuiControl, show, helperSafeZoneText
                }
                Else
                {
                    GuiControl, hide, helperSafeZoneText
                }
            case 2,3,4,5:
                GuiControl, show, helperSafeZoneText
        }
    }
    Else
    {
        if not extraSalvageHelperCkbox
        {
            GuiControl, Disable, extraSalvageHelperDropdown
        }
        GuiControl, Hide, helperSafeZoneText
    }
    Return
}

SetSkillQueue(){
    local
    Global tabslen
    Loop, %tabslen%
    {
        GuiControlGet, skillset%A_Index%useskillqueueckbox
        if skillset%A_Index%useskillqueueckbox
        {
            GuiControl, Enable, skillset%A_Index%useskillqueueedit
        }
        Else
        {
            GuiControl, Disable, skillset%A_Index%useskillqueueedit
        }
    }
    SetSkillQueueWarning()
    Return
}

spamSkillQueue(inv){
    local
    Global skillQueue, forceStandingKey, keysOnHold
    while (skillQueue.Count() > 0)
    {
        _k:=skillQueue.RemoveAt(1)
        k:=_k[1]
        if (_k[2]=3){
            for key, value in keysOnHold{
                if GetKeyState(key){
                    Send {%key% up}
                }
            }
            Sleep, inv//4
        }

        if (!GetKeyState(forceStandingKey) and (_k[2]=3 or k="LButton")){
            Send {Blind}{%forceStandingKey% down}{%k% down}
            if (_k[2]=3){
                Sleep, inv//4
            }
            Send {Blind}{%k% up}{%forceStandingKey% up}
        }
        Else{
            Send {%k%}
        }

        if (_k[2]=3){
            Sleep, inv//4
            for key, value in keysOnHold{
                if !GetKeyState(key){
                    Send {%key% down}
                }
            }
            Break
        }
    }
    Return
}

isDialogBoXOnScreen(D3W, D3H){
    point1:=[D3W/2-(3440/2-1655)*D3H/1440, 500*D3H/1440]
    point2:=[D3W/2+(3440/2-1800)*D3H/1440, 500*D3H/1440]
    c1:=getPixelRGB(point1)
    c2:=getPixelRGB(point2)
    if (c1[1]>c1[2] and c1[2]>c1[3] and c1[3]<5 and c1[2]<15 and c1[1]>25 and c2[1]>c2[2] and c2[2]>c2[3] and c2[3]<5 and c2[2]<15 and c2[1]>25)
    {
        Return True
    }
    Else
    {
        Return False
    }
}

getInventorySpaceXY(D3W, D3H, ID, zone){
    static _spaceSizeInnerW:=64
    static _spaceSizeInnerH:=63
    static _spaceSizeW:=67
    static _spaceSizeH:=66
    static _spaceBagX:=[2753,2820,2887,2954,3021,3089,3156,3223,3290,3357]
    static _spaceBagY:=[747,813,880,946,1013,1079]
    static _spaceKanaiX:=[242, 318, 394]
    static _spaceKanaiY:=[503, 579, 655]

    switch zone
    {
        case "bag":
            targetColumn:=(Mod(ID,10)=0)?10:Mod(ID,10)
            targetRow:=Floor((ID-1)/10)+1
            Return [Round(D3W-((3440-_spaceBagX[targetColumn]-_spaceSizeInnerW/2)*D3H/1440)), Round((_spaceBagY[targetRow]+_spaceSizeInnerH/2)*D3H/1440)
            , Round(D3W-((3440-_spaceBagX[targetColumn])*D3H/1440)), Round((_spaceBagY[targetRow])*D3H/1440)]
        case "kanai":
            targetColumn:=(Mod(ID,3)=0)?3:Mod(ID,3)
            targetRow:=Floor((ID-1)/3)+1
            Return [Round((_spaceKanaiX[targetColumn]+_spaceSizeInnerW/2)*D3H/1440), Round((_spaceKanaiY[targetRow]+_spaceSizeInnerH/2)*D3H/1440)
            , Round(_spaceKanaiX[targetColumn]*D3H/1440), Round((_spaceKanaiY[targetRow])*D3H/1440)]
    }
}

isSalvagePageOpen(D3W, D3H){
    c1:=getPixelRGB([Round(339*D3H/1440),Round(80*D3H/1440)])
    c2:=getPixelRGB([Round(351*D3H/1440),Round(107*D3H/1440)])
    c3:=getPixelRGB([Round(388*D3H/1440),Round(86*D3H/1440)])
    c4:=getPixelRGB([Round(673*D3H/1440),Round(1040*D3H/1440)])
    if (c1[3]>c1[2] and c1[2]>c1[1] and c1[3]>170 and c1[3]-c1[1]>80 and c3[3]>c3[2] and c3[2]>c3[1] and c3[3]>110 and c2[1]+c2[2]>350 and c4[1]>50 and c4[2]<15 and c4[3]<15){
        p:=getSalvageIconXY(D3W, D3H, "edge")
        cLeg:=getPixelRGB(p[1])
        cWhite:=getPixelRGB(p[2])
        cBlue:=getPixelRGB(p[3])
        cRare:=getPixelRGB(p[4])
        if (cBlue[3]>cBlue[2] and cBlue[2]>cBlue[1] and cRare[3]<20 and cRare[1]>cRare[2] and cRare[2]>cRare[3]) {
            Return [2, cLeg, cWhite, cBlue, cRare]
        } Else {
            Return [1]
        }
    }
    Else {
        Return [0]
    }
}

/*
Возвращает состояние окна Куба Канаи
0: Куб не открыт
1: Куб открыт, но страница неизвестна
2: Открыта страница перековки (рефордж)
3: Открыта страница улучшения (апгрейд)
4: Открыта страница конвертации материалов
*/
isKanaiCubeOpen(D3W, D3H){
    c1:=getPixelRGB([Round(353*D3H/1440),Round(85*D3H/1440)])
    c2:=getPixelRGB([Round(278*D3H/1440),Round(147*D3H/1440)])
    c3:=getPixelRGB([Round(330*D3H/1440),Round(140*D3H/1440)])

    if (c1[1]<50 and c1[2]<40 and c1[3]<35 and c2[1]>100 and c2[2]<30 and c2[3]<30 and abs(c3[3]-c3[2])<=8 and c3[1]<=55 and c3[1]<c3[2] and c3[1]<c3[3]){
        cc1:=getPixelRGB([Round(788*D3H/1440),Round(428*D3H/1440)])
        cc2:=getPixelRGB([Round(810*D3H/1440),Round(429*D3H/1440)])
        if (cc1[3]>230 and cc2[3]>230 and cc1[3]>cc1[2] and cc2[3]>cc2[2] and cc1[2]>cc1[1] and cc2[2]>cc2[1])
        {
            Return 2
        }
        else
        {
            WinGetTitle, gameWindowTitle, ahk_class D3 Main Window Class
            upgradeYOffset:=(gameWindowTitle="Diablo III")? 0:-22
            cc1:=getPixelRGB([Round(799*D3H/1440),Round((406+upgradeYOffset)*D3H/1440)])
            cc2:=getPixelRGB([Round(795*D3H/1440),Round((592+upgradeYOffset)*D3H/1440)])
            if (cc1[1]+cc1[2]+cc1[3]>550 and cc1[1]>cc1[3] and cc2[1]+cc2[2]>400 and cc2[1]>cc2[3])
            {
                Return 3
            }

            convertYOffset:=(gameWindowTitle="Diablo III")? 0:-43
            cc3:=getPixelRGB([Round(799*D3H/1440),Round((365+convertYOffset)*D3H/1440)])
            if (cc3[1]+cc3[2]+cc3[3]>600 and cc3[1]>cc3[2] and cc3[2]>cc3[3] and cc3[3]>110 and cc3[3]<200)
            {
                Return 4
            }
        }
        Return 1
    }
    Else {
        Return 0
    }
}

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

isGambleOpen(D3W, D3H){
    c1:=getPixelRGB([Round(320*D3H/1440),Round(96*D3H/1440)])
    c2:=getPixelRGB([Round(351*D3H/1440),Round(100*D3H/1440)])
    c4:=getPixelRGB([Round(194*D3H/1440),Round(67*D3H/1440)])
    c5:=getPixelRGB([Round(147*D3H/1440),Round(94*D3H/1440)])
    if (c1[3]>c1[1] and c1[1]>c1[2] and c1[3]>130 and c2[1]+c2[2]>330 and c4[1]+c4[2]+c4[3]+c5[1]+c5[2]+c5[3]<10){
        Return True
    }
    Else{
        Return False
    }
}

isInventoryOpen(D3W, D3H){
    c1:=getPixelRGB([Round(D3W - (3440-3086)*D3H/1440),Round(108*D3H/1440)])
    c2:=getPixelRGB([Round(D3W - (3440-3010)*D3H/1440),Round(147*D3H/1440)])
    c3:=getPixelRGB([Round(D3W - (3440-3425)*D3H/1440),Round(142*D3H/1440)])
    c4:=getPixelRGB([Round(D3W - (3440-3117)*D3H/1440),Round(84*D3H/1440)])
    if (c1[1]+c1[2]>240 and c2[1]>115 and c2[2]<30 and c2[3]<30 and abs(c3[1]-c3[2])<=10 and c3[3]<40 and c4[3]>c4[2]+60 and c4[2]>c4[1]){
        Return True
    }
    Else
    {
        Return False
    }
}

isStashOpen(D3W, D3H){
    c1:=getPixelRGB([Round(282*D3H/1440),Round(147*D3H/1440)])
    c2:=getPixelRGB([Round(382*D3H/1440),Round(77*D3H/1440)])
    c3:=getPixelRGB([Round(299*D3H/1440),Round(82*D3H/1440)])
    if (c1[1]>100 and c1[1]>c1[2]+80 and abs(c1[2]-c1[3])<10 and c2[2]>c2[3] and c2[3]>c2[1] and c2[2]-c2[1]>80 and c3[1]>c3[2] and c3[2]>c3[3] and c3[3]<40){
        Return 1
    }
    Else
    {
        Return 0
    }
}

isInventorySpaceEmpty(D3W, D3H, ID, ckpoints, zone){
    static _spaceSizeInnerW:=64
    static _spaceSizeInnerH:=63
    Global gameX, gameY
    m:=getInventorySpaceXY(D3W, D3H, ID, zone)
    if (ckpoints="")
    {
        c:=getPixelsRGB(Round(m[3]+0.2*_spaceSizeInnerW), Round(m[4]+0.2*_spaceSizeInnerH), Round(0.6*_spaceSizeInnerW), Round(0.6*_spaceSizeInnerH), "Max", True, gameX, gameY)
        if (c[1]>50 or c[2]>50 or c[3]>50)
        {
            Return False
        }
    }
    Else
    {
        for i, p in ckpoints
        {
            xy:=[Round(m[3]+_spaceSizeInnerW*ckpoints[i][1]*D3H/1440), Round(m[4]+_spaceSizeInnerH*ckpoints[i][2]*D3H/1440)]
            c:=getPixelRGB(xy)
            if !(c[1]<22 and c[2]<20 and c[3]<15 and c[1]>c[3] and c[2]>c[3])
            {
                Return False
            }
        }
    }
    Return True
}

getGameXYonScreen(GameX, GameY){
    VarSetCapacity(POINT, 8)
    NumPut(GameX, POINT, 0, "Int")
    NumPut(GameY, POINT, 4, "Int")
    DllCall("ClientToScreen", "ptr", WinExist("ahk_class D3 Main Window Class"), "ptr", &POINT)
    Return [NumGet(POINT, 0, "Int"), NumGet(POINT, 4, "Int")]
}

getGameResulution(ByRef D3W, ByRef D3H){
    local
    Global gameResolution, d3only
    if (gameResolution="Auto")
    {
        VarSetCapacity(rect, 16)
        DllCall("GetClientRect", "ptr", WinExist("ahk_class D3 Main Window Class"), "ptr", &rect)
        D3W:=NumGet(rect, 8, "Int")
        D3H:=NumGet(rect, 12, "Int")
        if (D3W*D3H=0 and d3only){
            MsgBox, % Format("Не удалось определить разрешение игры. Код ошибки: 0x{:X}. Попробуйте переключить игру в оконный режим.", A_LastError)
            Return False
        }
    }
    Else
    {
        _r:=StrSplit(gameResolution, "x", A_Space)
        D3W:=_r[1]
        D3H:=_r[2]
    }
    Return True
}

getKanaiCubeButtonPos(D3W, D3H){
    point1:=[Round(320*D3H/1440),Round(1105*D3H/1440)]
    point2:=[Round(955*D3H/1440),Round(1115*D3H/1440)]
    point3:=[Round(777*D3H/1440),Round(1117*D3H/1440)]
    point4:=[Round(1135*D3H/1440),Round(1117*D3H/1440)]
    Return [point1, point2, point3, point4]
}

getPixelRGB(point){
    PixelGetColor, cpixel, point[1], point[2], rgb
    Return splitRGB(cpixel)
}

getPixelsRGB(pointX, pointY, w, h, agg_func="", gdip=False, gameX=0, gameY=0){
    cpixelR:=[]
    cpixelG:=[]
    cpixelB:=[]
    if gdip
    {
        pBitmap:=Gdip_BitmapFromScreen(Format("{}|{}|{}|{}", pointX+gameX, pointY+gameY, w, h))
        Gdip_LockBits(pBitmap, 0, 0, Gdip_GetImageWidth(pBitmap), Gdip_GetImageHeight(pBitmap), Stride, Scan0, BitmapData)
        Loop, %w%
        {
            _x:=A_Index-1
            Loop, %h%
            {
                _y:=A_Index-1
                t:=splitRGB(Gdip_GetLockBitPixel(Scan0, _x, _y, Stride))
                cpixelR.Push(t[1])
                cpixelG.Push(t[2])
                cpixelB.Push(t[3])
            }
        }
        Gdip_UnlockBits(pBitmap, BitmapData)
        Gdip_DisposeImage(pBitmap)
    }
    Else
    {
        Loop, %w%
        {
            _x:=A_Index-1
            Loop, %h%
            {
                _y:=A_Index-1
                t:=getPixelRGB([_x+pointX, _y+pointY])
                cpixelR.Push(t[1])
                cpixelG.Push(t[2])
                cpixelB.Push(t[3])
            }
        }
    }
    if not agg_func {
        Return [cpixelR, cpixelG, cpixelB]
    }
    Else {
        Return [Func(agg_func).Call(cpixelR*), Func(agg_func).Call(cpixelG*), Func(agg_func).Call(cpixelB*)]
    }
}

keyJoin(sep, dict){
    for key,value in dict
        str .= key . sep
    return SubStr(str, 1, -StrLen(sep))
}

HasVal(haystack, needle) {
    for index, value in haystack
        if (value = needle)
            return index
    if !(IsObject(haystack))
        throw Exception("Bad haystack!", -1, haystack)
    return 0
}

make1DArray(len, fill=0){
    outArray:=[]
    Loop, %len%
    {
        outArray.Push(fill)
    }
    Return outArray
}

isArraysEqual(arrayA, arrayB, _t=0){
    if (arrayA.Length()!=arrayB.Length())
    {
        Return False
    }
    _l:=arrayA.Length()
    Loop, %_l%
    {
        if (abs(arrayA[A_Index]-arrayb[A_Index])>_t)
        {
            Return False
        }
    }
    Return True
}

dummyFunction(){
    Return
}

FillPixel(HWNDs, HexColor) {
    hBitmap := DllCall("CreateBitmap", "Int", 1, "Int", 1, "UInt", 1, "UInt", 32, "PtrP", HexColor, "Ptr")
    hBM := DllCall("CopyImage", "Ptr", hBitmap, "UInt", 0, "Int", 0, "Int", 0, "UInt", 0x2000|0x8|0x4, "Ptr")
    if IsObject(HWNDs)
    {
        for i, HWND in HWNDs
        {
            SendMessage, 0x172,, hBM,, ahk_id %HWND%
        }
    }
    Else
    {
        SendMessage, 0x172,, hBM,, ahk_id %HWNDs%
    }
    DllCall("DeleteObject", "Ptr", hBitmap)
    Return
}

GdipCreateFromBase64(B64, IsIcon := 0){
    VarSetCapacity(B64Len, 0)
    DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", B64Len, "Ptr", 0, "Ptr", 0)
    VarSetCapacity(B64Dec, B64Len, 0)
    DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &B64Dec, "UIntP", B64Len, "Ptr", 0, "Ptr", 0)
    pStream := DllCall("Shlwapi.dll\SHCreateMemStream", "Ptr", &B64Dec, "UInt", B64Len, "UPtr")
    VarSetCapacity(pBitmap, 0)
    DllCall("Gdiplus.dll\GdipCreateBitmapFromStreamICM", "Ptr", pStream, "PtrP", pBitmap)
    VarSetCapacity(hBitmap, 0)
    DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "UInt", pBitmap, "UInt*", hBitmap, "Int", 0x00FFFFFF)

    If (IsIcon) {
        DllCall("Gdiplus.dll\GdipCreateHICONFromBitmap", "Ptr", pBitmap, "PtrP", hIcon, "UInt", 0)
    }

    ObjRelease(pStream)
    return (IsIcon ? hIcon : hBitmap)
}

AddToolTip(con, text, duration=30000, Modify=0){
    Static TThwnd, GuiHwnd
    PtrSize := (A_PtrSize ? A_PtrSize : 4)
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
                    ,"UInt",0
                    ,"Str","tooltips_class32"
                    ,"UInt",0
                    ,"UInt",2147483648
                    ,"UInt",-2147483648
                    ,"UInt",-2147483648
                    ,"UInt",-2147483648
                    ,"UInt",-2147483648
                    ,"UInt",GuiHwnd
                    ,"UInt",0
                    ,"UInt",0
                    ,"UInt",0)
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
        SendMessage, %TTM_ADDTOOL%,, &TInfo,, ahk_id %TThwnd%
        SendMessage, %TTM_SETMAXTIPWIDTH%,, A_ScreenWidth,, ahk_id %TThwnd%
        SendMessage, %TTM_SETDELAYTIME%, TTF_AUTOPOP, duration,, ahk_id %TThwnd%
    }
    SendMessage, %TTM_UPDATETIPTEXT%,, &TInfo,, ahk_id %TThwnd%
    Return
}

objectSort(obj, keyName="", callbackFunc="", reverse=false)
{
    temp := Object()
    sorted := Object()
    
    for oneKey, oneValue in obj
    {
        if keyname
            value := oneValue[keyName]
        else
            value := oneValue
        
        if (callbackFunc)
            tempKey := %callbackFunc%(value)
        else
            tempKey := value
        
        if not isObject(temp[tempKey])
            temp[tempKey] := []
        temp[tempKey].push(oneValue)
    }
    
    for oneTempKey, oneValueList in temp
    {
        for oneValueIndex, oneValue in oneValueList
        {
            if (reverse)
                sorted.insertAt(1,oneValue)
            else
                sorted.push(oneValue)
        }
    }
    
    return sorted
}

Watchdog(wParam, lParam){
    Global
    If (wParam = 32772 or wParam = 4)     ; HSHELL_WINDOWCREATED 1, HSHELL_WINDOWACTIVATED 4, HSHELL_RUDEAPPACTIVATED 32772
    {
        helperBreak:=True
        if (lParam=0)
        {
            vFront:=True
            FillPixel(TitlebarID, 0x34495e)
            FillPixel([TitlebarLineID, BorderTopID, BorderBottomID, BorderLeftID, BorderRightID], 0x000000)
            Gui, Font, s11 +cFFFFFF Normal
            GuiControl, Font, TitleBarText
            GuiControl,, TitleBarText, % TITLE
            GuiControl,, % UIRightButtonID, % "HBITMAP:*" hBMPButtonClose_Normal
            GuiControl,, % UILeftButtonID, % "HBITMAP:*" hBMPButtonLeft_Normal
            if (hHookMouse){
                DllCall("UnhookWindowsHookEx", "Uint", hHookMouse)
            }
            hHookMouse:=DllCall("SetWindowsHookEx", "int", 14, "Ptr", RegisterCallback("MouseMove", "Fast"), "Ptr", DllCall("GetModuleHandle", "Ptr", 0 ,"Ptr"), "Uint", 0, "Ptr")
        }
        Else
        {
            if (hHookMouse){
                DllCall("UnhookWindowsHookEx", "Uint", hHookMouse)
                hHookMouse:=0
            }
            if (vFront)
            {
                FillPixel([TitlebarID, TitlebarLineID], 0x607e9d)
                FillPixel([BorderTopID, BorderBottomID, BorderLeftID, BorderRightID], 0x607e9d)
                GuiControl, +cEEEEEE, TitleBarText
                GuiControl,, TitleBarText, % TITLE
                vFront:=False
            }
            WinGetClass, AClass, ahk_id %lParam%
            if (vRunning and d3only and AClass != "D3 Main Window Class")
            {
                Gosub, StopMarco
            }
        }
    }
    Return
}

MouseMove(nCode, wParam, lParam)
{
    Global
    If (nCode=0)
    {
        MouseGetPos, , , , currentControlUnderMouse, 2
        switch wParam
        {
            case 0x200:
                switch currentControlUnderMouse
                {
                    case UIRightButtonID:
                        if (RightButtonState!=1)
                        {
                            GuiControl,, % UIRightButtonID, % "HBITMAP:*" hBMPButtonClose_Hover
                            RightButtonState:=1
                        }
                        if (LeftButtonState!=0)
                        {
                            GuiControl,, % UILeftButtonID, % "HBITMAP:*" hBMPButtonLeft_Normal
                            LeftButtonState:=0
                        }
                    case UILeftButtonID:
                        if (LeftButtonState!=1)
                        {
                            GuiControl,, % UILeftButtonID, % "HBITMAP:*" hBMPButtonLeft_Hover
                            LeftButtonState:=1
                        }
                        if (RightButtonState!=0)
                        {
                            GuiControl,, % UIRightButtonID, % "HBITMAP:*" hBMPButtonClose_Normal
                            RightButtonState:=0
                        }
                    Default:
                        if (RightButtonState!=0)
                        {
                            GuiControl,, % UIRightButtonID, % "HBITMAP:*" hBMPButtonClose_Normal
                            RightButtonState:=0
                        }
                        if (LeftButtonState!=0)
                        {
                            GuiControl,, % UILeftButtonID, % "HBITMAP:*" hBMPButtonLeft_Normal
                            LeftButtonState:=0
                        }
                        if (currentControlUnderMouse=TitleBarID or currentControlUnderMouse=TitleBarTextID)
                        {
                            PostMessage, 0xA1, 2,,, A
                        }
                }
            case 0x201,0x204:
                if (currentControlUnderMouse=UIRightButtonID)
                {
                    GuiControl,, % UIRightButtonID, % "HBITMAP:*" hBMPButtonClose_Pressed
                    RightButtonState:=2
                }
                if (currentControlUnderMouse=UILeftButtonID)
                {
                    GuiControl,, % UILeftButtonID, % "HBITMAP:*" hBMPButtonLeft_Pressed
                    LeftButtonState:=2
                }
            case 0x202,0x205:
                switch currentControlUnderMouse
                {
                    case UIRightButtonID:
                        if (wParam=0x202)
                        {
                            GuiClose()
                        }
                        Else
                        {
                            SetTimer, GuiExit, -1
                        }
                    case UILeftButtonID:
                        showMainWindow(isCompact? MainWindowW:CompactWindowW, MainWindowH)
                        isCompact:=!isCompact
                        hBMPButtonLeft_Normal := isCompact? hBMPButtonExpand_Normal:hBMPButtonBack_Normal
                        hBMPButtonLeft_Hover := isCompact? hBMPButtonExpand_Hover:hBMPButtonBack_Hover
                        hBMPButtonLeft_Pressed := isCompact? hBMPButtonExpand_Pressed:hBMPButtonBack_Pressed
                        if (LeftButtonState!=1)
                        {
                            GuiControl,, % UILeftButtonID, % "HBITMAP:*" hBMPButtonLeft_Hover
                            LeftButtonState:=1
                        }
                    Default:
                        if (RightButtonState!=0)
                        {
                            GuiControl,, % UIRightButtonID, % "HBITMAP:*" hBMPButtonClose_Normal
                            RightButtonState:=0
                        }
                        if (LeftButtonState!=0)
                        {
                            GuiControl,, % UILeftButtonID, % "HBITMAP:*" hBMPButtonLeft_Normal
                            LeftButtonState:=0
                        }
                }
        }
    }
    Return DllCall("CallNextHookEx", "Ptr", 0, "int", nCode, "Uint", wParam, "Ptr", lParam)
}

showMainWindow(windowSizeW, windowSizeH){
    global
    Gui Show, w%windowSizeW% h%windowSizeH%
    GuiControl, Move, TitleBar, % "w" windowSizeW-2
    GuiControl, Move, UIRightButton, % "x" windowSizeW-30-1
    GuiControl, Move, TitleBarText, % "x" (windowSizeW-TitleBarSizeW)/2
    GuiControl, Move, BorderTop, % "w" windowSizeW
    GuiControl, Move, BorderBottom, % "y" windowSizeH-1 " w" windowSizeW
    GuiControl, Move, BorderLeft, % "h" windowSizeH-2
    GuiControl, Move, BorderRight, % "x" windowSizeW-1 " h" windowSizeH-2
    WinSet, Redraw,, A
    Return
}
; ===================================== Метки (Subroutines) ===================================
spamSkillKeyA1:
spamSkillKeyA2:
spamSkillKeyA3:
spamSkillKeyA4:
spamSkillKeyA5:
spamSkillKeyA6:
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

SetTabFocus:
    Gui, Submit, NoHide
    GuiControl, , StatuesSkillsetText, % tabsarray[ActiveTab]
    currentProfile:=ActiveTab
    SetStartMode()
Return

SetQuickPause:
    Gui, Submit, NoHide
    GuiControlGet, skillset%currentProfile%clickpauseckbox
    GuiControlGet, skillset%currentProfile%clickpausedropdown2
    GuiControlGet, skillset%currentProfile%clickpausedropdown1
    mousePauseKeyArray:=["LButton", "RButton", "MButton", "XButton1", "XButton2"]
    currentQuickPauseHK:=mousePauseKeyArray[skillset%currentProfile%clickpausedropdown2]
    if skillset%currentProfile%clickpauseckbox
    {
        GuiControl, Enable, skillset%currentProfile%clickpausedropdown1
        GuiControl, Enable, skillset%currentProfile%clickpausedropdown2
        GuiControl, Enable, skillset%currentProfile%clickpausedropdown3
        GuiControl, Enable, skillset%currentProfile%clickpausetext1
        if (skillset%currentProfile%clickpausedropdown1!=3)
        {
            GuiControl, Enable, skillset%currentProfile%clickpauseedit
            GuiControl, Enable, skillset%currentProfile%clickpausetext2
        }
        Else
        {
            GuiControl, Disable, skillset%currentProfile%clickpauseedit
            GuiControl, Disable, skillset%currentProfile%clickpausetext2
        }
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
        GuiControl, Disable, skillset%currentProfile%clickpausedropdown3
        GuiControl, Disable, skillset%currentProfile%clickpausetext1
        GuiControl, Disable, skillset%currentProfile%clickpauseedit
        GuiControl, Disable, skillset%currentProfile%clickpausetext2
        Hotkey, ~*%currentQuickPauseHK%, quickPause, off
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
                GuiControl, Disable, skillset%currentPage%autostartmarcockbox
            case 2,3,4,5,6:
                GuiControl, Disable, skillset%currentPage%profilekeybindinghkbox
                GuiControl, Enable, skillset%currentPage%autostartmarcockbox
                ckey:=mouseKeyArray[skillset%currentPage%profilekeybindingdropdown]
                Hotkey, ~*%ckey%, SwitchProfile, on
                profileKeybinding[ckey]:=currentPage
            case 7:
                GuiControl, Enable, skillset%currentPage%profilekeybindinghkbox
                GuiControl, Enable, skillset%currentPage%autostartmarcockbox
                ckey:=skillset%currentPage%profilekeybindinghkbox
                if (ckey!="")
                {
                    Hotkey, ~*%ckey%, SwitchProfile, on
                    profileKeybinding[ckey]:=currentPage
                } 
        }
    }
Return

SwitchProfile:
    currentHK:=RegExReplace(A_ThisHotkey, "[~*]")
    if (currentProfile!=profileKeybinding[currentHK])
    {
        wasRunning:=vRunning
        wasPausing:=vPausing
        currentProfile:=profileKeybinding[currentHK]
        GuiControl , Choose, ActiveTab, % tabsarray[currentProfile]
        Gosub, SetTabFocus
        Gosub, StopMarco
        GuiControlGet, extraSoundonProfileSwitch
        if extraSoundonProfileSwitch
        {
            SoundBeep, 750, 250
        }
        if (wasRunning and !wasPausing and skillset%currentProfile%autostartmarcockbox and skillset%currentProfile%profilestartmodedropdown=1)
        {
            Gosub, RunMarco
        }
    }
Return

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
                GuiControl, Disable, skillset%A_Index%s6edit
                GuiControl, Disable, skillset%A_Index%s6delayedit
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

SetMovingHelper:
    Gui, Submit, NoHide
    Loop, %tabslen%
    {
        if (skillset%currentProfile%movingdropdown = 4)
        {
            GuiControl, Enable, skillset%A_Index%movingtext
            GuiControl, Enable, skillset%A_Index%movingedit
        }
        Else
        { 
            GuiControl, Disable, skillset%A_Index%movingtext
            GuiControl, Disable, skillset%A_Index%movingedit
        }
        if (skillset%currentProfile%potiondropdown > 1)
        {
            GuiControl, Enable, skillset%A_Index%potiontext
            GuiControl, Enable, skillset%A_Index%potionedit
        }
        Else
        { 
            GuiControl, Disable, skillset%A_Index%potiontext
            GuiControl, Disable, skillset%A_Index%potionedit
        }
    }
Return

SetSkillsetDropdown:
    Gui, Submit, NoHide
    Loop, %tabslen%
    {
        npage:=A_Index
        Loop, 6
        {
            switch skillset%npage%s%A_Index%dropdown
            {
                case 1,2:
                    GuiControl, Disable, skillset%npage%s%A_Index%edit
                    GuiControl, Disable, skillset%npage%s%A_Index%delayedit
                    GuiControl, Disable, skillset%npage%s%A_Index%randomckbox
                case 3:
                    GuiControl, Enable, skillset%npage%s%A_Index%edit
                    GuiControl, Enable, skillset%npage%s%A_Index%delayedit
                    GuiControl, Enable, skillset%npage%s%A_Index%randomckbox
                case 4:
                    GuiControl, Enable, skillset%npage%s%A_Index%edit
                    GuiControl, Disable, skillset%npage%s%A_Index%delayedit
                    GuiControl, Disable, skillset%npage%s%A_Index%randomckbox
                case 5:
                    GuiControl, Disable, skillset%npage%s%A_Index%edit
                    GuiControl, Enable, skillset%npage%s%A_Index%delayedit
                    GuiControl, Enable, skillset%npage%s%A_Index%randomckbox
            }
        }
    }
    SetSkillQueueWarning()
Return

MainMacro:
    GuiControlGet, skillset%currentProfile%profilestartmodedropdown
    switch skillset%currentProfile%profilestartmodedropdown
    {
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
            Gosub, RunMarco
            KeyWait, %startRunHK%
            Gosub, StopMarco
        case 3:
            Loop, 6
            {
                GuiControlGet, skillset%currentProfile%s%A_Index%dropdown
                GuiControlGet, skillset%currentProfile%s%A_Index%hotkey
                Switch skillset%currentProfile%s%A_Index%dropdown
                {
                    Case 2:
                        k:=skillset%currentProfile%s%A_Index%hotkey
                        Send {%k%}
                }
            }
    }
Return

RunMarco:
    Gui, Submit, NoHide
    GuiControlGet, extraCustomStanding
    GuiControlGet, extraCustomStandingHK
    forceStandingKey:=extraCustomStanding? extraCustomStandingHK:"LShift"
    GuiControlGet, extraCustomMoving
    GuiControlGet, extraCustomMovingHK
    forceMovingKey:=extraCustomMoving? extraCustomMovingHK:"e"
    potionKey:=extraCustomPotion? extraCustomPotionHK:"q"
    skillQueue:=[]
    syncTimer:={}
    syncDelay:={}
    if (!getGameResulution(D3W, D3H) and d3only)
    {
        Return
    }
    gameXY:=getGameXYonScreen(0,0)
    gameX:=gameXY[1]
    gameY:=gameXY[2]
    keyDelay:=[]
    Loop, 6
    {
        GuiControlGet, skillset%currentProfile%s%A_Index%dropdown
        GuiControlGet, skillset%currentProfile%s%A_Index%delayupdown
        GuiControlGet, skillset%currentProfile%s%A_Index%updown
        keyDelay.Push({"key":A_Index, "delay":(skillset%currentProfile%s%A_Index%dropdown=3)?mod(skillset%currentProfile%s%A_Index%updown + skillset%currentProfile%s%A_Index%delayupdown, skillset%currentProfile%s%A_Index%updown):0})
    }
    keyDelay:=ObjectSort(keyDelay, "delay", ,True)
    vRunning:=True
    for _, v in keyDelay
    {
        currentIndex:=v["key"]
        GuiControlGet, skillset%currentProfile%s%currentIndex%hotkey
        Switch skillset%currentProfile%s%currentIndex%dropdown
        {
            Case 2:
                k:=skillset%currentProfile%s%currentIndex%hotkey
                Send {%k% Down}
                keysOnHold[k]:=1
            Case 3, 4:
                if runOnStart{
                    SetTimer, spamSkillKeyA%currentIndex%, -1
                }
                GuiControlGet, skillset%currentProfile%s%currentIndex%updown
                SetTimer, spamSkillKey%currentIndex%, % skillset%currentProfile%s%currentIndex%updown
            case 5:
                k:=combats[currentProfile][currentIndex]["triggerbutton"]
                HotKey, ~*%k%, spamSkillKey%currentIndex%, on
            Default:
                SetTimer, spamSkillKey%currentIndex%, off
        }
        if (currentIndex <=4)
        {
            GuiControl, Disable, skillset%currentProfile%s%currentIndex%hotkey
        }
    }
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
            if runOnStart{
                Send {%extraCustomMovingHK%}
            }
            GuiControlGet, skillset%currentProfile%movingedit
            SetTimer, forceMoving, % skillset%currentProfile%movingedit

    }
    GuiControlGet, skillset%currentProfile%potiondropdown
    if (skillset%currentProfile%potiondropdown > 1)
    {
        if IsObject(pofunc){
            SetTimer, %pofunc%, off
        }
        GuiControlGet, skillset%currentProfile%potionedit
        pofunc:=Func("potionHelper").Bind(skillset%currentProfile%potiondropdown)
        SetTimer, %pofunc%, % skillset%currentProfile%potionupdown
    }
    if skillset%currentProfile%useskillqueueckbox{
        GuiControlGet, skillset%currentProfile%useskillqueueupdown
        sqfunc:=Func("spamSkillQueue").Bind(skillset%currentProfile%useskillqueueupdown)
        if runOnStart{
            SetTimer, %sqfunc%, -1
        }
        SetTimer, %sqfunc%, % skillset%currentProfile%useskillqueueupdown
    } 
    vPausing:=False
Return

StopMarco:
    if IsObject(sqfunc){
        SetTimer, %sqfunc%, off
    }
    if IsObject(pofunc){
        SetTimer, %pofunc%, off
    }
    skillQueue:=[]
    Loop, 6
    {
        SetTimer, spamSkillKey%A_Index%, off
        k:=combats[currentProfile][A_Index]["triggerbutton"]
        HotKey, ~*%k%, spamSkillKey%A_Index%, off
        if (A_Index <=4)
        {
            si:=A_Index
            Loop, %tabslen%
            {
                GuiControl, Enable, skillset%A_Index%s%si%hotkey
            }
        }
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

quickPause:
    GuiControlGet, skillset%currentProfile%clickpausedropdown1
    GuiControlGet, skillset%currentProfile%clickpausedropdown3
    GuiControlGet, skillset%currentProfile%clickpauseupdown
    switch skillset%currentProfile%clickpausedropdown1
    {
        case 1:
            If (A_PriorHotkey=A_ThisHotkey and A_TimeSincePriorHotkey < DblClickTime)
            {
                clickPauseMarco(skillset%currentProfile%clickpauseupdown, skillset%currentProfile%clickpausedropdown3)
            }
        case 2:
            clickPauseMarco(skillset%currentProfile%clickpauseupdown, skillset%currentProfile%clickpausedropdown3)
        case 3:
            clickPauseMarco(-1, skillset%currentProfile%clickpausedropdown3)
    }
Return

forceMoving:
    if !vPausing
    {
        Send {%forceMovingKey%}
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

NumpadIns::Numpad0
NumpadEnd::Numpad1
NumpadDown::Numpad2
NumpadPgDn::Numpad3
NumpadLeft::Numpad4
NumpadClear::Numpad5
NumpadRight::Numpad6
NumpadHome::Numpad7
NumpadUp::Numpad8
NumpadPgUp::Numpad9
NumpadDel::NumpadDot
; ===================================== System Functions ==================================
GuiClose(){
    Global
    Gui, Submit
    SaveCfgFile("d3oldsand.ini", tabs, currentProfile, safezone, VERSION)
    vFront:=False
    Return
}

GuiShowMainWindow(){
    Global
    Gui, Show,, %TIELE%
    Return
}

GuiExit(){
    Global
    Gui, Submit
    SaveCfgFile("d3oldsand.ini", tabs, currentProfile, safezone, VERSION)
    ExitApp
}

; ================================= GDIP библиотека ===============================
; https://github.com/mmikeww/AHKv2-Gdip
; Включены только необходимые функции
; =========================================================================
Gdip_BitmapFromScreen(Screen:=0, Raster:="")
{
    hhdc := 0
    Ptr := A_PtrSize ? "UPtr" : "UInt"
    if (Screen = 0)
    {
        _x := DllCall( "GetSystemMetrics", "Int", 76 )
        _y := DllCall( "GetSystemMetrics", "Int", 77 )
        _w := DllCall( "GetSystemMetrics", "Int", 78 )
        _h := DllCall( "GetSystemMetrics", "Int", 79 )
    }
    else if (SubStr(Screen, 1, 5) = "hwnd:")
    {
        Screen := SubStr(Screen, 6)
        if !WinExist("ahk_id " Screen)
            return -2
        WinGetRect(Screen,,, _w, _h)
        _x := _y := 0
        hhdc := GetDCEx(Screen, 3)
    }
    else if IsInteger(Screen)
    {
        M := GetMonitorInfo(Screen)
        _x := M.Left, _y := M.Top, _w := M.Right-M.Left, _h := M.Bottom-M.Top
    }
    else
    {
        S := StrSplit(Screen, "|")
        _x := S[1], _y := S[2], _w := S[3], _h := S[4]
    }

    if (_x = "") || (_y = "") || (_w = "") || (_h = "")
        return -1

    chdc := CreateCompatibleDC(), hbm := CreateDIBSection(_w, _h, chdc), obm := SelectObject(chdc, hbm), hhdc := hhdc ? hhdc : GetDC()
    BitBlt(chdc, 0, 0, _w, _h, hhdc, _x, _y, Raster)
    ReleaseDC(hhdc)

    pBitmap := Gdip_CreateBitmapFromHBITMAP(hbm)
    SelectObject(chdc, obm), DeleteObject(hbm), DeleteDC(hhdc), DeleteDC(chdc)
    return pBitmap
}

Gdip_LockBits(pBitmap, x, y, w, h, ByRef Stride, ByRef Scan0, ByRef BitmapData, LockMode := 3, PixelFormat := 0x26200a)
{
    Ptr := A_PtrSize ? "UPtr" : "UInt"

    CreateRect(_Rect, x, y, w, h)
    VarSetCapacity(BitmapData, 16+2*(A_PtrSize ? A_PtrSize : 4), 0)
    _E := DllCall("Gdiplus\GdipBitmapLockBits", Ptr, pBitmap, Ptr, &_Rect, "uint", LockMode, "int", PixelFormat, Ptr, &BitmapData)
    Stride := NumGet(BitmapData, 8, "Int")
    Scan0 := NumGet(BitmapData, 16, Ptr)
    return _E
}

Gdip_GetLockBitPixel(Scan0, x, y, Stride)
{
    return NumGet(Scan0+0, (x*4)+(y*Stride), "UInt")
}

Gdip_UnlockBits(pBitmap, ByRef BitmapData)
{
    Ptr := A_PtrSize ? "UPtr" : "UInt"

    return DllCall("Gdiplus\GdipBitmapUnlockBits", Ptr, pBitmap, Ptr, &BitmapData)
}

Gdip_DisposeImage(pBitmap)
{
    return DllCall("gdiplus\GdipDisposeImage", A_PtrSize ? "UPtr" : "UInt", pBitmap)
}

Gdip_GetImageWidth(pBitmap)
{
    Width := 0
    DllCall("gdiplus\GdipGetImageWidth", A_PtrSize ? "UPtr" : "UInt", pBitmap, "uint*", Width)
    return Width
}

Gdip_GetImageHeight(pBitmap)
{
    Height := 0
    DllCall("gdiplus\GdipGetImageHeight", A_PtrSize ? "UPtr" : "UInt", pBitmap, "uint*", Height)
    return Height
}

CreateCompatibleDC(hdc:=0)
{
    return DllCall("CreateCompatibleDC", A_PtrSize ? "UPtr" : "UInt", hdc)
}

CreateDIBSection(w, h, hdc:="", bpp:=32, ByRef ppvBits:=0)
{
    Ptr := A_PtrSize ? "UPtr" : "UInt"

    hdc2 := hdc ? hdc : GetDC()
    VarSetCapacity(bi, 40, 0)

    NumPut(w, bi, 4, "uint")
    , NumPut(h, bi, 8, "uint")
    , NumPut(40, bi, 0, "uint")
    , NumPut(1, bi, 12, "ushort")
    , NumPut(0, bi, 16, "uInt")
    , NumPut(bpp, bi, 14, "ushort")

    hbm := DllCall("CreateDIBSection"
                    , Ptr, hdc2
                    , Ptr, &bi
                    , "uint", 0
                    , A_PtrSize ? "UPtr*" : "uint*", ppvBits
                    , Ptr, 0
                    , "uint", 0, Ptr)

    if !hdc
        ReleaseDC(hdc2)
    return hbm
}

SelectObject(hdc, hgdiobj)
{
    Ptr := A_PtrSize ? "UPtr" : "UInt"

    return DllCall("SelectObject", Ptr, hdc, Ptr, hgdiobj)
}

GetDC(hwnd:=0)
{
    return DllCall("GetDC", A_PtrSize ? "UPtr" : "UInt", hwnd)
}

BitBlt(ddc, dx, dy, dw, dh, sdc, sx, sy, Raster:="")
{
    Ptr := A_PtrSize ? "UPtr" : "UInt"

    return DllCall("gdi32\BitBlt"
                    , Ptr, dDC
                    , "int", dx
                    , "int", dy
                    , "int", dw
                    , "int", dh
                    , Ptr, sDC
                    , "int", sx
                    , "int", sy
                    , "uint", Raster ? Raster : 0x00CC0020)
}

ReleaseDC(hdc, hwnd:=0)
{
    Ptr := A_PtrSize ? "UPtr" : "UInt"

    return DllCall("ReleaseDC", Ptr, hwnd, Ptr, hdc)
}

Gdip_CreateBitmapFromHBITMAP(hBitmap, Palette:=0)
{
    Ptr := A_PtrSize ? "UPtr" : "UInt"
    pBitmap := 0

    DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", Ptr, hBitmap, Ptr, Palette, A_PtrSize ? "UPtr*" : "uint*", pBitmap)
    return pBitmap
}

DeleteObject(hObject)
{
    return DllCall("DeleteObject", A_PtrSize ? "UPtr" : "UInt", hObject)
}

DeleteDC(hdc)
{
    return DllCall("DeleteDC", A_PtrSize ? "UPtr" : "UInt", hdc)
}

WinGetRect( hwnd, ByRef x:="", ByRef y:="", ByRef w:="", ByRef h:="" ) {
    Ptr := A_PtrSize ? "UPtr" : "UInt"
    CreateRect(winRect, 0, 0, 0, 0)
    DllCall( "GetWindowRect", Ptr, hwnd, Ptr, &winRect )
    x := NumGet(winRect,  0, "UInt")
    y := NumGet(winRect,  4, "UInt")
    w := NumGet(winRect,  8, "UInt") - x
    h := NumGet(winRect, 12, "UInt") - y
}

GetDCEx(hwnd, flags:=0, hrgnClip:=0)
{
    Ptr := A_PtrSize ? "UPtr" : "UInt"

    return DllCall("GetDCEx", Ptr, hwnd, Ptr, hrgnClip, "int", flags)
}

IsInteger(Var) {
    Static Integer := "Integer"
    If Var Is Integer
        Return True
    Return False
}

GetMonitorInfo(MonitorNum)
{
    Monitors := MDMF_Enum()
    for k,v in Monitors
        if (v.Num = MonitorNum)
            return v
}

CreateRect(ByRef Rect, x, y, w, h)
{
    VarSetCapacity(Rect, 16)
    NumPut(x, Rect, 0, "uint"), NumPut(y, Rect, 4, "uint"), NumPut(w, Rect, 8, "uint"), NumPut(h, Rect, 12, "uint")
}

MDMF_Enum(HMON := "") {
    Static CallbackFunc := Func(A_AhkVersion < "2" ? "RegisterCallback" : "CallbackCreate")
    Static EnumProc := CallbackFunc.Call("MDMF_EnumProc")
    Static Obj := (A_AhkVersion < "2") ? "Object" : "Map"
    Static Monitors := {}
    If (HMON = "") ; new enumeration
    {
        Monitors := %Obj%("TotalCount", 0)
        If !DllCall("User32.dll\EnumDisplayMonitors", "Ptr", 0, "Ptr", 0, "Ptr", EnumProc, "Ptr", &Monitors, "Int")
            Return False
    }
    Return (HMON = "") ? Monitors : Monitors.HasKey(HMON) ? Monitors[HMON] : False
}
