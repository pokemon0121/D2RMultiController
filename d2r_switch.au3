; ===== D2R 多开一键切换（AutoIt v3 兼容版） =====
; F1~F5：直达第 1~5 个 d2r.exe 窗口
; F6：弹出当前检测到的窗口清单
; F7：在所有 d2r.exe 窗口间循环切换
; 说明：每次按键都会即时重扫 d2r.exe 窗口，避免切换后句柄变化导致失效

#RequireAdmin                      ; 建议提权，若不想每次弹UAC，可注释掉这行
#include <Array.au3>
Opt("WinTitleMatchMode", 4)        ; 允许高级匹配模式

Global $g_sExe = "d2r.exe"
Global $g_iMaxSlots = 5
Global $g_bMinimizeOthers = False  ; 如需兜底，把它改为 True：先最小化其它实例再激活目标

; ---- 热键注册 ----
HotKeySet("{F1}", "_HK_F1")
HotKeySet("{F2}", "_HK_F2")
HotKeySet("{F3}", "_HK_F3")
HotKeySet("{F4}", "_HK_F4")
HotKeySet("{F5}", "_HK_F5")
HotKeySet("{F6}", "_HK_List")
HotKeySet("{F7}", "_HK_Cycle")

; 保持脚本存活
While 1
    Sleep(50)
WEnd

; ---- 热键处理 ----
Func _HK_F1()
    _SwitchTo(1)
EndFunc
Func _HK_F2()
    _SwitchTo(2)
EndFunc
Func _HK_F3()
    _SwitchTo(3)
EndFunc
Func _HK_F4()
    _SwitchTo(4)
EndFunc
Func _HK_F5()
    _SwitchTo(5)
EndFunc
Func _HK_List()
    Local $a = _GetD2RList()
    Local $msg = "当前找到 " & UBound($a) & " 个 " & $g_sExe & " 窗口:" & @CRLF & @CRLF
    For $i = 0 To UBound($a) - 1
        $msg &= StringFormat("%d) hwnd=0x%08X  title=%s", $i + 1, $a[$i], WinGetTitle($a[$i])) & @CRLF
    Next
    MsgBox(64, "D2R 窗口清单", $msg)
EndFunc
Func _HK_Cycle()
    Local $a = _GetD2RList()
    If UBound($a) = 0 Then
        TrayTip("提示", "未找到 " & $g_sExe & " 窗口", 1000)
        Return
    EndIf
    Local $hActive = WinGetHandle("[ACTIVE]")
    Local $idx = -1
    For $i = 0 To UBound($a) - 1
        If $a[$i] = $hActive Then
            $idx = $i
            ExitLoop
        EndIf
    Next
    Local $next = 0
    If $idx >= 0 And $idx < UBound($a) - 1 Then
        $next = $idx + 1
    Else
        $next = 0
    EndIf
    _ActivateHandle($a[$next])
EndFunc

; ---- 切到第 n 个实例 ----
Func _SwitchTo($n)
    If $n < 1 Or $n > $g_iMaxSlots Then Return
    Local $a = _GetD2RList() ; 每次即时重扫
    If UBound($a) < $n Then
        TrayTip("提示", "没有第 " & $n & " 个窗口（当前 " & UBound($a) & " 个）", 1200)
        Return
    EndIf
    _ActivateHandle($a[$n - 1])
EndFunc

; ---- 激活窗口（含强制前台兜底）----
Func _ActivateHandle($h)
    If Not WinExists($h) Then
        TrayTip("提示", "句柄已失效，请重试 (F6 可查看)", 1000)
        Return
    EndIf

    ; 可选兜底：先最小化其它实例
    If $g_bMinimizeOthers Then
        Local $aAll = _GetD2RList()
        For $i = 0 To UBound($aAll) - 1
            If $aAll[$i] <> $h Then
                DllCall("user32.dll", "int", "ShowWindow", "hwnd", $aAll[$i], "int", 6) ; SW_MINIMIZE
            EndIf
        Next
    EndIf

    ; 恢复并尝试激活
    DllCall("user32.dll", "int", "ShowWindow", "hwnd", $h, "int", 9) ; SW_RESTORE
    WinActivate($h)
    If WinWaitActive($h, "", 0.25) Then Return

    ; 兜底：调整Z序 -> 附加线程输入 -> 前台/焦点
    Local $HWND_TOPMOST = Ptr(-1), $HWND_NOTOPMOST = Ptr(-2)
    Local $SWP_NOMOVE = 0x2, $SWP_NOSIZE = 0x1, $SWP_SHOWWINDOW = 0x40

    DllCall("user32.dll", "int", "SetWindowPos", "hwnd", $h, "hwnd", $HWND_TOPMOST, _
            "int", 0, "int", 0, "int", 0, "int", 0, "uint", BitOR($SWP_NOMOVE, $SWP_NOSIZE, $SWP_SHOWWINDOW))
    DllCall("user32.dll", "int", "SetWindowPos", "hwnd", $h, "hwnd", $HWND_NOTOPMOST, _
            "int", 0, "int", 0, "int", 0, "int", 0, "uint", BitOR($SWP_NOMOVE, $SWP_NOSIZE, $SWP_SHOWWINDOW))

    ; ALT 闪一下（经典前台解锁）
    DllCall("user32.dll", "none", "keybd_event", "byte", 0x12, "byte", 0x38, "dword", 0, "ptr", 0) ; Alt down
    DllCall("user32.dll", "none", "keybd_event", "byte", 0x12, "byte", 0x38, "dword", 2, "ptr", 0) ; Alt up

    ; 附加前台线程输入
    Local $hFG = DllCall("user32.dll", "hwnd", "GetForegroundWindow")
    Local $tPID = DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hFG[0], "dword*", 0)
    Local $fgTID = $tPID[0]
    Local $curTID = DllCall("kernel32.dll", "dword", "GetCurrentThreadId")
    If $fgTID <> 0 Then
        DllCall("user32.dll", "int", "AttachThreadInput", "dword", $curTID[0], "dword", $fgTID, "int", 1)
    EndIf

    DllCall("user32.dll", "int", "BringWindowToTop", "hwnd", $h)
    DllCall("user32.dll", "int", "SetForegroundWindow", "hwnd", $h)
    DllCall("user32.dll", "hwnd", "SetActiveWindow", "hwnd", $h)
    DllCall("user32.dll", "hwnd", "SetFocus", "hwnd", $h)

    If $fgTID <> 0 Then
        DllCall("user32.dll", "int", "AttachThreadInput", "dword", $curTID[0], "dword", $fgTID, "int", 0)
    EndIf

    WinWaitActive($h, "", 0.35)
EndFunc

; ---- 获取 d2r.exe 的顶层可见窗口列表 —— 按“进程启动时间”升序（≈任务栏左→右）----
Func _GetD2RList()
    Local $aWin = WinList()
    Local $handles[0]

    ; 收集：所有可见顶层、exe=d2r.exe、无 owner 的窗口
    For $i = 1 To $aWin[0][0]
        Local $h = $aWin[$i][1]
        If $h = 0 Then ContinueLoop
        If BitAND(WinGetState($h), 2) = 0 Then ContinueLoop ; 不可见
        Local $pid = WinGetProcess($h)
        Local $pname = _GetProcessNameByPID($pid)
        If StringLower($pname) <> "d2r.exe" Then ContinueLoop
        Local $hOwner = DllCall("user32.dll", "hwnd", "GetWindow", "hwnd", $h, "uint", 4) ; GW_OWNER
        If IsArray($hOwner) And $hOwner[0] <> 0 Then ContinueLoop

        _ArrayAdd($handles, $h)
    Next

    ; 关联到“进程启动时间”并按升序排序
    If UBound($handles) > 1 Then
        ; 构建二维数组： [n][0]=hwnd, [n][1]=CreationTime(数值)
        Local $table[UBound($handles)][2]
        For $k = 0 To UBound($handles) - 1
            Local $h = $handles[$k]
            Local $pid = WinGetProcess($h)
            Local $ct = _GetProcessCreationTicks($pid) ; 越小越早启动
            $table[$k][0] = $h
            $table[$k][1] = $ct
        Next
        ; 按第2列(创建时间)升序
        _ArraySort2DByCol($table, 1, True)
        ; 提取排好序的 hwnd 列
        Local $sorted[UBound($handles)]
        For $m = 0 To UBound($handles) - 1
            $sorted[$m] = $table[$m][0]
        Next
        Return $sorted
    EndIf

    Return $handles
EndFunc

; ——用 WMI 取进程 CreationDate，并转成可比较的 ticks（越小越早）——
Func _GetProcessCreationTicks($pid)
    Local $oWMI = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    If @error Then Return 0x7FFFFFFF
    Local $col = $oWMI.ExecQuery("SELECT CreationDate FROM Win32_Process WHERE ProcessId=" & $pid)
    For $obj In $col
        ; WMI 的 CreationDate 形如 20250913xxxxxx.xxxxxx+000
        Local $s = StringLeft($obj.CreationDate, 14) ; YYYYMMDDhhmmss
        ; 转成一个可比较的整数
        Return Number($s)
    Next
    Return 0x7FFFFFFF
EndFunc

; ——从 PID 反查进程名（沿用你之前的实现）——
Func _GetProcessNameByPID($pid)
    Local $plist = ProcessList()
    For $i = 1 To $plist[0][0]
        If $plist[$i][1] = $pid Then
            Return $plist[$i][0]
        EndIf
    Next
    Return ""
EndFunc

; ——二维数组按指定列排序（升序=True / 降序=False）——
Func _ArraySort2DByCol(ByRef $arr, $colIndex = 0, $asc = True)
    Local $n = UBound($arr)
    If $n < 2 Then Return
    ; 简单插入排序，避免依赖额外 UDF
    For $i = 1 To $n - 1
        Local $key0 = $arr[$i][0], $key1 = $arr[$i][1]
        Local $j = $i - 1
        While $j >= 0 And _
            ( ($asc And ($arr[$j][$colIndex] > $key1)) Or (Not $asc And ($arr[$j][$colIndex] < $key1)) )
            $arr[$j + 1][0] = $arr[$j][0]
            $arr[$j + 1][1] = $arr[$j][1]
            $j -= 1
        WEnd
        $arr[$j + 1][0] = $key0
        $arr[$j + 1][1] = $key1
    Next
EndFunc
