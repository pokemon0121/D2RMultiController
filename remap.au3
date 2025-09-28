HotKeySet("{PGUP}", "_GoHome")
HotKeySet("{PGDN}", "_GoEnd")

While 1
    Sleep(50)
WEnd

Func _GoHome()
    Send("{HOME}")
EndFunc

Func _GoEnd()
    Send("{END}")
EndFunc
