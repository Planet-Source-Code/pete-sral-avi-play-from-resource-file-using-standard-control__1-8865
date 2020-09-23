Attribute VB_Name = "modStdAVI"
Option Explicit
'API Code provided by:
'Mattias Sjögren (MCSE) - mattiass@hem.passagen.se
'    VB+ http://hem.spray.se/mattias.sjogren/

'For more AVI's check out http://pjs-inc.com/vb-avi

Const WM_USER = &H400&
Const ACM_OPEN = WM_USER + 100&
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Public Enum AnimationType
    Globe = 100
    Busy = 101
    CdSpin = 102
    Defrag = 103
    Download = 104
    FileCopy = 105
    FileDelete = 106
    FileNuke = 108
    FileMove = 107
    TrashNuke = 115
    FindComputer = 109
    FindFile = 110
    FindFolder = 111
    Watch = 116
    InetDownload = 112
    InetSend = 113
    PrinterPrint = 114
End Enum

Public Function InIDE() As Boolean
    On Error GoTo InIDEError
    InIDE = False
    Debug.Print 1 / 0
    Exit Function

InIDEError:
    InIDE = True
    Exit Function
End Function

Public Sub LoadResAVI(pCtrlAnimation As Animation, pEnumAnimType As AnimationType)

    SendMessage pCtrlAnimation.hwnd, ACM_OPEN, ByVal App.hInstance, ByVal pEnumAnimType
    
End Sub

Public Sub ClearAnim(pCtrlAnimation As Animation)

On Error Resume Next
    
    'clear previous animation
    With pCtrlAnimation
        .AutoPlay = False
        .Close
        .AutoPlay = True
    End With
    
End Sub
