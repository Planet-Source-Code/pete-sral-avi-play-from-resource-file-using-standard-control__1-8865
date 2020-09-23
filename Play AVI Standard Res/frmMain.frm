VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Play Resource AVI with Standard Control"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin VB.ListBox lstAVI 
      Height          =   1815
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1931
      _Version        =   393216
      AutoPlay        =   -1  'True
      Center          =   -1  'True
      FullWidth       =   281
      FullHeight      =   73
   End
   Begin VB.Label Label1 
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   4215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Function GetID(psString As String) As Long


Dim iDelim As Integer
Dim sVal As String

    iDelim = InStr(1, psString, "=")
    
    If iDelim <> 0 Then
        sVal = Trim$(Mid$(psString, iDelim + 1))
        GetID = Val(sVal)
    Else
        GetID = 0
    End If
    
End Function

Private Sub cmdPlay_Click()


Dim lResID As Long

    If InIDE() Then
        MsgBox "The AVI will not appear in design mode, only at runtime!", vbInformation
        Exit Sub
    End If
    
    If lstAVI.ListIndex = -1 Then
        MsgBox "Please select AVI from list box!", vbExclamation
        Exit Sub
    End If
    
        
    lResID = GetID(lstAVI.List(lstAVI.ListIndex))
    
    If lResID <> 0 Then
               
        'refresh animation
        ClearAnim Animation1
        
        'load and play avi
        LoadResAVI Animation1, lResID
        
        
        'use the animation1 control's methods, properties as normal
        'the only problem is .Play - it errors (that is why AutoPlay = True)
    End If
    
    
    
    
End Sub

Private Sub Form_Load()
    With Label1
        .Caption = "Sample code to play an AVI File from a resource file using the standard Windows Animation Control"
        .Caption = .Caption & " The API code was provided by: Mattias Sjögren (MCSE) - VB+ http://hem.spray.se/mattias.sjogren/"
        .Caption = .Caption & " Please visit his site!" & Chr(13)
        .Caption = .Caption & " For more AVI files visit http://pjs-inc.com/vb-avi"
    End With
    
    'load listbox
    With lstAVI
        .AddItem "Globe = 100"
        .AddItem "Busy = 101"
        .AddItem "CdSpin = 102"
        .AddItem "Defrag = 103"
        .AddItem "Download = 104"
        .AddItem "FileCopy = 105"
        .AddItem "FileDelete = 106"
        .AddItem "FileNuke = 108"
        .AddItem "FileMove = 107"
        .AddItem "TrashNuke = 115"
        .AddItem "FindComputer = 109"
        .AddItem "FindFile = 110"
        .AddItem "FindFolder = 111"
        .AddItem "Watch = 116"
        .AddItem "InetDownload = 112"
        .AddItem "InetSend = 113"
        .AddItem "PrinterPrint = 114"
    End With
    
    'if False will get error
    Animation1.AutoPlay = True
    
End Sub
