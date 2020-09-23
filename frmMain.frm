VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "RegSvr32 for Windows 2.0"
   ClientHeight    =   3456
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5772
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3456
   ScaleWidth      =   5772
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Reg/Unreg from Shell"
      Height          =   495
      Left            =   3000
      TabIndex        =   9
      Top             =   2160
      Width           =   2412
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   5535
   End
   Begin VB.Frame frmFilename 
      Caption         =   "Options :"
      Height          =   2892
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton cmdAddToShell 
         Caption         =   "Add Reg/Unreg to Shell"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   2412
      End
      Begin VB.CommandButton cmdUnreg 
         Caption         =   "UNREGISTER FILE"
         Height          =   495
         Left            =   2880
         TabIndex        =   5
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CommandButton cmdReg 
         Caption         =   "REGISTER FILE"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse"
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtFilename 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label lblCopyright 
         Alignment       =   2  'Zentriert
         Caption         =   "RegSvr32 for Windows 2.0 (c) 2001 Thomas Sturm"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   3735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   30
         X2              =   5520
         Y1              =   1340
         Y2              =   1340
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   30
         X2              =   5520
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblFilename 
         Caption         =   "Filename :"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddToShell_Click()
Dim RegSvr32ForWindowsFileName As String
If Right$(App.Path, 1) = "\" Then
    RegSvr32ForWindowsFileName = App.Path & App.EXEName & ".exe"
Else
    RegSvr32ForWindowsFileName = App.Path & "\" & App.EXEName & ".exe"
End If

CreateNewKey HKEY_CLASSES_ROOT, "dllfile\shell\reg\command"
CreateNewKey HKEY_CLASSES_ROOT, "dllfile\shell\unreg\command"
CreateNewKey HKEY_CLASSES_ROOT, "ocxfile\shell\reg\command"
CreateNewKey HKEY_CLASSES_ROOT, "ocxfile\shell\unreg\command"

SetKeyValue HKEY_CLASSES_ROOT, "dllfile\shell\reg", (Standard), "Register", REG_SZ
SetKeyValue HKEY_CLASSES_ROOT, "dllfile\shell\unreg", (Standard), "Unregister", REG_SZ
SetKeyValue HKEY_CLASSES_ROOT, "dllfile\shell\reg\command", (Standard), Chr(34) & RegSvr32ForWindowsFileName & Chr(34) & " " & Chr(34) & "%1" & Chr(34), REG_SZ
SetKeyValue HKEY_CLASSES_ROOT, "dllfile\shell\unreg\command", (Standard), Chr(34) & RegSvr32ForWindowsFileName & Chr(34) & " /u " & Chr(34) & "%1" & Chr(34), REG_SZ

SetKeyValue HKEY_CLASSES_ROOT, "ocxfile\shell\reg", (Standard), "Register", REG_SZ
SetKeyValue HKEY_CLASSES_ROOT, "ocxfile\shell\unreg", (Standard), "Unregister", REG_SZ
SetKeyValue HKEY_CLASSES_ROOT, "ocxfile\shell\reg\command", (Standard), Chr(34) & RegSvr32ForWindowsFileName & Chr(34) & " " & Chr(34) & "%1" & Chr(34), REG_SZ
SetKeyValue HKEY_CLASSES_ROOT, "ocxfile\shell\unreg\command", (Standard), Chr(34) & RegSvr32ForWindowsFileName & Chr(34) & " /u " & Chr(34) & "%1" & Chr(34), REG_SZ

End Sub

Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdReg_Click()
If txtFilename.Text = "" Then Exit Sub
If FileExist(txtFilename.Text) Then
    If RegServer(txtFilename.Text) = True Then
        MsgBox txtFilename.Text & " was correctly registered.", vbInformation, "Success"
    Else
        MsgBox "Failure registering " & txtFilename.Text & ".", vbCritical, "Failure"
    End If
Else
    MsgBox "Specified File not found !", vbCritical, "Error"
End If
End Sub

Private Sub cmdUnreg_Click()
If txtFilename.Text = "" Then Exit Sub
If FileExist(txtFilename.Text) Then
    If UnRegServer(txtFilename.Text) = True Then
        MsgBox txtFilename.Text & " was correctly unregistered.", vbInformation, "Success"
    Else
        MsgBox "Failure unregistering " & txtFilename.Text & ".", vbCritical, "Failure"
    End If
Else
    MsgBox "Specified File not found !", vbCritical, "Error"
End If
End Sub

Function FileExist(sFileName As String) As Boolean
FileExist = IIf(Dir(sFileName) <> "", True, False)
End Function

Private Sub Command1_Click()
Load frmBrowse
frmBrowse.Show vbModal, Me
End Sub

Private Sub Command2_Click()
DeleteKey HKEY_CLASSES_ROOT, "dllfile\shell\reg\command"
DeleteKey HKEY_CLASSES_ROOT, "dllfile\shell\unreg\command"
DeleteKey HKEY_CLASSES_ROOT, "dllfile\shell\reg"
DeleteKey HKEY_CLASSES_ROOT, "dllfile\shell\unreg"

DeleteKey HKEY_CLASSES_ROOT, "ocxfile\shell\reg\command"
DeleteKey HKEY_CLASSES_ROOT, "ocxfile\shell\unreg\command"
DeleteKey HKEY_CLASSES_ROOT, "ocxfile\shell\reg"
DeleteKey HKEY_CLASSES_ROOT, "ocxfile\shell\unreg"
End Sub

Private Sub Form_Load()
Dim FileToReg As String
If Command$ <> "" Then
    Reg = Register(Command$)
    FileToReg = GetFileNameFromCommandLine(Command$)
    txtFilename.Text = Trim$(Mid$(FileToReg, 2, Len(FileToReg) - 2))
    If Reg = True Then
        cmdReg_Click
    ElseIf Reg = False Then
        cmdUnreg_Click
    End If
    End
End If
End Sub
