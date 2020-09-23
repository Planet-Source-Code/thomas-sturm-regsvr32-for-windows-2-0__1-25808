VERSION 5.00
Begin VB.Form frmBrowse 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Browse for File"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Frame frmFilesAndFolders 
      Caption         =   "Folders && Files :"
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.FileListBox filFiles 
         Height          =   3405
         Left            =   3240
         Pattern         =   "*.ocx;*.dll"
         TabIndex        =   3
         Top             =   350
         Width           =   3135
      End
      Begin VB.DriveListBox drvDrives 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   3425
         Width           =   2895
      End
      Begin VB.DirListBox dirDirectories 
         Height          =   3015
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Label lblFilename 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   4200
      Width           =   6615
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
frmMain.txtFilename.Text = lblFilename.Caption
Unload Me
End Sub

Private Sub dirDirectories_Change()
filFiles.Path = dirDirectories.Path
End Sub

Private Sub drvDrives_Change()
dirDirectories.Path = drvDrives.Drive
filFiles.Path = dirDirectories.Path
End Sub

Private Sub filFiles_Click()
If Right$(filFiles.Path, 1) = "\" Then
    lblFilename.Caption = filFiles.Path & filFiles.FileName
Else
    lblFilename.Caption = filFiles.Path & "\" & filFiles.FileName
End If
cmdOK.Enabled = True
End Sub

Private Sub Form_Load()
drvDrives.Drive = "C:\"
dirDirectories.Path = "C:\"
filFiles.Path = "C:\"
End Sub
