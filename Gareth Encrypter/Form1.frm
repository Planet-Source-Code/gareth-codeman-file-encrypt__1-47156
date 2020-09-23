VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "File Encrypter v1.0"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   6135
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Extra"
      Height          =   975
      Left            =   3600
      TabIndex        =   14
      Top             =   3000
      Width           =   2415
      Begin VB.OptionButton Option8 
         Caption         =   "Normal"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Read Only"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Hidden"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   2775
      Left            =   3600
      TabIndex        =   7
      Top             =   120
      Width           =   2415
      Begin VB.OptionButton Option1 
         Caption         =   "Recycle bin"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Briefcase"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "My Computer"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Control Panel"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton Option5 
         Caption         =   "History"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Decrypt"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Encrypt"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Shape Shape1 
      Height          =   2895
      Left            =   120
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label shity 
      Caption         =   "Save As"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label Label2 
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label Label1 
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next 'to stop any errors when choosing the file
cd1.ShowOpen 'to select a file
Label1.Caption = "FileName: " & cd1.FileName  'adds name & path of file to label
Label2.Caption = "FileTitle: " & cd1.FileTitle 'adds name of file to label
Label3.Caption = "FileSize: " & FileLen(cd1.FileTitle) & " bytes" 'adds size of file to label
End Sub

Private Sub Command2_Click()
On Error Resume Next

'if any option6 option7 or option8 values are true it will change the file attributes

'make file hidden
If Option6.Value = True Then
SetAttr cd1.FileName, vbHidden
End If

'make file readonly
If Option7.Value = True Then
SetAttr cd1.FileName, vbReadOnly
End If

'make file normal
If Option8.Value = True Then
SetAttr cd1.FileName, vbNormal
End If


'change file into recycle bin
If Option1.Value = True Then
Name cd1.FileTitle As "Recycle Bin.{645FF040-5081-101B-9F08-00AA002F954E}"
End If

'change file into briefcase
If Option2.Value = True Then
Name cd1.FileTitle As "Briefcase.{85BBD920-42A0-1069-A2E4-08002B30309D}"
End If

'change file into my computer
If Option3.Value = True Then
Name cd1.FileTitle As "My Computer.{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
End If

'change file into control panel
If Option4.Value = True Then
Name cd1.FileTitle As "Control Panel.{21EC2020-3AEA-1069-A2DD-08002B30309D}"
End If

'change file into history folder
If Option5.Value = True Then
Name cd1.FileTitle As "History.{FF393560-C2A7-11CF-BFF4-444553540000}"
End If





End Sub

Private Sub Command3_Click()
On Error Resume Next 'to stop errors if no file is selected
Name cd1.FileTitle As Text1.Text 'renaming your encrypted file
End Sub

