VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "File"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.CheckBox b 
      Caption         =   "Hidden"
      Height          =   435
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CheckBox c 
      Caption         =   "Archive"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CheckBox a 
      Caption         =   "Read Only"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub a_Click()
Command1.Enabled = True
End Sub
Private Sub b_Click()
Command1.Enabled = True
End Sub
Private Sub c_Click()
Command1.Enabled = True
End Sub
Private Sub Command1_Click()
Set fso = CreateObject("Scripting.FileSystemObject")
Set att = fso.GetFile(CommonDialog1.FileName)
If a.Value = 1 And b.Value = 0 And c.Value = 0 Then
att.Attributes = 1
End If
If a.Value = 1 And b.Value = 1 And c.Value = 0 Then
att.Attributes = 3
End If
If a.Value = 1 And b.Value = 1 And c.Value = 1 Then
att.Attributes = 35
End If
If a.Value = 1 And b.Value = 0 And c.Value = 1 Then
att.Attributes = 33
End If
If a.Value = 0 And b.Value = 1 And c.Value = 1 Then
att.Attributes = 34
End If
If a.Value = 0 And b.Value = 1 And c.Value = 0 Then
att.Attributes = 2
End If
If a.Value = 0 And b.Value = 0 And c.Value = 1 Then
att.Attributes = 32
End If
If a.Value = 0 And b.Value = 0 And c.Value = 0 Then
att.Attributes = 0
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then
Else
Set fso = CreateObject("Scripting.FileSystemObject")
Set att = fso.GetFile(CommonDialog1.FileName)
Me.Caption = CommonDialog1.FileName
If att.Attributes = 1 Then
a.Value = 1
b.Value = 0
c.Value = 0
End If
If att.Attributes = 3 Then
a.Value = 1
b.Value = 1
c.Value = 0
End If
If att.Attributes = 35 Then
a.Value = 1
b.Value = 1
c.Value = 1
End If
If att.Attributes = 33 Then
a.Value = 1
b.Value = 0
c.Value = 1
End If
If att.Attributes = 34 Then
a.Value = 0
b.Value = 1
c.Value = 1
End If
If att.Attributes = 2 Then
a.Value = 0
b.Value = 1
c.Value = 0
End If
If att.Attributes = 32 Then
a.Value = 0
b.Value = 0
c.Value = 1
End If
If att.Attributes = 0 Then
a.Value = 0
b.Value = 0
c.Value = 0
End If
a.Enabled = True
b.Enabled = True
c.Enabled = True
Command1.Enabled = False
End If
End Sub
Private Sub Form_Load()
Me.Caption = "No File Loaded"
Command1.Enabled = False
a.Enabled = False
b.Enabled = False
c.Enabled = False
End Sub
