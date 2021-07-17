VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BackColor       =   &H8000000B&
   Caption         =   "Form2"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8685
   LinkTopic       =   "Form2"
   Picture         =   "form2.frx":0000
   ScaleHeight     =   5040
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   170
      Left            =   5280
      Top             =   8160
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   855
      Left            =   6000
      TabIndex        =   0
      Top             =   1680
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   1508
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "        Loading... "
      BeginProperty Font 
         Name            =   "Brush Script MT"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6600
      TabIndex        =   1
      Top             =   360
      Width           =   7215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
If ProgressBar1.Value = ProgressBar1.max Then
Form3.Show
Unload Me
Else: ProgressBar1.Value = ProgressBar1.Value + 10
End If

End Sub

