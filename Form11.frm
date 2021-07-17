VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form11 
   ClientHeight    =   9645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17880
   LinkTopic       =   "Form11"
   ScaleHeight     =   9645
   ScaleWidth      =   17880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   615
      Left            =   21360
      TabIndex        =   35
      Top             =   0
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   6000
      TabIndex        =   0
      Top             =   4560
      Width           =   10935
      Begin VB.Frame Frame2 
         Height          =   3015
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   10455
         Begin VB.Label Label29 
            BackColor       =   &H80000014&
            Height          =   495
            Left            =   6360
            TabIndex        =   30
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label13 
            BackColor       =   &H80000014&
            Height          =   495
            Left            =   8280
            TabIndex        =   19
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label12 
            BackColor       =   &H80000014&
            Height          =   495
            Left            =   4680
            TabIndex        =   18
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label11 
            BackColor       =   &H80000014&
            Height          =   495
            Left            =   3120
            TabIndex        =   17
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label10 
            BackColor       =   &H80000014&
            Height          =   495
            Left            =   1560
            TabIndex        =   16
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label9 
            BackColor       =   &H80000014&
            Height          =   495
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label17 
            BackColor       =   &H80000014&
            Height          =   495
            Left            =   8280
            TabIndex        =   14
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "SGST(9%)"
            Height          =   495
            Left            =   6360
            TabIndex        =   13
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label16 
            BackColor       =   &H80000014&
            Height          =   495
            Left            =   8280
            TabIndex        =   12
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "CGST(9%)"
            Height          =   495
            Left            =   6360
            TabIndex        =   11
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label15 
            BackColor       =   &H80000014&
            Height          =   495
            Left            =   8280
            TabIndex        =   10
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label14 
            Caption         =   "       Net Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5400
            TabIndex        =   9
            Top             =   960
            Width           =   2415
         End
      End
      Begin VB.Label Label28 
         Caption         =   "Weight(in kg)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   29
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8640
         TabIndex        =   7
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Grand Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         TabIndex        =   6
         Top             =   4920
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "   Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8520
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Price(per kg)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5040
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "  Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "     Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Particulars"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Label Label33 
      BackColor       =   &H8000000E&
      Caption         =   "                    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16440
      TabIndex        =   34
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label32 
      BackColor       =   &H8000000E&
      Caption         =   "               "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16440
      TabIndex        =   33
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label31 
      Caption         =   "    Time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   18960
      TabIndex        =   32
      Top             =   10920
      Width           =   1215
   End
   Begin VB.Label Label30 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   20160
      TabIndex        =   31
      Top             =   10920
      Width           =   1935
   End
   Begin VB.Label Label27 
      Caption         =   "    Thank you                                                                                                     Visit Again"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   28
      Top             =   10680
      Width           =   11175
   End
   Begin VB.Label Label26 
      Caption         =   $"Form11.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7440
      TabIndex        =   27
      Top             =   1200
      Width           =   5655
   End
   Begin VB.Label Label25 
      Caption         =   "          Cake Shop"
      BeginProperty Font 
         Name            =   "Brush Script Std"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   26
      Top             =   360
      Width           =   7575
   End
   Begin VB.Label Label24 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   25
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Label Label23 
      Caption         =   "Contact No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   24
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label Label22 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15120
      TabIndex        =   23
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label21 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   22
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label Label20 
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   21
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label19 
      Caption         =   "Bill No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15120
      TabIndex        =   20
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim max As Integer


Private Sub Command1_Click()
Command1.Visible = False
CommonDialog1.ShowPrinter
con.Open
    rs.Open "bill", con, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs.Fields(1) = Label21.Caption
    rs.Fields(2) = Label24.Caption
    rs.Fields(3) = Label32.Caption
    rs.Fields(4) = Label9.Caption
    rs.Fields(5) = Label10.Caption
    rs.Fields(6) = Label11.Caption
    rs.Fields(7) = Label12.Caption
    rs.Fields(8) = Label29.Caption
    rs.Fields(9) = Label13.Caption
    rs.Fields(10) = Label15.Caption
    rs.Fields(11) = Label16.Caption
    rs.Fields(12) = Label17.Caption
    rs.Fields(13) = Label18.Caption
   
    rs.Update
    rs.Close
    con.Close
    MsgBox ("Records updated successfully")
    Unload Me
Command1.Visible = True
End Sub
'Private Sub Command2_Click()
'con.Open
 '   rs.Open "bill", con, adOpenDynamic, adLockOptimistic
  '  rs.AddNew
   ' rs.Fields(1) = Label21.Caption
'    rs.Fields(2) = Label24.Caption
 '   rs.Fields(3) = Label32.Caption
  '  rs.Fields(4) = Label9.Caption
   ' rs.Fields(5) = Label10.Caption
    'rs.Fields(6) = Label11.Caption
'    rs.Fields(7) = Label12.Caption
'    rs.Fields(8) = Label29.Caption
 '   rs.Fields(9) = Label13.Caption
  '  rs.Fields(10) = Label15.Caption
   ' rs.Fields(11) = Label16.Caption
'    rs.Fields(12) = Label17.Caption
 '   rs.Fields(13) = Label18.Caption
  
'    rs.Update
 '   rs.Close
  '  con.Close
   ' MsgBox ("Records updated successfully")
   ' Unload Me
    
'End Sub

'Private Sub Form_Click()
'Label16.Caption = ((9 * Label15.Caption) / 100)
'Label17.Caption = ((9 * Label15.Caption) / 100)
'Label18.Caption = (Val(Label15.Caption) + Val(Label16.Caption) + Val(Label17.Caption))

'End Sub

Private Sub Form_Load()
Label32.Caption = Date
Label30.Caption = Time




Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
rs.LockType = adLockOptimistic
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"

getMax
Label33.Caption = max





End Sub

Private Sub getMax()
con.Open
rs.Open "Select max(bill_no) from bill", con, adOpenDynamic
If rs.EOF <> True And rs.BOF <> True Then
max = CInt(rs.Fields(0)) + 1
Else
max = 1
End If
rs.Close
con.Close
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = ((9 * Label15.Caption) / 100)
Label17.Caption = ((9 * Label15.Caption) / 100)
Label18.Caption = (Val(Label15.Caption) + Val(Label16.Caption) + Val(Label17.Caption))
End Sub


