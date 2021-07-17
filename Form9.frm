VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   9615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17160
   LinkTopic       =   "Form10"
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   9615
   ScaleWidth      =   17160
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   1575
      Left            =   360
      TabIndex        =   18
      Top             =   6960
      Width           =   6855
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   4200
         TabIndex        =   20
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Confirm"
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5415
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   6975
      Begin VB.Frame Frame1 
         Caption         =   "OCCASION"
         Height          =   1215
         Left            =   360
         TabIndex        =   13
         Top             =   3600
         Width           =   6255
         Begin VB.OptionButton Option4 
            Caption         =   "party"
            Height          =   495
            Left            =   5040
            TabIndex        =   17
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Annversary"
            Height          =   495
            Left            =   3240
            TabIndex        =   16
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Birthday"
            Height          =   495
            Left            =   1680
            TabIndex        =   15
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Wedding"
            Height          =   495
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   1800
         TabIndex        =   12
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   1800
         TabIndex        =   10
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   1800
         TabIndex        =   8
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   1800
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Occasion"
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Email id"
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Phone No."
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Addr"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Cust_name"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   15240
      Top             =   8400
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   688
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select distinct cname from newcake"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11160
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "              Customer Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   21
      Top             =   240
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "Cust_id"
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
      Left            =   9600
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim max As Integer

Private Sub Command1_Click()
con.Open
rs.LockType = adLockOptimistic
rs.Open "customer", con
rs.AddNew
rs.Fields(1) = Text2
rs.Fields(2) = Text3
rs.Fields(3) = Text4
rs.Fields(4) = Text5
rs.Fields(5) = Text6
'rs.Fields(5) = Combo1.Text
'rs.Fields(6) = Text6
'rs.Fields(7) = Text7

rs.Update
MsgBox "Order Stored Successfully"
    

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
rs.LockType = adLockOptimistic
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"
'Adodc1.Refresh
'With Adodc1.Recordset
'Do Until .EOF
'Combo1.AddItem ![cname]
'.MoveNext
'Loop

'End With

getMax
Text1.Text = max


End Sub

Private Sub Option1_Click()
If Option1.Value = True Then Text6.Text = "Wedding"

End Sub

Private Sub Option2_Click()
If Option2.Value = True Then Text6.Text = "Birthday"

End Sub

Private Sub Option3_Click()
If Option3.Value = True Then Text6.Text = "Annversary"

End Sub

Private Sub Option4_Click()
If Option4.Value = True Then Text6.Text = "Annversary"

End Sub

Private Sub Text4_Validate(CANCEL As Boolean)
    If (Text4.Text <> "") Then
        If Len(Text4.Text) <> 10 Then
        MsgBox "Phone no. must be 10 digits", vbExclamation, "Error"
        CANCEL = True
        End If
    End If
End Sub

Private Sub Text5_Validate(CANCEL As Boolean)

If (Text5.Text <> "") Then
Dim strTmp, strEmail As String, n As Long, sEXT As String
strEmail = Text5.Text
sEXT = Text5.Text
Do While InStr(1, sEXT, ".") <> 0
    sEXT = Right(sEXT, Len(sEXT) - InStr(1, sEXT, "."))
Loop

If strEmail = "" Then
MsgBox " You did not enter an email address..!"
    CANCEL = True
ElseIf InStr(1, strEmail, "@") = 0 Then
   MsgBox " Invalid Email Address : Your email address does not contain an @ sign"
   CANCEL = True
ElseIf InStr(1, strEmail, "@") = 1 Then
   MsgBox " Invalid Email Address :  Your @ sign can not be the first character in your email address"
CANCEL = True
   ElseIf InStr(1, strEmail, "@") = Len(strEmail) Then
   MsgBox " Invalid Email Address : @ sign can not be the last character in email address"
   CANCEL = True
ElseIf Len(strEmail) < 6 Then
   MsgBox " Invalid Email Address : Email address can not be shorter than 6 characters"
   CANCEL = True
End If
strTmp = Text5.Text
Do While InStr(1, strTmp, "@") <> 0
   n = 1
   strTmp = Right(strTmp, Len(strTmp) - InStr(1, strTmp, "@"))
Loop
If n > 1 Then
   MsgBox " Invalid Email Address : More than 1 @ sign in email address"
   CANCEL = True
End If
End If
End Sub
Private Sub getMax()
con.Open
rs.Open "Select max(cid) from customer", con, adOpenDynamic
If rs.EOF <> True And rs.BOF <> True Then
max = CInt(rs.Fields(0)) + 1
Else
max = 1
End If
rs.Close
con.Close
End Sub

