VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   10005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15360
   LinkTopic       =   "Form11"
   Picture         =   "Form10.frx":0000
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3000
      TabIndex        =   24
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   23
      Top             =   9960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Bill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   22
      Top             =   9960
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      Height          =   735
      Left            =   3000
      TabIndex        =   21
      Top             =   8880
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Calculate Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   20
      Top             =   8880
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   3000
      TabIndex        =   19
      Top             =   8160
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   3000
      TabIndex        =   16
      Top             =   6720
      Width           =   2415
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   3000
      TabIndex        =   15
      Top             =   6000
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   3000
      TabIndex        =   14
      Top             =   5280
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   2400
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3000
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   10440
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      RecordSource    =   "select distinct cname  from newcake"
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   14280
      Top             =   2400
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      RecordSource    =   $"Form10.frx":4D9CA
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
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Price per kg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   18
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Cake Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   17
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
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
      Left            =   840
      TabIndex        =   12
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Weight"
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
      Left            =   840
      TabIndex        =   11
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Email id"
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
      Left            =   840
      TabIndex        =   10
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No"
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
      Left            =   720
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   720
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Occasion"
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
      Left            =   720
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cust_Id"
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
      Left            =   720
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cust_Name"
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
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "    Purchase "
      BeginProperty Font 
         Name            =   "Brush Script Std"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   5295
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con, con1 As ADODB.Connection
Dim rs, ros As ADODB.Recordset


Private Sub Combo1_Click()
Set rs = New ADODB.Recordset
con.Open
rs.Open "Select * from customer where cust_name = '" & Combo1.Text & "'", con
'select price from newcake where cname='" & Text2.Text & " ' ", con

If Not rs.EOF Then
Me.Text1.Text = rs!cid
Me.Text2.Text = rs!occassion
Me.Text3.Text = rs!addr
Me.Text4.Text = rs!phone
Me.Text5.Text = rs!email
'Me.Text6.Text = rs!qty
'Me.Text10.Text = rs!Weight
'Me.Text9.Text = rs!price

rs.MoveNext
End If
rs.Close
con.Close

Set rs = Nothing
End Sub


Private Sub Combo2_Click()
Set rs = New ADODB.Recordset
con.Open
rs.Open "Select * from newcake where cname = '" & Combo2.Text & "'", con

If Not rs.EOF Then
'Me.Text7.Text = rs!flavour
'Me.Text8.Text = rs!Weight
Me.Text9.Text = rs!price
rs.MoveNext
End If
rs.Close
con.Close

Set rs = Nothing
End Sub

Private Sub Command1_Click()

Dim s As String
s = "Update newcake set stock = stock -" & Val(Text6) & " where cname = '" & Combo2.Text & "'"
con.Open
con.Execute s
con.Close
'MsgBox " Inventory Updated"

con1.Open
ros.Open "purchase", con1, adOpenDynamic, adLockOptimistic
ros.AddNew
ros.Fields(0) = Combo1.Text
ros.Fields(1) = Text1.Text
ros.Fields(2) = Text2.Text
ros.Fields(3) = Text3.Text
ros.Fields(4) = Text4.Text
ros.Fields(5) = Text5.Text
ros.Fields(6) = Text6.Text
ros.Fields(7) = Text10.Text
ros.Fields(8) = Combo2.Text
ros.Fields(9) = Text9.Text
ros.Fields(10) = Text11.Text
ros.Update
ros.Close
con1.Close
'MsgBox ("Details Updated Successfully")

'DataReport1.Show




Form11.Show
Form11.Label9.Caption = Me.Combo2.Text
Form11.Label10.Caption = Me.Text9.Text
Form11.Label11.Caption = Me.Text6.Text
Form11.Label12.Caption = Me.Text9.Text
Form11.Label13.Caption = Me.Text11.Text
Form11.Label15.Caption = Me.Text11.Text
Form11.Label21.Caption = Me.Combo1.Text
Form11.Label24.Caption = Me.Text4.Text
Form11.Label29.Caption = Me.Text10.Text

Unload Me




End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
Text11.Text = (Text6.Text * Text9.Text * Text10.Text)
End Sub

Private Sub Command4_Click()


End Sub

Private Sub Form_Load()

Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
rs.LockType = adLockOptimistic
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"
Adodc1.Refresh
With Adodc1.Recordset
Do Until .EOF
Combo1.AddItem ![cust_name]
.MoveNext
Loop

End With

Adodc2.Refresh
With Adodc2.Recordset
Do Until .EOF
Combo2.AddItem ![cname]
.MoveNext
Loop

End With



Set con1 = New ADODB.Connection
Set ros = New ADODB.Recordset
ros.LockType = adLockOptimistic
con1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"
End Sub






