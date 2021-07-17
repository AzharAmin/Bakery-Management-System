VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   8955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14355
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   14355
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   480
      TabIndex        =   13
      Top             =   6360
      Width           =   7695
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5280
         TabIndex        =   17
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1680
         TabIndex        =   16
         Top             =   1560
         Width           =   4215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Update Selling Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2760
         TabIndex        =   15
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Add Selling price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "DETAILS"
      Height          =   4575
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   7695
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000E&
         Height          =   495
         Left            =   1920
         TabIndex        =   19
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Cost Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label10 
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
         Left            =   1920
         TabIndex        =   12
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Flavour"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Cake Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label7 
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
         Left            =   1920
         TabIndex        =   6
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label8 
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
         Left            =   1920
         TabIndex        =   5
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label9 
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
         Left            =   1920
         TabIndex        =   4
         Top             =   3120
         Width           =   1935
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3015
      Left            =   8640
      TabIndex        =   1
      Top             =   2640
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5318
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   16440
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
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
      RecordSource    =   "Select distinct cname from order11"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Cake Details"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   5535
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con, con1, con2, con3 As ADODB.Connection
Dim rs, ros As ADODB.Recordset
Dim str As String

Private Sub Combo1_Click()
Set rs = New ADODB.Recordset
con.Open
rs.Open "Select * from order11 where cname = '" & Combo1.Text & "'", con
If Not rs.EOF Then
'Me.Text1.Text = rs!flavour
Me.Label7.Caption = rs!flavour
Me.Label12.Caption = rs!cprice
'Me.Text3.Text = rs!Weight
Me.Label9.Caption = rs!Weight
'Me.Text4.Text = rs!stock
Me.Label10.Caption = rs!stock
rs.MoveNext
End If
rs.Close
con.Close

Set rs = Nothing

End Sub

Private Sub Command1_Click()
Label8.Caption = InputBox("Enter new selling price")

Dim s As String
s = "Update newcake set newcake.flavour='" + Label7.Caption + "',newcake.cprice=" + Label12.Caption + ",newcake.price=" + Label8.Caption + ",newcake.weight='" + Label9.Caption + "',newcake.stock =" + Label10.Caption + " where cname='" + Combo1.Text + "'"
con.Open
con.Execute s
con.Close
MsgBox "Updated"
End Sub

Private Sub Command2_Click()
'Dim str As String
Dim s1 As String
'str = InputBox("Enter cake name")

s1 = "delete from newcake where cname = '" & Combo1.Text & "'"
con.Open
con.Execute s1
Combo1.Text = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""
Label12.Caption = ""
MsgBox "Details Deleted"
con.Close


End Sub

Private Sub Command3_Click()
Unload Me

End Sub

Private Sub Command4_Click()
Label8.Caption = InputBox("Enter Selling Price")

con1.Open
ros.LockType = adLockOptimistic
ros.Open "newcake", con1
ros.AddNew
ros.Fields(0) = Combo1.Text
ros.Fields(1) = Label7.Caption
ros.Fields(2) = Label12.Caption
ros.Fields(3) = Label8.Caption
ros.Fields(4) = Label9.Caption
ros.Fields(5) = Label10.Caption

ros.Update
ros.Close
con1.Close
MsgBox ("Details Updated Successfully")
'Unload Me

End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set con1 = New ADODB.Connection
Set ros = New ADODB.Recordset
rs.LockType = adLockOptimistic
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"
con1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"
Adodc1.Refresh
With Adodc1.Recordset
Do Until .EOF
Combo1.AddItem ![cname]
.MoveNext
Loop

End With


Set con3 = New ADODB.Connection
Set rs = New ADODB.Recordset
Set ros = New ADODB.Recordset
con3.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"

Set con2 = New ADODB.Connection
con2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"

rs.CursorLocation = adUseClient
'rs.Open "Select * from newcake", con3, adOpenKeyset, adLockPessimistic, adcmdtxt
'Set DataGrid1.DataSource = rs
'DataGrid1.Refresh
'Set rs = Nothing



End Sub

