VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   10170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17805
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   10170
   ScaleWidth      =   17805
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   9375
      Left            =   8640
      TabIndex        =   1
      Top             =   1560
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   16536
      _Version        =   393216
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   67
      WrapCellPointer =   -1  'True
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
         Size            =   13.5
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
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   735
      Left            =   14760
      TabIndex        =   2
      Top             =   5520
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "    Stock Details"
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
      Left            =   5280
      TabIndex        =   0
      Top             =   480
      Width           =   5535
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim con1 As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ros As ADODB.Recordset


Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"
Set con1 = New ADODB.Connection
con1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"

rs.CursorLocation = adUseClient
rs.Open "Select cname,stock from newcake", con, adOpenKeyset, adLockPessimistic, adcmdtxt
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
Set rs = Nothing

End Sub

