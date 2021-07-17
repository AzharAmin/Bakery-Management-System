VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10500
   LinkTopic       =   "Form4"
   ScaleHeight     =   7755
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   12
      Top             =   6720
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   600
      Top             =   9480
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\final project\newcake.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\final project\newcake.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   11
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   2520
      TabIndex        =   10
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   2520
      TabIndex        =   8
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2520
      TabIndex        =   6
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   2760
      Width           =   1575
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
      Left            =   600
      TabIndex        =   9
      Top             =   5760
      Width           =   1575
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
      Left            =   600
      TabIndex        =   7
      Top             =   5040
      Width           =   1575
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
      Left            =   600
      TabIndex        =   5
      Top             =   4320
      Width           =   1575
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
      Left            =   600
      TabIndex        =   3
      Top             =   3600
      Width           =   1575
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
      Left            =   600
      TabIndex        =   1
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "      New Cake Details"
      BeginProperty Font 
         Name            =   "Brush Script Std"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4800
      TabIndex        =   0
      Top             =   360
      Width           =   8175
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con, con1 As ADODB.Connection
Dim rs, ros As ADODB.Recordset
Dim str As String

Private Sub Command2_Click()
con.Open
rs.Open "newcake", con
rs.AddNew
rs.Fields(0) = Text1
rs.Fields(1) = Text2
rs.Fields(2) = Text3
rs.Fields(3) = Text4
rs.Fields(4) = Text5
rs.Update
rs.Close
con.Close
MsgBox ("Details Updated Successfully")
Unload Me


End Sub

Private Sub Command3_Click()
Unload Me

End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
rs.LockType = adLockOptimistic
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"

End Sub
