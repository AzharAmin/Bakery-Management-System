VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10725
   ForeColor       =   &H80000011&
   LinkTopic       =   "Form1"
   Picture         =   "form1.frx":0000
   ScaleHeight     =   825
   ScaleMode       =   0  'User
   ScaleWidth      =   3952.985
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   13920
      Top             =   9960
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"
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
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "login"
      Height          =   495
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      DataField       =   "password"
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   "username"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cake Shop "
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   720
      TabIndex        =   6
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   960
      TabIndex        =   1
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
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
      Left            =   960
      TabIndex        =   0
      Top             =   3120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"


con.Open
rs.CursorType = adOpenDynamic
rs.Open "select * from loginn where username='" & Text1 & "' and password= '" & Text2 & "'", con
 
If rs.EOF <> True Then
Form2.Show
Unload Me

Else
MsgBox "Invalid Username/Password"
End If
rs.Close
con.Close
End Sub

Private Sub Command2_Click()
MsgBox "Are you sure you want to exit the system", vbCritical, "Exit"

Unload Me

End Sub

