VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   10965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13560
   LinkTopic       =   "Form7"
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   10965
   ScaleWidth      =   13560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   1575
      Left            =   720
      TabIndex        =   11
      Top             =   7080
      Width           =   4815
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2640
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000E&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   5415
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   4815
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFC0C0&
         Height          =   735
         Left            =   2160
         TabIndex        =   10
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFC0C0&
         Height          =   735
         Left            =   2160
         TabIndex        =   8
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFC0C0&
         Height          =   735
         Left            =   2160
         TabIndex        =   6
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFC0C0&
         Height          =   735
         Left            =   2160
         TabIndex        =   4
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "   Email Id"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   9
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "  Phone No."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   7
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "   Address"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   5
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "     Name"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3135
      Left            =   6720
      TabIndex        =   1
      Top             =   4800
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5530
      _Version        =   393216
      BackColor       =   -2147483633
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
      Height          =   855
      Left            =   360
      Top             =   10320
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Details"
      BeginProperty Font 
         Name            =   "Brush Script Std"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   5175
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim con1 As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ros As ADODB.Recordset

Private Sub Command1_Click()
con1.Open
ros.LockType = adLockOptimistic
ros.Open "supplier", con1
ros.AddNew
ros.Fields(0) = Text1
ros.Fields(1) = Text2
ros.Fields(2) = Text3
ros.Fields(3) = Text4
ros.Update
MsgBox "Record Inserted"
    
Unload Me
End Sub




Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set ros = New ADODB.Recordset
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"

Set con1 = New ADODB.Connection
con1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"

rs.CursorLocation = adUseClient
rs.Open "Select * from supplier", con, adOpenKeyset, adLockPessimistic, adcmdtxt
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
Set rs = Nothing



End Sub

Private Sub Text3_Validate(CANCEL As Boolean)
    If (Text3.Text <> "") Then
        If Len(Text3.Text) <> 10 Then
        MsgBox "Contact no. must be 10 digits..!", vbExclamation, "Error"
        CANCEL = True
        End If
    End If
End Sub

Private Sub Text4_Validate(CANCEL As Boolean)

If (Text4.Text <> "") Then
Dim strTmp, strEmail As String, n As Long, sEXT As String
strEmail = Text4.Text
sEXT = Text4.Text
Do While InStr(1, sEXT, ".") <> 0
    sEXT = Right(sEXT, Len(sEXT) - InStr(1, sEXT, "."))
Loop

If strEmail = "" Then
MsgBox " You did not enter an email address..!"
    CANCEL = True
ElseIf InStr(1, strEmail, "@") = 0 Then
   MsgBox " Invalid Email Address : Your email address does not contain an @ sign..!"
   CANCEL = True
ElseIf InStr(1, strEmail, "@") = 1 Then
   MsgBox " Invalid Email Address :  Your @ sign can not be the first character in your email address..!"
CANCEL = True
   ElseIf InStr(1, strEmail, "@") = Len(strEmail) Then
   MsgBox " Invalid Email Address : @ sign can not be the last character in email address..!"
   CANCEL = True
ElseIf Len(strEmail) < 6 Then
   MsgBox " Invalid Email Address : Email address can not be shorter than 6 characters..!"
   CANCEL = True
End If
strTmp = Text4.Text
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

