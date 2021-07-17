VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   9015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14280
   LinkTopic       =   "Form8"
   Picture         =   "Form8.frx":0000
   ScaleHeight     =   9015
   ScaleWidth      =   14280
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   6240
      TabIndex        =   12
      Top             =   3360
      Width           =   10215
      Begin MSFlexGridLib.MSFlexGrid fg 
         Height          =   3015
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5318
         _Version        =   393216
         Cols            =   4
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   960
      Top             =   8160
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      RecordSource    =   "select distinct sname from supplier"
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
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   11
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   10
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2760
      TabIndex        =   8
      Top             =   5760
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2760
      TabIndex        =   7
      Top             =   4920
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2760
      TabIndex        =   6
      Top             =   4080
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2760
      TabIndex        =   5
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Email Id"
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
      Left            =   360
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No."
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
      Left            =   360
      TabIndex        =   3
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   360
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   360
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Dealer Details"
      BeginProperty Font 
         Name            =   "Brush Script Std"
         Size            =   27.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   4695
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim con2, con3 As ADODB.Connection
Dim VID As Integer


Private Sub Combo1_Click()
Set rs = New ADODB.Recordset
con.Open
rs.Open "Select * from supplier where sname = '" & Combo1.Text & "'", con
If Not rs.EOF Then
Me.Text1.Text = rs!addr
Me.Text2.Text = rs!phone
Me.Text3.Text = rs!email
rs.MoveNext
End If
rs.Close
con.Close

Set rs = Nothing

End Sub

Private Sub Command1_Click()
Dim s As String
s = "Update supplier set supplier.addr='" + Text1.Text + "',supplier.phone=" + Text2.Text + ",supplier.email='" + Text3.Text + "' where sname='" + Combo1.Text + "'"
con.Open
con.Execute s
con.Close
MsgBox "Updated Successfully"
End Sub

Private Sub Command2_Click()
Dim s1 As String
con.Open
s1 = "delete from supplier where sname = '" & Combo1.Text & "'"
con.Execute s1
Combo1.Text = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
MsgBox "Record Deleted"
con.Close

End Sub

Private Sub Command3_Click()
Unload Me

End Sub

'Private Sub Command4_Click()
'If (VID = 0) Then
 '    MsgBox "Select a supplier to View..!", vbInformation, "Select name"
'Else
 '    VID = (fg.Text)
  '   Dim RSS As String
   '  RSS = "select * from supplier where sname = ' " & VID & " ' "
    ' con.Open
    ' Set rs = New ADODB.Recordset
    ' rs.Open RSS, con, adOpenDynamic, adLockOptimistic
    ' If (rs.BOF Or rs.EOF) Then
     '   MsgBox "No Record Found"
      '  rs.Close
      '  con.Close
     'Else
    ' rs.MoveFirst
    '    Combo1.Text = rs.Fields(0)
     '   Text1.Text = rs.Fields(1)
     '   Text2.Text = rs.Fields(2)
      '  Text3.Text = rs.Fields(3)
        
   '  rs.Close
   '  con.Close
    ' End If
'End If
'End Sub

Private Sub fg_Click()
VID = CInt(fg.Text)
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set con2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset
rs.LockType = adLockOptimistic
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"
con2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"


MAINR

Adodc1.Refresh
With Adodc1.Recordset
Do Until .EOF
Combo1.AddItem ![sname]
.MoveNext
Loop

End With

'Set con3 = New ADODB.Connection
'Set rs = New ADODB.Recordset
'Set ros = New ADODB.Recordset
'con3.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"

'Set con2 = New ADODB.Connection
'con2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Azhar\Desktop\newwwwwww\loginn.mdb;Persist Security Info=False"

'rs.CursorLocation = adUseClient
'rs.Open "Select * from supplier", con3, adOpenKeyset, adLockPessimistic, adcmdtxt
'Set DataGrid1.DataSource = rs
'DataGrid1.Refresh
'Set rs = Nothing




End Sub
Private Sub LoadFG()
    con2.Open
    rs2.Open "supplier", con2, adOpenDynamic, adLockOptimistic
    fg.Cols = 4
    fg.Rows = 1
    Dim i As Integer
    Do While Not rs2.EOF
    fg.Rows = fg.Rows + 1
    fg.Row = fg.Rows - 1
            fg.Col = 0
            fg.Text = rs2(0).Value & ""
            fg.CellAlignment = flexAlignCenterBottom
            fg.Col = 1
            fg.CellAlignment = flexAlignCenterBottom
            fg.Text = rs2(1).Value & " "
            fg.Col = 2
            fg.CellAlignment = flexAlignLeftBottom
            fg.Text = rs2(2).Value & ""
            fg.Col = 3
            fg.CellAlignment = flexAlignLeftBottom
            fg.Text = rs2(3).Value & ""
            
            
            
    rs2.MoveNext
    Loop
    rs2.Close
    con2.Close
    VID = 0
End Sub
Private Sub MAINR()
    Dim items(4) As String
    Dim colndx As Integer
    items(0) = "Name"
    items(1) = "Address"
    items(2) = "Phone"
    items(3) = "Email id"
    
    

    fg.ColWidth(0) = 2500
    fg.ColWidth(1) = 2500
    fg.ColWidth(2) = 2500
    fg.ColWidth(3) = 2500
    
    
    fg.Row = 0
    For colndx = 0 To fg.Cols - 1
        fg.Col = colndx
        fg.CellFontBold = True
        fg.CellAlignment = flexAlignCenterBottom
        fg.Text = items(colndx)
    Next
    Call LoadFG
End Sub
  
