VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   Caption         =   "form4"
   ClientHeight    =   9810
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15300
   LinkTopic       =   "Form9"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   9810
   ScaleWidth      =   15300
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Back"
      Height          =   495
      Left            =   13560
      TabIndex        =   33
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "view"
      Height          =   495
      Left            =   9480
      TabIndex        =   32
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000014&
      Height          =   6015
      Left            =   8160
      TabIndex        =   30
      Top             =   1440
      Width           =   9135
      Begin MSFlexGridLib.MSFlexGrid fg 
         Height          =   2655
         Left            =   240
         TabIndex        =   31
         Top             =   720
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4683
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColorBkg    =   -2147483633
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Caption         =   "Details"
      Height          =   9375
      Left            =   720
      TabIndex        =   2
      Top             =   1440
      Width           =   6855
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   1320
         TabIndex        =   35
         Top             =   7440
         Width           =   4695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   615
         Left            =   3000
         TabIndex        =   28
         Top             =   8280
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Update"
         Height          =   615
         Left            =   480
         TabIndex        =   27
         Top             =   8280
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   1320
         TabIndex        =   26
         Top             =   6000
         Width           =   4695
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   1320
         TabIndex        =   24
         Top             =   4080
         Width           =   4335
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000014&
         Caption         =   "Topping Flavour"
         Height          =   975
         Left            =   240
         TabIndex        =   18
         Top             =   4800
         Width           =   6615
         Begin VB.OptionButton Option7 
            BackColor       =   &H80000014&
            Caption         =   "Chocolate"
            Height          =   495
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option8 
            BackColor       =   &H80000014&
            Caption         =   "Vanilla"
            Height          =   495
            Left            =   1680
            TabIndex        =   21
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option9 
            BackColor       =   &H80000014&
            Caption         =   "Strawberry"
            Height          =   495
            Left            =   3120
            TabIndex        =   20
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option10 
            BackColor       =   &H80000014&
            Caption         =   "Mango"
            Height          =   495
            Left            =   4920
            TabIndex        =   19
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1560
         TabIndex        =   17
         Top             =   480
         Width           =   4455
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
         Left            =   1560
         TabIndex        =   13
         Top             =   2880
         Width           =   4695
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000014&
         Caption         =   "CAKES"
         Height          =   1575
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   6615
         Begin VB.OptionButton Option1 
            BackColor       =   &H80000014&
            Caption         =   "Blackforest"
            Height          =   495
            Left            =   360
            TabIndex        =   11
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H80000014&
            Caption         =   "Pineapple"
            Height          =   495
            Left            =   360
            TabIndex        =   10
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H80000014&
            Caption         =   "Red Velvet"
            Height          =   495
            Left            =   2040
            TabIndex        =   9
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H80000014&
            Caption         =   "MudPie"
            Height          =   495
            Left            =   2040
            TabIndex        =   8
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H80000014&
            Caption         =   "MangoTwin"
            Height          =   495
            Left            =   3720
            TabIndex        =   7
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H80000014&
            Caption         =   "ChocoPie"
            Height          =   495
            Left            =   3720
            TabIndex        =   6
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   1320
         TabIndex        =   4
         Top             =   6720
         Width           =   4695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   1560
         TabIndex        =   15
         Top             =   3480
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   873
         _Version        =   393216
         Format          =   126877697
         CurrentDate     =   43341
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000E&
         Caption         =   "Cost Price"
         Height          =   495
         Left            =   360
         TabIndex        =   34
         Top             =   7560
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000014&
         Caption         =   "stock"
         Height          =   495
         Left            =   360
         TabIndex        =   25
         Top             =   6840
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000014&
         Caption         =   "Description"
         Height          =   495
         Left            =   360
         TabIndex        =   23
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000014&
         Caption         =   "Fllavour"
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000014&
         Caption         =   "Order_date"
         Height          =   495
         Left            =   360
         TabIndex        =   14
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000014&
         Caption         =   "Supplier_name"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000014&
         Caption         =   "Cake Name"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1320
      Top             =   11160
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   9960
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "       Add Stock"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   1080
      TabIndex        =   29
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "   Order id"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   7800
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim con1 As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ros As ADODB.Recordset
Dim max As Integer
Dim diff As Integer
Dim rs2 As ADODB.Recordset
Dim con2 As ADODB.Connection
Dim VID As Integer

Private Sub Command1_Click()


Dim s As String
s = "Update newcake set stock = stock +" & Val(Text5) & " where cname = '" & Text2.Text & "'"
con.Open
con.Execute s
con.Close





Dim d As Date
d = DTPicker1.Value
s = CStr(d)
con.Open
    rs.Open "order11", con, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs.Fields(1) = DTPicker1.Value
    rs.Fields(2) = Combo1.Text
    rs.Fields(3) = Text3.Text
    rs.Fields(4) = Text4.Text
    rs.Fields(5) = Text5.Text
    rs.Fields(6) = Text2.Text
    rs.Fields(7) = Text6.Text

    rs.Update
    MsgBox ("order sent")
End Sub

Private Sub Command2_Click()
Unload Me

End Sub


Private Sub Command3_Click()
If (VID = 0) Then
     MsgBox "Select a order to View..!", vbInformation, "Select order"
Else
     VID = CInt(fg.Text)
     Dim RSS As String
     RSS = "select * from order11 where oid = " & VID & " "
     con.Open
     Set rs = New ADODB.Recordset
     rs.Open RSS, con, adOpenDynamic, adLockOptimistic
     If (rs.BOF Or rs.EOF) Then
        MsgBox "No Record Found"
        rs.Close
        con.Close
     Else
     rs.MoveFirst
        Text1.Text = rs.Fields(0)
        Text5.Text = rs.Fields(5)
        Combo1.Text = rs.Fields(2)
        Text4.Text = rs.Fields(4)
        DTPicker1.Value = rs.Fields(1)
        Text2.Text = rs.Fields(6)
        Text3.Text = rs.Fields(3)
        Frame1.Visible = False
        Frame2.Visible = False
        
     rs.Close
     con.Close
     End If
End If
End Sub



Private Sub Command4_Click()
Form3.Show

End Sub

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
Adodc1.Refresh
With Adodc1.Recordset
Do Until .EOF
Combo1.AddItem ![sname]
.MoveNext
Loop
MAINR
getMax
Text1.Text = max

End With
End Sub

Private Sub getMax()
con.Open
rs.Open "Select max(oid) from order11", con, adOpenDynamic
If rs.EOF <> True And rs.BOF <> True Then
max = CInt(rs.Fields(0)) + 1
Else
max = 1
End If
rs.Close
con.Close
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then Text2.Text = "BlackForest"
If Option1.Value = True Then Text6.Text = "300"
End Sub

Private Sub Option10_Click()
If Option10.Value = True Then Text3.Text = "Mango"
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then Text2.Text = "PineApple"
If Option2.Value = True Then Text6.Text = "275"
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then Text2.Text = "Red Velvet"
If Option3.Value = True Then Text6.Text = "350"
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then Text2.Text = "MudPie"
If Option4.Value = True Then Text6.Text = "375"
End Sub

Private Sub Option5_Click()
If Option5.Value = True Then Text2.Text = "MangoTwin"
If Option5.Value = True Then Text6.Text = "365"
End Sub

Private Sub Option6_Click()
If Option6.Value = True Then Text2.Text = "ChocoPie"
If Option6.Value = True Then Text6.Text = "385"
End Sub

Private Sub Option7_Click()
If Option7.Value = True Then Text3.Text = "Chocolate"
End Sub

Private Sub Option8_Click()
If Option8.Value = True Then Text3.Text = "Vanilla"
End Sub

Private Sub Option9_Click()
If Option9.Value = True Then Text3.Text = "Strawberry"
End Sub
Private Sub LoadFG()
    con2.Open
    rs2.Open "order11", con2, adOpenDynamic, adLockOptimistic
    fg.Cols = 9
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
            fg.Col = 4
            fg.CellAlignment = flexAlignCenterBottom
            fg.Text = rs2(4).Value & " "
            fg.Col = 5
            fg.CellAlignment = flexAlignCenterBottom
            fg.Text = rs2(5).Value & " "
            fg.Col = 6
            fg.CellAlignment = flexAlignCenterBottom
            fg.Text = rs2(6).Value & " "
             fg.Col = 7
            fg.CellAlignment = flexAlignCenterBottom
            fg.Text = rs2(7).Value & " "
            
            
    rs2.MoveNext
    Loop
    rs2.Close
    con2.Close
    VID = 0
End Sub
Private Sub MAINR()
    Dim items(7) As String
    Dim colndx As Integer
    items(0) = "Oid"
    items(1) = "Odate"
    items(2) = "Sname"
    items(3) = "Flavour"
    items(4) = "weight"
    items(5) = "Stock"
    items(6) = "Cake name"
     items(7) = "Cost Price"
    

    fg.ColWidth(0) = 1000
    fg.ColWidth(1) = 1000
    fg.ColWidth(2) = 1000
    fg.ColWidth(3) = 1000
    fg.ColWidth(4) = 1000
    fg.ColWidth(5) = 1000
    fg.ColWidth(6) = 1000
    
    fg.Row = 0
    For colndx = 0 To fg.Cols - 1
        fg.Col = colndx
        fg.CellFontBold = True
        fg.CellAlignment = flexAlignCenterBottom
        fg.Text = items(colndx)
    Next
    Call LoadFG
End Sub
'Private Sub getMaxId()
'con.Open
'rs.Open "select max(CAMP_NO) from CAMP ", con, adOpenDynamic
'If rs.EOF <> True And rs.BOF <> True Then
   ' max = CInt(rs.Fields(0)) + 1
'Else
    'max = 1
'End If
'rs.close
'con.close

'End Sub




Private Sub Text1_Change()

End Sub
