VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form10 
   BackColor       =   &H00000040&
   Caption         =   "counter zone"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   10830
   LinkTopic       =   "Form10"
   ScaleHeight     =   9000
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00404080&
      Height          =   1455
      Left            =   10080
      TabIndex        =   30
      Top             =   9000
      Width           =   4575
      Begin VB.CommandButton Command7 
         BackColor       =   &H00404080&
         Caption         =   "E X I T"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404080&
      Height          =   1455
      Left            =   360
      TabIndex        =   13
      Top             =   9000
      Width           =   9495
      Begin VB.TextBox Text8 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6480
         TabIndex        =   21
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00404080&
         Caption         =   "TODAY"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackColor       =   &H00404080&
         Caption         =   "VIEW ORDERS HISTORY"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackColor       =   &H00404080&
         Caption         =   "TODAY S DEALING "
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Top             =   840
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404080&
      Height          =   8055
      Left            =   360
      TabIndex        =   10
      Top             =   480
      Width           =   14295
      Begin VB.CommandButton Command5 
         BackColor       =   &H00404080&
         Caption         =   "PAID"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   7200
         Width           =   3615
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11640
         TabIndex        =   20
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00404080&
         Caption         =   "ADD "
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5880
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   19
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9480
         TabIndex        =   18
         Top             =   6000
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1935
         Left            =   3240
         TabIndex        =   17
         Top             =   2160
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   3413
         _Version        =   393216
         Rows            =   10
         Cols            =   6
         BackColor       =   12640511
         BackColorFixed  =   8421504
         BackColorSel    =   16576
         BackColorBkg    =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11640
         TabIndex        =   16
         Top             =   5040
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9480
         TabIndex        =   2
         Top             =   5040
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00404080&
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4320
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         TabIndex        =   12
         Top             =   5040
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00404080&
         Caption         =   "REMOVE"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5880
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00404080&
         Caption         =   "FINAL TOTAL"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5880
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   11
         Top             =   5040
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3240
         TabIndex        =   1
         Top             =   5040
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form10.frx":0000
         Left            =   1080
         List            =   "Form10.frx":0002
         TabIndex        =   0
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H00404080&
         Caption         =   "ADD TO RECORDS"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   31
         Top             =   7320
         Width           =   3015
      End
      Begin VB.Label Label8 
         BackColor       =   &H00404080&
         Caption         =   "QUANTITY"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9480
         TabIndex        =   29
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404080&
         Caption         =   "ITEM PRICE"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         TabIndex        =   28
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404080&
         Caption         =   "ITEM CODE"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   27
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404080&
         Caption         =   "ITEM NAME"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   26
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404080&
         Caption         =   "ITEM TYPE"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   25
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404080&
         Caption         =   "ORDER NO."
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11640
         TabIndex        =   24
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404080&
         Caption         =   "DATE"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404080&
         Caption         =   "COUNTER ZONE"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   22
         Top             =   360
         Width           =   7695
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cn As ADODB.Connection 'this is the connection
Private rs As ADODB.Recordset 'this is the recordset
Private rs1 As ADODB.Recordset
Private rs2 As ADODB.Recordset
Dim cmb1 As String
Dim qnt As Integer
Dim ttl As Integer
Dim pri As Integer
Dim ord As Integer
Dim r As Integer
Dim td As Integer
Dim ft As Integer


Private Sub Combo1_Click()
clearall
If Combo1.Text = "bevereges" Then
   clearall
   rs1.Close
   Set rs1 = New ADODB.Recordset
   rs1.Open "bevereges", cn, adOpenKeyset, adLockPessimistic, adCmdTable
   rs1.MoveFirst 'moves to the first record
    Do Until rs1.EOF = True 'this is the Loop to add items to the combo box
      Combo2.AddItem rs1.Fields("itm_name") 'this adds items from field1 into the combo box
      rs1.MoveNext 'moves next record
    Loop
   rs1.MoveFirst
End If

If Combo1.Text = "meals" Then
   clearall
   rs1.Close
   Set rs1 = New ADODB.Recordset
   rs1.Open "meals", cn, adOpenKeyset, adLockPessimistic, adCmdTable
   rs1.MoveFirst 'moves to the first record
    Do Until rs1.EOF = True 'this is the Loop to add items to the combo box
     Combo2.AddItem rs1.Fields("itm_name") 'this adds items from field1 into the combo box
     rs1.MoveNext 'moves next record
    Loop
 rs1.MoveFirst
End If


If Combo1.Text = "snacks" Then
   clearall
   rs1.Close
  Set rs1 = New ADODB.Recordset
  rs1.Open "snacks", cn, adOpenKeyset, adLockPessimistic, adCmdTable
  rs1.MoveFirst 'moves to the first record
Do Until rs1.EOF = True 'this is the Loop to add items to the combo box
Combo2.AddItem rs1.Fields("itm_name") 'this adds items from field1 into the combo box
rs1.MoveNext 'moves next record
Loop
rs1.MoveFirst
End If

If Combo1.Text = "deserts" Then
   clearall
   rs1.Close
  Set rs1 = New ADODB.Recordset
  rs1.Open "deserts", cn, adOpenKeyset, adLockPessimistic, adCmdTable
  rs1.MoveFirst 'moves to the first record
Do Until rs1.EOF = True
Combo2.AddItem rs1.Fields("itm_name")
rs1.MoveNext 'moves next record
Loop
rs1.MoveFirst
End If
End Sub

Private Sub Combo2_Click()
Text3 = ""
Text4 = ""

cmb1 = Combo2.Text
rs1.MoveFirst
Do Until rs1.EOF = True
If cmb1 = rs1!itm_name Then
   Text1 = rs1!itm_code
   Text2 = rs1!price
   Text3 = "1"
   Text3.SetFocus
   Text4 = rs1!price
   Exit Sub
Else
rs1.MoveNext
End If
Loop

End Sub

Private Sub Command2_Click()
MSFlexGrid1.Col = 1
ft = ft - Val(MSFlexGrid1.Text)
MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
r = r - 1
End Sub

Private Sub Command3_Click()
If Text3 = "" Then Text3 = "1"
qnt = Val(Text3.Text)
pri = Val(Text2.Text)
ttl = qnt * pri
Text4 = ttl
End Sub

Private Sub Command4_Click()

If Text3 = "" Then Text3 = "1"
qnt = Val(Text3.Text)
pri = Val(Text2.Text)
ttl = qnt * pri
Text4 = ttl
ft = ft + Text4
MSFlexGrid1.Row = r + 1
MSFlexGrid1.Col = 1
MSFlexGrid1.Text = Text1
MSFlexGrid1.Col = 2
MSFlexGrid1.Text = Combo2.Text
MSFlexGrid1.Col = 3
MSFlexGrid1.Text = Text2
MSFlexGrid1.Col = 4
MSFlexGrid1.Text = Text3
MSFlexGrid1.Col = 5
MSFlexGrid1.Text = Text4
r = r + 1
End Sub

Private Sub Command5_Click()
rs2.AddNew
rs2.Fields(1) = Text6
rs2.Fields(2) = Text5
rs2.update
rs2.MoveLast
Text7 = rs2.Fields(0) + 1
MsgBox "Paid"
clearall
End Sub

Private Sub Command6_Click()
rs2.MoveFirst
If rs2.BOF Then
rs2.MoveFirst
End If

td = 0

Do Until rs2.EOF = True
 If rs2.Fields(1) = Text6 Then
    td = td + rs2.Fields(2)
    rs2.MoveNext
 Else
    rs2.MoveNext
 End If
 Loop
 
 
Text8 = td
 
End Sub

Private Sub Command7_Click()
Unload Me
Set Form10 = Nothing
End Sub

Private Sub Form_Load()

Dim a As Integer
Dim arr(1 To 5) As String
arr(1) = "Item Code"
arr(2) = "Item Name"
arr(3) = "Price"
arr(4) = "Quantity"
arr(5) = "Total Price"

MSFlexGrid1.Row = 0
For a = 1 To 5
MSFlexGrid1.Col = a
MSFlexGrid1.Text = arr(a)
Next

Combo2.Clear
Text1.Text = " "
Text2.Text = " "

Text6 = Date

Me.MousePointer = 11

    Set cn = New ADODB.Connection
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rest1.mdb"
      
    cn.Open
    Set rs = New ADODB.Recordset
    rs.Open "itm_type", cn, adOpenKeyset, adLockPessimistic, adCmdTable
    Set rs1 = New ADODB.Recordset
    rs1.Open "bevereges", cn, adOpenKeyset, adLockPessimistic, adCmdTable
      Set rs2 = New ADODB.Recordset
      rs2.Open "paid", cn, adOpenDynamic, adLockOptimistic
      rs2.MoveLast
      If rs2.EOF = True Then
         rs2.MoveLast
         End If
      
      ord = rs2!ordno
      Text7 = ord + 1
Do Until rs.EOF = True 'this is the Loop to add items to the combo box
Combo1.AddItem rs.Fields("itm_type") 'this adds items from field1 into the combo box
rs.MoveNext 'moves next record
Loop
rs.MoveFirst
'fillfields 'i'll explain this later on.
Me.MousePointer = 0 'sets the mouse pointer to the normal arrow


'

End Sub

Private Sub Command1_Click()
Text5 = ft
End Sub
Public Sub clearall()
Combo2.Clear

Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "

End Sub



Private Sub Label17_Click()
End Sub

