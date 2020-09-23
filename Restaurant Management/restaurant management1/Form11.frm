VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form11 
   BackColor       =   &H00000040&
   Caption         =   "job ids and codes"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6630
   LinkTopic       =   "Form11"
   ScaleHeight     =   7935
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00404080&
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   6480
      Width           =   6135
      Begin VB.CommandButton Command1 
         BackColor       =   &H00404080&
         Caption         =   "E X I T "
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
         Top             =   480
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404080&
      Height          =   5655
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6135
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3735
         Left            =   600
         TabIndex        =   1
         Top             =   1320
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   6588
         _Version        =   393216
         Rows            =   15
         Cols            =   4
         BackColor       =   12640511
         BackColorFixed  =   8421504
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404080&
         Caption         =   "JOB IDS AND THEIR CODES"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   5775
      End
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cn As ADODB.Connection 'this is the connection
Private rs As ADODB.Recordset 'this is the recordset
Dim r As Integer


Private Sub Command1_Click()
Unload Me
Set Form11 = Nothing

End Sub

Private Sub Form_Load()

Dim a As Integer
'Dim arr(1 To 3) As String
'arr(1) = "S.NO"
'arr(2) = "EMP CODES"
'arr(3) = "POST"

'MSFlexGrid1.Row = 0
'For a = 1 To 3
'MSFlexGrid1.Col = a
'MSFlexGrid1.Text = arr(a)
'Next

MSFlexGrid1.FormatString = " |S.NO| CODES|        POST NAME         "

Set cn = New ADODB.Connection
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;data source=" & App.Path & "\rest1.mdb"
    cn.Open
    Set rs = New ADODB.Recordset
    rs.Open "post", cn, adOpenKeyset, adLockPessimistic, adCmdTable
    
rs.MoveFirst
If rs.BOF = True Then rs.MoveFirst

Do Until rs.EOF = True
   update
   Loop
   


End Sub

Public Sub update()
   MSFlexGrid1.Row = r + 1
   MSFlexGrid1.Col = 1
   MSFlexGrid1.Text = rs.Fields(0)
   MSFlexGrid1.Col = 2
   MSFlexGrid1.Text = rs.Fields(1)
   MSFlexGrid1.Col = 3
   MSFlexGrid1.Text = rs.Fields(2)
   r = r + 1
   rs.MoveNext
End Sub


