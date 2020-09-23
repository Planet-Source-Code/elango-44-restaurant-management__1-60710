VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00000040&
   Caption         =   "Manager's Area"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   10695
   LinkTopic       =   "Form6"
   ScaleHeight     =   8730
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00404080&
      Height          =   8175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   10215
      Begin VB.CommandButton Command6 
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
         Height          =   615
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7080
         Width           =   3495
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00404080&
         Caption         =   "CHOSEN ONES"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5520
         Width           =   3615
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00404080&
         Caption         =   "IDS AND CODES"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4560
         Width           =   3615
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00404080&
         Caption         =   "STAFF RECORDS"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3600
         Width           =   3615
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00404080&
         Caption         =   "PERSONAL RECORDS"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2640
         Width           =   3615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00404080&
         Caption         =   "COUNTER ZONE"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404080&
         Caption         =   "VIEW CHOSEN ONES"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   5640
         Width           =   5175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404080&
         Caption         =   "JOB IDS AND CODES"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   4680
         Width           =   5175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404080&
         Caption         =   "VIEW STAFF RECORDS"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   3720
         Width           =   5415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404080&
         Caption         =   "VIEW PERSONAL RECORDS"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   2760
         Width           =   5295
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404080&
         Caption         =   "GO TO COUNTER ZONE"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   1800
         Width           =   5295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404080&
         Caption         =   "MANAGERS AREA"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   1
         Top             =   480
         Width           =   5655
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cn As ADODB.Connection 'this is the connection
Private rs As ADODB.Recordset   'this is the recordset
Dim leaves As Long
Dim basic As Long
Dim hra As Long
Dim others As Long
Dim deductions As Long
Dim total As Long



Private Sub Command1_Click()
Form10.Show
End Sub

Private Sub Command2_Click()
Load Form13
Form13.Show
Form13.var = 1
End Sub

Private Sub Command3_Click()
    Set cn = New ADODB.Connection '
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rest1.mdb"
      
    cn.Open
    Set rs = New ADODB.Recordset
    rs.Open "employee", cn, adOpenDynamic, adLockOptimistic
    rs.MoveFirst
    If rs.BOF = True Then rs.MoveFirst
    
    Form7.Show
    Form7.Text2 = rs.Fields(0)
    Form7.Text1 = rs.Fields(2)
    Form7.Text6 = rs.Fields(1)
    Form7.Text3 = rs.Fields(3)
    basic = rs.Fields(3)
    Form7.Text7 = rs.Fields(4)
    leaves = rs.Fields(4)
    Form7.Text5 = rs.Fields(5)
    hra = rs.Fields(5)
    Form7.Text8 = rs.Fields(6)
    others = rs.Fields(6)
    Form7.Text9 = rs.Fields(7)
    deductions = rs.Fields(7)
    
    total = basic + hra + others - deductions - 700 * (leaves)
    Form7.Text10 = total
    Form7.Label1.Caption = "EMPLOYEE RECORDS"
    Me.Hide
    
End Sub

Private Sub Command4_Click()
Form11.Show
End Sub

Private Sub Command5_Click()
Set Form8 = Nothing
Form8.Show

End Sub

Private Sub Command6_Click()
Unload Me
Set Form6 = Nothing

End Sub



