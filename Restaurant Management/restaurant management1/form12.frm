VERSION 5.00
Begin VB.Form form12 
   BackColor       =   &H00404080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manager's Login"
   ClientHeight    =   2730
   ClientLeft      =   2835
   ClientTop       =   3570
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1612.974
   ScaleMode       =   0  'User
   ScaleWidth      =   4760.456
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Embossing Tape 2 (BRK)"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00404080&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Embossing Tape 2 (BRK)"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00404080&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Embossing Tape 2 (BRK)"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1260
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Embossing Tape 2 (BRK)"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00404080&
      Caption         =   "&User Name"
      BeginProperty Font 
         Name            =   "Embossing Tape 2 (BRK)"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1560
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00404080&
      Caption         =   "&Password"
      BeginProperty Font 
         Name            =   "Embossing Tape 2 (BRK)"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1545
   End
End
Attribute VB_Name = "FORM12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cn As ADODB.Connection 'this is the connection
Private rs As ADODB.Recordset 'this is the recordset

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Set cn = New ADODB.Connection '
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rest1.mdb"
      
    cn.Open
    Set rs = New ADODB.Recordset
    rs.Open "pwd", cn, adOpenDynamic, adLockOptimistic
    rs.MoveFirst
    If rs.BOF = True Then rs.MoveFirst
    
    Do Until rs.EOF = True
       check
       rs.MoveNext
    Loop
    
    If LoginSucceeded = False Then
        MsgBox "Invalid Password, or userid try again!", , "Login"
        txtUserName.SetFocus
        txtPassword = ""
        SendKeys "{Home}+{End}"
     End If
     
    
    
End Sub
 Public Sub check()
  If txtUserName = rs.Fields(0) And txtPassword = rs.Fields(1) Then
    'check for correct password
    'place code to here to pass the
    'success to the calling sub
    'setting a global var is the easiest
    LoginSucceeded = True
    
    Form6.Show
    Unload Me
    Set FORM12 = Nothing
    
        
End If
  
     
 End Sub

