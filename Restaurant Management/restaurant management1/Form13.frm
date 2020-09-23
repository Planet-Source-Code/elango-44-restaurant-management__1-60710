VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H00000040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "record's login"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6495
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00404080&
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00404080&
         Caption         =   "CANCEL"
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
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00404080&
         Caption         =   "O.K."
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
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00C0E0FF&
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
         IMEMode         =   3  'DISABLE
         Left            =   3840
         PasswordChar    =   "#"
         TabIndex        =   4
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H00C0E0FF&
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
         Left            =   3840
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404080&
         Caption         =   "EMPLOYEE PASSWORD"
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
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404080&
         Caption         =   "ENTER EMPLOYEE ID"
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
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   3495
      End
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cn As ADODB.Connection 'this is the connection
Private rs As ADODB.Recordset   'this is the recordset
Public empid As String
Dim leaves As Long
Dim basic As Long
Dim hra As Long
Dim others As Long
Dim deductions As Long
Dim total As Long
Public var As Integer


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
    rs.Open "employee", cn, adOpenDynamic, adLockOptimistic
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
  If txtUserName = rs.Fields(0) And txtPassword = rs.Fields(8) Then
    'check for correct password
    'place code to here to pass the
    'success to the calling sub
    'setting a global var is the easiest
    LoginSucceeded = True
    empid = rs.Fields(0)
    Form7.Show
    Form7.Text2 = empid
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
    
    
    If var = 1 Then
    Me.Hide
    Unload Me
    
    End If
    
    If var = 2 Then
    Form7.BorderStyle = 3
    Form7.Height = 8370
    Form7.ScaleHeight = 7860
    
    Me.Hide
    Unload Me
    End If
       
    
End If
  
     
 End Sub

