VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00000040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Records"
   ClientHeight    =   10395
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10695
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10395
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00404080&
      Height          =   2295
      Left            =   240
      TabIndex        =   22
      Top             =   7920
      Width           =   10215
      Begin VB.CommandButton Command9 
         BackColor       =   &H00404080&
         Caption         =   "LAST "
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
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00404080&
         Caption         =   "NEXT "
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
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00404080&
         Caption         =   " PREVIOUS"
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
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00404080&
         Caption         =   " FIRST"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00404080&
         Caption         =   "UPDATE"
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
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00404080&
         Caption         =   "CLEAR ALL"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00404080&
         Caption         =   "DELETE"
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
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00404080&
         Caption         =   "ADD NEW"
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
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404080&
      Height          =   7455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   10215
      Begin VB.CommandButton Command1 
         BackColor       =   &H00404080&
         Caption         =   "close"
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
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   6600
         Width           =   2295
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   20
         Top             =   6720
         Width           =   2175
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8880
         TabIndex        =   19
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         TabIndex        =   18
         Top             =   5280
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         TabIndex        =   17
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   12
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   11
         Top             =   5280
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   10
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   9
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7800
         TabIndex        =   8
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00404080&
         Caption         =   "PAY THIS MONTH"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   16
         Top             =   3720
         Width           =   4935
      End
      Begin VB.Label Label10 
         BackColor       =   &H00404080&
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   6840
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00404080&
         Caption         =   "DEDUCTIONS"
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
         Left            =   7200
         TabIndex        =   14
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H00404080&
         Caption         =   "LEAVES"
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
         Left            =   3840
         TabIndex        =   13
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404080&
         Caption         =   "OTHERS"
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
         Left            =   3840
         TabIndex        =   7
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404080&
         Caption         =   "HRA"
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
         Left            =   480
         TabIndex        =   6
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404080&
         Caption         =   "BASIC"
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
         Left            =   480
         TabIndex        =   5
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404080&
         Caption         =   "EMPLOYEE CODE"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404080&
         Caption         =   "EMPLOYEE ID"
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
         Left            =   240
         TabIndex        =   3
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404080&
         Caption         =   "EMPLOYEE NAME"
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
         Left            =   5280
         TabIndex        =   2
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404080&
         Caption         =   "PERSONAL RECORD"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   1
         Top             =   480
         Width           =   5415
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cn As ADODB.Connection 'this is the connection
Private rs As ADODB.Recordset 'this is the recordset
Dim flag As Boolean





Private Sub Command1_Click()
Unload Me
Set Form7 = Nothing

End Sub

Private Sub Command2_Click()
 
 flag = False
 check
 rs.MoveFirst
 Do Until rs.EOF = True
    
    If rs.Fields(0) = Text2 Then
       MsgBox "User ID Already Exist"
       flag = True
    Else
    rs.MoveNext
    
    End If
Loop

If flag = False Then
   rs.AddNew
   rs.Fields(0) = Text2
   rs.Fields(1) = Text6
   rs.Fields(2) = Text1
   rs.Fields(3) = Text3
   rs.Fields(4) = Text7
   rs.Fields(5) = Text5
   rs.Fields(6) = Text8
   rs.Fields(7) = Text9
   rs.Fields(8) = "aaa"
   rs.update
   
     
   
End If
   
End Sub
Public Sub check()
 
 If Text2 = "" Or Text3 = "" Or Text6 = "" Then
   MsgBox "Enter The Required Fields Properly"
   flag = True
   
 End If
 
End Sub

Private Sub Command3_Click()
rs.MoveFirst
Do Until rs.EOF = True
   If rs.Fields(0) = Text2 Then
      rs.Delete
      rs.update
      End Sub
   Else
      rs.MoveNext
   End If
Loop

End Sub

Private Sub Command4_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
End Sub

Private Sub Command5_Click()
rs.MoveFirst
Do Until rs.EOF = True
   If rs.Fields(0) = Text2 Then
      rs.Fields(1) = Text6
      rs.Fields(2) = Text1
      rs.Fields(3) = Text3
      rs.Fields(4) = Text7
      rs.Fields(5) = Text5
      rs.Fields(6) = Text8
      rs.Fields(7) = Text9
      rs.update
      End Sub
   Else
      rs.MoveNext
   End If
Loop
End Sub

Private Sub Command6_Click()
rs.MoveFirst
Text2 = rs.Fields(0)
Text6 = rs.Fields(1)
Text1 = rs.Fields(2)
Text3 = rs.Fields(3)
Text7 = rs.Fields(4)
Text5 = rs.Fields(5)
Text8 = rs.Fields(6)
Text9 = rs.Fields(7)

End Sub

Private Sub Command7_Click()
rs.MovePrevious
If rs.BOF Then
rs.MoveFirst
MsgBox "u r in first record"
End If
loadall
End Sub

Private Sub Command8_Click()
rs.MoveNext
If rs.EOF Then
rs.MoveLast
MsgBox "u r in last record"
End If
loadall
End Sub

Private Sub Command9_Click()
rs.MoveLast
loadall
If rs.EOF = False Then
MsgBox "hello"
End If

End Sub

Private Sub Form_Load()


    Set cn = New ADODB.Connection
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rest1.mdb"
      
    cn.Open
    Set rs = New ADODB.Recordset
    rs.Open "employee", cn, adOpenKeyset, adLockPessimistic, adCmdTable
    
rs.MoveFirst



End Sub

Public Sub loadall()

Text2 = rs.Fields(0)
Text6 = rs.Fields(1)
Text1 = rs.Fields(2)
Text3 = rs.Fields(3)
Text7 = rs.Fields(4)
Text5 = rs.Fields(5)
Text8 = rs.Fields(6)
Text9 = rs.Fields(7)


End Sub

