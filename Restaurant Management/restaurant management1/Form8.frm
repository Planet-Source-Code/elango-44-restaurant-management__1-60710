VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00000040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chosen Ones"
   ClientHeight    =   8625
   ClientLeft      =   3540
   ClientTop       =   2715
   ClientWidth     =   6885
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   6885
   Begin VB.Frame Frame2 
      BackColor       =   &H00404080&
      Height          =   4455
      Left            =   240
      TabIndex        =   1
      Top             =   3960
      Width           =   6495
      Begin VB.CommandButton Command1 
         BackColor       =   &H00404080&
         Caption         =   "ADD"
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
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "add this to main record"
         Top             =   3480
         Width           =   2535
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
         Left            =   3720
         TabIndex        =   7
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox Text4 
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
         Left            =   3720
         TabIndex        =   6
         Top             =   2040
         Width           =   2535
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
         Left            =   3720
         TabIndex        =   5
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00404080&
         Caption         =   "ADD NEW RECORD"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   17
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label8 
         BackColor       =   &H00404080&
         Caption         =   "ADD THIS TO RECORD"
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
         TabIndex        =   16
         Top             =   3600
         Width           =   3495
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404080&
         Caption         =   "FOR THE MONTH"
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
         TabIndex        =   15
         Top             =   2880
         Width           =   3495
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404080&
         Caption         =   "COUNTER PERSON"
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
         TabIndex        =   14
         Top             =   2160
         Width           =   3495
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404080&
         Caption         =   "SALES PERSON"
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
         TabIndex        =   13
         Top             =   1440
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404080&
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6495
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
         Left            =   3720
         TabIndex        =   4
         Top             =   1200
         Width           =   2535
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
         Left            =   3720
         TabIndex        =   3
         Top             =   2640
         Width           =   2535
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
         Left            =   3720
         TabIndex        =   2
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404080&
         Caption         =   "FOR THE MONTH"
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
         TabIndex        =   12
         Top             =   2760
         Width           =   3495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404080&
         Caption         =   "SALES PERSON"
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
         TabIndex        =   11
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404080&
         Caption         =   "COUNTER PERSON"
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
         TabIndex        =   10
         Top             =   2040
         Width           =   3495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404080&
         Caption         =   "CHOSEN ONES THIS MONTH"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   4575
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private cn As ADODB.Connection 'this is the connection
Private rs As ADODB.Recordset 'this is the recordset

Private Sub Command1_Click()
If Text5 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox ("Fill in The Required Fields")
Exit Sub
End If

rs.AddNew
rs.Fields(1) = Text5
rs.Fields(2) = Text3
rs.Fields(3) = Text4
rs.update
clearall
loadall
End Sub
Public Sub clearall()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""

End Sub
Public Sub loadall()
    rs.MoveLast
    Text1 = rs.Fields(2)
    Text2 = rs.Fields(3)
    Text6 = rs.Fields(1)
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection 'we've declared it as a ADODB connection lets set it.
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source= j:\rest1.mdb"
    cn.Open
    Set rs = New ADODB.Recordset
    rs.Open "chosenones", cn, adOpenKeyset, adLockPessimistic, adCmdTable
    rs.MoveLast
    Text1 = rs.Fields(2)
    Text2 = rs.Fields(3)
    Text6 = rs.Fields(1)
    
End Sub



