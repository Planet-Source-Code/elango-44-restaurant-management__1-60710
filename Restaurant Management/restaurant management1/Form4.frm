VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000040&
   Caption         =   "login page"
   ClientHeight    =   5370
   ClientLeft      =   3135
   ClientTop       =   2970
   ClientWidth     =   8385
   BeginProperty Font 
      Name            =   "Embossing Tape 2 (BRK)"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   5370
   ScaleWidth      =   8385
   Begin VB.Frame Frame1 
      BackColor       =   &H00404080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7935
      Begin VB.CommandButton Command2 
         BackColor       =   &H00404080&
         Caption         =   "MEMBERS LOGIN"
         Height          =   615
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3240
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00404080&
         Caption         =   "MANAGERS LOGIN"
         Height          =   615
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00404080&
         Caption         =   "STAFF MEMBERS AREA"
         Height          =   495
         Left            =   720
         TabIndex        =   4
         Top             =   3240
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00404080&
         Caption         =   "MANAGERS AREA"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404080&
         Caption         =   "LOGIN PAGE"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
FORM12.Show
End Sub

Private Sub Command2_Click()
form2.Show
End Sub

