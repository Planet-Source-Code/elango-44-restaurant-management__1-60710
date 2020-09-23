VERSION 5.00
Begin VB.Form form1 
   BackColor       =   &H00000040&
   Caption         =   "welcome"
   ClientHeight    =   4365
   ClientLeft      =   4125
   ClientTop       =   3855
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Electrofied"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   7200
   WindowState     =   2  'Maximized
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
      Height          =   9735
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   14655
      Begin VB.CommandButton Command1 
         BackColor       =   &H00404080&
         Caption         =   "click to go"
         DisabledPicture =   "Form1.frx":0000
         DownPicture     =   "Form1.frx":163C
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9480
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   8280
         Width           =   4455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404080&
         Caption         =   "RESTAURANT MANAGEMENT SYSTEM"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   600
         TabIndex        =   2
         Top             =   2520
         Width           =   13695
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Form4.Show

End Sub

