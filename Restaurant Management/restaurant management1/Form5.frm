VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00000040&
   Caption         =   "Staff Member's Area"
   ClientHeight    =   7890
   ClientLeft      =   2385
   ClientTop       =   2100
   ClientWidth     =   10020
   LinkTopic       =   "Form5"
   ScaleHeight     =   7890
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00404080&
      Height          =   7215
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   9495
      Begin VB.CommandButton Command4 
         BackColor       =   &H00404080&
         Caption         =   "BACK"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6120
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00404080&
         Caption         =   "MONTH  FAVOURITES"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00404080&
         Caption         =   "RECORDS"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2760
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00404080&
         Caption         =   "COUNTER"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404080&
         Caption         =   "THE CHOSEN ONES"
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
         Left            =   480
         TabIndex        =   6
         Top             =   3960
         Width           =   3495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404080&
         Caption         =   "VIEW PERSONNEL RECORDS"
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
         Left            =   480
         TabIndex        =   4
         Top             =   3000
         Width           =   3855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404080&
         Caption         =   "GO TO THE COUNTER ZONE"
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
         Left            =   480
         TabIndex        =   2
         Top             =   2040
         Width           =   3975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404080&
         Caption         =   "STAFF MEMBERS AREA"
         BeginProperty Font 
            Name            =   "Embossing Tape 2 (BRK)"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   5535
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Form10.Show
End Sub

Private Sub Command2_Click()
Load Form13
Form13.Show
Form13.var = 2
End Sub

Private Sub Command3_Click()
Form8.Height = 4470
Form8.ScaleHeight = 3960
Form8.Show

Form5.Hide
    
    
End Sub

Private Sub Command4_Click()
Unload Me
Set Form5 = Nothing
End Sub

