VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   $"frmAbout.frx":0000
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   4335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Author E-Mail: muhy3@ihug.com.au"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   4335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Version: 1.0.00"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   4335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "PicView"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_click()
 Unload Me
End Sub

Private Sub Frame1_click()
 Unload Me
End Sub

Private Sub Label1_Click()
 Unload Me
End Sub

Private Sub Label2_Click()
 Unload Me
End Sub

Private Sub Label3_Click()
 Unload Me
End Sub

Private Sub Label4_Click()
 Unload Me
End Sub
