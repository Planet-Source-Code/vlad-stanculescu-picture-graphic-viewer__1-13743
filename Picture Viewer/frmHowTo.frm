VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHowTo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "How to use this program"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   8910
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   3840
      Width           =   1815
   End
   Begin RichTextLib.RichTextBox hlpbox 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   5106
      _Version        =   393217
      BackColor       =   -2147483648
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmHowTo.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Caption         =   "PicView Help Centre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmHowTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    hlpbox.Text = hlpbox.Text & "Follow these few quick steps:"
    hlpbox.Text = hlpbox.Text & vbCrLf & " " & vbCrLf & "" & vbCrLf
    hlpbox.Text = hlpbox.Text & "1) Go to File" & vbCrLf
    hlpbox.Text = hlpbox.Text & "2) Select the option which sais 'Open'" & vbCrLf
    hlpbox.Text = hlpbox.Text & "3) Select the location of your file (drive)" & vbCrLf
    hlpbox.Text = hlpbox.Text & "4) Select the directory of your file" & vbCrLf
    hlpbox.Text = hlpbox.Text & "5) When you have located your image, double click on it or single click on it to view a preview of the image to check wether it is the one you are looking for. Otherwise select another image and double click on it." & vbCrLf
    hlpbox.Text = hlpbox.Text & "6) Once the file has opened there aren't many things you can do with this program but view the picture. You are welcome to re-distribute this program with any changes to it." & vbCrLf
    hlpbox.Text = hlpbox.Text & vbCrLf & " " & vbCrLf & "" & vbCrLf
    hlpbox.Text = hlpbox.Text & "If this has not helped you enough.. experiment, this program cannot cause any pain on your system what so ever."
End Sub
