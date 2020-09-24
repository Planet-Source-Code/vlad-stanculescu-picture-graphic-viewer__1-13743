VERSION 5.00
Begin VB.Form frmLoadPic 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Load a Picture"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6975
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.PictureBox displaypic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2590
         Left            =   4320
         ScaleHeight     =   2565
         ScaleWidth      =   2265
         TabIndex        =   4
         Top             =   600
         Width           =   2295
      End
      Begin VB.DirListBox dirbox 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   990
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   4095
      End
      Begin VB.DriveListBox drivebox 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4095
      End
      Begin VB.FileListBox filebox 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1590
         Left            =   120
         Pattern         =   "*.bmp;*.jpg;*.gif"
         TabIndex        =   1
         Top             =   1600
         Width           =   4095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         Caption         =   "Preview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmLoadPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frmPic(1 To 1000) As New frmTemplate
Private Sub dirbox_Change()
    filebox.Path = dirbox.Path
End Sub
Private Sub drivebox_Change()
    dirbox.Path = drivebox.Drive
End Sub
Private Sub filebox_Click()
    displaypic.Picture = LoadPicture(dirbox.Path & "\" & filebox.FileName)
End Sub
Private Sub filebox_dblClick()
    Dim pic As String, pic2 As String
    displaypic.Picture = LoadPicture(dirbox.Path & "\" & filebox.FileName)
    pic = dirbox.Path & "\" & filebox.FileName
    pic2 = filebox.FileName
    Unload Me
    INCNR = (INCNR + 1)
    Load frmPic(INCNR)
    frmPic(INCNR).img.Picture = LoadPicture(pic)
    frmPic(INCNR).Caption = pic2 & " (" & (frmPic(INCNR).img.Width / 15) & "," & (frmPic(INCNR).img.Height / 15) & ")"
    frmPic(INCNR).Label1.Caption = frmPic(INCNR).img.Width
    frmPic(INCNR).Label2.Caption = frmPic(INCNR).img.Height
    frmPic(INCNR).Width = frmPic(INCNR).img.Width
    frmPic(INCNR).Height = frmPic(INCNR).img.Height + 380
End Sub

