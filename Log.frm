VERSION 4.00
Begin VB.Form Log 
   Caption         =   "Log"
   ClientHeight    =   2430
   ClientLeft      =   2850
   ClientTop       =   6585
   ClientWidth     =   8190
   Height          =   2835
   Icon            =   "Log.frx":0000
   Left            =   2790
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   8190
   Top             =   6240
   Width           =   8310
   Begin VB.PictureBox LogText 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   541
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.PictureBox LogTempHolder 
         Height          =   975
         Left            =   0
         ScaleHeight     =   915
         ScaleWidth      =   1635
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Log"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Public Sub LogScroll(Pixels As Integer)
    Rem Grab the currently printed text image and shift it down.
    LogTempHolder.Picture = LogText.Image
    LogText.Cls
    LogText.PaintPicture LogTempHolder.Picture, 0, Pixels
End Sub

Public Sub LogTalk(Name As String, Message As String)
    Rem TODO: Move scrolling to printline so we can scroll if it's multiple lines.
    LogScroll LogText.TextHeight(Name & ": ")
    
    Rem Print the character talking's name in blue.
    LogText.ForeColor = vbBlue
    LogText.CurrentY = 0
    LogText.Print Name & ": "
    
    LogText.ForeColor = vbBlack
    LogText.CurrentY = 0
    LogText.CurrentX = LogText.TextWidth(Name & ": ")
    LogText.Print Message
End Sub

Private Sub Form_Load()
    View.MenuLog.Checked = True
    LogText.Picture = LogText.Image
End Sub

Public Sub LogDebug(Message As String)
    If View.MenuDebugLog.Checked Then
        LogScroll LogText.TextHeight(Message)
        LogText.ForeColor = vbRed
        LogText.CurrentY = 0
        LogText.Print Message
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    View.MenuLog.Checked = False
End Sub


