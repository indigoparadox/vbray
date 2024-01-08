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
   Begin VB.TextBox LogText 
      Height          =   2415
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "Log"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Public Sub LogTalk(Name As String, Message As String)
    LogText.Text = Name & ": " & Message & vbCrLf & LogText.Text
End Sub

Private Sub Form_Load()
    View.MenuLog.Checked = True
End Sub

Private Sub Form_Resize()
    LogText.Width = Log.Width - 115
    LogText.Height = Log.Height - 405
End Sub

Public Sub LogDebug(Message As String)
    If View.MenuDebugLog.Checked Then
        LogText.Text = Message & vbCrLf & LogText.Text
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    View.MenuLog.Checked = False
End Sub


