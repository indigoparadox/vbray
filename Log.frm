VERSION 4.00
Begin VB.Form Log 
   Caption         =   "Log"
   ClientHeight    =   2430
   ClientLeft      =   1755
   ClientTop       =   6735
   ClientWidth     =   8190
   Height          =   2835
   Left            =   1695
   LinkTopic       =   "Form1"
   Picture         =   "Log.frx":0000
   ScaleHeight     =   2430
   ScaleWidth      =   8190
   Top             =   6390
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

Private Sub Form_Load()
    View.MenuLog.Checked = True
End Sub

Private Sub Form_Resize()
    logtext.Width = Log.Width - 115
    logtext.Height = Log.Height - 405
End Sub

Public Sub LogLine(Message As String)
    logtext.text = logtext.text & Message & vbCrLf
End Sub

Private Sub Form_Unload(Cancel As Integer)
    View.MenuLog.Checked = False
End Sub

