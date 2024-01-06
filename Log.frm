VERSION 4.00
Begin VB.Form Log 
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   1530
   ClientTop       =   6645
   ClientWidth     =   6135
   Height          =   4860
   Left            =   1470
   LinkTopic       =   "Form1"
   Picture         =   "Log.frx":0000
   ScaleHeight     =   4455
   ScaleWidth      =   6135
   Top             =   6300
   Width           =   6255
   Begin VB.TextBox LogText 
      Height          =   4455
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "Log"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Form_Resize()
    logtext.Width = Log.Width - 115
    logtext.Height = Log.Height - 405
End Sub

Public Sub LogLine(Message As String)
    logtext.text = logtext.text & Message & vbCrLf
End Sub
