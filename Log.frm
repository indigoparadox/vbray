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
   ScaleHeight     =   162
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   546
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
         ScaleHeight     =   61
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   109
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

Public Function LogWord(Word As String)
    Dim LineCurrentX As Integer
    Dim LineCurrentY As Integer
    
    Rem Check and wrap the line if the next word will be off-screeen.
    Rem If this word will be off-screen, then scroll and print it *below* the current line.
    If LogText.CurrentX + LogText.TextWidth(Word) > LogText.ScaleWidth Then
        LogText.CurrentX = 0
        LogText.CurrentY = LogText.CurrentY + LogText.TextHeight(Word)
    End If
    
    Rem Print the word and reset the current position.
    LineCurrentX = LogText.CurrentX
    LineCurrentY = LogText.CurrentY
    
    LogText.Print Word
    
    Rem Reset current X to what it was before printing + size of word.
    LogText.CurrentX = LineCurrentX + LogText.TextWidth(Word)
    LogText.CurrentY = LineCurrentY
End Function

Public Sub LogLine(ByVal Line As String)
    Dim LineFirstWord As String
    
    Rem Print each word in the line.
    While InStr(Line, Chr$(32))
        LineFirstWord = Left(Line, InStr(Line, Chr$(32)))
        Line = Right(Line, Len(Line) - InStr(Line, Chr$(32)))
        LogWord LineFirstWord
    Wend
    LogWord Line
    
    Rem Reset X/Y after line is finished printing.
    LogText.CurrentX = 0
    LogText.CurrentY = 0
End Sub

Public Function LogLineHeight(ByVal Line As String) As Integer
    Dim LineBreakCount As Integer
    Dim LinePxSz As Integer
    
    Rem Avoid weird VB4 banker's rounding issues with division.
    LinePxSz = LogText.TextWidth(Line)
    LineBreakCount = 0
    While LinePxSz > 0:
        LineBreakCount = LineBreakCount + 1
        LinePxSz = LinePxSz - LogText.ScaleWidth
    Wend
    
    LogLineHeight = LineBreakCount * LogText.TextHeight(Line)
End Function

Public Sub LogTalk(Name As String, Message As String)
    LogScroll LogLineHeight(Name & ": " & Message)
    
    Rem Print the character talking's name in blue.
    LogText.ForeColor = vbBlue
    LogWord Name & ": "
    
    Rem Print what they're saying in black, allowing for line breaks.
    LogText.ForeColor = vbBlack
    LogLine Message
End Sub

Private Sub Form_Load()
    View.MenuLog.Checked = True
    LogText.Picture = LogText.Image
End Sub

Public Sub LogDebug(Message As String)
    If View.MenuDebugLog.Checked Then
        LogScroll LogLineHeight(Message)
        LogText.ForeColor = vbRed
        LogLine Message
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    View.MenuLog.Checked = False
End Sub


