VERSION 4.00
Begin VB.Form Log 
   Caption         =   "Log"
   ClientHeight    =   2445
   ClientLeft      =   405
   ClientTop       =   7020
   ClientWidth     =   8190
   Height          =   2850
   Icon            =   "Log.frx":0000
   Left            =   345
   LinkTopic       =   "Form1"
   ScaleHeight     =   163
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   546
   Top             =   6675
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
      Begin VB.PictureBox InlinePicture 
         AutoSize        =   -1  'True
         Height          =   615
         Left            =   2520
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
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
Option Explicit

Dim LogHistoryText() As String
Dim LogHistoryCount As Integer

Public Function LogLineHeight(ByVal Line As String) As Integer
    Dim LineBreakCount As Integer
    Dim LinePxSz As Integer
    Dim LinePrefix As String
    Dim PicturesWidth As Integer
    
    PicturesWidth = 0
    
    Rem Trim out all tags in the line, one by one.
    While InStr(Line, Chr$(60))
        LinePrefix = Left(Line, InStr(Line, Chr$(60)) - 1)
        Line = Right(Line, Len(Line) - InStr(Line, Chr$(60)))
        Tag = Left(Line, InStr(Line, Chr$(62)) - 1)
        Line = LinePrefix & Right(Line, Len(Line) - InStr(Line, Chr$(62)))
    
        Rem Find any images loaded and add their width.
        If "img" = Left(Tag, 3) Then
            InlinePicture.Picture = LoadPicture(Right(Tag, Len(Tag) - InStr(Tag, Chr$(58))))
            PicturesWidth = PicturesWidth + InlinePicture.ScaleWidth
        End If
    Wend

    Rem Avoid weird VB4 banker's rounding issues with division.
    LinePxSz = LogText.TextWidth(Line) + PicturesWidth
    LineBreakCount = 0
    While LinePxSz > 0:
        LineBreakCount = LineBreakCount + 1
        LinePxSz = LinePxSz - LogText.ScaleWidth
    Wend
    
    LogLineHeight = LineBreakCount * LogText.TextHeight(Line)
End Function

Public Sub LogScroll(Pixels As Integer)
    Rem Grab the currently printed text image and shift it down.
    LogTempHolder.Picture = LogText.Image
    LogText.Cls
    LogText.PaintPicture LogTempHolder.Picture, 0, Pixels
End Sub

Public Sub LogPicture(PicturePath As String)
    Dim LineCurrentX As Integer
        
    LineCurrentX = LogText.CurrentX
    
    InlinePicture.Picture = LoadPicture(PicturePath)
    LogText.PaintPicture InlinePicture.Picture, LineCurrentX, LogText.CurrentY
        
    LogText.CurrentX = LineCurrentX + InlinePicture.ScaleWidth
    
End Sub

Public Sub LogParseTags(Word As String)
    Rem Check if the word starts with "<"
    While 1 = InStr(Word, Chr$(60))
        Rem Split off any attached words from the tag ending with ">"
        Tag = Left(Word, InStr(Word, Chr$(62)))
        Word = Right(Word, Len(Word) - InStr(Word, Chr$(62)))
        
        Select Case Tag
        Case "<green>"
            LogText.ForeColor = &H8000&
        Case "<yellow>"
            LogText.ForeColor = &H8080&
        Case "<red>"
            LogText.ForeColor = &H80&
        Case "<blue>"
            LogText.ForeColor = &H800000
        Case "<violet>"
            LogText.ForeColor = &H800080
        Case "<black>"
            LogText.ForeColor = vbBlack
        Case "<b>"
            LogText.Font.Bold = True
        Case "</b>"
            LogText.Font.Bold = False
        Case "<u>"
            LogText.Font.Underline = True
        Case "</u>"
            LogText.Font.Underline = False
        Case "<i>"
            LogText.Font.Italic = True
        Case "</i>"
            LogText.Font.Italic = False
        Case "<s>"
            LogText.Font.Strikethrough = True
        Case "</s>"
            LogText.Font.Strikethrough = False
        Case Else
            If "<img" = Left(Tag, 4) Then
                LogPicture Left(Right(Tag, Len(Tag) - 5), Len(Tag) - 6)
            End If
        End Select
    Wend
End Sub

Public Sub LogWord(ByVal Word As String)
    Dim LineCurrentX As Integer
    Dim LineCurrentY As Integer
    Dim LineRemainderTag As String
    
    LineRemainderTag = ""
    
    Rem This will trim any prefix tags off, as well as applying their formatting.
    LogParseTags Word
    
    Rem If the word ends in a tag, save that for later.
    If InStr(Word, Chr$(60)) Then
        LineRemainderTag = Right(Word, Len(Word) - InStr(Word, Chr$(60)) + 1)
        Word = Left(Word, InStr(Word, Chr$(60)) - 1)
    End If
    
    Rem Now the word should be itself, sans tags.
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
    
    If 0 < Len(LineRemainderTag) Then
        LogWord LineRemainderTag
    End If
End Sub

Public Sub LogLine(ByVal Line As String, ByVal AddHistory As Boolean)
    Dim LineFirstWord As String
    
    Rem Add logged line to in-memory text log.
    If AddHistory Then
        ReDim Preserve LogHistoryText(LogHistoryCount)
        LogHistoryText(LogHistoryCount) = Line
        LogHistoryCount = LogHistoryCount + 1
    End If
    
    LogText.ForeColor = vbBlack
    LogText.Font.Bold = False
    
    LogScroll LogLineHeight(Line)
    
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

Public Sub LogTalk(Name As String, Message As String)
    Rem Print the character talking's name in blue.
    Rem Print what they're saying in black, allowing for line breaks.
    LogLine "<blue>" & Name & ":</blue> " & Message, True
End Sub

Private Sub Form_Load()
    View.MenuLog.Checked = True
    LogText.Picture = LogText.Image
    LogLine "(Beginning)", True
End Sub

Public Sub LogDebug(Message As String)
    If View.MenuDebugLog.Checked Then
        LogScroll LogLineHeight(Message)
        LogLine Message, True
    End If
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    
    Rem Resize, clear, and redraw the log.
    LogText.Width = ScaleWidth - 1
    LogText.Height = ScaleHeight - 1
    LogText.Cls
    
    For i = 0 To LogHistoryCount - 1
        LogLine LogHistoryText(i), False
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    View.MenuLog.Checked = False
End Sub


