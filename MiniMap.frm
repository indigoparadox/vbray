VERSION 4.00
Begin VB.Form MiniMap 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map"
   ClientHeight    =   1770
   ClientLeft      =   1860
   ClientTop       =   3660
   ClientWidth     =   1710
   Height          =   2175
   Icon            =   "MiniMap.frx":0000
   Left            =   1800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   31.221
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   30.163
   Top             =   3315
   Width           =   1830
   Begin VB.Line Rays 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   6.35
      X2              =   23.283
      Y1              =   4.233
      Y2              =   4.233
   End
End
Attribute VB_Name = "MiniMap"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Public Sub RayStart(Idx As Integer, ByVal X As Single, ByVal Y As Single)
    Rays(Idx).X1 = X * 3
    Rays(Idx).Y1 = Y * 3
End Sub

Public Sub RayEnd(Idx As Integer, ByVal X As Single, ByVal Y As Single, ByVal Color As Long)
    Rays(Idx).X2 = X * 3
    Rays(Idx).Y2 = Y * 3
    Rays(Idx).BorderColor = Color
    Rays(Idx).Visible = True
End Sub
Private Sub Form_Load()
    Dim XOff As Integer
    
    Rem TODO: Get rid of this!
    Const ViewW = 320
    
    Rem Setup the wall lines.
    For XOff = 0 To ViewW - 1
        Rem Expand control array as needed.
        If XOff > 0 Then
            Load Rays(XOff)
        End If
        Rem Bring to front.
        Rays(XOff).ZOrder
    Next XOff
End Sub


