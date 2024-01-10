VERSION 4.00
Begin VB.Form MiniMap 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map"
   ClientHeight    =   2595
   ClientLeft      =   1635
   ClientTop       =   2625
   ClientWidth     =   2880
   Height          =   3000
   Icon            =   "MiniMap.frx":0000
   Left            =   1575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   45.773
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   50.8
   Top             =   2280
   Width           =   3000
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
Option Explicit




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
Public Sub SetupLines(ViewW As Integer, TilemapW As Integer, TilemapH As Integer, UnloadLines As Boolean)
    Dim XOff As Integer
    
    If 0 < TilemapW Then
        ScaleWidth = TilemapW * 3
    End If
    If 0 < TilemapH Then
        ScaleHeight = TilemapH * 3
    End If
    
    Rem Setup the wall lines.
    For XOff = 0 To ViewW - 1
        Rem Expand control array as needed.
        If XOff > 0 Then
            If UnloadLines Then
                Unload Rays(XOff)
            Else
                Load Rays(XOff)
            End If
        End If
        If Not UnloadLines Then
            Rem Bring to front.
            Rays(XOff).ZOrder
        End If
    Next XOff
End Sub

Private Sub Form_Load()
    View.MenuMiniMap.Checked = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
    View.MenuMiniMap.Checked = False
End Sub


