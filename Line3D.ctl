VERSION 5.00
Begin VB.UserControl Line3D 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "Line3D.ctx":0000
End
Attribute VB_Name = "Line3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/* Author: Morgan Haueisen (morganh@hartcom.net)
'/* Copyright (c) 2004
'/* Version 1.0.0

Option Explicit

Public Enum enuLineTypes
   [Flat Line] = 0
   [Inserted Line] = 1
   [Raised Line] = 2
End Enum

Private mudtLineType      As enuLineTypes
Private mlng3DHighlight   As OLE_COLOR
Private mlng3DShadow      As OLE_COLOR

Public Property Get DrawStyle() As DrawStyleConstants
   
   DrawStyle = UserControl.DrawStyle
   
End Property

Public Property Let DrawStyle(ByVal vNewValue As DrawStyleConstants)
   
   UserControl.DrawStyle = vNewValue
   PropertyChanged "DrawStyle"
   UserControl.Cls
   Call UserControl_Resize
   
End Property

Public Property Get Border3DHighlight() As OLE_COLOR
   
   Border3DHighlight = mlng3DHighlight
   
End Property

Public Property Let Border3DHighlight(ByVal vNewValue As OLE_COLOR)
   
   mlng3DHighlight = vNewValue
   PropertyChanged "Border3DHighlight"
   Call UserControl_Resize
   
End Property

Public Property Get Border3DShadow() As OLE_COLOR
   
   Border3DShadow = mlng3DShadow
   
End Property

Public Property Let Border3DShadow(ByVal vNewValue As OLE_COLOR)
   
   mlng3DShadow = vNewValue
   PropertyChanged "Border3DShadow"
   Call UserControl_Resize
   
End Property

Public Property Get LineTypes() As enuLineTypes

   LineTypes = mudtLineType

End Property

Public Property Let LineTypes(ByVal vNewValue As enuLineTypes)

   mudtLineType = vNewValue
   PropertyChanged "LineType"
   Call UserControl_Resize
   
End Property

Private Sub UserControl_InitProperties()
   
   mudtLineType = 1
   mlng3DHighlight = vb3DHighlight
   mlng3DShadow = vb3DShadow

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   On Error Resume Next
   
   With PropBag
      mudtLineType = .ReadProperty("LineType", 1)
      mlng3DHighlight = .ReadProperty("3DHighlight", vb3DHighlight)
      mlng3DShadow = .ReadProperty("3DShadow", vb3DShadow)
      UserControl.DrawStyle = .ReadProperty("DrawStyle", UserControl.DrawStyle)
   End With


End Sub

Private Sub UserControl_Resize()
   
   With UserControl
   
      If .Width >= .Height Then
         
         '/* Horizontal Line
         Select Case mudtLineType
          Case 2 'Raised
            UserControl.Line (0, 0)-(.ScaleWidth, 0), mlng3DHighlight
            UserControl.Line (0, 1)-(.ScaleWidth, 1), mlng3DShadow
            .Height = 2 * Screen.TwipsPerPixelY
          Case 1 'Inserted
            UserControl.Line (0, 0)-(.ScaleWidth, 0), mlng3DShadow
            UserControl.Line (0, 1)-(.ScaleWidth, 1), mlng3DHighlight
            .Height = 2 * Screen.TwipsPerPixelY
          Case Else ' Flat
            UserControl.Line (0, 0)-(.ScaleWidth, 0), mlng3DShadow
            .Height = Screen.TwipsPerPixelY
         End Select
         
      Else
         
         '/* Vertical Line
         Select Case mudtLineType
          Case 2 'Raised
            UserControl.Line (0, 0)-(0, .ScaleHeight), mlng3DHighlight
            UserControl.Line (1, 0)-(1, .ScaleHeight), mlng3DShadow
            .Width = 2 * Screen.TwipsPerPixelX
          Case 1 'Inserted
            UserControl.Line (0, 0)-(0, .ScaleHeight), mlng3DShadow
            UserControl.Line (1, 0)-(1, .ScaleHeight), mlng3DHighlight
            .Width = 2 * Screen.TwipsPerPixelX
          Case Else 'Flat
            UserControl.Line (0, 0)-(0, .ScaleHeight), mlng3DShadow
            .Width = Screen.TwipsPerPixelX
         End Select
      
      End If
      
      .Refresh
   End With
  
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   On Error Resume Next
   
   With PropBag
      .WriteProperty "LineType", mudtLineType
      .WriteProperty "3DHighlight", mlng3DHighlight
      .WriteProperty "3DShadow", mlng3DShadow
      .WriteProperty "DrawStyle", UserControl.DrawStyle
   End With

End Sub

