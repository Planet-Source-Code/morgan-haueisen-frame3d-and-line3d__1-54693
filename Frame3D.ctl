VERSION 5.00
Begin VB.UserControl Frame3D 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   FillColor       =   &H00FFC19F&
   ForeColor       =   &H80000010&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "Frame3D.ctx":0000
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   195
      TabIndex        =   0
      Top             =   0
      Width           =   45
   End
   Begin VB.Label lblCaptionShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000014&
      Height          =   195
      Left            =   210
      TabIndex        =   1
      Top             =   15
      Width           =   45
   End
   Begin VB.Label lblCaptionBlank 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   45
   End
End
Attribute VB_Name = "Frame3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/*************************************************************************************************
'/* Author: Morgan Haueisen (morganh@hartcom.net)
'/* Copyright (c) 2004
'/* Version 1.0.0
'/*
'/* Legal:
'/*
'/*      Redistribution of this code, whole or in part, as source code or in binary form, alone or
'/*      as part of a larger distribution or product, is forbidden for any commercial or for-profit
'/*      use without the author's explicit written permission.
'/*
'/*      Redistribution of this code, as source code or in binary form, with or without
'/*      modification, is permitted provided that the following conditions are met:
'/*
'/*      Redistributions of source code must include this list of conditions, and the following
'/*      acknowledgment:
'/*
'/*      This code was developed by Morgan Haueisen.  <morganh@hartcom.net>
'/*      Source code, written in Visual Basic, is freely available for non-commercial,
'/*      non-profit use at www.planetsourcecode.com.
'/*
'/*      Redistributions in binary form, as part of a larger project, must include the above
'/*      acknowledgment in the end-user documentation.  Alternatively, the above acknowledgment
'/*      may appear in the software itself, if and wherever such third-party acknowledgments
'/*      normally appear.
'/*************************************************************************************************

Option Explicit

'/* Used to draw the object's rounded border
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, _
   ByVal left As Long, ByVal top As Long, ByVal right As Long, _
   ByVal bottom As Long, ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long

'/* Used to make the rounded corners transparent
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, _
   ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal RectX1 As Long, _
   ByVal RectY1 As Long, ByVal RectX2 As Long, ByVal RectY2 As Long, _
   ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long

'/* The GetSysColor function retrieves the current color of the specified display element
'/* Used to add gradient fill
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Enum enuBorderTypes
   [None Border] = 0
   [Frame Inserted] = 1
   [Frame Raised] = 2
   [Panel Flat Shadow] = 3
   [Panel Flat Highlight] = 4
   [Panel Raised] = 5
   [Panel Inserted] = 6
   [rFrame Inserted] = 7
   [rFrame Raised] = 8
   [rPanel Flat Shadow] = 9
   [rPanel Flat Highlight] = 10
   [rPanel Raised] = 11
   [rPanel Inserted] = 12
   [rNone Border] = 13
End Enum

Public Enum enuBevelInner
   [None Bevel] = 0
   [Inserted Bevel] = 1
   [Raised Bevel] = 2
   [Flat Shadow] = 3
   [Flat Highlight] = 4
End Enum

Public Enum enuCaption3D
   [Flat Caption] = 0
   [Inserted Caption] = 1
   [Raised Caption] = 2
End Enum

Public Enum enuCaptionLocation
   [Inside Frame] = 0
   [In Frame] = 1
End Enum

Public Enum enuFloodType
   [Left To Right] = 0
   [Bottom To Top] = 1
End Enum

Public Enum enuCaptionAlignment
   [Top Left] = 0
   [Top Center] = 1
   [Top Right] = 2
   [Middle Left] = 3
   [Middle Center] = 4
   [Middle Right] = 5
   [Bottom Left] = 6
   [Bottom Center] = 7
   [Bottom Right] = 8
End Enum

Public Enum enuFillGradient
   [None Gradient] = 0
   [Fill Horizontal] = 1
   [Fill Vertical] = 2
End Enum

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Private mudtBorderType           As enuBorderTypes
Private mudtBevelInner           As enuBevelInner
Private mudtCaption3D            As enuCaption3D
Private mudtCaptionAlignment     As enuCaptionAlignment
Private mudtCaptionLocation      As enuCaptionLocation
Private mudtFloodType            As enuFloodType
Private mudtFillGradient         As enuFillGradient
Private mudtCaptionMAlignment    As AlignmentConstants

Private mlngBevelWidth           As Long
Private mlng3DHighlight          As OLE_COLOR
Private mlng3DShadow             As OLE_COLOR
Private mlngBackColor            As OLE_COLOR
Private mblnEnabled              As Boolean
Private msngTop                  As Single
Private msngLeft                 As Single
Private mlngFloodValue           As Long
Private mblnFloodShowPct         As Boolean
Private mlngFloodColor           As OLE_COLOR
Private msngInsideBorder         As Single
Private mlngCornerRadius         As Long
Private mlngCornerDia            As Long
Private mlngInsideHeight         As Long
Private mlngInsideWidth          As Long
Private mlngInsideLeft           As Long
Private mlngInsideTop            As Long

Public Property Get BackColor() As OLE_COLOR
   
   BackColor = mlngBackColor
   
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
   
   mlngBackColor = vNewValue
   PropertyChanged "BackColor"
   Call UserControl_Resize
   
End Property

Public Property Get BevelInner() As enuBevelInner
   
   BevelInner = mudtBevelInner
   
End Property

Public Property Let BevelInner(ByVal vNewValue As enuBevelInner)
   
   mudtBevelInner = vNewValue
   PropertyChanged "BevelInner"
   Call UserControl_Resize
   
End Property

Public Property Get BevelWidth() As Long
   
   BevelWidth = mlngBevelWidth
   
End Property

Public Property Let BevelWidth(ByVal vNewValue As Long)
   
   mlngBevelWidth = vNewValue
   If mlngBevelWidth < 3 Then mlngBevelWidth = 3
   PropertyChanged "BevelWidth"
   Call UserControl_Resize
   
End Property

Public Property Get BorderType() As enuBorderTypes
   
   BorderType = mudtBorderType
   
End Property

Public Property Let BorderType(ByVal vNewValue As enuBorderTypes)
   
   mudtBorderType = vNewValue
   PropertyChanged "BorderType"
   Call UserControl_Resize
   
End Property

Public Property Get Border3DHighlight() As OLE_COLOR
   
   Border3DHighlight = mlng3DHighlight
   
End Property

Public Property Let Border3DHighlight(ByVal vNewValue As OLE_COLOR)
   
   mlng3DHighlight = vNewValue
   lblCaptionShadow.ForeColor = mlng3DHighlight
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

Public Property Get Caption() As String
   
   Caption = lblCaption.Caption
   
End Property

Public Property Let Caption(ByVal vstrNewValue As String)
   
   lblCaption.Caption = vstrNewValue
   lblCaptionShadow.Caption = lblCaption.Caption
   PropertyChanged "Caption"
   Call UserControl_Resize
   
End Property

Public Property Get CaptionAlignment() As enuCaptionAlignment
   
   CaptionAlignment = mudtCaptionAlignment
   
End Property

Public Property Let CaptionAlignment(ByVal vNewValue As enuCaptionAlignment)
   
   mudtCaptionAlignment = vNewValue
   PropertyChanged "CaptionAlignment"
   Call UserControl_Resize
   
End Property

Public Property Get CaptionMAlignment() As AlignmentConstants
   
   CaptionMAlignment = lblCaption.Alignment
   
End Property

Public Property Let CaptionMAlignment(ByVal vNewValue As AlignmentConstants)
   
   lblCaption.Alignment = vNewValue
   lblCaptionShadow.Alignment = vNewValue
   PropertyChanged "CaptionMAlignment"
   Call UserControl_Resize
   
End Property

Public Property Get CaptionLocation() As enuCaptionLocation
   
   CaptionLocation = mudtCaptionLocation
   
End Property

Public Property Let CaptionLocation(ByVal vNewValue As enuCaptionLocation)
   
   mudtCaptionLocation = vNewValue
   PropertyChanged "CaptionLocation"
   Call UserControl_Resize
   
End Property

Public Property Get Caption3D() As enuCaption3D
   
   Caption3D = mudtCaption3D
   
End Property

Public Property Let Caption3D(ByVal newValue As enuCaption3D)
   
   mudtCaption3D = newValue
   PropertyChanged "Caption3D"
   Call UserControl_Resize
   
End Property
Public Property Get CornerRadius() As Long
   
   CornerRadius = mlngCornerRadius
   
End Property

Public Property Let CornerRadius(ByVal newValue As Long)
   
   mlngCornerRadius = newValue
   mlngCornerDia = mlngCornerRadius * 2
   PropertyChanged "CornerDiameter"
   Call UserControl_Resize
   
End Property

Private Sub DrawBevelInner()
   
  Dim sngBevelWidth As Single
  Dim lngCorner     As Long
   
   On Error GoTo Err_Proc
   
   If mudtBevelInner = 0 Then GoTo Exit_Proc
  
   sngBevelWidth = mlngBevelWidth + msngInsideBorder
   lblCaptionBlank.Visible = False
   
   Select Case mudtBorderType
   Case 7 To 13  '/* Rounded Corners
      lngCorner = mlngCornerRadius
   Case Else
      lngCorner = 0&
   End Select
   
   With UserControl
      Select Case mudtBevelInner
      Case 2 '/* Raised
      
         '/* Top
         UserControl.Line (sngBevelWidth + lngCorner, sngBevelWidth)-(.ScaleWidth - lngCorner - 1 - sngBevelWidth, sngBevelWidth), mlng3DHighlight
         '/* Left
         UserControl.Line (sngBevelWidth, lngCorner + sngBevelWidth)-(sngBevelWidth, .ScaleHeight - sngBevelWidth - lngCorner - 1), mlng3DHighlight
         '/* Right
         UserControl.Line (.ScaleWidth - sngBevelWidth - 1, lngCorner + sngBevelWidth)-(.ScaleWidth - sngBevelWidth - 1, .ScaleHeight - lngCorner - sngBevelWidth - 1), mlng3DShadow
         '/* Bottom
         UserControl.Line (lngCorner + sngBevelWidth, .ScaleHeight - sngBevelWidth - 1)-(.ScaleWidth - sngBevelWidth - lngCorner - 1, .ScaleHeight - sngBevelWidth - 1), mlng3DShadow
         
         If lngCorner Then
            '/* Top Left
            UserControl.Circle (lngCorner + sngBevelWidth, lngCorner + sngBevelWidth), lngCorner, mlng3DHighlight, 1.57, 3.14
            '/* Bottom Left
            UserControl.Circle (lngCorner + sngBevelWidth, .ScaleHeight - sngBevelWidth - lngCorner - 1), lngCorner, mlng3DHighlight, 3.14, 4.71
            '/* Top Right
            UserControl.Circle (.ScaleWidth - lngCorner - sngBevelWidth - 1, sngBevelWidth + lngCorner), lngCorner, mlng3DShadow, 6.28, 1.57
            '/* Bottom Right
            UserControl.Circle (.ScaleWidth - lngCorner - sngBevelWidth - 1, .ScaleHeight - sngBevelWidth - lngCorner - 1), lngCorner, mlng3DShadow, 4.71
         End If
         
      Case 1 '/* Inserted
         
         '/* Top
         UserControl.Line (sngBevelWidth + lngCorner, sngBevelWidth)-(.ScaleWidth - lngCorner - 1 - sngBevelWidth, sngBevelWidth), mlng3DShadow
         '/* Left
         UserControl.Line (sngBevelWidth, lngCorner + sngBevelWidth)-(sngBevelWidth, .ScaleHeight - sngBevelWidth - lngCorner - 1), mlng3DShadow
         '/* Right
         UserControl.Line (.ScaleWidth - sngBevelWidth - 1, lngCorner + sngBevelWidth)-(.ScaleWidth - sngBevelWidth - 1, .ScaleHeight - lngCorner - sngBevelWidth - 1), mlng3DHighlight
         '/* Bottom
         UserControl.Line (lngCorner + sngBevelWidth, .ScaleHeight - sngBevelWidth - 1)-(.ScaleWidth - sngBevelWidth - lngCorner - 1, .ScaleHeight - sngBevelWidth - 1), mlng3DHighlight
         
         If lngCorner Then
            '/* Top Left
            UserControl.Circle (lngCorner + sngBevelWidth, lngCorner + sngBevelWidth), lngCorner, mlng3DShadow, 1.57, 3.14
            '/* Bottom Left
            UserControl.Circle (lngCorner + sngBevelWidth, .ScaleHeight - sngBevelWidth - lngCorner - 1), lngCorner, mlng3DShadow, 3.14, 4.71
            '/* Top Right
            UserControl.Circle (.ScaleWidth - lngCorner - sngBevelWidth - 1, sngBevelWidth + lngCorner), lngCorner, mlng3DHighlight, 6.28, 1.57
            '/* Bottom Right
            UserControl.Circle (.ScaleWidth - lngCorner - sngBevelWidth - 1, .ScaleHeight - sngBevelWidth - lngCorner - 1), lngCorner, mlng3DHighlight, 4.71
         End If
         
      Case 3 '/* Flat Shadow
      
         With UserControl
            .ForeColor = mlng3DShadow
            Call RoundRect(.hdc, sngBevelWidth, sngBevelWidth, .ScaleWidth - sngBevelWidth, .ScaleHeight - sngBevelWidth, lngCorner, lngCorner)
         End With

      Case 4 '/* Flat Highlight
      
         With UserControl
            .ForeColor = mlng3DHighlight
            Call RoundRect(.hdc, sngBevelWidth, sngBevelWidth, .ScaleWidth - sngBevelWidth, .ScaleHeight - sngBevelWidth, lngCorner, lngCorner)
         End With

      End Select
   End With
   
   '/* Get inside workspace size in twips
   mlngInsideHeight = (UserControl.ScaleHeight - sngBevelWidth - sngBevelWidth - 2) * Screen.TwipsPerPixelY
   mlngInsideWidth = (UserControl.ScaleWidth - sngBevelWidth - sngBevelWidth - 2) * Screen.TwipsPerPixelX
   mlngInsideLeft = (sngBevelWidth + 1) * Screen.TwipsPerPixelY
   mlngInsideTop = (sngBevelWidth + 1) * Screen.TwipsPerPixelX
   

Exit_Proc:
   If mlngFloodValue > 0 Then Call DrawFlood
   
   Exit Sub
   
Err_Proc:
   ErrHandler False, Err.Number, Err.Description, "Frame3D", "DrawBevelInner"
   Err.Clear
   Resume Next
   
End Sub

Private Sub DrawCaptionAlignment()
   
  Dim sngWidth   As Single
  Dim sngHeight  As Single
  Dim sngOffset  As Long
   
   On Error GoTo Err_Proc
   
   If Len(lblCaption.Caption) = 0 Then GoTo Exit_Proc
   
   Select Case mudtBorderType
   Case 7 To 13 '/* Rounded Corners
      sngOffset = mlngCornerRadius
      If mudtCaptionLocation Then '/* [In Frame]
         sngOffset = sngOffset + 3
      End If
   Case 1 To 6  '/* Square Corners
      If mudtCaptionLocation Then '/* [In Frame]
         sngOffset = 6
      Else '/* [Inside]
         sngOffset = 3
      End If
   Case Else '/* No border
      sngOffset = 1
   End Select
   
   sngWidth = UserControl.ScaleWidth
   sngHeight = UserControl.ScaleHeight
   
   If mudtBevelInner Then
      sngWidth = UserControl.ScaleWidth - (mlngBevelWidth * 2)
      sngHeight = UserControl.ScaleHeight - (mlngBevelWidth * 2)
   End If
   
   Select Case mudtCaptionAlignment
   Case 0 '/* Top left
      msngLeft = sngOffset
      msngTop = 1
   Case 1 '/* Top Center
      msngLeft = (sngWidth - lblCaption.Width) / 2
      msngTop = 1
   Case 2 '/* Top right
      msngLeft = sngWidth - lblCaption.Width - sngOffset
      msngTop = 1
      
   Case 3 '/* Mid Left
      msngLeft = 3
      msngTop = (sngHeight - lblCaption.Height) / 2
   Case 4 '/* Mid Center
      msngLeft = (sngWidth - lblCaption.Width) / 2
      msngTop = (sngHeight - lblCaption.Height) / 2
   Case 5 '/* Mid Right
      msngLeft = sngWidth - lblCaption.Width - 3
      msngTop = (sngHeight - lblCaption.Height) / 2
      
   Case 6 '/* Bot Left
      msngLeft = sngOffset
      msngTop = sngHeight - lblCaption.Height - 3
   Case 7 '/* Bot Center
      msngLeft = (sngWidth - lblCaption.Width) / 2
      msngTop = sngHeight - lblCaption.Height - 3
   Case 8 '/* Bot Right
      msngLeft = sngWidth - lblCaption.Width - sngOffset
      msngTop = sngHeight - lblCaption.Height - 3
   End Select
   
   If mudtBevelInner Then
      msngLeft = msngLeft + mlngBevelWidth
      msngTop = msngTop + mlngBevelWidth
   End If
   
   Call DrawCaptionStyle
   
Exit_Proc:
   
   UserControl.Refresh
   
   Exit Sub
   
Err_Proc:
   ErrHandler False, Err.Number, Err.Description, "Frame3D", "DrawCaptionAlignment"
   Err.Clear
   Resume Next
   
End Sub

Private Sub DrawCaptionStyle()
   
   On Error GoTo Err_Proc
   
   lblCaption.Move msngLeft, msngTop
   
   Select Case mudtCaption3D
   Case 0 '/* Flat
      lblCaptionShadow.Visible = False
   Case 1 '/* Inserted
      lblCaptionShadow.Visible = True
      lblCaptionShadow.Move lblCaption.left + 1, lblCaption.top + 1
   Case 2 '/* Raised
      lblCaptionShadow.Visible = True
      lblCaptionShadow.Move lblCaption.left - 1, lblCaption.top - 1
   End Select
   
   lblCaptionBlank.Move lblCaption.left - 1, lblCaption.top - 1, lblCaption.Width + 2, lblCaption.Height + 2
   
   Exit Sub
   
Err_Proc:
   ErrHandler False, Err.Number, Err.Description, "Frame3D", "DrawCaptionStyle"
   Err.Clear
   Resume Next
   
End Sub

Private Sub DrawFlood()
   
  Dim sngBevelWidth As Single
  Dim sngNewValue   As Single
   
   On Error GoTo Err_Proc
   
   '/* Show percent complete?
   If mblnFloodShowPct Then
      lblCaption.Caption = CStr(mlngFloodValue) & "%"
      lblCaptionShadow.Caption = CStr(mlngFloodValue) & "%"
      Call DrawCaptionAlignment
   End If
   
   If mlngFloodValue Then
      '/* Is there an inside border showing?
      If mudtBevelInner Then
         sngBevelWidth = mlngBevelWidth + msngInsideBorder + 1
      Else
         sngBevelWidth = msngInsideBorder + 1
      End If
      
      '/* Flood Fill
      If mudtFloodType Then  '/* [Bottom To Top]
         
         sngNewValue = UserControl.ScaleHeight - sngBevelWidth - sngBevelWidth - 1
         sngNewValue = sngNewValue - (sngNewValue * (mlngFloodValue / 100))
         
         UserControl.Line (sngBevelWidth, UserControl.ScaleHeight - sngBevelWidth - 1)- _
                          (UserControl.ScaleWidth - sngBevelWidth - 1, sngNewValue + sngBevelWidth), _
                          mlngFloodColor, BF
      
      Else '/[Left To Right]
         
         sngNewValue = (UserControl.ScaleWidth - sngBevelWidth - sngBevelWidth) * mlngFloodValue / 100
         
         UserControl.Line (sngBevelWidth, sngBevelWidth)- _
                          (sngNewValue + sngBevelWidth - 1, UserControl.ScaleHeight - sngBevelWidth - 1), _
                          mlngFloodColor, BF
      End If
   End If
   
Exit_Proc:
   Exit Sub
   
Err_Proc:
   ErrHandler False, Err.Number, Err.Description, "Frame3D", "DrawFlood"
   Err.Clear
   Resume Next
   
End Sub

Public Sub DrawGradiant()
  
  Dim lngI    As Long
  Dim lngStep As Long
  Dim sngRed1 As Single
  Dim sngGrn1 As Single
  Dim sngBlu1 As Single
  Dim sngRed2 As Single
  Dim sngGrn2 As Single
  Dim sngBlu2 As Single
      
    On Error Resume Next
    
    Call GetRGBColor(UserControl.FillColor, sngRed1, sngGrn1, sngBlu1)
    Call GetRGBColor(UserControl.BackColor, sngRed2, sngGrn2, sngBlu2)
    
    With UserControl
        Select Case mudtFillGradient
        Case 1 '/* [Horizontal]
            lngStep = .ScaleWidth - mlngCornerRadius
            '/* Get gradient color step
            sngRed2 = (sngRed2 - sngRed1) / lngStep
            sngGrn2 = (sngGrn2 - sngGrn1) / lngStep
            sngBlu2 = (sngBlu2 - sngBlu1) / lngStep
            '/* Begin drawing horizontal gradient
            For lngI = 0 To lngStep
                UserControl.Line (lngI, 0)-(lngI, .ScaleHeight), RGB(CInt(sngRed1), CInt(sngGrn1), CInt(sngBlu1))
                sngRed1 = sngRed1 + sngRed2
                sngGrn1 = sngGrn1 + sngGrn2
                sngBlu1 = sngBlu1 + sngBlu2
            Next lngI
            
        Case 2 '/* [Vertical]
            lngStep = .ScaleHeight - mlngCornerRadius
            '/* Get gradient color step
            sngRed2 = (sngRed2 - sngRed1) / lngStep
            sngGrn2 = (sngGrn2 - sngGrn1) / lngStep
            sngBlu2 = (sngBlu2 - sngBlu1) / lngStep
            '/* Begin drawing vertical gradient
            For lngI = 0 To lngStep
                UserControl.Line (0, lngI)-(.ScaleWidth, lngI), RGB(CInt(sngRed1), CInt(sngGrn1), CInt(sngBlu1))
                sngRed1 = sngRed1 + sngRed2
                sngGrn1 = sngGrn1 + sngGrn2
                sngBlu1 = sngBlu1 + sngBlu2
            Next lngI
        
        End Select
    End With
    
End Sub

Public Property Get DrawStyle() As DrawStyleConstants
   
   DrawStyle = UserControl.DrawStyle
   
End Property

Public Property Let DrawStyle(ByVal vNewValue As DrawStyleConstants)
   
   UserControl.DrawStyle = vNewValue
   PropertyChanged "DrawStyle"
   Call UserControl_Resize
   
End Property

Public Property Get Enabled() As Boolean
   
   Enabled = mblnEnabled
   
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
   
   mblnEnabled = vNewValue
   PropertyChanged "Enabled"
   Call UserControl_Resize
   
End Property

Private Sub ErrHandler(Optional ByVal vblnDisplayError As Boolean = True, _
                       Optional ByVal vstrErrNumber As String = vbNullString, _
                       Optional ByVal vstrErrDescription As String = vbNullString, _
                       Optional ByVal vstrModuleName As String = vbNullString, _
                       Optional ByVal vstrProcName As String = vbNullString)
   
  Dim strTemp As String
  Dim lngFN   As Long
   
   '/* Purpose: Error handling - On Error
   
   '/* Show Error Message
   If vblnDisplayError Then
      strTemp = "Error occured: "
      If Len(vstrErrNumber) > 0 Then
         strTemp = strTemp & vstrErrNumber & vbNewLine
      Else
         strTemp = strTemp & vbNewLine
      End If
      If Len(vstrErrDescription) > 0 Then strTemp = strTemp & "Description: " & vstrErrDescription & vbNewLine
      If Len(vstrModuleName) > 0 Then strTemp = strTemp & "Module: " & vstrModuleName & vbNewLine
      If Len(vstrProcName) > 0 Then strTemp = strTemp & "Function: " & vstrProcName
      MsgBox strTemp, vbCritical, App.Title & " - ERROR"
   End If
   
   '/* Write error log
   lngFN = FreeFile
   Open App.Path & "\ErrorLog.txt" For Append As #lngFN
   Write #lngFN, Now, vstrErrNumber, vstrErrDescription, vstrModuleName, vstrProcName, _
     App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision, Environ$("username"), Environ$("computername")
   Close #lngFN
   
End Sub

Public Property Get FillColor() As OLE_COLOR
   
   FillColor = UserControl.FillColor
   
End Property

Public Property Let FillColor(ByVal vNewValue As OLE_COLOR)
   
   UserControl.FillColor = vNewValue
   PropertyChanged "FillColor"
   Call UserControl_Resize
   
End Property

Public Property Get FillGradient() As enuFillGradient
   
   FillGradient = mudtFillGradient

End Property

Public Property Let FillGradient(ByVal vNewValue As enuFillGradient)

   mudtFillGradient = vNewValue
   PropertyChanged "FillGradient"
   Call UserControl_Resize

End Property

Public Property Get FillStyle() As FillStyleConstants
   
   FillStyle = UserControl.FillStyle
   
End Property

Public Property Let FillStyle(ByVal vNewValue As FillStyleConstants)
   
   UserControl.FillStyle = vNewValue
   PropertyChanged "FillStyle"
   Call UserControl_Resize
   
End Property

Public Property Get FloodColor() As OLE_COLOR
   
   FloodColor = mlngFloodColor
   
End Property

Public Property Let FloodColor(ByVal vNewValue As OLE_COLOR)
   
   mlngFloodColor = vNewValue
   PropertyChanged "FloodColor"
   Call DrawFlood
   
End Property

Public Property Get FloodPercent() As Long
   
   FloodPercent = mlngFloodValue
   
End Property

Public Property Let FloodPercent(ByVal vNewValue As Long)
   
   '/* Fix the value
   Select Case vNewValue
   Case Is > 100
      vNewValue = 100
   Case Is < 0
      vNewValue = 0
   End Select
   
   '/* Clear old values if decreasing
   If vNewValue <= mlngFloodValue Then
      '/* Save new property value
      mlngFloodValue = vNewValue
      PropertyChanged "FloodPercent"
      Call DrawFlood
      Call UserControl_Resize
   Else
      '/* Save new property value
      mlngFloodValue = vNewValue
      PropertyChanged "FloodPercent"
      Call DrawFlood
   End If
   
End Property

Public Property Get FloodShowPct() As Boolean
   
   FloodShowPct = mblnFloodShowPct
   
End Property

Public Property Let FloodShowPct(ByVal vNewValue As Boolean)
   
   mblnFloodShowPct = vNewValue
   PropertyChanged "FloodShowPct"
   
   If mblnFloodShowPct Then
      lblCaption.Caption = CStr(mlngFloodValue) & "%"
      lblCaptionShadow.Caption = CStr(mlngFloodValue) & "%"
   Else
      lblCaption.Caption = vbNullString
      lblCaptionShadow.Caption = vbNullString
   End If
   
End Property

Public Property Get FloodType() As enuFloodType
   
   FloodType = mudtFloodType
   
End Property

Public Property Let FloodType(ByVal vNewValue As enuFloodType)
   
   mudtFloodType = vNewValue
   PropertyChanged "FloodType"
   Call UserControl_Resize
   
End Property

Public Property Get Font() As Font
   
   Set Font = lblCaption.Font
   
End Property

Public Property Set Font(ByRef vNewValue As Font)
   
   Set lblCaption.Font = vNewValue
   Set lblCaptionShadow.Font = vNewValue
   PropertyChanged "Font"
   Call UserControl_Resize
   
End Property

Public Property Get FontBold() As Boolean
   
   FontBold = lblCaption.FontBold
   
End Property

Public Property Let FontBold(ByVal vNewValue As Boolean)
   
   lblCaption.FontBold = vNewValue
   lblCaptionShadow.FontBold = vNewValue
   PropertyChanged "FontBold"
   Call UserControl_Resize
   
End Property

Public Property Get FontItalic() As Boolean
   
   FontItalic = lblCaption.FontItalic
   
End Property

Public Property Let FontItalic(ByVal vNewValue As Boolean)
   
   lblCaption.FontItalic = vNewValue
   lblCaptionShadow.FontItalic = vNewValue
   PropertyChanged "FontItalic"
   Call UserControl_Resize
   
End Property

Public Property Get FontName() As String
   
   FontName = lblCaption.FontName
   
End Property

Public Property Let FontName(ByVal vNewValue As String)
   
   lblCaption.FontName = vNewValue
   lblCaptionShadow.FontName = vNewValue
   PropertyChanged "FontName"
   Call UserControl_Resize
   
End Property

Public Property Get FontSize() As Long
   
   FontSize = lblCaption.FontSize
   
End Property

Public Property Let FontSize(ByVal vNewValue As Long)
   
   lblCaption.FontSize = vNewValue
   lblCaptionShadow.FontSize = vNewValue
   PropertyChanged "FontSize"
   Call UserControl_Resize
   
End Property

Public Property Get FontUnderline() As Boolean
   
   FontUnderline = lblCaption.FontUnderline
   
End Property

Public Property Let FontUnderline(ByVal vNewValue As Boolean)
   
   lblCaption.FontUnderline = vNewValue
   lblCaptionShadow.FontUnderline = vNewValue
   PropertyChanged "FontUnderline"
   
End Property

Public Property Get ForeColor() As OLE_COLOR
   
   ForeColor = lblCaption.ForeColor
   
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
   
   lblCaption.ForeColor = vNewValue
   PropertyChanged "ForeColor"
   
End Property

Private Sub GetRGBColor(ByVal vlngColor As Long, _
                        ByRef rsngRed As Single, _
                        ByRef rsngGrn As Single, _
                        ByRef rsngBlu As Single)
    
    '/* Is the color a VB color constant?
    If vlngColor < 0 Then
        '/* Retrieves the current color of the specified display element
        vlngColor = GetSysColor(vlngColor And &HFF&)
    End If

    '/* Separate the color into it's RGB values
    rsngRed = CSng((vlngColor And &HFF&))
    rsngGrn = CSng((vlngColor And &HFF00&) \ &H100&)
    rsngBlu = CSng((vlngColor And &HFF0000) \ &H10000)
    
    '/* These values would normally be declared as Longs but
    '/* the calling sub requires Singles

End Sub

Public Property Get hdc() As Long

    hdc = UserControl.hdc

End Property

Public Property Get hWnd() As Long

    hWnd = UserControl.hWnd

End Property

Public Property Get InsideHeight() As Long
   
   InsideHeight = mlngInsideHeight

End Property

Public Property Get InsideLeft() As Long
   
   InsideLeft = mlngInsideLeft

End Property

Public Property Get InsideTop() As Long
   
   InsideTop = mlngInsideTop

End Property

Public Property Get InsideWidth() As Long
   
   InsideWidth = mlngInsideWidth

End Property

Public Property Get MouseIcon() As StdPicture
   
   Set MouseIcon = UserControl.MouseIcon
   
End Property

Public Property Set MouseIcon(ByVal vNewValue As StdPicture)
   
   On Local Error Resume Next
   Set UserControl.MouseIcon = vNewValue
   PropertyChanged "MouseIcon"
   On Local Error GoTo 0
   
End Property

Public Property Get MousePointer() As MousePointerConstants
   
   MousePointer = UserControl.MousePointer
   
End Property

Public Property Let MousePointer(ByVal vNewValue As MousePointerConstants)
   
   UserControl.MousePointer = vNewValue
   PropertyChanged "MousePointer"
   
End Property

Public Property Get Picture() As StdPicture
   
   Set Picture = UserControl.Picture
   
End Property

Public Property Set Picture(ByVal vNewValue As StdPicture)
   
   On Local Error Resume Next
   Set UserControl.Picture = vNewValue
   PropertyChanged "Picture"
   On Local Error GoTo 0
   
End Property

Public Property Get UseMnemonic() As Boolean
   
   UseMnemonic = lblCaption.UseMnemonic
   
End Property

Public Property Let UseMnemonic(ByVal vNewValue As Boolean)
   
   lblCaption.UseMnemonic = vNewValue
   lblCaptionShadow.UseMnemonic = vNewValue
   PropertyChanged "UseMnemonic"
   Call UserControl_Resize
   
End Property

Private Sub lblCaption_Click()
   RaiseEvent Click
End Sub

Private Sub lblCaption_DblClick()
   RaiseEvent DblClick
End Sub


Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   RaiseEvent MouseDown(Button, Shift, X, y)
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
   RaiseEvent MouseMove(Button, Shift, X, y)
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
   RaiseEvent MouseUp(Button, Shift, X, y)
End Sub

Private Sub lblCaptionBlank_Click()
   RaiseEvent Click
End Sub


Private Sub lblCaptionBlank_DblClick()
   RaiseEvent DblClick
End Sub


Private Sub lblCaptionBlank_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   RaiseEvent MouseDown(Button, Shift, X, y)
End Sub

Private Sub lblCaptionBlank_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
   RaiseEvent MouseMove(Button, Shift, X, y)
End Sub

Private Sub lblCaptionBlank_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
   RaiseEvent MouseUp(Button, Shift, X, y)
End Sub

Private Sub lblCaptionShadow_Click()
   RaiseEvent Click
End Sub


Private Sub lblCaptionShadow_DblClick()
   RaiseEvent DblClick
End Sub


Private Sub lblCaptionShadow_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   RaiseEvent MouseDown(Button, Shift, X, y)
End Sub

Private Sub lblCaptionShadow_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
   RaiseEvent MouseMove(Button, Shift, X, y)
End Sub

Private Sub lblCaptionShadow_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
   RaiseEvent MouseUp(Button, Shift, X, y)
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
   Call UserControl_Click
End Sub

Private Sub UserControl_Click()
   RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
   RaiseEvent DblClick
End Sub

Private Sub UserControl_InitProperties()
   
   On Error GoTo Err_Proc
   
   mlngBevelWidth = 3&
   mudtBorderType = 1 '/* [Frame Inserted]
   mlngBackColor = UserControl.Parent.BackColor
   mlng3DHighlight = vb3DHighlight
   mlng3DShadow = vb3DShadow
   mblnEnabled = True
   mlngFloodValue = 0&
   mblnFloodShowPct = False
   mudtFloodType = 0 '/* [Left To Right]
   mlngFloodColor = UserControl.FillColor
   lblCaptionShadow.ForeColor = mlng3DHighlight
   lblCaption.UseMnemonic = False
   lblCaptionShadow.UseMnemonic = False
   mudtCaptionLocation = [Inside Frame]
   mlngCornerRadius = 7&
   mlngCornerDia = mlngCornerRadius * 2

   
Exit_Proc:
   Exit Sub
   
Err_Proc:
   ErrHandler True, Err.Number, Err.Description, "Frame3D", "UserControl_InitProperties"
   Err.Clear
   Resume Next
   
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
   
   RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub


Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   RaiseEvent MouseDown(Button, Shift, X, y)
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
   RaiseEvent MouseMove(Button, Shift, X, y)
End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
   RaiseEvent MouseUp(Button, Shift, X, y)
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   
   On Error GoTo Err_Proc
   
   With PropBag
      mudtBorderType = .ReadProperty("BorderType", 1)
      mlngBevelWidth = .ReadProperty("BevelWidth", 3&)
      mudtBevelInner = .ReadProperty("BevelInner", 0)
      mudtCaption3D = .ReadProperty("Caption3D", 0)
      mudtCaptionAlignment = .ReadProperty("CaptionAlignment", 0)
      mudtCaptionLocation = .ReadProperty("CaptionLocation", 0)
      mlngBackColor = .ReadProperty("BackColor", UserControl.Parent.BackColor)
      mlng3DHighlight = .ReadProperty("Border3DHighlight", vb3DHighlight)
      mlng3DShadow = .ReadProperty("Border3DShadow", vb3DShadow)
      mblnEnabled = .ReadProperty("Enabled", True)
      mlngCornerRadius = .ReadProperty("CornerDiameter", 7&)
      mlngCornerDia = mlngCornerRadius * 2
      
      mlngFloodValue = .ReadProperty("FloodPercent", 0&)
      mblnFloodShowPct = .ReadProperty("FloodShowPct", 0)
      mudtFloodType = .ReadProperty("FloodType", 0)
      mlngFloodColor = .ReadProperty("FloodColor", UserControl.FillColor)
      mudtFillGradient = .ReadProperty("FillGradient", 0)
      
      UserControl.FillColor = .ReadProperty("FillColor", UserControl.FillColor)
      UserControl.FillStyle = .ReadProperty("FillStyle", UserControl.FillStyle)
      UserControl.DrawStyle = .ReadProperty("DrawStyle", UserControl.DrawStyle)
      UserControl.MousePointer = .ReadProperty("MousePointer", UserControl.MousePointer)
      UserControl.MouseIcon = .ReadProperty("MouseIcon", UserControl.MouseIcon)
      UserControl.Picture = .ReadProperty("Picture", UserControl.Picture)
      
      lblCaption.Alignment = .ReadProperty("CaptionMAlignment", lblCaption.Alignment)
      lblCaption.Font = .ReadProperty("Font", lblCaption.Font)
      lblCaption.FontBold = .ReadProperty("FontBold", lblCaption.FontBold)
      lblCaption.FontItalic = .ReadProperty("FontItalic", lblCaption.FontItalic)
      lblCaption.FontName = .ReadProperty("FontName", lblCaption.FontName)
      lblCaption.FontSize = .ReadProperty("FontSize", lblCaption.FontSize)
      lblCaption.FontStrikethru = .ReadProperty("FontStrikethru", lblCaption.FontStrikethru)
      lblCaption.FontUnderline = .ReadProperty("FontUnderline", lblCaption.FontUnderline)
      lblCaption.ForeColor = .ReadProperty("ForeColor", lblCaption.ForeColor)
      lblCaption.Caption = .ReadProperty("Caption", lblCaption.Caption)
      lblCaption.UseMnemonic = .ReadProperty("UseMnemonic", lblCaption.UseMnemonic)
      
   End With
   
   If mblnFloodShowPct Then
      lblCaption.Caption = CStr(mlngFloodValue) & "%"
      lblCaptionShadow.Caption = CStr(mlngFloodValue) & "%"
   Else
      '/* Trick to fix right justified text
      lblCaption.Caption = lblCaption.Caption & " "
      lblCaption.Caption = left$(lblCaption.Caption, Len(lblCaption.Caption) - 1)
   End If
   
   With lblCaptionShadow
      .Alignment = lblCaption.Alignment
      .Font = lblCaption.Font
      .FontBold = lblCaption.FontBold
      .FontItalic = lblCaption.FontItalic
      .FontName = lblCaption.FontName
      .FontSize = lblCaption.FontSize
      .FontStrikethru = lblCaption.FontStrikethru
      .FontUnderline = lblCaption.FontUnderline
      .Caption = lblCaption.Caption
      .UseMnemonic = lblCaption.UseMnemonic
   End With
   
   Exit Sub
   
Err_Proc:
   ErrHandler True, Err.Number, Err.Description, "Frame3D", "UserControl_ReadProperties"
   Err.Clear
   Resume Next
   
End Sub

Private Sub UserControl_Resize()
   
  Dim sngBottom     As Single
  Dim sngTop        As Single
  Dim mlngCornerDia As Long
   
   On Error GoTo Err_Proc
   
   sngBottom = UserControl.ScaleHeight
   sngTop = 0!
   lblCaptionBlank.Visible = False
   
   If mudtBevelInner = 0 Then
      Select Case mudtCaptionAlignment
      Case 0, 1, 2 '/* [Top Left], [Top Center], [Top Right]
         If mudtCaptionLocation = 1 And LenB(lblCaption.Caption) Then
            sngTop = lblCaption.Height / 2.25!
            lblCaptionBlank.Visible = True
         End If
      Case 6, 7, 8 '/* [Bottom Left], [Bottom Center], [Bottom Right]
         If mudtCaptionLocation = 1 And LenB(lblCaption.Caption) Then
            sngBottom = UserControl.ScaleHeight - (lblCaption.Height / 2.25!)
            lblCaptionBlank.Visible = True
         End If
      End Select
   End If
   
   msngInsideBorder = sngTop
   mlngCornerDia = mlngCornerRadius * 2
   
   With UserControl
      .DrawMode = vbCopyPen
      .Cls
      .BackColor = mlngBackColor
      lblCaptionBlank.BackColor = mlngBackColor
      .Enabled = mblnEnabled
      If mudtFillGradient Then Call DrawGradiant
      
      
      '/* Get inside workspace size in twips
      Select Case mudtBorderType
      Case 1, 2, 7, 8 '/* Frame Inserted/Raised
         mlngInsideHeight = (.ScaleHeight - 4) * Screen.TwipsPerPixelY
         mlngInsideWidth = (.ScaleWidth - 4) * Screen.TwipsPerPixelX
         mlngInsideLeft = 2 * Screen.TwipsPerPixelX
         mlngInsideTop = 2 * Screen.TwipsPerPixelY
      Case Else
         mlngInsideHeight = (.ScaleHeight - 2) * Screen.TwipsPerPixelY
         mlngInsideWidth = (.ScaleWidth - 2) * Screen.TwipsPerPixelX
         mlngInsideLeft = Screen.TwipsPerPixelX
         mlngInsideTop = Screen.TwipsPerPixelY
      End Select
  
      
      '/* Draw Border Type
      Select Case mudtBorderType
      Case 1 '/* [Frame Inserted]
         .ForeColor = mlng3DShadow
         Call RoundRect(.hdc, 0&, sngTop, .ScaleWidth - 1, sngBottom - 1, 0&, 0&)
         .ForeColor = mlng3DHighlight
         Call RoundRect(.hdc, 1&, sngTop + 1&, .ScaleWidth, sngBottom, 0&, 0&)
         Call SetWindowRgn(.hWnd, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, 0&, 0&), True)
         
      Case 2 '/* [Frame Raised]
         .ForeColor = mlng3DHighlight
         Call RoundRect(.hdc, 0&, sngTop, .ScaleWidth - 1, sngBottom - 1, 0&, 0&)
         .ForeColor = mlng3DShadow
         Call RoundRect(.hdc, 1&, sngTop + 1&, .ScaleWidth, sngBottom, 0&, 0&)
         Call SetWindowRgn(.hWnd, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, 0&, 0&), True)
      
      Case 3 '/* [Panel Flat Shadow]
         .ForeColor = mlng3DShadow
         Call RoundRect(.hdc, 0&, sngTop, .ScaleWidth - 1, sngBottom - 1, 0&, 0&)
         Call SetWindowRgn(.hWnd, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, 0&, 0&), True)
      
      Case 4 '/* [Panel Flat Highlight]
         .ForeColor = mlng3DHighlight
         Call RoundRect(.hdc, 0&, sngTop, .ScaleWidth, sngBottom, 0&, 0&)
         Call SetWindowRgn(.hWnd, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, 0&, 0&), True)
      
      Case 5 '/* [Panel Raised]
         lblCaptionBlank.Visible = False
         '/* Trick to allow FillStyle to work
         .ForeColor = .BackColor
         Call RoundRect(.hdc, 0&, 0&, .ScaleWidth, .ScaleHeight, 0&, 0&)
         UserControl.Line (.ScaleWidth - 1, .ScaleHeight - 1)-(.ScaleWidth - 1, 0), mlng3DShadow
         UserControl.Line (.ScaleWidth - 1, .ScaleHeight - 1)-(0, .ScaleHeight - 1), mlng3DShadow
         UserControl.Line (.ScaleWidth - 1, 0)-(0, 0), mlng3DHighlight
         UserControl.Line (0, .ScaleHeight - 1)-(0, 0), mlng3DHighlight
         Call SetWindowRgn(.hWnd, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, 0&, 0&), True)
         
      Case 6 '/* [Inserted Panel Square Corners]
         lblCaptionBlank.Visible = False
         '/* Trick to allow FillStyle to work
         .ForeColor = .BackColor
         Call RoundRect(.hdc, 0&, 0&, .ScaleWidth, .ScaleHeight, 0&, 0&)
         UserControl.Line (.ScaleWidth - 1, .ScaleHeight - 1)-(.ScaleWidth - 1, 0), mlng3DHighlight
         UserControl.Line (.ScaleWidth - 1, .ScaleHeight - 1)-(0, .ScaleHeight - 1), mlng3DHighlight
         UserControl.Line (.ScaleWidth - 2, 0)-(0, 0), mlng3DShadow
         UserControl.Line (0, .ScaleHeight - 2)-(0, 0), mlng3DShadow
         Call SetWindowRgn(.hWnd, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, 0&, 0&), True)
      
      Case 7 '/* [rFrame Inserted]
         .ForeColor = mlng3DShadow
         Call RoundRect(.hdc, 0&, sngTop, .ScaleWidth - 1, sngBottom - 1, mlngCornerDia, mlngCornerDia)
         .ForeColor = mlng3DHighlight
         Call RoundRect(.hdc, 1&, sngTop + 1&, .ScaleWidth, sngBottom, mlngCornerDia, mlngCornerDia)
         '/* Make corners transparent
         Call SetWindowRgn(.hWnd, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, mlngCornerDia, mlngCornerDia), True)
               
      Case 8 '/* [rFrame Raised]
         .ForeColor = mlng3DHighlight
         Call RoundRect(.hdc, 0&, sngTop, .ScaleWidth - 1, sngBottom - 1, mlngCornerDia, mlngCornerDia)
         .ForeColor = mlng3DShadow
         Call RoundRect(.hdc, 1&, sngTop + 1&, .ScaleWidth, sngBottom, mlngCornerDia, mlngCornerDia)
         '/* Make corners transparent
         Call SetWindowRgn(.hWnd, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, mlngCornerDia, mlngCornerDia), True)
                  
      Case 9 '/* [rPanel Flat Shadow]
         .ForeColor = mlng3DShadow
         Call RoundRect(.hdc, 0&, sngTop, .ScaleWidth, sngBottom, mlngCornerDia, mlngCornerDia)
         Call SetWindowRgn(.hWnd, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, mlngCornerDia, mlngCornerDia), True)
               
      Case 10 '/* [rPanel Flat Highlight]
         .ForeColor = mlng3DHighlight
         Call RoundRect(.hdc, 0&, sngTop, .ScaleWidth, sngBottom, mlngCornerDia, mlngCornerDia)
         Call SetWindowRgn(.hWnd, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, mlngCornerDia, mlngCornerDia), True)
         
      Case 11 '/* [rPanel Raised]
         lblCaptionBlank.Visible = False
         '/* Trick to allow FillStyle to work
         .ForeColor = .BackColor
         Call RoundRect(.hdc, 0&, 0&, .ScaleWidth, .ScaleHeight, mlngCornerDia, mlngCornerDia)
         '/* Top
         UserControl.Line (mlngCornerRadius, 0)-(.ScaleWidth - mlngCornerRadius - 1, 0), mlng3DHighlight
         '/* Left
         UserControl.Line (0, mlngCornerRadius)-(0, .ScaleHeight - mlngCornerRadius - 1), mlng3DHighlight
         '/* Right
         UserControl.Line (.ScaleWidth - 1, mlngCornerRadius)-(.ScaleWidth - 1, .ScaleHeight - mlngCornerRadius - 1), mlng3DShadow
         '/* Bottom
         UserControl.Line (mlngCornerRadius, .ScaleHeight - 1)-(.ScaleWidth - mlngCornerRadius - 1, .ScaleHeight - 1), mlng3DShadow
         '/* Top Left
         UserControl.Circle (mlngCornerRadius, mlngCornerRadius), mlngCornerRadius, mlng3DHighlight, 1.57, 3.14
         '/* Bottom Left
         UserControl.Circle (mlngCornerRadius, .ScaleHeight - mlngCornerRadius - 1), mlngCornerRadius, mlng3DHighlight, 3.14, 4.71
         '/* Top Right
         UserControl.Circle (.ScaleWidth - mlngCornerRadius - 1, mlngCornerRadius), mlngCornerRadius, mlng3DShadow, 6.28, 1.57
         '/* Bottom Right
         UserControl.Circle (.ScaleWidth - mlngCornerRadius - 1, .ScaleHeight - mlngCornerRadius - 1), mlngCornerRadius, mlng3DShadow, 4.71
         Call SetWindowRgn(.hWnd, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, mlngCornerDia, mlngCornerDia), True)
         
      Case 12 '/* [Inserted Panel Round Corners]
         lblCaptionBlank.Visible = False
         '/* Trick to allow FillStyle to work
         .ForeColor = .BackColor
         Call RoundRect(.hdc, 0&, 0&, .ScaleWidth, .ScaleHeight, mlngCornerDia, mlngCornerDia)
         '/* Top
         UserControl.Line (mlngCornerRadius, 0)-(.ScaleWidth - mlngCornerRadius - 1, 0), mlng3DShadow
         '/* Left
         UserControl.Line (0, mlngCornerRadius)-(0, .ScaleHeight - mlngCornerRadius - 1), mlng3DShadow
         '/* Right
         UserControl.Line (.ScaleWidth - 1, mlngCornerRadius)-(.ScaleWidth - 1, .ScaleHeight - mlngCornerRadius - 1), mlng3DHighlight
         '/* Bottom
         UserControl.Line (mlngCornerRadius, .ScaleHeight - 1)-(.ScaleWidth - mlngCornerRadius - 1, .ScaleHeight - 1), mlng3DHighlight
         '/* Top Left
         UserControl.Circle (mlngCornerRadius, mlngCornerRadius), mlngCornerRadius, mlng3DShadow, 1.57, 3.14
         '/* Bottom Left
         UserControl.Circle (mlngCornerRadius, .ScaleHeight - mlngCornerRadius - 1), mlngCornerRadius, mlng3DShadow, 3.14, 4.71
         '/* Top Right
         UserControl.Circle (.ScaleWidth - mlngCornerRadius - 1, mlngCornerRadius), mlngCornerRadius, mlng3DHighlight, 6.28, 1.57
         '/* Bottom Right
         UserControl.Circle (.ScaleWidth - mlngCornerRadius - 1, .ScaleHeight - mlngCornerRadius - 1), mlngCornerRadius, mlng3DHighlight, 4.71
         Call SetWindowRgn(.hWnd, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, mlngCornerRadius, mlngCornerRadius), True)
         
      Case Else '/* [None]
         '/* No Border
         '/* Trick to allow FillStyle to work
         .ForeColor = .BackColor
         Call RoundRect(.hdc, 0&, 0&, .ScaleWidth, .ScaleHeight, 0&, 0&)
         If mudtBorderType = 13 Then '/* Rounded region
            Call SetWindowRgn(.hWnd, CreateRoundRectRgn(1&, 1&, .ScaleWidth, .ScaleHeight, mlngCornerDia, mlngCornerDia), True)
         Else
            Call SetWindowRgn(.hWnd, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, 0&, 0&), True)
         End If
      
      End Select
      
   End With
   
   Call DrawBevelInner
   Call DrawCaptionAlignment
   
   Exit Sub
   
Err_Proc:
   ErrHandler False, Err.Number, Err.Description, "Frame3D", "UserControl_Resize"
   Err.Clear
   Resume Next
   
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   
   On Error GoTo Err_Proc
   
   With PropBag
      .WriteProperty "BorderType", mudtBorderType
      .WriteProperty "BevelWidth", mlngBevelWidth
      .WriteProperty "BevelInner", mudtBevelInner
      .WriteProperty "Caption3D", mudtCaption3D
      .WriteProperty "CaptionAlignment", mudtCaptionAlignment
      .WriteProperty "CaptionLocation", mudtCaptionLocation
      .WriteProperty "BackColor", mlngBackColor
      .WriteProperty "CornerDiameter", mlngCornerRadius
      
      .WriteProperty "FillColor", UserControl.FillColor
      .WriteProperty "FillStyle", UserControl.FillStyle
      .WriteProperty "DrawStyle", UserControl.DrawStyle
      
      .WriteProperty "FloodPercent", mlngFloodValue
      .WriteProperty "FloodShowPct", mblnFloodShowPct
      .WriteProperty "FloodType", mudtFloodType
      .WriteProperty "FloodColor", mlngFloodColor
      .WriteProperty "FillGradient", mudtFillGradient
      
      .WriteProperty "MousePointer", UserControl.MousePointer
      .WriteProperty "MouseIcon", UserControl.MouseIcon
      .WriteProperty "Picture", UserControl.Picture
      
      .WriteProperty "Border3DHighlight", mlng3DHighlight
      .WriteProperty "Border3DShadow", mlng3DShadow
      
      .WriteProperty "Enabled", mblnEnabled
      
      .WriteProperty "CaptionMAlignment", lblCaption.Alignment
      .WriteProperty "Font", lblCaption.Font
      .WriteProperty "FontBold", lblCaption.FontBold
      .WriteProperty "FontItalic", lblCaption.FontItalic
      .WriteProperty "FontName", lblCaption.FontName
      .WriteProperty "FontSize", lblCaption.FontSize
      .WriteProperty "FontStrikethru", lblCaption.FontStrikethru
      .WriteProperty "FontUnderline", lblCaption.FontUnderline
      .WriteProperty "ForeColor", lblCaption.ForeColor
      .WriteProperty "Caption", lblCaption.Caption
      .WriteProperty "UseMnemonic", lblCaption.UseMnemonic
   End With
   
   Exit Sub
   
Err_Proc:
   ErrHandler True, Err.Number, Err.Description, "Frame3D", "UserControl_WriteProperties"
   Err.Clear
   Resume Next
   
End Sub
