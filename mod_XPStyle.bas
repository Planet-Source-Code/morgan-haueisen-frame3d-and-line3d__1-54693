Attribute VB_Name = "mod_XPStyle"
'/* Author: Morgan Haueisen (morganh@hartcom.net)
'/* Copyright (c) 2003
'/* Version 1.0.0

Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long

'/* Operating system version information
Private Type OSVersionInfo
    OSVSize       As Long
    dwVerMajor    As Long
    dwVerMinor    As Long
    dwBuildNumber As Long
    PlatformID    As Long
    szCSDVersion  As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVersionInfo) As Long

Public gblnOsIsXp As Boolean

Public Sub ManifestWrite(Optional ByVal vblnOnlyIfXP As Boolean = True)

  Dim lngFN         As Long
  Dim strEXEName    As String
  Dim strXPLookXML  As String
  Dim osv           As OSVersionInfo
    
    '/* Get OS compatability flag
   osv.OSVSize = Len(osv)
   If GetVersionEx(osv) = 1 Then
      If osv.PlatformID = 2 Then
         If osv.dwVerMajor >= 5 Then
            If osv.dwVerMinor = 1 Then
               gblnOsIsXp = True '/* OS is XP
            End If
         End If
      End If
   End If
  
   '/* If OS is XP or force always write then continue
   If gblnOsIsXp Or Not vblnOnlyIfXP Then
      '/* Standard manifest file as defined at:
      '/* http://support.microsoft.com/default.aspx?scid=kb;en-us;309366
      strXPLookXML = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf & _
         "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">" & vbCrLf & _
         "<assemblyIdentity" & vbCrLf & _
         "   version=""1.0.0.0""" & vbCrLf & _
         "   processorArchitecture=""X86""" & vbCrLf & _
         "   name=""" & App.EXEName & """" & vbCrLf & _
         "   type=""win32""" & vbCrLf & _
         "/>" & vbCrLf & _
         "<description>XP-Look</description>" & vbCrLf & _
         "<dependency>" & vbCrLf & _
         "   <dependentAssembly>" & vbCrLf & _
         "      <assemblyIdentity" & vbCrLf & _
         "         type=""win32""" & vbCrLf & _
         "         name=""Microsoft.Windows.Common-Controls""" & vbCrLf & _
         "         version=""6.0.0.0""" & vbCrLf & _
         "         processorArchitecture=""X86""" & vbCrLf & _
         "         publicKeyToken=""6595b64144ccf1df""" & vbCrLf & _
         "         language=""*""" & vbCrLf & _
         "      />" & vbCrLf & _
         "   </dependentAssembly>" & vbCrLf & _
         "</dependency>" & vbCrLf & _
         "</assembly>"
      
      On Error Resume Next
      
      '/* Create manifest file if it is missing
      strEXEName = App.Path & "\" & App.EXEName & ".exe.Manifest"
      If LenB(Dir$(strEXEName)) = 0 Then
           lngFN = FreeFile
           Open strEXEName For Output As lngFN
           Print #lngFN, strXPLookXML
           Close lngFN
      End If
   
      '/* Link XP themes to application
      Call InitCommonControls
      
      On Error GoTo 0
   
   End If 'gblnOsIsXp

End Sub

Public Sub EndApp()
  
  '/* Call from the closing Form's Form_Unload event
  '/* Example:
  '/*   Call EndApp
  '/*   Set FormName = Nothing
  '/*   '/* End Program

  Dim Frm As Form
  Const SEM_NOGPFAULTERRORBOX As Long = &H2&
  
    On Error Resume Next
    
    '/* Close all open Forms
    For Each Frm In Forms
        Unload Frm
        Set Frm = Nothing
    Next Frm

    '/* Some versions of ComCtl32.DLL version 6.0 cause a crash at shutdown
    '/* when you enable XP Visual Styles in an application that has a VB User Control.
    '/* This code instructs Windows to NOT display the UAE message box that
    '/* invites you to send Microsoft information about the problem.
    If CBool(VB.App.LogMode()) Then '/* Not running in IDE
        Call SetErrorMode(SEM_NOGPFAULTERRORBOX)
    End If
             
End Sub

