VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@==========================================================================
' http://vbplanet.ath.cx
'
' CDialog Class is used for common dialog API Calls, Feel free to use it
' in your commercial programs. Please do not distribute it in uncompiled form
' or display this code on any webpages. Thank you
'============================================================================
Option Compare Text
Option Explicit
' ** API CONSTANTS **
' Constants for GetSaveFileName API
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
' Constants used by GetOpenFileName
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_CREATEPROMPT = &H2000 ' prompt to overwrite
Private Const OFN_NODEREFERENCELINKS = &H100000
' ** API TYPES **
' OpenFileName Type is used by GetSaveFileName/OpenSaveFilename API
Private Type OPENFILENAME
    lStructSize As Long ' size of type/structure
    hwndOwner As Long ' Handle of owner window
    hInstance As Long
    lpstrFilter As String ' Filters used to select files
    lpstrCustomFilter As String
    nMaxCustomFilter As Long
    nFilterIndex As Long ' index of Filter to start with
    lpstrFile As String ' Holds filepath and name
    nMaxFile As Long    ' Maximum Filepath and name length
    lpstrFileTitle As String ' Filename
    nMaxFileTitle As Long ' Max Length of filename
    lpstrInitialDir As String ' Starting Directory
    lpstrTitle As String ' Title of window
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
' For Color Dialog
Private Type ChooseColorTYPE
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        Flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

' ** API CALLS **
' Brings up common dialog prompting user to save a file
Private Declare Function GetSaveFileName Lib "comdlg32.dll" _
       Alias "GetSaveFileNameA" (lpofn As OPENFILENAME) As Long
' We Declare GetOpenFileName
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
 "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
' color dialog
Private Declare Function ChooseColorDLG Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColorTYPE) As Long
' CopyMemory used for color dialog
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'@================================================================
' GetSaveName:
'   Uses API to prompt user to save a file
' Returns: A string containing filepath and name
'=================================================================
Function GetSaveName(Optional ByVal WindowTitle As String = "Save File", _
         Optional ByVal FilterStr As String = "Windows Bitmap" + vbNullChar + "*.bmp*") As String
    Dim DlgInfo As OPENFILENAME  ' Open file name type
    Dim RET As Long           ' Used for API return Values
    ' Set the Dlg windows properties
    With DlgInfo
        ' set the Size of structure/type
        .lStructSize = Len(DlgInfo)
        ' Set parent window handle
        .hwndOwner = 0
        ' File filters
        .lpstrFilter = FilterStr
        ' Set Default filter to be first one in our set
        .nFilterIndex = 1
        ' Make room for filepath string and set maxlen
        .lpstrFile = Space(512) & vbNullChar
        .nMaxFile = Len(.lpstrFile)
        ' Make room for filename and set maxlen
        .lpstrFileTitle = Space(256) & vbNullChar
        .nMaxFileTitle = Len(.lpstrFileTitle)
        ' Set Initial Directory
        .lpstrInitialDir = CurDir & vbNullChar
        ' Set title of window
        .lpstrTitle = WindowTitle & vbNullChar
        ' Prompt if user wants to overwrite
        .Flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
        ' Unused options
        .nMaxCustomFilter = 0
        .nFileOffset = 0
        .nFileExtension = 0
        .lCustData = 0
        .lpfnHook = 0
        .hInstance = 0
    End With
    ' Display the dialog box.
    RET = GetSaveFileName(DlgInfo)
    ' if return value then we got a filename
    If Not RET = 0 Then
        ' Remove null
        GetSaveName = Left(DlgInfo.lpstrFile, InStr(DlgInfo.lpstrFile, vbNullChar) - 1)
    Else ' Nothing selected
        GetSaveName = ""
    End If
End Function

'@=========================================================================
' GetOpenName:
'    Prompts user for a file to open.
' Parameters:
'
' Returns: Filename and path user selects
'==========================================================================
Function GetOpenName(Optional ByVal WindowTitle As String = "Load File", _
                     Optional ByVal Filters As String = "All Files" + vbNullChar + "*.*", _
                     Optional ByVal DefaultFileName As String = "")
 Dim RET As Long ' Return values
 Dim DlgInfo As OPENFILENAME ' Dim our type for Dialog setup
 ' Setup the way the dialog looks and acts
 With DlgInfo
      ' set size of structure/type
      .lStructSize = Len(DlgInfo)
      ' Set Parent window
      .hwndOwner = 0
      ' What files to show
      .lpstrFilter = Filters
      ' Set Filter index
      .nFilterIndex = 1
      ' Set the initial filename/path leaving space left over incase user selects larger
      .lpstrFile = DefaultFileName & Space$(1024) & vbNullChar & vbNullChar
      ' set max len of filename/path
      .nMaxFile = Len(.lpstrFile)
      ' whether to apply an extension if non is supplie
      .lpstrDefExt = vbNullChar & vbNullChar
      ' Space for filename
      .lpstrFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
      .nMaxFileTitle = Len(.lpstrFileTitle)
      'starting folder, double-null terminated
      .lpstrInitialDir = CurDir + vbNullChar
      ' Set the window title
      .lpstrTitle = WindowTitle
      ' Set open Flags
      .Flags = OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_NODEREFERENCELINKS
 End With
 ' Get Filename from dialog
  GetOpenName = GetOpenFileName(DlgInfo)
  ' remove unneeded chars
  GetOpenName = Left(DlgInfo.lpstrFile, InStr(DlgInfo.lpstrFile, vbNullChar) - 1)
End Function
'@============================================================================
' PickColor: Lets the user pick a color and returns that color
' NOTE: if no color is picked, -1 is returned
'=============================================================================
Function PickColor() As Long
    Dim TChooseColor As ChooseColorTYPE
    Dim CustomColors(15) As Long ' Holds custom colors
    Dim RET As Long ' Return value of API call
    Dim A As Long ' General purpose
    ' Load Custom colors
    For A = 0 To 15
     CustomColors(A) = CLng(GetSetting("Temp", "CustomColors", CStr(A), "0"))
    Next A
    ' set type structure.
    With TChooseColor
        ' Set the handle for the owner of the window.
        .hwndOwner = 0
        ' Pass the custom colors array after converting
        .lpCustColors = VarPtr(CustomColors(0))
        ' it to Unicode using the StrConv function.
        .Flags = &H101 Or &H2
        ' Set the size of the type structure.
        .lStructSize = Len(TChooseColor)
    End With
    ' display dialog
    RET = ChooseColorDLG(TChooseColor)
    ' If user picked color we get it otherwise -1 is returned
    If RET = 0 Then
      PickColor = -1
     Else
      PickColor = TChooseColor.rgbResult
     End If
     ' update custom Colors Long is 32 bytes
     CopyMemory CustomColors(0), ByVal (TChooseColor.lCustData), 32 * 16
     ' save custom Colors
     For A = 0 To 15
      SaveSetting "Temp", "CustomColors", CStr(A), CustomColors(A)
     Next A
End Function
         
