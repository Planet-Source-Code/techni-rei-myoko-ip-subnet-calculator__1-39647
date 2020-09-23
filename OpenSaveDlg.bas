Attribute VB_Name = "OpenSaveDlg"
Public HTMLFileName As String ' Global Variable - avaliable to all forms


'-------------------------------------------------------'
' Some Of This Code Was Taken From PSC                  '
' Thanks To: Brand-X Software For Original Code         '
' Edited By Arvinder Sehmi                              '
'-------------------------------------------------------'
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type
Public SaveFileDialog As OPENFILENAME
Public OpenFileDialog As OPENFILENAME
Private rv As Long
Private sv As Long
Public Function fileexists(filename As String) As Boolean
If Dir(filename) = "" Then fileexists = False Else fileexists = True
End Function
Public Function Open_File(hWnd As Long) As String
   rv& = GetOpenFileName(OpenFileDialog)
   
   If (rv&) Then
      Open_File = Trim$(OpenFileDialog.lpstrFile)
      Open_File = left(Open_File, Len(Open_File) - 1)
   Else
      Open_File = ""
   End If
End Function
Public Function Save_File(hWnd As Long) As String
   sv& = GetSaveFileName(SaveFileDialog)
   If (sv&) Then
      Save_File = Trim$(SaveFileDialog.lpstrFile)
      Save_File = left(Save_File, Len(Save_File) - 1)
      If fileexists(Save_File) Then
        Dim temp As Integer
        temp = MsgBox("This file already exists" & vbNewLine & Save_File & vbNewLine & "Do you want to overwrite it?", vbYesNo, "File already exists")
        If temp = vbNo Then Save_File = ""
      End If
   Else
      Save_File = ""
   End If
End Function

Public Sub InitSaveDlg()
  With SaveFileDialog
     .lStructSize = Len(SaveFileDialog)
     .hwndOwner = hWnd&
     .hInstance = App.hInstance
     .lpstrFilter = "Text File, Or Html" + Chr$(0) + "*.txt;*.htm*"
     .lpstrFile = Space$(254)
     .nMaxFile = 255
     .lpstrFileTitle = Space$(254)
     .nMaxFileTitle = 255
     .lpstrInitialDir = CurDir
     .lpstrTitle = "Save IP details"
     .flags = 0
  End With
End Sub
Public Sub InitOpenDlg()
   With OpenFileDialog
     .lStructSize = Len(OpenFileDialog)
     .hwndOwner = hWnd&
     .hInstance = App.hInstance
     .lpstrFilter = "Image Formats" + Chr$(0) + "*.bmp;*.jpg;*.gif;*.pcx;*.wmf;*.emf;*.dib"
     .lpstrFile = Space$(254)
     .nMaxFile = 255
     .lpstrFileTitle = Space$(254)
     .nMaxFileTitle = 255
     .lpstrInitialDir = CurDir
     .lpstrTitle = "Load Image..."
     .flags = 0
   End With
End Sub
Public Sub InitDlgs()
 Call InitSaveDlg
 Call InitOpenDlg
End Sub
