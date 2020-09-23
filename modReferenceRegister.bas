Attribute VB_Name = "modReferenceRegister"
' =========================================================================
' EZe Component Register
' Copyright © 2001 Tony Wilson (tonyscomp@europe.com)
'
' Registers your controls by simply double click them
' =========================================================================
Private Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function FindFirstFileA Lib "kernel32" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Global CpyToSystem As Single
Global SystemDir As String

Sub Main()
    Dim SLength As Long         ' receives length of the string returned
    Dim sysfilename As String   ' new name of file copied to system dir

    SystemDir = Space(255)                          ' initialize buffer to receive the string
    SLength = GetSystemDirectory(SystemDir, 255)    ' read the path of the Windows directory
    SystemDir = Left(SystemDir, SLength)            ' extract the returned string from the buffer

    'Register myself so that I can be used by simply right clicking the icon
    'or as the default for opening the component.
    SelfRegister
    
    'Get the saved setting on whether to copy to the system
    CpyToSystem = GetSetting("EZe Register", "Load Settings", "Copy To System Dir", 0)
      
    If Command = "" Then
        'The program was loaded manually - show it
        frmMain.Show
    Else
        'Register the control
        RegisterIt Command, CpyToSystem
    End If
    
End Sub

Sub SelfRegister()

    bSetRegValue HKEY_CLASSES_ROOT, "ocxfile\shell\Register", "", "Register by EZe Component Register"
    bSetRegValue HKEY_CLASSES_ROOT, "ocxfile\shell\Register\command", "", """" & App.Path & "\" & App.EXEName & ".exe" & """" & " ""%1" & """"
    
    bSetRegValue HKEY_CLASSES_ROOT, "dllfile\shell\Register", "", "Register by EZe Component Register"
    bSetRegValue HKEY_CLASSES_ROOT, "dllfile\shell\Register\command", "", """" & App.Path & "\" & App.EXEName & ".exe" & """" & " ""%1" & """"
    
    bSetRegValue HKEY_CLASSES_ROOT, "olbfile\shell\Register", "", "Register by EZe Component Register"
    bSetRegValue HKEY_CLASSES_ROOT, "olbfile\shell\Register\command", "", """" & App.Path & "\" & App.EXEName & ".exe" & """" & " ""%1" & """"
    
    bSetRegValue HKEY_CLASSES_ROOT, "tlbfile\shell\Register", "", "Register by EZe Component Register"
    bSetRegValue HKEY_CLASSES_ROOT, "tlbxfile\shell\Register\command", "", """" & App.Path & "\" & App.EXEName & ".exe" & """" & " ""%1" & """"
    
    bSetRegValue HKEY_CLASSES_ROOT, "exefile\shell\Register", "", "Register by EZe Component Register"
    bSetRegValue HKEY_CLASSES_ROOT, "exefile\shell\Register\command", "", """" & App.Path & "\" & App.EXEName & ".exe" & """" & " ""%1" & """"

End Sub

Public Function RegisterIt(ByVal sFile As String, ByVal sCopyToSystem As Single) As Boolean
    
    Dim sysfilename As String
    Dim DidIt As Boolean
    'trim the string of commas
    If Left(sFile, 1) = """" Then sFile = Mid(sFile, 2, Len(sFile) - 2)

    'Should we copy to the system directory?
    If sCopyToSystem = 1 Then

        'Get new name if the file is copied
        sysfilename = FixPath(SystemDir) & FileName(sFile)
        
        'Check to see if the file already exists
        If FileExists(sysfilename) Then
        
            'File exists, ask for permission to replace
            If MsgBox("That file already exists, would you like to replace it?", vbYesNo + vbQuestion, "Confirm") = vbYes Then
                FileCopy sFile, sysfilename
            Else
                Shell "regsvr32 " & sFile & " /s"
                DidIt = True
            End If
        Else
            FileCopy sFile, sysfilename
            
            'Register the copied file
            sFile = sysfilename
            
            Shell "regsvr32 " & sFile & " /s"
            DidIt = True
        End If
            
    Else
        'Register the file
        Shell "regsvr32 " & sFile & " /s"
        DidIt = True
    End If

    If DidIt Then
        RegisterIt = True
        MsgBox "Your component was successfully registered." & vbCrLf & "Your registration has been saved to the log file", vbInformation + vbOKOnly, "Success"
        WriteToLog sFile
    
    End If
    Exit Function

Err:

    MsgBox "There was an error" & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    RegisterIt = False


End Function

Private Sub WriteToLog(sFile As String)
    Open FixPath(App.Path) & "RegLog.txt" For Append As #1
    
        Print #1, sFile & " - " & Date & " - " & Time
    
    Close #1


End Sub

Public Function FileExists(ByVal sFile As String) As Boolean
    Dim R As Long
    Dim uFIND_DATA As WIN32_FIND_DATA
    
    R = FindFirstFileA(sFile, uFIND_DATA)
    
    If R = INVALID_HANDLE_VALUE Then
        FileExists = False
    Else
        FileExists = True
        Call FindClose(R)
    End If

End Function

Public Function FileName(ByVal sFullPath As String) As String
    Dim I As Integer
    
    I = InStrRev(sFullPath, "\")

    If I > 0 Then
        FileName = Mid$(sFullPath, I + 1)
    Else
        FileName = sFullPath
    End If
End Function

Public Function FixPath(ByVal sPath As String) As String

    If Right(sPath, 1) = "]" Then Exit Function
    
    If Right(sPath, 1) = "\" Then
        FixPath = sPath
    Else
        FixPath = sPath & "\"
    End If

End Function
