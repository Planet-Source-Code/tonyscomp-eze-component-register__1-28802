VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register your components easily!"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FAE4D3&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   735
      ScaleWidth      =   6255
      TabIndex        =   8
      Top             =   0
      Width           =   6255
      Begin VB.Label lblVote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vote"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5865
         MouseIcon       =   "frmMain.frx":68E2
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   240
         Width           =   330
      End
      Begin VB.Label lblAbout 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5760
         MouseIcon       =   "frmMain.frx":6BEC
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   0
         Width           =   435
      End
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      Height          =   255
      Left            =   5685
      TabIndex        =   1
      Top             =   960
      Width           =   325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   3600
      TabIndex        =   2
      Top             =   3480
      Width           =   1150
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   330
      Left            =   4920
      TabIndex        =   3
      Top             =   3480
      Width           =   1150
   End
   Begin VB.CheckBox chkCopy 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copy to system folder before registering"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox txtInfo 
      Height          =   1575
      Left            =   2760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   6240
      Y1              =   730
      Y2              =   730
   End
   Begin VB.Image Image1 
      Height          =   3270
      Left            =   0
      Picture         =   "frmMain.frx":6EF6
      Top             =   720
      Width           =   2100
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Info:"
      Height          =   195
      Left            =   2280
      TabIndex        =   7
      Top             =   1320
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Path:"
      Height          =   195
      Left            =   2280
      TabIndex        =   6
      Top             =   960
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================================
' EZe Component Register
' Copyright © 2001 Tony Wilson (tonyscomp@europe.com)
'
' Registers your controls by simply double click them
' =========================================================================



Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, lpString2 As Any) As Long

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If Not txtFile.Text = "" And FileExists(txtFile.Text) Then
        RegisterIt txtFile.Text, CpyToSystem
    Else
        MsgBox "Please specify a correct file address", vbExclamation + vbOKOnly, "Error"
    End If
End Sub

Private Sub cmdOpen_Click()
    Dim die As clsFile
    Set die = New clsFile

    die.DialogTitle = "Select Component"
    die.Filter = "ActiveX Control (*.ocx)|*.ocx|ActiveX Server (*.dll)|*.dll|Type Libraries (*.tlb)|*.tlb|Object Library (*.olb)|*.olb|Application (*.exe)|*.exe"
    die.ShowOpen
    txtFile.Text = die.FileName
    
    ShowInfo txtFile.Text

End Sub

'Works only On 32-bit Files
Private Sub ShowInfo(sFile As String)
    Dim sVerInfo(7) As String
    Dim sInfo As String
    Dim sValue As String
    Dim sData() As Byte
    Dim lSize As Long
    Dim lPointer As Long
    Dim iIndex As Integer

    txtInfo = ""
    'get the Length of the FileVersion Information
    lSize = GetFileVersionInfoSize(sFile, ByVal 0&)
    'Create a Buffer to hold the Version Info
    ReDim sData(lSize)
    'get the Version Info
    If GetFileVersionInfo(sFile, 0&, lSize, sData(0)) Then
        'Extract the Details of the Version Info
        sVerInfo(0) = "CompanyName"
        sVerInfo(1) = "FileDescription"
        sVerInfo(2) = "FileVersion"
        sVerInfo(3) = "InternalName"
        sVerInfo(4) = "LegalCopyright"
        sVerInfo(5) = "OriginalFileName"
        sVerInfo(6) = "ProductName"
        sVerInfo(7) = "ProductVersion"
        For iIndex = 0 To 7
            sInfo = "\StringFileInfo\040904E4\" & sVerInfo(iIndex)
            If VerQueryValue(sData(0), sInfo, lPointer, lSize) Then
                sValue = Space(lSize)
                lstrcpy ByVal sValue, ByVal lPointer
                txtInfo.SelText = sVerInfo(iIndex) & ": " & sValue
                txtInfo.SelText = vbCrLf
            End If
         Next
    End If

    If txtInfo.Text = "" Then
        txtInfo.Text = "No Information Available"
    End If

End Sub

Private Sub Form_Load()
    
    If GetSetting("EZe Register", "Vote Settings", "Number Times Voted", 0) = 0 Then
    
        If MsgBox("Thank You For Downloading My Project" & vbCrLf & "I hope you find this program useful as I have put a lot of work into it.  If you do like my program I would really be greatful if you voted for me." & vbCrLf & "Would you like to vote now?", vbQuestion + vbYesNo, "EZe Component Register") = vbYes Then
            Call Shell("Start.exe " & "http://www.planetsourcecode.com/xq/ASP/txtCodeId.28802/lngWId.1/qx/vb/scripts/ShowCode.htm", 0)
            SaveSetting "EZe Register", "Vote Settings", "Number Times Voted", 1
            MsgBox "Thank you, this box will not be displayed the next time you run this program"
        End If

    End If
    
    chkCopy.ToolTipText = SystemDir
    chkCopy.Value = CpyToSystem
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "EZe Register", "Load Settings", "Copy To System Dir", chkCopy.Value
End Sub

Private Sub lblAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub lblVote_Click()

    Call Shell("Start.exe " & "http://www.planetsourcecode.com/xq/ASP/txtCodeId.28802/lngWId.1/qx/vb/scripts/ShowCode.htm", 0)
    SaveSetting "EZe Register", "Vote Settings", "Number Times Voted", 1
    MsgBox "Thank you for your vote!"

End Sub

