VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About EZe Component Register"
   ClientHeight    =   3495
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4365
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2412.311
   ScaleMode       =   0  'User
   ScaleWidth      =   4098.96
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2880
      TabIndex        =   7
      Top             =   3000
      Width           =   1260
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00CC600F&
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   810
      TabIndex        =   0
      Top             =   120
      Width           =   863
      Begin VB.Image Image2 
         Height          =   720
         Left            =   50
         Picture         =   "frmAbout.frx":0000
         Top             =   50
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FAE4D3&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   6015
      TabIndex        =   1
      Top             =   -120
      Width           =   6015
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version: 1.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   6
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Component Register"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1080
         TabIndex        =   2
         Top             =   360
         Width           =   2700
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please feel free to distribute the original compiled version of this program along with your Visual Basic programs."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1080
      TabIndex        =   8
      Top             =   1800
      Width           =   3255
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Index           =   0
      X1              =   112.686
      X2              =   3944.016
      Y1              =   1905.001
      Y2              =   1905.001
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2001 Tony Wilson"
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
      Left            =   120
      MouseIcon       =   "frmAbout.frx":1B42
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Tag             =   "http://ezevb.cjb.net/copyright.htm"
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "tonyscomp@europe.com"
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
      Left            =   120
      MouseIcon       =   "frmAbout.frx":1E4C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Tag             =   "tonyscomp@europe.com"
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":2156
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   3255
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   5634.309
      Y1              =   424.484
      Y2              =   424.484
   End
End
Attribute VB_Name = "frmAbout"
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
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub lblCopyright_Click()
    ShellExecute hwnd, "open", lblCopyright.Tag, vbNullString, vbNullString, 5
End Sub

Private Sub lblEmail_Click()
    ShellExecute hwnd, "open", "mailto:" & lblEmail.Tag, vbNullString, vbNullString, 5
End Sub
