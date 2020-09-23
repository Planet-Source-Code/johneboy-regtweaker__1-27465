VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About regTweaker"
   ClientHeight    =   3150
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5535
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmAbout.frx":030A
   ScaleHeight     =   2174.186
   ScaleMode       =   0  'User
   ScaleWidth      =   5197.651
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   120
      Picture         =   "frmAbout.frx":0614
      ScaleHeight     =   1395
      ScaleWidth      =   5235
      TabIndex        =   2
      Top             =   120
      Width           =   5295
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      DisabledPicture =   "frmAbout.frx":4B56
      Height          =   345
      Left            =   2040
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   1260
   End
   Begin VB.Label Label2 
      Caption         =   "Questions, Comments, Suggestions? "
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Johneboy@mindspring.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3120
      MouseIcon       =   "frmAbout.frx":4F98
      MousePointer    =   4  'Icon
      TabIndex        =   3
      ToolTipText     =   "Send me an E-Mail"
      Top             =   2160
      UseMnemonic     =   0   'False
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   6648.486
      Y1              =   1739.349
      Y2              =   1739.349
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      Caption         =   "A great little ultility for customizing Windows 98® without having to mess with the registry. "
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   6648.486
      Y1              =   1739.349
      Y2              =   1739.349
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
      Private Declare Function ShellExecute Lib _
     "shell32.dll" Alias "ShellExecuteA" _
     (ByVal hwnd As Long, ByVal lpOperation _
     As String, ByVal lpFile As String, ByVal _
     lpParameters As String, ByVal lpDirectory _
     As String, ByVal nShowCmd As Long) As Long
Private Sub cmdOK_Click()


Unload Me
End Sub


Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Left = 2929
cmdOK.BackColor = &H80FFFF

End Sub

Private Sub Command1_Click()
Form1.Show 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.FontBold = False
Me.MousePointer = 0
Label1.Left = 2929
cmdOK.BackColor = &H8000000F

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me


End Sub

Private Sub Label1_Click()
ShellExecute 0&, vbNullString, "mailto:johneboy@mindspring.com?Subject=About regTweaker...", vbNullString, _
      vbNullString, SW_SHOWNORMAL
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label1.FontBold = True
   Label1.Left = 2800
   

    Me.MousePointer = 99


    Label1.MouseIcon = LoadPicture("c:\Windows\system\hand.cur")


End Sub


