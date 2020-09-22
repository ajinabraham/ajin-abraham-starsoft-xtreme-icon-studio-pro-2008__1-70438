VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3945
   ClientLeft      =   2340
   ClientTop       =   1545
   ClientWidth     =   5745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2722.909
   ScaleMode       =   0  'User
   ScaleWidth      =   5394.852
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   465
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   1620
   End
   Begin VB.Image IconImg 
      Height          =   375
      Left            =   0
      Picture         =   "frmAbout.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "Menu"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Closeb 
      Height          =   375
      Left            =   5280
      Picture         =   "frmAbout.frx":08CA
      Stretch         =   -1  'True
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image ic 
      Height          =   1095
      Left            =   240
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":1AD5
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   945
      Left            =   285
      TabIndex        =   3
      Top             =   2625
      Width           =   3870
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   126.772
      X2              =   5337.57
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Starsoft Xtreme "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   1965
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Icon Studio Pro 2008"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   3165
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5337.57
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label WindowCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image Maximize 
      Height          =   375
      Left            =   4080
      Picture         =   "frmAbout.frx":1B87
      Stretch         =   -1  'True
      ToolTipText     =   "Maximize"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Minimize 
      Height          =   375
      Left            =   4440
      Picture         =   "frmAbout.frx":2A45
      Stretch         =   -1  'True
      ToolTipText     =   "Minimize"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Restore 
      Height          =   375
      Left            =   4920
      Picture         =   "frmAbout.frx":3636
      Stretch         =   -1  'True
      ToolTipText     =   "Restore"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image TitleBar 
      Height          =   405
      Left            =   0
      Picture         =   "frmAbout.frx":461E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6435
   End
   Begin VB.Image bg 
      Height          =   3615
      Left            =   -120
      Stretch         =   -1  'True
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer, n As Integer
Dim GapX As Integer, GapY As Integer





Private Sub cmdOK_Click()
 
End Sub

Private Sub cmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
SetupTitlebar Me
SetTrans 30, Me.hwnd
bg.Picture = Form1.Image1.Picture
ic.Picture = Form1.Icon

End Sub

Private Sub Closeb_Click()
End
End Sub

Private Sub Form_Resize()
SetupTitlebar Me

End Sub

Private Sub Closeb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MVCloseb Me
End Sub

Private Sub Maximize_Click()
Me.WindowState = vbMaximized
Maximize.Visible = False
End Sub

Private Sub maximize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MVMaximize Me
End Sub

Private Sub Minimize_Click()
Form1.WindowState = vbMinimized
End Sub

Private Sub Restore_Click()
Me.WindowState = vbNormal
Maximize.Visible = True
End Sub

Private Sub restore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MVRestore Me
End Sub

Private Sub minimize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MVMinimize Me
End Sub





Private Sub TitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
GapX = X
GapY = Y
End If
End Sub

Private Sub TitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then _
TitleBar.Parent.Move TitleBar.Parent.Left + X - GapX, TitleBar.Parent.Top + Y - GapY
End Sub

Private Sub windowcaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
GapX = X
GapY = Y
End If
End Sub

Private Sub windowcaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then _
TitleBar.Parent.Move TitleBar.Parent.Left + X - GapX, TitleBar.Parent.Top + Y - GapY
End Sub



