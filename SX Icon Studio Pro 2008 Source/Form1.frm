VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   5115
   ClientLeft      =   5055
   ClientTop       =   4605
   ClientWidth     =   7215
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "About "
      Height          =   855
      Left            =   6000
      Picture         =   "Form1.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "About Icon Studio Pro"
      Top             =   480
      Width           =   1215
   End
   Begin MSComctlLib.Toolbar tbSmall 
      Height          =   630
      Left            =   5400
      TabIndex        =   7
      Top             =   1680
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.CheckBox msk 
      BackColor       =   &H000080FF&
      Caption         =   "Use Back Color"
      Height          =   255
      Left            =   840
      TabIndex        =   20
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Icons"
      Height          =   255
      Left            =   5280
      TabIndex        =   19
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picSmall 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5520
      ScaleHeight     =   240
      ScaleMode       =   0  'User
      ScaleWidth      =   240
      TabIndex        =   12
      Top             =   4320
      Width           =   240
   End
   Begin VB.PictureBox picLarge 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4560
      ScaleHeight     =   495
      ScaleMode       =   0  'User
      ScaleWidth      =   495
      TabIndex        =   11
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000080FF&
      Caption         =   "Back"
      Height          =   615
      Left            =   3240
      Picture         =   "Form1.frx":2855C
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Save as Icon or Cursor"
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H000080FF&
      Caption         =   "Next"
      Height          =   615
      Left            =   3240
      Picture         =   "Form1.frx":28C16
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Save as Icon or Cursor"
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton S 
      BackColor       =   &H000080FF&
      Caption         =   "Save Extracted Icon"
      Height          =   855
      Left            =   4800
      Picture         =   "Form1.frx":292D0
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Save as Icon "
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Open_cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "Open Image File"
      Height          =   855
      Left            =   0
      Picture         =   "Form1.frx":2998A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Open an Image File"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      BackColor       =   &H000080FF&
      Caption         =   "Open Exe& DLLs"
      Height          =   855
      Left            =   3600
      Picture         =   "Form1.frx":2A044
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Open exe or dll"
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Save as Cursor"
      Height          =   855
      Left            =   2400
      Picture         =   "Form1.frx":2A6FE
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Save as Cursor"
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Save_cmd 
      BackColor       =   &H000080FF&
      Caption         =   "Save as Icon"
      Height          =   855
      Left            =   1200
      Picture         =   "Form1.frx":2B0C4
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Save as Icon "
      Top             =   480
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CDBoxClr 
      Left            =   120
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CDBox2 
      Left            =   120
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdbox 
      Left            =   120
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All Files(*.*)|*.*"
   End
   Begin MSComctlLib.ImageList imglst 
      Left            =   120
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   320
      ImageHeight     =   320
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   720
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbLarge 
      Height          =   630
      Left            =   4560
      TabIndex        =   6
      Top             =   1680
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   480
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FilterIndex     =   1
   End
   Begin MSComctlLib.ImageList imgLarge 
      Left            =   480
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgSmall 
      Left            =   360
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Main Icon      Large           Small"
      Height          =   255
      Left            =   3600
      TabIndex        =   21
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblIcon 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5760
      TabIndex        =   18
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Icon Index:"
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblIcons 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5640
      TabIndex        =   16
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Number of  Icons:"
      Height          =   195
      Left            =   4200
      TabIndex        =   15
      Top             =   3120
      Width           =   1260
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Small"
      Height          =   255
      Left            =   5400
      TabIndex        =   14
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Large"
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   3120
      TabIndex        =   8
      Top             =   2400
      Width           =   3495
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000D&
      Height          =   1095
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   5
      X1              =   3000
      X2              =   3000
      Y1              =   1320
      Y2              =   5280
   End
   Begin VB.Image IconView 
      Height          =   1335
      Left            =   720
      MouseIcon       =   "Form1.frx":2B77E
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":2B8D0
      Stretch         =   -1  'True
      ToolTipText     =   "Preview"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Image IconImg 
      Height          =   375
      Left            =   0
      Picture         =   "Form1.frx":46326
      Stretch         =   -1  'True
      ToolTipText     =   "Menu"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Restore 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6480
      Picture         =   "Form1.frx":46BF0
      Stretch         =   -1  'True
      ToolTipText     =   "Restore"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Minimize 
      Height          =   375
      Left            =   5760
      Picture         =   "Form1.frx":47BD8
      Stretch         =   -1  'True
      ToolTipText     =   "Minimize"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Closeb 
      Height          =   375
      Left            =   6840
      Picture         =   "Form1.frx":487C9
      Stretch         =   -1  'True
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Maximize 
      Height          =   375
      Left            =   6120
      Picture         =   "Form1.frx":499D4
      Stretch         =   -1  'True
      ToolTipText     =   "Maximize"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label WindowCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "     Starsoft Xtreme Icon Studio Proffesional"
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
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000040C0&
      BorderWidth     =   2
      Height          =   2295
      Left            =   600
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Image TitleBar 
      Height          =   405
      Left            =   0
      Picture         =   "Form1.frx":4A892
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8235
   End
   Begin VB.Image Image1 
      Height          =   4935
      Left            =   0
      Picture         =   "Form1.frx":4D4C5
      Stretch         =   -1  'True
      Top             =   360
      Width           =   7275
   End
   Begin VB.Menu MnuControl 
      Caption         =   "ControlBox"
      Visible         =   0   'False
      Begin VB.Menu MnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuMinimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnu_removeTrans 
         Caption         =   "Remove Transparancy"
      End
      Begin VB.Menu Mnu_Transp 
         Caption         =   "Make Transparent"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim glLargeIcons() As Long
Dim glSmallIcons() As Long
Dim lIndex         As Long
Dim lIcons         As Long
Dim sExeName       As String

Const LARGE_ICON As Integer = 32
Const SMALL_ICON As Integer = 16
Const DI_NORMAL = 3
Private Declare Function DrawIconEx Lib "user32" _
    (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, _
    ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, _
    ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, _
    ByVal diFlags As Long) As Long

Private Declare Function ExtractIconEx Lib "shell32.dll" _
    Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, _
    phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long

Dim GapX As Integer, GapY As Integer
Dim fso As New FileSystemObject
Dim widIco As Integer
Dim HiIco As Integer

Sub resetIcon()
If IconImg.Width <> widIco Or IconImg.Height <> HiIco Then
IconImg.Width = widIco
IconImg.Height = HiIco
End If
End Sub

Private Sub cc_Click()

End Sub

Private Sub cmdBack_Click()
'
' Get the previous icon.
'
If lIndex > 0 Then
    lIndex = lIndex - 1
    Call pGetIcon
End If

End Sub

Private Sub cmdBrowse_Click()
Dim btn    As Button
Dim imgObj As ListImage
'
' Initialize labels. Clear the picture boxes.
'
lIcons = 0
lIndex = 0
lblIcons = 0
lblIcon = 0
lblFile = ""
picSmall.Picture = LoadPicture("")
picLarge.Picture = LoadPicture("")
'
' Remove all toolbar buttons and the
' unbind the ImageList controls.
'
tbLarge.Buttons.Clear
tbLarge.ImageList = Nothing
tbSmall.Buttons.Clear
tbSmall.ImageList = Nothing
'
' Remove all images from the ImageList controls
' and set their size properties.
'
With imgLarge
    .ListImages.Clear
    .ImageHeight = LARGE_ICON
    .ImageWidth = LARGE_ICON
End With

With imgSmall
    .ListImages.Clear
    .ImageHeight = SMALL_ICON
    .ImageWidth = SMALL_ICON
End With
'
' Display the File Open dialog.
' Filter out all files except exe's and dll's.
'
cdlOpen.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNHideReadOnly
cdlOpen.FileName = ""
cdlOpen.Filter = "Executable Files (*.exe) | *.exe|Application Extension (*.dll) | *.dll"
On Error GoTo CancelButton
cdlOpen.Action = 1
sExeName = cdlOpen.FileName
lblFile = sExeName
'
' Get the total number of Icons in the file.
'
lIcons = ExtractIconEx(sExeName, -1, 0, 0, 0)
'
' Enable various controls.
'
lblIcons = lIcons
cmdBack.Enabled = (lIcons > 1)
cmdNext.Enabled = (lIcons > 1)
lblIcons.Enabled = True
lblIcon.Enabled = True
picSmall.Enabled = True
picLarge.Enabled = True
Label1.Enabled = True
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Frame2.Enabled = True
'
' Dimension the arrays to the number of icons.
' Get the icons' handles.
'
ReDim glLargeIcons(lIcons)
ReDim glSmallIcons(lIcons)
Call pGetIcon
'
' Add the Large icon to the Large ImageList control.
' Bind the large ImageList to the large ToolBar.
' Add a button to the toolbar and populate its ToolTip text.
'
' Note: The "Key" fields of both the ImageList and ToolBar
'       control are set to the same value.  This is what
'       binds a particular image in the ImageList to a
'       given button on the ToolBar control.
'
'           Syntax is:    ...Add(Index, Key, Image)
Set imgObj = imgLarge.ListImages.Add(1, sExeName, picLarge.Image)

With tbLarge
    .ImageList = imgLarge
    ' Syntax is:    ...Add(Index, Key, Caption, Style, Image)
    Set btn = .Buttons.Add(.Buttons.Count + 1, sExeName, , , sExeName)
    .Buttons(1).ToolTipText = sExeName
End With
'
' Repeat for the small icon.
'
Set imgObj = imgSmall.ListImages.Add(1, sExeName, picSmall.Image)
With tbSmall
    .ImageList = imgSmall
    Set btn = .Buttons.Add(.Buttons.Count + 1, sExeName, , , sExeName)
    .Buttons(1).ToolTipText = sExeName
End With

CancelButton:
    'We end up here when hitting Cancel on the Open File dialog.

End Sub

Private Sub cmdNext_Click()
'
' Get the next icon.
'
If lIndex < lIcons - 1 Then
    lIndex = lIndex + 1
    Call pGetIcon
End If

End Sub

Private Sub cmdSave_Click()
End Sub

Private Sub Command1_Click()
On Error GoTo er
With CDBox2
   .Filter = "Cursor File(.cur)|*.cur"
   .ShowSave
End With
If fso.FileExists(CDBox2.FileName) = True Then
 If MsgBox("File '" & CDBox2.FileTitle & "' already exists!" & vbNewLine & "Do you want to overwrite?", vbYesNo + vbQuestion, "File Exists") = vbYes Then
 SavePicture imglst.ListImages(1).ExtractIcon, CDBox2.FileName
 End If
Else
 SavePicture imglst.ListImages(1).ExtractIcon, CDBox2.FileName
End If
Exit Sub
er:

End Sub

Private Sub Command2_Click()
frmAbout.Show
End Sub

Private Sub Form_Load()
SetTrans 30, Me.hwnd
SetupTitlebar Me
'SetTrans 30, Me.hwnd
widIco = IconImg.Width
HiIco = IconImg.Height
lIndex = 0

cmdBack.Enabled = False
cmdNext.Enabled = False
lblIcons.Enabled = False
lblIcon.Enabled = False
picSmall.Enabled = False
picLarge.Enabled = False
Label1.Enabled = False
Label2.Enabled = False
Label3.Enabled = False
Label4.Enabled = False

'
' Align the toolbars to the top of the form.
'
With tbLarge
    .AllowCustomize = False
    .Wrappable = False
    .BorderStyle = ccNone
End With

With tbSmall
    .AllowCustomize = False
    .Wrappable = False
    .BorderStyle = ccNone
End With
'
' Set the dimensions of the PictureBox controls where the
' icons will be drawn.  We will use 32x32 and 16x16 icons.
' Each size uses its own PictureBox.
'
picLarge.Height = LARGE_ICON * Screen.TwipsPerPixelY
picLarge.Width = LARGE_ICON * Screen.TwipsPerPixelX
picSmall.Height = SMALL_ICON * Screen.TwipsPerPixelY
picSmall.Width = SMALL_ICON * Screen.TwipsPerPixelX

End Sub



Private Sub GoldButton1_Click()
On Error GoTo er

cdbox.ShowOpen
imglst.ListImages.Add 1, , LoadPicture(cdbox.FileName)

IconView.Picture = imglst.ListImages(1).ExtractIcon

Exit Sub

er:
If cdbox.FileName <> "" Then
MsgBox "Invalid File", vbInformation, "Error"
End If

End Sub

Private Sub IconImg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PopupMenu MnuControl
End Sub



Private Sub IconImg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

IconImg.Width = 500
IconImg.Height = 500


End Sub



Private Sub IconView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
IconView.ZOrder 0
IconView.Width = 2000
IconView.Height = 2000
End If
End Sub

Private Sub IconView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
WindowCaption.ForeColor = vbWhite
resetIcon
End Sub

Private Sub IconView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
IconView.Width = 1335
IconView.Height = 1335
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
WindowCaption.ForeColor = vbWhite
resetIcon

End Sub

Private Sub mnu_removeTrans_Click()
SetTrans 0, Me.hwnd

End Sub

Private Sub Mnu_Transp_Click()
SetTrans 30, Me.hwnd
End Sub

Private Sub MnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuMinimize_Click()
Form1.WindowState = vbMinimized

End Sub



Private Sub msk_Click()
On Error Resume Next
If msk.Value Then
imglst.UseMaskColor = True
imglst.ListImages.Add 1, , LoadPicture(cdbox.FileName)

IconView.Picture = imglst.ListImages(1).ExtractIcon
Else
maskclr.Visible = False
imglst.UseMaskColor = False
imglst.ListImages.Add 1, , LoadPicture(cdbox.FileName)

IconView.Picture = imglst.ListImages(1).ExtractIcon
End If
End Sub

Private Sub Open_cmd_Click()

On Error GoTo er

cdbox.ShowOpen
imglst.ListImages.Add 1, , LoadPicture(cdbox.FileName)

IconView.Picture = imglst.ListImages(1).ExtractIcon

Exit Sub

er:
If cdbox.FileName <> "" Then
MsgBox "Invalid File", vbInformation, "Error"
End If
End Sub



Private Sub S_Click()
 With cd1
        'Set The Title Of The Save Window To "Choose a filename to save"
        .DialogTitle = "Choose a filename to save"
        'Set The Save Filter Of The Save Window To "24-bit Bitmap (*.bmp)|*.bmp"
        .Filter = "Icon file (*.ico)|*.ico"
        'Set The Filters Index To 1
        .FilterIndex = 1
        'Set The File Name To "" (Blank)
        .FileName = ""
        'Show The Save Window
        .ShowSave
        
    'If The File Name Is "" (Blank) Then Exit Sub
    If .FileName = "" Then
        Exit Sub
    End If
        'Save Picture PicBox's Image as .FileName
        SavePicture Form1.picLarge.Image, .FileName
    'End With cd1 (Common Dialog)
    End With

End Sub

Private Sub Save_cmd_Click()
On Error GoTo er
With CDBox2
   .Filter = "Icon File(.ico)|*.ico|"
   .ShowSave
End With
If fso.FileExists(CDBox2.FileName) = True Then
 If MsgBox("File '" & CDBox2.FileTitle & "' Already exists" & vbNewLine & "Do you want to overwrite?", vbYesNo + vbQuestion, "File Exists") = vbYes Then
 SavePicture imglst.ListImages(1).ExtractIcon, CDBox2.FileName
 End If
Else
 SavePicture imglst.ListImages(1).ExtractIcon, CDBox2.FileName
End If
Exit Sub
er:

End Sub
Private Sub Closeb_Click()
End
End Sub

Private Sub Form_Resize()
'SetupTitlebar Me

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
resetIcon
WindowCaption.ForeColor = vbBlue
If Button = 1 Then _
TitleBar.Parent.Move TitleBar.Parent.Left + X - GapX, TitleBar.Parent.Top + Y - GapY
End Sub

Private Sub windowcaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
resetIcon
If Button = 1 Then
GapX = X
GapY = Y
End If
End Sub

Private Sub windowcaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
WindowCaption.ForeColor = vbBlue
If Button = 1 Then _
TitleBar.Parent.Move TitleBar.Parent.Left + X - GapX, TitleBar.Parent.Top + Y - GapY
End Sub

Public Sub pGetIcon()
Dim l As Long
'
' Get the handle of the icon indicated by lIndex.
'
Call ExtractIconEx(sExeName, lIndex, glLargeIcons(lIndex), glSmallIcons(lIndex), 1)
'
' Draw the icon to respective picturebox control.
'
With picLarge
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    Call DrawIconEx(.hdc, 0, 0, glLargeIcons(lIndex), LARGE_ICON, LARGE_ICON, 0, 0, DI_NORMAL)
    .Refresh
End With

With picSmall
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    Call DrawIconEx(.hdc, 0, 0, glSmallIcons(lIndex), SMALL_ICON, SMALL_ICON, 0, 0, DI_NORMAL)
    .Refresh
End With
lblIcon = lIndex
End Sub


