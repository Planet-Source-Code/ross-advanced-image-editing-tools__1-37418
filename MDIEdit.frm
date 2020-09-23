VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIEdit 
   BackColor       =   &H8000000C&
   Caption         =   "Pro Picture Editor"
   ClientHeight    =   7125
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7035
   Icon            =   "MDIEdit.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   120
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.TextBox Progress 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         HideSelection   =   0   'False
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog C 
      Left            =   1680
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu saveas 
         Caption         =   "Save &As"
      End
      Begin VB.Menu b 
         Caption         =   "-"
      End
      Begin VB.Menu close 
         Caption         =   "&Close"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu undo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu d 
         Caption         =   "-"
      End
      Begin VB.Menu brighten 
         Caption         =   "Brighten"
      End
      Begin VB.Menu colmono 
         Caption         =   "Colour to Mono"
      End
      Begin VB.Menu diffuse 
         Caption         =   "Diffuse"
      End
      Begin VB.Menu emboss 
         Caption         =   "Emboss"
      End
      Begin VB.Menu engrave 
         Caption         =   "Engrave"
      End
      Begin VB.Menu invert 
         Caption         =   "Invert"
      End
      Begin VB.Menu midday 
         Caption         =   "Midday"
      End
      Begin VB.Menu mosaic 
         Caption         =   "Mosaic"
      End
      Begin VB.Menu neon 
         Caption         =   "Neon"
      End
      Begin VB.Menu sharpen 
         Caption         =   "Sharpen"
      End
      Begin VB.Menu smooth 
         Caption         =   "Smooth"
      End
      Begin VB.Menu f 
         Caption         =   "-"
      End
      Begin VB.Menu ered 
         Caption         =   "Edit Red"
      End
      Begin VB.Menu egreen 
         Caption         =   "Edit Green"
      End
      Begin VB.Menu eblue 
         Caption         =   "Edit Blue"
      End
      Begin VB.Menu h 
         Caption         =   "-"
      End
      Begin VB.Menu tonebalance 
         Caption         =   "Tone Balance"
      End
      Begin VB.Menu colbala 
         Caption         =   "Colour Balance"
      End
   End
End
Attribute VB_Name = "MDIEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub brighten_Click()
    ActiveForm.PicEdit.brighten
End Sub

Private Sub close_Click()
    On Error Resume Next
    Unload ActiveForm
End Sub

Private Sub colbala_Click()
    ActiveForm.PicEdit.ColourBalance
End Sub

Private Sub colmono_Click()
    ActiveForm.PicEdit.ColourToMono
End Sub

Private Sub Command1_Click()
    ActiveForm.PicEdit.ColorBalance
End Sub

Private Sub diffuse_Click()
    ActiveForm.PicEdit.diffuse
End Sub

Private Sub eblue_Click()
    ActiveForm.PicEdit.EditBlue
End Sub

Private Sub egreen_Click()
    ActiveForm.PicEdit.EditGreen
End Sub

Private Sub Emboss_Click()
    ActiveForm.PicEdit.emboss
End Sub

Private Sub engrave_Click()
    ActiveForm.PicEdit.engrave
End Sub

Private Sub ered_Click()
    ActiveForm.PicEdit.EditRed
End Sub

Private Sub exit_Click()
    Unload Me
    DoEvents
    End
End Sub

Private Sub invert_Click()
    ActiveForm.PicEdit.invert
End Sub

Private Sub MDIForm_Load()
    Me.WindowState = 2
End Sub

Private Sub MDIForm_Resize()
ProgressBar.Width = Me.Width - 1800
End Sub

Private Sub midday_Click()
    ActiveForm.PicEdit.midday
End Sub

Private Sub mosaic_Click()
    ActiveForm.PicEdit.mosaic
End Sub

Private Sub neon_Click()
    ActiveForm.PicEdit.neon
End Sub

Private Sub new_Click()
    OpenNew
End Sub

Function OpenNew()
    On Error Resume Next
    Static Count As Long
    Dim frm As FrmEdit
    Count = Count + 1
    Set frm = New FrmEdit
    frm.Caption = "Picture Editor"
    frm.Show
End Function

Private Sub open_Click()
    On Error Resume Next
    Dim Filename As String
    If ActiveForm Is Nothing Then OpenNew
    C.DialogTitle = "Open"
    C.Filename = ""
    C.Filter = "Picture Files|*.bmp;*.jpg;*.gif"
    C.ShowOpen
    If C.Filename = "" Then Exit Sub
    Filename = C.Filename
    ActiveForm.PicEdit.OpenPicture Filename
    ActiveForm.Caption = "Picture Editor"
    ActiveForm.Filename = Filename
End Sub

Private Sub save_Click()
    On Error Resume Next
    Dim Filename As String
    Filename = ActiveForm.Filename
    ActiveForm.PicEdit.SavePic Filename
End Sub

Private Sub saveas_Click()
    On Error Resume Next
    Dim Filename As String
    C.DialogTitle = "Save As"
    C.Filename = ActiveForm.Filename
    C.Filter = "Bitmap (*.bmp)|*.bmp|JPEG (*.jpg)|*.jpg|GIF (*.gif)|*.gif"
    C.ShowSave
    If C.Filename = "" Then Exit Sub
    Filename = C.Filename
    ActiveForm.PicEdit.SavePic Filename
    ActiveForm.Filename = Filename
End Sub

Private Sub sharpen_Click()
    ActiveForm.PicEdit.sharpen
End Sub

Private Sub smooth_Click()
    ActiveForm.PicEdit.smooth
End Sub

Private Sub tonebalance_Click()
    ActiveForm.PicEdit.tonebalance
End Sub

Private Sub Undo_Click()
    ActiveForm.PicEdit.undo
End Sub
