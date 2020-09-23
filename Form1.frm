VERSION 5.00
Begin VB.Form FrmEdit 
   Caption         =   "Form1"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3435
   ScaleWidth      =   4935
   Begin VB.PictureBox UndoPic 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   1320
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin PictureEditor.PicEdit PicEdit 
      Height          =   2895
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5106
   End
End
Attribute VB_Name = "FrmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Filename As String

Private Sub Form_Activate()
    On Error Resume Next
    Me.Width = PicEdit.Width * 2
    Me.Height = PicEdit.Height * 2
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Width = PicEdit.Width * 2
    Me.Height = PicEdit.Height * 2
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    PicEdit.Left = (Me.Width / 2) - (PicEdit.Width / 2)
    PicEdit.Top = (Me.Height / 2) - (PicEdit.Height / 2)
End Sub

Private Sub PicEdit_Progress(Prog As Integer, State As String)
    On Error Resume Next
    MDIEdit.ProgressBar.Value = Prog
    MDIEdit.Progress.Text = State
End Sub

Private Sub PicEdit_Resize(Width As Long, Height As Long)
    On Error Resume Next
    Me.Width = PicEdit.Width * 2
    Me.Height = PicEdit.Height * 2
End Sub

