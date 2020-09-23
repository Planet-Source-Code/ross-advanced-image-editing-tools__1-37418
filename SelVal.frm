VERSION 5.00
Begin VB.Form SelVal 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Picture Editor"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4635
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1515
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   0
      Width           =   4630
      Begin VB.CommandButton Cancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2320
         TabIndex        =   4
         Top             =   900
         Width           =   2125
      End
      Begin VB.CommandButton OK 
         Caption         =   "OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   900
         Width           =   2125
      End
      Begin VB.HScrollBar ProgBar 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   400
         Width           =   4335
      End
      Begin VB.Label Val 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   2
         Top             =   75
         Width           =   975
      End
   End
End
Attribute VB_Name = "SelVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ProgMax As Integer, ProgMin As Integer, ProgVal As Integer, EditVal As String

Private Sub Cancel_Click()
    ProgVal = 1000
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ProgBar.Min = ProgMin
    ProgBar.Max = ProgMax
    ProgBar.Value = ProgVal
    Me.Caption = EditVal
End Sub

Private Sub OK_Click()
    ProgVal = ProgBar.Value
    Unload Me
End Sub

Private Sub ProgBar_Change()
    Val.Caption = ProgBar.Value
End Sub
