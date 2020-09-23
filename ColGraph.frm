VERSION 5.00
Begin VB.Form ColGraph 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Colour Balance"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   6720
      Width           =   8055
      Begin VB.ComboBox Combo 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "ColGraph.frx":0000
         Left            =   200
         List            =   "ColGraph.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1320
         Width           =   2655
      End
      Begin VB.HScrollBar HScroll 
         Height          =   375
         Left            =   1260
         Max             =   255
         Min             =   -255
         TabIndex        =   8
         Top             =   360
         Width           =   5535
      End
      Begin VB.CommandButton Cancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   6120
         TabIndex        =   7
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton Ok 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   495
         Left            =   4200
         TabIndex        =   6
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   " Dark "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   600
         TabIndex        =   13
         Top             =   360
         Width           =   585
      End
      Begin VB.Label lblmax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   " Max = 255 "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6360
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblmin 
         Appearance      =   0  'Flat
         Caption         =   " Min = 0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   " Bright "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6840
         TabIndex        =   10
         Top             =   360
         Width           =   660
      End
      Begin VB.Label lblval 
         Alignment       =   2  'Center
         Caption         =   " Edit Value = 0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   840
         Width           =   2295
      End
   End
   Begin VB.PictureBox Temp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   9120
      ScaleHeight     =   1335
      ScaleWidth      =   2055
      TabIndex        =   1
      Top             =   6240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox View 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   6015
      Left            =   120
      ScaleHeight     =   5985
      ScaleMode       =   0  'User
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8160
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   6240
      Y2              =   6360
   End
   Begin VB.Line Line3 
      X1              =   4140
      X2              =   4140
      Y1              =   6240
      Y2              =   6360
   End
   Begin VB.Line Line4 
      X1              =   8160
      X2              =   8160
      Y1              =   6240
      Y2              =   6360
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "127"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4005
      TabIndex        =   3
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   2
      Top             =   6480
      Width           =   615
   End
End
Attribute VB_Name = "ColGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MinVal As Integer, MaxVal As Integer, Val As Integer, Opt As Integer

Function Draw()
    Dim Prev(2) As Long, NewVal As Long, hNo As Long, ColV(2) As ColorConstants
    hNo = 0
    For i = 0 To 255
        For j = 1 To 3
            If ColVal(j, i) > hNo Then hNo = ColVal(j, i)
        Next
    Next
    
    ColV(0) = vbRed
    ColV(1) = vbGreen
    ColV(2) = vbBlue
    
    View.ScaleHeight = hNo + 100
    View.ScaleWidth = 256
    
    For i = 0 To 2
        Prev(i) = ColVal(i + 1, 0)
    Next
    
    For i = 1 To 255
        For j = 0 To 2
            NewVal = View.ScaleHeight - ColVal(j + 1, i) - 50
            View.Line (i - 1, Prev(j))-(i, NewVal), ColV(j)
            Prev(j) = NewVal
        Next
    Next
    Temp.Picture = View.Image
End Function

Private Sub Cancel_Click()
    MaxVal = -1
    Unload Me
End Sub

Private Sub Combo_Change()
    Opt = Combo.ListIndex
End Sub

Private Sub Form_Activate()
    Combo.ListIndex = 0
    MinVal = 0
    MaxVal = 255
    Val = 0
    Opt = 0
    Draw
End Sub

Private Sub HScroll_Change()
    lblval.Caption = " Edit Value = " & HScroll.Value
    Val = -HScroll.Value
End Sub

Private Sub Ok_Click()
    Opt = Combo.ListIndex
    Unload Me
End Sub

Private Sub View_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim SelVal As Integer
    SelVal = X
    If Button = vbRightButton Then
        View.Cls
        View.Picture = Temp.Picture
        If SelVal <= MinVal Then
            SelVal = MinVal + 1
        End If
        MaxVal = SelVal
        View.Line (SelVal, 0)-(SelVal, View.ScaleHeight), vbBlack
        View.Line (MinVal, 0)-(MinVal, View.ScaleHeight), RGB(127, 127, 127)
    Else
        If Button = vbLeftButton Then
            View.Cls
            View.Picture = Temp.Picture
            If SelVal >= MaxVal Then
                SelVal = MaxVal - 1
            End If
            MinVal = SelVal
            View.Line (SelVal, 0)-(SelVal, View.ScaleHeight), RGB(127, 127, 127)
            View.Line (MaxVal, 0)-(MaxVal, View.ScaleHeight), vbBlack
        End If
    End If
    lblmin.Caption = " Min = " & MinVal
    lblmax.Caption = " Max = " & MaxVal
End Sub
