VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   Picture         =   "NewMain.frx":0000
   ScaleHeight     =   3315
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmrMove 
      Interval        =   1
      Left            =   3840
      Top             =   2880
   End
   Begin VB.Label LblC 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by Yehia Ahmed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   -720
      TabIndex        =   3
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1800
      MouseIcon       =   "NewMain.frx":55C2
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address Book"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1800
      MouseIcon       =   "NewMain.frx":58CC
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Image Img2 
      Height          =   720
      Left            =   960
      MouseIcon       =   "NewMain.frx":5BD6
      MousePointer    =   99  'Custom
      Picture         =   "NewMain.frx":5EE0
      Top             =   1800
      Width           =   720
   End
   Begin VB.Image Img1 
      Height          =   720
      Left            =   960
      MouseIcon       =   "NewMain.frx":A41A
      MousePointer    =   99  'Custom
      Picture         =   "NewMain.frx":A724
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start With:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1350
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub Form_Load()
s = CreateEllipticRgn(2, 218, 283, 1)
Call SetWindowRgn(hWnd, s, True)
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Long
    If Button = 1 Then
        ReleaseCapture
        I = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If

End Sub

Private Sub Img1_Click()
Me.Hide: Ab.Show
End Sub

Private Sub Img2_Click()
Me.Hide: Calc.Show
End Sub

Private Sub Lbl1_Click()
Me.Hide: Ab.Show
End Sub

Private Sub Lbl2_Click()
Me.Hide: Calc.Show
End Sub

Private Sub TmrMove_Timer()
If LblC.Left < 3820 Then LblC.Left = LblC.Left + 20
End Sub
