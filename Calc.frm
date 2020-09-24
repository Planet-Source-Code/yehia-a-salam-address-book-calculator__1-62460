VERSION 5.00
Begin VB.Form Calc 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Calculator"
   ClientHeight    =   4305
   ClientLeft      =   5115
   ClientTop       =   2010
   ClientWidth     =   3360
   Icon            =   "Calc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   3360
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton CmdCalculations 
      Caption         =   "nCr"
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
      Index           =   6
      Left            =   2040
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton CmdCalculations 
      Caption         =   "nPr"
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
      Index           =   5
      Left            =   1440
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton CmdOpr 
      Caption         =   "cos"
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
      Index           =   7
      Left            =   2640
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton CmdOpr 
      Caption         =   "tan"
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
      Index           =   6
      Left            =   2040
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton CmdOpr 
      Caption         =   "n!"
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
      Index           =   5
      Left            =   1440
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton CmdCalculations 
      Caption         =   "xª"
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
      Index           =   4
      Left            =   240
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton CmdOpr 
      Caption         =   "x³"
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
      Index           =   4
      Left            =   840
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton CmdCopy 
      Caption         =   "Copy"
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
      Left            =   1920
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton CmdPaste 
      Caption         =   "Paste"
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
      Left            =   2520
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton CmdOpr 
      Caption         =   "1/x"
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
      Index           =   0
      Left            =   2640
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton CalcMemory 
      Caption         =   "Mr"
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
      Index           =   0
      Left            =   240
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton CmdOpr 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   2040
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton CmdOpr 
      Caption         =   "Sqr"
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
      Index           =   1
      Left            =   2040
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton CmdNbr 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   1440
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton CmdPlusMinus 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton CmdOpr 
      Caption         =   "Del"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   2040
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton CmdOpr 
      Caption         =   "X x 1"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton CalcMemory 
      Caption         =   "M-"
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
      Index           =   2
      Left            =   240
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton CalcMemory 
      Caption         =   "M+"
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
      Index           =   1
      Left            =   840
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton CmdCalculations 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton CmdOpr 
      Caption         =   "x²"
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
      Index           =   3
      Left            =   840
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton CmdCalculations 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton CmdEqual 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton CmdCalculations 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2040
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton CmdCalculations 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton AC 
      Caption         =   "AC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton CmdNbr 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton CmdNbr 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   1440
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton CmdNbr 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   840
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton CmdNbr 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   240
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton CmdNbr 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1440
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton CmdNbr 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   840
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton CmdNbr 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton CmdNbr 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1440
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton CmdNbr 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton CmdNbr 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label LblSwitch 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H005499FD&
      Caption         =   "Switch To..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      MouseIcon       =   "Calc.frx":1E72
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1025
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   240
      TabIndex        =   24
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, Y, Z As Double
Dim Number01, Number02, Nbr As Double
Dim A As Double
Dim Clip As String
Dim Fact01, Fact02, Fact03 As Double
Dim Temp As Double
Private Sub AC_Click()
Lbl = ""
X = 0
Y = 0
Z = 0
End Sub

Private Sub CalcMemory_Click(Index As Integer)
Select Case Index
    Case 0 'Mr (add to memory)
        Clip = Lbl
    Case 1 'M+ (put to the display)
        Lbl = Lbl + Clip
    Case 2 'M- (clear memory)
        Clip = ""
        MsgBox "Memory Cleared..", vbInformation, "Calculator"
End Select
End Sub

Private Sub CmdCalculations_Click(Index As Integer)
Select Case Index
    Case 0 '+
        X = Val(Lbl)
        A = 1
    Case 1 '-
        X = Lbl
        A = 2
    Case 2 'x
        X = Lbl
        A = 3
    Case 3 ':
        X = Lbl
        A = 4
    Case 4 'X^y
        X = Lbl
        A = 5
    Case 5 'nPr
        X = Lbl
        A = 6
    Case 6 'nCr
        X = Lbl
        A = 7
End Select
Lbl = ""
End Sub

Private Sub CmdCopy_Click()
Clipboard.SetText Lbl, vbCFText
End Sub

Private Sub CmdNbr_Click(Index As Integer)
Lbl = Lbl + (CmdNbr(Index).Caption)
End Sub

Private Sub CmdPaste_Click()
Lbl = Val(Clipboard.GetText)
End Sub
Private Sub CmdOpr_Click(Index As Integer)
On Error GoTo EH
Y = Lbl
Select Case CmdOpr.Item(Index).Caption
    Case "1/x"
        Lbl = 1 / Lbl
    Case "Sqr"
        Lbl = Sqr(Lbl)
    Case "cos"
        Lbl = Cos(Lbl)
    Case "tan"
        Lbl = Tan(Lbl)
    Case "X x 1"
        CalcMulti.Show
    Case "x²"
        Lbl = Lbl * Lbl
    Case "x³"
        Lbl = Lbl * Lbl * Lbl
    Case "n!"
        GetFactor Lbl
        Lbl = Temp
    Case "%"
        Lbl = X / Y
        Lbl = Lbl * 100
    Case "Del"
        Lbl = Left(Lbl, Len(Lbl) - 1)
End Select
EH:
Call ErrorHandler
End Sub

Private Sub CmdPlusMinus_Click()
re = InStr(1, Lbl, "-", vbTextCompare)
Select Case re
    Case Is = 0 'Minus Not Found
        Lbl = "-" + Lbl
    Case Is <> 0    'Minus Found
          Lbl = Right(Lbl, (Len(Lbl) - 1))
End Select
End Sub

Private Sub CmdEqual_Click()
Y = Lbl
On Error GoTo EH
Select Case A
    Case 1  'Plus
        Lbl = X + Y
    Case 2  'Minus
        Lbl = X - Y
    Case 3  'Multi
        Lbl = X * Y
    Case 4  'Divide
        Lbl = X / Y
    Case 5  'Square By Y
    Nbr = X
        For I = 1 To Y - 1
            X = X * Nbr
        Next I
        Lbl = X
    Case 6 'nPr
        FactorOpr X, Y
        Lbl = Fact01 / Fact03
    Case 7 'nCr
        FactorOpr X, Y
        Lbl = Fact01 / (Fact03 * Fact02)
End Select

EH:
Call ErrorHandler
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKey0, vbKeyNumpad0
        Lbl = Lbl + "0"
        re = SendMessage(CmdNbr(0).hWnd, BM_SETSTATE, 1, ByVal 0&)
    Case vbKey1, vbKeyNumpad1
        Lbl = Lbl + "1"
        re = SendMessage(CmdNbr(1).hWnd, BM_SETSTATE, 1, ByVal 0&)
    Case vbKey2, vbKeyNumpad2
        Lbl = Lbl + "2"
        re = SendMessage(CmdNbr(2).hWnd, BM_SETSTATE, 1, ByVal 0&)
    Case vbKey3, vbKeyNumpad3
        Lbl = Lbl + "3"
        re = SendMessage(CmdNbr(3).hWnd, BM_SETSTATE, 1, ByVal 0&)
    Case vbKey4, vbKeyNumpad4
        Lbl = Lbl + "4"
        re = SendMessage(CmdNbr(4).hWnd, BM_SETSTATE, 1, ByVal 0&)
    Case vbKey5, vbKeyNumpad5
        Lbl = Lbl + "5"
        re = SendMessage(CmdNbr(5).hWnd, BM_SETSTATE, 1, ByVal 0&)
    Case vbKey6, vbKeyNumpad6
        Lbl = Lbl + "6"
        re = SendMessage(CmdNbr(6).hWnd, BM_SETSTATE, 1, ByVal 0&)
    Case vbKey7, vbKeyNumpad7
        Lbl = Lbl + "7"
        re = SendMessage(CmdNbr(7).hWnd, BM_SETSTATE, 1, ByVal 0&)
    Case vbKey8, vbKeyNumpad8
        Lbl = Lbl + "8"
        re = SendMessage(CmdNbr(8).hWnd, BM_SETSTATE, 1, ByVal 0&)
    Case vbKey9, vbKeyNumpad9
        Lbl = Lbl + "9"
        re = SendMessage(CmdNbr(9).hWnd, BM_SETSTATE, 1, ByVal 0&)
    Case vbKeyDecimal
        Lbl = Lbl + "."
        re = SendMessage(CmdNbr(10).hWnd, BM_SETSTATE, 1, ByVal 0&)
    Case vbKeyAdd
        Call CmdCalculations_Click(0)
        re = SendMessage(CmdCalculations(0).hWnd, BM_SETSTATE, 1, ByVal 0&)
    Case vbKeySubtract
        Call CmdCalculations_Click(1)
        re = SendMessage(CmdCalculations(1).hWnd, BM_SETSTATE, 1, ByVal 0&)
    Case vbKeyMultiply
        Call CmdCalculations_Click(2)
        re = SendMessage(CmdCalculations(2).hWnd, BM_SETSTATE, 1, ByVal 0&)
    Case vbKeyDivide
        Call CmdCalculations_Click(3)
        re = SendMessage(CmdCalculations(3).hWnd, BM_SETSTATE, 1, ByVal 0&)
    Case vbKeyDelete, vbKeyBack
        Call CmdOpr_Click(9)
        re = SendMessage(CmdOpr(9).hWnd, BM_SETSTATE, 1, ByVal 0&)
    Case vbKeyReturn, vbKeySeparator
        Call CmdEqual_Click
        re = SendMessage(CmdEqual.hWnd, BM_SETSTATE, 1, ByVal 0&)

End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
For I = 0 To 10
    re = SendMessage(CmdNbr(I).hWnd, BM_SETSTATE, ByVal 0&, ByVal 0&)
Next I
For I = 0 To 3
    re = SendMessage(CmdCalculations(I).hWnd, BM_SETSTATE, ByVal 0&, ByVal 0&)
Next I
For I = 0 To 9
    re = SendMessage(CmdOpr(I).hWnd, BM_SETSTATE, ByVal 0&, ByVal 0&)
Next I
End Sub

Private Sub Form_Load()
Calc.Picture = Ab.Picture
End Sub

Private Sub Lbl_Change()
If Len(Lbl) > 21 Then Lbl = Right(Lbl, 21)
End Sub

Private Sub ErrorHandler()
Select Case Err.Number
    Case 11 'Division By Zero
        Lbl = "Cannot divide by Zero"
    Case 6 'OverFlow
        Lbl = "OverFlow"
End Select
End Sub
Function GetFactor(Nbr)
Temp = 1
For I = 1 To Nbr
Temp = Temp * I
Next I
End Function

Function FactorOpr(Number01, Number02)
If Number01 > 69 Or Number01 < Number02 Then
ErrorHandler
Exit Function
End If
GetFactor Number01
Fact01 = Temp 'X! Factor For Number One
GetFactor Number02
Fact02 = Temp 'X! Factor For Number Two
GetFactor Number01 - Number02
Fact03 = Temp 'X!Factor For Number One-Two
End Function

Private Sub LblSwitch_Click()
PopupMenu CalcMulti.MnuSwitch, , 1300, 600
End Sub
