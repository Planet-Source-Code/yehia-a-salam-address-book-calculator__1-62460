VERSION 5.00
Begin VB.Form Ab 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Address Book"
   ClientHeight    =   4455
   ClientLeft      =   5025
   ClientTop       =   1965
   ClientWidth     =   5955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AdressBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "AdressBook.frx":0ECA
   ScaleHeight     =   4455
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   2040
      ScaleHeight     =   825
      ScaleWidth      =   3705
      TabIndex        =   19
      Top             =   480
      Width           =   3735
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   3000
         TabIndex        =   24
         Top             =   570
         Width           =   495
      End
      Begin VB.Image ImgTools 
         Height          =   480
         Index           =   5
         Left            =   3000
         Picture         =   "AdressBook.frx":E2F4
         Stretch         =   -1  'True
         Top             =   120
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   2280
         TabIndex        =   23
         Top             =   570
         Width           =   555
      End
      Begin VB.Image ImgTools 
         Height          =   480
         Index           =   4
         Left            =   2280
         Picture         =   "AdressBook.frx":1457E
         Stretch         =   -1  'True
         Top             =   135
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1530
         TabIndex        =   22
         Top             =   570
         Width           =   585
      End
      Begin VB.Image ImgTools 
         Height          =   465
         Index           =   3
         Left            =   1560
         Picture         =   "AdressBook.frx":1A2D0
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modify"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   840
         TabIndex        =   21
         Top             =   570
         Width           =   570
      End
      Begin VB.Image ImgTools 
         Height          =   480
         Index           =   2
         Left            =   840
         Picture         =   "AdressBook.frx":1FAB2
         Top             =   135
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   570
         Width           =   330
      End
      Begin VB.Line LineUp 
         BorderColor     =   &H00DDDBE3&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   120
         X2              =   720
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line LineRight 
         BorderColor     =   &H009C9B9F&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   720
         X2              =   720
         Y1              =   120
         Y2              =   780
      End
      Begin VB.Line LineLeft 
         BorderColor     =   &H00DDDBE3&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   780
      End
      Begin VB.Line LineDown 
         BorderColor     =   &H009C9B9F&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   120
         X2              =   720
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Image ImgTools 
         Height          =   480
         Index           =   1
         Left            =   165
         Picture         =   "AdressBook.frx":2077C
         Top             =   135
         Width           =   480
      End
   End
   Begin VB.TextBox TxtSearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3360
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3360
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3360
      TabIndex        =   9
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton OldUpdate 
      Caption         =   "Old Modify"
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
      Left            =   120
      Picture         =   "AdressBook.frx":21446
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox TxtCounter 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3720
      Width           =   1815
   End
   Begin VB.ListBox List 
      Appearance      =   0  'Flat
      ForeColor       =   &H00400000&
      Height          =   3015
      IntegralHeight  =   0   'False
      ItemData        =   "AdressBook.frx":21BF0
      Left            =   120
      List            =   "AdressBook.frx":21BF2
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txt3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3240
      TabIndex        =   5
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txt2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3240
      TabIndex        =   4
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox txt1 
      Appearance      =   0  'Flat
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
      Height          =   270
      Left            =   3240
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
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
      Left            =   4920
      MouseIcon       =   "AdressBook.frx":21BF4
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   4200
      Width           =   1065
   End
   Begin VB.Label LblCompName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome "
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
      Height          =   195
      Left            =   5040
      TabIndex        =   25
      Top             =   120
      Width           =   825
   End
   Begin VB.Label LblSearchForm 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Form"
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
      Left            =   2520
      TabIndex        =   18
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      Visible         =   0   'False
      X1              =   2280
      X2              =   2400
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      Index           =   3
      Visible         =   0   'False
      X1              =   3720
      X2              =   5640
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      Visible         =   0   'False
      X1              =   2280
      X2              =   5640
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label LblChoices 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   3720
      TabIndex        =   17
      Top             =   3360
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label LblChoices 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tel."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   3720
      TabIndex        =   16
      Top             =   3600
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label LblChoices 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   3720
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lblSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search For :"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2400
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblOptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search By :"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2400
      TabIndex        =   13
      Top             =   3120
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image ImgSearch 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   5280
      MouseIcon       =   "AdressBook.frx":21EFE
      MousePointer    =   99  'Custom
      Picture         =   "AdressBook.frx":22208
      Top             =   2760
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Lbladd 
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Lblname 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "Ab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Source As Variant
Private Sub CmdUpdate_Click()
'Open App.Path + "\List.dat" For Input As #1
'Open App.Path + "\List.tmp" For Append As #2
'UpText = InputBox("Enter a Name to search for")
'Do Until EOF(1)
'  Line Input #1, X
'  Line Input #1, Y
'  Line Input #1, z
'  If UpText = Trim(X) Then
'  X = InputBox("Enter The New Name")
'  Y = InputBox("Enter The New Adrees")
'  z = InputBox("Enter The New Number")
' AnyThing = 1
' End If
' Print #2, X
' Print #2, Y
' Print #2, z
' Loop
' If AnyThing <> 1 Then
' MsgBox "Not Found"
' End If
' Close #1
' Close #2
'Kill App.Path + "\List.dat"
'Name App.Path + "\List.tmp" As App.Path + "\List.dat"
End Sub
Private Sub ListFill()
List.Clear
Open Source For Append As #1
Close #1
Open App.Path + "\List.dat" For Input As #1
Do Until EOF(1)
    Line Input #1, X
    Line Input #1, Y
    Line Input #1, Z
    List.AddItem (X)
Loop
Close #1
TxtCounter = List.ListCount & " Entries are found."
End Sub

Private Sub Form_Load()
Source = App.Path + "\List.dat"
Call ListFill
Dim Compname As String, re As Long
Compname = Space(255)
re = GetComputerName(Compname, 255)
Compname = Left(Compname, InStr(Compname, vbNullChar) - 1)
LblCompName = LblCompName + Compname
End Sub

Private Sub ImgClose_Click()
End
End Sub

Private Sub ImgTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Long
    If Button = 1 Then
        ReleaseCapture
        I = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub

Private Sub ImgSearch_Click()
Open App.Path + "\List.dat" For Input As #1
Do Until EOF(1)
    Line Input #1, X
    Line Input #1, Y
    Line Input #1, Z
    If Option1.Value = True Then Option1.Tag = X
    If Option2.Value = True Then Option1.Tag = Y
    If Option3.Value = True Then Option1.Tag = Z

    If TxtSearch.Text = Option1.Tag Then
        Anything = 1
        txt1 = X
        txt2 = Y
        txt3 = Z
        List.Text = X
    End If
Loop
If Anything <> 1 Then
    Msg = MsgBox("Entry Not Found !", vbInformation, "Alert")
End If
Close #1
End Sub



Private Sub ImgTools_Click(Index As Integer)
Select Case Index
    Case 1 'Add
        Open App.Path + "\List.dat" For Append As #1
        Print #1, txt1
        Print #1, txt2
        Print #1, txt3
        Close #1
        List.AddItem (txt1)
        txt1 = ""
        txt2 = ""
        txt3 = ""
        Call ListFill
    Case 2 'Modify
        Open App.Path + "\List.dat" For Input As #1
        Open App.Path + "\List.tmp" For Append As #2
        Do Until EOF(1)
            Line Input #1, X
            Line Input #1, Y
            Line Input #1, Z
                If X = List.Text Then
                    X = txt1
                    Y = txt2
                    Z = txt3
                    List.Text = txt1.Text
                End If
            Print #2, X
            Print #2, Y
            Print #2, Z
        Loop
        Close #1
        Close #2
        Kill App.Path + "\List.dat"
        Name App.Path + "\List.tmp" As App.Path + "\List.dat"
        Call ListFill
    Case 3 'Search
        LblSearchForm.Visible = Not LblSearchForm.Visible
        Line2.Item(1).Visible = Not Line2.Item(1).Visible
        Line2.Item(2).Visible = Not Line2.Item(2).Visible
        Line2.Item(3).Visible = Not Line2.Item(3).Visible
        Option1.Visible = Not Option1.Visible
        Option2.Visible = Not Option2.Visible
        Option3.Visible = Not Option3.Visible
        ImgSearch.Visible = Not ImgSearch.Visible
        lblSearch.Visible = Not lblSearch.Visible
        lblOptions.Visible = Not lblOptions.Visible
        LblChoices(1).Visible = Not LblChoices(1).Visible
        LblChoices(2).Visible = Not LblChoices(2).Visible
        LblChoices(3).Visible = Not LblChoices(3).Visible
        TxtSearch.Visible = Not TxtSearch.Visible
    Case 4 'Delete
        Msg = MsgBox("Are you sure that you want to Delete " & "(" & List.Text & ") ?", vbYesNo + vbQuestion, "Alert")
        If Msg = 6 Then
        Open App.Path + "\List.dat" For Input As #1
        Open App.Path + "\List.tmp" For Append As #2
            Do Until EOF(1)
                Line Input #1, X
                Line Input #1, Y
                Line Input #1, Z
            If List.Text <> Trim(X) Then
                Print #2, X
                Print #2, Y
                Print #2, Z
            End If
            Loop
            List.RemoveItem List.ListIndex
        Close #1
        Close #2
        Kill App.Path + "\List.dat"
        Name App.Path + "\List.tmp" As App.Path + "\List.dat"
        Call ListFill
        End If
    Case 5 'Reset
        Alert = MsgBox("This will delete all entries,Procces ?", vbYesNo + vbInformation, "Alert")
        If Alert = 6 Then
            Kill App.Path + "\List.dat"
            Open App.Path + "\List.dat" For Output As #1
            List.Clear
            Close #1
        End If
        Call ListFill
End Select
End Sub

Private Sub ImgTools_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
LineUp.BorderColor = &H9C9B9F: LineDown.BorderColor = &HDDDBE3
LineRight.BorderColor = &HDDDBE3: LineLeft.BorderColor = &H9C9B9F
End Sub

Private Sub ImgTools_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
LineUp.Visible = True: LineDown.Visible = True: LineLeft.Visible = True: LineRight.Visible = True
LineUp.X1 = ImgTools(Index).Left - 40: LineUp.X2 = LineUp.X1 + 600
LineLeft.X1 = ImgTools(Index).Left - 40: LineLeft.X2 = ImgTools(Index).Left - 40
LineRight.X1 = LineLeft.X2 + 620: LineRight.X2 = LineRight.X1
LineDown.X1 = LineUp.X1: LineDown.X2 = LineUp.X2
End Sub

Private Sub ImgTools_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
LineDown.BorderColor = &H9C9B9F: LineRight.BorderColor = &H9C9B9F
LineLeft.BorderColor = &HDDDBE3: LineUp.BorderColor = &HDDDBE3
End Sub

Private Sub LblSwitch_Click()
PopupMenu CalcMulti.MnuSwitch, , 3100, 3600
End Sub

Private Sub List_Click()
Open App.Path + "\List.dat" For Input As #3
L = List.Text
Do Until EOF(3)
        Line Input #3, X
        Line Input #3, Y
        Line Input #3, Z
    If L = X Then
        txt1 = X
        txt2 = Y
        txt3 = Z
    End If
Loop
Close #3
End Sub



Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LineUp.Visible = False: LineDown.Visible = False
LineLeft.Visible = False: LineRight.Visible = False
End Sub
