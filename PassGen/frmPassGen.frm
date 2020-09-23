VERSION 5.00
Begin VB.Form frmPassGen 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "~@PassGen@~"
   ClientHeight    =   4215
   ClientLeft      =   5580
   ClientTop       =   1800
   ClientWidth     =   4815
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPassGen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4815
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   3240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   3615
   End
   Begin VB.Frame fraSetting 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2000
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   4575
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1550
         Left            =   2520
         TabIndex        =   12
         Top             =   240
         Width           =   1700
         Begin VB.HScrollBar bar 
            Height          =   340
            Left            =   960
            Max             =   99
            Min             =   1
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   840
            Value           =   1
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Length (1~99)"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblLength 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "99"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1550
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   1700
         Begin VB.CheckBox chkSym 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Symbols"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkNum 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Numeric"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   840
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkLower 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Lowercase"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   600
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkUpper 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Uppercase"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Value           =   1  'Checked
            Width           =   1215
         End
      End
   End
   Begin VB.Frame fraMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2000
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   4575
      Begin VB.ListBox lst 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1200
         Left            =   120
         TabIndex        =   2
         Top             =   195
         Width           =   4335
      End
      Begin VB.Label lblGen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   2880
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblGen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Copy to ClipBoard"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   2775
      End
   End
   Begin VB.Shape Shape1 
      Height          =   4215
      Left            =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   360
      Picture         =   "frmPassGen.frx":74F2
      Top             =   120
      Width           =   3435
   End
   Begin VB.Label lblGen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   19
      Top             =   360
      Width           =   255
   End
   Begin VB.Label lblGen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   4080
      TabIndex        =   18
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      Height          =   255
      Left            =   4305
      TabIndex        =   17
      Top             =   4005
      Width           =   495
   End
   Begin VB.Label lblGen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Option"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   2640
      TabIndex        =   16
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   240
      TabIndex        =   5
      Top             =   1125
      Width           =   4335
   End
   Begin VB.Label lblGen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Generate"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4800
      Y1              =   525
      Y2              =   510
   End
End
Attribute VB_Name = "frmPassGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/// Password Generater Ver.1 by K3N
'/// I don't really know who wants this crap, but i just felt like uploading!
'/// any questions, suggestions, shit to me? email at wwhc3@hotmail.com
Option Explicit
Dim upper, lower, Numeric, Sym As Boolean
Dim length As Byte
Dim c(92) As String
Dim blnMain As Boolean

Public Function PassGen()
Dim char, password As String
Dim i, s As Integer

Call createArray

'generate  passwords
Randomize
For i = 0 To length - 1
a:
    s = Int(Rnd * 92)
    
    'check characters based on the user setting, re-generate if necessary
    If upper = False Then
        If s >= 41 And s <= 66 Then GoTo a:
    End If

    If lower = False Then
        If s >= 67 And s <= 92 Then GoTo a:
    End If
    
    If Numeric = False Then
        If s >= 31 And s <= 40 Then GoTo a:
    End If

    If Sym = False Then
        If s >= 0 And s <= 30 Then GoTo a:
    End If

    char = c(s)
    password = password & char
Next

PassGen = password
End Function

Public Sub createArray()
Dim i, X As Byte

'33 ~ 47, 58 ~ 64, 91 ~ 96, 123 ~ 126 except 34 = symbols - c(0 - 30)  chr(34) is " which causes an run time error when reading back from a text file
'48 ~ 57 =  0 ~ 9 - c(31 - 40)
'65 ~ 90 = A ~ Z - c(41 - 66)
'97 ~ 122 = a ~ z - c(67 - 92)

'put all characters, numbers, and symbols into an array
X = 33
For i = 0 To 30
    c(i) = Chr(X)
    X = X + 1
    If X = 34 Then
        X = 35
    ElseIf X = 48 Then
        X = 58
    ElseIf X = 65 Then
        X = 91
    ElseIf X = 97 Then
        X = 123
    ElseIf X = 127 Then
        Exit For
    End If
Next

X = 48
For i = 31 To UBound(c)
    c(i) = Chr(X)
    X = X + 1
    If X = 58 Then
        X = 65
    ElseIf X = 91 Then
        X = 97
    ElseIf X = 123 Then
        Exit For
    End If
Next

End Sub

Private Sub bar_Change()
lblLength = bar.Value
End Sub

Private Sub chkLower_Click()
lower = chkLower.Value
End Sub

Private Sub chkNum_Click()
Numeric = chkNum.Value
End Sub

Private Sub chkSym_Click()
Sym = chkSym.Value
End Sub

Private Sub chkUpper_Click()
upper = chkUpper.Value
End Sub

Private Sub Form_Load()

If Dir("passgen.ini") <> "" Then
    Open "PassGen.ini" For Input As #1
        Input #1, upper, lower, Numeric, Sym, length
    Close #1
Else
'if the file is not found
    lower = True
    length = 10
End If

If upper = True Then
    chkUpper.Value = 1
Else
    chkUpper.Value = 0
End If

If lower = True Then
   chkLower.Value = 1
Else
    chkLower.Value = 0
End If

If Numeric = True Then
    chkNum.Value = 1
Else
    chkNum.Value = 0
End If

If Sym = True Then
    chkSym.Value = 1
Else
    chkSym.Value = 0
End If

lblLength.Caption = length
bar.Value = length

fraSetting.Left = Me.Width
fraSetting.Top = fraMain.Top

blnMain = True

lblStatus = "Password Generater Ver.1" & vbCrLf & "Created By K3N"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte
'reset the labels' color
For i = 0 To 5
lblGen(i).BackColor = vbWhite
lblGen(i).ForeColor = vbBlack
Next

lblStatus = "Password Generater Ver.1" & vbCrLf & "Created By K3N"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call SaveSetting
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblStatus = "Check/Uncheck options you want to include/exclude."
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblStatus = "Set the password length!"
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
'logo made with photshop by me
End Sub

Private Sub Label2_Click()
MsgBox "Password Generater Version 1" & vbCrLf & _
        "Inspired by Tweak-XP" & vbCrLf & _
        "Created by K3N" & vbCrLf & _
        "With Micro$oft Vi$ual Ba$ic 6" & vbCrLf & _
        "In approximately 6 hours" & vbCrLf & _
        "Copyright @ hell - No rights reserved", vbInformation, "About this crap..."
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStatus = "About this crap..."
End Sub


Private Sub lblGen_Click(Index As Integer)
Dim pass As String

'if generate button(label) is clicked
If Index = 0 Then
    Call Validate
    'if the generate frame is on the form then generate
    If blnMain = True Then
        pass = PassGen
        lst.AddItem (pass)
    'if the option frame is on the form then scroll right to show the generate frame
    Else
        Timer1.Enabled = False
        Timer2.Enabled = True
        blnMain = True
    End If
'if the copy button is clicked
ElseIf Index = 1 Then
    Clipboard.Clear
    Clipboard.SetText (lst.Text)
    MsgBox "Copied the following text:" & vbCrLf & Clipboard.GetText, vbOKOnly, "Password Generater Version.1"
'if the remove button is clicked
ElseIf Index = 2 Then
    If lst.ListIndex = -1 Then
        lst.Clear
    Else
        lst.RemoveItem (lst.ListIndex)
    End If
'if the option button is clicked and if the generate frame is on the form then scroll left to show the option frame
ElseIf Index = 3 Then
    If blnMain = True Then
        Timer1.Enabled = True
        Timer2.Enabled = False
        blnMain = False
    End If
'if the minimum button is clicked
ElseIf Index = 4 Then
    Call min
    Me.WindowState = 1
'shutdown
ElseIf Index = 5 Then
    Unload Me
    End
End If
End Sub

Private Sub lblGen_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblGen(Index).BackColor = vbBlack
lblGen(Index).ForeColor = vbRed
End Sub

Private Sub lblGen_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte

For i = 0 To 5
    If i <> Index Then
        lblGen(i).BackColor = vbWhite
        lblGen(i).ForeColor = vbBlack
    End If
Next
    lblGen(Index).BackColor = vbRed
    lblGen(Index).ForeColor = vbWhite
    
    If Index = 0 Then
        lblStatus = "Click to generate passwords!"
    ElseIf Index = 1 Then
        lblStatus = "Click to copy the selected password onto the clipboard!"
    ElseIf Index = 2 Then
        lblStatus = "Click to remove the selected password or clear the list"
    ElseIf Index = 3 Then
        lblStatus = "Click here for some setting"
    ElseIf Index = 4 Then
        lblStatus = "Click to minimize the window."
    ElseIf Index = 5 Then
        lblStatus = "Click to exit the program."
    End If
End Sub

Private Sub lblGen_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblGen(Index).BackColor = vbRed
lblGen(Index).ForeColor = vbWhite
End Sub

Private Sub lblLength_Change()
length = Val(lblLength.Caption)
End Sub

Private Sub lblMin_Click()
Me.WindowState = 1
End Sub

Private Sub lst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte
For i = 0 To 5
lblGen(i).BackColor = vbWhite
lblGen(i).ForeColor = vbBlack
Next

lblStatus = "Select a number you want to copy or remove!"
End Sub

Public Sub SaveSetting()

If chkUpper.Value = 1 Then
    upper = True
Else
    upper = False
End If

If chkLower.Value = 1 Then
    lower = True
Else
    lower = False
End If

If chkNum.Value = 1 Then
    Numeric = True
Else
    Numeric = False
End If

If chkSym.Value = 1 Then
    Sym = True
Else
    Sym = False
End If

Open "PassGen.ini" For Output As #1
    Write #1, upper, lower, Numeric, Sym, length
Close #1
End Sub

Private Sub Timer1_Timer()
'show the option frame
    fraMain.Left = fraMain.Left - 50
    fraSetting.Left = fraSetting.Left - 50
    If fraSetting.Left <= 120 Then
        Timer1.Enabled = False
    End If
End Sub

Private Sub Timer2_Timer()
'show the generate form
    fraMain.Left = fraMain.Left + 50
    fraSetting.Left = fraSetting.Left + 50
    If fraMain.Left >= 120 Then
        Timer2.Enabled = False
    End If
End Sub

Public Sub Validate()
    If chkUpper.Value = 0 And chkLower.Value = 0 And chkNum.Value = 0 And chkSym.Value = 0 Then
        chkLower.Value = 1
    End If
End Sub

Private Sub min()
Dim j As Integer

For j = 0 To 100
    Me.Top = Me.Top + 200
    Me.Width = Me.Width - 41
    Me.Height = Me.Height - 41
Next

Me.Height = 4215
Me.Width = 4815
Me.Top = 1800
Me.Left = 5580

End Sub
