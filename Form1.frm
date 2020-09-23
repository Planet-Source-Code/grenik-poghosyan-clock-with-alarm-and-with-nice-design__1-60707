VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   1320
   ClientLeft      =   5655
   ClientTop       =   180
   ClientWidth     =   1725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   1725
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1290
      Top             =   510
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1290
      Top             =   30
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   1245
      TabIndex        =   0
      Top             =   0
      Width           =   1245
      Begin VB.Line Line3 
         X1              =   630
         X2              =   630
         Y1              =   630
         Y2              =   1020
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   630
         X2              =   975
         Y1              =   630
         Y2              =   645
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   630
         X2              =   630
         Y1              =   630
         Y2              =   360
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "AG_Souvenir"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   360
         TabIndex        =   1
         Top             =   420
         Width           =   585
      End
      Begin VB.Label lblSecond 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "AG_Souvenir"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   420
         TabIndex        =   2
         Top             =   720
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Const VK_SCROLL = &H91
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim cx As Single, cy As Single, dx As Single, dy As Single
Dim bDrag As Boolean
Dim min As Integer
Dim Caspra As String
Dim HColor As String
Dim MColor As String
Dim SColor As String
Dim Message As String
Dim Music As String
Dim Program As String
Dim Skin As String
Private Sub Form_Load()
On Error Resume Next
MkDir App.Path & "\Skins"
MoveFile App.Path & "\Aqua.ico", App.Path & "\Skins\Aqua.ico"
MoveFile App.Path & "\Matrix.ico", App.Path & "\Skins\Matrix.ico"
MoveFile App.Path & "\DcClock.bmp", App.Path & "\Skins\DcClock.bmp"
Skin = GetString(HKEY_CURRENT_USER, "Software\Caspra", "Skins")
Me.Left = GetString(HKEY_CURRENT_USER, "Software\Caspra", "x")
Me.Top = GetString(HKEY_CURRENT_USER, "Software\Caspra", "y")
Caspra = GetString(HKEY_CURRENT_USER, "Software\Caspra", "Caspra Clock")
HColor = GetString(HKEY_CURRENT_USER, "Software\Caspra", "HColor")
MColor = GetString(HKEY_CURRENT_USER, "Software\Caspra", "MColor")
SColor = GetString(HKEY_CURRENT_USER, "Software\Caspra", "SColor")
If Skin = "Matrix" Then
pic.Picture = LoadPicture(App.Path & "\Skins\matrix.ico")
frmMenu.Matrix.Checked = True
End If
If Skin = "Aqua" Then
pic.Picture = LoadPicture(App.Path & "\Skins\aqua.ico")
frmMenu.Aqua.Checked = True
End If
If Skin = "Digital" Then
pic.Picture = LoadPicture(App.Path & "\Skins\DCClock.bmp")
frmMenu.Digital.Checked = True
End If
If Skin = "" Then
pic.Picture = LoadPicture(App.Path & "\Skins\DCClock.bmp")
frmMenu.Digital.Checked = True
End If
If HColor <> "" Then
Line1.BorderColor = HColor
Line2.BorderColor = MColor
Line3.BorderColor = SColor
Else
Line1.BorderColor = vbBlack
Line2.BorderColor = vbBlack
Line3.BorderColor = vbRed
End If
If Caspra = "" Then
SaveKey HKEY_CURRENT_USER, "Software\Caspra"
SaveKey HKEY_CURRENT_USER, "Software\Caspra\Alarm"
SaveKey HKEY_CURRENT_USER, "Software\Caspra\Alarm\Shell"
SaveKey HKEY_CURRENT_USER, "Software\Caspra\Alarm\Play"
SaveString HKEY_CURRENT_USER, "Software\Caspra", "Caspra Clock", App.Path & "\Clock.exe"
End If
App.TaskVisible = False
On Error Resume Next
Dim St As String, c, Cc
pic.Move 0, 0
For c = 0 To 2
Cc = Cc + (pic.Width \ 100) * 10
Next c
Main1
lblTime = Hour(Time) & ":" & Minute(Time)
lblSecond = Second(Time)
Timer1_Timer
End Sub

Private Sub lblSecond_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu frmMenu.File
End If
If Button = 1 Then
bDrag = True
dx = X: dy = Y: cx = Me.Left: cy = Me.Top
End If
End Sub

Private Sub lblSecond_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bDrag Then
cx = cx + X - dx
cy = cy + Y - dy
Me.Move cx, cy
SaveString HKEY_CURRENT_USER, "Software\Caspra", "X", Me.Left
SaveString HKEY_CURRENT_USER, "Software\Caspra", "y", Me.Top
End If

End Sub

Private Sub lblSecond_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
bDrag = False
End Sub

Private Sub lblTime_Change()
Message = GetString(HKEY_CURRENT_USER, "Software\Caspra\Alarm", lblTime)
Music = GetString(HKEY_CURRENT_USER, "Software\Caspra\Alarm\Play", lblTime)
Program = GetString(HKEY_CURRENT_USER, "Software\Caspra\Alarm\Shell", lblTime)
If Message <> "" Then
MsgBox Message, vbOKOnl, "Attention"
End If
If Program <> "" Then
On Error Resume Next
ShellExecute Me.hWnd, "Open", Program, vbNullString, vbNullString, 1
End If
If Music <> "" Then
On Error Resume Next
ShellExecute Me.hWnd, "Open", Music, vbNullString, vbNullString, 1
End If
End Sub

Private Sub lblTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu frmMenu.File
End If
If Button = 1 Then
bDrag = True
dx = X: dy = Y: cx = Me.Left: cy = Me.Top
End If

End Sub

Private Sub lblTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bDrag Then
cx = cx + X - dx
cy = cy + Y - dy
Me.Move cx, cy
SaveString HKEY_CURRENT_USER, "Software\Caspra", "X", Me.Left
SaveString HKEY_CURRENT_USER, "Software\Caspra", "y", Me.Top
End If

End Sub

Private Sub lblTime_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
bDrag = False
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu frmMenu.File
End If
If Button = 1 Then
bDrag = True
dx = X: dy = Y: cx = Me.Left: cy = Me.Top
End If
End Sub
Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bDrag Then
cx = cx + X - dx
cy = cy + Y - dy
Me.Move cx, cy
SaveString HKEY_CURRENT_USER, "Software\Caspra", "X", Me.Left
SaveString HKEY_CURRENT_USER, "Software\Caspra", "y", Me.Top
End If
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
bDrag = False
End Sub
Private Sub Timer1_Timer()
frmMenu.Date.Caption = "Date: " & Date
frmMenu.Time.Caption = "Time: " & Time
min = Minute(Time)
If min < 10 Then
lblTime.Caption = Hour(Time) & ":0" & Minute(Time)
Else
lblTime.Caption = Hour(Time) & ":" & Minute(Time)
End If
If lblSecond.Caption < 9 Then
lblSecond.Caption = "0" & Second(Time)
Else
lblSecond = Second(Time)
End If
m = Hour(Time) * 30 + 270
l = ((Line1.X1 - Line1.X2) ^ 2 + (Line1.Y1 - Line1.Y2) ^ 2) ^ (1 / 2)
Line1.X2 = Line1.X1 + l * Cos(m * 3.14 / 180)
Line1.Y2 = Line1.Y1 + l * Sin(m * 3.14 / 180)
m = Minute(Time) * 6 + 270
l = ((Line2.X1 - Line2.X2) ^ 2 + (Line2.Y1 - Line2.Y2) ^ 2) ^ (1 / 2)
Line2.X2 = Line2.X1 + l * Cos(m * 3.14 / 180)
Line2.Y2 = Line2.Y1 + l * Sin(m * 3.14 / 180)
m = Second(Time) * 6 + 270
l = ((Line3.X1 - Line3.X2) ^ 2 + (Line3.Y1 - Line3.Y2) ^ 2) ^ (1 / 2)
Line3.X2 = Line3.X1 + l * Cos(m * 3.14 / 180)
Line3.Y2 = Line3.Y1 + l * Sin(m * 3.14 / 180)
End Sub
Private Sub Timer2_Timer()
If GetAsyncKeyState(VK_SCROLL) Then
Me.Visible = True
Timer2.Enabled = False
End If
End Sub
