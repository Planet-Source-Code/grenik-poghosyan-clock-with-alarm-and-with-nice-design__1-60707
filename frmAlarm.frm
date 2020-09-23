VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAlarm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alarm"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAlarm.frx":0000
   ScaleHeight     =   3690
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4890
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowseP 
      Caption         =   "..."
      Height          =   285
      Left            =   4890
      TabIndex        =   15
      Top             =   1170
      Width           =   465
   End
   Begin VB.CommandButton cmdBrowseM 
      Caption         =   "..."
      Height          =   285
      Left            =   4890
      TabIndex        =   14
      Top             =   780
      Width           =   465
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   315
      Index           =   0
      Left            =   210
      TabIndex        =   13
      Top             =   3330
      Width           =   855
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   210
      TabIndex        =   12
      Top             =   1530
      Width           =   4605
   End
   Begin VB.TextBox txtMinute 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      MaxLength       =   2
      TabIndex        =   10
      Top             =   60
      Width           =   315
   End
   Begin VB.TextBox txtHour 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3990
      MaxLength       =   2
      TabIndex        =   9
      Top             =   60
      Width           =   315
   End
   Begin VB.TextBox txtProgram 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2310
      TabIndex        =   7
      Top             =   1170
      Width           =   2535
   End
   Begin VB.TextBox txtMusic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2310
      TabIndex        =   6
      Top             =   780
      Width           =   2535
   End
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2310
      TabIndex        =   5
      Top             =   390
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   315
      Left            =   3690
      TabIndex        =   4
      Top             =   3330
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   2370
      TabIndex        =   3
      Top             =   3330
      Width           =   1125
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   195
      Left            =   4350
      TabIndex        =   11
      Top             =   90
      Width           =   135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      Height          =   225
      Left            =   3570
      TabIndex        =   8
      Top             =   90
      Width           =   465
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Shell Program"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   990
      TabIndex        =   2
      Top             =   1170
      Width           =   3825
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Set Music"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   990
      TabIndex        =   1
      Top             =   780
      Width           =   3825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Set Message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1050
      TabIndex        =   0
      Top             =   420
      Width           =   3825
   End
   Begin VB.Shape Shape3 
      DrawMode        =   10  'Mask Pen
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   900
      Top             =   1170
      Width           =   3945
   End
   Begin VB.Shape Shape2 
      DrawMode        =   10  'Mask Pen
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   900
      Top             =   780
      Width           =   3945
   End
   Begin VB.Shape Shape1 
      DrawMode        =   10  'Mask Pen
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   900
      Top             =   390
      Width           =   3945
   End
End
Attribute VB_Name = "frmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Dim ListS As Integer
Dim Selected As String
Dim H
Dim M
Private Sub cmdBrowseM_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler
CommonDialog1.FilterIndex = 2
CommonDialog1.Filter = "Music Files (*.wav,*.wma,*.cda,*.mid,*.midi,*.mp3)|*.wav;*.wma;*.cda;*.mid;*.midi;*.mp3|"
CommonDialog1.ShowOpen
txtMusic.Text = CommonDialog1.FileName
Exit Sub
ErrHandler:
Exit Sub
End Sub

Private Sub cmdBrowseP_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler
CommonDialog1.FilterIndex = 2
CommonDialog1.Filter = "Programs (*.exe)|*.exe|"
CommonDialog1.ShowOpen
txtProgram.Text = CommonDialog1.FileName
Exit Sub
ErrHandler:
Exit Sub

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdOK_Click()
On Error Resume Next
If txtHour.Text <> "" And txtHour.Text <= "24" And txtMinute.Text < "60" And txtMinute.Text <> "" Then
If txtMessage.Text <> "" Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Alarm", txtHour & ":" & txtMinute, txtMessage
End If
If txtMusic.Text <> "" Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Alarm\Play", txtHour & ":" & txtMinute, txtMusic
End If
If txtProgram.Text <> "" Then
SaveString HKEY_CURRENT_USER, "Software\Caspra\Alarm\Shell", txtHour & ":" & txtMinute, txtProgram
End If
List1.AddItem txtHour.Text & ":" & txtMinute
n = n + 1
SaveSetting "txt", "txt", "count", Str(n)
For i = 1 To n
SaveSetting "txt", "txt", "item" & i, List1.List(i - 1)
Next
Else
MsgBox "You must Type time in Currect form", vbOKOnly + vbInformation, "Attention"
End If

End Sub
Private Sub cmdRemove_Click(Index As Integer)
On Error Resume Next
ListS = List1.SelCount
If ListS > 0 Then
List1.RemoveItem List1.ListIndex
n = n - 1
DeleteValue HKEY_CURRENT_USER, "Software\Caspra\Alarm", Selected
DeleteValue HKEY_CURRENT_USER, "Software\Caspra\Alarm\Play", Selected
DeleteValue HKEY_CURRENT_USER, "Software\Caspra\Alarm\Shell", Selected
SaveSetting "txt", "txt", "count", Str(n)
For i = 1 To n
SaveSetting "txt", "txt", "item" & i, List1.List(i - 1)
Next
End If
txtMessage.Text = ""
txtMusic.Text = ""
txtProgram.Text = ""
End Sub

Private Sub Form_Load()
On Error Resume Next
If GetSetting("txt", "txt", "count") = "" Then
n = 0
Else
n = Int(GetSetting("txt", "txt", "count"))
End If
For i = 1 To n
List1.AddItem GetSetting("txt", "txt", "Item" & i)
Next
End Sub
Private Sub List1_Click()
Selected = List1.Text
txtMessage.Text = GetString(HKEY_CURRENT_USER, "Software\Caspra\Alarm", Selected)
txtMusic.Text = GetString(HKEY_CURRENT_USER, "Software\Caspra\Alarm\Play", Selected)
txtProgram.Text = GetString(HKEY_CURRENT_USER, "Software\Caspra\Alarm\Shell", Selected)
End Sub
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Selected = List1.Text
txtMessage.Text = GetString(HKEY_CURRENT_USER, "Software\Caspra\Alarm", Selected)
txtMusic.Text = GetString(HKEY_CURRENT_USER, "Software\Caspra\Alarm\Play", Selected)
txtProgram.Text = GetString(HKEY_CURRENT_USER, "Software\Caspra\Alarm\Shell", Selected)
End Sub
Private Sub List1_KeyPress(KeyAscii As Integer)
Selected = List1.Text
txtMessage.Text = GetString(HKEY_CURRENT_USER, "Software\Caspra\Alarm", Selected)
txtMusic.Text = GetString(HKEY_CURRENT_USER, "Software\Caspra\Alarm\Play", Selected)
txtProgram.Text = GetString(HKEY_CURRENT_USER, "Software\Caspra\Alarm\Shell", Selected)
End Sub
Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
Selected = List1.Text
txtMessage.Text = GetString(HKEY_CURRENT_USER, "Software\Caspra\Alarm", Selected)
txtMusic.Text = GetString(HKEY_CURRENT_USER, "Software\Caspra\Alarm\Play", Selected)
txtProgram.Text = GetString(HKEY_CURRENT_USER, "Software\Caspra\Alarm\Shell", Selected)
End Sub
Private Sub txtHour_Change()
H = txtHour.Text
If txtHour.Text <> "" Then
If IsNumeric(txtHour.Text) = False Then
txtHour.Text = ""
End If
If H > 23 Then
txtHour.Text = ""
End If
End If
End Sub
Private Sub txtMinute_Change()
M = txtMinute.Text
If txtMinute.Text <> "" Then
If IsNumeric(txtMinute.Text) = False Then
txtMinute.Text = ""
End If
If M > 59 Then
txtMinute.Text = ""
End If
End If
End Sub
