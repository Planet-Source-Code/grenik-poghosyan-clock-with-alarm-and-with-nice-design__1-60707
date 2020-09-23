VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caspra Clock"
   ClientHeight    =   1785
   ClientLeft      =   150
   ClientTop       =   765
   ClientWidth     =   1950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   1950
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Date 
         Caption         =   "Date"
         Enabled         =   0   'False
      End
      Begin VB.Menu Time 
         Caption         =   "Time"
         Enabled         =   0   'False
      End
      Begin VB.Menu Alarm 
         Caption         =   "Alarm"
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu AlwaysOn 
         Caption         =   "Always On Top"
      End
      Begin VB.Menu RunOn 
         Caption         =   "Run On StartUp"
      End
      Begin VB.Menu Line2 
         Caption         =   "-"
      End
      Begin VB.Menu Skins 
         Caption         =   "Skins"
         Begin VB.Menu Digital 
            Caption         =   "Digital"
         End
         Begin VB.Menu Aqua 
            Caption         =   "Aqua"
         End
         Begin VB.Menu Matrix 
            Caption         =   "Matrix"
         End
      End
      Begin VB.Menu ArrowC 
         Caption         =   "Arrow Color"
      End
      Begin VB.Menu Line3 
         Caption         =   "-"
      End
      Begin VB.Menu SetupD 
         Caption         =   "Setup Date && Time"
      End
      Begin VB.Menu Line4 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "About..."
      End
      Begin VB.Menu Hide 
         Caption         =   "Hide           ScrollLock"
      End
      Begin VB.Menu Close 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StartUp As String
Dim TopMost As String

Private Sub About_Click()
MsgBox "Caspra Clock Version 1.5" & vbCrLf & "This program was created by Grenik Poghosyan", vbOKOnly + vbInformation, "Attention"
End Sub

Private Sub Alarm_Click()
frmAlarm.Show 1
End Sub
Private Sub AlwaysOn_Click()
If AlwaysOn.Checked = True Then
SaveString HKEY_CURRENT_USER, "Software\Caspra", "TopMost", "0"
AlwaysOn.Checked = False
SetFormPosition frmMain.hWnd, False
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra", "TopMost", "1"
AlwaysOn.Checked = True
SetFormPosition frmMain.hWnd, True
End If
End Sub

Private Sub Aqua_Click()
SaveString HKEY_CURRENT_USER, "Software\Caspra", "Skins", "Aqua"
Digital.Checked = False
Matrix.Checked = False
Aqua.Checked = True
frmMain.pic.Picture = LoadPicture(App.Path & "\Skins\aqua.ico")
End Sub

Private Sub ArrowC_Click()
frmColor.Show 1
End Sub

Private Sub Close_Click()
Unload frmMain
Unload Me
End Sub

Private Sub Digital_Click()
SaveString HKEY_CURRENT_USER, "Software\Caspra", "Skins", "Digital"
Digital.Checked = True
Matrix.Checked = False
Aqua.Checked = False
frmMain.pic.Picture = LoadPicture(App.Path & "\Skins\DCClock.bmp")
End Sub

Private Sub Form_Load()
TopMost = GetString(HKEY_CURRENT_USER, "Software\Caspra", "TopMost")
StartUp = GetString(HKEY_CURRENT_USER, "Software\Caspra", "StartUp")
If StartUp = "1" Then
RunOn.Checked = True
Else
RunOn.Checked = False
End If
If TopMost = "1" Then
AlwaysOn.Checked = True
SetFormPosition frmMain.hWnd, True
Else
AlwaysOn.Checked = False
SetFormPosition frmMain.hWnd, False
End If
End Sub

Private Sub Hide_Click()
frmMain.Visible = False
frmMain.Timer2.Enabled = True
End Sub

Private Sub Matrix_Click()
SaveString HKEY_CURRENT_USER, "Software\Caspra", "Skins", "Matrix"
Digital.Checked = False
Matrix.Checked = True
Aqua.Checked = False
frmMain.pic.Picture = LoadPicture(App.Path & "\Skins\Matrix.ico")
End Sub

Private Sub RunOn_Click()
If RunOn.Checked = True Then
SaveString HKEY_CURRENT_USER, "Software\Caspra", "StartUp", "0"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Caspra Clock"
RunOn.Checked = False
Else
SaveString HKEY_CURRENT_USER, "Software\Caspra", "StartUp", "1"
SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Caspra Clock", App.Path & "\Clock.exe"
RunOn.Checked = True
End If
End Sub
Private Sub SetupD_Click()
On Error Resume Next
Call Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", vbNormalFocus)
End Sub
