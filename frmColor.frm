VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColor 
   BackColor       =   &H80000003&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Arrow Color"
   ClientHeight    =   2160
   ClientLeft      =   3810
   ClientTop       =   3435
   ClientWidth     =   4605
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmColor.frx":0000
   ScaleHeight     =   2160
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   780
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3270
      TabIndex        =   7
      Top             =   1770
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   2220
      TabIndex        =   6
      Top             =   1770
      Width           =   915
   End
   Begin VB.PictureBox SColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3750
      ScaleHeight     =   255
      ScaleWidth      =   525
      TabIndex        =   5
      Top             =   1110
      Width           =   555
   End
   Begin VB.PictureBox MColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3750
      ScaleHeight     =   255
      ScaleWidth      =   525
      TabIndex        =   4
      Top             =   630
      Width           =   555
   End
   Begin VB.PictureBox HColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3750
      ScaleHeight     =   255
      ScaleWidth      =   525
      TabIndex        =   3
      Top             =   180
      Width           =   555
   End
   Begin VB.Line Line4 
      X1              =   3000
      X2              =   3750
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line3 
      X1              =   2700
      X2              =   3750
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Line Line2 
      X1              =   2370
      X2              =   3750
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SColor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2010
      TabIndex        =   2
      Top             =   1110
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "MColor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   630
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "HColor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1290
      TabIndex        =   0
      Top             =   150
      Width           =   855
   End
   Begin VB.Shape Shape3 
      DrawMode        =   6  'Mask Pen Not
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   1170
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   1845
   End
   Begin VB.Shape Shape2 
      DrawMode        =   6  'Mask Pen Not
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   870
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1845
   End
   Begin VB.Shape Shape1 
      DrawMode        =   6  'Mask Pen Not
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   570
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1845
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdOK_Click()
SaveString HKEY_CURRENT_USER, "Software\Caspra", "HColor", HColor.BackColor
SaveString HKEY_CURRENT_USER, "Software\Caspra", "MColor", MColor.BackColor
SaveString HKEY_CURRENT_USER, "Software\Caspra", "SColor", SColor.BackColor
frmMain.Line1.BorderColor = HColor.BackColor
frmMain.Line2.BorderColor = MColor.BackColor
frmMain.Line3.BorderColor = SColor.BackColor
Unload Me
End Sub
Private Sub Form_Load()
HColor.BackColor = frmMain.Line1.BorderColor
MColor.BackColor = frmMain.Line2.BorderColor
SColor.BackColor = frmMain.Line3.BorderColor
End Sub

Private Sub HColor_DblClick()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler
CommonDialog1.Flags = cdlCCRGBInit
CommonDialog1.ShowColor
HColor.BackColor = CommonDialog1.Color
Exit Sub
ErrHandler:
Exit Sub

End Sub

Private Sub MColor_DblClick()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler
CommonDialog1.Flags = cdlCCRGBInit
CommonDialog1.ShowColor
MColor.BackColor = CommonDialog1.Color
Exit Sub
ErrHandler:
Exit Sub
End Sub

Private Sub SColor_DblClick()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler
CommonDialog1.Flags = cdlCCRGBInit
CommonDialog1.ShowColor
SColor.BackColor = CommonDialog1.Color
Exit Sub
ErrHandler:
Exit Sub
End Sub
