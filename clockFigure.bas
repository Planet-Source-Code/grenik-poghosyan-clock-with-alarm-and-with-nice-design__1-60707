Attribute VB_Name = "Module1"
Public Const RGN_COPY = 5
Public Const RGN_AND = 1
Public Const RGN_XOR = 3
Public Const RGN_OR = 2

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Private Const HWND_BOTTOM = 1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOP = 0

Public StartColor As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Function lGetRegion(pic As PictureBox, lBackColor As Long) As Long
Dim lRgn As Long
Dim lSkinRgn As Long
Dim lStart As Long
Dim lX As Long
Dim lY As Long
Dim lHeight As Long
Dim lWidth As Long


lSkinRgn = CreateRectRgn(0, 0, 0, 0)

With pic

    lHeight = .Height / Screen.TwipsPerPixelY
    lWidth = .Width / Screen.TwipsPerPixelX
    For lX = 0 To lHeight - 1
        lY = 0
        Do While lY < lWidth

            Do While lY < lWidth And GetPixel(.hdc, lY, lX) = lBackColor
                lY = lY + 1
            Loop

            If lY < lWidth Then
                lStart = lY
                Do While lY < lWidth And GetPixel(.hdc, lY, lX) <> lBackColor
                    lY = lY + 1
                Loop
                If lY > lWidth Then lY = lWidth

                lRgn = CreateRectRgn(lStart, lX, lY, lX + 1)
                CombineRgn lSkinRgn, lSkinRgn, lRgn, RGN_OR
                DeleteObject lRgn
            End If
        Loop
    Next
End With

lGetRegion = lSkinRgn
End Function


Public Sub SetFormPosition(hWnd As Long, TopPosition As Boolean)
    If TopPosition Then
         SetWindowPos frmMain.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
                      SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
     Else
         SetWindowPos frmMain.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, _
                      SWP_NOSIZE Or SWP_NOMOVE
     End If
     
End Sub

Sub Main1()
Dim lRgn As Long
StartColor = GetPixel(frmMain.pic.hdc, 0, 0)
    'Load frmclock
    lRgn = lGetRegion(frmMain.pic, StartColor)
    SetWindowRgn frmMain.hWnd, lRgn, True
    DeleteObject lRgn
    'frmclock.Show
    
End Sub


Sub Main()
    If App.PrevInstance = True Then
        End
    Else
        frmFunny.Show
    End If
End Sub
