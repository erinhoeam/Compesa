Attribute VB_Name = "mForms"
Public Declare Function CreateRoundRectRgn Lib _
        "gdi32" (ByVal X1 As Long, ByVal Y1 As _
        Long, ByVal X2 As Long, ByVal Y2 As Long, _
        ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" _
        (ByVal hwnd As Long, ByVal hRgn As Long, _
        ByVal bRedraw As Boolean) As Long
Public Declare Function GetClientRect Lib "user32" _
        (ByVal hwnd As Long, lpRect As Rect) As Long
Public Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
' Declaração de API's necessários
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
            ByVal hWndInsertAfter As Long, _
            ByVal x As Long, _
            ByVal y As Long, _
            ByVal cx As Long, _
            ByVal cy As Long, _
            ByVal wFlags As Long) As Long

' Definição de constantes

Public Const GWL_EXSTYLE = (-20)
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2
Public Const ULW_OPAQUE = &H4
Public Const WS_EX_LAYERED = &H80000

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
         As Long
         If Topmost = True Then 'Make the window topmost
            SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, _
               0, FLAGS)
         Else
            SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, _
               0, 0, FLAGS)
            SetTopMostWindow = False
         End If
End Function
' Define o form como transparente
Public Sub MakeTransparent(ByVal hwnd As Long, ByVal bAlpha As Integer)

        Dim msg As Long

      

        ' Ignora possíveis erros

        On Error Resume Next

      

        ' Caso o valor seja inferior a 255 e superior

        ' a 0 aplica uma nova transparência

        If bAlpha > 0 Or bAlpha < 255 Then

      

            msg = GetWindowLong(hwnd, GWL_EXSTYLE)

            msg = msg Or WS_EX_LAYERED

            SetWindowLong hwnd, GWL_EXSTYLE, msg

      

            SetLayeredWindowAttributes hwnd, 0, bAlpha, LWA_ALPHA

        End If

      

    End Sub

      

      

    ' Define o form com opaco

    Public Sub MakeOpaque(ByVal hwnd As Long)

        Dim msg As Long

      

        ' Ignora possíveis erros

        On Error Resume Next

      

        msg = GetWindowLong(hwnd, GWL_EXSTYLE)

        msg = msg And Not WS_EX_LAYERED

        SetWindowLong hwnd, GWL_EXSTYLE, msg

        SetLayeredWindowAttributes hwnd, 0, 0, LWA_ALPHA

      

    End Sub

Public Sub Retangulo(m_hWnd As Long, Fator As Byte)
  Dim RGN As Long
  Dim RC As Rect
  Call GetClientRect(m_hWnd, RC)
  RGN = CreateRoundRectRgn(RC.Left, RC.Top, RC.Right, _
                           RC.Bottom, Fator, Fator)
  SetWindowRgn m_hWnd, RGN, True
End Sub
'Fator é a distância da curvatura do canto arredondado


