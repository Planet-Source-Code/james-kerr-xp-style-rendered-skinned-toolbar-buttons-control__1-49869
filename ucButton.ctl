VERSION 5.00
Begin VB.UserControl ucButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1755
   FillColor       =   &H80000007&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   825
   ScaleWidth      =   1755
   ToolboxBitmap   =   "ucButton.ctx":0000
   Begin VB.Shape shBord 
      BorderColor     =   &H80000015&
      Height          =   420
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   135
      Width           =   465
   End
   Begin VB.Label lblDD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1395
      TabIndex        =   1
      Top             =   315
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   495
      TabIndex        =   0
      Top             =   180
      Width           =   45
   End
End
Attribute VB_Name = "ucButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const m_def_Caption               As String = ""
Private m_Caption                         As String
Public Event Click()
Public Enum eCaptionPos
    CaptionRight = 0
    CaptionBottom = 1
End Enum
Private Enum eButtonState
    Normal = 0
    Over = 1
    Pressed = 2
    Disabled = 4
End Enum
Private MyState                           As eButtonState
Public Enum eButtonStyle
    Standard = 0
    Check = 1
    DropDown = 2
End Enum
Private m_Hwnd                            As Long
Private bClosing                          As Boolean
Private Const WM_MOUSEMOVE                As Long = &H200
Private Const WM_MOUSEHOVER               As Long = &H2A1
Private Const WM_MOUSELEAVE               As Long = &H2A3
Private Const WM_LBUTTONDOWN              As Long = &H201
Private Const WM_LBUTTONUP                As Long = &H202
Private Const TME_HOVER                   As Long = &H1
Private Const TME_LEAVE                   As Long = &H2
Private Const TME_QUERY                   As Long = &H40000000
Private Const TME_CANCEL                  As Long = &H80000000
Private Const HOVER_DEFAULT               As Long = &HC00000
Private Type tagTRACKMOUSEEVENT
    cbSize                                  As Long
    dwFlags                                 As Long
    hwndTrack                               As Long
    dwHoverTime                             As Long
End Type
Private bTracking                         As Boolean
Implements ISubclass
Private Const m_def_ImageHeight           As Integer = 16
Private Const m_def_Enabled               As Boolean = True
Private Const m_def_CaptionPosition       As Integer = 0
Private m_ImageHeight                     As Integer
Private m_ButtonType                      As Integer
Private m_Enabled                         As Boolean
Private m_ButtonPicture                   As StdPicture
Private m_Font                            As Font
Private m_CaptionPosition                 As eCaptionPos
Private Const DST_COMPLEX                 As Long = &H0
Private Const DST_TEXT                    As Long = &H1
Private Const DST_PREFIXTEXT              As Long = &H2
Private Const DST_ICON                    As Long = &H3
Private Const DST_BITMAP                  As Long = &H4
Private Const DSS_NORMAL                  As Long = &H0
Private Const DSS_UNION                   As Long = &H10
Private Const DSS_DISABLED                As Long = &H20
Private Const DSS_MONO                    As Long = &H80
Private Const DSS_RIGHT                   As Long = &H8000
Private Const CLR_INVALID                 As Integer = -1
Private Const m_def_Checked               As Boolean = False
Private Const m_def_ButtonStyle           As Integer = 0
Private m_Checked                         As Boolean
Private m_ButtonStyle                     As Integer
Private Type RECT
    Left                                    As Long
    Top                                     As Long
    Right                                   As Long
    Bottom                                  As Long
End Type
Private Type POINTAPI
    x                                       As Long
    y                                       As Long
End Type
Private Type PointSng
    x                                       As Double
    y                                       As Double
End Type
Private Type RectAPI
    Left                                    As Long
    Top                                     As Long
    Right                                   As Long
    Bottom                                  As Long
End Type
Private Const PS_SOLID                    As Long = 0
Private Const PI                          As Double = 3.14159265358979
Private Const RADS                        As Double = PI / 180
Private Const RGBMAX                      As Long = 255
Private Const HSLMAX                      As Long = RGBMAX
Private Type tHSL
    h                                       As Double
    s                                       As Double
    L                                       As Double
End Type
Private Type tRGB
    R                                       As Long
    g                                       As Long
    b                                       As Long
End Type
Private Enum GradBlendMode
    gbmRGB = 0
    gbmHSL = 1
End Enum
Private Enum GradType
    gtNormal = 0
    gtElliptical = 1
    gtRectangular = 2
End Enum
Private mlColor1                          As Long
Private mlColor2                          As Long
Private m_BackColour                      As OLE_COLOR
Private Const m_def_ForeColour            As Long = vbMenuText
Private m_ForeColour                      As OLE_COLOR
Private Const m_def_UseParentColour       As Boolean = False
Private m_UseParentColour                 As Boolean
Private Const m_def_OfficeXPStyle         As Boolean = False
Private Const m_def_HilightColour         As Integer = 0
Private m_OfficeXPStyle                   As Boolean
Private m_HilightColour                   As OLE_COLOR
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As tagTRACKMOUSEEVENT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
                                                 ByVal hWndNewParent As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, _
                                                                    ByVal hBrush As Long, _
                                                                    ByVal lpDrawStateProc As Long, _
                                                                    ByVal lParam As Long, _
                                                                    ByVal wParam As Long, _
                                                                    ByVal n1 As Long, _
                                                                    ByVal n2 As Long, _
                                                                    ByVal n3 As Long, _
                                                                    ByVal n4 As Long, _
                                                                    ByVal un As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
                                                               ByVal HPALETTE As Long, _
                                                               pccolorref As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
                                                  ByVal nCmdShow As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               ByVal crColor As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RectAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
                                             ByVal x As Long, _
                                             ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               lpPoint As POINTAPI) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long
Const m_def_HasSeparator = False
Dim m_HasSeparator As Boolean
Const m_def_BackColour = vbButtonFace
Const m_def_HiColourDisplay = True
Dim m_HiColourDisplay As Boolean




Private Sub Attach()

    m_Hwnd = UserControl.hwnd
    AttachMessage Me, m_Hwnd, WM_MOUSEMOVE
    AttachMessage Me, m_Hwnd, WM_MOUSELEAVE
    AttachMessage Me, m_Hwnd, WM_LBUTTONDOWN
    AttachMessage Me, m_Hwnd, WM_LBUTTONUP

End Sub

Private Function BlendColors(ByVal lColor1 As Long, _
                             ByVal lColor2 As Long, _
                             ByVal lSteps As Long, _
                             ByVal fRepetitions As Double, _
                             laRetColors() As Long) As Long

  Dim lIdx    As Long
  Dim lIdx2   As Long
  Dim lRed    As Long
  Dim lGrn    As Long
  Dim lBlu    As Long
  Dim fRedStp As Double
  Dim fGrnStp As Double
  Dim fBluStp As Double

    If lSteps < 2 Then
        lSteps = 2
    End If
    ReDim laRetColors(lSteps * 2)
    lRed = (lColor1 And &HFF&)
    lGrn = (lColor1 And &HFF00&) / &H100
    lBlu = (lColor1 And &HFF0000) / &H10000
    fRedStp = Div((lColor2 And &HFF&) - lRed, lSteps / fRepetitions)
    fGrnStp = Div(((lColor2 And &HFF00&) / &H100&) - lGrn, lSteps / fRepetitions)
    fBluStp = Div(((lColor2 And &HFF0000) / &H10000) - lBlu, lSteps / fRepetitions)
    laRetColors(0) = lColor1    'First Color
    laRetColors(Int(lSteps / fRepetitions)) = lColor2        'Last Color
    laRetColors(Int(lSteps / fRepetitions) + 1) = lColor2    'Last Color
    For lIdx = 1 To Int(lSteps / fRepetitions) - 1           'All Colors between
        laRetColors(lIdx) = CLng(lRed + (fRedStp * lIdx)) + (CLng(lGrn + (fGrnStp * lIdx)) * &H100&) + (CLng(lBlu + (fBluStp * lIdx)) * &H10000)
    Next lIdx
    If Int(fRepetitions) >= 1 Then
        For lIdx2 = 1 To Int(fRepetitions) + 1
            If lIdx2 / 2 = Int(lIdx2 / 2) Then
                For lIdx = 0 To Int(lSteps / fRepetitions)
                    laRetColors(((lIdx2 - 1) * Int(lSteps / fRepetitions)) + lIdx) = laRetColors((lSteps / fRepetitions) - lIdx)
                Next lIdx
             Else 'NOT LIDX2...
                For lIdx = 0 To Int(lSteps / fRepetitions)
                    laRetColors(((lIdx2 - 1) * Int(lSteps / fRepetitions)) + lIdx) = laRetColors(lIdx)
                Next lIdx
            End If
        Next lIdx2
    End If
    BlendColors = lSteps

End Function

Private Function BlendColour(ByVal oColorFrom As OLE_COLOR, _
                             ByVal oColorTo As OLE_COLOR, _
                             Optional ByVal alpha As Long = 128) As Long

  Dim lCFrom As Long
  Dim lCTo   As Long
  Dim lSrcR  As Long
  Dim lSrcG  As Long
  Dim lSrcB  As Long
  Dim lDstR  As Long
  Dim lDstG  As Long
  Dim lDstB  As Long

    On Local Error Resume Next
    lCFrom = TranslateColour(oColorFrom)
    lCTo = TranslateColour(oColorTo)
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
    BlendColour = RGB(((lSrcR * alpha) / 255) + ((lDstR * (255 - alpha)) / 255), ((lSrcG * alpha) / 255) + ((lDstG * (255 - alpha)) / 255), ((lSrcB * alpha) / 255) + ((lDstB * (255 - alpha)) / 255))

End Function

Public Property Get ButtonPicture() As Picture

    Set ButtonPicture = m_ButtonPicture
    'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
    'MemberInfo=11,0,0,0

End Property

Public Property Set ButtonPicture(ByVal New_ButtonPicture As Picture)

    Set m_ButtonPicture = New_ButtonPicture
    PropertyChanged "ButtonPicture"
    Call DrawButton(5)

End Property

Public Property Get ButtonStyle() As eButtonStyle

    ButtonStyle = m_ButtonStyle
    'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
    'MemberInfo=7,0,0,0

End Property

Public Property Let ButtonStyle(ByVal New_ButtonStyle As eButtonStyle)

    m_ButtonStyle = New_ButtonStyle
    PropertyChanged "ButtonStyle"
    Call DrawButton(5)

End Property

Public Property Get Caption() As String

    Caption = m_Caption
    'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
    'MemberInfo=13,0,0,

End Property

Public Property Let Caption(ByVal New_Caption As String)

    m_Caption = New_Caption
    PropertyChanged "Caption"
    Call DrawButton(5)

End Property

Public Property Get CaptionPosition() As eCaptionPos

    CaptionPosition = m_CaptionPosition
    'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
    'MemberInfo=7,0,0,0

End Property

Public Property Let CaptionPosition(ByVal New_CaptionPosition As eCaptionPos)

    m_CaptionPosition = New_CaptionPosition
    PropertyChanged "CaptionPosition"
    Call DrawButton(5)

End Property

Public Property Get Checked() As Boolean

    Checked = m_Checked
    'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
    'MemberInfo=0,0,0,False

End Property

Public Property Let Checked(ByVal New_Checked As Boolean)

    m_Checked = New_Checked
    PropertyChanged "Checked"
    Call DrawButton(5)

End Property

Private Sub Detach()

    If Not m_Hwnd = 0 Then
        DetachMessage Me, m_Hwnd, WM_MOUSEMOVE
        DetachMessage Me, m_Hwnd, WM_MOUSELEAVE
        DetachMessage Me, m_Hwnd, WM_LBUTTONDOWN
        DetachMessage Me, m_Hwnd, WM_LBUTTONUP
        ShowWindow m_Hwnd, 0
        SetParent m_Hwnd, 0
        m_Hwnd = 0
    End If

End Sub

Private Function Div(ByVal dNumer As Double, _
                     ByVal dDenom As Double) As Double

    On Error Resume Next
    Div = dNumer / dDenom
    On Error GoTo 0

End Function

Private Function Draw(ByVal MyHwnd As Long, _
                      ByVal MyHDC As Long) As Boolean

  Dim lRet    As Long
  Dim lIdx    As Long
  Dim uRect   As RectAPI

    On Local Error Resume Next
    mlColor1 = TranslateColour(vbWhite)
    mlColor2 = TranslateColour(m_BackColour)
    lRet = GetClientRect(MyHwnd, uRect)
    If lRet <> 0 Then
        With uRect
            If .Right > 1 Then
                If .Bottom > 1 Then
                    lIdx = DrawGradient(MyHDC, .Right, .Bottom)
                    Draw = (lIdx > 0)
                End If
            End If
        End With 'uRect
    End If

End Function

Private Sub DrawButton(ToState As eButtonState)

  Dim XPos       As Long
  Dim YPos       As Long
  Dim lPicHeight As Long
  Dim lPicWidth  As Long
  Dim iAddWidth  As Integer
  Dim MyColour   As OLE_COLOR
  
    On Local Error Resume Next
    If ToState <> 5 Then
        If MyState = ToState Then
            Exit Sub
        End If
    End If
    If (m_ImageHeight = 24 Or m_OfficeXPStyle) Then
        lPicHeight = UserControl.ScaleY(m_ImageHeight, vbPixels, vbTwips) + 135
        If m_HasSeparator Then
            lPicWidth = UserControl.ScaleX(m_ImageHeight, vbPixels, vbTwips) + 180
            XPos = UserControl.ScaleX(90, vbTwips, vbPixels)
        Else
            lPicWidth = UserControl.ScaleX(m_ImageHeight, vbPixels, vbTwips) + 135
            XPos = UserControl.ScaleX(60, vbTwips, vbPixels)
        End If
        YPos = UserControl.ScaleY(45, vbTwips, vbPixels)
     Else 'NOT (M_IMAGEHEIGHT...
        lPicHeight = UserControl.ScaleY(m_ImageHeight, vbPixels, vbTwips) + 270
        lPicWidth = UserControl.ScaleX(m_ImageHeight, vbPixels, vbTwips) + 180
        XPos = UserControl.ScaleX(90, vbTwips, vbPixels)
        YPos = UserControl.ScaleY(90, vbTwips, vbPixels)
    End If
    UserControl.Cls
    'Separator
    With UserControl
        iAddWidth = 0
        If (m_UseParentColour Or m_OfficeXPStyle) Then
            .BackColor = vbButtonFace
         Else 'NOT (M_USEPARENTCOLOUR...
            .BackColor = m_BackColour
        End If
        If LenB(m_Caption) = 0 Then
            If m_ButtonStyle = 2 Then
                iAddWidth = iAddWidth + 235
            End If
            .Height = lPicHeight
            .Width = lPicWidth + iAddWidth
            If m_ButtonStyle = 2 Then
                lblDD.Move .ScaleWidth - 250, (.ScaleHeight - 195) / 2
                lblDD.ForeColor = m_ForeColour
                lblDD.Visible = True
             Else 'NOT M_BUTTONSTYLE...
                lblDD.Visible = False
            End If
         Else 'NOT LEN(M_CAPTION)...
            If m_ButtonStyle = 2 Then
                iAddWidth = iAddWidth + 180
            End If
            lblCaption.Caption = m_Caption
            Select Case m_CaptionPosition
             Case 0 'right
                .Height = lPicHeight
                .Width = lPicWidth + 45 + lblCaption.Width + iAddWidth
                lblCaption.Move lPicWidth - 30, (((.ScaleHeight - 45) - lblCaption.Height) / 2)
             Case 1 'bottom
                .Height = lPicHeight + lblCaption.Height
                If lblCaption.Width > lPicWidth Then
                    .Width = lblCaption.Width + 235 + iAddWidth
                 Else 'NOT LBLCAPTION.WIDTH...
                    .Width = lPicWidth + 235 + iAddWidth
                End If
                lblCaption.Move ((.ScaleWidth - lblCaption.Width) - iAddWidth) / 2, lPicHeight - 135
                XPos = .ScaleX(((.ScaleWidth - (lPicHeight - 270)) - iAddWidth) / 2, vbTwips, vbPixels)
            End Select
            If m_ButtonStyle = 2 Then
                lblDD.Move .ScaleWidth - 250, ((.ScaleHeight - 45) - 195) / 2
                lblDD.Visible = True
             Else 'NOT M_BUTTONSTYLE...
                lblDD.Visible = False
            End If
        End If
        shBord.Visible = False
        If Not m_Enabled Then
            MyState = Disabled
            If m_OfficeXPStyle Then
                lblCaption.ForeColor = vbButtonShadow
            Else
                If m_HiColourDisplay Then
                    lblCaption.ForeColor = BlendColour(m_ForeColour, m_BackColour, 70)
                Else
                    lblCaption.ForeColor = vb3DDKShadow
                End If
            End If
            If m_ButtonStyle = 2 Then
                lblDD.ForeColor = lblCaption.ForeColor
            End If
            Line (1, 1)-(.ScaleWidth, .ScaleHeight), .BackColor, BF
            If Not m_ButtonPicture Is Nothing Then
                Call DrawButtonImage(m_ButtonPicture, XPos, YPos, False, False)
            End If
         Else
            MyState = ToState
            lblCaption.ForeColor = vbMenuText
            If lblDD.Visible Then
                lblDD.ForeColor = vbMenuText
            End If
            If ToState = Over Then
                Draw .hwnd, .hdc
                lblCaption.ForeColor = m_HilightColour
                If lblDD.Visible Then
                    lblDD.ForeColor = m_HilightColour
                End If
                If m_OfficeXPStyle Then
                    If m_HiColourDisplay Then
                        MyColour = BlendColour(vbHighlight, vbWindowBackground, 70)
                    Else
                        MyColour = vbInfoBackground
                    End If
                    If m_HasSeparator Then
                        Line (1, 1)-(.ScaleWidth, .ScaleHeight), vbButtonFace, BF
                        Line (40, 1)-(.ScaleWidth, .ScaleHeight), MyColour, BF
                        Line (40, 1)-(.ScaleWidth - 20, .ScaleHeight - 20), vbHighlight, B
                    Else
                        Line (1, 1)-(.ScaleWidth, .ScaleHeight), MyColour, BF
                        Line (1, 1)-(.ScaleWidth - 20, .ScaleHeight - 20), vbHighlight, B
                    End If
                 Else 'M_OFFICEXPSTYLE = FALSE/0
                    Line (45, 1)-(.ScaleWidth - 45, 1), vb3DHighlight
                    Line (1, 45)-(1, .ScaleHeight - 45), vb3DHighlight
                    Call DrawMenuShadow(.hwnd, .hdc, .ScaleX(.ScaleWidth, vbHimetric, vbPixels), .ScaleY(.ScaleHeight, vbHimetric, vbPixels))
                End If
                If m_HiColourDisplay Then Call DrawButtonImage(m_ButtonPicture, (XPos + 1), (YPos + 1), True, True)
                Call DrawButtonImage(m_ButtonPicture, (XPos - 1), (YPos - 1), False, True)
             ElseIf ToState = Pressed Then 'NOT TOSTATE...
                If m_HiColourDisplay Then
                    MyColour = BlendColour(vbHighlight, m_BackColour, 70)
                Else
                    MyColour = vbButtonShadow
                End If
                lblCaption.ForeColor = vb3DHighlight
                If lblDD.Visible Then
                    lblDD.ForeColor = vb3DHighlight
                End If
                If m_OfficeXPStyle Then
                    If m_HasSeparator Then
                        Line (1, 1)-(.ScaleWidth, .ScaleHeight), vbButtonFace, BF
                        Line (40, 1)-(.ScaleWidth, .ScaleHeight), MyColour, BF
                        Line (40, 1)-(.ScaleWidth - 20, .ScaleHeight - 20), vb3DDKShadow, B
                    Else
                        Line (1, 1)-(.ScaleWidth, .ScaleHeight), MyColour, BF
                        Line (1, 1)-(.ScaleWidth - 20, .ScaleHeight - 20), vb3DDKShadow, B
                    End If
                 Else 'M_OFFICEXPSTYLE = FALSE/0
                    shBord.Shape = 4
                    shBord.BorderColor = BlendColour(vbButtonShadow, m_ForeColour, 120)
                    shBord.Move 4, 8, .ScaleWidth, .ScaleHeight - 24
                    shBord.Visible = True
                    Line (12, 24)-(.ScaleWidth - 24, .ScaleHeight - 48), MyColour, BF
                End If
                Call DrawButtonImage(m_ButtonPicture, (XPos), (YPos), False, True)
             ElseIf m_ButtonStyle = 1 And m_Checked Then 'NOT TOSTATE...
                If m_OfficeXPStyle Then
                    If m_HiColourDisplay Then
                        MyColour = BlendColour(vbHighlight, vbWindowBackground, 40)
                    Else
                        MyColour = vbHighlight
                    End If
                    Line (1, 1)-(.ScaleWidth, .ScaleHeight), MyColour, BF
                    If m_HiColourDisplay Then
                        MyColour = BlendColour(vbHighlight, vbWindowBackground, 100)
                    Else
                        MyColour = vbHighlightText
                    End If
                    Line (1, 1)-(.ScaleWidth - 20, .ScaleHeight - 20), MyColour, B
                 Else 'M_OFFICEXPSTYLE = FALSE/0
                    Draw .hwnd, .hdc
                    shBord.Shape = 0
                    If m_HiColourDisplay Then
                        shBord.BorderColor = BlendColour(vbButtonShadow, m_BackColour, 120)
                    Else
                        shBord.BorderColor = vbButtonShadow
                    End If
                    shBord.Move 0, 0, .ScaleWidth, .ScaleHeight
                    shBord.Visible = True
                End If
                Call DrawButtonImage(m_ButtonPicture, XPos, YPos, False, True)
             Else 'NOT M_BUTTONSTYLE...
                lblCaption.ForeColor = m_ForeColour
                If lblDD.Visible Then
                    lblDD.ForeColor = m_ForeColour
                End If
                Line (1, 1)-(.ScaleWidth, .ScaleHeight), .BackColor, BF
                Call DrawButtonImage(m_ButtonPicture, XPos, YPos, False, True)
            End If
        End If
        If m_HasSeparator And m_OfficeXPStyle Then
            Line (1, 40)-(1, .ScaleHeight - 40), vbButtonShadow, BF
        End If
        If Not Ambient.UserMode Then
            If m_HasSeparator Then
                Line (40, 1)-(.ScaleWidth - 20, .ScaleHeight - 20), vbBlack, B
            Else
                Line (1, 1)-(.ScaleWidth - 20, .ScaleHeight - 20), vbBlack, B
            End If
        End If
    End With 'USERCONTROL

End Sub

Private Sub DrawButtonImage(ByRef m_Picture As StdPicture, _
                            ByVal x As Long, _
                            ByVal y As Long, _
                            ByVal bShadow As Boolean, _
                            ByVal Enabled As Boolean)

  Dim lFlags As Long
  Dim hBrush As Long

    On Local Error Resume Next
    Select Case m_Picture.Type
     Case vbPicTypeBitmap
        lFlags = DST_BITMAP
     Case vbPicTypeIcon
        lFlags = DST_ICON
     Case Else
        lFlags = DST_COMPLEX
    End Select
    If bShadow Then
        If m_OfficeXPStyle Then
            hBrush = CreateSolidBrush(BlendColour(vbHighlight, vbButtonShadow, 10))
         Else 'M_OFFICEXPSTYLE = FALSE/0
            hBrush = CreateSolidBrush(BlendColour(vbButtonShadow, m_BackColour, 60))
        End If
    End If
    If Enabled Then
        DrawState UserControl.hdc, IIf(bShadow, hBrush, 0), 0, m_Picture.Handle, 0, x, y, UserControl.ScaleX(m_Picture.Width, vbHimetric, vbPixels), UserControl.ScaleY(m_Picture.Height, vbHimetric, vbPixels), lFlags Or IIf(bShadow, DSS_MONO, DSS_NORMAL)
     Else 'ENABLED = FALSE/0
        DrawState UserControl.hdc, IIf(bShadow, hBrush, 0), 0, m_Picture.Handle, 0, x, y, UserControl.ScaleX(m_Picture.Width, vbHimetric, vbPixels), UserControl.ScaleY(m_Picture.Height, vbHimetric, vbPixels), lFlags Or DSS_DISABLED
    End If
    If bShadow Then
        DeleteObject hBrush
    End If

End Sub

Private Function DrawGradient(ByVal hdc As Long, _
                              ByVal lWidth As Long, _
                              ByVal lHeight As Long) As Long

  Dim bDone       As Boolean
  Dim iIncX       As Integer
  Dim iIncY       As Integer
  Dim lIdx        As Long
  Dim lRet        As Long
  Dim hPen        As Long
  Dim hOldPen     As Long

  Dim laColors()  As Long
  Dim fMovX       As Double
  Dim fMovY       As Double
  Dim fDist       As Double
  Dim fAngle      As Double
  Dim fLongSide   As Double
  Dim uTmpPt      As POINTAPI
  Dim uaPts()     As POINTAPI
  Dim uaTmpPts()  As PointSng
    On Local Error Resume Next
    ReDim uaTmpPts(2)
    uaTmpPts(2).x = Int(lWidth / 2)
    uaTmpPts(2).y = Int(lHeight / 2)
    fLongSide = IIf(lWidth > lHeight, lWidth, lHeight)
    fDist = (Sqr((fLongSide ^ 2) + (fLongSide ^ 2)) + 2) / 2
    uaTmpPts(0).x = uaTmpPts(2).x - fDist
    uaTmpPts(0).y = uaTmpPts(2).y
    uaTmpPts(1).x = uaTmpPts(2).x + fDist
    uaTmpPts(1).y = uaTmpPts(2).y
    fAngle = 90 Mod 360
    Call RotatePoint(uaTmpPts(2), uaTmpPts(0), fAngle)
    Call RotatePoint(uaTmpPts(2), uaTmpPts(1), fAngle)
    If Abs(uaTmpPts(0).x - uaTmpPts(1).x) <= Abs(uaTmpPts(0).y - uaTmpPts(1).y) Then
        fMovX = IIf(uaTmpPts(0).x > uaTmpPts(1).x, -uaTmpPts(0).x, -uaTmpPts(1).x)
        fMovY = 0
        iIncX = 1
        iIncY = 0
     Else 'NOT ABS(UATMPPTS(0).X...
        fMovX = 0
        fMovY = IIf(uaTmpPts(0).y > uaTmpPts(1).y, lHeight - uaTmpPts(1).y, lHeight - uaTmpPts(0).y)
        iIncX = 0
        iIncY = -1
    End If
    ReDim uaPts(999)
    uaPts(0).x = uaTmpPts(0).x + fMovX
    uaPts(0).y = uaTmpPts(0).y + fMovY
    uaPts(1).x = uaTmpPts(1).x + fMovX
    uaPts(1).y = uaTmpPts(1).y + fMovY
    lIdx = 2
    Do While Not bDone
        uaPts(lIdx).x = uaPts(lIdx - 2).x + iIncX
        uaPts(lIdx).y = uaPts(lIdx - 2).y + iIncY
        lIdx = lIdx + 1
        Select Case True
         Case iIncX > 0  'Moving Left to Right
            bDone = uaPts(lIdx - 1).x > lWidth And uaPts(lIdx - 2).x > lWidth
         Case iIncX < 0  'Moving Right to Left
            bDone = uaPts(lIdx - 1).x < 0 And uaPts(lIdx - 2).x < 0
         Case iIncY > 0  'Moving Top to Bottom
            bDone = uaPts(lIdx - 1).y > lHeight And uaPts(lIdx - 2).y > lHeight
         Case iIncY < 0  'Moving Bottom to Top
            bDone = uaPts(lIdx - 1).y < 0 And uaPts(lIdx - 2).y < 0
        End Select
        If (lIdx Mod 1000) = 0 Then
            ReDim Preserve uaPts(UBound(uaPts) + 1000)
        End If
    Loop
    ReDim Preserve uaPts(lIdx - 1)
    lRet = BlendColors(mlColor1, mlColor2, lIdx / 2, 1, laColors)
    For lIdx = 0 To UBound(uaPts) - 1 Step 2
        'Move to next point
        lRet = MoveToEx(hdc, uaPts(lIdx).x, uaPts(lIdx).y, uTmpPt)
        'Create the colored pen and select it into the DC
        hPen = CreatePen(PS_SOLID, 1, laColors(Int(lIdx / 2)))
        hOldPen = SelectObject(hdc, hPen)
        'Draw the line
        lRet = LineTo(hdc, uaPts(lIdx + 1).x, uaPts(lIdx + 1).y)
        'Get the pen back out of the DC and destroy it
        lRet = SelectObject(hdc, hOldPen)
        lRet = DeleteObject(hPen)
    Next lIdx
    DrawGradient = lIdx

End Function

Private Sub DrawMenuShadow(ByVal hwnd As Long, _
                           ByVal hdc As Long, _
                           ByVal xOrg As Long, _
                           ByVal yOrg As Long)

  
  Dim Rec  As RECT
  Dim winW As Long
  Dim winH As Long

  Dim x    As Long
  Dim y    As Long
  Dim c    As Long
    
    GetWindowRect hwnd, Rec
    winW = (Rec.Right - Rec.Left)
    winH = (Rec.Bottom - Rec.Top)
    c = TranslateColour(UserControl.BackColor)
    For x = 1 To 4
        For y = 0 To 3
            SetPixel hdc, winW - x, y, c
        Next y
        For y = 4 To 7
            SetPixel hdc, winW - x, y, pMask(3 * x * (y - 3), c)
        Next y
        For y = 8 To winH - 5
            SetPixel hdc, winW - x, y, pMask(15 * x, c)
        Next y
        For y = winH - 4 To winH - 1
            SetPixel hdc, winW - x, y, pMask(3 * x * -(y - winH), c)
        Next y
    Next x
    For y = 1 To 4
        For x = 0 To 3
            SetPixel hdc, x, winH - y, c
        Next x
        For x = 4 To 7
            SetPixel hdc, x, winH - y, pMask(3 * (x - 3) * y, c)
        Next x
        For x = 8 To winW - 5
            SetPixel hdc, x, winH - y, pMask(15 * y, c)
        Next x
    Next y

End Sub

Public Property Get Enabled() As Boolean

    Enabled = m_Enabled
    'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
    'MemberInfo=0,0,0,True

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    If m_Enabled Then
        Call DrawButton(Normal)
     Else 'M_ENABLED = FALSE/0
        Call DrawButton(Disabled)
    End If

End Property

Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512

    Set Font = m_Font
    'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
    'MemberInfo=6,0,0,0

End Property

Public Property Set Font(ByVal New_Font As Font)

    Set m_Font = New_Font
    PropertyChanged "Font"
    Set lblCaption.Font = m_Font
    Call DrawButton(5)

End Property

Public Property Get ForeColour() As OLE_COLOR

    ForeColour = m_ForeColour
    'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
    'MemberInfo=10,0,0,0

End Property

Public Property Let ForeColour(ByVal New_ForeColour As OLE_COLOR)

    m_ForeColour = New_ForeColour
    PropertyChanged "ForeColour"
    Call DrawButton(5)

End Property

Public Property Get HilightColour() As OLE_COLOR

    HilightColour = m_HilightColour
    'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
    'MemberInfo=10,0,0,0

End Property

Public Property Let HilightColour(ByVal New_HilightColour As OLE_COLOR)

    m_HilightColour = New_HilightColour
    PropertyChanged "HilightColour"

End Property

Public Property Get ImageHeight() As Integer

    ImageHeight = m_ImageHeight
    'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
    'MemberInfo=7,0,0,16

End Property

Public Property Let ImageHeight(ByVal New_ImageHeight As Integer)

    m_ImageHeight = New_ImageHeight
    PropertyChanged "ImageHeight"
    Call DrawButton(5)

End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse

    If CurrentMessage = WM_MOUSEMOVE Or CurrentMessage = WM_MOUSELEAVE Or CurrentMessage = WM_LBUTTONUP Or CurrentMessage = WM_LBUTTONDOWN Then
        ISubclass_MsgResponse = emrPreprocess
    End If

End Property

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)

  '

End Property

Private Sub lblCaption_Click()

    If Not m_Enabled Then
        Exit Sub
    End If
    RaiseEvent Click
    Call DrawButton(Normal)
    bTracking = False

End Sub

Private Sub lblDD_Click()

    If Not m_Enabled Then
        Exit Sub
    End If
    RaiseEvent Click
    Call DrawButton(Normal)
    bTracking = False

End Sub

Public Property Get OfficeXPStyle() As Boolean

    OfficeXPStyle = m_OfficeXPStyle
    'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
    'MemberInfo=0,0,0,False

End Property

Public Property Let OfficeXPStyle(ByVal New_OfficeXPStyle As Boolean)

    m_OfficeXPStyle = New_OfficeXPStyle
    PropertyChanged "OfficeXPStyle"
    Call DrawButton(5)

End Property

Private Function pMask(ByVal lScale As Long, _
                       ByVal lColor As Long) As Long

  Dim R        As Long
  Dim g        As Long
  Dim b        As Long
  Dim MyColour As Long

    MyColour = TranslateColour(lColor)
    R = MyColour And &HFF
    g = (MyColour And &HFF00&) \ &H100&
    b = (MyColour And &HFF0000) \ &H10000
    R = pTransform(lScale, R)
    g = pTransform(lScale, g)
    b = pTransform(lScale, b)
    pMask = RGB(R, g, b)

End Function

Private Function pTransform(ByVal lScale As Long, _
                            ByVal lColor As Long) As Long

    pTransform = lColor - Int(lColor * lScale / 255)
    ' - Function pTransform converts
    ' a RGB subcolor using a scale
    ' where 0 = 0 and 255 = lScale

End Function

Private Sub RotatePoint(uAxisPt As PointSng, _
                        uRotatePt As PointSng, _
                        ByVal fDegrees As Double)

  Dim fDX         As Double
  Dim fDY         As Double
  Dim fRadians    As Double

    fRadians = fDegrees * RADS
    With uRotatePt
        fDX = .x - uAxisPt.x
        fDY = .y - uAxisPt.y
        .x = uAxisPt.x + ((fDX * Cos(fRadians)) + (fDY * Sin(fRadians)))
        .y = uAxisPt.y + -((fDX * Sin(fRadians)) - (fDY * Cos(fRadians)))
    End With 'uRotatePt

End Sub

Private Sub StartMouseTracking()

  Dim tET As tagTRACKMOUSEEVENT
  Dim lR  As Long

    On Error Resume Next
    If Not bTracking Then
        With tET
            .cbSize = Len(tET)
            .dwFlags = TME_HOVER Or TME_LEAVE
            .dwHoverTime = HOVER_DEFAULT
            .hwndTrack = m_Hwnd
        End With 'tET
        lR = TrackMouseEvent(tET)
        bTracking = True
    End If
    On Error GoTo 0

End Sub

Private Function TranslateColour(ByVal oClr As OLE_COLOR, _
                                 Optional hPal As Long = 0) As Long

    If OleTranslateColor(oClr, hPal, TranslateColour) Then
        TranslateColour = CLR_INVALID
    End If

End Function

Public Property Get UseParentColour() As Boolean

    UseParentColour = m_UseParentColour
    'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
    'MemberInfo=0,0,0,False

End Property

Public Property Let UseParentColour(ByVal New_UseParentColour As Boolean)

    m_UseParentColour = New_UseParentColour
    PropertyChanged "UseParentColour"
    Call DrawButton(5)

End Property

Private Sub UserControl_Click()

    If Not m_Enabled Then
        Exit Sub
    End If
    RaiseEvent Click
    Call DrawButton(Normal)
    bTracking = False

End Sub

Private Sub UserControl_Initialize()

    bTracking = False
    bClosing = False
    Call Attach

End Sub

Private Sub UserControl_InitProperties()

    On Error Resume Next
    m_Caption = m_def_Caption
    m_CaptionPosition = m_def_CaptionPosition
    Set m_Font = Ambient.Font
    Set m_ButtonPicture = Nothing
    m_Enabled = m_def_Enabled
    m_ImageHeight = m_def_ImageHeight
    m_ButtonStyle = m_def_ButtonStyle
    m_Checked = m_def_Checked
    m_BackColour = m_def_BackColour
    m_ForeColour = m_def_ForeColour
    m_UseParentColour = m_def_UseParentColour
    m_HilightColour = m_def_HilightColour
    m_OfficeXPStyle = m_def_OfficeXPStyle
    m_HasSeparator = m_def_HasSeparator
    m_HiColourDisplay = m_def_HiColourDisplay
    On Error GoTo 0

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    On Error Resume Next
    With PropBag
        m_Caption = .ReadProperty("Caption", m_def_Caption)
        m_CaptionPosition = .ReadProperty("CaptionPosition", m_def_CaptionPosition)
        Set m_Font = .ReadProperty("Font", Ambient.Font)
        Set m_ButtonPicture = .ReadProperty("ButtonPicture", Nothing)
        m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
        m_ImageHeight = .ReadProperty("ImageHeight", m_def_ImageHeight)
        m_ButtonStyle = .ReadProperty("ButtonStyle", m_def_ButtonStyle)
        m_Checked = .ReadProperty("Checked", m_def_Checked)
        m_BackColour = .ReadProperty("BackColour", m_def_BackColour)
        m_ForeColour = .ReadProperty("ForeColour", m_def_ForeColour)
        m_UseParentColour = .ReadProperty("UseParentColour", m_def_UseParentColour)
        m_HilightColour = .ReadProperty("HilightColour", m_def_HilightColour)
        m_OfficeXPStyle = .ReadProperty("OfficeXPStyle", m_def_OfficeXPStyle)
        m_HasSeparator = .ReadProperty("HasSeparator", m_def_HasSeparator)
        m_HiColourDisplay = .ReadProperty("HiColourDisplay", m_def_HiColourDisplay)
    End With 'PropBag
    Set lblCaption.Font = m_Font
    Call DrawButton(5)
    On Error GoTo 0

End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
    If Not Ambient.UserMode Then
        If Not bClosing Then
            Call DrawButton(5)
        End If
    End If
    On Error GoTo 0

End Sub

Private Sub UserControl_Terminate()

    bClosing = True
    Call Detach

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    On Error Resume Next
    With PropBag
        Call .WriteProperty("Caption", m_Caption, m_def_Caption)
        Call .WriteProperty("CaptionPosition", m_CaptionPosition, m_def_CaptionPosition)
        Call .WriteProperty("Font", m_Font, Ambient.Font)
        Call .WriteProperty("ButtonPicture", m_ButtonPicture, Nothing)
        Call .WriteProperty("Enabled", m_Enabled, m_def_Enabled)
        Call .WriteProperty("ImageHeight", m_ImageHeight, m_def_ImageHeight)
        Call .WriteProperty("ButtonStyle", m_ButtonStyle, m_def_ButtonStyle)
        Call .WriteProperty("Checked", m_Checked, m_def_Checked)
        Call .WriteProperty("BackColour", m_BackColour, m_def_BackColour)
        Call .WriteProperty("ForeColour", m_ForeColour, m_def_ForeColour)
        Call .WriteProperty("UseParentColour", m_UseParentColour, m_def_UseParentColour)
        Call .WriteProperty("HilightColour", m_HilightColour, m_def_HilightColour)
        Call .WriteProperty("OfficeXPStyle", m_OfficeXPStyle, m_def_OfficeXPStyle)
        Call .WriteProperty("HasSeparator", m_HasSeparator, m_def_HasSeparator)
        Call .WriteProperty("HiColourDisplay", m_HiColourDisplay, m_def_HiColourDisplay)
    End With 'PropBag
    On Error GoTo 0

End Sub

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Not m_Enabled Then
        Exit Function
    End If
    If hwnd <> m_Hwnd Then
        Exit Function
    End If
    Select Case iMsg
        Case WM_MOUSEMOVE
            If MyState <> Pressed Then
                Call DrawButton(Over)
                Call StartMouseTracking
            End If
        Case WM_MOUSELEAVE
            If MyState <> Normal Then
                Call DrawButton(Normal)
            End If
            bTracking = False
        Case WM_LBUTTONDOWN
            If MyState <> Pressed Then
                Call DrawButton(Pressed)
            End If
        Case WM_LBUTTONUP
            If bTracking Then
                If MyState <> Over Then
                    Call DrawButton(Over)
                End If
            Else
                If MyState <> Normal Then
                    Call DrawButton(Normal)
                End If
            End If
    End Select
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get HasSeparator() As Boolean
    HasSeparator = m_HasSeparator
End Property

Public Property Let HasSeparator(ByVal New_HasSeparator As Boolean)
    m_HasSeparator = New_HasSeparator
    PropertyChanged "HasSeparator"
    Call DrawButton(5)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColour() As OLE_COLOR
    BackColour = m_BackColour
End Property

Public Property Let BackColour(ByVal New_BackColour As OLE_COLOR)
    m_BackColour = New_BackColour
    PropertyChanged "BackColour"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get HiColourDisplay() As Boolean
    HiColourDisplay = m_HiColourDisplay
End Property

Public Property Let HiColourDisplay(ByVal New_HiColourDisplay As Boolean)
    m_HiColourDisplay = New_HiColourDisplay
    PropertyChanged "HiColourDisplay"
End Property

