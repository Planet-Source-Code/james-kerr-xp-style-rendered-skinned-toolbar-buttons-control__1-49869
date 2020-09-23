VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4125
   ClientLeft      =   1530
   ClientTop       =   1935
   ClientWidth     =   5190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   810
      Top             =   4230
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stye Options ..."
      Height          =   3480
      Left            =   45
      TabIndex        =   1
      Top             =   585
      Width           =   5100
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3165
         Left            =   90
         ScaleHeight     =   3165
         ScaleWidth      =   4920
         TabIndex        =   2
         Top             =   225
         Width           =   4920
         Begin VB.CheckBox Check2 
            Alignment       =   1  'Right Justify
            Caption         =   "Skin Form"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3645
            TabIndex        =   15
            Top             =   2880
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   330
            Index           =   2
            Left            =   2925
            TabIndex        =   11
            Top             =   2385
            Width           =   420
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   330
            Index           =   1
            Left            =   2925
            TabIndex        =   10
            Top             =   1980
            Width           =   420
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   330
            Index           =   0
            Left            =   2925
            TabIndex        =   9
            Top             =   1575
            Width           =   420
         End
         Begin VB.CheckBox Check1 
            Caption         =   "High Colour Display"
            Height          =   285
            Left            =   540
            TabIndex        =   5
            Top             =   630
            Value           =   1  'Checked
            Width           =   3840
         End
         Begin VB.OptionButton optXP 
            Caption         =   "Use Rendered Buttons"
            Height          =   285
            Index           =   1
            Left            =   225
            TabIndex        =   4
            Top             =   1215
            Width           =   4470
         End
         Begin VB.OptionButton optXP 
            Caption         =   "Use Office XP Style Buttons"
            Height          =   285
            Index           =   0
            Left            =   225
            TabIndex        =   3
            Top             =   225
            Value           =   -1  'True
            Width           =   4470
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C00000&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Index           =   2
            Left            =   3465
            TabIndex        =   14
            Top             =   2385
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label Label2 
            BackColor       =   &H00AD735A&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Index           =   1
            Left            =   3465
            TabIndex        =   13
            Top             =   1980
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Index           =   0
            Left            =   3465
            TabIndex        =   12
            Top             =   1575
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Select Highlight Colour:"
            Enabled         =   0   'False
            Height          =   330
            Index           =   2
            Left            =   540
            TabIndex        =   8
            Top             =   2430
            Width           =   2310
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Select Back Colour:"
            Enabled         =   0   'False
            Height          =   330
            Index           =   1
            Left            =   540
            TabIndex        =   7
            Top             =   2025
            Width           =   2310
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Select Fore Colour:"
            Enabled         =   0   'False
            Height          =   330
            Index           =   0
            Left            =   540
            TabIndex        =   6
            Top             =   1620
            Width           =   2310
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   135
      Top             =   4140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":015A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":02B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":040E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0568
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":06C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":081C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   5190
      TabIndex        =   0
      Top             =   0
      Width           =   5190
      Begin prjTestUcButton.ucButton ucButton1 
         Height          =   375
         Index           =   0
         Left            =   45
         ToolTipText     =   "Button Tooltip ...."
         Top             =   15
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Caption         =   "New"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OfficeXPStyle   =   -1  'True
      End
      Begin prjTestUcButton.ucButton ucButton1 
         Height          =   375
         Index           =   1
         Left            =   855
         ToolTipText     =   "Button Tooltip ...."
         Top             =   20
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "Open"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OfficeXPStyle   =   -1  'True
         HasSeparator    =   -1  'True
      End
      Begin prjTestUcButton.ucButton ucButton1 
         Height          =   375
         Index           =   2
         Left            =   1755
         ToolTipText     =   "Button Tooltip ...."
         Top             =   20
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   1
         Checked         =   -1  'True
         OfficeXPStyle   =   -1  'True
         HasSeparator    =   -1  'True
      End
      Begin prjTestUcButton.ucButton ucButton1 
         Height          =   375
         Index           =   3
         Left            =   2205
         ToolTipText     =   "Button Tooltip ...."
         Top             =   20
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   1
         OfficeXPStyle   =   -1  'True
      End
      Begin prjTestUcButton.ucButton ucButton1 
         Height          =   375
         Index           =   4
         Left            =   2610
         ToolTipText     =   "Button Tooltip ...."
         Top             =   20
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   1
         OfficeXPStyle   =   -1  'True
      End
      Begin prjTestUcButton.ucButton ucButton1 
         Height          =   375
         Index           =   5
         Left            =   3060
         ToolTipText     =   "Button Tooltip ...."
         Top             =   20
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   661
         Caption         =   "Select"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   2
         OfficeXPStyle   =   -1  'True
         HasSeparator    =   -1  'True
      End
      Begin prjTestUcButton.ucButton ucButton1 
         Height          =   375
         Index           =   6
         Left            =   4185
         ToolTipText     =   "Close the Form!"
         Top             =   20
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "Close"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OfficeXPStyle   =   -1  'True
         HasSeparator    =   -1  'True
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "none"
      Visible         =   0   'False
      Begin VB.Menu mnuPopSelect1 
         Caption         =   "Drop Down Selection One"
      End
      Begin VB.Menu mnuPopSelect2 
         Caption         =   "Drop Down Selection Two"
      End
      Begin VB.Menu mnuPopSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopSelect3 
         Caption         =   "Drop Down Selection Three"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'***********************************
'Not the prettiest code - just run up quickly
'to give you an idea how to use the control ...
'
'Form needs a grab handle or coding to move it, Error handlers, etc ...
'
'Note- Subclassing - Don't hit the stop button on the VB toolbar!
'***********************************

Private Sub SetButtonStyle(ByVal OfficeXP As Boolean, Optional ByVal HiColour As Boolean, Optional ByVal ForeColour As Long, Optional ByVal BackColour As Long, Optional ByVal HighLight As Long, Optional ByVal SkinIt As Boolean)
    
    Dim lVar As Long
    Dim lButtonCount As Long
    Dim MyControl As Control
    
    'stop any redraw flicker
    Call LockWindowUpdate(Me.hwnd)
    
    lButtonCount = ucButton1.UBound
    
    For lVar = 0 To lButtonCount
        
        'Set the button Style
        '    OfficeXPStyle TRUE (self explanatory - modified by the HiColourDisplay setting)
        '                  Ignores Forecolour/BackColour/HighlighColour
        '    OfficeXPStyle FALSE (use rendering)
        '                  Requires Forecolour/BackColour/HighlighColour
        
        ucButton1(lVar).OfficeXPStyle = OfficeXP
        
        'If HiColour is set to FALSE, the control will provide XP Style buttons for 16/256 colour displays
        
        ucButton1(lVar).HiColourDisplay = HiColour
        
        If Not OfficeXP Then
            
            'set the height and background of the Toolbar container ...
            'Note that rendered buttons are a bit bigger than the XP ones
            
            picToolbar.BackColor = BackColour
            picToolbar.Height = 512
            
            'Rendered buttons look better at the top of the container
            ucButton1(lVar).Top = 0
            
            'we need to set up colours ...
            ucButton1(lVar).BackColour = BackColour
            ucButton1(lVar).ForeColour = ForeColour
            ucButton1(lVar).HilightColour = HighLight
        
        Else
        
            'XP Style buttons need a bit of space from the top
            ucButton1(lVar).Top = 20
            
            'Set up standard colours for XP Style ...
            ucButton1(lVar).BackColour = vbButtonFace
            ucButton1(lVar).ForeColour = vbButtonText
            ucButton1(lVar).HilightColour = vbButtonText
            
            picToolbar.BackColor = vbButtonFace
            picToolbar.Height = 415
            
        End If
        
        'assign the bitmaps to the toolbar ...
        'Note - this version of the ucButton control doesn't really support 'Caption Only' buttons
        '16x16, 24x24 and 32x32,etc. icons (all colour depths) are supported - set the ImageHeight Property
        'on the control accordingly.
        
        Set ucButton1(lVar).ButtonPicture = ImageList1.ListImages(lVar + 1).Picture
        
        'Because Button Width can change ... move them together
        
        If lVar > 0 Then
            ucButton1(lVar).Left = (ucButton1(lVar - 1).Left + ucButton1(lVar - 1).Width + 45)
        End If
    
        '... and optionally skin the form (Messy and quick way!)
        
        If SkinIt Then
            For Each MyControl In Me.Controls
                Select Case TypeName(MyControl)
                    Case "CheckBox", "OptionButton", "Frame"
                        MyControl.BackColor = BackColour
                        MyControl.ForeColor = ForeColour
                End Select
            Next
            Picture1.BackColor = BackColour
            Me.BackColor = BackColour
            Label1(0).ForeColor = ForeColour
            Label1(1).ForeColor = ForeColour
            Label1(2).ForeColor = ForeColour
        Else
            For Each MyControl In Me.Controls
                Select Case TypeName(MyControl)
                    Case "CheckBox", "OptionButton", "Frame"
                        MyControl.BackColor = vbButtonFace
                        MyControl.ForeColor = vbButtonText
                End Select
            Next
            Picture1.BackColor = vbButtonFace
            Me.BackColor = vbButtonFace
            Label1(0).ForeColor = vbButtonText
            Label1(1).ForeColor = vbButtonText
            Label1(2).ForeColor = vbButtonText
        End If
    Next

    'unlock window updates
    Call LockWindowUpdate(0&)

End Sub

Private Sub Check1_Click()
    
    If Check1.Value = 1 Then
        Call SetButtonStyle(True, True)
    Else
        Call SetButtonStyle(True, False)
    End If

End Sub

Private Sub Check2_Click()
    
    If Check2.Value = 1 Then
    
        Call SetButtonStyle(False, False, Label2(0).BackColor, Label2(1).BackColor, Label2(2).BackColor, True)
    Else
        Call SetButtonStyle(False, False, Label2(0).BackColor, Label2(1).BackColor, Label2(2).BackColor, False)
        
    End If

End Sub

Private Sub Command1_Click(Index As Integer)
    
    With CommonDialog1
        .Color = Label2(Index).BackColor
        .ShowColor
        Label2(Index).BackColor = .Color
        
        If Check2.Value = 1 Then
        
            Call SetButtonStyle(False, False, Label2(0).BackColor, Label2(1).BackColor, Label2(2).BackColor, True)
        Else
            Call SetButtonStyle(False, False, Label2(0).BackColor, Label2(1).BackColor, Label2(2).BackColor, False)
            
        End If
        
    End With
    
End Sub

Private Sub Form_Load()

    Call SetButtonStyle(True, True)
    
End Sub

Private Sub optXP_Click(Index As Integer)
    
    Dim lVar As Long
    
    If optXP(0).Value Then
        Check1.Enabled = True
        
        Check2.Enabled = False
        For lVar = 0 To 2
            Label1(lVar).Enabled = False
            Label2(lVar).Visible = False
            Command1(lVar).Enabled = False
        Next
        
        If Check1.Value = 1 Then
            Call SetButtonStyle(True, True)
        Else
            Call SetButtonStyle(True, False)
        End If
    Else
        Check1.Enabled = False
        
        Check2.Enabled = True
        For lVar = 0 To 2
            Label1(lVar).Enabled = True
            Label2(lVar).Visible = True
            Command1(lVar).Enabled = True
        Next
        
        If Check2.Value = 1 Then
        
            Call SetButtonStyle(False, False, Label2(0).BackColor, Label2(1).BackColor, Label2(2).BackColor, True)
        Else
            Call SetButtonStyle(False, False, Label2(0).BackColor, Label2(1).BackColor, Label2(2).BackColor, False)
            
        End If
        
    End If
    
End Sub

Private Sub ucButton1_Click(Index As Integer)
    Select Case Index
        Case 2
            ucButton1(2).Checked = Not ucButton1(2).Checked
        Case 5
            Me.PopupMenu mnuPopUp, 0, ucButton1(5).Left, (ucButton1(5).Top + ucButton1(5).Height)
        Case 6
            Unload Me
    End Select
End Sub
