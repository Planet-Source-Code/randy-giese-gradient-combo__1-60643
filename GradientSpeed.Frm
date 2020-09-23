VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form GradientSpeed 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GradientSpeed"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   89
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   681
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSpeed 
      BorderStyle     =   0  'None
      Height          =   1035
      Index           =   0
      Left            =   150
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   660
      TabIndex        =   0
      Top             =   150
      Width           =   9900
      Begin MSComctlLib.Slider sldSpeed 
         Height          =   495
         Left            =   0
         TabIndex        =   1
         Top             =   495
         Width           =   9900
         _ExtentX        =   17463
         _ExtentY        =   873
         _Version        =   393216
         Min             =   -50
         Max             =   50
         TickFrequency   =   5
      End
      Begin VB.Label lblSpeed 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Display Speed!"
         Height          =   360
         Index           =   0
         Left            =   3300
         TabIndex        =   4
         Top             =   75
         Width           =   3300
      End
      Begin VB.Label lblSpeed 
         Alignment       =   2  'Center
         Caption         =   "Slower"
         Height          =   300
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Top             =   75
         Width           =   3300
      End
      Begin VB.Label lblSpeed 
         Alignment       =   2  'Center
         Caption         =   "Faster"
         Height          =   300
         Index           =   2
         Left            =   6600
         TabIndex        =   2
         Top             =   75
         Width           =   3300
      End
   End
End
Attribute VB_Name = "GradientSpeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()

'
'   *************************************************************************
'   ****                                                                 ****
'   ****    The variable (blnSpeed) indicates that the Speed control     ****
'   ****    is visible and halts the displays countdown timer.  The      ****
'   ****    display is allowed to continue on (without the timer) so     ****
'   ****    that the user can see the effect of any Speed Control        ****
'   ****    changes.  The timer is also halted to prevent a different    ****
'   ****    display from starting.  This way the user knows exactly      ****
'   ****    which display is being changed and how much.                 ****
'   ****                                                                 ****
'   ****    I save the difference (lngSpeedDifference) between the       ****
'   ****    current displays start time (lngGradientStart) and the       ****
'   ****    current time, which is held in the (lngGradientStop)         ****
'   ****    variable.  This allows me to correctly reset the             ****
'   ****    countdown timer when the Speed control is closed.            ****
'   ****                                                                 ****
'   *************************************************************************
'

    lngSpeedDifference = lngGradientStop - lngGradientStart
    blnSpeed = True
    sldSpeed.Value = lngSpeed(lngSubNum)

    With Me
        .Top = 300
        .Left = (Screen.Width - .Width) \ 2
    End With    '   Me

    SetCursorPos Screen.Width \ Screen.TwipsPerPixelX \ 2, (Me.Top \ Screen.TwipsPerPixelY) + (Me.Height \ Screen.TwipsPerPixelY \ 2) + 21

End Sub

Private Sub Form_Unload(Cancel As Integer)

    blnSpeed = False
    Unload Me

End Sub

Private Sub sldSpeed_Change()

'
'   *************************************************************************
'   ****                                                                 ****
'   ****    The lngSpeed variable is used to maintain the position of    ****
'   ****    the pointer even though the form is unloaded.                ****
'   ****                                                                 ****
'   *************************************************************************
'

    lngSpeed(lngSubNum) = sldSpeed.Value
    subSpeed

End Sub
':)Code Fixer V3.0.9 (5/22/2005 8:19:58 PM) 2 + 72 = 74 Lines Thanks Ulli for inspiration and lots of code.
