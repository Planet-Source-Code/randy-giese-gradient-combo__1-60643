VERSION 5.00
Begin VB.Form GradientCombo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1980
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   10185
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawWidth       =   2
   Icon            =   "GradientCombo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   132
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   679
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTime 
      Height          =   1575
      Left            =   3000
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   396
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   6000
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   840
         TabIndex        =   3
         Top             =   900
         Width           =   1320
      End
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close!"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3840
         TabIndex        =   4
         Top             =   967
         Width           =   1320
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Caption         =   "How long do you want this display to be active?  Enter the number of seconds."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   150
         TabIndex        =   2
         Top             =   150
         Width           =   5700
      End
   End
   Begin VB.PictureBox picSpeed 
      AutoRedraw      =   -1  'True
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1440
      ScaleHeight     =   330
      ScaleWidth      =   465
      TabIndex        =   0
      Top             =   1440
      Width           =   525
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "&Setup"
      Visible         =   0   'False
      Begin VB.Menu mnuTime 
         Caption         =   "&Run Time"
      End
      Begin VB.Menu mnuSpeed 
         Caption         =   "Speed"
      End
   End
End
Attribute VB_Name = "GradientCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Gradient_API uses the 'GradientFillTriangle' API to          ****
'   ****    create two opposed Gradient triangles.  The effect is        ****
'   ****    that you get a screen sized display that is a Gradient in    ****
'   ****    all directions.  By using the API, I can achieve very        ****
'   ****    fast graphics.                                               ****
'   ****                                                                 ****
'   ****    I set it up to use a Timer (tmrGrad_API) to continually      ****
'   ****    refresh the display every second.  I included a popup        ****
'   ****    Speed control, accessible by right-clicking the form, so     ****
'   ****    that the user may control the display speed.  I set the      ****
'   ****    minimum time to 10 milliseconds, but I find that to be       ****
'   ****    too fast.  It causes an annoying flashing.  On the other     ****
'   ****    hand, the maximum of 10 seconds is very slow (too slow       ****
'   ****    for me).                                                     ****
'   ****                                                                 ****
'   ****    The colors of the 4 corners, and the direction that the      ****
'   ****    triangles face, are chosen at random, so the display will    ****
'   ****    be constantly changing.                                      ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'

'** Variables **
Private blnPause                         As Boolean

Private curEnd                           As Currency
Private curQPerfFreq                     As Currency
Private curStart                         As Currency

Private fsoIniFile                       As New Scripting.FileSystemObject
Private GradTriangle                     As GRADIENT_TRIANGLE

Private lng4Corner(0 To 6)               As Long
Private lngClockPrev                     As Long
Private lngCntr                          As Long
Private lngColorNum(0 To 3)              As Long
Private lngCorner                        As Long
Private lngCornerNum(0 To 3)             As Long
Private lngDiagonal                      As Long
Private lngHowLong(1 To 4)               As Long
Private lngLoop1                         As Long
Private lngLoop2                         As Long
Private lngNum                           As Long
Private lngPorM(0 To 3)                  As Long
Private lngRGB(0 To 3)                   As Long
Private lngUorD                          As Long

Private strTemp                          As String
Private strWinDir                        As String

Private TriVert(0 To 3)                  As TRIVERTEX

'
'   *************************************************************************
'   ****                                                                 ****
'   ****    In this mode, an array of TRIVERTEX structures is passed     ****
'   ****    to GDI along with a list of array indexes that describe      ****
'   ****    separate triangles.  GDI performs linear interpolation       ****
'   ****    between triangle vertices and fills the interior.            ****
'   ****    Drawing is done directly in 24- and 32-bpp modes.            ****
'   ****    Dithering is performed in 16-, 8-, 4-, and 1-bpp mode.       ****
'   ****                                                                 ****
'   *************************************************************************
'

Private Const GRADIENT_FILL_TRIANGLE     As Long = &H2

Private Type TRIVERTEX
    X                                        As Long
    Y                                        As Long
    Red                                      As Integer '  Ushort Red value
    Green                                    As Integer '  Ushort Green value
    Blue                                     As Integer '  Ushort Blue value
    Alpha                                    As Integer '  Ushort Alpha value
End Type

Private Type GRADIENT_TRIANGLE
    Vertex1                                  As Long    '  In reality this is an UNSIGNED Long
    Vertex2                                  As Long    '  In reality this is an UNSIGNED Long
    Vertex3                                  As Long    '  In reality this is an UNSIGNED Long
End Type

'** API Declarations **
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetInputState Lib "user32" () As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_TRIANGLE, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Sub Form_DblClick()

    subQuit

End Sub

Private Sub Form_Load()

Dim strSpeed As String
Dim strTime  As String

    Randomize Timer
    Me.Show

    QueryPerformanceFrequency curQPerfFreq  '  Get the timer frequency
    curQPerfFreq = curQPerfFreq / 1000      '  Break it down to a millisecond

    curFrequency(1) = curQPerfFreq * 350    '  Base Frequency for this routine
    curFrequency(2) = curQPerfFreq * 100    '  Base Frequency for this routine
    curFrequency(3) = curQPerfFreq / 2      '  Base Frequency for this routine
    curFrequency(4) = curQPerfFreq * 100    '  Base Frequency for this routine

'   Create a buffer string
    strTemp = String$(200, vbNullChar)
'   Get the windows directory
    strWinDir = Left$(strTemp, GetWindowsDirectory(strTemp, Len(strTemp))) & "\Grad_Comb.Ini"

    If fsoIniFile.FileExists(strWinDir) Then
'       Read setup values if they exist.
        Open strWinDir For Input As #1

        For lngSubNum = 1 To 4
'           Get the current Speed values
            Line Input #1, strSpeed
            lngSpeed(lngSubNum) = Val(strSpeed)
'           Then get the default Time values
            Line Input #1, strTime
            lngHowLong(lngSubNum) = Val(strTime)

            subSpeed    '   Set the Speed for this display
        Next lngSubNum
    Else      '  fsoIniFile.FileExists(strWinDir) = FALSE/0
'       If there is no Config file, create one and fill all values.
        Open strWinDir For Output As #1

'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    This controls the speed of the display.  Each display has    ****
'   ****    it's own speed rating.                                       ****
'   ****                                                                 ****
'   ****    curFrequency() is the base Frequency for the routine.        ****
'   ****    The speed control can be adjusted faster or slower based     ****
'   ****    on this figure.  That results in the (curFreq())             ****
'   ****    variable, which is then used in (subSlowDown) to actually    ****
'   ****    control the display speed.                                   ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'

        For lngLoop1 = 1 To 4
'           Save the current Speed values
            curFreq(lngLoop1) = curFrequency(lngLoop1)
            Print #1, lngSpeed(lngLoop1)
'           Then save the default Time values
            lngHowLong(lngLoop1) = 6 '   The number of seconds to run this routine
            Print #1, lngHowLong(lngLoop1)
        Next lngLoop1
    End If

    Close #1
    blnPause = False
    lngNum = 1
    lngUorD = -1

    With picSpeed
        .Top = (Me.Height \ Screen.TwipsPerPixelY) - .Height - 5
        .Left = (Me.Width \ Screen.TwipsPerPixelX) - .Width - 5
    End With    '   picSpeed

    With picTime
        .Top = 20
        .Left = ((Me.Width \ Screen.TwipsPerPixelX) - .Width) \ 2
    End With    '   picTime

    subDoubleTriFullScr

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Click the Right Mouse button to Pop-up the Menu.             ****
'   ****                                                                 ****
'   ****    Click the Left Mouse button to Freeze or Thaw the            ****
'   ****    display, depending on the current screen status.             ****
'   ****                                                                 ****
'   *************************************************************************
'

    If Button = vbRightButton Then
        Me.PopupMenu mnuSetup
    Else      '  NOT Button....

'       Click to Pause the display.
        blnPause = Not blnPause
        subFreeze
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    subQuit

End Sub

Private Sub lblClose_Click()

Dim lngTime As Long

    picTime.Visible = False
    lngTime = Val(txtTime.Text)

    If lngTime < 100000 Then
        lngHowLong(lngSubNum) = lngTime
        blnPause = False
    End If

End Sub

Private Function LongToUShort(Unsigned As Long) As Long

    LongToUShort = CInt(Unsigned - &H10000)

End Function

Private Sub mnuSpeed_Click()

    BringWindowToTop GradientSpeed.hwnd
    GradientSpeed.Show

End Sub

Private Sub mnuTime_Click()

    txtTime.Text = lngHowLong(lngSubNum)
    picTime.Visible = True

    With picTime
        .Top = 20
        .Left = ((Me.Width \ Screen.TwipsPerPixelX) - .Width) \ 2
    End With    '   picTime

    SetCursorPos Screen.Width \ Screen.TwipsPerPixelX \ 2, picTime.Top + (picTime.Height \ 2) + 21

    blnPause = True
    subFreeze

End Sub

Private Sub subChangeColors()

'   This updates the Gradient's colors
    lngRGB(1) = lngRGB(1) + lngPorM(1)      '  Change first color value

'   Check all three color groups to make sure they are within legal bounds.
    For lngCntr = 1 To 3
        If lngRGB(lngCntr) < 1 Or lngRGB(lngCntr) > 254 Then   '  If color reaches min or max,
            subIncrementColors
        End If
    Next lngCntr

End Sub

Private Sub subColor()

'
'   *************************************************************************
'   ****                                                                 ****
'   ****    VB uses the color groups: Red, Green & Blue.                 ****
'   ****    The value for each color group is chosen at random and       ****
'   ****    can be anywhere from 0 to 255.                               ****
'   ****                                                                 ****
'   ****    The 'Select Case' below is used to make the colors even      ****
'   ****    more random by randomly selecting which of the 3 color       ****
'   ****    groups to include in each display.  The possible choices     ****
'   ****    in this program are:                                         ****
'   ****                                                                 ****
'   ****    Red, Green, Blue                                             ****
'   ****    Red & Blue                                                   ****
'   ****    Red & Green                                                  ****
'   ****    Green & Blue                                                 ****
'   ****    Red                                                          ****
'   ****    Green                                                        ****
'   ****    Blue                                                         ****
'   ****                                                                 ****
'   *************************************************************************
'

    Select Case lngCorner
    Case 0
        With TriVert(lngLoop1)
            .Red = LongToUShort(lngRGB(1))
            .Green = LongToUShort(lngRGB(2))
            .Blue = LongToUShort(lngRGB(3))
        End With             '  TriVert(lngLoop1)

    Case 1
        With TriVert(lngLoop1)
            .Red = LongToUShort(lngRGB(1))
            .Green = 0&
            .Blue = 0&
        End With             '  TriVert(lngLoop1)

    Case 2
        With TriVert(lngLoop1)
            .Red = 0&
            .Green = LongToUShort(lngRGB(2))
            .Blue = 0&
        End With             '  TriVert(lngLoop1)

    Case 3
        With TriVert(lngLoop1)
            .Red = 0&
            .Green = 0&
            .Blue = LongToUShort(lngRGB(3))
        End With             '  TriVert(lngLoop1)

    Case 4
        With TriVert(lngLoop1)
            .Red = LongToUShort(lngRGB(1))
            .Green = 0&
            .Blue = LongToUShort(lngRGB(3))
        End With             '  TriVert(lngLoop1)

    Case 5
        With TriVert(lngLoop1)
            .Red = 0&
            .Green = LongToUShort(lngRGB(2))
            .Blue = LongToUShort(lngRGB(3))
        End With             '  TriVert(lngLoop1)

    Case 6
        With TriVert(lngLoop1)
            .Red = LongToUShort(lngRGB(1))
            .Green = LongToUShort(lngRGB(2))
            .Blue = 0&
        End With             '  TriVert(lngLoop1)
    End Select

    lngColorNum(lngNum) = lngColorNum(lngNum) + lngUorD

End Sub

Private Sub subDoubleTriangle()

    Me.AutoRedraw = False

'   Select new random colors
    For lngLoop1 = 1 To 3
        lngColorNum(lngLoop1) = Int(128 * Rnd) + 128
        lngRGB(lngLoop1) = CLng(lngColorNum(lngLoop1)) * 256
    Next lngLoop1

'       Choose which colors are to be used in the display.
    For lngLoop1 = 0 To 3
        lngCorner = Int(7 * Rnd)
        lngCornerNum(lngLoop1) = lngCorner
        subColor
    Next lngLoop1

    Me.AutoRedraw = True

    If lngDiagonal < 3 Then
        lngDiagonal = Int(2 * Rnd) + 1
    End If

'
'   *************************************************************************
'   ****                                                                 ****
'   ****    This routine is where I draw the Triangles.                  ****
'   ****                                                                 ****
'   *************************************************************************
'

    Select Case lngDiagonal
    Case 1  '   This draws 2 Triangles facing each other
        With GradTriangle
            .Vertex1 = 0
            .Vertex2 = 1
            .Vertex3 = 2
            GradientFillTriangle Me.hdc, TriVert(0), 4, GradTriangle, 1, GRADIENT_FILL_TRIANGLE
            .Vertex1 = 2
            .Vertex2 = 3
            .Vertex3 = 0
        End With             '  GradTriangle

    Case 2  '   This draws 2 Triangles facing each other
        With GradTriangle
            .Vertex1 = 1
            .Vertex2 = 2
            .Vertex3 = 3
            GradientFillTriangle Me.hdc, TriVert(0), 4, GradTriangle, 1, GRADIENT_FILL_TRIANGLE
            .Vertex1 = 3
            .Vertex2 = 0
            .Vertex3 = 1
        End With             '  GradTriangle

    Case 3  '   This draws a single Triangle
        With GradTriangle
            .Vertex1 = 1
            .Vertex2 = 2
            .Vertex3 = 3
        End With             '  GradTriangle
    End Select

'
'   *************************************************************************
'   ****                                                                 ****
'   ****    This is the draw command for the 2nd of 2 Triangles, or      ****
'   ****    the single Triangle, depending on which Triangle is          ****
'   ****    currently being drawn.                                       ****
'   ****                                                                 ****
    GradientFillTriangle Me.hdc, TriVert(0), 4, GradTriangle, 1, GRADIENT_FILL_TRIANGLE
'   ****                                                                 ****
'   ****    I save two statements by moving this command outside the     ****
'   ****    'Select Case'.  I know it's not a big deal, but I try to     ****
'   ****    slim down the code when/where I can.                         ****
'   ****                                                                 ****
'   *************************************************************************
'

    Me.Refresh

'   Display the (Time left) for the current routine.
    picSpeed.Cls
    picSpeed.Print (lngGradientStart + lngHowLong(lngSubNum)) - lngGradientStop

End Sub

Private Sub subDoubleTriFullScr()

'   Set up the TriVertex info for the full screen display
    lng4Corner(0) = 0
    lng4Corner(1) = 0
    lng4Corner(2) = Me.ScaleWidth
    lng4Corner(3) = Me.ScaleHeight

'   Draw a screen-sized gradient
    subFourCorners

    lngSubNum = 1
    lngClockPrev = 999
    lngGradientStart = Timer
    lngGradientStop = Timer

    Do
        subDoubleTriangle
        subSlowDown

        lngGradientStop = Timer

        If blnSpeed Then
            lngGradientStart = lngGradientStop - lngSpeedDifference
        End If
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Note:                                                        ****
'   ****                                                                 ****

    Loop Until lngGradientStop > lngGradientStart + lngHowLong(lngSubNum)

'   ****    The variable (lngHowLong()) in the line above, is the        ****
'   ****    number of seconds that this particular display will run.     ****
'   ****    Each of the following subRoutines has it's own               ****
'   ****    identification number (lngSubNum) as shown below. You        ****
'   ****    may, of course, change any of these numbers to make a        ****
'   ****    particular display run longer or shorter.                    ****
'   ****                                                                 ****
'   ****    1 - subDoubleTriFullScr - Full Screen Gradient               ****
'   ****    2 - subDoubleTriMini    - Smaller rectangular Gradients      ****
'   ****    3 - subGradientCircles  - Gradient Circles                   ****
'   ****    4 - subTriangles        - Gradient Triangles                 ****
'   ****                                                                 ****
'   *************************************************************************
'
    Cls
    subDoubleTriMini

End Sub

Private Sub subDoubleTriMini()

    lngSubNum = 2
    lngClockPrev = 999
    lngGradientStart = Timer

    Do
'       Set up the TriVertex info for the mini display
        lng4Corner(0) = (Me.ScaleWidth * 0.5) * Rnd
        lng4Corner(1) = (Me.ScaleHeight * 0.5) * Rnd
        lng4Corner(2) = ((Me.ScaleWidth * 0.5) * Rnd) + (Me.ScaleWidth * 0.5)
        lng4Corner(3) = ((Me.ScaleHeight * 0.5) * Rnd) + (Me.ScaleHeight * 0.5)

'       Draw the mini-sized gradients
        subFourCorners
        subDoubleTriangle
        subSlowDown

        lngGradientStop = Timer

        If blnSpeed Then
            lngGradientStart = lngGradientStop - lngSpeedDifference
        End If
    Loop Until lngGradientStop > lngGradientStart + lngHowLong(lngSubNum)

    Cls
    subGradientCircles

End Sub

Private Sub subFourCorners()

    For lngLoop1 = 1 To 3
        lngColorNum(lngLoop1) = Int(128 * Rnd) + 128
        lngRGB(lngLoop1) = CLng(lngColorNum(lngLoop1)) * 256
    Next lngLoop1

'   Define the four corners of the form.
    With TriVert(0)
        .X = lng4Corner(0)
        .Y = lng4Corner(1)
        .Alpha = 0&
    End With  '  TriVert(0)

    With TriVert(1)
        .X = lng4Corner(2)
        .Y = lng4Corner(1)
        .Alpha = 0&
    End With  '  TriVert(1)

    With TriVert(2)
        .X = lng4Corner(2)
        .Y = lng4Corner(3)
        .Alpha = 0&
    End With  '  TriVert(2)

    With TriVert(3)
        .X = lng4Corner(0)
        .Y = lng4Corner(3)
        .Alpha = 0&
    End With  '  TriVert(3)

End Sub

Private Sub subFreeze()

'
'   *************************************************************************
'   ****                                                                 ****
'   ****    The Pause routine is activated by pressing a Mouse           ****
'   ****    button.  Press a button once to Freeze (Pause) the           ****
'   ****    screen.  Press a button again to Thaw the screen.            ****
'   ****                                                                 ****
'   ****    The local variables (lngPauseStop & lngPauseStart) allow     ****
'   ****    me to freeze the clock while the screen is being paused.     ****
'   ****                                                                 ****
'   *************************************************************************
'

Dim lngPauseStart As Long
Dim lngPauseStop  As Long

    lngPauseStart = Timer

    Do While blnPause
        DoEvents             '  Yield for other processing.
    Loop

    lngPauseStop = Timer
    lngGradientStart = lngGradientStart + (lngPauseStop - lngPauseStart)

End Sub

Private Sub subGradientCircles()

'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    GradientCircles                                              ****
'   ****                                                                 ****
'   ****    Written by:  Randy Giese                                     ****
'   ****    Written at:  www.RandyGrams.com                              ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    Thank you Robert Rayment !!!!!                               ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    I received a great deal of help from Robert Rayment in       ****
'   ****    debugging this program.  Robert; I am deeply indebted to     ****
'   ****    you for all of your assistance.                              ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    I wrote GradientCircles a few years ago as part of a         ****
'   ****    larger project.  It has always been one of my favorite       ****
'   ****    Graphics programs.  But then I'm a sucker for colorful       ****
'   ****    random graphics.  I've seen some other interesting           ****
'   ****    graphic displays on PSC that are based on a specific         ****
'   ****    formula or pattern, so that once you see it.... you've       ****
'   ****    seen it.  In other words, it does one thing nicely, but      ****
'   ****    that's the only thing it can do.                             ****
'   ****    GradientCircles is my example of just the opposite.  It      ****
'   ****    is colorful and always changing.  The colors of the          ****
'   ****    circles are always different, the circle size always         ****
'   ****    changes and sometimes the circle builds from the center      ****
'   ****    out and other times it shrinks from the outer edge to the    ****
'   ****    center.                                                      ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'

Dim lngCenterCol As Long
Dim lngCenterRow As Long
Dim lngCircleDir As Long
Dim lngFrom      As Long
Dim lngRadius    As Long
Dim lngTo        As Long

    Me.AutoRedraw = False
    picSpeed.Cls
    picSpeed.AutoRedraw = False
    lngSubNum = 3
    lngClockPrev = 999
    lngGradientStart = Timer

    Do
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Determine the Size and Location of the Circle.               ****
'   ****                                                                 ****
'   *************************************************************************
'
        lngCenterCol = ((ScaleWidth * 0.75) * Rnd) + (ScaleWidth * 0.125)     '  Starting Column
        lngCenterRow = ((ScaleHeight * 0.75) * Rnd) + (ScaleHeight * 0.125)   '  Starting Row
        lngCircleDir = Sgn(Rnd - 0.5)       '  Initial Direction of Color Change

'       This makes the radius anywhere from 1/4 to 1/2 screen height.
        lngRadius = (ScaleHeight * 0.25 * Rnd) + (ScaleHeight * 0.25) '  Radius of Circle

'       Determine whether the Circle Shrinks or Grows
        If lngCircleDir > 0 Then
'           The Circle Grows
            lngFrom = 1
            lngTo = lngRadius
        Else  '  Not lngCircleDir
'           The Circle Shrinks
            lngFrom = lngRadius
            lngTo = 1
        End If
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Determine the Circle's colors.                               ****
'   ****                                                                 ****
'   *************************************************************************
'
        subGradientColor     '  Setup the initial Color Values and direction of change
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Draw the Circles.                                            ****
'   ****                                                                 ****
'   *************************************************************************
'
        For lngLoop2 = lngFrom To lngTo Step lngCircleDir
            QueryPerformanceCounter curStart               '  Get the start time

            Me.ForeColor = RGB(lngRGB(1), lngRGB(2), lngRGB(3))
            Ellipse Me.hdc, lngCenterCol + lngLoop2, lngCenterRow + lngLoop2, lngCenterCol - lngLoop2, lngCenterRow - lngLoop2

'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Comment out this statement:                                  ****
'   ****                                                                 ****
            subSlowDown
'   ****                                                                 ****
'   ****    to make GradientCircles run at max speed.                    ****
'   ****                                                                 ****
'   *************************************************************************
'
            subChangeColors
        Next lngLoop2

        lngGradientStop = Timer

        If blnSpeed Then
            lngGradientStart = lngGradientStop - lngSpeedDifference
        End If
    Loop Until lngGradientStop > lngGradientStart + lngHowLong(lngSubNum)

    picSpeed.AutoRedraw = True
    Cls
    subTriangles

End Sub

Private Sub subGradientColor()

'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Get the initial color and direction of the color change.     ****
'   ****                                                                 ****
'   ****    lngPorM(x) = Plus 1 or Minus 1.                              ****
'   ****                                                                 ****
'   ****    When a color reaches 255, lngPorM will change to a '-1'      ****
'   ****    and the lngRGB numbers will start getting smaller...         ****
'   ****    until it reaches '0' and starts the whole process over       ****
'   ****    again.                                                       ****
'   ****                                                                 ****
'   *************************************************************************
'

    For lngCntr = 1 To 3     '  Setup the initial Color Values and direction of change
        lngPorM(lngCntr) = (Sgn(Rnd - 0.5)) '  A Plus 1 or Minus 1 - lngPorM neans Plus or Minus
        lngRGB(lngCntr) = Int(254 * Rnd) + 1
    Next lngCntr

End Sub

Private Sub subIncrementColors()

    lngPorM(lngCntr) = -lngPorM(lngCntr)    '  Reverse the color direction
    If lngCntr < 3 Then      '  If it's not the last color group (Green),
        lngRGB(lngCntr + 1) = lngRGB(lngCntr + 1) + lngPorM(lngCntr + 1)  '  then add 1 to next color group
    End If

    picSpeed.Cls
    picSpeed.Print (lngGradientStart + lngHowLong(lngSubNum)) - lngGradientStop

End Sub

Private Sub subQuit()

    On Error Resume Next

    Open strWinDir For Output As #1

    For lngLoop1 = 1 To 4
'       Save the current Speed values
        Print #1, lngSpeed(lngLoop1)
'       Then save the default Time values
        Print #1, lngHowLong(lngLoop1)
    Next lngLoop1

    Close

    For lngCntr = (Forms.Count - 1) To 0 Step -1
        Unload Forms(lngCntr)               '  Unloads ALL forms
    Next lngCntr

    Set GradientCombo = Nothing
    End

    On Error GoTo 0

End Sub

Private Sub subSlowDown()

'
'   *************************************************************************
'   ****                                                                 ****
'   ****    subSlowDown causes a very brief pause between the drawing    ****
'   ****    of each ring of the circle.  GradientCircles runs way too    ****
'   ****    fast on my PC without some sort of Slow Down routine.  If    ****
'   ****    you think it should run faster, try commenting out the       ****
'   ****    subSlowDown statement right after the 'Ellipse' statement    ****
'   ****    in 'subGradientCircles'.  If that looks better to you        ****
'   ****    then you can always delete this SubRoutine later.            ****
'   ****                                                                 ****
'   *************************************************************************
'

    QueryPerformanceCounter curStart        '  Get the start time

    Do While curEnd < curStart + curFreq(lngSubNum)
        QueryPerformanceCounter curEnd      '  Get the end time

        If GetInputState Then               '  Yield for other processing.
            DoEvents
        End If
    Loop

End Sub

Private Sub subTriangles()

    lngSubNum = 4
    lngDiagonal = 3
    lngClockPrev = 999
    lngGradientStart = Timer

    Do
        For lngLoop1 = 1 To 3
            lngColorNum(lngLoop1) = Int(128 * Rnd) + 128
            lngRGB(lngLoop1) = CLng(lngColorNum(lngLoop1)) * 256
        Next lngLoop1

'       Set up the TriVertex info for the mini display
        lng4Corner(1) = ScaleWidth * Rnd
        lng4Corner(2) = ScaleHeight * Rnd
        lng4Corner(3) = ScaleWidth * Rnd
        lng4Corner(4) = ScaleHeight * Rnd
        lng4Corner(5) = ScaleWidth * Rnd
        lng4Corner(6) = ScaleHeight * Rnd

'       Define the Three corners of the form.
        With TriVert(1)
            .X = lng4Corner(1)
            .Y = lng4Corner(2)
            .Alpha = 0&
        End With  '  TriVert(1)

        With TriVert(2)
            .X = lng4Corner(3)
            .Y = lng4Corner(4)
            .Alpha = 0&
        End With  '  TriVert(2)

        With TriVert(3)
            .X = lng4Corner(5)
            .Y = lng4Corner(6)
            .Alpha = 0&
        End With  '  TriVert(3)

'       Draw the mini-sized Triangles
        subDoubleTriangle
        subSlowDown

        lngGradientStop = Timer

        If blnSpeed Then
            lngGradientStart = lngGradientStop - lngSpeedDifference
        End If
    Loop Until lngGradientStop > lngGradientStart + lngHowLong(lngSubNum)

    Cls
    lngDiagonal = 0
    subDoubleTriFullScr

End Sub

Private Sub txtTime_Click()

    txtTime.Text = vbNullString

End Sub

Private Sub txtTime_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        lblClose_Click
    End If

End Sub
