Attribute VB_Name = "modGradient"

Option Explicit

Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    These are Public Variables and/or Functions that are         ****
'   ****    needed on more than 1 form.                                  ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'

Public blnSpeed                As Boolean

Public curFreq(1 To 4)         As Currency
Public curFrequency(1 To 4)    As Currency

Public lngGradientStart        As Long
Public lngGradientStop         As Long
Public lngSpeed(1 To 4)        As Long
Public lngSpeedDifference      As Long
Public lngSubNum               As Long

Public Sub subSpeed()

    Select Case lngSpeed(lngSubNum)
    Case Is < 0
'       Slow down the circles.
        curFreq(lngSubNum) = curFrequency(lngSubNum) * Abs(lngSpeed(lngSubNum))
    Case 0
'       Default speed of the circles.
        curFreq(lngSubNum) = curFrequency(lngSubNum)
    Case Is > 0
'       Speed up the circles.
        curFreq(lngSubNum) = curFrequency(lngSubNum) / lngSpeed(lngSubNum)
    End Select

End Sub
':)Code Fixer V3.0.9 (5/22/2005 8:19:58 PM) 29 + 0 = 29 Lines Thanks Ulli for inspiration and lots of code.
