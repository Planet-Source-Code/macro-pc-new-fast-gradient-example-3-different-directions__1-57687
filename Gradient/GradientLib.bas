Attribute VB_Name = "GradientLib"
Public Type RECT_DOUBLE
Left As Double
Top As Double
Right As Double
Bottom As Double
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Public Const BITSPIXEL As Long = 12
Public Const PLANES As Long = 14
Public Const RED_BITS As Long = 255
Public Const GREEN_BITS As Long = 65280
Public Const BLUE_BITS As Long = 16711680
Public Const BLUE_COLOR As Long = &H10000
Public Const GREEN_COLOR As Long = &H100
Public Const SRCCOPY = &HCC0020


Public Function CreateGradient(RefForm As Form, Direction As Long, _
    StartingColor As Long, EndingColor As Long) As Boolean

On Error GoTo err_CreateGradient

Dim dblCurrentRed As Double
Dim dblCurrentBlue As Double
Dim dblCurrentGreen As Double
Dim dblBlueInc As Double
Dim dblGreenInc As Double
Dim dblRedInc As Double
Dim dblX As Double
Dim dblY As Double
Dim lngBitsPerPixel As Long
Dim lngBrushHandle As Long
Dim lngColorBits As Long
Dim lngColorDiff As Long
Dim lngCurrentRed As Long
Dim lngCurrentBlue As Long
Dim lngCurrentGreen As Long
Dim lngCurrentScale As Long
Dim lngEndColor As Long
Dim lngFormHeight As Long
Dim lngFormWidth As Long
Dim lngPlanes As Long
Dim lngRegionCount As Long
Dim lngRet As Long
Dim lngStartColor As Long
Dim lngC As Long
Dim udtArealng As RECT
Dim udtAreadbl As RECT_DOUBLE

If Not RefForm Is Nothing Then

    lngStartColor = EndingColor
    lngEndColor = StartingColor
    
    lngBitsPerPixel = GetDeviceCaps(RefForm.hdc, BITSPIXEL)
    lngPlanes = GetDeviceCaps(RefForm.hdc, PLANES)
    lngColorBits = (lngBitsPerPixel * lngPlanes)
    
    If lngcolorbis > 15 Then
        lngRegionCount = 256
    Else
        lngRegionCount = 32
    End If
    
    With RefForm
        
        lngCurrentScale = .ScaleMode
        .ScaleMode = vbPixels
        lngFormHeight = .ScaleHeight
        lngFormWidth = .ScaleWidth
        .ScaleMode = lngCurrentScale
        
    End With
    
    dblRedInc = CDbl((-1 * ((lngEndColor And RED_BITS) - _
        (lngstartscale And RED_BITS))) / lngRegionCount)
        
    dblGreenInc = CDbl((-1 * (((lngEndColor And _
    GREEN_BITS) / RED_BITS) - ((lngStartColor And GREEN_BITS) / _
    RED_BITS))) / lngRegionCount)
    
    dblBlueInc = CDbl((-1 * (((lngEndColor And BLUE_BITS) / _
    GREEN_BITS) - ((lngStartColor And BLUE_BITS) / _
    GREEN_BITS))) / lngRegionCount)
    
    dblCurrentRed = lngEndColor And RED_BITS
    lngCurrentRed = IIf(dblCurrentRed > 255, 255, _
        IIf(dblCurrentRed < 0, 0, CLng(dblCurrentRed)))
    dblCurrentGreen = (lngEndColor And GREEN_BITS) / RED_BITS
    lngCurrentGreen = IIf(dblCurrentGreen > 255, 255, _
        IIf(dblCurrentGreen < 0, 0, CLng(dblCurrentGreen)))
    dblCurrentBlue = (lngEndColor And BLUE_BITS) / GREEN_BITS
    lngCurrentBlue = IIf(dblCurrentBlue > 255, 255, _
        IIf(dblCurrentBlue < 0, 0, CLng(dblCurrentBlue)))
    
    dblX = lngFormWidth / lngRegionCount
    dblY = lngFormHeight / lngRegionCount
    
    udtArealng.Left = 0
    udtAreadbl.Left = 0
    udtArealng.Top = 0
    udtAreadbl.Top = 0
    udtArealng.Right = lngFormWidth
    udtAreadbl.Right = lngFormWidth
    udtArealng.Bottom = lngFormHeight
    udtAreadbl.Bottom = lngFormHeight
    
    For lngC = 0 To (lngRegionCount - 1)
        lngBrushHandle = CreateSolidBrush(RGB _
        (lngCurrentRed, lngCurrentGreen, lngCurrentBlue))
        
        If Direction = 0 Then

            'diagonal
            udtAreadbl.Top = udtAreadbl.Bottom - dblY
            udtArealng.Top = CLng(udtAreadbl.Top)
            udtAreadbl.Left = 0
            udtArealng.Left = 0
            lngRet = FillRect(RefForm.hdc, udtArealng, _
                lngBrushHandle)
            udtAreadbl.Top = 0
            udtArealng.Top = 0
            udtAreadbl.Left = udtAreadbl.Right - dblX
            udtArealng.Left = CLng(udtAreadbl.Left)
            lngRet = FillRect(RefForm.hdc, udtArealng, _
                lngBrushHandle)
            udtAreadbl.Bottom = udtAreadbl.Bottom - dblY
            udtArealng.Bottom = CLng(udtAreadbl.Bottom)
            udtAreadbl.Right = udtAreadbl.Right - dblX
            udtArealng.Right = CLng(udtAreadbl.Right)
            
    ElseIf Direction = 1 Then
        'vertical
    
        udtAreadbl.Top = udtAreadbl.Bottom - dblY
        udtArealng.Top = CLng(udtAreadbl.Top)
        lngRet = FillRect(RefForm.hdc, udtArealng, _
            lngBrushHandle)
    
        udtAreadbl.Bottom = udtAreadbl.Bottom - dblY
        udtArealng.Bottom = CLng(udtAreadbl.Bottom)
    
    ElseIf Direction = 2 Then
        'horizontal
        udtAreadbl.Left = udtAreadbl.Right - dblX
        udtArealng.Left = CLng(udtAreadbl.Left)
        lngRet = FillRect(RefForm.hdc, udtArealng, _
            lngBrushHandle)
        udtAreadbl.Right = udtAreadbl.Right - dblX
        udtArealng.Right = CLng(udtAreadbl.Right)
        
    Else
    
    End If
    
        lngRet = DeleteObject(lngBrushHandle)
    
        dblCurrentRed = dblCurrentRed + dblRedInc
        lngCurrentRed = IIf(dblCurrentRed > 255, 255, _
            IIf(dblCurrentRed < 0, 0, CLng(dblCurrentRed)))
        dblCurrentGreen = dblCurrentGreen + dblGreenInc
        lngCurrentGreen = IIf(dblCurrentGreen > 255, 255, _
            IIf(dblCurrentGreen < 0, 0, CLng(dblCurrentGreen)))
        dblCurrentBlue = dblCurrentBlue + dblBlueInc
        lngCurrentBlue = IIf(dblCurrentBlue > 255, 255, _
            IIf(dblCurrentBlue < 0, 0, CLng(dblCurrentBlue)))
        
    Next lngC
    
    udtArealng.Top = 0
    udtArealng.Left = 0
    lngBrushHandle = CreateSolidBrush(RGB _
        (lngCurrentRed, lngCurrentGreen, lngCurrentBlue))
    lngRet = FillRect(RefForm.hdc, udtArealng, lngBrushHandle)
    lngRet = DeleteObject(lngBrushHandle)
    
    CreateGradient = True
    
Else

    CreateGradient = False

End If

Exit Function
 
err_CreateGradient:
    
    CreateGradient = False
    
    If Not RefForm Is Nothing Then
        RefForm.Tag = Err.Number & "  " & Err.Description
    End If

End Function

