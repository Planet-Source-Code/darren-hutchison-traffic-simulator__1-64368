Attribute VB_Name = "modDirectX"
' This module is derived from Jim Camel's DX Rotation program from Planet Source Code
' Some subs and all comments have been removed, and a couple of new subs added in

Option Explicit
Private Const SRCCOPY = &HCC0020
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private DX As New DirectX7
Private DD As DirectDraw7
Private surfPrimary As DirectDrawSurface7
Private surfBackBuffer As DirectDrawSurface7
Private Const PI As Double = 3.14159265358979

Public Function initDirectDraw(ByVal frm As Form, ByVal lngWidth As Long, ByVal lngHeight As Long, lngBitColor As Long) As Boolean

    Dim ddsdTemp As DDSURFACEDESC2
    Dim caps As DDSCAPS2
    
    Set DD = DX.DirectDrawCreate("")

    frm.Show

    DD.SetCooperativeLevel frm.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
    DD.SetDisplayMode lngWidth, lngHeight, lngBitColor, 0, DDSDM_DEFAULT

    ddsdTemp.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsdTemp.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    ddsdTemp.lBackBufferCount = 1
    Set surfPrimary = DD.CreateSurface(ddsdTemp)
    
    caps.lCaps = DDSCAPS_BACKBUFFER
    Set surfBackBuffer = surfPrimary.GetAttachedSurface(caps)
    surfBackBuffer.SetForeColor 16777215
    
End Function

Public Sub rotateSurface(ByRef surfSource As DirectDrawSurface7, surfDestination As DirectDrawSurface7, iAngle As Integer, Optional XDest As Long = 0, Optional Ydest As Long = 0, Optional Transparency As Long = -1)

  Dim ddsdSource As DDSURFACEDESC2
  Dim lngXI As Long, lngYI As Long
  Dim lngXO As Long, lngYO As Long
  Dim rEmpty As RECT
  Dim sngA As Single, SinA As Single, CosA As Single
  Dim dblRMax As Double
  Dim lngColor As Long
  Dim lWidth As Long, lHeight As Long

    sngA = iAngle * PI / 180
    SinA = Sin(sngA)
    CosA = Cos(sngA)

    surfSource.GetSurfaceDesc ddsdSource
    lWidth = ddsdSource.lWidth
    lHeight = ddsdSource.lHeight
    dblRMax = Sqr(lWidth ^ 2 + lHeight ^ 2)
    
    XDest = XDest + lWidth / 2
    Ydest = Ydest + lHeight / 2
    
    surfDestination.Lock rEmpty, ddsdSource, DDLOCK_WAIT, 0
    surfSource.Lock rEmpty, ddsdSource, DDLOCK_WAIT, 0
    For lngXI = -dblRMax To dblRMax
        For lngYI = -dblRMax To dblRMax
            lngXO = lWidth / 2 - (lngXI * CosA + lngYI * SinA)
            lngYO = lHeight / 2 - (lngXI * SinA - lngYI * CosA)
            If lngXO >= 0 And lngYO >= 0 Then
                If lngXO < lWidth And lngYO < lHeight Then
                        lngColor = surfSource.GetLockedPixel(lngXO, lngYO)
                        If lngColor <> Transparency Then
                            surfDestination.SetLockedPixel XDest + lngXI, Ydest + lngYI, lngColor
                        End If
                End If
            End If
        Next lngYI
    Next lngXI
    
    surfSource.Unlock rEmpty
    surfDestination.Unlock rEmpty

End Sub

Public Sub terminateDirectX()
    Set surfBackBuffer = Nothing
    Set surfPrimary = Nothing
    Set DD = Nothing
    Set DX = Nothing
End Sub

Private Function createSurfaceFromFile(DirectDraw As DirectDraw7, ByVal FileName As String, SurfaceDesc As DDSURFACEDESC2) As DirectDrawSurface7

  Dim Picture As StdPicture
  Dim Width As Long, Height As Long
  Dim surface As DirectDrawSurface7
  Dim hdcPicture As Long, hdcSurface As Long
  Dim ddtrans1 As DDCOLORKEY
  
    Set Picture = LoadPicture(FileName)

    Width = CLng((Picture.Width * 0.001) * 567 / Screen.TwipsPerPixelX)
    Height = CLng((Picture.Height * 0.001) * 567 / Screen.TwipsPerPixelY)
    
    With SurfaceDesc
        If .lFlags = 0 Then .lFlags = DDSD_CAPS
        .lFlags = .lFlags Or DDSD_WIDTH Or DDSD_HEIGHT
        If .ddsCaps.lCaps = 0 Then .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        If .lWidth = 0 Then .lWidth = Width
        If .lHeight = 0 Then .lHeight = Height
    End With

    Set surface = DirectDraw.CreateSurface(SurfaceDesc)
    hdcPicture = CreateCompatibleDC(0)
    SelectObject hdcPicture, Picture.Handle
    hdcSurface = surface.GetDC
    StretchBlt hdcSurface, 0, 0, SurfaceDesc.lWidth, SurfaceDesc.lHeight, hdcPicture, 0, 0, Width, Height, SRCCOPY
    
    surface.ReleaseDC hdcSurface

    ddtrans1.low = vbMagenta
    ddtrans1.high = ddtrans1.low
    surface.SetColorKey DDCKEY_SRCBLT, ddtrans1

    DeleteDC hdcPicture
    Set Picture = Nothing
    Set createSurfaceFromFile = surface
    Set surface = Nothing

End Function

Public Function loadDirectXSurface(Optional FileName As String, Optional Width As Long, Optional Height As Long) As DirectDrawSurface7

  Dim ddsd1 As DDSURFACEDESC2

    If Len(FileName) > 0 Then
        If Width > 0 Then
            ddsd1.lFlags = DDSD_CAPS
            ddsd1.lHeight = Height
            ddsd1.lWidth = Width
            Set loadDirectXSurface = createSurfaceFromFile(DD, FileName, ddsd1)
          Else
            ddsd1.lFlags = DDSD_CAPS
            Set loadDirectXSurface = createSurfaceFromFile(DD, FileName, ddsd1)
        End If
    Else
        ddsd1.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        ddsd1.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE
        ddsd1.lHeight = Height
        ddsd1.lWidth = Width
        Set loadDirectXSurface = DD.CreateSurface(ddsd1)
    End If
End Function

Public Sub bltToBack(destRect As RECT, Source As DirectDrawSurface7, sourceRect As RECT, flags As CONST_DDBLTFLAGS)
    surfBackBuffer.Blt destRect, Source, sourceRect, flags
End Sub

Public Sub textToBack(x As Long, y As Long, text As String)
    surfBackBuffer.DrawText x, y, text, False
End Sub

Public Sub flipPrimary()
    surfPrimary.Flip Nothing, DDFLIP_WAIT
End Sub

Public Function GetMaxRotateSize(surface As DirectDrawSurface7) As Long
    Dim ddsdSource As DDSURFACEDESC2
    
    surface.GetSurfaceDesc ddsdSource
    GetMaxRotateSize = Sqr(ddsdSource.lWidth ^ 2 + ddsdSource.lHeight ^ 2)

End Function
Public Sub SetTexttoBack(color As Long)

surfBackBuffer.SetForeColor color

End Sub
Public Sub SetTextSize(name As Variant)

surfBackBuffer.SetFont name

End Sub
Public Sub AddColorKey(surface As DirectDrawSurface7, ColorKey As DDCOLORKEY, low As Long, high As Long)

ColorKey.low = low
ColorKey.high = high
surface.SetColorKey DDCKEY_SRCBLT, ColorKey

End Sub

Public Sub rotate90(ByRef surfSource As DirectDrawSurface7, surfDestination As DirectDrawSurface7, lngAngle As Long, Optional XDest As Long = 0, Optional Ydest As Long = 0, Optional Transparency As Long = -1)
Dim ddsdSource As DDSURFACEDESC2
Dim lWidth As Long, lHeight As Long
Dim lngXI As Long, lngYI As Long
Dim lngColor As Long, rEmpty As RECT

    surfSource.GetSurfaceDesc ddsdSource
    lWidth = ddsdSource.lWidth
    lHeight = ddsdSource.lHeight
    lngAngle = lngAngle Mod 4
    
    If lngAngle = 0 Then
        Dim r As RECT, ddColKey As DDCOLORKEY, ddColBack As DDCOLORKEY
        r.Right = lWidth: r.Bottom = lHeight
        If Transparency > -1 Then
            surfSource.GetColorKey DDCKEY_SRCBLT, ddColBack
            ddColKey.low = Transparency: ddColKey.high = Transparency
            surfSource.SetColorKey DDCKEY_SRCBLT, ddColKey
            surfDestination.Blt r, surfSource, r, DDBLT_KEYSRC
            surfSource.SetColorKey DDCKEY_SRCBLT, ddColBack
        Else
            surfDestination.Blt r, surfSource, r, DDBLT_WAIT
        End If
        Exit Sub
    End If

    surfDestination.Lock rEmpty, ddsdSource, DDLOCK_WAIT, 0
    surfSource.Lock rEmpty, ddsdSource, DDLOCK_WAIT, 0
    
    For lngYI = 0 To lHeight
        For lngXI = 0 To lWidth
                lngColor = surfSource.GetLockedPixel(lngXI, lngYI)
                If lngColor <> Transparency Then
                    Select Case lngAngle
                    Case 1
                        Call surfDestination.SetLockedPixel(lWidth - lngYI - 1, lngXI, lngColor)
                    Case 2
                        Call surfDestination.SetLockedPixel(lHeight - lngXI - 1, lWidth - lngYI - 1, lngColor)
                    Case 3
                        Call surfDestination.SetLockedPixel(lngYI + 1, lHeight - lngXI - 1, lngColor)
                    End Select
                End If
        Next lngXI
    Next lngYI
      
    surfSource.Unlock rEmpty
    surfDestination.Unlock rEmpty
      
End Sub

