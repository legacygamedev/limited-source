Attribute VB_Name = "modWeather"
Option Explicit
Type DropRainRec
    x As Long
    y As Long
    Randomized As Boolean
    speed As Byte
End Type

' Used for atmosphere
Public GameWeather As Long
Public GameTime As Long
Public RainIntensity As Long

Public MAX_RAINDROPS As Long
Public BLT_RAIN_DROPS As Long
Public DropRain() As DropRainRec

Public BLT_SNOW_DROPS As Long
Public DropSnow() As DropRainRec


Sub BltWeather()
Dim i As Long, c As Long

    Call DD_BackBuffer.SetForeColor(RGB(0, 0, 250))
    
    If GameWeather = WEATHER_RAINING Or GameWeather = WEATHER_THUNDER Then
        For i = 1 To MAX_RAINDROPS
            If DropRain(i).Randomized = False Then
                If frmMirage.tmrRainDrop.Enabled = False Then
                    BLT_RAIN_DROPS = 1
                    frmMirage.tmrRainDrop.Enabled = True
                    If frmMirage.tmrRainDrop.Tag = "" Then
                        frmMirage.tmrRainDrop.Interval = 200
                        frmMirage.tmrRainDrop.Tag = "123"
                    End If
                End If
            End If
        Next i
    ElseIf GameWeather = WEATHER_SNOWING Then
        For i = 1 To MAX_RAINDROPS
            If DropSnow(i).Randomized = False Then
                If frmMirage.tmrSnowDrop.Enabled = False Then
                    BLT_SNOW_DROPS = 1
                    frmMirage.tmrSnowDrop.Enabled = True
                    If frmMirage.tmrSnowDrop.Tag = "" Then
                        frmMirage.tmrSnowDrop.Interval = 200
                        frmMirage.tmrSnowDrop.Tag = "123"
                    End If
                End If
            End If
        Next i
    Else
        If BLT_RAIN_DROPS > 0 And BLT_RAIN_DROPS <= RainIntensity Then
            Call ClearRainDrop(BLT_RAIN_DROPS)
        End If
        frmMirage.tmrRainDrop.Tag = ""
    End If
    
    For i = 1 To MAX_RAINDROPS
        If Not ((DropRain(i).x = 0) Or (DropRain(i).y = 0)) Then
            DropRain(i).x = DropRain(i).x + DropRain(i).speed
            DropRain(i).y = DropRain(i).y + DropRain(i).speed
            Call DD_BackBuffer.DrawLine((PIC_X * (MAX_MAPX + 1)) + ((DropRain(i).x)), (PIC_Y * (MAX_MAPY + 1)) + ((DropRain(i).y)), (PIC_X * (MAX_MAPX + 1)) + (DropRain(i).x + DropRain(i).speed), (PIC_Y * (MAX_MAPY + 1)) + (DropRain(i).y + DropRain(i).speed))
            If (DropRain(i).x > (MAX_MAPX + 1) * PIC_X) Or (DropRain(i).y > (MAX_MAPY + 1) * PIC_Y) Then
                DropRain(i).Randomized = False
            End If
        End If
    Next i
    
    rec.Top = 0
    rec.Bottom = rec.Top + PIC_Y

    For i = 1 To MAX_RAINDROPS

        If Not ((DropSnow(i).x = 0) Or (DropSnow(i).y = 0)) Then
            rec.Left = Rand(0, 0) * PIC_X
            rec.Right = rec.Left + PIC_X
            
            DropSnow(i).x = DropSnow(i).x + DropSnow(i).speed
            DropSnow(i).y = DropSnow(i).y + DropSnow(i).speed
            Call DD_BackBuffer.BltFast((PIC_X * (MAX_MAPX + 1)) + (DropSnow(i).x + DropSnow(i).speed), (PIC_Y * (MAX_MAPY + 1)) + (DropSnow(i).y + DropSnow(i).speed), DD_WeatherSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            If (DropSnow(i).x > (MAX_MAPX + 1) * PIC_X) Or (DropSnow(i).y > (MAX_MAPY + 1) * PIC_Y) Then
                DropSnow(i).Randomized = False
            End If
        End If
    Next i
        
    ' If it's thunder, make the screen randomly flash white
    If GameWeather = WEATHER_THUNDER Then
        If Int((100 - 1 + 1) * Rnd) + 1 = 8 Then
            DD_BackBuffer.SetFillColor RGB(255, 255, 255)
            Call PlaySound("Thunder.wav")
            Call DD_BackBuffer.DrawBox((PIC_X * (MAX_MAPX + 1)), (PIC_Y * (MAX_MAPY + 1)), (PIC_X * (MAX_MAPX + 1)) + ((MAX_MAPX + 1) * PIC_X), (PIC_Y * (MAX_MAPY + 1)) + ((MAX_MAPY + 1) * PIC_Y))
        End If
    End If
End Sub

Sub RNDRainDrop(ByVal RDNumber As Long)
Start:
    DropRain(RDNumber).x = Int((((MAX_MAPX + 1) * PIC_X) * Rnd) + 1)
    DropRain(RDNumber).y = Int((((MAX_MAPY + 1) * PIC_Y) * Rnd) + 1)
    If (DropRain(RDNumber).y > (MAX_MAPY + 1) * PIC_Y / 4) And (DropRain(RDNumber).x > (MAX_MAPX + 1) * PIC_X / 4) Then GoTo Start
    DropRain(RDNumber).speed = Int((10 * Rnd) + 6)
    DropRain(RDNumber).Randomized = True
End Sub

Sub ClearRainDrop(ByVal RDNumber As Long)
On Error Resume Next
    DropRain(RDNumber).x = 0
    DropRain(RDNumber).y = 0
    DropRain(RDNumber).speed = 0
    DropRain(RDNumber).Randomized = False
End Sub

Sub RNDSnowDrop(ByVal RDNumber As Long)
Start:
    DropSnow(RDNumber).x = Int((((MAX_MAPX + 1) * PIC_X) * Rnd) + 1)
    DropSnow(RDNumber).y = Int((((MAX_MAPY + 1) * PIC_Y) * Rnd) + 1)
    If (DropSnow(RDNumber).y > (MAX_MAPY + 1) * PIC_Y / 4) And (DropSnow(RDNumber).x > (MAX_MAPX + 1) * PIC_X / 4) Then GoTo Start
    DropSnow(RDNumber).speed = Int((10 * Rnd) + 6)
    DropSnow(RDNumber).Randomized = True
End Sub

Sub ClearSnowDrop(ByVal RDNumber As Long)
On Error Resume Next
    DropSnow(RDNumber).x = 0
    DropSnow(RDNumber).y = 0
    DropSnow(RDNumber).speed = 0
    DropSnow(RDNumber).Randomized = False
End Sub
