Attribute VB_Name = "MPrinter"
Option Explicit
'Enum VBRUN.ScaleModeConstants
'    vbUser              =  0
'    vbTwips             =  1
'    vbPoints            =  2
'    vbPixels            =  3
'    vbCharacters        =  4
'    vbInches            =  5
'    vbMillimeters       =  6
'    vbCentimeters       =  7
'    vbHimetric          =  8 'invalid property value for Printer.ScaleMode
'    vbContainerPosition =  9
'    vbContainerSize     = 10
'End Enum
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Public PrinterDpi_VB  As Long
Public PrinterDpi_API As Long
Public ScreenDpi_VB   As Long
Public ScreenDpi_API  As Long

Sub Init()
    PrinterDpi_VB = 1440 / Printer.TwipsPerPixelX
    PrinterDpi_API = Printer_ResolutionDpi
    ScreenDpi_VB = 1440 / Screen.TwipsPerPixelX
    ScreenDpi_API = Screen_ResolutionDpi
End Sub

Public Function Printer_ResolutionDpi() As Long
    Dim dpiX As Long: dpiX = GetDeviceCaps(Printer.hDC, LOGPIXELSX)
    'Dim dpiY As Long: dpiY = GetDeviceCaps(Printer.hDC, LOGPIXELSY)
    Printer_ResolutionDpi = dpiX
End Function

Public Function Screen_ResolutionDpi() As Long
    Dim hDC  As Long:  hDC = GetDC(0)
    Dim dpiX As Long: dpiX = GetDeviceCaps(hDC, LOGPIXELSX)
    'Dim dpiY As Long: dpiY = GetDeviceCaps(hDC, LOGPIXELSY)
    Screen_ResolutionDpi = dpiX
End Function

Public Sub SelectPrinter(ByVal PrinterName As String)
    Dim i As Long
    For i = 0 To Printers.Count - 1
        If UCase(Printers(i).DeviceName) = UCase(PrinterName) Then 'e.g.: "Microsoft Print to PDF"
            Set Printer = Printers(i)
            Exit For
        End If
    Next
End Sub

Public Function Millimeter_Scale(ByVal sm As ScaleModeConstants, ByVal Value As Double) As Double
    'Value in Millimeters, the output is according to the scalemodeconstant
    'Screen.Twips and Printer.Twips are not the same!"
    'e.g.  Screen.TwipsPerPixelX = 15   ' at  96 dpi
    'e.g. Printer.TwipsPerPixelX = 2.4  ' at 600 dpi
    Select Case sm
    'Case ScaleModeConstants.vbUser:        Millimeter_Scale = Value
    Case ScaleModeConstants.vbTwips:       Millimeter_Scale = Value * 1440 / 25.4
    Case ScaleModeConstants.vbPoints:      Millimeter_Scale = Value * 72 / 25.4
    Case ScaleModeConstants.vbPixels:      Millimeter_Scale = Value * 1440 / Printer.TwipsPerPixelX / 25.4
                                                       'a twip is a 1/1440 of an inch

    Case ScaleModeConstants.vbCharacters:  Millimeter_Scale = Value / 10 '???
    
    Case ScaleModeConstants.vbInches:      Millimeter_Scale = Value / 25.4
    Case ScaleModeConstants.vbMillimeters: Millimeter_Scale = Value
    Case ScaleModeConstants.vbCentimeters: Millimeter_Scale = Value / 10
    'Case ScaleModeConstants.vbHimetric:   'ungültiger Eigenschaftswert für Printer
    'Case ScaleModeConstants.vbContainerPosition
    'Case ScaleModeConstants.vbContainerSize
    End Select
End Function

Public Sub PaintPictureW(aPic As StdPicture, ByVal X As Single, ByVal Y As Single, ByVal Width As Single)
    'Prints the picture to the X-Y-position with the width of Width, and scales the height to ratio 1
    Dim W As Single: W = aPic.Width
    Dim H As Single: H = aPic.Height
    Printer.PaintPicture aPic, X, Y, Width, Width * H / W
End Sub

Public Sub PaintPictureH(aPic As StdPicture, ByVal X As Single, ByVal Y As Single, ByVal Height As Single)
    'Prints the picture to the X-Y-position with the height of Height, and scales the width to ratio 1
    Dim W As Single: W = aPic.Width
    Dim H As Single: H = aPic.Height
    Printer.PaintPicture aPic, X, Y, Height * W / H, Height
End Sub

Public Sub PaintPictureFit(aPic As StdPicture, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal OrHeight As Single)
    If Width < OrHeight Then
        PaintPictureH aPic, X, Y, Width 'OrHeight
    Else
        PaintPictureW aPic, X, Y, OrHeight 'Width
    End If
End Sub


'Private Function Millimeter_ToTwips(ByVal mm As Double) As Single
'    Dim dpi    As Single:    dpi = 96   ' dots per inch
'    Dim ppi    As Single:    ppi = 72   'point per inch
'    Dim mmpi   As Single:   mmpi = 25.4 '  mm  per inch
'    Dim TPPX   As Single:   TPPX = Screen.TwipsPerPixelX
'    'Dim sc     As Single. SC 0
'    Millimeter_ToTwips = mm * TPPX * dpi / mmpi 'dpi / ppi * mmpi
'End Function




'Hi Wolfgang,
'
'[q]so einfach ist das leider nicht.[/q]
'ich versteh ehrlichgesagt Deine Probleme nicht.
'
'[q]Bei 100% ist ein Pixel ein Pixel. [/q]
'Japp
'
'[q]Alles was jenseits von diesen 100% ist, wirft Fragen auf[/q]
'Nein! Definition:
'1 Pixel ist die kleinste Einheit eines Bildes bzw des auf dem Monitor dargestellten Bildes.
'Hat z.B. Dein Monitor eine Auflösung von Full-HD also 1920*1080 Pixel
'Dann ist ein Pixel der 1920-ste Teil der Breite des Bildes auf dem Bildschirms.
'Bzw der 1080-ste Teil der Höhe des Bildes auf dem Bildschirm.
'Bei heutigen Flachbildschirmen, ist es soviel ich weiß nicht mehr möglich das Bild zu quetschen oder
'zu stauchen, damit ist ein Pixel immer genau so breit wie hoch.
'Wenn man z.B. einen Screenshot des gesamten Bildschirms macht, das Bild in ein Bildbearbeitungs-
'programm kopiert wie z.B. MS-Paint und dann reinzoomst, dann kommt man irgendwann auf die kleinste
'Einheit auf einen Pixel, kleiner wirds dann einfach nimmer und das wird unter einem Pixel verstanden.
'
'Dadurch dass die Pixeldichte Monitore immer

