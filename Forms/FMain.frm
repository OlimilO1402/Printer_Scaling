VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Printer Scaling"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6375
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "FMain"
   ScaleHeight     =   8895
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   8070
      Left            =   120
      Picture         =   "FMain.frx":1782
      ScaleHeight     =   8010
      ScaleWidth      =   6105
      TabIndex        =   3
      Top             =   720
      Width           =   6165
   End
   Begin VB.CommandButton BtnPrintPdf 
      Caption         =   "Print to pdf"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label LblPrinterDpi 
      AutoSize        =   -1  'True
      Caption         =   "Printer resolution:"
      Height          =   195
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label LblScreenDpi 
      AutoSize        =   -1  'True
      Caption         =   "Screen resolution:"
      Height          =   195
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1275
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    LblScreenDpi.Caption = LblScreenDpi.Caption & " " & MPrinter.Screen_ResolutionDpi & " dpi"
    LblPrinterDpi.Caption = LblPrinterDpi.Caption & MPrinter.Printer_ResolutionDpi & " dpi"
End Sub

'For drawing graphics onto the printer use the methods Circle, Cls, Line, PaintPicture, Point, Print und PSet
'for printing text use the Print-method

Private Sub BtnPrintPdf_Click()
Try: On Error GoTo Catch
    'in the following rectangle all values are given in millimeters
    'the rectangle should start at 50 mm from the left and 70 mm from the top, and it should be 110 mm wide and 60 mmm high
    Dim r0 As MPGeom.AARect: r0 = MPGeom.New_AARect(MPGeom.New_Point(50, 70), MPGeom.New_Size(110, 60))
    Dim r1 As MPGeom.PPRect: r1 = MPGeom.AARect_ToPPRect(r0)
    Dim r  As MPGeom.PPRect, ar As MPGeom.AARect
    Dim sc As Double
    
    SelectPrinter "Microsoft Print to PDF"
    With Printer
        '.ScaleMode = ScaleModeConstants.vbTwips 'default!
        Printer.Print "Printer.ScaleMode = Twips"
        sc = Millimeter_Scale(.ScaleMode, 1)
        r = MPGeom.PPRect_Mul(r1, sc)
        Printer.Line (r.P1.X, r.P1.Y)-(r.P2.X, r.P2.Y), , B
    End With
    
    Printer.NewPage
    
    With Printer
        .ScaleMode = ScaleModeConstants.vbPoints
        Printer.Print "Printer.ScaleMode = Points"
        sc = Millimeter_Scale(.ScaleMode, 1)
        r = MPGeom.PPRect_Mul(r1, sc)
        Printer.Line (r.P1.X, r.P1.Y)-(r.P2.X, r.P2.Y), , B
    End With
    
    Printer.NewPage
    
    With Printer
        .ScaleMode = ScaleModeConstants.vbPixels
        Printer.Print "Printer.ScaleMode = Pixels"
        sc = Millimeter_Scale(.ScaleMode, 1)
        r = MPGeom.PPRect_Mul(r1, sc)
        Printer.Line (r.P1.X, r.P1.Y)-(r.P2.X, r.P2.Y), , B
    End With
    
    Printer.NewPage
    
    With Printer
        .ScaleMode = ScaleModeConstants.vbInches
        Printer.Print "Printer.ScaleMode = Inches"
        sc = Millimeter_Scale(.ScaleMode, 1)
        r = MPGeom.PPRect_Mul(r1, sc)
        Printer.Line (r.P1.X, r.P1.Y)-(r.P2.X, r.P2.Y), , B
    End With
    
    Printer.NewPage
    
    With Printer
        .ScaleMode = ScaleModeConstants.vbMillimeters
        Printer.Print "Printer.ScaleMode = Millimeters"
        sc = Millimeter_Scale(.ScaleMode, 1)
        r = MPGeom.PPRect_Mul(r1, sc)
        Printer.Line (r.P1.X, r.P1.Y)-(r.P2.X, r.P2.Y), , B
    End With
    
    Printer.NewPage
    
    With Printer
        .ScaleMode = ScaleModeConstants.vbCentimeters
        Printer.Print "Printer.ScaleMode = Centimeters"
        sc = Millimeter_Scale(.ScaleMode, 1)
        r = MPGeom.PPRect_Mul(r1, sc)
        Printer.Line (r.P1.X, r.P1.Y)-(r.P2.X, r.P2.Y), , B
    End With
    
'    Printer.NewPage
'
'    With Printer
'        .ScaleMode = ScaleModeConstants.vbCharacters
'        sc = Millimeter_Scale(.ScaleMode, 1)
'        r = MPGeom.PPRect_Mul(r1, sc)
'        Printer.Line (r.P1.X, r.P1.Y)-(r.P2.X, r.P2.Y), , B
'    End With
    
    'now drawing positioning and scaling a Picture
    'What we want:
    'we want to draw a picture at a certain position with either a certain width and height scaled with ratio=1 or height and width scaled with ratio=1
    'in cm or millimeter
    
    'https://learn.microsoft.com/en-us/previous-versions/bb918079(v=vs.140)
    
    Dim pic As StdPicture: Set pic = Picture1.Picture
    
    Printer.NewPage
    
    With Printer
        .ScaleMode = ScaleModeConstants.vbTwips
        Printer.Print "Printer.ScaleMode = Twips"
        Printer.Print "Picture scaled to rectangle width"
        sc = Millimeter_Scale(.ScaleMode, 1)
        r = MPGeom.PPRect_Mul(r1, sc)
        Printer.Line (r.P1.X, r.P1.Y)-(r.P2.X, r.P2.Y), , B
        ar = MPGeom.AARect_Mul(r0, sc)
        MPrinter.PaintPictureW pic, ar.Pt.X, ar.Pt.Y, ar.Sz.Width
    End With
    
    Printer.NewPage
    
    With Printer
        .ScaleMode = ScaleModeConstants.vbTwips
        Printer.Print "Printer.ScaleMode = Twips"
        Printer.Print "Picture scaled to rectangle height"
        sc = Millimeter_Scale(.ScaleMode, 1)
        r = MPGeom.PPRect_Mul(r1, sc)
        Printer.Line (r.P1.X, r.P1.Y)-(r.P2.X, r.P2.Y), , B
        ar = MPGeom.AARect_Mul(r0, sc)
        MPrinter.PaintPictureH pic, ar.Pt.X, ar.Pt.Y, ar.Sz.Height
    End With
    
    
'    Dim r0 As MPGeom.AARect: r0 = MPGeom.New_AARect(MPGeom.New_Point(50, 70), MPGeom.New_Size(110, 60))
'    Dim r1 As MPGeom.PPRect: r1 = MPGeom.AARect_ToPPRect(r0)
'    Dim r  As MPGeom.PPRect, ar As MPGeom.AARect
    
    Printer.NewPage
    
    With Printer
        .ScaleMode = ScaleModeConstants.vbTwips
        Printer.Print "Printer.ScaleMode = Twips"
        Printer.Print "Picture scaled to fit into rectangle"
        sc = Millimeter_Scale(.ScaleMode, 1)
        r0.Sz.Height = 1.5 * r0.Sz.Height
        r1 = MPGeom.AARect_ToPPRect(r0)
        r = MPGeom.PPRect_Mul(r1, sc)
        Printer.Line (r.P1.X, r.P1.Y)-(r.P2.X, r.P2.Y), , B
        ar = MPGeom.AARect_Mul(r0, sc)
        MPrinter.PaintPictureFit pic, ar.Pt.X, ar.Pt.Y, ar.Sz.Height, ar.Sz.Width
    End With
    
    Printer.NewPage
    
    With Printer
        .ScaleMode = ScaleModeConstants.vbTwips
        Printer.Print "Printer.ScaleMode = Twips"
        Printer.Print "Picture scaled to fit into rectangle"
        sc = Millimeter_Scale(.ScaleMode, 1)
        r0.Sz.Width = 0.5 * r0.Sz.Width
        r1 = MPGeom.AARect_ToPPRect(r0)
        r = MPGeom.PPRect_Mul(r1, sc)
        Printer.Line (r.P1.X, r.P1.Y)-(r.P2.X, r.P2.Y), , B
        ar = MPGeom.AARect_Mul(r0, sc)
        MPrinter.PaintPictureFit pic, ar.Pt.X, ar.Pt.Y, ar.Sz.Height, ar.Sz.Width
    End With
    
    Printer.EndDoc
    Exit Sub
Catch:
    MsgBox "Error " & Err.Description
End Sub
