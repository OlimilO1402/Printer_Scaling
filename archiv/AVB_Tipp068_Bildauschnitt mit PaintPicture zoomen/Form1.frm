VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "www.ActiveVB.de"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   1440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ende"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.PictureBox PBZoom 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   1440
      ScaleHeight     =   2535
      ScaleWidth      =   2175
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.PictureBox Source 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   1320
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1320
      ScaleWidth      =   1125
      TabIndex        =   0
      Top             =   120
      Width           =   1125
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !

Option Explicit

Private m_Wert    As Single
Private m_Schritt As Single

Private Sub Form_Load()
    m_Wert = 100
    m_Schritt = 1
End Sub

Private Sub Command1_Click()
    Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Zoom m_Wert
    
    If m_Wert >= 300 Then m_Schritt = -m_Schritt
    If m_Wert <= 10 Then m_Schritt = -m_Schritt
    
    If m_Wert < 50 Then m_Schritt = Sgn(m_Schritt) * 2
    If m_Wert >= 50 Then m_Schritt = Sgn(m_Schritt) * 5
    
    m_Wert = m_Wert + m_Schritt
End Sub

Private Sub Zoom(ByVal Prozent As Single)
    Dim w0 As Single: w0 = Source.Width
    Dim h0 As Single: h0 = Source.Height
    
    Dim w1 As Single: w1 = w0 * Prozent / 100
    Dim h1 As Single: h1 = h0 * Prozent / 100
    PBZoom.Cls
    PBZoom.PaintPicture Source, 0, 0, w1, h1
End Sub
