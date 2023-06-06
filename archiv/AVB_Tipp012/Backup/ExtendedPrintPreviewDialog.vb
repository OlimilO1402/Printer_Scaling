Option Explicit On 
Option Strict On
Option Compare Binary

Imports System.Drawing
Imports System.Windows.Forms

' <remarks>
'   Erweitert den Druckvorschau-Dialog 
'   um einige optische Effekte.
' </remarks>
Public Class ExtendedPrintPreviewDialog
    Inherits System.Windows.Forms.PrintPreviewDialog

#Region " Vom Windows Form Designer generierter Code "
    Public Sub New()
        MyBase.New()
        InitializeComponent()

        ' Einige Anpassungen vornehmen. Vor Allem der Schliessen-Button ist etwas mager
        ' ausgefallen, daher wird er im Systemstyle dargestellt und an den unteren
        ' Rand des Dialogs gerückt.
        With Me
            Dim b As Button = DirectCast(.Controls(1).Controls(2), Button)
            b.Location = New Point(0, 0)
            b.FlatStyle = FlatStyle.System
            Me.MinimumSize = New Size(Me.MinimumSize.Width - b.Width, Me.MinimumSize.Height)
            Dim p As Panel = New Panel()
            b.Size = New Size(80, 24)
            p.Size = b.Size
            p.Controls.Add(b)
            b.Anchor = AnchorStyles.None
            p.Height = 40
            p.Dock = DockStyle.Bottom
            .Controls.Add(p)
            With DirectCast(.Controls(1), ToolBar)
                .Buttons.RemoveAt(8)
                .Divider = False
            End With
        End With
    End Sub

    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
Me.SuspendLayout()
'
'ExtendedPrintPreviewDialog
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(400, 300)
Me.Name = "ExtendedPrintPreviewDialog"
Me.ResumeLayout(False)

    End Sub
#End Region
End Class