Option Explicit On 
Option Strict On
Option Compare Binary

Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

' <remarks>
'   Hauptformular der Anwendung.
' </remarks>
Public Class MainForm
    Inherits System.Windows.Forms.Form

    Private m_pd As New Printing.PrintDocument()
    Private m_intCurrentPage As Integer

#Region " Vom Windows Form Designer generierter Code "
    Public Sub New()
        MyBase.New()
        InitializeComponent()
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

    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboPrinters As System.Windows.Forms.ComboBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnPreview As System.Windows.Forms.Button
    Friend WithEvents btnChoosePrinter As System.Windows.Forms.Button
    Friend WithEvents btnPageSetup As System.Windows.Forms.Button

    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.cboPrinters = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.btnPreview = New System.Windows.Forms.Button()
        Me.btnChoosePrinter = New System.Windows.Forms.Button()
        Me.btnPageSetup = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cboPrinters
        '
        Me.cboPrinters.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPrinters.Location = New System.Drawing.Point(16, 32)
        Me.cboPrinters.Name = "cboPrinters"
        Me.cboPrinters.Size = New System.Drawing.Size(176, 21)
        Me.cboPrinters.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Pri&nter:"
        '
        'btnPrint
        '
        Me.btnPrint.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnPrint.Location = New System.Drawing.Point(216, 96)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 24)
        Me.btnPrint.TabIndex = 5
        Me.btnPrint.Text = "&Print"
        '
        'btnPreview
        '
        Me.btnPreview.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnPreview.Location = New System.Drawing.Point(16, 96)
        Me.btnPreview.Name = "btnPreview"
        Me.btnPreview.Size = New System.Drawing.Size(88, 24)
        Me.btnPreview.TabIndex = 4
        Me.btnPreview.Text = "Page P&review..."
        '
        'btnChoosePrinter
        '
        Me.btnChoosePrinter.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnChoosePrinter.Location = New System.Drawing.Point(200, 30)
        Me.btnChoosePrinter.Name = "btnChoosePrinter"
        Me.btnChoosePrinter.Size = New System.Drawing.Size(96, 24)
        Me.btnChoosePrinter.TabIndex = 2
        Me.btnChoosePrinter.Text = "Prin&ter Setup..."
        '
        'btnPageSetup
        '
        Me.btnPageSetup.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnPageSetup.Location = New System.Drawing.Point(16, 64)
        Me.btnPageSetup.Name = "btnPageSetup"
        Me.btnPageSetup.Size = New System.Drawing.Size(88, 24)
        Me.btnPageSetup.TabIndex = 3
        Me.btnPageSetup.Text = "P&age Setup..."
        '
        'MainForm
        '
        Me.AcceptButton = Me.btnPrint
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(314, 136)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnPageSetup, Me.btnChoosePrinter, Me.btnPreview, Me.btnPrint, Me.Label1, Me.cboPrinters})
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "MainForm"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "The .NET Print Framework"
        Me.ResumeLayout(False)

    End Sub
#End Region

    Private Sub MainForm_Load( _
      ByVal sender As System.Object, _
      ByVal e As System.EventArgs) _
      Handles MyBase.Load
        With cboPrinters
            Dim s As String
            For Each s In Printing.PrinterSettings.InstalledPrinters
                .Items.Add(s)
            Next s
            If .Items.Count > 0 Then
                .SelectedIndex = 0
            Else
                MessageBox.Show("No printers installed, quitting!", _
                  Application.ProductName, _
                  MessageBoxButtons.OK, _
                  MessageBoxIcon.Exclamation)
                Me.Close()
            End If
        End With
        m_pd.DocumentName = "Unser erstes Dokument"
        AddHandler m_pd.PrintPage, AddressOf m_pd_PrintPage
        m_intCurrentPage = 0
    End Sub

    Private Sub cboPrinters_SelectedIndexChanged( _
      ByVal sender As System.Object, _
      ByVal e As System.EventArgs) _
      Handles cboPrinters.SelectedIndexChanged
        If cboPrinters.SelectedIndex <> -1 Then
            m_pd.PrinterSettings.PrinterName = cboPrinters.Text
        End If
    End Sub

    Private Sub m_pd_PrintPage( _
      ByVal sender As System.Object, _
      ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        m_intCurrentPage += 1

        Select Case m_intCurrentPage

            ' Drucken der ersten Seite.
            Case 1

                ' Zeichnen eines elliptischen Bereichs 
                ' über die gesamte Seite (ohne Randabstände).
                e.Graphics.FillEllipse( _
                    New Drawing2D.HatchBrush( _
                        HatchStyle.Percent10, _
                        Color.Red, _
                        Color.White _
                    ), _
                    e.MarginBounds _
                )

                ' Text genau in der Mitte der Seite ausgeben. 
                ' Je nach Randabständen muss das
                ' nicht unbedingt genau in der Mitte der Ellipse sein!
                Dim strText As String = "Das ist die Seite 1"
                Dim fntFont As New Font("Arial", 18)
                e.Graphics.DrawString( _
                    strText, _
                    fntFont, _
                    New SolidBrush(Color.Blue), _
                    CSng( _
                        ( _
                            e.PageBounds.Width - _
                            e.Graphics.MeasureString(strText, fntFont).Width _
                        ) * 0.5 _
                    ), _
                    CSng(200) _
                )

                ' Es folgen noch weitere Seiten.
                e.HasMorePages = True

            ' Drucken der zweiten Seite.
            Case 2

                ' Seitennummer im linken oberen Eck ausgeben.
                e.Graphics.DrawString( _
                    "Seite 2", _
                    New Font("Times New Roman", 12), _
                    New SolidBrush(Color.Black), _
                    e.MarginBounds.Left, _
                    e.MarginBounds.Top _
                )

                ' Das war die letzte Seite.
                e.HasMorePages = False

                ' Seitenzähler wieder zurücksetzen.
                m_intCurrentPage = 0
        End Select
    End Sub

    Private Sub btnPrint_Click( _
      ByVal sender As System.Object, _
      ByVal e As System.EventArgs) _
      Handles btnPrint.Click
        m_pd.Print()
    End Sub

    Private Sub btnPreview_Click( _
      ByVal sender As System.Object, _
      ByVal e As System.EventArgs) _
      Handles btnPreview.Click

        ' Seitenzähler initialisieren.
        m_intCurrentPage = 0

        ' Vorschaudialog erstellen und anzeigen.
        Dim ppdlg As ExtendedPrintPreviewDialog = _
          New ExtendedPrintPreviewDialog()
        With ppdlg

            ' Der Druckvorschau das Dokument zuweisen.
            .Document = m_pd

            ' Die Druckvorschau soll maximiert gezeigt werden.
            .WindowState = FormWindowState.Maximized

            ' Druckvorschau anzeigen.
            .ShowDialog(Me)
        End With
    End Sub

    Private Sub btnChoosePrinter_Click( _
      ByVal sender As System.Object, _
      ByVal e As System.EventArgs) _
      Handles btnChoosePrinter.Click
        Dim pdlg As PrintDialog = New PrintDialog()
        With pdlg

            ' Dokument an Printerdialog weiterreichen.
            .Document = m_pd

            .PrinterSettings = m_pd.PrinterSettings
            .AllowPrintToFile = False
            If .ShowDialog(Me) = DialogResult.OK Then

                ' Die Einstellungen in unserem 
                ' Formular werden nun angepasst...
                SelectPrinter(cboPrinters, .PrinterSettings.PrinterName)
            End If
        End With
    End Sub

    Private Sub btnPageSetup_Click( _
      ByVal sender As System.Object, _
      ByVal e As System.EventArgs) _
      Handles btnPageSetup.Click
        Dim psdlg As PageSetupDialog = New PageSetupDialog()
        With psdlg
            .PrinterSettings = m_pd.PrinterSettings
            .PageSettings = m_pd.DefaultPageSettings
            If .ShowDialog(Me) = DialogResult.OK Then

                ' Hier wird ein Fehler (?!) ausgebügelt: VB .NET 
                ' konvertiert anscheinend alle Werte von Inch in 
                ' Millimeter, da (vermutlich) im englischen Dialog 
                ' die(Werte) in Inch eingegeben werden. Allerdings 
                ' ist der Umrechnungsfaktor nicht genau Inch:Millimeter, 
                ' sondern etwas mehr, sodass beim Wert 10 bei 
                ' erneutem Aufruf 9.9 in der TextBox steht.
                .PageSettings.Margins = _
                  Printing.PrinterUnitConvert.Convert( _
                  .PageSettings.Margins, _
                  Drawing.Printing.PrinterUnit.ThousandthsOfAnInch, _
                  Drawing.Printing.PrinterUnit.HundredthsOfAMillimeter)

                ' Die Einstellungen in unserem Formular 
                ' werden nun angepasst...
                SelectPrinter(cboPrinters, .PrinterSettings.PrinterName)
            End If
        End With
    End Sub

    ' <summary>
    '   Wählt den in <paramref name="strPrinterName"/> 
    '   angegebenen Drucker in den Einträgen
    '   der im Parameter <paramref name="cboComboBox"/> ComboBox aus.
    ' </summary>
    ' <param name="cboComboBox"></param>
    ' <param name="strPrinterName"></param>
    Private Sub SelectPrinter( _
      ByVal cboComboBox As ComboBox, _
      ByVal strPrinterName As String)
        Dim i As Integer
        For i = 0 To cboComboBox.Items.Count - 1
            If Convert.ToString(cboComboBox.Items(i)) = strPrinterName Then
                cboComboBox.SelectedIndex = i
                Exit For
            End If
        Next i
    End Sub
End Class