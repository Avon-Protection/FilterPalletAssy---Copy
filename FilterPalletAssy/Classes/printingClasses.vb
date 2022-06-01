Imports Microsoft.VisualBasic, System.Drawing, System.Web, System
Imports System.Net
Imports System.net.mail
Imports System.IO
Imports Neodynamic.WebControls.BarcodeProfessional

Namespace commonPrinting

    Public Class printingClass
        Private drawString As String
        Private drawFont As Font
        Private drawBrush As SolidBrush
        Private drawPoints As PointF
        Private drawFormat As New StringFormat

        Public Sub renderTextData(ByVal xCoord As Single, ByVal yCoord As Single, ByVal fontType As String, ByVal fontSize As Integer, ByVal drawString As String, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
            drawFont = New Font(fontType, fontSize)
            drawBrush = New SolidBrush(Color.Black)
            drawPoints = New PointF(xCoord, yCoord)

            If drawString <> "" Then
                e.Graphics.DrawString(drawString, drawFont, drawBrush, drawPoints, drawFormat)
            End If
        End Sub
		'Public Sub renderBarcodeData(ByVal xCoord As Single, ByVal yCoord As Single, ByVal barHeight As Single, ByVal barText As String, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
		'    Dim bcp As New Neodynamic.WebControls.BarcodeProfessional.BarcodeProfessional()
		'    'BarcodeProfessional.LicenseKey = "UBCPMXWRRNWSTFWUZ5BJX6YJTMHZFPLMHYSQXEBHZ8UGFFRG52GQ"
		'    'BarcodeProfessional.LicenseOwner = "Avon Rubber p.l.c.-Standard Edition-Developer License"
		'    BarcodeProfessional.LicenseKey = "BM2F6PTFL9H7LL4UVL8J2FSMRQPTRH47XKG4QLYXLB45CZUT96PQ"
		'    BarcodeProfessional.LicenseOwner = "Avon Rubber p.l.c.-Ultimate Edition-Developer License"
		'    bcp.Symbology = Neodynamic.WebControls.BarcodeProfessional.Symbology.Code39

		'    bcp.BarHeight = barHeight
		'    bcp.QuietZoneWidth = 0

		'    ' BARCODE TEXT
		'    bcp.Code = barText
		'    bcp.DisplayStartStopChar = False
		'    bcp.BarWidth = 0.02F
		'    bcp.AddChecksum = False

		'    If barText <> "" Then
		'        bcp.DrawOnCanvas(e.Graphics, New System.Drawing.PointF(xCoord / 100, yCoord / 100))
		'    End If
		'End Sub
		'Public Sub renderBarcodeDataLabel(ByVal xCoord As Single, ByVal yCoord As Single, ByVal barHeight As Single, ByVal barText As String, ByVal barWidth As Single, ByVal barRatio As Single, ByVal fontSize As Integer, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
		'    Dim bcp As New Neodynamic.WebControls.BarcodeProfessional.BarcodeProfessional()
		'    'BarcodeProfessional.LicenseKey = "UBCPMXWRRNWSTFWUZ5BJX6YJTMHZFPLMHYSQXEBHZ8UGFFRG52GQ"
		'    'BarcodeProfessional.LicenseOwner = "Avon Rubber p.l.c.-Standard Edition-Developer License"
		'    BarcodeProfessional.LicenseKey = "BM2F6PTFL9H7LL4UVL8J2FSMRQPTRH47XKG4QLYXLB45CZUT96PQ"
		'    BarcodeProfessional.LicenseOwner = "Avon Rubber p.l.c.-Ultimate Edition-Developer License"

		'    ' Use Code 128 symbology
		'    bcp.Symbology = Neodynamic.WebControls.BarcodeProfessional.Symbology.Code39

		'    bcp.BarHeight = barHeight
		'    bcp.QuietZoneWidth = 0
		'    bcp.BarRatio = barRatio
		'    bcp.Font.Bold = True
		'    bcp.Font.Size = fontSize

		'    '' BARCODE TEXT
		'    bcp.Code = barText
		'    bcp.DisplayStartStopChar = False
		'    bcp.BarWidth = barWidth
		'    bcp.AddChecksum = False

		'    If barText <> "" Then
		'        bcp.DrawOnCanvas(e.Graphics, New System.Drawing.PointF(xCoord / 100, yCoord / 100))
		'    End If
		'End Sub
		'Public Sub render2DBarcodeData(ByVal xCoord As Single, ByVal yCoord As Single, ByVal barWidth As Single, ByVal barText As String, ByVal barRatio As Integer, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
		'    Dim bcp As New Neodynamic.WebControls.BarcodeProfessional.BarcodeProfessional()
		'    'BarcodeProfessional.LicenseKey = "UBCPMXWRRNWSTFWUZ5BJX6YJTMHZFPLMHYSQXEBHZ8UGFFRG52GQ"
		'    BarcodeProfessional.LicenseKey = "BM2F6PTFL9H7LL4UVL8J2FSMRQPTRH47XKG4QLYXLB45CZUT96PQ"
		'    BarcodeProfessional.LicenseOwner = "Avon Rubber p.l.c.-Ultimate Edition-Developer License"
		'    'BarcodeProfessional.LicenseOwner = "Avon Rubber p.l.c.-Standard Edition-Developer License"
		'    bcp.Symbology = Neodynamic.WebControls.BarcodeProfessional.Symbology.Pdf417
		'    bcp.Pdf417CompactionType = Pdf417CompactionType.Text
		'    bcp.Pdf417ErrorCorrectionLevel = Neodynamic.WebControls.BarcodeProfessional.Pdf417ErrorCorrection.Level2

		'    bcp.BarWidth = barWidth
		'    bcp.BarRatio = barRatio
		'    bcp.BarHeight = barRatio * barWidth
		'    bcp.Pdf417Rows = 0
		'    bcp.Pdf417Columns = 0
		'    bcp.QuietZoneWidth = 0
		'    'bcp.QuietZoneWidth = bcp.BarWidth * 2
		'    bcp.TopMargin = bcp.BarWidth * 2
		'    bcp.BottomMargin = bcp.BarWidth * 2
		'    bcp.Code = barText
		'    xCoord = xCoord / 100
		'    yCoord = yCoord / 100


		'    bcp.DrawOnCanvas(e.Graphics, New System.Drawing.PointF(xCoord, yCoord))
		'End Sub
		Public Sub renderImage(ByVal xCoord As Int32, ByVal yCoord As Int32, ByVal width As Int32, ByVal height As Int32, ByVal imageFile As String, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
            Dim printImage As System.Drawing.Image = System.Drawing.Image.FromFile(imageFile)
            
            e.Graphics.DrawImage(printImage, xCoord, yCoord, width, height)


        End Sub
    End Class
End Namespace

