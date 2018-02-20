Imports iTextSharp.text.pdf
Imports iTextSharp.text
Imports System.IO
Public Class OsmLOG

    Public Sub CreatePDF()

        Dim pDoc As New Document(PageSize.A4, 72, 72, 72, 72)
        Dim pstrOutFilePath As String = "C:\Users\Chris.Nicolatos\Desktop\WIP\pdf\test.pdf"
        Dim gif As Image = Image.GetInstance("C:\Users\Chris.Nicolatos\Desktop\WIP\pdf\download.png")
        Dim x As Drawing.Image

        'Dim LabelFont As New Font(Font.COURIER, 9.0F, Font.NORMAL, New Color(163, 21, 21))
        Dim pArial11 As Font = FontFactory.GetFont("arial", 11, FontStyle.Regular)
        Dim pArial12 As Font = FontFactory.GetFont("arial", 12, FontStyle.Regular)
        Dim pArial16b As Font = FontFactory.GetFont("arial", 16, FontStyle.Bold)
        PdfWriter.GetInstance(pDoc, New FileStream(pstrOutFilePath, FileMode.Create))
        pDoc.Open()
        gif.ScalePercent(40)
        pDoc.Add(gif)

        pDoc.Add(AddParagraph("LETTER OF GUARANTEE", pArial16b, 28, 28, "Center"))
        pDoc.Add(AddParagraph("07 Feb 2018", pArial12, 0, 14, "Right"))

        pDoc.Add(AddParagraph("To Whom It May Concern", pArial11, 0, 14, "Left"))
        pDoc.Add(AddParagraph("The holder of this letter is an employee of OSM CREW MANAGEMENT traveling to or from a Vessel/Rig.", pArial11, 0, 14, "Left"))

        pDoc.Add(MakeTable(pArial11))
        pDoc.Add(AddParagraph("OSM CREW MANAGEMENT will be responsible for all his expenses and guarantees his repatriation as and when required.", pArial11, 14, 14, "Left"))
        pDoc.Add(AddParagraph("Please allow him clearance for further transport to his destination and extend all travel courtesies to him.", pArial11, 0, 28, "Left"))
        pDoc.Add(AddParagraph("MACIEJ PRYKASZCZYK", pArial11, 0, 0, "Left"))
        pDoc.Add(AddParagraph("Recruitment Officer", pArial11, 0, 0, "Left"))
        pDoc.Add(AddParagraph("OSM Maritime Group", pArial11, 0, 0, "Left"))
        pDoc.Add(AddParagraph("P: +48 58 661 59 61", pArial11, 0, 0, "Left"))
        pDoc.Add(AddParagraph("D: +48 58 660 90 72", pArial11, 0, 0, "Left"))
        pDoc.Add(AddParagraph("M:+48 603 132 311", pArial11, 0, 0, "Left"))
        pDoc.Add(AddParagraph("A:      AL.Grunwaldzka 472D, 80 - 309 Gdansk, Poland", pArial11, 0, 0, "Left"))

        pDoc.Close()

    End Sub

    Private Function AddParagraph(ByVal pText As String, ByVal pFont As Font, ByVal pSpacingBefore As Integer, ByVal pSpacingAfter As Integer, ByVal pAlignment As String) As Paragraph

        Dim x2 As New Paragraph(pText, pFont) With {
            .SpacingBefore = pSpacingBefore,
            .SpacingAfter = pSpacingAfter
        }
        x2.SetAlignment(pAlignment)

        AddParagraph = x2

    End Function

    Private Function MakeTable(ByVal pFont As Font) As PdfPTable


        Dim Table As New PdfPTable(2) With {
            .LockedWidth = False,
            .HorizontalAlignment = 0,
            .SpacingBefore = 20,
            .SpacingAfter = 30
        }
        'relative col widths in proportions - 1/3 And 2/3
        Dim widths() As Single = {1, 2}
        Table.SetWidths(widths)


        Table.AddCell(AddCell("Name:", pFont))
        Table.AddCell(AddCell("HRACIUK SLAWOMIR ALEKSANDER", pFont))

        Table.AddCell(AddCell("Job Title:", pFont))
        Table.AddCell(AddCell("2nd Engineer", pFont))

        Table.AddCell(AddCell("Vessel:", pFont))
        Table.AddCell(AddCell("Troms Mira", pFont))

        Table.AddCell(AddCell("Date of Travel:", pFont))
        Table.AddCell(AddCell("09 February 2018", pFont))

        Table.AddCell(AddCell("Route:", pFont))
        Table.AddCell(AddCell("WAW LON EDI", pFont))

        Table.AddCell(AddCell("Destination", pFont))
        Table.AddCell(AddCell("EDI –EDINBURGH- UNITED KINGDOM", pFont))

        MakeTable = Table

    End Function
    Private Function AddCell(ByVal pText As String, ByVal pFont As Font) As PdfPCell
        Dim c1 As New PdfPCell(New Phrase(pText, pFont)) With {
                    .Border = Rectangle.NO_BORDER
                }
        AddCell = c1
    End Function
End Class
