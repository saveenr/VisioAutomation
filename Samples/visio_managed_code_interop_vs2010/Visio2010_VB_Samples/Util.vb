Imports IVisio = Microsoft.Office.Interop.Visio

Public NotInheritable Class Util
    Private Sub New()
    End Sub
    Public Shared Function CreateStandardShape(page As IVisio.Page) As IVisio.Shape
        Dim shape = page.DrawRectangle(1, 1, 4, 3)
        Dim cell_width = shape.CellsU("Width")
        Dim cell_height = shape.CellsU("Height")
        cell_width.Formula = "=(1.0+2.5)"
        cell_height.Formula = "=(0.0+1.5)"
        Return shape
    End Function

    Public Shared Function CreateStandardPage(doc As IVisio.Document, pagename As String) As IVisio.Page
        Dim pages = doc.Pages
        Dim page = pages.Add()
        page.NameU = pagename
        Return page
    End Function
End Class
