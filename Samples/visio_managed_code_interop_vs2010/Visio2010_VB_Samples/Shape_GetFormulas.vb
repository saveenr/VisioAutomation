Imports IVisio = Microsoft.Office.Interop.Visio


Partial Public Class VS2010_VB_Samples
    Shared Sub Shape_GetFormulas (ByVal doc As Microsoft.Office.Interop.Visio.Document)

        Dim page = Util.CreateStandardPage(doc, "SGF")
        Dim shape = Util.CreateStandardShape(page)

        ' CREATE REQUEST
        Dim request = {New With { _
                Key .Section = CShort (IVisio.VisSectionIndices.visSectionObject), _
                Key .Row = CShort (IVisio.VisRowIndices.visRowXFormOut), _
                Key .Cell = CShort (IVisio.VisCellIndices.visXFormWidth) _
                }, New With { _
                Key .Section = CShort (IVisio.VisSectionIndices.visSectionObject), _
                Key .Row = CShort (IVisio.VisRowIndices.visRowXFormOut), _
                Key .Cell = CShort (IVisio.VisCellIndices.visXFormHeight) _
                }}

        ' MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        Dim SRCStream = New Short(request.Length*3 - 1) {}
        For i As Integer = 0 To request.Length - 1
            SRCStream ((i*3) + 0) = request (i).Section
            SRCStream ((i*3) + 1) = request (i).Row
            SRCStream ((i*3) + 2) = request (i).Cell
        Next

        ' EXECUTE THE REQUEST
        Dim formulas_sa As System.Array = Nothing
        shape.GetFormulasU (SRCStream, formulas_sa)

        ' MAP OUTPUT BACK TO SOMETHING USEFUL 
        Dim formulas = New String(formulas_sa.Length - 1) {}
        formulas_sa.CopyTo (formulas, 0)

        ' DISPLAY THE INFORMATION
        shape.Text = String.Format ("Formulas={0},{1}", formulas (0), formulas (1))
    End Sub
End Class

