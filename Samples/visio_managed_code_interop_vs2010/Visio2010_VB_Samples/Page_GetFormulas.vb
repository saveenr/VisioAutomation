Imports IVisio = Microsoft.Office.Interop.Visio


Partial Public Class VS2010_VB_Samples
    Shared Sub Page_GetFormulas(ByVal doc As Microsoft.Office.Interop.Visio.Document)
        Dim page = VisioInterop.Util.CreateStandardPage(doc, "PGF")
        Dim shape = VisioInterop.Util.CreateStandardShape(page)

        ' CREATE REQUEST
        Dim request = {New With { _
                Key .ID = CShort(shape.ID16), _
                Key .Section = CShort(IVisio.VisSectionIndices.visSectionObject), _
                Key .Row = CShort(IVisio.VisRowIndices.visRowXFormOut), _
                Key .Cell = CShort(IVisio.VisCellIndices.visXFormWidth) _
                }, New With { _
                Key .ID = CShort(shape.ID16), _
                Key .Section = CShort(IVisio.VisSectionIndices.visSectionObject), _
                Key .Row = CShort(IVisio.VisRowIndices.visRowXFormOut), _
                Key .Cell = CShort(IVisio.VisCellIndices.visXFormHeight) _
                }}

        ' MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        Dim SID_SRCStream = New Short(request.Length * 4 - 1) {}
        For i As Integer = 0 To request.Length - 1
            SID_SRCStream((i * 4) + 0) = request(i).ID
            SID_SRCStream((i * 4) + 1) = request(i).Section
            SID_SRCStream((i * 4) + 2) = request(i).Row
            SID_SRCStream((i * 4) + 3) = request(i).Cell
        Next

        ' EXECUTE THE REQUEST
        Dim formulas_sa As System.Array
        page.GetFormulasU(SID_SRCStream, formulas_sa)

        ' MAP OUTPUT BACK TO SOMETHING USEFUL 
        Dim formulas = New String(formulas_sa.Length - 1) {}
        formulas_sa.CopyTo(formulas, 0)

        ' DISPLAY THE INFORMATION
        shape.Text = String.Format("Formulas={0},{1}", formulas(0), formulas(1))

    End Sub
End Class
