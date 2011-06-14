Imports IVisio = Microsoft.Office.Interop.Visio


Partial Public Class VS2010_VB_Samples
    Shared Sub Shape_SetFormulas (ByVal doc As Microsoft.Office.Interop.Visio.Document)
        Dim page = Util.CreateStandardPage(doc, "SSF")
        Dim shape = Util.CreateStandardShape(page)

        ' CREATE REQUEST
        Dim request = {New With { _
                Key .Section = CShort (IVisio.VisSectionIndices.visSectionObject), _
                Key .Row = CShort (IVisio.VisRowIndices.visRowXFormOut), _
                Key .Cell = CShort (IVisio.VisCellIndices.visXFormWidth), _
                Key .Formula = "2.0" _
                }, New With { _
                Key .Section = CShort (IVisio.VisSectionIndices.visSectionObject), _
                Key .Row = CShort (IVisio.VisRowIndices.visRowXFormOut), _
                Key .Cell = CShort (IVisio.VisCellIndices.visXFormHeight), _
                Key .Formula = "3.0" _
                }}

        ' MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS

        Dim SRCStream = New Short(request.Length*3 - 1) {}
        Dim formulas_objects = New Object(request.Length - 1) {}
        For i As Integer = 0 To request.Length - 1
            SRCStream ((i*3) + 0) = request (i).Section
            SRCStream ((i*3) + 1) = request (i).Row
            SRCStream ((i*3) + 2) = request (i).Cell
            formulas_objects (i) = request (i).Formula
        Next

        ' EXECUTE THE REQUEST
        Dim _
            flags As Short = _
                CShort (IVisio.VisGetSetArgs.visSetBlastGuards Or IVisio.VisGetSetArgs.visSetUniversalSyntax)
        Dim count As Integer = shape.SetFormulas (SRCStream, formulas_objects, flags)

        ' DISPLAY THE INFORMATION
        shape.Text = "SetFormulas"
    End Sub
End Class
