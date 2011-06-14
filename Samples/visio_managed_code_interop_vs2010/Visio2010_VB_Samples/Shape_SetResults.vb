Imports IVisio = Microsoft.Office.Interop.Visio


Partial Public Class VS2010_VB_Samples

    Shared Sub Shape_SetResults(ByVal doc As Microsoft.Office.Interop.Visio.Document)
        Dim page = VisioInterop.Util.CreateStandardPage(doc, "SSR")
        Dim shape = VisioInterop.Util.CreateStandardShape(page)

        ' CREATE REQUEST
        Dim request = {New With { _
                Key .Section = CShort(IVisio.VisSectionIndices.visSectionObject), _
                Key .Row = CShort(IVisio.VisRowIndices.visRowXFormOut), _
                Key .Cell = CShort(IVisio.VisCellIndices.visXFormWidth), _
                Key .UnitCode = CShort(IVisio.VisUnitCodes.visNoCast), _
                Key .Result = CDbl(8.2) _
                }, _
                       New With { _
                Key .Section = CShort(IVisio.VisSectionIndices.visSectionObject), _
                Key .Row = CShort(IVisio.VisRowIndices.visRowXFormOut), _
                Key .Cell = CShort(IVisio.VisCellIndices.visXFormHeight), _
                Key .UnitCode = CShort(IVisio.VisUnitCodes.visNoCast), _
                Key .Result = CDbl(1.3) _
                }}

        ' MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        Dim SRCStream = New Short(request.Length * 3 - 1) {}
        Dim results_objects = New Object(request.Length - 1) {}
        Dim unitcodes = New Object(request.Length - 1) {}
        For i As Integer = 0 To request.Length - 1
            SRCStream((i * 3) + 0) = request(i).Section
            SRCStream((i * 3) + 1) = request(i).Row
            SRCStream((i * 3) + 2) = request(i).Cell
            results_objects(i) = request(i).Result
            unitcodes(i) = request(i).UnitCode
        Next

        ' EXECUTE THE REQUEST
        Dim flags As Short = 0
        Dim count As Integer = shape.SetResults(SRCStream, unitcodes, results_objects, flags)

        ' DISPLAY THE INFORMATION
        shape.Text = "SetResults"

    End Sub

End Class
