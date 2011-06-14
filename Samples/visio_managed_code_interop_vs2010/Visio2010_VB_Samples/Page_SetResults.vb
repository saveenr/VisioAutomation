Imports IVisio = Microsoft.Office.Interop.Visio


Partial Public Class VS2010_VB_Samples

    Shared Sub Page_SetResults (ByVal doc As Microsoft.Office.Interop.Visio.Document)

        Dim page = VisioInterop.Util.CreateStandardPage(doc, "PSR")
        Dim shape = VisioInterop.Util.CreateStandardShape(page)

        ' CREATE REQUEST
        Dim request = {New With { _
 Key .ID = CShort(shape.ID16), _
 Key .Section = CShort(IVisio.VisSectionIndices.visSectionObject), _
 Key .Row = CShort(IVisio.VisRowIndices.visRowXFormOut), _
 Key .Cell = CShort(IVisio.VisCellIndices.visXFormWidth), _
 Key .UnitCode = CShort(IVisio.VisUnitCodes.visNoCast), _
 Key .Result = CDbl(8.0) _
}, New With { _
 Key .ID = CShort(shape.ID16), _
 Key .Section = CShort(IVisio.VisSectionIndices.visSectionObject), _
 Key .Row = CShort(IVisio.VisRowIndices.visRowXFormOut), _
 Key .Cell = CShort(IVisio.VisCellIndices.visXFormHeight), _
 Key .UnitCode = CShort(IVisio.VisUnitCodes.visNoCast), _
 Key .Result = CDbl(1.3) _
}}

        ' MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        Dim SID_SRCStream = New Short(request.Length * 4 - 1) {}
        Dim results_objects = New Object(request.Length - 1) {}
        Dim unitcodes = New Object(request.Length - 1) {}
        For i As Integer = 0 To request.Length - 1
            SID_SRCStream((i * 4) + 0) = request(i).ID
            SID_SRCStream((i * 4) + 1) = request(i).Section
            SID_SRCStream((i * 4) + 2) = request(i).Row
            SID_SRCStream((i * 4) + 3) = request(i).Cell
            results_objects(i) = request(i).Result
            unitcodes(i) = request(i).UnitCode
        Next

        ' EXECUTE THE REQUEST
        Dim flags As Short = 0
        Dim count As Integer = page.SetResults(SID_SRCStream, unitcodes, results_objects, flags)

        ' DISPLAY THE INFORMATION
        shape.Text = "SetResults"

    End Sub

End Class
