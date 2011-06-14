Imports IVisio = Microsoft.Office.Interop.Visio


Partial Public Class VS2010_VB_Samples
    Shared Sub Shape_GetResults (ByVal doc As Microsoft.Office.Interop.Visio.Document)
        Dim page = VisioInterop.Util.CreateStandardPage (doc, "SGR")
        Dim shape = VisioInterop.Util.CreateStandardShape (page)

        ' CREATE REQUEST
        Dim request = {New With { _
                Key .Section = CShort (IVisio.VisSectionIndices.visSectionObject), _
                Key .Row = CShort (IVisio.VisRowIndices.visRowXFormOut), _
                Key .Cell = CShort (IVisio.VisCellIndices.visXFormWidth), _
                Key .UnitCode = CShort (IVisio.VisUnitCodes.visNoCast) _
                }, New With { _
                Key .Section = CShort (IVisio.VisSectionIndices.visSectionObject), _
                Key .Row = CShort (IVisio.VisRowIndices.visRowXFormOut), _
                Key .Cell = CShort (IVisio.VisCellIndices.visXFormHeight), _
                Key .UnitCode = CShort (IVisio.VisUnitCodes.visNoCast) _
                }}

        ' MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        Dim SRCStream = New Short(request.Length*3 - 1) {}
        Dim unitcodes = New Object(request.Length - 1) {}
        For i As Integer = 0 To request.Length - 1
            SRCStream ((i*3) + 0) = request (i).Section
            SRCStream ((i*3) + 1) = request (i).Row
            SRCStream ((i*3) + 2) = request (i).Cell
            unitcodes (i) = request (i).UnitCode
        Next

        ' EXECUTE THE REQUEST
        Dim flags As Short = CShort (IVisio.VisGetSetArgs.visGetFloats)
        Dim results_sa As System.Array = Nothing
        shape.GetResults (SRCStream, flags, unitcodes, results_sa)

        ' MAP OUTPUT BACK TO SOMETHING USEFUL 
        Dim results_doubles = New Double(results_sa.Length - 1) {}
        results_sa.CopyTo (results_doubles, 0)

        ' DISPLAY THE INFORMATION
        shape.Text = String.Format ("Results={0},{1}", results_doubles (0), results_doubles (1))

    End Sub
End Class
