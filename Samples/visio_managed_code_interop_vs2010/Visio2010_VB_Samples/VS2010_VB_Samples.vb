Imports IVisio = Microsoft.Office.Interop.Visio


Public Class VS2010_VB_Samples

    Shared Sub Shape_GetFormulas(ByVal doc As Microsoft.Office.Interop.Visio.Document)

        Dim page = VisioInterop.Util.CreateStandardPage(doc, "SGF")
        Dim shape = VisioInterop.Util.CreateStandardShape(page)
        Dim request = VisioInterop.Util.Create_SGF_Request()

        ' MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        Dim SRCStream = VS2010_VB_Samples.CreateShortArray(request.Length * 3)
        For i = 0 To request.Length - 1
            Dim item = request(i)
            SRCStream(i * 3 + 0) = item.CellSRC.SectionIndex
            SRCStream(i * 3 + 1) = item.CellSRC.RowIndex
            SRCStream(i * 3 + 2) = item.CellSRC.CellIndex
        Next i

        ' EXECUTE THE REQUEST
        Dim formulas_sa As System.Array = Nothing
        shape.GetFormulasU(SRCStream, formulas_sa)

        ' MAP OUTPUT BACK TO SOMETHING USEFUL 
        Dim formulas(request.Length) As String
        formulas_sa.CopyTo(formulas, 0)

        ' DISPLAY THE INFORMATION
        shape.Text = String.Format("Formulas={0},{1}", formulas(0), formulas(1))

    End Sub


    Shared Sub Shape_GetResults(ByVal doc As Microsoft.Office.Interop.Visio.Document)

        Dim page = VisioInterop.Util.CreateStandardPage(doc, "SGR")
        Dim shape = VisioInterop.Util.CreateStandardShape(page)
        Dim request = VisioInterop.Util.Create_SGR_Request()

        ' MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        Dim SRCStream = VS2010_VB_Samples.CreateShortArray(request.Length * 3)
        Dim unitcodes = VS2010_VB_Samples.CreateObjectArray(request.Length)
        For i = 0 To request.Length - 1
            Dim item = request(i)
            SRCStream(i * 3 + 0) = item.CellSRC.SectionIndex
            SRCStream(i * 3 + 1) = item.CellSRC.RowIndex
            SRCStream(i * 3 + 2) = item.CellSRC.CellIndex
            unitcodes(i) = item.UnitCode
        Next i

        ' EXECUTE THE REQUEST
        Dim flags = CShort(Microsoft.Office.Interop.Visio.VisGetSetArgs.visGetFloats)
        Dim results_sa As System.Array = Nothing
        shape.GetResults(SRCStream, flags, unitcodes, results_sa)

        ' MAP OUTPUT BACK TO SOMETHING USEFUL 
        Dim results(request.Length) As Double
        results_sa.CopyTo(results, 0)

        ' DISPLAY THE INFORMATION
        shape.Text = String.Format("Formulas={0},{1}", results(0), results(1))

    End Sub

    Shared Sub Shape_SetFormulas(ByVal doc As Microsoft.Office.Interop.Visio.Document)

        Dim page = VisioInterop.Util.CreateStandardPage(doc, "SSF")
        Dim shape = VisioInterop.Util.CreateStandardShape(page)
        Dim request = VisioInterop.Util.Create_SSF_Request()

        ' MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        Dim SRCStream = VS2010_VB_Samples.CreateShortArray(request.Length * 3)
        Dim formulas = VS2010_VB_Samples.CreateObjectArray(request.Length)
        For i = 0 To request.Length - 1
            Dim item = request(i)
            SRCStream(i * 3 + 0) = item.CellSRC.SectionIndex
            SRCStream(i * 3 + 1) = item.CellSRC.RowIndex
            SRCStream(i * 3 + 2) = item.CellSRC.CellIndex
            formulas(i) = item.Formula
        Next i

        ' EXECUTE THE REQUEST
        Dim flags = CShort(IVisio.VisGetSetArgs.visSetBlastGuards Or IVisio.VisGetSetArgs.visSetUniversalSyntax)
        Dim count = shape.SetFormulas(SRCStream, formulas, flags)

        ' DISPLAY THE INFORMATION
        shape.Text = String.Format("SetFormulas")

    End Sub

    Shared Sub Shape_SetResults(ByVal doc As Microsoft.Office.Interop.Visio.Document)

        Dim page = VisioInterop.Util.CreateStandardPage(doc, "SSR")
        Dim shape = VisioInterop.Util.CreateStandardShape(page)
        Dim request = VisioInterop.Util.Create_SSR_Request()

        ' MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        Dim SRCStream = VS2010_VB_Samples.CreateShortArray(request.Length * 3)
        Dim results = VS2010_VB_Samples.CreateObjectArray(request.Length)
        Dim unitcodes = VS2010_VB_Samples.CreateObjectArray(request.Length)
        For i = 0 To request.Length - 1
            Dim item = request(i)
            SRCStream(i * 3 + 0) = item.CellSRC.SectionIndex
            SRCStream(i * 3 + 1) = item.CellSRC.RowIndex
            SRCStream(i * 3 + 2) = item.CellSRC.CellIndex
            results(i) = item.Result
            unitcodes(i) = item.Result
        Next i

        ' EXECUTE THE REQUEST
        Dim flags = CShort(0)
        Dim count = shape.SetResults(SRCStream, unitcodes, results, flags)

        ' DISPLAY THE INFORMATION
        shape.Text = String.Format("SetResults")

    End Sub


    Shared Sub Page_GetFormulas(ByVal doc As Microsoft.Office.Interop.Visio.Document)

        Dim page = VisioInterop.Util.CreateStandardPage(doc, "PGF")
        Dim shape = VisioInterop.Util.CreateStandardShape(page)
        Dim request = VisioInterop.Util.Create_PGF_Request(shape)

        ' MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        Dim SRCStream = VS2010_VB_Samples.CreateShortArray(request.Length * 4)
        For i = 0 To request.Length - 1
            Dim item = request(i)
            SRCStream(i * 4 + 0) = item.ShapeID
            SRCStream(i * 4 + 1) = item.CellSRC.SectionIndex
            SRCStream(i * 4 + 2) = item.CellSRC.RowIndex
            SRCStream(i * 4 + 3) = item.CellSRC.CellIndex
        Next i

        ' EXECUTE THE REQUEST
        Dim formulas_sa As System.Array = Nothing
        page.GetFormulasU(SRCStream, formulas_sa)

        ' MAP OUTPUT BACK TO SOMETHING USEFUL 
        Dim formulas(request.Length) As String
        formulas_sa.CopyTo(formulas, 0)

        ' DISPLAY THE INFORMATION
        shape.Text = String.Format("Formulas={0},{1}", formulas(0), formulas(1))

    End Sub

    Shared Sub Page_GetResults(ByVal doc As Microsoft.Office.Interop.Visio.Document)

        Dim page = VisioInterop.Util.CreateStandardPage(doc, "PGR")
        Dim shape = VisioInterop.Util.CreateStandardShape(page)
        Dim request = VisioInterop.Util.Create_PGR_Request(shape)

        ' MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        Dim SRCStream = VS2010_VB_Samples.CreateShortArray(request.Length * 4)
        Dim unitcodes = VS2010_VB_Samples.CreateObjectArray(request.Length)
        For i = 0 To request.Length - 1
            Dim item = request(i)
            SRCStream(i * 4 + 0) = item.ShapeID
            SRCStream(i * 4 + 1) = item.CellSRC.SectionIndex
            SRCStream(i * 4 + 2) = item.CellSRC.RowIndex
            SRCStream(i * 4 + 3) = item.CellSRC.CellIndex
            unitcodes(i) = item.UnitCode
        Next i

        ' EXECUTE THE REQUEST
        Dim flags = CShort(Microsoft.Office.Interop.Visio.VisGetSetArgs.visGetFloats)
        Dim results_sa As System.Array = Nothing
        page.GetResults(SRCStream, flags, unitcodes, results_sa)

        ' MAP OUTPUT BACK TO SOMETHING USEFUL 
        Dim results(request.Length) As Double
        results_sa.CopyTo(results, 0)

        ' DISPLAY THE INFORMATION
        shape.Text = String.Format("Formulas={0},{1}", results(0), results(1))

    End Sub



    Shared Sub Page_SetFormulas(ByVal doc As Microsoft.Office.Interop.Visio.Document)

        Dim page = VisioInterop.Util.CreateStandardPage(doc, "PSF")
        Dim shape = VisioInterop.Util.CreateStandardShape(page)
        Dim request = VisioInterop.Util.Create_PSF_Request(shape)

        ' MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        Dim flags = CShort(IVisio.VisGetSetArgs.visSetBlastGuards Or IVisio.VisGetSetArgs.visSetUniversalSyntax)
        Dim SRCStream = VS2010_VB_Samples.CreateShortArray(request.Length * 4)
        Dim formulas = VS2010_VB_Samples.CreateObjectArray(request.Length)
        For i = 0 To request.Length - 1
            Dim item = request(i)
            SRCStream(i * 4 + 0) = item.ShapeID
            SRCStream(i * 4 + 1) = item.CellSRC.SectionIndex
            SRCStream(i * 4 + 2) = item.CellSRC.RowIndex
            SRCStream(i * 4 + 3) = item.CellSRC.CellIndex
            formulas(i) = item.Formula
        Next i

        ' EXECUTE THE REQUEST
        Dim count = page.SetFormulas(SRCStream, formulas, flags)


        ' DISPLAY THE INFORMATION
        shape.Text = String.Format("SetFormulas")

    End Sub

    Shared Sub Page_SetResults(ByVal doc As Microsoft.Office.Interop.Visio.Document)

        Dim page = VisioInterop.Util.CreateStandardPage(doc, "PSR")
        Dim shape = VisioInterop.Util.CreateStandardShape(page)
        Dim request = VisioInterop.Util.Create_PSR_Request(shape)

        ' MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        Dim SRCStream = VS2010_VB_Samples.CreateShortArray(request.Length * 4)
        Dim results = VS2010_VB_Samples.CreateObjectArray(request.Length)
        Dim unitcodes = VS2010_VB_Samples.CreateObjectArray(request.Length)
        For i = 0 To request.Length - 1
            Dim item = request(i)
            SRCStream(i * 4 + 0) = item.ShapeID
            SRCStream(i * 4 + 1) = item.CellSRC.SectionIndex
            SRCStream(i * 4 + 2) = item.CellSRC.RowIndex
            SRCStream(i * 4 + 3) = item.CellSRC.CellIndex
            results(i) = item.Result
            unitcodes(i) = item.Result
        Next i

        ' EXECUTE THE REQUEST
        Dim flags = CShort(0)
        Dim count = page.SetResults(SRCStream, unitcodes, results, flags)


        ' DISPLAY THE INFORMATION
        shape.Text = String.Format("SetResults")

    End Sub

    Shared Function CreateShortArray(ByVal length As Integer) As System.Array
        Dim s = CType(Array.CreateInstance(GetType(Short), length), Short())
        Return s
    End Function

    Shared Function CreateObjectArray(ByVal length As Integer) As System.Array
        Dim s = CType(Array.CreateInstance(GetType(Object), length), Object())
        Return s
    End Function

    Shared Function CreateStringArray(ByVal length As Integer) As System.Array
        Dim a = CType(Array.CreateInstance(GetType(String), length), String())
        Return a
    End Function

End Class
