Module Module1


    Sub Main()
        Dim visapp = New Microsoft.Office.Interop.Visio.Application()
        Dim docs = visapp.Documents
        Dim doc = docs.Add("")
        Visio2010_VB_Samples.VS2010_VB_Samples.Shape_GetFormulas(doc)
        Visio2010_VB_Samples.VS2010_VB_Samples.Shape_GetResults(doc)
        Visio2010_VB_Samples.VS2010_VB_Samples.Shape_SetFormulas(doc)
        Visio2010_VB_Samples.VS2010_VB_Samples.Shape_SetResults(doc)

        Visio2010_VB_Samples.VS2010_VB_Samples.Page_GetFormulas(doc)
        Visio2010_VB_Samples.VS2010_VB_Samples.Page_GetResults(doc)
        Visio2010_VB_Samples.VS2010_VB_Samples.Page_SetFormulas(doc)
        Visio2010_VB_Samples.VS2010_VB_Samples.Page_SetResults(doc)

    End Sub

End Module
