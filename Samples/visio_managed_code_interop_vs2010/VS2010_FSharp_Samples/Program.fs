
type FSharp_Samples =
  
    static member Shape_GetFormulas( doc: Microsoft.Office.Interop.Visio.Document ) =

        let page = VisioInterop.Util.CreateStandardPage(doc,"SGF");
        let shape= VisioInterop.Util.CreateStandardShape(page);
        let request= VisioInterop.Util.Create_SGF_Request();

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        let SRCStream  = FSharp_Samples.CreateShortArray(request.Length*3)
        for i in 0 .. request.Length-1 do
            let item = request.[i]
            SRCStream.[i*3 + 0] <- item.CellSRC.SectionIndex
            SRCStream.[i*3 + 1] <- item.CellSRC.RowIndex
            SRCStream.[i*3 + 2] <- item.CellSRC.CellIndex

        //let formulas_sa : System.Array = null
        let formulas_sa_ref = ref null
        let SRCStream_sa : System.Array = SRCStream :> System.Array
        shape.GetFormulasU(ref SRCStream_sa, formulas_sa_ref);

        // MAP OUTPUT BACK TO SOMETHING USEFUL 
        let empty = System.String.Empty
        let formulas = Array.create (request.Length) (empty)
        formulas_sa_ref.Value.CopyTo(formulas, 0);

        // DISPLAY THE INFORMATION
        shape.Text <- System.String.Format("Formulas={0},{1}", formulas.[0], formulas.[1]);

        ()


    static member Shape_GetResults( doc: Microsoft.Office.Interop.Visio.Document ) =

        let page = VisioInterop.Util.CreateStandardPage(doc,"SGR");
        let shape= VisioInterop.Util.CreateStandardShape(page);
        let request= VisioInterop.Util.Create_SGR_Request();

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        let SRCStream  = FSharp_Samples.CreateShortArray( request.Length*3)
        let unitcodes = FSharp_Samples.CreateObjectArray(request.Length)
        for i in 0 .. request.Length-1 do
            let item = request.[i]
            SRCStream.[i*3 + 0] <- item.CellSRC.SectionIndex
            SRCStream.[i*3 + 1] <- item.CellSRC.RowIndex
            SRCStream.[i*3 + 2] <- item.CellSRC.CellIndex
            unitcodes.[i] <- box ( (int16)Microsoft.Office.Interop.Visio.VisUnitCodes.visNoCast )

        //let formulas_sa : System.Array = null
        let flags = (int16)Microsoft.Office.Interop.Visio.VisGetSetArgs.visGetFloats;
        let results_sa_ref = ref null
        let SRCStream_sa : System.Array = SRCStream :> System.Array
        let unitcodes_sa : System.Array = unitcodes :> System.Array
        shape.GetResults(ref SRCStream_sa, flags, ref unitcodes_sa, results_sa_ref);

        // MAP OUTPUT BACK TO SOMETHING USEFUL 
        let empty = System.String.Empty
        let results = Array.create (request.Length) (0.0)
        results_sa_ref.Value.CopyTo(results, 0);

        // DISPLAY THE INFORMATION
        shape.Text <- System.String.Format("Formulas={0},{1}", results.[0], results.[1]);

        ()


    static member Shape_SetFormulas( doc: Microsoft.Office.Interop.Visio.Document ) =

        let page = VisioInterop.Util.CreateStandardPage(doc,"SSF");
        let shape= VisioInterop.Util.CreateStandardShape(page);
        let request= VisioInterop.Util.Create_SSF_Request();

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        let SRCStream  = FSharp_Samples.CreateShortArray( request.Length*3)
        let formulas = FSharp_Samples.CreateObjectArray(request.Length)
        for i in 0 .. request.Length-1 do
            let item = request.[i]
            SRCStream.[i*3 + 0] <- item.CellSRC.SectionIndex
            SRCStream.[i*3 + 1] <- item.CellSRC.RowIndex
            SRCStream.[i*3 + 2] <- item.CellSRC.CellIndex
            formulas.[i] <- item.Formula :> System.Object

        //let formulas_sa : System.Array = null
        let flags = (int16)(Microsoft.Office.Interop.Visio.VisGetSetArgs.visSetBlastGuards ||| Microsoft.Office.Interop.Visio.VisGetSetArgs.visSetUniversalSyntax);
        let formulas_sa : System.Array = formulas :> System.Array
        let SRCStream_sa : System.Array = SRCStream :> System.Array
        let count = shape.SetFormulas(ref SRCStream_sa, ref formulas_sa, flags)

        // DISPLAY THE INFORMATION
        shape.Text <- System.String.Format("SetFormulas")

        ()


    static member Shape_SetResults( doc: Microsoft.Office.Interop.Visio.Document ) =

        let page = VisioInterop.Util.CreateStandardPage(doc,"SSR");
        let shape= VisioInterop.Util.CreateStandardShape(page);
        let request= VisioInterop.Util.Create_SSR_Request();

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        let SRCStream  = FSharp_Samples.CreateShortArray(request.Length*3)
        let results = FSharp_Samples.CreateObjectArray(request.Length)
        let unitcodes = FSharp_Samples.CreateObjectArray(request.Length)
        for i in 0 .. request.Length-1 do
            let item = request.[i]
            SRCStream.[i*3 + 0] <- item.CellSRC.SectionIndex
            SRCStream.[i*3 + 1] <- item.CellSRC.RowIndex
            SRCStream.[i*3 + 2] <- item.CellSRC.CellIndex
            results.[i] <- box ( item.Result )
            unitcodes.[i] <- box ( (int16)item.UnitCode)

        //let formulas_sa : System.Array = null
        let flags = (int16)(0);
        let results_sa : System.Array = results :> System.Array
        let unitcodes_sa : System.Array = unitcodes :> System.Array
        let SRCStream_sa : System.Array = SRCStream :> System.Array
        let count = shape.SetResults(ref SRCStream_sa, ref unitcodes_sa, ref results_sa, flags)

        // DISPLAY THE INFORMATION
        shape.Text <- System.String.Format("SetResults")

        ()


    static member Page_GetFormulas( doc: Microsoft.Office.Interop.Visio.Document ) =

        let page = VisioInterop.Util.CreateStandardPage(doc,"PGF");
        let shape= VisioInterop.Util.CreateStandardShape(page);
        let request= VisioInterop.Util.Create_PGF_Request(shape);

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        //let SRCStream  = Array.create (request.Length*3) ((int16)0)
        let SRCStream  = FSharp_Samples.CreateShortArray( request.Length*4)

        for i in 0 .. request.Length-1 do
            let item = request.[i]
            SRCStream.[i*4 + 0] <- item.ShapeID
            SRCStream.[i*4 + 1] <- item.CellSRC.SectionIndex
            SRCStream.[i*4 + 2] <- item.CellSRC.RowIndex
            SRCStream.[i*4 + 3] <- item.CellSRC.CellIndex

        //let formulas_sa : System.Array = null
        let formulas_sa_ref = ref null
        let SRCStream_sa : System.Array = SRCStream :> System.Array
        page.GetFormulasU(ref SRCStream_sa, formulas_sa_ref);

        // MAP OUTPUT BACK TO SOMETHING USEFUL 
        let empty = System.String.Empty
        let formulas = Array.create (request.Length) (empty)
        formulas_sa_ref.Value.CopyTo(formulas, 0);

        // DISPLAY THE INFORMATION
        shape.Text <- System.String.Format("Formulas={0},{1}", formulas.[0], formulas.[1]);

        ()


    static member Page_GetResults( doc: Microsoft.Office.Interop.Visio.Document ) =

        let page = VisioInterop.Util.CreateStandardPage(doc,"PGR");
        let shape= VisioInterop.Util.CreateStandardShape(page);
        let request= VisioInterop.Util.Create_PGR_Request(shape);

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        let SRCStream  = FSharp_Samples.CreateShortArray(request.Length*4)
        let unitcodes = FSharp_Samples.CreateObjectArray(request.Length)
        for i in 0 .. request.Length-1 do
            let item = request.[i]
            SRCStream.[i*4 + 0] <- item.ShapeID
            SRCStream.[i*4 + 1] <- item.CellSRC.SectionIndex
            SRCStream.[i*4 + 2] <- item.CellSRC.RowIndex
            SRCStream.[i*4 + 3] <- item.CellSRC.CellIndex
            unitcodes.[i] <- box ( (int16)Microsoft.Office.Interop.Visio.VisUnitCodes.visNoCast )

        //let formulas_sa : System.Array = null
        let flags = (int16)Microsoft.Office.Interop.Visio.VisGetSetArgs.visGetFloats;
        let results_sa_ref = ref null
        let SRCStream_sa : System.Array = SRCStream :> System.Array
        let unitcodes_sa : System.Array = unitcodes :> System.Array
        page.GetResults(ref SRCStream_sa, flags, ref unitcodes_sa, results_sa_ref);

        // MAP OUTPUT BACK TO SOMETHING USEFUL 
        let empty = System.String.Empty
        let results = Array.create (request.Length) (0.0)
        results_sa_ref.Value.CopyTo(results, 0);

        // DISPLAY THE INFORMATION
        shape.Text <- System.String.Format("Formulas={0},{1}", results.[0], results.[1]);

        ()


    static member Page_SetFormulas( doc: Microsoft.Office.Interop.Visio.Document ) =

        let page = VisioInterop.Util.CreateStandardPage(doc,"PSF");
        let shape= VisioInterop.Util.CreateStandardShape(page);
        let request= VisioInterop.Util.Create_PSF_Request(shape);

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        let SRCStream  = FSharp_Samples.CreateShortArray( request.Length*4)
        let formulas = FSharp_Samples.CreateObjectArray(request.Length)
        for i in 0 .. request.Length-1 do
            let item = request.[i]
            SRCStream.[i*4 + 0] <- item.ShapeID
            SRCStream.[i*4 + 1] <- item.CellSRC.SectionIndex
            SRCStream.[i*4 + 2] <- item.CellSRC.RowIndex
            SRCStream.[i*4 + 3] <- item.CellSRC.CellIndex
            formulas.[i] <- item.Formula :> System.Object

        //let formulas_sa : System.Array = null
        let flags = (int16)(Microsoft.Office.Interop.Visio.VisGetSetArgs.visSetBlastGuards ||| Microsoft.Office.Interop.Visio.VisGetSetArgs.visSetUniversalSyntax);
        let formulas_sa : System.Array = formulas :> System.Array
        let SRCStream_sa : System.Array = SRCStream :> System.Array
        let count = page.SetFormulas(ref SRCStream_sa, ref formulas_sa, flags)

        // DISPLAY THE INFORMATION
        shape.Text <- System.String.Format("SetFormulas")

        ()

    static member Page_SetResults( doc: Microsoft.Office.Interop.Visio.Document ) =

        let page = VisioInterop.Util.CreateStandardPage(doc,"PSR");
        let shape= VisioInterop.Util.CreateStandardShape(page);
        let request= VisioInterop.Util.Create_PSR_Request(shape);

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        let SRCStream  = FSharp_Samples.CreateShortArray(request.Length*4)
        let results = FSharp_Samples.CreateObjectArray(request.Length)
        let unitcodes = FSharp_Samples.CreateObjectArray(request.Length)
        for i in 0 .. request.Length-1 do
            let item = request.[i]
            SRCStream.[i*4 + 0] <- item.ShapeID
            SRCStream.[i*4 + 1] <- item.CellSRC.SectionIndex
            SRCStream.[i*4 + 2] <- item.CellSRC.RowIndex
            SRCStream.[i*4 + 3] <- item.CellSRC.CellIndex
            results.[i] <- box ( item.Result )
            unitcodes.[i] <- box ( (int16)item.UnitCode)

        //let formulas_sa : System.Array = null
        let flags = (int16)(0);
        let results_sa : System.Array = results :> System.Array
        let unitcodes_sa : System.Array = unitcodes :> System.Array
        let SRCStream_sa : System.Array = SRCStream :> System.Array
        let count = page.SetResults(ref SRCStream_sa, ref unitcodes_sa, ref results_sa, flags)

        // DISPLAY THE INFORMATION
        shape.Text <- System.String.Format("SetResults")

        ()


    static member CreateShortArray( length: int ) : int16 [] =
        Array.create (length) ((int16)0)

    static member CreateObjectArray( length: int ) : System.Object [] =
        Array.create (length) (null)



[<EntryPoint>]
let main (args : string[]) =

    let visapp = new Microsoft.Office.Interop.Visio.ApplicationClass()
    let doc = visapp.Documents.Add("")

    FSharp_Samples.Shape_GetFormulas(doc) |> ignore
    FSharp_Samples.Shape_GetResults(doc) |> ignore
    FSharp_Samples.Shape_SetFormulas(doc) |> ignore
    FSharp_Samples.Shape_SetResults(doc) |> ignore

    FSharp_Samples.Page_GetFormulas(doc) |> ignore
    FSharp_Samples.Page_GetResults(doc) |> ignore
    FSharp_Samples.Page_SetFormulas(doc) |> ignore
    FSharp_Samples.Page_SetResults(doc) |> ignore

    0
