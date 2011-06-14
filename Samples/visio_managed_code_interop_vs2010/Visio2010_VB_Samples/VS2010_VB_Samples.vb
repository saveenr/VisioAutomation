Imports IVisio = Microsoft.Office.Interop.Visio


Partial Public Class VS2010_VB_Samples
    Shared Function CreateShortArray(ByVal length As Integer) As System.Array
        Dim s = CType(Array.CreateInstance(GetType(Short), length), Short())
        Return s
    End Function

    Shared Function CreateObjectArray (ByVal length As Integer) As System.Array
        Dim s = CType (Array.CreateInstance (GetType (Object), length), Object())
        Return s
    End Function

    Shared Function CreateStringArray (ByVal length As Integer) As System.Array
        Dim a = CType (Array.CreateInstance (GetType (String), length), String())
        Return a
    End Function
End Class
