Module Module1
    Public cn As New ADODB.Connection
    Public rs As New ADODB.Recordset
    Sub con()
        With cn
            If .State = 1 Then .Close()
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=dbsample.mdb; Persist Security Info=false;"
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .Open()
        End With
    End Sub
End Module
