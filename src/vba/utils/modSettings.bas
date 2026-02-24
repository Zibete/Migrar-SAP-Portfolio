Attribute VB_Name = "modSettings"

Public Function GetConfigValue(ByVal rangeName As String, Optional ByVal defaultValue As Variant = "") As Variant

    On Error GoTo Handle
    GetConfigValue = Hoja3.Range(rangeName).Value
    Exit Function

Handle:
    GetConfigValue = defaultValue
    On Error GoTo 0

End Function

Public Sub SetConfigValue(ByVal rangeName As String, ByVal value As Variant)

    On Error GoTo Handle
    Hoja3.Range(rangeName).Value = value

Handle:
    On Error GoTo 0

End Sub

Public Function GetVendorFilter() As String
    GetVendorFilter = CStr(GetConfigValue("Vend", ""))
End Function

Public Function GetRutaCarpeta() As String
    GetRutaCarpeta = CStr(GetConfigValue("Ruta", ""))
End Function

Public Sub SetRutaCarpeta(ByVal ruta As String)
    SetConfigValue "Ruta", ruta
End Sub

Public Function GetOrigenDatos() As String
    GetOrigenDatos = CStr(GetConfigValue("origenDatos", ""))
End Function

Public Function GetMantenerDatos() As String
    GetMantenerDatos = CStr(GetConfigValue("mantenerDatos", ""))
End Function

Public Function GetPagoPendiente() As String
    GetPagoPendiente = CStr(GetConfigValue("PagoPendiente", ""))
End Function

Public Function GetEliminarDuplicados() As String
    GetEliminarDuplicados = CStr(GetConfigValue("EliminarDuplicados", ""))
End Function
