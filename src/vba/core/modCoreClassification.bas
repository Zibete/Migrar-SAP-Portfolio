Attribute VB_Name = "modCoreClassification"

' Pure classification helpers (no Excel/SAP/IE access).

Public Function CoreIsFCE(ByVal esPyme As Boolean, ByVal totalBruto As Double, ByVal montoFCE As Double) As Boolean
    CoreIsFCE = (esPyme And totalBruto >= montoFCE)
End Function

Public Function CoreTipoDocFromRetailWeb(ByVal retailWebValue As Variant, ByVal remitoRef As String, ByVal esFCE As Boolean) As String

    Dim prefijo As String

    prefijo = Left(CStr(retailWebValue), 1)

    If InStr(1, remitoRef, "A") > 0 Or InStr(1, remitoRef, "C") > 0 Then
        If prefijo = "1" And Not esFCE Then CoreTipoDocFromRetailWeb = "FC-REC"
        If prefijo = "1" And esFCE Then CoreTipoDocFromRetailWeb = "FCE-REC"
        If prefijo = "2" And Not esFCE Then CoreTipoDocFromRetailWeb = "NC-DEV"
        If prefijo = "2" And esFCE Then CoreTipoDocFromRetailWeb = "NCE-DEV"
    ElseIf InStr(1, remitoRef, "R") > 0 Then
        If prefijo = "1" Then CoreTipoDocFromRetailWeb = "FC-REM"
        If prefijo = "2" Then CoreTipoDocFromRetailWeb = "NC-REM"
    End If

End Function
