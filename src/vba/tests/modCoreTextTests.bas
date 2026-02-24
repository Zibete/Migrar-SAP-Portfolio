Attribute VB_Name = "modCoreTextTests"
'@TestModule
Option Explicit

'@TestMethod("CoreAppendUniqueToken")
Public Sub CoreAppendUniqueToken_Appends()
    AssertEqual "CoreAppendUniqueToken/Appends", "A-B", CoreAppendUniqueToken("A", "B", "-")
End Sub

'@TestMethod("CoreAppendUniqueToken")
Public Sub CoreAppendUniqueToken_Dedup()
    AssertEqual "CoreAppendUniqueToken/Dedup", "A-B", CoreAppendUniqueToken("A-B", "A", "-")
End Sub

'@TestMethod("CoreAppendUniqueToken")
Public Sub CoreAppendUniqueToken_EmptyBase()
    AssertEqual "CoreAppendUniqueToken/EmptyBase", "Token", CoreAppendUniqueToken("", "Token", "-")
End Sub

'@TestMethod("CoreTruncateText")
Public Sub CoreTruncateText_NoChange()
    AssertEqual "CoreTruncateText/NoChange", "abc", CoreTruncateText("abc", 200)
End Sub

'@TestMethod("CoreTruncateText")
Public Sub CoreTruncateText_LongNoSpaces()
    Dim inputValue As String
    inputValue = String(201, "A")
    AssertEqual "CoreTruncateText/LongNoSpaces", String(200, "A"), CoreTruncateText(inputValue, 200)
End Sub

'@TestMethod("CoreBuildAutoComment")
Public Sub CoreBuildAutoComment_IncludesDiff()
    Dim result As String
    result = CoreBuildAutoComment("", "", "Ok", "FC-REC", "REF1", "COMP1", 10, "2026-02-22-USER")
    AssertContains "CoreBuildAutoComment/FechaUser", result, "2026-02-22-USER"
    AssertContains "CoreBuildAutoComment/Comp", result, "COMP1"
    AssertContains "CoreBuildAutoComment/DifContra", result, "Dif. en contra:"
End Sub

'@TestMethod("CoreBuildAutoComment")
Public Sub CoreBuildAutoComment_SkipDiffOnPendienteReingreso()
    Dim result As String
    result = CoreBuildAutoComment("", "", "Pendiente de Reingreso", "FC-REC", "REF1", "", 10, "2026-02-22-USER")
    AssertNotContains "CoreBuildAutoComment/SkipDif", result, "Dif. en contra:"
End Sub

'@TestMethod("CoreBuildNombreBase")
Public Sub CoreBuildNombreBase_BuildsExpectedName()
    Dim result As String
    result = CoreBuildNombreBase("SITE1", "FC-REM", "REF1", "01.01.2026", False, "Contabilizado")
    AssertEqual "CoreBuildNombreBase/Expected", "SITE1-FC-REM-REF1-Fecha base 01.01.2026-Sin RW-Contabilizado", result
End Sub
