Attribute VB_Name = "modCoreClassificationTests"
'@TestModule
Option Explicit

'@TestMethod("CoreIsFCE")
Public Sub CoreIsFCE_WhenPymeAndThreshold()
    AssertTrue "CoreIsFCE/True", CoreIsFCE(True, 100, 100)
    AssertTrue "CoreIsFCE/FalseNoPyme", Not CoreIsFCE(False, 100, 100)
    AssertTrue "CoreIsFCE/FalseThreshold", Not CoreIsFCE(True, 99, 100)
End Sub

'@TestMethod("CoreTipoDocFromRetailWeb")
Public Sub CoreTipoDocFromRetailWeb_Rec()
    AssertEqual "CoreTipoDocFromRetailWeb/FC-REC", "FC-REC", CoreTipoDocFromRetailWeb("1", "A123", False)
    AssertEqual "CoreTipoDocFromRetailWeb/NCE-DEV", "NCE-DEV", CoreTipoDocFromRetailWeb("2", "A123", True)
End Sub

'@TestMethod("CoreTipoDocFromRetailWeb")
Public Sub CoreTipoDocFromRetailWeb_Rem()
    AssertEqual "CoreTipoDocFromRetailWeb/FC-REM", "FC-REM", CoreTipoDocFromRetailWeb("1", "R123", False)
    AssertEqual "CoreTipoDocFromRetailWeb/NC-REM", "NC-REM", CoreTipoDocFromRetailWeb("2", "R123", False)
End Sub
