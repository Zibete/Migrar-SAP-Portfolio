Attribute VB_Name = "modCoreValidationTests"
'@TestModule
Option Explicit

'@TestMethod("CoreShouldMarkPendienteRevisar")
Public Sub CoreShouldMarkPendienteRevisar_Basic()
    AssertTrue "CoreShouldMarkPendienteRevisar/True", CoreShouldMarkPendienteRevisar(10, "", 5)
    AssertTrue "CoreShouldMarkPendienteRevisar/FalseLow", Not CoreShouldMarkPendienteRevisar(3, "", 5)
    AssertTrue "CoreShouldMarkPendienteRevisar/FalseDifConNC", Not CoreShouldMarkPendienteRevisar(10, 2, 5)
    AssertTrue "CoreShouldMarkPendienteRevisar/TrueDifConNC", CoreShouldMarkPendienteRevisar(10, 6, 5)
End Sub

'@TestMethod("CoreIsSiteMismatch")
Public Sub CoreIsSiteMismatch_Basic()
    AssertTrue "CoreIsSiteMismatch/True", CoreIsSiteMismatch("A", "B")
    AssertTrue "CoreIsSiteMismatch/FalseSame", Not CoreIsSiteMismatch("A", "A")
    AssertTrue "CoreIsSiteMismatch/FalseEmpty", Not CoreIsSiteMismatch("", "A")
End Sub

'@TestMethod("CoreBuildSiteMismatchComment")
Public Sub CoreBuildSiteMismatchComment_Rem()
    AssertEqual "CoreBuildSiteMismatchComment/REM", "Anular: Ingresan un RTO. de la Sucursal X", CoreBuildSiteMismatchComment("FC-REM", "X")
    AssertEqual "CoreBuildSiteMismatchComment/REC", "Anular: Ingresan una FC de la Sucursal X", CoreBuildSiteMismatchComment("FC-REC", "X")
End Sub

'@TestMethod("CoreBuildDoaMessage")
Public Sub CoreBuildDoaMessage_Basic()
    Dim msg As String
    msg = CoreBuildDoaMessage(DateSerial(2026, 2, 21), 0, 100, DateSerial(2026, 2, 22))
    AssertContains "CoreBuildDoaMessage/Prefix", msg, MSG_DOA_PREFIJO
    AssertContains "CoreBuildDoaMessage/Suffix", msg, MSG_DOA_SUFIJO
End Sub

'@TestMethod("CoreIsErrorReferencia")
Public Sub CoreIsErrorReferencia_Varios()
    AssertTrue "CoreIsErrorReferencia/True", CoreIsErrorReferencia("Varios", "123", "FC-REC", "X", 13, "A")
    AssertTrue "CoreIsErrorReferencia/False", Not CoreIsErrorReferencia("Varios", "123456A789012", "FC-REC", "123456A789012", 13, "A")
    AssertTrue "CoreIsErrorReferencia/TrimmedValid", Not CoreIsErrorReferencia("Varios", "123456A789012", "FC-REC", " 123456A789012 ", 13, "A")
    AssertTrue "CoreIsErrorReferencia/Len12", CoreIsErrorReferencia("Varios", "12345A789012", "FC-REC", "12345A789012", 13, "A")
End Sub
