Attribute VB_Name = "modTestRunner"

Option Explicit

Public Function RunCoreTests() As Long
    Dim passed As Long
    Dim failed As Long
    Dim currentTest As String
    Dim summary As String

    On Error GoTo UnexpectedError

    TestLogInit
    TestLogLine "Core tests run: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    TestLogLine "Details file: " & TestLogPath()

    currentTest = "Test_CoreAppendUniqueToken"
    Test_CoreAppendUniqueToken passed, failed
    currentTest = "Test_CoreTruncateText"
    Test_CoreTruncateText passed, failed
    currentTest = "Test_CoreBuildAutoComment"
    Test_CoreBuildAutoComment passed, failed
    currentTest = "Test_CoreBuildNombreBase"
    Test_CoreBuildNombreBase passed, failed
    currentTest = "Test_CoreIsFCE"
    Test_CoreIsFCE passed, failed
    currentTest = "Test_CoreTipoDocFromRetailWeb"
    Test_CoreTipoDocFromRetailWeb passed, failed
    currentTest = "Test_CoreShouldMarkPendienteRevisar"
    Test_CoreShouldMarkPendienteRevisar passed, failed
    currentTest = "Test_CoreIsSiteMismatch"
    Test_CoreIsSiteMismatch passed, failed
    currentTest = "Test_CoreBuildSiteMismatchComment"
    Test_CoreBuildSiteMismatchComment passed, failed
    currentTest = "Test_CoreBuildDoaMessage"
    Test_CoreBuildDoaMessage passed, failed
    currentTest = "Test_CoreIsErrorReferencia"
    Test_CoreIsErrorReferencia passed, failed

    summary = "Core tests: " & passed & " passed, " & failed & " failed"
    Debug.Print summary
    TestLogLine summary
    RunCoreTests = failed
    GoTo CleanExit

UnexpectedError:
    failed = failed + 1
    Debug.Print "ERROR: " & currentTest & " err=" & Err.Number & " " & Err.Description
    TestLogLine "ERROR: " & currentTest & " err=" & Err.Number & " " & Err.Description
    summary = "Core tests: " & passed & " passed, " & failed & " failed"
    Debug.Print summary
    TestLogLine summary
    RunCoreTests = failed
    Err.Clear

CleanExit:
    TestLogFlushOrClose
End Function

Private Sub Test_CoreAppendUniqueToken(ByRef passed As Long, ByRef failed As Long)
    AssertEqual "CoreAppendUniqueToken/Appends", "A-B", CoreAppendUniqueToken("A", "B", "-"), passed, failed
    AssertEqual "CoreAppendUniqueToken/Dedup", "A-B", CoreAppendUniqueToken("A-B", "A", "-"), passed, failed
    AssertEqual "CoreAppendUniqueToken/EmptyBase", "Token", CoreAppendUniqueToken("", "Token", "-"), passed, failed
End Sub

Private Sub Test_CoreTruncateText(ByRef passed As Long, ByRef failed As Long)
    AssertEqual "CoreTruncateText/NoChange", "abc", CoreTruncateText("abc", 200), passed, failed
    AssertEqual "CoreTruncateText/LongNoSpaces", String(200, "A"), CoreTruncateText(String(201, "A"), 200), passed, failed
End Sub

Private Sub Test_CoreBuildAutoComment(ByRef passed As Long, ByRef failed As Long)
    Dim result As String

    result = CoreBuildAutoComment("", "", "Ok", "FC-REC", "REF1", "COMP1", 10, "2026-02-22-USER")
    AssertContains "CoreBuildAutoComment/FechaUser", result, "2026-02-22-USER", passed, failed
    AssertContains "CoreBuildAutoComment/Comp", result, "COMP1", passed, failed
    AssertContains "CoreBuildAutoComment/DifContra", result, "Dif. en contra:", passed, failed

    result = CoreBuildAutoComment("", "", "Pendiente de Reingreso", "FC-REC", "REF1", "", 10, "2026-02-22-USER")
    AssertNotContains "CoreBuildAutoComment/SkipDif", result, "Dif. en contra:", passed, failed
End Sub

Private Sub Test_CoreBuildNombreBase(ByRef passed As Long, ByRef failed As Long)
    Dim result As String
    result = CoreBuildNombreBase("SITE1", "FC-REM", "REF1", "01.01.2026", False, "Contabilizado")
    AssertEqual "CoreBuildNombreBase/Expected", "SITE1-FC-REM-REF1-Fecha base 01.01.2026-Sin RW-Contabilizado", result, passed, failed
End Sub

Private Sub Test_CoreIsFCE(ByRef passed As Long, ByRef failed As Long)
    AssertTrue "CoreIsFCE/True", CoreIsFCE(True, 100, 100), passed, failed
    AssertTrue "CoreIsFCE/FalseNoPyme", Not CoreIsFCE(False, 100, 100), passed, failed
    AssertTrue "CoreIsFCE/FalseThreshold", Not CoreIsFCE(True, 99, 100), passed, failed
End Sub

Private Sub Test_CoreTipoDocFromRetailWeb(ByRef passed As Long, ByRef failed As Long)
    AssertEqual "CoreTipoDocFromRetailWeb/FC-REC", "FC-REC", CoreTipoDocFromRetailWeb("1", "A123", False), passed, failed
    AssertEqual "CoreTipoDocFromRetailWeb/NCE-DEV", "NCE-DEV", CoreTipoDocFromRetailWeb("2", "A123", True), passed, failed
    AssertEqual "CoreTipoDocFromRetailWeb/FC-REM", "FC-REM", CoreTipoDocFromRetailWeb("1", "R123", False), passed, failed
    AssertEqual "CoreTipoDocFromRetailWeb/NC-REM", "NC-REM", CoreTipoDocFromRetailWeb("2", "R123", False), passed, failed
End Sub

Private Sub Test_CoreShouldMarkPendienteRevisar(ByRef passed As Long, ByRef failed As Long)
    AssertTrue "CoreShouldMarkPendienteRevisar/True", CoreShouldMarkPendienteRevisar(10, "", 5), passed, failed
    AssertTrue "CoreShouldMarkPendienteRevisar/FalseLow", Not CoreShouldMarkPendienteRevisar(3, "", 5), passed, failed
    AssertTrue "CoreShouldMarkPendienteRevisar/FalseDifConNC", Not CoreShouldMarkPendienteRevisar(10, 2, 5), passed, failed
    AssertTrue "CoreShouldMarkPendienteRevisar/TrueDifConNC", CoreShouldMarkPendienteRevisar(10, 6, 5), passed, failed
End Sub

Private Sub Test_CoreIsSiteMismatch(ByRef passed As Long, ByRef failed As Long)
    AssertTrue "CoreIsSiteMismatch/True", CoreIsSiteMismatch("A", "B"), passed, failed
    AssertTrue "CoreIsSiteMismatch/FalseSame", Not CoreIsSiteMismatch("A", "A"), passed, failed
    AssertTrue "CoreIsSiteMismatch/FalseEmpty", Not CoreIsSiteMismatch("", "A"), passed, failed
End Sub

Private Sub Test_CoreBuildSiteMismatchComment(ByRef passed As Long, ByRef failed As Long)
    AssertEqual "CoreBuildSiteMismatchComment/REM", "Anular: Ingresan un RTO. de la Sucursal X", CoreBuildSiteMismatchComment("FC-REM", "X"), passed, failed
    AssertEqual "CoreBuildSiteMismatchComment/REC", "Anular: Ingresan una FC de la Sucursal X", CoreBuildSiteMismatchComment("FC-REC", "X"), passed, failed
End Sub

Private Sub Test_CoreBuildDoaMessage(ByRef passed As Long, ByRef failed As Long)
    Dim msg As String
    msg = CoreBuildDoaMessage(DateSerial(2026, 2, 21), 0, 100, DateSerial(2026, 2, 22))
    AssertContains "CoreBuildDoaMessage/Prefix", msg, MSG_DOA_PREFIJO, passed, failed
    AssertContains "CoreBuildDoaMessage/Suffix", msg, MSG_DOA_SUFIJO, passed, failed
End Sub

Private Sub Test_CoreIsErrorReferencia(ByRef passed As Long, ByRef failed As Long)
    AssertTrue "CoreIsErrorReferencia/True", CoreIsErrorReferencia("Varios", "123", "FC-REC", "X", 13, "A"), passed, failed
    AssertTrue "CoreIsErrorReferencia/False", Not CoreIsErrorReferencia("Varios", "123456A789012", "FC-REC", "123456A789012", 13, "A"), passed, failed
    AssertTrue "CoreIsErrorReferencia/TrimmedValid", Not CoreIsErrorReferencia("Varios", "123456A789012", "FC-REC", " 123456A789012 ", 13, "A"), passed, failed
    AssertTrue "CoreIsErrorReferencia/Len12", CoreIsErrorReferencia("Varios", "12345A789012", "FC-REC", "12345A789012", 13, "A"), passed, failed
End Sub

Private Sub AssertEqual(ByVal name As String, ByVal expected As Variant, ByVal actual As Variant, ByRef passed As Long, ByRef failed As Long)
    Dim message As String

    If expected = actual Then
        passed = passed + 1
    Else
        failed = failed + 1
        message = "FAIL: " & name & " expected=" & expected & " actual=" & actual
        Debug.Print message
        TestLogLine message
    End If
End Sub

Private Sub AssertTrue(ByVal name As String, ByVal condition As Boolean, ByRef passed As Long, ByRef failed As Long)
    Dim message As String

    If condition Then
        passed = passed + 1
    Else
        failed = failed + 1
        message = "FAIL: " & name & " expected=True actual=False"
        Debug.Print message
        TestLogLine message
    End If
End Sub

Private Sub AssertContains(ByVal name As String, ByVal value As String, ByVal fragment As String, ByRef passed As Long, ByRef failed As Long)
    Dim message As String

    If InStr(1, value, fragment) > 0 Then
        passed = passed + 1
    Else
        failed = failed + 1
        message = "FAIL: " & name & " expected_contains=" & fragment & " actual=" & value
        Debug.Print message
        TestLogLine message
    End If
End Sub

Private Sub AssertNotContains(ByVal name As String, ByVal value As String, ByVal fragment As String, ByRef passed As Long, ByRef failed As Long)
    Dim message As String

    If InStr(1, value, fragment) = 0 Then
        passed = passed + 1
    Else
        failed = failed + 1
        message = "FAIL: " & name & " expected_not_contains=" & fragment & " actual=" & value
        Debug.Print message
        TestLogLine message
    End If
End Sub
