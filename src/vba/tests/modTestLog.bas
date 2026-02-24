Attribute VB_Name = "modTestLog"

Option Explicit

Private mLogFileNum As Integer
Private mLogIsOpen As Boolean

Public Sub TestLogInit()

    On Error GoTo HandleError

    TestLogFlushOrClose
    TestLogEnsureArtifactsDir

    mLogFileNum = FreeFile
    Open TestLogPath() For Output As #mLogFileNum
    mLogIsOpen = True
    Exit Sub

HandleError:
    mLogIsOpen = False
    Debug.Print "WARN: TestLogInit failed: " & Err.Number & " - " & Err.Description
    Err.Clear

End Sub

Public Sub TestLogLine(ByVal text As String)

    On Error GoTo HandleError

    If Not mLogIsOpen Then
        TestLogEnsureArtifactsDir
        mLogFileNum = FreeFile
        Open TestLogPath() For Append As #mLogFileNum
        mLogIsOpen = True
    End If

    Print #mLogFileNum, text
    Exit Sub

HandleError:
    Debug.Print "WARN: TestLogLine failed: " & Err.Number & " - " & Err.Description
    Err.Clear

End Sub

Public Sub TestLogFlushOrClose()

    On Error Resume Next
    If mLogIsOpen Then
        Close #mLogFileNum
        mLogIsOpen = False
    End If
    On Error GoTo 0

End Sub

Public Function TestLogPath() As String

    TestLogPath = TestArtifactsDirPath() & Application.PathSeparator & "core-tests-details.txt"

End Function

Private Function TestArtifactsDirPath() As String

    TestArtifactsDirPath = ParentFolderPath(ThisWorkbook.Path) & Application.PathSeparator & "artifacts"

End Function

Private Function ParentFolderPath(ByVal folderPath As String) As String

    Dim pos As Long

    pos = InStrRev(folderPath, Application.PathSeparator)
    If pos > 0 Then
        ParentFolderPath = Left$(folderPath, pos - 1)
    Else
        ParentFolderPath = folderPath
    End If

End Function

Private Sub TestLogEnsureArtifactsDir()

    Dim artifactsPath As String

    artifactsPath = TestArtifactsDirPath()
    If Dir$(artifactsPath, vbDirectory) = vbNullString Then MkDir artifactsPath

End Sub
