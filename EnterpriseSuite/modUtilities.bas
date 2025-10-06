Attribute VB_Name = "modUtilities"
Option Explicit

Private Const LOG_FILE_NAME As String = "InsuranceSuite.log"

Public Sub Log(ByVal message As String)
    On Error Resume Next
    Dim fileNo As Integer
    fileNo = FreeFile
    Open CombinePath(App.Path, LOG_FILE_NAME) For Append As #fileNo
    Print #fileNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - " & message
    Close #fileNo
End Sub

Public Function CombinePath(ByVal basePath As String, ByVal relativePath As String) As String
    If Right$(basePath, 1) = "\" Then
        CombinePath = basePath & relativePath
    Else
        CombinePath = basePath & "\" & relativePath
    End If
End Function

Public Function AppDataPath() As String
    AppDataPath = CombinePath(App.Path, "data")
End Function

Public Sub EnsureFolderExists(ByVal folderPath As String)
    If Dir$(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
End Sub

Public Function FileExists(ByVal filePath As String) As Boolean
    FileExists = (Dir$(filePath, vbNormal) <> "")
End Function

Public Function FormatCurrencyValue(ByVal amount As Variant) As String
    If IsNull(amount) Or amount = "" Then
        FormatCurrencyValue = "$0.00"
    Else
        FormatCurrencyValue = Format$(amount, "Currency")
    End If
End Function

Public Function FormatDateValue(ByVal value As Variant) As String
    If IsNull(value) Or value = "" Then
        FormatDateValue = "N/A"
    Else
        FormatDateValue = Format$(CDate(value), "yyyy-mm-dd")
    End If
End Function

Public Function NullToString(ByVal value As Variant) As String
    If IsNull(value) Then
        NullToString = ""
    Else
        NullToString = CStr(value)
    End If
End Function

Public Function NullToCurrency(ByVal value As Variant) As Currency
    If IsNull(value) Then
        NullToCurrency = 0
    Else
        NullToCurrency = CCur(value)
    End If
End Function

Public Function NullToDate(ByVal value As Variant) As Date
    If IsNull(value) Then
        NullToDate = 0
    Else
        NullToDate = CDate(value)
    End If
End Function
