<%
DynamicInclude_MustBeGlobal = false
' Checks if File does exist in System.
'
' @param {string} File
' @param {FileSystemObject} System
' @return {bool}
Function FileDoExist( File, System )
    FileDoExist = System.FileExists(Server.MapPath(File))
End Function
' Checks if File exist.
'
' @param {string} File
' @return {bool}
Function FileExists(File)
    Set System = FileSystem()

    FileExists = FileDoExist(File, System)
    Set System = Nothing
End Function
' Syntax sugar.
'
' @return {FileSystemObject}
Function FileSystem( )
    Set FileSystem = Server.CreateObject("Scripting.FileSystemObject")
End Function
' Reads File as ASP code.
'
' @param {string} File
' @return {string}
Function ParseFile( File )
    Dim CloseTag
    Dim JoinSize
    Dim LineJoin
    Dim Match
    Dim Rows
    Dim OpenTag
    Dim PlainRow
    Dim Regex
    Dim WriteTag

    DynamicInclude_MustBeGlobal = false

    CloseTag = Chr(37)&Chr(62)
    LineJoin = """ & VbCrLf & """
    JoinSize = LEN(LineJoin)
    OpenTag  = Chr(60)&Chr(37)
    WriteTag = OpenTag & "="

    Set Regex = new RegExp
    With Regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
    End With
    ' Reading file
    ParseFile = ReadFile(File)
    ' Removing out comments
    Regex.Pattern = "^\s*'.*$"
    ParseFile = Regex.Replace(ParseFile, "")
    ' Removing duplicate new lines
    Regex.Pattern = "(?:\s*\n)+"
    ParseFile = Regex.Replace(ParseFile, VbCrLf)
    ' Replacing write tag with expanded command
    ParseFile = Replace(ParseFile, WriteTag, OpenTag & " Response.Write ")
    ' Adding ASP tags when needed
    If LEFT(ParseFile, 2) <> OpenTag Then
        ParseFile = OpenTag & CloseTag & VbCrLf & ParseFile
    End If
    If RIGHT(ParseFile, 2) <> CloseTag Then
        ParseFile =  ParseFile & VbCrLf & OpenTag & CloseTag
    End If
    ' Replacing free percentage symbols (but keeping ASP Tags)
    ParseFile = Mid(ParseFile, 3, LEN(ParseFile) -4)
    ParseFile = Replace(ParseFile, "%", "&percnt;")
    ParseFile = Replace(ParseFile, "<&percnt;", OpenTag)
    ParseFile = Replace(ParseFile, "&percnt;>", CloseTag)
    ' Inserting plain texts in Response.write commands
    Regex.Pattern = CloseTag & "([^%])+" & OpenTag
    Set Rows = Regex.Execute(ParseFile)
    For Each Match in Rows
        ' Removing delimiting close and open tags
        PlainRow = Mid(Match.value, 3, LEN(Match.value) -4)
        if (LEFT(PlainRow, 2) = VbCrLf) Then
            PlainRow = Right(PlainRow, LEN(PlainRow) - 2)
        end if
        ' Doubling quotes
        PlainRow = Replace(PlainRow, """", """""")
        ' Replacing new lines with line joins
        PlainRow = Replace(PlainRow, VbCrLf, LineJoin)
        if (Right(PlainRow, JoinSize) = LineJoin) Then
            ' Keeping last quote
            PlainRow = LEFT(PlainRow, LEN(PlainRow) - JoinSize)
        end if
        PlainRow = VbCrLf & "Response.write """ & PlainRow & """" & VbCrLf
        ' Replacing original line with corrected one
        ParseFile = Replace(ParseFile, Match.value, PlainRow)
    Next
    ' Detecting class declarations
    Regex.Pattern = OpenTag & "(?:.|\n)+class(?:.|\n)+end class(?:.|\n)+" & CloseTag
    Set Rows = Regex.Execute(ParseFile)
    if Rows.Count <> 0 then
        DynamicInclude_MustBeGlobal = true
    end if

    Set Regex = Nothing
End Function
' Reads File as plain text.
'
' @param {string} File
' @return {string}
Function ReadFile(File)
    Set System = FileSystem()
    Set FileData = System.OpenTextFile(Server.MapPath(File), 1)

    if FileDoExist(File, System) then
        ReadFile = FileData.ReadAll
    else
        Call Err.Raise(53)
    end if

    Set FileData = Nothing
    Set System = Nothing
End Function
' Tries to include File on current script.
' In case of failure, continues silently.
'
' @param {string} File
Sub Include(File)
    On Error Resume Next
        ' Execute the code or fails silently
        ' Uses class declaration detection to prevent 
        if DynamicInclude_MustBeGlobal then
            Call ExecuteGlobal(ParseFile(File))
        else
            Call Execute(ParseFile(File))
        end if
    On Error Goto 0
End Sub
' Tries to include File on current script.
' In case of failure, ends script execution with error message.
'
' @param {string} File
Sub Require(File)
    On Error Resume Next
        ' Execute the code
        Call Execute(ParseFile(File))
        ' If an error occurs, detect it here and stop parsing after displaying error.
        If Err.Number > 0 Then
            Response.Write "---FATAL ERROR: while trying to execute <b>" & File & "</b>---<br>" & VbCrLf
            Response.Write "---Error description: " & Err.Description & "<br>" & VbCrLf
            Response.Write "---Error Source: " & Err.Source & "<hr>" & VbCrLf
            Response.Write "---Error Line: " & Err.Line & "<hr>" & VbCrLf
            Call Response.End()
        End If
    On Error Goto 0
End Sub
%>