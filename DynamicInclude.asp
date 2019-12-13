<%
' Self-explained Constant
Const DI_CurrentFolder = "./"

' "Constants" (needed to be fake), declared rightaway to avoid multipe string
' concatenations at runtime
' ASP block closing tag
Private DI_CloseTag : DI_CloseTag = Chr(37)&Chr(62)
' ASP block opening tag
Private DI_OpenTag : DI_OpenTag = Chr(60)&Chr(37)
' Opening tag followed by a closing tag
Private DI_OpenClose : DI_OpenClose = DI_OpenTag & DI_CloseTag
Private DI_OpenBreakClose : DI_OpenBreakClose = DI_OpenTag & (vbNewLine & DI_CloseTag)
' Closing tag followed by a opening tag
Private DI_CloseOpen : DI_CloseOpen = DI_CloseTag & DI_OpenTag
Private DI_CloseBreakOpen : DI_CloseBreakOpen = DI_CloseTag & (vbNewLine & DI_OpenTag)

' ASP write block opening tag
Private DI_WriteTag : DI_WriteTag = DI_OpenTag & "="
' ASP common Response.Write statement
Private DI_ResponseWrite : DI_ResponseWrite = DI_OpenTag & " Response.Write "
' Dynamic require call
Private DI_Require : DI_Require = DI_OpenTag & (" Require($1) " & DI_CloseTag)



' Public variables
' Path of the last included file, used for recursive file importing.
' May be used along with *DynamicInclude_CurrentPath* to change how recursive imports will behave.
Public DynamicInclude_PreviousPath : DynamicInclude_PreviousPath = vbNullString
' Path of current file to include, used for recursive file importing.
' May be used along with *DynamicInclude_CurrentPath* to change how recursive imports will behave.
Public DynamicInclude_CurrentPath  : DynamicInclude_CurrentPath  = DI_CurrentFolder
' If HTML text should be trimmed while parsing.
Public DynamicInclude_TrimHtml : DynamicInclude_TrimHtml = false
' If duplicated new lines hould be removed from parsed text.
Public DynamicInclude_TrimNewlines : DynamicInclude_TrimNewlines = false



' Executes an imported file in the global namespace, dealing with paths for recursive calls.
'
' @param {string} File
Sub ExecuteFile( ByVal File )
    Const DI_Bar = "/"
    Const DI_ReverseBar = "\"

    Dim BarIndex
    Dim Parsed
    Dim Path

    'Path = FilePath(File)
    Path = Replace(File, DI_ReverseBar, DI_Bar)
    BarIndex = InStrRev(Path, DI_Bar)
    if IsNull(BarIndex) or BarIndex = 0 then
        Path = vbNullString
    else
        Path = Mid(File, 1, BarIndex)
    end if

    ' Path operations for recursive file importing
    DynamicInclude_PreviousPath = DynamicInclude_CurrentPath
    DynamicInclude_CurrentPath = DynamicInclude_CurrentPath & Path

    'File = FileName(File)
    File = Replace(File, DI_ReverseBar, DI_Bar)
    BarIndex = InStrRev(File, DI_Bar)
    File = Mid(File, BarIndex + 1, LEN(File) - BarIndex)

    ' Parsing ASP file
    Parsed = ParseFile(DynamicInclude_CurrentPath & File)
    ' Always importing in global namespace (to prevent errors)
    ExecuteGlobal Trim(Parsed)

    ' Restoring path operation
    DynamicInclude_CurrentPath = DI_CurrentFolder
    DynamicInclude_PreviousPath = vbNullString
End Sub
' Reads File as ASP code.
'
' @param {string} File
' @return {string}
Function ParseFile(ByRef File)
    ' Operator uses to join lines while parsing
    Const DI_LineJoin = """ & vbNewLine & """
    ' Self-explaining constants
    Const DI_Space = " "
    Const DI_Quote = """"
    Const DI_DoubleQuote = """"""

    ' Size of the line-joining operator
    Dim DI_JoinSize : DI_JoinSize = LEN(DI_LineJoin)

    Dim LineArray(4)
    Dim KeepLeft
    Dim KeepRight
    Dim Match
    Dim Rows
    Dim PlainRow
    Dim Regex

    ' Reading file
    ParseFile = ReadFile(File)

    ' Adding nem line separator in tags (otherwise the rows regex will fail)
    ParseFile = Replace(ParseFile, DI_CloseOpen, DI_CloseBreakOpen)
    ParseFile = Replace(ParseFile, DI_OpenClose, DI_OpenBreakClose)
    ' Replacing write tag with expanded command
    ParseFile = Replace(ParseFile, DI_WriteTag, DI_ResponseWrite)

    Set Regex = new RegExp
    With Regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = "<!--\s*#include file=(""[^""]+"")\s*-->"
    End With
    ' Converting ASP imports to Require calls
    ParseFile = Regex.Replace(ParseFile, DI_Require)

    ' Adding ASP tags when needed
    If LEFT(ParseFile, 2) <> DI_OpenTag Then
        ParseFile = (DI_OpenClose & vbNewLine) & ParseFile
    End If
    If RIGHT(ParseFile, 2) <> DI_CloseTag Then
        ParseFile =  ParseFile & (vbNewLine & DI_OpenClose)
    End If
    ' Trimming fist open tag and last close tag
    ParseFile = Mid(ParseFile, 3, LEN(ParseFile) -4)
    ' Replacing free percentage symbols (but keeping ASP Tags)
    ParseFile = Replace(ParseFile, "%", "&percnt;")
    ParseFile = Replace(ParseFile, "<&percnt;", DI_OpenTag)
    ParseFile = Replace(ParseFile, "&percnt;>", DI_CloseTag)
    if DynamicInclude_TrimHtml then
        ' Trim tags (except text content):
        Regex.Pattern = "([^%]>)\s+(<[^%]\/?)"
        ParseFile = Regex.Replace(ParseFile, "$1$2")
        ' Trim text content (except <pre>formatted text:
        Regex.Pattern = "([^%](?!pre)>)\s*([^\s]+)\s*(<[^%](?!pre))"
        ParseFile = Regex.Replace(ParseFile, "$1$2$3")
    end if

    ' Inserting plain texts in Response.write commands
    Regex.Pattern = (DI_CloseTag & "([^%])+") & DI_OpenTag
    Set Rows = Regex.Execute(ParseFile)
    ' Using fixed array to reduce string parsing
    LineArray(0) = vbNewLine
    LineArray(1) = "Response.write """
    LineArray(3) = DI_Quote
    LineArray(4) = vbNewLine
    For Each Match in Rows
        ' Removing delimiting close and open tags
        PlainRow = Mid(Match.value, 3, LEN(Match.value) -4)
        if (LEFT(PlainRow, 2) = vbNewLine) Then
            PlainRow = Right(PlainRow, LEN(PlainRow) - 2)
        end if
        ' Doubling quotes
        PlainRow = Replace(PlainRow, DI_Quote, DI_DoubleQuote)
        ' Replacing new lines with line joins
        PlainRow = Replace(PlainRow, vbNewLine, DI_LineJoin)
        if (Right(PlainRow, DI_JoinSize) = DI_LineJoin) Then
            ' Keeping last quote
            PlainRow = LEFT(PlainRow, LEN(PlainRow) - DI_JoinSize)
        end if
        if DynamicInclude_TrimHtml then
            KeepLeft = (Left(PlainRow, 1) = DI_Space)
            KeepRight = (Right(PlainRow, 1) = DI_Space)

            PlainRow = Trim(PlainRow)
            if KeepLeft Then
                PlainRow = DI_Space & PlainRow
            end if
            if KeepRight Then
                PlainRow = PlainRow & DI_Space
            end if
        end if

        LineArray(2) = PlainRow
        ' Replacing original line with corrected one
        ParseFile = Replace(ParseFile, Match.value, Join(LineArray, vbNullString))
    Next
    Set Rows = Nothing : Set Match = Nothing : Erase LineArray
    ' Removing empty Response.write commands
    ParseFile = Replace(ParseFile, "Response.write """"" & vbNewLine, vbNullString)
    ' Removing duplicate new lines
    if DynamicInclude_TrimNewlines then
        Regex.Pattern = "(?:\s+\n)+"
        ParseFile = Regex.Replace(ParseFile, vbNewLine)
    end if
    ' Unescaping &percnt symbols
    ParseFile = Replace(ParseFile, "&percnt;", "%")

    Set Regex = Nothing
End Function
' Reads File as plain text.
'
' @param {string} File
' @return {string}
Function ReadFile(ByRef File)
    Dim System : Set System = Server.CreateObject("Scripting.FileSystemObject")
    Dim Path
    Dim FileData

    if File = vbNullString then
        File = DI_CurrentFolder
    end if

    Path = Server.MapPath(File)

    if System.FileExists(Path) then
        Set FileData = System.OpenTextFile(Path, 1)

        ReadFile = FileData.ReadAll

        Set FileData = Nothing
    else
        Err.Raise 53, Join(Array( _
            "File '", File, "' was not found'.", vbNewLine _
        ), vbNullString)
    end if

    Set System = Nothing
End Function
' Tries to include File on current script.
' In case of failure, continues silently.
'
' @param {string} File
Sub Include(ByVal File)
    On Error Resume Next
        ' Execute the code or fails silently
        ExecuteFile File
    On Error Goto 0
End Sub
' Tries to include File on current script.
' In case of failure, ends script execution with error message.
'
' @param {string} File
Sub Require(ByVal File)
    On Error Resume Next
        ' Execute the code
        ExecuteFile File
        ' If an error occurs, detect it here and stop execution.
        If Err.Number > 0 Then
            Response.Write Join(Array( _
                "---FATAL ERROR: while trying to execute <b>", File, "</b>---<br>", vbNewLine, _
                "---Error description: ", Err.Description, "<br>", vbNewLine, _
                "---Error Source: ", Err.Source, "<hr>", vbNewLine, _
                "---Error Line: ", Err.Line, "<hr>", vbNewLine _
            ), vbNullString)
            Response.End
        End If
    On Error Goto 0
End Sub
%>