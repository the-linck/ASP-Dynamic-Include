<% ' Deprecated functionalities that may be useful for users

' Checks if File does exist in System.
' Deprecated due to unnecessary increase of call stack.
'
' @param {string} File
' @param {FileSystemObject} System
' @return {bool}
Function FileDoExist( ByRef File, ByRef System )
    FileDoExist = System.FileExists(Server.MapPath(File))
End Function

' Checks if File exist.
' Deprecated because it was not used.
'
' @param {string} File
' @return {bool}
Function FileExists(ByRef File)
    Dim System : Set System = Server.CreateObject("Scripting.FileSystemObject")

    FileExists = System.FileExists(Server.MapPath(File))
    Set System = Nothing
End Function

' Gets the name of the file.
' Deprecated due to unnecessary increase of call stack.
'
' @param {string} File
' @return {string}
Function FileName(ByRef File)
    Dim BarIndex

    File = Replace(File, DynamicInclude_ReverseBar, DynamicInclude_Bar)
    BarIndex = InStrRev(File, DynamicInclude_Bar)

    FileName = Mid(File, BarIndex + 1, LEN(File) - BarIndex)
End Function

' Gets the path of a file.
' Deprecated due to unnecessary increase of call stack.
'
' @param {string} File
' @return {string}
Function FilePath(File)
    Dim BarIndex

    File = Replace(File, DynamicInclude_ReverseBar, DynamicInclude_Bar)
    BarIndex = InStrRev(File, DynamicInclude_Bar)

    if IsNull(BarIndex) or BarIndex = 0 then
        FilePath = vbNullString
    else
        FilePath = Mid(File, 1, BarIndex)
    end if
End Function

' Syntax sugar.
' Deprecated because it was completely unnecessary.
'
' @return {FileSystemObject}
Function FileSystem( )
    Set FileSystem = Server.CreateObject("Scripting.FileSystemObject")
End Function

' Gets the absolute equivalent of Path.
' Deprecated due to unnecessary increase of call stack.
'
' @param {string} Path
' @return {string}
Function MapPath( Path )
    if Path = vbNullString then
        Path = DynamicInclude_CurrentFolder
    end if

    MapPath = Server.MapPath(Path)
End Function
%>