# VBScript Dynamic Includes

A simple utilitary library to provide __real__ dynamic includes in Classic ASP.

There are not many relevent new resources to be added in this project, so certainly it won't be updated frequenly - except in case of errors or problems.



## How to use

Just use a classic server-side include to add *DynamicInclude.asp* on your project and call the include methods provided by this library to load files dynamically.

*Hint: obviously its faster to use standard includes when posssible than calling everything by this library.*



## Provided Include Methods

There are two include methods, inspired in PHP's way of including files:
* __Include__(string __File__)  
Tries to include a file on current script. If no errors occur file is included, else just fails quietly.
* __Require__(string __File__)  
Tries to include a file on current script. If no errors occur file is included, else ends script execution with an error message.



## Control Flags

This variables are provided to control the behavior of the including process:

* string __DynamicInclude_PreviousPath__  
Path of the last included file, used for recursive file importing.  
May be used along with *DynamicInclude_CurrentPath* to change how recursive imports will behave.
* string __DynamicInclude_CurrentPath__  
Path of current file to include, used for recursive file importing.  
May be used along with *DynamicInclude_PreviousPath* to change how recursive imports will behave.
* boolean __DynamicInclude_TrimHtml__  
If HTML text should be trimmed while parsing.
* boolean __DynamicInclude_TrimNewlines__  
If duplicated new lines hould be removed from parsed text.

More control flags may be added in the future.
Flags meant only for internal use are not listed for obvious reasons.



## Parsing proccess

The inclusion is made reading and parsing the files, converting all plain-text
blocks to sequences of _Response.Write_ commands. Short ASP output tags (__&lt;%=__) in  are expanded to _Response.Write_ commands.

During the those steps, all ASP tags (__&lt;%__ and __&gt;%__) are removed from the code.  
At the end of this proccess, all text is conveted to pain *VBScript* output with valid code.


**_Warning:_ Everything will be executed on the global namespace.A flag to control this behavior may be added in the future.**

**_Warning:_ Do not import/require files the have VBScript Class declarations inside control flow statements (if/else/case) or loops. Conditional class declaration isn't  supported by Classic ASP.**

*This library is not meant to be used in pure VBScript projects, because such complex parsing isn't even needed for plain __.vbs__ files.*



## Additional Methods

The following  methods are also provided with the library, because they are internaly used by *Include* and *Require*:
* __ExecuteFile__( string __File__, FileSystemObject __System__)  
Imports and executes File in the global namespace, dealing with paths for recursive calls.  
* string __ParseFile__( string __File__)  
Reads File as ASP code.
* string __ReadFile__( string __File__)  
Reads File as plain text.



## Deprecated methods

This extra methods were removed from main code file (DynamicInclude.asp) because were not really needed for the library functionality or even were bad for performance.

Those methods are:

* boolean __FileDoExist__( string __File__, FileSystemObject __System__)  
Checks if _File_ does exist in _System_. Useful to avoid extra FileSystemObject allocation.  
*Deprecated due to unnecessary increase of call stack*
* boolean __FileExists__( string __File__)  
Checks if File exist.  
*Deprecated because it was not used*
* string __FileName__( string __File__)  
Gets the name of the file.  
*Deprecated due to unnecessary increase of call stack*
* string __FilePath__( string __File__)  
Gets the path of a file.  
*Deprecated due to unnecessary increase of call stack*
* FileSystemObject __FileSystem__( )  
Syntax sugar.  
*Deprecated because it was completely unnecessary*
* string __MapPath__( string __File__)  
Gets the absolute equivalent of Path.  
*Deprecated due to unnecessary increase of call stack*