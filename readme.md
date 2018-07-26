# VBScript Dynamic Includes

A simple utilitary library to provide __real__ dynamic includes in Classic ASP.

There are not many relevent new resources to be added in this project, so certainly it won't be updated frequenly.



## Parsing proccess

The inclusion is made reading and parsing the files, converting all plain-text
blocks to sequences of _Response.Write_ commands. Short ASP output tags (__&lt;=__) in  are expanded to _Response.Write_ commands.

During the previous steps, all ASP tags (__&lt;__ and __&gt;__) are removed from the code.  
At the end of this proccess, all plain text is conveted to VBScript output with valid code.

**_Warning:_ If a included file contains *class* declarations it will be executed on the global namespace - it's a simple fix for a VBScript limitation.**


*This library is not meant to be used in pure VBScript projects, because such complex parsing isn't even needed for plain __.vbs__ files.*



## Provided Include Methods

There are two include methods, inspired in PHP way of including files:
* __Include__(string __File__)  
Tries to include a file on current script. If no errors occur file is included, else just fails quietly.
* __Require__(string __File__)  
Tries to include a file on current script. If no errors occur file is included, else ends script execution with error message.




## Additional Methods

The following extra methods are also provided with the library, because they are internaly used by *Include* and *Require*:
* boolean __FileDoExist__( string __File__, FileSystemObject __System__)  
Checks if _File_ does exist in _System_. Useful to avoid extra FileSystemObject allocation.
* boolean __FileExists__( string __File__)  
Checks if File exist.
* FileSystemObject __FileSystem__( )  
Syntax sugar.
* string __ParseFile__( string __File__)  
Reads File as ASP code.
* string __ReadFile__( string __File__)  
Reads File as plain text.