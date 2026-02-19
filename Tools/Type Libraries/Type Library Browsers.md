# Type Library Browsers

Type libraries contain the specification for one or more COM elements, including classes, interfaces, enumerations, and more. These files are stored in a standard binary format. A type library can be a stand-alone file with the .tlb file name extension, or it can be stored as a resource in an executable file, which can have an .ocx, .dll, or .exe file name extension. The type library viewers and conversion tools described following read this format to gain information about the COM elements in the library.
Before you can program an object in a particular programming language, you must be able to view its type library in that language. Doing this provides you with the proper syntax for the classes, interfaces, methods, properties, and events of the COM object.

https://msdn.microsoft.com/en-us/library/windows/desktop/ms680581(v=vs.85).aspx

Microsoft provides an OLE-COM Viewer, Oleview.exe, an application supplied with Visual C++ that displays the COM objects installed on your computer and the interfaces they support. You can use this object viewer to view type libraries. Visual Basic 6 has its own object browser.

# The making of a Type Library browser

What we need is a type library browser able to generate code in the proper syntax to be used with FreeBasic.

* Retrieving information from the registry
* Loading the type library
* Parsing the information
* Enumerating the constants
* Enumerating structures and constants
* Aliases and typedefs
* CoClasses
* Interfaces
* Methods and properties
* Parameters
