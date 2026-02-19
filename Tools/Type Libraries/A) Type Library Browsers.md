# Type Library Browsers

Type libraries contain the specification for one or more COM elements, including classes, interfaces, enumerations, and more. These files are stored in a standard binary format. A type library can be a stand-alone file with the `.tlb` file name extension, or it can be stored as a resource in an executable file, which can have an `.ocx`, `.dll`, or `.exe` file name extension. The type library viewers and conversion tools described following read this format to gain information about the COM elements in the library.

Before you can program an object in a particular programming language, you must be able to view its type library in that language. Doing this provides you with the proper syntax for the classes, interfaces, methods, properties, and events of the COM object.

[Type Library Viewers and Conversion Tools](https://msdn.microsoft.com/en-us/library/windows/desktop/ms680581(v=vs.85).aspx)

Microsoft provides an OLE-COM Viewer, `Oleview.exe`, an application supplied with Visual C++ that displays the COM objects installed on your computer and the interfaces they support. You can use this object viewer to view type libraries. Visual Basic 6 has its own object browser.

# The making of a Type Library browser

Creating a Type Library (TypeLib) browser involves building an application that reads binary metadata (.tlb, .dll, .ocx) and parses the COM (Component Object Model) interface information—such as CoClasses, interfaces, enums, and methods—into a readable format. A customized browser allows developers to understand component object models, generate wrappers, and explore API documentation.

#### Core components for making a TypeLib Browser

The foundational Windows APIs used to load and parse `.tlb` files, i.e. the **LoadTypeLib** function and the **ITypeLib** and **ITypeInfo** interfaces.

A user interface to display the hierarchy, such as a **TreeView** (for library -> coclasses -> interfaces -> methods -> parameters) and information such guids, constants, enumerations and structures.

The following articles describe with code how to build a `TypeLib Browser` with `Freebasic`.

* [Retrieving information from the registry](https://github.com/JoseRoca/AfxNova/blob/main/Tools/Type%20Libraries/B\)%20Retrieving%20the%20type%20libraries.md)
* [Loading the type library](https://github.com/JoseRoca/AfxNova/blob/main/Tools/Type%20Libraries/C\)%20Loading%20the%20type%20library.md)
* [Parsing the information](https://github.com/JoseRoca/AfxNova/blob/main/Tools/Type%20Libraries/D\)%20Parsing%20the%20information.md)
* [Enumerating the constants](https://github.com/JoseRoca/AfxNova/blob/main/Tools/Type%20Libraries/E\)%20Enumerating%20the%20constants.md)
* [Enumerating structures and constants](https://github.com/JoseRoca/AfxNova/blob/main/Tools/Type%20Libraries/F\)%20Enumerating%20structures%20and%20unions.md)
* [Aliases and typedefs](https://github.com/JoseRoca/AfxNova/blob/main/Tools/Type%20Libraries/G\)%20Aliases%20and%20typedefs.md)
* [Enumerations and modules](https://github.com/JoseRoca/AfxNova/blob/main/Tools/Type%20Libraries/H\)%20Enumerations%20and%20modules.md)
* [CoClasses](https://github.com/JoseRoca/AfxNova/blob/main/Tools/Type%20Libraries/I\)%20Retrieving%20the%20CoClasses.md)
* [Interfaces](https://github.com/JoseRoca/AfxNova/blob/main/Tools/Type%20Libraries/J\)%20Retireving%20the%20Interfaces.md)
* [Methods and properties](https://github.com/JoseRoca/AfxNova/blob/main/Tools/Type%20Libraries/K\)%20Retrieving%20methods%20and%20properties.md)
* [Parameters](https://github.com/JoseRoca/AfxNova/blob/main/Tools/Type%20Libraries/L\)%20Retrieving%20the%20parameters.md)

### Step-by-step development approach

*  Step 1: Load the Type Library
+ Use the **LoadTypeLib** API function to load the `.tlb` file into memory. If the file is embedded within a DLL or EXE, use **LoadTypeLibEx** to extract it from the resource section.
* Step 2: Enumerate type information
 * Iterate through the **ITypeLib** interface to get the count of type descriptions. Use **GetTypeInfo** to retrieve each **ITypeInfo** interface. This provides access to the details of each class, interface, and enum.
* Step 3: Parse the metadata
For each **ITypeInfo**, extract the following:
 * CoClasses: Component classes that can be instantiated.
 * Interfaces: The methods, properties, and functions (VTable entries).
 * GUIDs: The IID (Interface ID) and CLSID (Class ID).
* Step 4: Display/Generate Output
 * View: Map the parsed information into a TreeView.
 * Generate code: Convert the parsed data into readable programming language syntax (e.g., `FreeBasic`).
