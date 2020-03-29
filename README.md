# 
In Order to create Dll and use it with Java, we need to have a VB6 environement with SP6 Update, witch can work in Windows 10.
    To install, follow instructions : (en) https://www.raymond.cc/blog/install-visual-basic-6-vb6-in-windows-7-without-microsoft-virtual-machine-for-java/
                                      (fr) http://www.cmdvb.fr/tutoriel-installation-de-visual-basic-6-sp5-sp6-sous-windows-10/
																			
Understand différence between ActiveX Dll and Standard Dll :
    ActiveX is Microsoft's name for COM compliant components. An ActiveX DLL is therefore one that can
    be registered using RegSvr32, and complies with all of the requires for COM.
    A general DLL, such as might be created with C or C++, may (or may not) be compliant with COM (for instance,
    most, if not all of the WINDOWS API DLLs ARE NOT COM-compliant) and as such are not registered, but are
    accessed through the API declarations and call. The potential problem is that a C or C++ DLL may not
    export the functions and procedures n a format that VB can use.
    COM executable programs and DLLs are libraries of classes. Client applications use COM objects by creating
    instances of classes provided by the COM .exe or .dll file. Clients call the properties, methods, and
    events provided by each COM object.
    In Visual Basic, the project templates you use to create a COM executable program or COM DLL are referred
    to as ActiveX EXE and ActiveX DLL, respectively.
    Visual Basic handles much of the complexity of creating COM .exe and .dll files, such as creating a
    type library and registering the component, automatically.
    In-process:
    An in-process component is implemented as a dynamic-link library (DLL). It runs in the same process
    space as its client application. This enables the most efficient communication between client and component,
    because you need only call the function to obtain the required functionality. Each client application
    that uses the component starts a new instance of the component.
    Based on above, any DLL created thru Visual Basic is termed as ActiveX DLL. If created in other languages,
    it is called DLL (Not ActiveX DLL).
    '=============
    It is not possible to create a non-ActiveX dll in VB alone, but some third party tools exist that claim to be
    able to do so. 
		'============= 
		Because AxtiveX Dll don't expose Methods, we have to found a way to transform an ActiveX Dll to a Standard Dll: See 
	
    
