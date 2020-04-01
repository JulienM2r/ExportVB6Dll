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
#
Because ActiveX Dll does not expose Methods, we have to find a way to transform an ActiveX Dll into a Standard Dll.
In fact, when you Make an ActiveX Dll, VB send some instructions to the Linker (LINK.EXE), the solution is in the Hack of this transfer of instructions: replace the linker with another VB6.exe project.
	Step 1 : Download or Copy the "ConsoleApplication1" folder and open the "ConsoleAppDll.vbp" file in VB6 IDE. (Never forget to Launch VB6 on admin)
		-> The Module DllFunction.bas contains the Functions we want exposed, if you build it as DLL file then you don't Export inside methods and you don't have a standard Dll. 
	Step 2 : Download or Copy the "NewLinker" folder and open the "NewLinker.vbp" file in VB6 IDE. (Never forget to Launch VB6 on admin)
		-> Generate a new .EXE file with de project, this one change the vb instruction and add descriptor in the Dll for .DEF and .REF informations. Then locate the LINK.EXE file on your VB instalation.(In mine, there is in : "C:\Program Files (x86)\Microsoft Visual Studio\VB98", and rename this "linklnk.EXE". Then Rename your NewLinker Project Exe "LINK.EXE" and copy it in the same Folder than "linklnk.EXE". 
		-> Open your favorite Text Editor and Paste This : 
					-----------------------------------------------------
						NAME ConsoleAppDll
						LIBRARY TestDllBL
						DESCRIPTION "Test d'une dll standard"
						EXPORTS DllMain @1
							Fn_calc @2
							Testmethod @3
							GetData @4
							FunctionCalled @5
					-----------------------------------------------------
		-> Save it as "ConsoleAppDll.DEF" (The name does have the same Name that de the .dll File)
		The Name and Library Tag are mandatory but you can choose free's Name
		Exports's part does contains the DllMain part who pointed on the Main of the Dll, after you have to enumerate all the 			functions (does have Public Access) you want to expose, and give them all an ordinal number(Incremental)(the @1 is reserved		  for the Main one)
		-> then Make Dll with your ConsoleAppDll.vbp 
		
	-> You can now use your Dll as a standard Dll
		
		
