/// Utility methods
Class isc.poi.Utils
{

/// Main class to import
Parameter CLASS = "isc.poi.Main";

/// Gateway name to create/use
Parameter GATEWAY = "POI";

/// Path to jar dir. Set with:
/// set sc=$system.OBJ.UpdateConfigParam("isc.poi.Utils","DIR","C:\InterSystems\poi\")
Parameter DIR As CONFIGVALUE;

/// Library files to load
/// poi.jar should be first
Parameter LIBS = {$lb("poi.jar", "poi-3.17.jar","poi-ooxml-3.17.jar", "poi-ooxml-schemas-3.17.jar", $lb("ooxml-lib", "xmlbeans-2.6.0.jar"), $lb("lib","commons-collections4-4.1.jar"))};

/// Test class in poi.jar
Parameter CLASSTEST = "isc.poi.Test";

/// DEV tools.
/// Path to built JAR. Set with:
/// set sc=$system.OBJ.UpdateConfigParam("isc.poi.Utils","BUILDCLASSPATH","D:\Cache\POI\java\out\artifacts\java_jar\java.jar")
Parameter BUILDCLASSPATH As CONFIGVALUE;

/// Create JGW. Java home must point to 1.8 jre.
/// Write $System.Status.GetErrorText(##class(isc.poi.Utils).createGateway())
ClassMethod createGateway(gatewayName = {..#GATEWAY}, javaHome = {$SYSTEM.Util.GetEnviron("JAVA_HOME")}, path As %String = {..getPath()}, port As %Integer = 55556)
{
	set sys = ##class(%Net.Remote.ObjectGateway).%New()
	set sys.Name = gatewayName
	set sys.Type = 1
	set sys.JavaHome = javaHome
	set sys.ClassPath = path
	set sys.Port = port
	quit sys.%Save()
}

/// w ##class(isc.poi.Utils).getPath()
ClassMethod getPath()
{
	set separator = $select($$$isWINDOWS:"\",1:"/")
	set path = ""
	for i=1:1:$ll(..#LIBS) {
		set lib = $lg(..#LIBS, i)
		set:$listValid(lib) lib = $lts(lib, separator)
		set path = path _ $lb(..#DIR _ lib)
		
	}
	quit $lts(path)
}

/// Load Jar from path.
/// Write $System.Status.GetErrorText(##class(isc.poi.Utils).updateJar())
ClassMethod updateJar(gatewayName = {..#GATEWAY}, path As %String = {..getPath()})
{
	#Dim sc As %Status = $$$OK
	set sc = ##class(%Net.Remote.Service).StopGateway(gatewayName)
	
	set result = 0
	/*set:$system["ed-pc" sc = ##class(%File).CopyFile(..#BUILDCLASSPATH,..#DIR _ $lg(..#LIBS), $$$YES, .result)
	if result=-32 {
		hang 1
		set sc = ##class(%File).CopyFile(..#BUILDCLASSPATH,..#DIR _ $lg(..#LIBS), $$$YES, .result)
	}
	
	quit:sc'=1 $$$ERROR($$$GeneralError, $$$FormatText("File copy failed with error: %1", result))*/
	
	#dim gateway As %Net.Remote.Gateway
	set gateway = ..connect(gatewayName, path, .sc)
	quit:$$$ISERR(sc) sc
	
	set sc = gateway.%Import(..#CLASS)
	//set:$system["ed-pc" sc = gateway.%Import(..#CLASSTEST)
	quit:$$$ISERR(sc) sc
	set:'##class(%Dictionary.CompiledClass).%ExistsId(..#CLASS) sc = $$$ERROR($$$GeneralError, $$$FormatText("Class '%1' does not exist",..#CLASS))
	
	quit sc
}

/// Get JGW object
ClassMethod connect(gatewayName As %String = {..#GATEWAY}, path As %String = {..getPath()}, Output sc As %Status) As %Net.Remote.Gateway
{
	set gateway = ""
	set sc = ##class(%Net.Remote.Service).OpenGateway(gatewayName, .gatewayConfig)
	quit:$$$ISERR(sc) gateway
	set sc = ##class(%Net.Remote.Service).ConnectGateway(gatewayConfig, .gateway, path, $$$YES)
	quit gateway
}

/// Load xlsx
/// Write $System.Status.GetErrorText(##class(isc.poi.Utils).getBook("D:\Cache\POI\Книга1.xlsx",.b))
ClassMethod getBook(file As %String, ByRef book As isc.poi.Book, debug As %Boolean = {$$$NO}) As %Status
{
    #dim gateway as %Net.Remote.Gateway
    #dim exception as %Exception.AbstractException

    set sc = $$$OK
    try {
        set gateway = ..connect()
        set start = $zh
        
        set list = ##class(isc.poi.Main).getBook(gateway, file)
        
        set end1 = $zh
        set book = ##class(isc.poi.Book).%New(list) 
		set end2 = $zh
		
		if debug {
			write !,"JGW: ",end1-start
			write !,"Instantiation: ",end2-end1
			
			//set exec = "set list = " _result
			//XECUTE exec
			//zwrite list
	        //zwrite result
	        //break
		}

        set sc = gateway.%Disconnect()

    } catch ex {
	    break:debug
        set sc = $$$ADDSC(ex.AsStatus(), $g(%objlasterror))
    }

    quit sc
}

}

