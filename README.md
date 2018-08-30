# apache-poi-os
ObjectScript part for apache-poi integration. Allows reading xlsx.

# Installation

1. Download [latest release](https://github.com/intersystems-ru/apache-poi-os/releases) and import it into any namespace.
2. Download [latest release](https://github.com/intersystems-ru/apache-poi/releases): poi archive and jar from [intersystems-ru/apache-poi repository](https://github.com/intersystems-ru/apache-poi/releases).
3. Extract archive and copy `poi.jar` into one directory, later referenced as `<DIR-WITH-JARS>`. Caché should have access to this directory. 
4. Execute: `set sc=$system.OBJ.UpdateConfigParam("isc.poi.Utils","DIR", ##class(%File).NormalizeDirectory("<DIR-WITH-JARS>"))`. Don't forget to check `sc` for errors.
5. Create Java Gateway: `Write $System.Status.GetErrorText(##class(isc.poi.Utils).createGateway(name, home, path, port))`, where:
   - `name` is a name of the Java Gateway, defaults to `POI`.
   - `home` path to Java 1.8 JRE. Defaults to `JAVA_HOME` environment variable.
   - `path` is a path to jars, should be set automatically.
   - `port` Java Gateway port. Defaults to `55556`.
6. Load jar into Caché: `Write $System.Status.GetErrorText(##class(isc.poi.Utils).updateJar())`
7. Check that everything works: `Write $System.Status.GetErrorText(##class(isc.poi.Utils).getBook(file, .book))`

# API 

Excel books are loaded into `isc.poi.Book` objects. Check it for API documenttion.
