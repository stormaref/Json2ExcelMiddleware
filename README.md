# Json2ExcelMiddleware

This middleware convert your .net controller json result to excel file without adding anything to any controller

how to add package:

using pm>

```
Install-Package Json2ExcelMiddleware -Version 1.2.0
```

how to use: 

1) first add this line of code to your startup file configure section:

```c#
app.UseJson2Excel();
```
* after app.UseRouting()

2) add this header to your http request:

```
x-excel:1
```

that's it!

it works for every kind of http request (get,post,put,etc.)
