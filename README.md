# Json2ExcelMiddleware  

**Json2ExcelMiddleware** allows you to effortlessly convert your .NET controller's JSON responses into Excel files with minimal setup.  

## Installation  

Install the package via NuGet:  
```bash
Install-Package Json2ExcelMiddleware -Version 1.2.4
```

## How to Use  

1. **Configure Middleware**  
   Add the following line to the `Configure` method in your `Startup.cs` after `app.UseRouting()`:  
   ```csharp
   app.UseJson2Excel();
   ```
2. **Set Header**
   Include this header in your HTTP request:
   ```
   x-excel:1
   ```
That’s it! The middleware works with all HTTP methods (GET, POST, PUT, etc.).

## Example
### Before
A typical JSON response:
```json
[
  { "Name": "Alice", "Age": 25 },
  { "Name": "Bob", "Age": 30 }
]
```
### After
Downloadable Excel file with the same data!

## License
This project is licensed under [Apache License](https://github.com/stormaref/Json2ExcelMiddleware/blob/main/LICENSE)
