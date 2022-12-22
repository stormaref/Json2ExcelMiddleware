using ClosedXML.Excel;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json;
using System.Data;
using System.IO;
using System.Threading.Tasks;

namespace Json2ExcelMiddleware;

internal class Json2ExcelMiddleware
{
    private readonly RequestDelegate _next;
    private const string FileName = "result.xlsx";
    private const string SheetName = "Result";

    public Json2ExcelMiddleware(RequestDelegate next)
    {
        _next = next;
    }

    public async Task Invoke(HttpContext context)
    {
        var isExcel = context.Request.Headers.Keys.Contains("x-excel");
        if (!isExcel)
        {
            await _next(context);
            return;
        }

        context.Response.OnStarting(() => OnStarting(context));
        await WriteToBody(context);
    }

    Task OnStarting(HttpContext context)
    {
        context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        context.Response.Headers.Add("content-disposition", $"attachment;filename=\"{FileName}\"");
        context.Response.ContentLength = null;
        return Task.CompletedTask;
    }

    private async Task WriteToBody(HttpContext context)
    {
        var originalBody = context.Response.Body;
        try
        {
            using var stream = new MemoryStream();
            context.Response.Body = stream;
            await _next(context);
            stream.Position = 0;
            var jsonResponse = await new StreamReader(stream).ReadToEndAsync();
            var workbook = CreateWorkbook(jsonResponse);
            await WriteWorkbookToStream(context, workbook, stream.Length, originalBody);
        }
        finally
        {
            context.Response.Body = originalBody;
        }
    }

    private static XLWorkbook CreateWorkbook(string json)
    {
        var dataTable = GetDataTable(json);
        using var workbook = new XLWorkbook();
        workbook.Worksheets.Add(dataTable, SheetName);
        return workbook;
    }

    private static async Task WriteWorkbookToStream(HttpContext context, IXLWorkbook workbook, long length,
        Stream originalBody)
    {
        using var memoryStream = new MemoryStream();
        workbook.SaveAs(memoryStream);
        memoryStream.Position = 0;
        context.Response.ContentLength = length;
        await memoryStream.CopyToAsync(originalBody);
        memoryStream.Close();
    }

    private static DataTable GetDataTable(string responseBody)
    {
        var jsonString = responseBody.StartsWith("{") ? $"[{responseBody}]" : responseBody;
        return JsonConvert.DeserializeObject<DataTable>(jsonString);
    }
}

public static class Json2ExcelMiddlewareExtensions
{
    public static IApplicationBuilder UseJson2Excel(this IApplicationBuilder builder)
    {
        return builder.UseMiddleware<Json2ExcelMiddleware>();
    }
}