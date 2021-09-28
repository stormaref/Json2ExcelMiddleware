using ClosedXML.Excel;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using System.Data;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;

namespace Json2ExcelMiddleware
{
    internal class Json2ExcelMiddleware
    {
        private readonly RequestDelegate _next;
        private string fileName = "result.xlsx";

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

            context.Response.OnStarting(() =>
            {
                context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                context.Response.Headers.Add("content-disposition", $"attachment;filename=\"{fileName}\"");
                context.Response.ContentLength = null;
                return Task.CompletedTask;
            });

            Stream originalBody = context.Response.Body;

            try
            {
                using (var ms = new MemoryStream())
                {
                    context.Response.Body = ms;

                    await _next(context);

                    ms.Position = 0;
                    string responseBody = new StreamReader(ms).ReadToEnd();
                    DataTable dt = GetDataTable(responseBody);
                    using (var workbook = new XLWorkbook())
                    {
                        workbook.Worksheets.Add(dt, "Result");
                        using (var memoryStream = new MemoryStream())
                        {
                            workbook.SaveAs(memoryStream);
                            memoryStream.Position = 0;
                            context.Response.ContentLength = ms.Length;
                            await memoryStream.CopyToAsync(originalBody);
                            memoryStream.Close();
                        }
                    }
                }
            }
            finally
            {
                context.Response.Body = originalBody;
            }
        }

        private DataTable GetDataTable(string responseBody)
        {
            var jsonString = responseBody.StartsWith("{") ? $"[{responseBody}]" : responseBody;
            return (DataTable)JsonSerializer.Deserialize(jsonString, typeof(DataTable));
        }
    }

    public static class Json2ExcelMiddlewareExtensions
    {
        public static IApplicationBuilder UseJson2Excel(this IApplicationBuilder builder)
        {
            return builder.UseMiddleware<Json2ExcelMiddleware>();
        }
    }
}
