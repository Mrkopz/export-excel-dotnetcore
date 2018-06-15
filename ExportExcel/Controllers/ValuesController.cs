using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.FileProviders;
using OfficeOpenXml;

namespace ExportExcel.Controllers
{
    [Route("api/[controller]")]
    public class ValuesController : Controller
    {
        private readonly IHostingEnvironment _hosting;

        public ValuesController(IHostingEnvironment hosting)
        {
            _hosting = hosting;
        }
        // GET api/values
        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/values/5
        [HttpGet("{id}")]
        public string Get(int id)
        {
            return "value";
        }

        // POST api/values
        [HttpPost]
        public void Post([FromBody]string value)
        {
        }

        // PUT api/values/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/values/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }

        [HttpGet, Route("export")]
        public IActionResult ExportExcel()
        {
            var webRoot = _hosting.WebRootPath;
            var fileName = @"test.xlsx";
            var url = string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, fileName);
            var file = new FileInfo(Path.Combine(webRoot, fileName));
            if (file.Exists)
            {
                file.Delete();
                file = new FileInfo(Path.Combine(webRoot, fileName));
            }

            using (var package = new ExcelPackage(file))
            {
                var worksheet = package.Workbook.Worksheets.Add("test");

                worksheet.Cells[1, 1].Value = "Id";
                worksheet.Cells[1, 2].Value = "Name";

                worksheet.Cells["A2"].Value = 1000;
                worksheet.Cells["B2"].Value = "AAA";

                worksheet.Cells["A3"].Value = 1001;
                worksheet.Cells["B3"].Value = "BBB";

                package.Save();
            }

            return DownloadFile(webRoot, fileName);
        }

        public FileResult DownloadFile(string filePath, string filename)
        {
            var provider = new PhysicalFileProvider(filePath);
            var fileInfo = provider.GetFileInfo(filename);
            var readStream = fileInfo.CreateReadStream();
            var mimeType = "application/vnd.ms-excel";

            //var test = new FileContentResult();
            return File(readStream, mimeType, filename);
        }
    }
}
