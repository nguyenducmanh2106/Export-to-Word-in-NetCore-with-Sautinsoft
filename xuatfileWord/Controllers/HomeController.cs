using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using SautinSoft.Document;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using xuatfileWord.Models;

namespace xuatfileWord.Controllers
{
    [Route("Home")]
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IHostingEnvironment _hostingEnvironment;
        public HomeController(ILogger<HomeController> logger,IHostingEnvironment hostingEnvironment)
        {
            _logger = logger;
            _hostingEnvironment = hostingEnvironment;
        }
        [HttpGet("Index")]
        public IActionResult Index()
        {
            
            return View();
        }

        [HttpGet("xuat-file")]
       public IActionResult FirstProcess()
        {
            string htmlPage = "Index";
            string rootPath = $"{_hostingEnvironment.ContentRootPath}";
            string path = $"{rootPath}/Views/Home/{htmlPage}.cshtml"; 
            byte[] outputData = null;
            if (System.IO.File.Exists(path))
            {
                //string outputFile = $"F:/DocMy.docx";//Tạo tệp có sẵn
                byte[] inputFile = System.IO.File.ReadAllBytes(path);
                
                using (MemoryStream memoryStream = new MemoryStream(inputFile))
                {
                    DocumentCore document = DocumentCore.Load(memoryStream,new HtmlLoadOptions());
                   using(MemoryStream ms=new MemoryStream())
                    {
                        document.Save(ms, new DocxSaveOptions());
                        outputData = ms.ToArray();
                    }
                    //if (outputData != null)
                    //{
                    //    //System.IO.File.WriteAllBytes(outputFile,outputData); //Ghi đề tệp outputData->outputFile
                    //}
                }      
            }
            return File(outputData, "application/msword", "THONG_KE_KHU_DIEM_CAP_XEP_HANG.docx");
            //return File(outputData, "application/force-download", "THONG_KE_KHU_DIEM_CAP_XEP_HANG.docx");

        }
       
    }
}
