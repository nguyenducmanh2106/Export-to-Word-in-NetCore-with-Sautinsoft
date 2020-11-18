using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using xuatfileWord.Models;
using xuatfileWord.Util;
namespace xuatfileWord.Controllers
{
    [Route("Home")]
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IHostingEnvironment _hostingEnvironment;
        public HomeController(ILogger<HomeController> logger, IHostingEnvironment hostingEnvironment)
        {
            _logger = logger;
            _hostingEnvironment = hostingEnvironment;
        }
        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }
        [HttpGet("CreateFileWord")]
        public async Task<IActionResult> CreateFileWord(string html="")
        {
            try
            {
              
                return File(HtmlToWord.HtmlToWordMethod(html), "application/force-download", "thongkevanban.doc");
            }
            catch (Exception ex)
            {
                return Json("");
            }
        }

    }
}
