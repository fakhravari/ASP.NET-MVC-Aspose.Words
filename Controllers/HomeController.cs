using System;
using System.Web.Mvc;
using WebApplication2.Models;

namespace WebApplication2.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            string[] Field = new string[] { "Shomareh", "Tarikh", "Peyvast", "HtmlMatn", "HtmlEmza", "HtmlRoonevesht", "HtmlParaf" };
            object[] data;

            data = new object[] {
                "Shomareh",
                "ShamsiDateNow",
                "",
                "PreMatnName",
                "getEmzaFormatHtmlFA",
                "Roonevesht",
                "ShowparaphsPrint"
            };

            string path = iAspose.BuildPrintLetter(Field, data, "~/Content/LT_HavooRush.docx", "~/PigisFileServer/Temp/");
            TempData["file"] = path;

            return View();
        }
    }
}