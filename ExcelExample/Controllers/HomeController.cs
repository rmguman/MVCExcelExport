using ClosedXML.Excel;
using ExcelExample.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ExcelExample.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            List<UserVM> model = new List<UserVM>();

            Random rnd = new Random();

            for (int i = 1; i < 11; i++)
            {
                UserVM user = new UserVM();
                user.ID = i;
                user.Name = "Name " + (rnd.Next(1, 100)).ToString();
                user.Surname = "Surname " + (rnd.Next(1, 100)).ToString();

                model.Add(user);
            }

            return View(model);
        }

        public FileResult DownloadExcel()
        {
            List<UserVM> model = new List<UserVM>();

            Random rnd = new Random();

            for (int i = 1; i < 11; i++)
            {
                UserVM user = new UserVM();
                user.ID = i;
                user.Name = "Name " + (rnd.Next(1, 100)).ToString();
                user.Surname = "Surname " + (rnd.Next(1, 100)).ToString();

                model.Add(user);
            }


            //4.
            DataTable dt = new DataTable("Grid");
            dt.Columns.AddRange(new DataColumn[3] { new DataColumn("ID"),
                                            new DataColumn("Name"),
                                            new DataColumn("Surname") });

            var customers = model;

            foreach (var customer in customers)
            {
                dt.Rows.Add(customer.ID, customer.Name, customer.Surname);
            }

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                  return  File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Grid.xlsx");
                }
            }



        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}