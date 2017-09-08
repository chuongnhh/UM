using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using UM.Models;

namespace UM.Controllers
{
    public class EmployeeController : Controller
    {
        public static List<Employee> lstEmployee;
        // GET: Emloyee
        public ActionResult Index()
        {
            lstEmployee = new List<Employee>();
            return View();
        }

        public JsonResult GetData()
        {
            using (StreamReader r = new StreamReader(Server.MapPath(@"/scripts/data.js")))
            {
                string json = r.ReadToEnd();
                List<Employee> items = JsonConvert.DeserializeObject<List<Employee>>(json);
                lstEmployee.AddRange(items);
                //db.SaveChanges();
            }
            return Json(new { data = lstEmployee }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult Search(string txtsearch)
        {
            return Json(new
            {
                data = lstEmployee
                .Where(x => x.Name.Contains(txtsearch) ||
                x.PhoneNumber.Contains(txtsearch) ||
                txtsearch.Contains(x.Name) ||
                txtsearch.Contains(x.PhoneNumber)).ToList()
            },
                JsonRequestBehavior.AllowGet);
        }
    }
}