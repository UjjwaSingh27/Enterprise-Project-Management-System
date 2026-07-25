using Project_Manager.DAL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Project_Manager.Controllers
{
    public class ChartController : Controller
    {
        public ActionResult Index()
        {
            using (var con = OracleDbHelper.GetConnection())
            {
                con.Open();
            }

            ViewBag.Message = "Oracle Connected Successfully";
            return View();
        }
    }
}