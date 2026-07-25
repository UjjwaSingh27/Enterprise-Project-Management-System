using DocumentFormat.OpenXml.Spreadsheet;
using Oracle.ManagedDataAccess.Client;
using Project_Manager.DAL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
namespace Project_Manager.Controllers
{
    public class AccountController : Controller
    {
        // GET: Account
        public ActionResult Login()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Login(string username, string password)
        {
            using (var con = OracleDbHelper.GetConnection())
            {
                con.Open();

                string query = @"SELECT USER_ID, USERNAME, ROLE_ID
                     FROM USERS
                     WHERE UPPER(USERNAME) = UPPER(:username)
                     AND PASSWORD_HASH = :password
                     AND STATUS = 'ACTIVE'";

                OracleCommand cmd = new OracleCommand(query, con);
                cmd.BindByName = true;

                cmd.Parameters.Add("username", username.Trim());
                cmd.Parameters.Add("password", password.Trim());

                System.Diagnostics.Debug.WriteLine("Username: [" + username + "]");
                System.Diagnostics.Debug.WriteLine("Password: [" + password + "]");

                System.Diagnostics.Debug.WriteLine("Running login query...");

                OracleDataReader dr = cmd.ExecuteReader();

                if (dr.Read())
                {
                    Session["USER_ID"] = dr["USER_ID"].ToString();
                    Session["USERNAME"] = dr["USERNAME"].ToString();

                    string roleId = dr["ROLE_ID"].ToString();
                    Session["ROLE_ID"] = roleId;

                    if (roleId == "1")
                        Session["Role"] = "ADMIN";
                    else if (roleId == "2")
                        Session["Role"] = "MANAGER";
                    else if (roleId == "3")
                        Session["Role"] = "USER";
                    else
                        Session["Role"] = "VIEWER";

                    return RedirectToAction("Index", "Home");
                }
            }

            ViewBag.Error = "Invalid Username or Password";
            return View();
        }

        public ActionResult Logout()
        {
            Session.Clear();
            Session.Abandon();

            return RedirectToAction("Login", "Account");
        }
    }
}