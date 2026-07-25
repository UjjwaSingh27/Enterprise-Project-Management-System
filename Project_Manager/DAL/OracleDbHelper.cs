using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using Oracle.ManagedDataAccess.Client;
using System.Configuration;

namespace Project_Manager.DAL
{
    public class OracleDbHelper
    {
        public static OracleConnection GetConnection()
        {
            string connStr =
                ConfigurationManager.ConnectionStrings["OracleDb"].ConnectionString;

            return new OracleConnection(connStr);
        }
    }
}

    