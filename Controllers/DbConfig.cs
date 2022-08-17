using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRUD.MVC.WebApp.Utility
{
    public class DbConfig
    {
        public static string ConnectionString
        {
            get
            {
                return ConfigurationManager.ConnectionStrings["sqlConnection"].ConnectionString;
            }
        }
    }
}
