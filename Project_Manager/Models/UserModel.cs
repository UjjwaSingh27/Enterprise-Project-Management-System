using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Project_Manager.Models
{
    public class UserModel
    {
        public int USER_ID { get; set; }
        public string USERNAME { get; set; }
        public string PASSWORD_HASH { get; set; }
        public int ROLE_ID { get; set; }
    }
}