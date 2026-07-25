using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Project_Manager.Models
{
    public class DependencyModel
    {
        public int id { get; set; }
        public int predecessorId { get; set; }
        public int successorId { get; set; }
        public int type { get; set; }

    }
}