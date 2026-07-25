using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Project_Manager.Models
{
    public class ProjectModel
    {
        public string projectName { get; set; }

        public string GroupName { get; set; }

        public string categoryName { get; set; }


        public List<TaskModel> tasks { get; set; }
    }
}