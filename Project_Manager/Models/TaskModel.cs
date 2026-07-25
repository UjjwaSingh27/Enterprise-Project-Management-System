using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Project_Manager.Models
{
    public class TaskModel
    {
        public int? id { get; set; }

        public int? projectId { get; set; }

        public int? parentId { get; set; }
        public string name { get; set; }
        public int? orderId { get; set; }

        public string title { get; set; }

        public string assignedToName { get; set; }   // stores username from UI
        public int? assignedTo { get; set; }         // stores USER_ID in DB
        public string attachmentPath { get; set; }
        public string attachmentName { get; set; }

        public int? attachedBy { get; set; }
        public string attachedByName { get; set; }
        public DateTime? attachedOn { get; set; }

        public DateTime? start { get; set; }
        public DateTime? end { get; set; }

        public double? percentComplete { get; set; }


        public bool summary { get; set; }

        public bool expanded { get; set; }
    }
}