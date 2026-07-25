using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office.Word;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Ajax.Utilities;
using Oracle.ManagedDataAccess.Client;
using Project_Manager.DAL;
using Project_Manager.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Project_Manager.Controllers
{
    public class TaskController : Controller
    {

        private bool CanEdit()
        {
            var role = Session["Role"]?.ToString();

            return role == "ADMIN" || role == "MANAGER";
        }

        private void AddAuditLog(
            OracleConnection con,
            int taskId,
            string actionType,
            string oldValue,
            string newValue)
        {
            string query = @"INSERT INTO AUDIT_LOGS
                (LOG_ID, USER_ID, TASK_ID, ACTION_TYPE,
                 OLD_VALUE, NEW_VALUE, ACTION_TIME)
                VALUES
                (AUDIT_SEQ.NEXTVAL, :userId, :taskId, :actionType,
                 :oldValue, :newValue, SYSDATE)";

            OracleCommand cmd = new OracleCommand(query, con);

            cmd.Parameters.Add(new OracleParameter(
                "userId",
                Convert.ToInt32(Session["USER_ID"])
            ));

            cmd.Parameters.Add(new OracleParameter("taskId", taskId));
            cmd.Parameters.Add(new OracleParameter("actionType", actionType));
            cmd.Parameters.Add(new OracleParameter("oldValue", oldValue));
            cmd.Parameters.Add(new OracleParameter("newValue", newValue));

            cmd.ExecuteNonQuery();
        }

        public JsonResult GetProjects()
        {
            var projects = new List<object>();   // MOVE HERE

            using (var con = OracleDbHelper.GetConnection())
            {
                con.Open();

                OracleCommand userCmd = new OracleCommand("SELECT USER FROM DUAL", con);
                var currentUser = userCmd.ExecuteScalar();

                System.Diagnostics.Debug.WriteLine("CURRENT USER: " + currentUser);

                OracleCommand countCmd = new OracleCommand("SELECT COUNT(*) FROM PROJECTS", con);
                var count = countCmd.ExecuteScalar();

                System.Diagnostics.Debug.WriteLine("PROJECT COUNT: " + count);

                string query = "SELECT PROJECT_ID, PROJECT_NAME, PROJECT_GROUP, PROJECT_CATEGORY\r\nFROM PROJECTS\r\nORDER BY PROJECT_GROUP, PROJECT_CATEGORY, PROJECT_NAME";

                using (OracleCommand cmd = new OracleCommand(query, con))
                {
                    using (OracleDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            try
                            {
                                int id = Convert.ToInt32(dr["PROJECT_ID"]);
                                string name = dr["PROJECT_NAME"].ToString();

                                System.Diagnostics.Debug.WriteLine("Loaded: " + id + " - " + name);

                                projects.Add(new
                                {
                                    id = dr["PROJECT_ID"],
                                    name = dr["PROJECT_NAME"].ToString(),
                                    group = dr["PROJECT_GROUP"].ToString(),
                                    category = dr["PROJECT_CATEGORY"] == DBNull.Value
                                        ? "Uncategorized"
                                        : dr["PROJECT_CATEGORY"].ToString()
                                });
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine("Error: " + ex.Message);
                            }
                        }
                    }
                }
            }

            System.Diagnostics.Debug.WriteLine("Total: " + projects.Count);

            return Json(projects, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetGroups()
        {


            var groups = new List<object>();

            using (var con = OracleDbHelper.GetConnection())
            {
                con.Open();

                string query = @"
                    SELECT GROUP_ID, GROUP_NAME
                    FROM PROJECT_GROUPS
                    ORDER BY GROUP_NAME";

                OracleCommand cmd = new OracleCommand(query, con);
                OracleDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    groups.Add(new
                    {
                        id = dr["GROUP_ID"],
                        name = dr["GROUP_NAME"].ToString()
                    });
                }
            }

            return Json(groups, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetTasks()
        {
            int projectId = Convert.ToInt32(Request["projectId"] ?? "1");

            var tasks = new List<object>();

            using (var con = OracleDbHelper.GetConnection())
            {
                con.Open();

                string query = @"
                    SELECT 
                        T.TASK_ID,
                        T.PROJECT_ID,
                        T.PARENT_TASK_ID,
                        T.TASK_NAME,
                        T.ASSIGNED_TO,
                        U.USERNAME AS ASSIGNED_NAME,
                        T.ATTACHMENT_PATH,
                        T.ATTACHMENT_NAME,
                        T.ATTACHED_BY,
                        T.ATTACHED_ON,
                        AU.USERNAME AS ATTACHED_BY_NAME,
                        T.PLANNED_START,
                        T.PLANNED_END,
                        T.ACTUAL_START,
                        T.ACTUAL_END,
                        T.PROGRESS
                    FROM TASKS T
                    LEFT JOIN USERS U ON T.ASSIGNED_TO = U.USER_ID
                    LEFT JOIN USERS AU ON T.ATTACHED_BY = AU.USER_ID
                    WHERE T.PROJECT_ID = :projectId
                    ORDER BY NVL(T.PARENT_TASK_ID, T.TASK_ID), T.TASK_ID";

                OracleCommand cmd = new OracleCommand(query, con);
                cmd.BindByName = true;

                cmd.Parameters.Add(new OracleParameter("projectId", projectId));

                System.Diagnostics.Debug.WriteLine("Received Project ID: " + projectId);
                OracleDataReader dr = cmd.ExecuteReader();

                int order = 0;

                while (dr.Read())
                {
                    DateTime startDate = Convert.ToDateTime(dr["PLANNED_START"]);
                    DateTime endDate = Convert.ToDateTime(dr["PLANNED_END"]);

                    DateTime? actualEnd = dr["ACTUAL_END"] == DBNull.Value
                        ? (DateTime?)null
                        : Convert.ToDateTime(dr["ACTUAL_END"]);

                    int variance = 0;

                    if (actualEnd.HasValue)
                    {
                        variance = (actualEnd.Value - endDate).Days;
                    }


                    int progress = Convert.ToInt32(dr["PROGRESS"]);

                    string status;



                    if (progress == 100)
                    {
                        status = "COMPLETED";
                    }
                    else if (DateTime.Now > endDate)
                    {
                        status = "DELAYED";
                    }
                    else if (progress > 0)
                    {
                        status = "IN_PROGRESS";
                    }
                    else
                    {
                        status = "PENDING";
                    }
                    System.Diagnostics.Debug.WriteLine(
                        "Task: " + dr["TASK_NAME"] +
                        " End: " + endDate +
                        " Progress: " + progress +
                        " Status: " + status
                    );

                    tasks.Add(new
                    {
                        id = Convert.ToInt32(dr["TASK_ID"]),

                        projectId = Convert.ToInt32(dr["PROJECT_ID"]),

                        parentId = dr["PARENT_TASK_ID"] == DBNull.Value
                            ? (int?)null
                            : Convert.ToInt32(dr["PARENT_TASK_ID"]),
                        orderId = order++,
                        title = dr["TASK_NAME"].ToString(),

                        attachmentPath = dr["ATTACHMENT_PATH"]?.ToString(),
                        attachmentName = dr["ATTACHMENT_NAME"]?.ToString(),

                        assignedToName = dr["ASSIGNED_NAME"] == DBNull.Value
                            ? ""
                            : dr["ASSIGNED_NAME"].ToString(),
                        assignedTo = dr["ASSIGNED_TO"] == DBNull.Value
                            ? (int?)null
                            : Convert.ToInt32(dr["ASSIGNED_TO"]),

                        attachedBy = dr["ATTACHED_BY"] == DBNull.Value
                            ? (int?)null
                            : Convert.ToInt32(dr["ATTACHED_BY"]),

                        attachedByName = dr["ATTACHED_BY_NAME"]?.ToString(),

                        attachedOn = dr["ATTACHED_ON"] == DBNull.Value
                            ? (DateTime?)null
                            : Convert.ToDateTime(dr["ATTACHED_ON"]),

                        start = Convert.ToDateTime(dr["PLANNED_START"]),
                        end = Convert.ToDateTime(dr["PLANNED_END"]),

                        actualStart = dr["ACTUAL_START"] == DBNull.Value
                            ? (DateTime?)null
                            : Convert.ToDateTime(dr["ACTUAL_START"]),

                        actualEnd = dr["ACTUAL_END"] == DBNull.Value
                            ? (DateTime?)null
                            : Convert.ToDateTime(dr["ACTUAL_END"]),

                        variance = variance,


                        status = status,
                        percentComplete = Convert.ToInt32(dr["PROGRESS"]) / 100.0

                    });
                }
            }

            return Json(tasks, JsonRequestBehavior.AllowGet);
        }


        public JsonResult GetDependencies(int projectId)
        {
            var dependencies = new List<object>();

            using (var con = OracleDbHelper.GetConnection())
            {
                con.Open();

                string query = @"SELECT * FROM TASK_DEPENDENCIES
                                WHERE PROJECT_ID = :projectId";

                OracleCommand cmd = new OracleCommand(query, con);

                cmd.Parameters.Add(new OracleParameter("projectId", projectId));

                OracleDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    dependencies.Add(new
                    {
                        id = Convert.ToInt32(dr["DEPENDENCY_ID"]),
                        predecessorId = Convert.ToInt32(dr["PREDECESSOR_TASK_ID"]),
                        successorId = Convert.ToInt32(dr["SUCCESSOR_TASK_ID"]),
                        type = Convert.ToInt32(dr["DEPENDENCY_TYPE"])
                    });
                }
            }

            return Json(dependencies, JsonRequestBehavior.AllowGet);
        }

        public ActionResult Planner()
        {
            return View();
        }

        public JsonResult GetDashboardStats()
        {
            int projectId = Convert.ToInt32(Request["projectId"] ?? "1");

            using (var con = OracleDbHelper.GetConnection())
            {
                con.Open();

                string query = @"   
                    SELECT
                        COUNT(*) AS TOTAL,

                        SUM(CASE 
                            WHEN PROGRESS = 100 THEN 1 
                            ELSE 0 
                        END) AS COMPLETED,

                        SUM(CASE 
                            WHEN PROGRESS > 0 
                                 AND PROGRESS < 100 
                                 AND PLANNED_END >= SYSDATE
                            THEN 1 
                            ELSE 0 
                        END) AS IN_PROGRESS,

                        SUM(CASE 
                            WHEN PROGRESS = 0 
                                 AND PLANNED_END >= SYSDATE
                            THEN 1 
                            ELSE 0 
                        END) AS PENDING,

                        SUM(CASE 
                            WHEN PROGRESS < 100 
                                 AND PLANNED_END < SYSDATE
                            THEN 1 
                            ELSE 0 
                        END) AS DELAYED,

                        ROUND(AVG(PROGRESS), 2) AS AVG_PROGRESS
                    FROM TASKS
                    WHERE PROJECT_ID = :projectId";

                OracleCommand cmd = new OracleCommand(query, con);

                cmd.Parameters.Add(new OracleParameter("projectId", projectId));

                System.Diagnostics.Debug.WriteLine("Dashboard Project: " + projectId);

                OracleDataReader dr = cmd.ExecuteReader();

                if (dr.Read())
                {
                    return Json(new
                    {
                        total = dr["TOTAL"],
                        completed = dr["COMPLETED"],
                        inProgress = dr["IN_PROGRESS"],
                        pending = dr["PENDING"],
                        delayed = dr["DELAYED"],
                        completion = dr["AVG_PROGRESS"]
                    }, JsonRequestBehavior.AllowGet);
                }
            }

            return Json(null, JsonRequestBehavior.AllowGet);
        }

        public ActionResult ExportToExcel()
        {
            using (var con = OracleDbHelper.GetConnection())
            {
                con.Open();

                string query = @"
                    SELECT
                        T.TASK_NAME,
                        U.USERNAME AS ASSIGNED_NAME,
                        T.PLANNED_START,
                        T.PLANNED_END,
                        T.ACTUAL_START,
                        T.ACTUAL_END,
                        T.STATUS,
                        T.PROGRESS
                    FROM TASKS T
                    LEFT JOIN USERS U
                        ON T.ASSIGNED_TO = U.USER_ID";

                OracleCommand cmd = new OracleCommand(query, con);
                OracleDataReader dr = cmd.ExecuteReader();

                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("Tasks");

                    // Headers
                    ws.Cell(1, 1).Value = "Task Name";
                    ws.Cell(1, 2).Value = "Assigned To";
                    ws.Cell(1, 3).Value = "Planned Start";
                    ws.Cell(1, 4).Value = "Planned End";
                    ws.Cell(1, 5).Value = "Actual Start";
                    ws.Cell(1, 6).Value = "Actual End";
                    ws.Cell(1, 7).Value = "Status";
                    ws.Cell(1, 8).Value = "Progress (%)";

                    int row = 2;

                    while (dr.Read())
                    {
                        ws.Cell(row, 1).Value = dr["TASK_NAME"].ToString();
                        ws.Cell(row, 2).Value = dr["ASSIGNED_NAME"].ToString();
                        ws.Cell(row, 3).Value = dr["PLANNED_START"].ToString();
                        ws.Cell(row, 4).Value = dr["PLANNED_END"].ToString();
                        ws.Cell(row, 5).Value = dr["ACTUAL_START"].ToString();
                        ws.Cell(row, 6).Value = dr["ACTUAL_END"].ToString();
                        ws.Cell(row, 7).Value = dr["STATUS"].ToString();
                        ws.Cell(row, 8).Value = dr["PROGRESS"].ToString();

                        row++;
                    }

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        stream.Position = 0;

                        return File(
                            stream.ToArray(),
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            "ProjectTasks.xlsx"
                        );
                    }
                }
            }
        }

        public ActionResult ExportToPdf()
        {
            using (var con = OracleDbHelper.GetConnection())
            {
                con.Open();

                string query = @"
                    SELECT
                        T.TASK_NAME,
                        U.USERNAME AS ASSIGNED_NAME,
                        T.PLANNED_START,
                        T.PLANNED_END,
                        T.ACTUAL_START,
                        T.ACTUAL_END,
                        T.STATUS,
                        T.PROGRESS
                    FROM TASKS T
                    LEFT JOIN USERS U
                        ON T.ASSIGNED_TO = U.USER_ID";

                OracleCommand cmd = new OracleCommand(query, con);
                OracleDataReader dr = cmd.ExecuteReader();

                using (MemoryStream ms = new MemoryStream())
                {
                    Document doc = new Document(PageSize.A4.Rotate());
                    PdfWriter.GetInstance(doc, ms);
                    doc.Open();

                    doc.Add(new Paragraph("Project Tasks Report"));
                    doc.Add(new Paragraph(" "));

                    PdfPTable table = new PdfPTable(8);

                    table.AddCell("Task Name");
                    table.AddCell("Assigned To");
                    table.AddCell("Planned Start");
                    table.AddCell("Planned End");
                    table.AddCell("Actual Start");
                    table.AddCell("Actual End");
                    table.AddCell("Status");
                    table.AddCell("Progress");

                    while (dr.Read())
                    {
                        table.AddCell(dr["TASK_NAME"].ToString());
                        table.AddCell(dr["ASSIGNED_NAME"].ToString());
                        table.AddCell(dr["PLANNED_START"].ToString());
                        table.AddCell(dr["PLANNED_END"].ToString());
                        table.AddCell(dr["ACTUAL_START"].ToString());
                        table.AddCell(dr["ACTUAL_END"].ToString());
                        table.AddCell(dr["STATUS"].ToString());
                        table.AddCell(dr["PROGRESS"].ToString() + "%");
                    }

                    doc.Add(table);
                    doc.Close();

                    return File(ms.ToArray(), "application/pdf", "ProjectTasks.pdf");
                }
            }
        }

        public JsonResult GetProjectDocuments(int projectId)
        {
            var docs = new List<object>();

            using (var con = OracleDbHelper.GetConnection())
            {
                con.Open();

                string query = @"
                    SELECT
                        D.FILE_NAME,
                        D.FILE_PATH,
                        T.TASK_NAME,
                        P.PROJECT_NAME,
                        P.PROJECT_GROUP,
                        P.PROJECT_CATEGORY,
                        NVL(U.USERNAME, 'Unknown') AS UPLOADED_BY,
                        D.UPLOADED_ON
                    FROM TASK_DOCUMENTS D
                    LEFT JOIN TASKS T ON D.TASK_ID = T.TASK_ID
                    LEFT JOIN USERS U ON D.UPLOADED_BY = U.USER_ID
                    LEFT JOIN PROJECTS P ON D.PROJECT_ID = P.PROJECT_ID
                    WHERE D.PROJECT_ID = :projectId
                    ORDER BY D.UPLOADED_ON DESC";

                OracleCommand cmd = new OracleCommand(query, con);
                cmd.BindByName = true;

                cmd.Parameters.Add("projectId", projectId);


                OracleDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    docs.Add(new
                    {
                        fileName = dr["FILE_NAME"].ToString(),
                        filePath = dr["FILE_PATH"].ToString(),

                        taskName = dr["TASK_NAME"].ToString(),
                        projectName = dr["PROJECT_NAME"].ToString(),

                        groupName = dr["PROJECT_GROUP"] == DBNull.Value
                            ? "No Group"
                            : dr["PROJECT_GROUP"].ToString(),
                        categoryName = dr["PROJECT_CATEGORY"].ToString(),

                        uploadedByName = dr["UPLOADED_BY"].ToString(),

                        uploadedOn =
                            dr["UPLOADED_ON"] == DBNull.Value
                                ? ""
                        : Convert.ToDateTime(dr["UPLOADED_ON"])
                            .ToString("dd-MMM-yyyy hh:mm tt")
                    });
                }
            }

            return Json(docs, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public JsonResult AddProject(ProjectModel model)
        {
            using (var con = OracleDbHelper.GetConnection())
            {
                con.Open();

                // STEP 1: Insert Project
                string projectQuery = @"
                    INSERT INTO PROJECTS
                    (PROJECT_ID, PROJECT_NAME,PROJECT_GROUP, PROJECT_CATEGORY)
                    VALUES
                    (PROJECT_SEQ.NEXTVAL, :projectName, :groupName, :categoryName)
                    RETURNING PROJECT_ID INTO :newProjectId";

                OracleCommand projectCmd =
                    new OracleCommand(projectQuery, con);

                projectCmd.BindByName = true;

                OracleParameter newProjectId =
                    new OracleParameter("newProjectId", OracleDbType.Int32);

                newProjectId.Direction =
                    System.Data.ParameterDirection.Output;

                projectCmd.Parameters.Add("projectName", model.projectName);
                projectCmd.Parameters.Add(newProjectId);
                projectCmd.Parameters.Add("groupName", model.GroupName);

                projectCmd.Parameters.Add("categoryName", model.categoryName);

                projectCmd.ExecuteNonQuery();

                int projectId =
                    ((Oracle.ManagedDataAccess.Types.OracleDecimal)
                        newProjectId.Value).ToInt32();






                if (model.tasks == null)
                {
                    return Json(new
                    {
                        success = false,
                        message = "No tasks received"
                    });
                }
                // STEP 3: Insert Tasks
                foreach (var task in model.tasks)
                {

                    string taskQuery = @"
                    INSERT INTO TASKS
                    (
                        TASK_ID,
                        PROJECT_ID,
                        TASK_NAME,
                        PLANNED_START,
                        PLANNED_END,
                        STATUS,
                        PROGRESS
                    )
                    VALUES
                    (
                        TASK_SEQ.NEXTVAL,
                        :projectId,
                        :taskName,
                        :plannedStart,
                        :plannedEnd,
                        'PENDING',
                        0
                    )";

                    OracleCommand taskCmd =
                        new OracleCommand(taskQuery, con);

                    taskCmd.BindByName = true;

                    taskCmd.Parameters.Add("projectId", projectId);
                    taskCmd.Parameters.Add("taskName", task.title);
                    taskCmd.Parameters.Add("plannedStart", Convert.ToDateTime(task.start));
                    taskCmd.Parameters.Add("plannedEnd", Convert.ToDateTime(task.end));

                    System.Diagnostics.Debug.WriteLine("Project Name: " + model.projectName);

                    if (model.tasks != null)
                    {
                        foreach (var t in model.tasks)
                        {
                            System.Diagnostics.Debug.WriteLine("Task: " + t.title);
                            System.Diagnostics.Debug.WriteLine("Start: " + t.start);
                            System.Diagnostics.Debug.WriteLine("End: " + t.end);
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("Tasks are NULL");
                    }

                    taskCmd.ExecuteNonQuery();
                }
            }

            return Json(new { success = true });
        }

        [HttpPost]
        public JsonResult ImportProjects(HttpPostedFileBase file, string groupName, string sheetName, int projectCol, int taskCol, int startCol, int endCol, int statusCol, int categoryCol)
        { 
        
            if (file == null || file.ContentLength == 0)
                return Json(new { success = false, message = "No file selected" });

            using (var con = OracleDbHelper.GetConnection())
            {
                con.Open();

                using (var workbook = new XLWorkbook(file.InputStream))
                {
                    var ws = workbook.Worksheet(sheetName);

                    if (ws == null)
                    {
                        return Json(new
                        {
                            success = false,
                            message = "Selected Sheet not found"
                        });
                    }

                    int row = 2;

                    for (int c = 1; c <= 15; c++)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            "Row " + row + " Col " + c + ": " + ws.Cell(row, c).GetString()
                        );
                    }
                    int count = 0;
                    //int taskCount = 0;



                    

                    while (true)
                    {


                        string projectName = ws.Cell(row, projectCol).GetString().Trim();
                        string taskName =
                             taskCol == 0 ? "" : ws.Cell(row, taskCol).GetString().Trim();

                        if (string.IsNullOrWhiteSpace(projectName))
                            break;

                        string categoryName =
                            categoryCol == 0 ? "" : ws.Cell(row, categoryCol).GetString().Trim();

                        string startText =
                            startCol == 0 ? "" : ws.Cell(row, startCol).GetString().Trim();
                        string endText =
                            endCol == 0 ? "" : ws.Cell(row, endCol).GetString().Trim();
                        string status =
                            statusCol == 0 ? "PENDING" : ws.Cell(row, statusCol).GetString().Trim();




                        DateTime plannedStart;
                        DateTime plannedEnd;

                        DateTime.TryParse(startText, out plannedStart);
                        DateTime.TryParse(endText, out plannedEnd);

                        int projectId = 0;

                        // Check project
                        string checkQuery = @"
                            SELECT PROJECT_ID
                            FROM PROJECTS
                            WHERE UPPER(TRIM(PROJECT_NAME)) = UPPER(TRIM(:projectName))
                            AND UPPER(TRIM(PROJECT_GROUP)) = UPPER(TRIM(:groupName))";

                        OracleCommand checkCmd = new OracleCommand(checkQuery, con);
                        checkCmd.Parameters.Add("projectName", projectName);
                        checkCmd.Parameters.Add("groupName", groupName);

                        object existingProject = checkCmd.ExecuteScalar();

                        if (existingProject == null)
                        {
                            string query = @"
                            INSERT INTO PROJECTS
                            (
                                PROJECT_ID,
                                PROJECT_NAME,
                                START_DATE,
                                END_DATE,
                                STATUS,
                                PROJECT_GROUP,
                                PROJECT_CATEGORY        
                            )
                            VALUES
                            (
                                PROJECT_SEQ.NEXTVAL,
                                :projectName,
                                :startDate,
                                :endDate,
                                :status,
                                :projectGroup,
                                :projectCategory
                            )";

                            OracleCommand cmd = new OracleCommand(query, con);

                            cmd.Parameters.Add("projectName", projectName);
                            cmd.Parameters.Add("startDate", plannedStart);
                            cmd.Parameters.Add("endDate", plannedEnd);
                            cmd.Parameters.Add("status", status);
                            cmd.Parameters.Add("projectGroup", groupName);
                            cmd.Parameters.Add("projectCategory", categoryName);

                            cmd.ExecuteNonQuery();

                            OracleCommand getIdCmd =
                                new OracleCommand("SELECT PROJECT_SEQ.CURRVAL FROM DUAL", con);

                            projectId = Convert.ToInt32(getIdCmd.ExecuteScalar());

                            count++;
                        }
                        else
                        {
                            projectId = Convert.ToInt32(existingProject);
                        }

                        // Insert task only if task column has value
                        if (!string.IsNullOrWhiteSpace(taskName))
                        {
                            string checkTaskQuery = @"
                                SELECT COUNT(*)
                                FROM TASKS
                                WHERE PROJECT_ID = :projectId
                                AND UPPER(TRIM(TASK_NAME)) = UPPER(TRIM(:taskName))";

                            OracleCommand checkTaskCmd =
                                new OracleCommand(checkTaskQuery, con);

                            checkTaskCmd.Parameters.Add("projectId", projectId);
                            checkTaskCmd.Parameters.Add("taskName", taskName);

                            int taskExists =
                                Convert.ToInt32(checkTaskCmd.ExecuteScalar());

                            if (taskExists == 0)
                            {
                                string taskQuery = @"
                                INSERT INTO TASKS
                                (
                                    TASK_ID,
                                    PROJECT_ID,
                                    TASK_NAME,
                                    PLANNED_START,
                                    PLANNED_END,
                                    STATUS,
                                    PROGRESS
                                )
                                VALUES
                                (
                                    TASK_SEQ.NEXTVAL,
                                    :projectId,
                                    :taskName,
                                    :plannedStart,
                                    :plannedEnd,
                                    :status,
                                    0
                                )";

                                OracleCommand taskCmd = new OracleCommand(taskQuery, con);

                                taskCmd.Parameters.Add("projectId", projectId);
                                taskCmd.Parameters.Add("taskName", taskName);
                                taskCmd.Parameters.Add("plannedStart", plannedStart);
                                taskCmd.Parameters.Add("plannedEnd", plannedEnd);
                                taskCmd.Parameters.Add("status", status);

                                taskCmd.ExecuteNonQuery();
                            }
                        }

                        row++;
                    }

                    return Json(new
                    {
                        success = true,
                        message = count + " projects imported successfully"
                    });
                }
            }
        }


        [HttpPost]
        public JsonResult AddGroup(string groupName)
        {
            using (var con = OracleDbHelper.GetConnection())
            {
                con.Open();

                string query = @"
            INSERT INTO PROJECT_GROUPS
            (
                GROUP_ID,
                GROUP_NAME
            )
            VALUES
            (
                GROUP_SEQ.NEXTVAL,
                :groupName
            )";

                OracleCommand cmd = new OracleCommand(query, con);
                cmd.Parameters.Add("groupName", groupName);

                cmd.ExecuteNonQuery();
            }

            return Json(new { success = true });
        }

        [HttpPost]
        public JsonResult GetSheetColumns(HttpPostedFileBase file, string sheetName)
        {
            var columns = new List<object>();

            using (var workbook = new XLWorkbook(file.InputStream))
            {
                var ws = workbook.Worksheet(sheetName);

                int col = 1;

                while (true)
                {
                    string header = ws.Cell(1, col).GetString().Trim();

                    if (string.IsNullOrWhiteSpace(header))
                        break;

                    System.Diagnostics.Debug.WriteLine(
                        "Column " + col + " Header: " + header
                    );

                    columns.Add(new
                    {
                        Index = col,
                        Name = header
                    });

                    col++;
                }
            }

            return Json(columns, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public JsonResult AddTask()
        {
            string body;

            using (var reader = new System.IO.StreamReader(Request.InputStream))
            {
                body = reader.ReadToEnd();
            }

            var task = Newtonsoft.Json.JsonConvert.DeserializeObject<TaskModel>(body);

            if (string.IsNullOrWhiteSpace(task.title))
            {
                task.title = "New Task";
            }

            if (!CanEdit())
                return Json(new { success = false, message = "Unauthorized" });

            using (var con = OracleDbHelper.GetConnection())
            {
                con.Open();

                task.projectId = task.projectId == 0 ? 1 : task.projectId;

                string query = @"
                    INSERT INTO TASKS
                    (   
                        TASK_ID,
                        PROJECT_ID,
                        TASK_NAME,
                        ASSIGNED_TO,
                        PLANNED_START,
                        PLANNED_END,
                        STATUS,
                        PROGRESS
                    )
                    VALUES
                    (
                        TASK_SEQ.NEXTVAL,
                        :projectId,
                        :taskName,
                        :assignedTo,
                        :plannedStart,
                        :plannedEnd,
                        :status,
                        :progress
                    )
                    RETURNING TASK_ID INTO :newId";

                OracleCommand cmd = new OracleCommand(query, con);
                cmd.BindByName = true;

                OracleParameter newIdParam = new OracleParameter("newId", OracleDbType.Int32);
                newIdParam.Direction = System.Data.ParameterDirection.Output;

                object assignedValue = DBNull.Value;

                if (task.assignedTo.HasValue)
                {
                    string checkUser =
                        "SELECT COUNT(*) FROM USERS WHERE USER_ID = :userId";

                    OracleCommand checkCmd =
                        new OracleCommand(checkUser, con);

                    checkCmd.Parameters.Add("userId", task.assignedTo.Value);

                    int exists = Convert.ToInt32(checkCmd.ExecuteScalar());

                    if (exists > 0)
                    {
                        assignedValue = task.assignedTo.Value;
                    }
                }

                DateTime startDate =
                    task.start.HasValue
                        ? task.start.Value.AddHours(5).AddMinutes(30).Date
                        : DateTime.Today;

                DateTime endDate =
                    task.end.HasValue
                        ? task.end.Value.AddHours(5).AddMinutes(30).Date
                        : startDate.AddDays(7);

                string status = "PENDING";

                if (task.percentComplete.HasValue &&
                    task.percentComplete > 0 &&
                    task.percentComplete < 1)
                    status = "IN_PROGRESS";
                else if (task.percentComplete.HasValue &&
                        task.percentComplete == 1)
                    status = "COMPLETED";

                cmd.Parameters.Add("projectId", task.projectId);
                cmd.Parameters.Add("taskName", task.title);
                OracleParameter assignedParam =
                    new OracleParameter("assignedTo", OracleDbType.Int32);

                assignedParam.Value = assignedValue;

                cmd.Parameters.Add(assignedParam);
                cmd.Parameters.Add("plannedStart", startDate);
                cmd.Parameters.Add("plannedEnd", endDate);
                cmd.Parameters.Add("status", status);
                cmd.Parameters.Add(
                    "progress",
                    Convert.ToInt32((task.percentComplete ?? 0) * 100)
                );
                cmd.Parameters.Add(newIdParam);

                System.Diagnostics.Debug.WriteLine("Task: " + task.title);
                System.Diagnostics.Debug.WriteLine("Project: " + task.projectId);

                System.Diagnostics.Debug.WriteLine(
                    "Saving ASSIGNED_TO: " + assignedValue
                );

                cmd.ExecuteNonQuery();


                int newTaskId = ((Oracle.ManagedDataAccess.Types.OracleDecimal)newIdParam.Value).ToInt32();
                task.id = newTaskId;

                System.Diagnostics.Debug.WriteLine("New Task ID: " + newTaskId);
                AddAuditLog(
                    con,
                    newTaskId,
                    "CREATE",
                    null,
                    "Created task: " + task.title
                );
            }

            return Json(task);
        }

        [HttpPost]
        public JsonResult UpdateTask()
        {
            string body;

            using (var reader = new System.IO.StreamReader(Request.InputStream))
            {
                body = reader.ReadToEnd();
            }

            var task = Newtonsoft.Json.JsonConvert.DeserializeObject<TaskModel>(body);

            if (string.IsNullOrWhiteSpace(task.title))
            {
                task.title = "New Task";
            }
            System.Diagnostics.Debug.WriteLine("UpdateTask HIT");
            if (!CanEdit())
                return Json(new { success = false, message = "Unauthorized" });

            using (var con = OracleDbHelper.GetConnection())
            {
                con.Open();

                double progress = task.percentComplete ?? 0;

                DateTime? actualStart = null;
                DateTime? actualEnd = null;

                if (progress > 0)
                    actualStart = DateTime.Today;

                if (progress == 1)
                    actualEnd = DateTime.Today;

                string status = "PENDING";

                if (progress > 0 && progress < 1)
                    status = "IN_PROGRESS";
                else if (progress == 1)
                    status = "COMPLETED";

                string query = @"UPDATE TASKS
                         SET PROJECT_ID = :projectId,
                             TASK_NAME = :taskName,
                             ASSIGNED_TO = :assignedTo,
                             PARENT_TASK_ID = :parentId,
                             PLANNED_START = :plannedStart,
                             PLANNED_END = :plannedEnd,
                             ACTUAL_START = NVL(ACTUAL_START, :actualStart),
                             ACTUAL_END = :actualEnd,
                             STATUS = :status,
                             PROGRESS = :progress
                         WHERE TASK_ID = :taskId";

                OracleCommand cmd = new OracleCommand(query, con);
                cmd.BindByName = true;

                object assignedValue = DBNull.Value;

                if (!string.IsNullOrWhiteSpace(task.assignedToName))
                {
                    string userQuery =
                        "SELECT USER_ID FROM USERS WHERE UPPER(TRIM(USERNAME)) = UPPER(TRIM(:username))";

                    OracleCommand userCmd = new OracleCommand(userQuery, con);
                    userCmd.BindByName = true;

                    userCmd.Parameters.Add("username", task.assignedToName.Trim());

                    object result = userCmd.ExecuteScalar();

                    if (result != null && result != DBNull.Value)
                    {
                        assignedValue = Convert.ToInt32(result);
                    }
                }

                cmd.Parameters.Add("projectId", task.projectId);
                cmd.Parameters.Add("taskName", task.title);
                OracleParameter assignedParam =
                     new OracleParameter("assignedTo", OracleDbType.Int32);

                assignedParam.Value = assignedValue;

                cmd.Parameters.Add(assignedParam);
                cmd.Parameters.Add("parentId",
                    task.parentId.HasValue ? (object)task.parentId.Value : DBNull.Value);
                cmd.Parameters.Add(
                    "plannedStart",
                    task.start.Value.AddHours(5).AddMinutes(30).Date
                );

                cmd.Parameters.Add(
                    "plannedEnd",
                    task.end.Value.AddHours(5).AddMinutes(30).Date
                );
                cmd.Parameters.Add("actualStart",
                    actualStart.HasValue ? (object)actualStart.Value : DBNull.Value);
                cmd.Parameters.Add("actualEnd",
                    actualEnd.HasValue ? (object)actualEnd.Value : DBNull.Value);
                cmd.Parameters.Add("status", status);
                cmd.Parameters.Add("progress", Convert.ToInt32(progress * 100));
                cmd.Parameters.Add("taskId", task.id);

                string oldQuery = "SELECT TASK_NAME FROM TASKS WHERE TASK_ID = :taskId";
                OracleCommand oldCmd = new OracleCommand(oldQuery, con);
                oldCmd.Parameters.Add(new OracleParameter("taskId", task.id));

                object oldResult = oldCmd.ExecuteScalar();
                string oldTaskName = oldResult != null ? oldResult.ToString() : "";

                System.Diagnostics.Debug.WriteLine("AssignedToName: " + task.assignedToName);
                System.Diagnostics.Debug.WriteLine("Resolved AssignedValue: " + assignedValue);

                int rows = cmd.ExecuteNonQuery();



                AddAuditLog(
                    con,
                    task.id ?? 0,
                    "UPDATE",
                    "Old Task: " + oldTaskName,
                    "Updated Task: " + task.title
                );

                return Json(new
                {
                    success = true,
                    updated = rows
                });
            }
        }

        [HttpPost]
        public JsonResult DeleteTask(TaskModel task)
        {
            if (!CanEdit())
                return Json(new { success = false, message = "Unauthorized" });

            using (var con = OracleDbHelper.GetConnection())
            {
                con.Open();

                string oldQuery =
                    "SELECT TASK_NAME FROM TASKS WHERE TASK_ID = :taskId";

                OracleCommand oldCmd =
                    new OracleCommand(oldQuery, con);

                oldCmd.Parameters.Add("taskId", task.id);

                object oldResult = oldCmd.ExecuteScalar();

                string oldTaskName =
                    oldResult != null ? oldResult.ToString() : "";

                AddAuditLog(
                    con,
                    task.id ?? 0,
                    "DELETE",
                    "Deleted Task: " + oldTaskName,
                    null
                );
                System.Diagnostics.Debug.WriteLine("Step 1: Audit done");

                // Delete dependencies first
                string depQuery =
                     @"DELETE FROM TASK_DEPENDENCIES
                      WHERE PREDECESSOR_TASK_ID = :taskId
                      OR SUCCESSOR_TASK_ID = :taskId";
                OracleCommand depCmd =
                    new OracleCommand(depQuery, con);

                depCmd.Parameters.Add("taskId", task.id);
                depCmd.ExecuteNonQuery();

                System.Diagnostics.Debug.WriteLine("Step 2: Dependencies deleted");

                // Delete child tasks
                string childQuery =
                    "DELETE FROM TASKS WHERE PARENT_TASK_ID = :taskId";

                OracleCommand childCmd =
                    new OracleCommand(childQuery, con);

                childCmd.Parameters.Add("taskId", task.id);
                childCmd.ExecuteNonQuery();

                System.Diagnostics.Debug.WriteLine("Step 3: Children deleted");

                // Delete actual task
                string query =
                    "DELETE FROM TASKS WHERE TASK_ID = :taskId";

                OracleCommand cmd =
                    new OracleCommand(query, con);

                cmd.Parameters.Add("taskId", task.id);

                System.Diagnostics.Debug.WriteLine("Deleting Task ID: " + task.id);

                cmd.ExecuteNonQuery();

                System.Diagnostics.Debug.WriteLine("Task Deleted Successfully");
            }

            return Json(new { data = task });
        }

        [HttpPost]
        public JsonResult GetExcelSheets(HttpPostedFileBase file)
        {
            var sheets = new List<string>();

            using (var workbook = new XLWorkbook(file.InputStream))
            {
                foreach (var ws in workbook.Worksheets)
                {
                    sheets.Add(ws.Name);
                }
            }

            return Json(sheets);
        }

        [HttpPost]
        public JsonResult UploadAttachment(int taskId, HttpPostedFileBase file)
        {

            var path = Path.Combine(Server.MapPath("~/Uploads"), file.FileName);
            file.SaveAs(path);
            if (file == null || file.ContentLength == 0)
                return Json(new { success = false });

            string folder = Server.MapPath("~/Uploads/");

            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);

            string fileName = Path.GetFileName(file.FileName);
            string uniqueName = taskId + "_" + DateTime.Now.Ticks + "_" + fileName;

            string fullPath = Path.Combine(folder, uniqueName);

            file.SaveAs(fullPath);

            using (var con = OracleDbHelper.GetConnection())
            {
                con.Open();

                object sessionUser = Session["USER_ID"];

                int? userId = null;

                if (sessionUser != null)
                {
                    userId = Convert.ToInt32(sessionUser);
                }

                // Get project id of task
                string getProjectQuery = @"
                    SELECT PROJECT_ID
                    FROM TASKS
                    WHERE TASK_ID = :taskId";

                OracleCommand projectCmd =
                    new OracleCommand(getProjectQuery, con);

                projectCmd.Parameters.Add("taskId", taskId);

                int projectId =
                    Convert.ToInt32(projectCmd.ExecuteScalar());

                // Check if same filename exists for same task
                string checkQuery = @"
                    SELECT DOCUMENT_ID
                    FROM TASK_DOCUMENTS
                    WHERE TASK_ID = :taskId
                    AND UPPER(TRIM(FILE_NAME)) = UPPER(TRIM(:fileName))";

                OracleCommand checkCmd =
                    new OracleCommand(checkQuery, con);

                checkCmd.Parameters.Add("taskId", taskId);
                checkCmd.Parameters.Add("fileName", fileName);

                object existingDoc = checkCmd.ExecuteScalar();

                if (existingDoc != null)
                {
                    // overwrite same file
                    string updateQuery = @"
                        UPDATE TASK_DOCUMENTS
                        SET FILE_PATH = :filePath,
                            UPLOADED_BY = :uploadedBy,
                            UPLOADED_ON = SYSDATE
                        WHERE DOCUMENT_ID = :documentId";

                    OracleCommand updateCmd =
                        new OracleCommand(updateQuery, con);

                    updateCmd.Parameters.Add("filePath", "/Uploads/" + uniqueName);
                    updateCmd.Parameters.Add("uploadedBy",
                        userId.HasValue ? (object)userId.Value : DBNull.Value);
                    updateCmd.Parameters.Add("documentId", existingDoc);

                    updateCmd.ExecuteNonQuery();
                }
                else
                {
                    // add new file
                    string insertQuery = @"
                    INSERT INTO TASK_DOCUMENTS
                    (
                        DOCUMENT_ID,
                        TASK_ID,
                        PROJECT_ID,
                        FILE_NAME,
                        FILE_PATH,
                        UPLOADED_BY,
                        UPLOADED_ON
                    )
                    VALUES
                    (
                        TASK_DOCUMENTS_SEQ.NEXTVAL,
                        :taskId,
                        :projectId,
                        :fileName,
                        :filePath,
                        :uploadedBy,
                        SYSDATE
                    )";

                    OracleCommand insertCmd =
                        new OracleCommand(insertQuery, con);

                    insertCmd.Parameters.Add("taskId", taskId);
                    insertCmd.Parameters.Add("projectId", projectId);
                    insertCmd.Parameters.Add("fileName", fileName);
                    insertCmd.Parameters.Add("filePath", "/Uploads/" + uniqueName);
                    insertCmd.Parameters.Add("uploadedBy",
                        userId.HasValue ? (object)userId.Value : DBNull.Value);

                    insertCmd.ExecuteNonQuery();
                }
            }

            return Json(new { success = true });
        }


    }
}