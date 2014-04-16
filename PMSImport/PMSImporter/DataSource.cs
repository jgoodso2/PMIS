using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using PSLibrary = Microsoft.Office.Project.Server.Library;

namespace PMSImporter
{
    public class DataSource:IDataSource
    {

        public DataSet ReadData(string fileName)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Project UID", typeof(Guid));
            dt.Columns.Add("Project Type", typeof(int));
            dt.Columns.Add("Project Name", typeof(string));
            dt.Columns.Add("Project Author", typeof(string));
            dt.Columns.Add("Project Manager", typeof(string));
            dt.Columns.Add("Project Title", typeof(string));
            dt.Columns.Add("Task UID", typeof(Guid));
            dt.Columns.Add("Task Parent UID", typeof(Guid));
            dt.Columns.Add("Task Name", typeof(string));
            dt.Columns.Add("Task ID", typeof(int));
            dt.Columns.Add("Task Duration", typeof(string));
            dt.Columns.Add("Task Start", typeof(DateTime));
            dt.Columns.Add("Task Finish", typeof(DateTime));
            dt.Columns.Add("TB Start", typeof(DateTime));
            dt.Columns.Add("TB Finish", typeof(DateTime));

            var projectId = Guid.NewGuid();
             var prow  =dt.NewRow();
             prow["Project UID"] = projectId;
             prow["Project Type"] = (int)Enum.Parse(typeof(PSLibrary.Project.ProjectType), "Project");
             prow["Project Name"] = "Test Project 1";
             prow["Project Manager"] = "Nishant";
             prow["Project Author"] = "Nishant";
            dt.Rows.Add(prow);
            for (int i = 0; i < 20; i++)
            {
                var row = dt.NewRow();
                row["Project UID"] = projectId;
                row["Project Type"] = (int)Enum.Parse(typeof(PSLibrary.Project.ProjectType), "Project");
                row["Project Name"] = "Test Project 1";
                row["Project Manager"] = "Nishant";
                row["Project Author"] = "Nishant";
                row["Task UID"] = Guid.NewGuid();
                row["Task Name"] = "Test Task" + i;
                row["Task ID"] =  i;
                row["Task Duration"] = i;
                row["Task Start"] = DateTime.Now.AddDays(i);
                row["Task Finish"] = DateTime.Now.AddDays(i + 5);
                row["TB Start"] = DateTime.Now.AddDays(i);
                row["TB Finish"] = DateTime.Now.AddDays(i + 5);

                dt.Rows.Add(row);
            }
            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
            return ds;
        }
    }
}
