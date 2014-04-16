using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;

namespace PMSImporter
{
    public class XLDataSource :IDataSource
    {
        public DataSet ReadData(string fileName)
        {
            return GenerateExcelData(fileName);
        }
        private DataSet GenerateExcelData(string fileName)
        {
            Console.WriteLine("generate excel data started");
            DataSet projectDs = new DataSet();
            DataTable projectTable = new DataTable();
            OleDbConnection oledbConn = null;
            try
        {
            if (!File.Exists(fileName)) throw new Exception("No excel input file specified");
            oledbConn = new OleDbConnection(String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 8.0;HDR=YES;\""));
             
           
                oledbConn.Open();
                OleDbCommand cmd = new OleDbCommand(); ;
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                cmd.Connection = oledbConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "select * from [Sheet1$]";
                oleda = new OleDbDataAdapter(cmd);
                oleda.Fill(ds, "Project");
                var groups = ds.Tables[0].AsEnumerable().GroupBy(t => t.Field<string>("Project Name"));
            int i = 0;
                foreach (var group in groups)
            {
                projectTable = ds.Tables[0].Clone();
                projectTable.TableName = "Project" + i++;
                var rowList = ds.Tables[0].AsEnumerable().Where(t => t.Field<string>("Project Name") == group.Key);
                foreach (DataRow row in rowList)
                    projectTable.ImportRow(row);
                projectDs.Tables.Add(projectTable);
            }
            
            return projectDs;
        }
        // need to catch possible exceptions
        catch (Exception ex)
        {
            throw ex;
        }
        finally
        {
             oledbConn.Close();
        }
            Console.WriteLine("generate excel data done successfully");
        } // close of method GemerateExceLData
    }
}
