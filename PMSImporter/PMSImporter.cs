using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using SvcProject;

namespace PMSImporter
{
    public class PMSImporter
    {
        //public static void Import(string fileName)
        //{
        //    XLDataSource source = new XLDataSource();
        //    Repository repository = new Repository();
        //    List<string> successfulProjects = new List<string>();
        //    List<string> failedProjects = new List<string>();
        //    List<string> errors;
        //    DataSet ds = source.ReadData(fileName);
        //    if (ds.Tables.Count > 0 && !DataSetBuilder.Validate(ds.Tables[0],out errors))
        //    {
        //        Console.WriteLine("Project Import aborted due to validation error");
        //        return ;
        //    }
        //    foreach (DataTable table in ds.Tables)
        //    {
        //        ProjectDataSet projectDataSet;
        //        ProjectDataSet.ProjectRow row = null;
        //        try
        //        {
        //            bool iscreate;
        //            DataSetBuilder.BuildProjectDataSetForCreate(table, out projectDataSet, out row,out iscreate);
        //            DataSetBuilder.BuildProjectDataSetForUpdate(table, projectDataSet, row, iscreate);
        //            successfulProjects.Add(row.PROJ_NAME);
        //        }
        //        catch (Exception ex)
        //        {
        //            if (row != null)
        //            {
        //                Console.WriteLine("An error occured.Skipping Project import for {0}. Failure reason = {1}", row.PROJ_NAME, ex.Message);
        //                failedProjects.Add(row.PROJ_NAME);
        //            }
        //            else
        //            {
        //                Console.WriteLine("An error occured.Skipping Project import. Failure reason = {0}", ex.Message);
        //                failedProjects.Add("");
        //            }
        //            continue;
        //        }
        //    }
        //    Console.WriteLine("=================================================================================================================");
        //    Console.WriteLine("SUMMARY OF PROJECT IMPORT");

        //    Console.WriteLine("Total no of projects to import = {0}", (successfulProjects.Count + failedProjects.Count).ToString());
        //    if (successfulProjects.Count > 0)
        //    {
        //        Console.WriteLine("No of projects successfully imported = {0}", successfulProjects.Count.ToString());

        //        foreach (string project in successfulProjects)
        //        {
        //            Console.WriteLine("{0}", project);
        //        }
        //    }

        //    if (failedProjects.Count > 0)
        //    {
        //        Console.WriteLine("No of projects failed to import = {0}", failedProjects.Count.ToString());

        //        foreach (string project in failedProjects)
        //        {
        //            Console.WriteLine("{0}", project);
        //        }
        //    }
        //}
    }
}