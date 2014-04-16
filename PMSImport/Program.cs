using System;
using System.Data;
using PMSImporter;
using SvcProject;
using System.Configuration;

namespace PMSImport
{
    class Program
    {
        static void Main()
        {
            Repository.SetProjectServerUrl(ConfigurationManager.AppSettings["ProjectServerUrl"]);
            //PMSImporter.PMSImporter.Import(ConfigurationManager.AppSettings["XLFileName"]);
            Console.ReadKey();
        }
    }
}
