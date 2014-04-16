using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PMSImporter;

namespace TestPMSImport
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            Repository.SaveProject(new Guid("CABDE72A-B847-4A16-A188-5C961D25A641"));

        }
    }
}
