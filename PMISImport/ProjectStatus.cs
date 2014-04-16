using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PMISImport
{
    public class ProjectStatus
    {
        public string ProjectName {get;set;}
        public int SuccessCount{get;set;}
        public int FailedCount { get; set; }
        public string Status{get;set;}
    }
}
