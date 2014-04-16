

using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Security.Principal;
using System.ServiceModel;
using System.Xml;
using Microsoft.Office.Project.Server.Library;
using SvcCustomFields;
using SvcProject;

namespace PMSImporter
{
    public class DataSetBuilder
    {
        private static CustomFieldDataSet customDataSet;
        private static Mapping _mapping;
        private static Dictionary<string, Guid> resUids = new Dictionary<string, Guid>();
        private static List<string> _resourcesList;

        public static Mapping Mapping { get { return _mapping; } }
        static DataSetBuilder()
        {
            Console.WriteLine("Starting Set up for PMS Import");
            customDataSet = Repository.customFieldsClient.ReadCustomFields("", false);
            
            Console.WriteLine("Set up for PMS Import done successfully");
        }

        public static bool Validate(DataTable inputTable,string[] projectNames, out List<string> errors)
        {
            errors = new List<string>();
            return ValidateCustomFields(inputTable, errors) && ValidateProjectCheckedOut(projectNames,errors);
        }

        private static bool ValidateProjectCheckedOut(string[] projectNames, List<string> errors)
        {
            bool valid = true;
            foreach(string projectName in projectNames)
           {
                if(Repository.IsProjectCheckedOut(projectName))
                {
                    errors.Add(string.Format("Project {0} is already checked out.",projectName));
                    valid = false;
                }
           }
            return valid;
        }

        public static bool ValidateCustomFields(DataTable inputTable, List<string> errors)
        {
            return (ValidateProjectCustomFields(inputTable, errors) && ValidateTaskCustomFields(inputTable, errors)) && ValidateResourceColumns(inputTable, errors);
        }

        public static bool ValidateSourceFields(DataTable inputTable, List<string> errors)
        {
            bool valid = true;
            foreach (DataColumn column in inputTable.Columns)
            {
                if (!(_mapping.ProjectMap.ContainsValue(column.ColumnName) || _mapping.TaskMap.ContainsValue(column.ColumnName)
                    || _mapping.AssignmentMap.ContainsValue(column.ColumnName) || _mapping.ProjectCustomFieldsMap.ContainsValue(column.ColumnName)
                     || _mapping.TaskCustomFieldsMap.ContainsValue(column.ColumnName)))
                {
                    Console.WriteLine("Validation Error: Source mapping not found  in the Mapping file for the source column = {0}", column.ColumnName);
                    errors.Add(string.Format("Validation Error: Source mapping not found  in the Mapping file for the source column = {0}", column.ColumnName));
                    valid = false;
                }

            }
            return valid;
        }

        public static bool ValidateResourceColumns(DataTable inputTable, List<string> errors)
        {
            return ValidateResourceColumns(_mapping.Project.Fields.Where(t => t.IsResourceColumn == true), inputTable, errors)
                && ValidateResourceColumns(_mapping.Project.CustomFields.Where(t => t.IsResourceColumn == true), inputTable, errors)
                && ValidateResourceColumns(_mapping.Assignment.Fields.Where(t => t.IsResourceColumn == true), inputTable, errors)
                //&& ValidateResourceColumns(_mapping.Assignment.CustomFields.Where(t => t.IsResourceColumn == true), inputTable, errors)
                && ValidateResourceColumns(_mapping.Task.Fields.Where(t => t.IsResourceColumn == true), inputTable, errors)
                && ValidateResourceColumns(_mapping.Task.CustomFields.Where(t => t.IsResourceColumn == true), inputTable, errors);
        }

        private static bool ValidateResourceColumns(IEnumerable<Field> fields, DataTable inputTable, List<string> errors)
        {
            bool valid = true;
            if (fields != null)
            {
                foreach (var field in fields)
                {
                    
                    if (inputTable.HasColumn(field.Source) )
                    {
                        foreach (DataRow row in inputTable.AsEnumerable().GroupBy(t => t.Field<string>(field.Source)).Select(t=>t.First()))
                        {
                            if (row[field.Source] != System.DBNull.Value && row[field.Source] != null && !string.IsNullOrEmpty(row[field.Source].ToString()))
                            {
                                if (!Repository.GetResources().Resources.Any(t => t.RES_NAME.Trim().ToUpper() == row[field.Source].ToString().Trim().ToUpper()))
                            {
                                valid = false;
                                errors.Add(string.Format("Resource Validation Error: Resource {0} specified in the column name ={1} not found", row[field.Source].ToString(),field.Source));
                            }
                            }
                        }
                    }
                }
            }
            return valid;
        }

        public static bool ValidateProjectCustomFields(DataTable inputTable, List<string> errors)
        {
            bool valid = true;
            if (Mapping.Project !=null && Mapping.Project.CustomFields != null)
            {
                foreach (var customField in _mapping.Project.CustomFields)
                {
                    var sourcecolumn = customField.Source;

                    if (!inputTable.HasColumn(sourcecolumn))
                        continue;

                    var targetColumn = customField.Target;
                    if (!customDataSet.CustomFields.Any(t => t.MD_PROP_NAME == targetColumn))
                    {
                        Console.WriteLine("Validation Error: Incorrect target mapping found for project custom field source ={0},target={1}", sourcecolumn, targetColumn);
                        errors.Add(string.Format("Validation Error: Incorrect target mapping found for project custom field source ={0},target={1}", sourcecolumn, targetColumn));
                        valid = false;
                    }
                }
            }
            return valid;
        }

        public static bool ValidateTaskCustomFields(DataTable inputTable, List<string> errors)
        {
            bool valid = true;
            foreach (var customField in _mapping.Task.CustomFields)
            {
                var sourcecolumn = customField.Source;
                if (!inputTable.HasColumn(sourcecolumn))
                    continue;

                var targetColumn = customField.Target;
                if (!customDataSet.CustomFields.Any(t => t.MD_PROP_NAME == targetColumn))
                {
                    Console.WriteLine("Validation Error: Incorrect target mapping found for task custom field source ={0},target={1}", sourcecolumn, targetColumn);
                    errors.Add(string.Format("Validation Error: Incorrect target mapping found for task custom field source ={0},target={1}", sourcecolumn, targetColumn));
                    valid = false;
                }
            }
            return valid;
        }

        public static ProjectDataSet BuildProjectDataSetForCreate(DataTable table, out ProjectDataSet projectDataSet, out ProjectDataSet.ProjectRow projectRow,out bool iscreate)
        {
            var inputProjectRow = table.Rows[0];
            Console.WriteLine("Starting import of project for {0}", inputProjectRow[_mapping.ProjectMap["PROJ_NAME"]].ToString());
            projectDataSet = new ProjectDataSet();

            iscreate = false;

            projectRow = projectDataSet.Project.NewProjectRow();
            Guid projectGuid = Guid.NewGuid();

            projectRow.PROJ_UID = projectGuid;
            projectRow.PROJ_NAME = inputProjectRow[_mapping.ProjectMap["PROJ_NAME"]].ToString();
         
            //projectRow.ENTERPRISE_PROJECT_TYPE_NAME = inputProjectRow[_mapping.ProjectMap["ENTERPRISE_PROJECT_TYPE_NAME"]].ToString();
            // set the project start date to min of start dates for all tasks provided mapping exists for task start with key = TASK_START_DATE
            if (ProjectHasTasks(table) && _mapping.TaskMap.ContainsKey("TASK_START_DATE") && table.Columns.Contains(_mapping.TaskMap["TASK_START_DATE"].ToString()))
            {
                DateTime minDate = table.AsEnumerable().Min(t => t.Field<DateTime>(_mapping.TaskMap["TASK_START_DATE"]));
                projectRow.PROJ_INFO_START_DATE = minDate;
            }
            projectDataSet.Project.AddProjectRow(projectRow);
           
            if (!Repository.CheckIfProjectExists(inputProjectRow[_mapping.ProjectMap["PROJ_NAME"]].ToString()))
            {
                Console.WriteLine("Starting create project for {0}", inputProjectRow[_mapping.ProjectMap["PROJ_NAME"]].ToString());
                
                    BuildProjectCustomFields(projectDataSet, inputProjectRow, projectGuid);
                projectGuid =  new Repository().CreateProject(projectDataSet);  //create with minimal initiation data including EPT Type
                iscreate = true;
            }
            else
            {
                Console.WriteLine("Project already exists for {0}, Starting an update", inputProjectRow[_mapping.ProjectMap["PROJ_NAME"]].ToString());
                projectGuid = Repository.GetProjectList().Project
                    .First(t => t.PROJ_NAME.Trim().ToUpper() == inputProjectRow[_mapping.ProjectMap["PROJ_NAME"]].ToString().Trim().ToUpper()).PROJ_UID;
            }
           
            projectDataSet = Repository.ReadProject(projectGuid);
            projectRow = projectDataSet.Project.Rows[0] as SvcProject.ProjectDataSet.ProjectRow;
            BuildResources(table, projectDataSet, projectGuid); // build team
            new Repository().UpdateProjectTeam(projectDataSet);
            projectDataSet = Repository.ReadProject(projectGuid);
            projectRow = projectDataSet.Project.Rows[0] as SvcProject.ProjectDataSet.ProjectRow;
            return projectDataSet;
        }

        public static ProjectDataSet BuildProjectDataSetForUpdate(DataTable table, ProjectDataSet projectDataSet, 
            ProjectDataSet.ProjectRow projectRow, bool updateCustomFieldsnotrequired)
        {
            Console.WriteLine("Starting build project entities");
            Guid projectGuid = projectDataSet.Project[0].PROJ_UID;

            var inputProjectRow = table.Rows[0];
          
                BuildProject(inputProjectRow, projectRow, projectDataSet, projectGuid,updateCustomFieldsnotrequired);
            BuildTasks(table, projectDataSet, projectGuid);
            new Repository().UpdateProject(projectDataSet);
            Console.WriteLine("Build project entities done successfully");
            Console.WriteLine("Import of project done successfully for {0}", inputProjectRow[_mapping.ProjectMap["PROJ_NAME"]].ToString());
            return projectDataSet;
        }


        public static void BuildResources(DataTable table, ProjectDataSet resources, Guid projectGuid)
        {
            Console.WriteLine("Starting build project resources");
            DataRow inputProjectRow = table.Rows[0];
            BuildResource(table, resources, inputProjectRow, projectGuid, ResourceMode.Project);
            Console.WriteLine("Build project resources done successfully");
            for (int i = 1; i < table.Rows.Count; i++)
            {
                Console.WriteLine("Starting build task resource");
                BuildResource(table, resources, table.Rows[i], projectGuid, ResourceMode.Task);
                Console.WriteLine("Build task resource done successfully");
            }

        }

        private static void BuildResource(DataTable table, ProjectDataSet resources, DataRow inputProjectRow, Guid projectGuid, ResourceMode mode)
        {
            _resourcesList = new List<string>();
            IEnumerable<Field> data;

            if (mode == ResourceMode.Project)
            {
                data = _mapping.Project.Fields.Where(t => t.IsResourceColumn);
                if (_mapping.Project.Fields.Any(t => t.IsResourceColumn))
                {
                    foreach (var projectField in data)
                    {
                        var sourcecolumn = projectField.Source;
                        var targetColumn = projectField.Target;
                        if (!inputProjectRow.Table.HasColumn(sourcecolumn))
                        {
                            continue;
                        }
                        var inputValue = inputProjectRow[sourcecolumn];

                        if (inputValue == null || inputValue == System.DBNull.Value || inputValue.ToString() == string.Empty)
                        {
                            continue;
                        }
                        if (!_resourcesList.Contains(inputValue))
                        {
                            _resourcesList.Add(inputValue.ToString());
                        }
                    }
                }
            }
            else
            {

                for (int i = 1; i < table.Rows.Count; i++)
                {
                    if (_mapping.Task.Fields.Any(t => t.IsResourceColumn))
                    {
                        data = _mapping.Task.Fields.Where(t => t.IsResourceColumn);
                        foreach (var taskField in data)
                        {
                            var sourcecolumn = taskField.Source;
                            var targetColumn = taskField.Target;
                            var inputValue = table.Rows[i][sourcecolumn];
                            if (!_resourcesList.Contains(inputValue))
                            {
                                _resourcesList.Add(inputValue.ToString());
                            }
                        }
                    }

                    if (_mapping.Assignment.Fields.Any(t => t.IsResourceColumn))
                    {
                        data = _mapping.Assignment.Fields.Where(t => t.IsResourceColumn);
                        foreach (var assnField in data)
                        {
                            var sourcecolumn = assnField.Source;
                            var targetColumn = assnField.Target;
                            var inputValue = table.Rows[i][sourcecolumn];
                            if (!_resourcesList.Contains(inputValue))
                            {
                                _resourcesList.Add(inputValue.ToString());
                            }
                        }
                    }
                }
            }


            /// at this point the _resourcesList contains every distinct resource identified in the input file
            foreach (var newResource in _resourcesList)
            {
                if (!Repository.GetResources().Resources.Any(t => t.RES_NAME.ToUpper() == newResource.ToUpper()))
                {
                    var newRow = resources.ProjectResource.NewProjectResourceRow();
                    newRow.RES_UID = Guid.NewGuid();//Repository.GetResourceUidFromNtAccount(newResource, out isWindowsUser);
                    newRow.RES_NAME = newResource;
                    newRow.PROJ_UID = projectGuid;
                    resources.ProjectResource.Rows.Add(newRow);
                }
                else
                {
                    if (!resources.ProjectResource.Any(t => t.RES_NAME.Trim().ToUpper() == newResource.Trim().ToUpper()))
                    {
                        var newRow = resources.ProjectResource.NewProjectResourceRow();
                        if (!Repository.GetResources()
                                .Resources.Any(t => t.RES_NAME.Trim().ToUpper() == newResource.Trim().ToUpper()))
                            throw new ArgumentException("Resource not found for" + newResource);
                        newRow.RES_UID =
                            Repository.GetResources()
                                .Resources.First(t => t.RES_NAME.Trim().ToUpper() == newResource.Trim().ToUpper())
                                .RES_UID;
                        newRow.RES_NAME = newResource;
                        newRow.PROJ_UID = projectGuid;

                        resources.ProjectResource.Rows.Add(newRow);
                        newRow.AcceptChanges();
                    }
                }
            }
        }

        private static void BuildTasks(DataTable table, ProjectDataSet projectDataSet, Guid projectGuid)
        {
            Console.WriteLine("Starting build tasks");
            if (ProjectHasTasks(table))
            {
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    //Build Tasks
                    var inputTaskRow = table.Rows[i];
                    SvcProject.ProjectDataSet.TaskRow taskRow; bool isCreate = false;
                    if (_mapping.TaskMap.ContainsKey("TASK_ID") && inputTaskRow.Table.HasColumn(_mapping.TaskMap["TASK_ID"].ToString()) &&
                        inputTaskRow[_mapping.TaskMap["TASK_ID"].ToString()].ToString() != null &&
                        !string.IsNullOrEmpty(inputTaskRow[_mapping.TaskMap["TASK_ID"]].ToString()))
                    {
                        if (projectDataSet.Task.Any(t => (!t.IsTASK_IDNull() && t.TASK_ID == Convert.ToInt32(inputTaskRow[_mapping.TaskMap["TASK_ID"].ToString()].ToString())) && t.PROJ_UID == projectGuid))
                        {
                            taskRow = projectDataSet.Task.First(t => (t.TASK_ID == Convert.ToInt32(inputTaskRow[_mapping.TaskMap["TASK_ID"].ToString()].ToString())) && t.PROJ_UID == projectGuid);
                        }
                        else
                        {
                            taskRow = projectDataSet.Task.NewTaskRow();
                            isCreate = true;
                        }

                    }
                    else
                    {
                        taskRow = projectDataSet.Task.NewTaskRow();
                        isCreate = true;
                    }

                    BuildTask(projectDataSet, inputTaskRow, taskRow, projectGuid, isCreate);
                    if (isCreate)
                    {
                        projectDataSet.Task.Rows.Add(taskRow);
                    }
                    //BuildAssignments(inputTaskRow, projectDataSet, projectGuid);
                    BuildTaskCustomFields(projectDataSet, inputTaskRow, projectGuid);
                }
            }
            Console.WriteLine("build tasks done successfully");
        }

        private static bool ProjectHasTasks(DataTable table)
        {
            if (table.Rows.Count > 0)
            {
                DataRow row = table.Rows[0];
                foreach (string column in _mapping.TaskMap.Values)
                {
                    if(table.HasColumn(column) && row[column] != System.DBNull.Value && !string.IsNullOrEmpty(row[column].ToString()))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private static void BuildAssignments(DataRow inputTaskRow,
            ProjectDataSet projectDataSet, Guid projectGuid)
        {
            Console.WriteLine("Starting build assignment");
            if (inputTaskRow.Table.Columns.Contains(_mapping.AssignmentMap["RES_NAME"]))
            {
                foreach (var assnField in _mapping.Assignment.Fields)
                {
                    var sourcecolumn = assnField.Source;
                    var targetColumn = assnField.Target;

                    if (!inputTaskRow.Table.HasColumn(sourcecolumn))
                        continue;
                    var assignmentRow = projectDataSet.Assignment.NewAssignmentRow();
                    var inputValue = inputTaskRow[sourcecolumn];
                    assignmentRow.ASSN_UID = Guid.NewGuid();
                    assignmentRow.PROJ_UID = projectGuid;
                    assignmentRow.TASK_UID =
                        projectDataSet.Task.First(t => t.TASK_NAME == inputTaskRow[_mapping.TaskMap["TASK_NAME"]].ToString())
                            .TASK_UID;
                    if (resUids.ContainsKey(inputTaskRow[_mapping.AssignmentMap["RES_NAME"]].ToString()))
                    {
                        assignmentRow.RES_UID = resUids[inputTaskRow[_mapping.AssignmentMap["RES_NAME"]].ToString()];
                    }
                    else
                    {

                        if (
                            Repository.GetResources()
                                .Resources.Any(
                                    t =>
                                        t.RES_NAME.ToUpper() ==
                                        inputTaskRow[_mapping.AssignmentMap["RES_NAME"]].ToString().ToUpper()))
                        {
                            assignmentRow.RES_UID = Repository.GetResources()
                                .Resources.First(
                                    t =>
                                        t.RES_NAME.ToUpper() ==
                                        inputTaskRow[_mapping.AssignmentMap["RES_NAME"]].ToString().ToUpper()).RES_UID;
                        }
                        else
                        {
                            assignmentRow.RES_UID = projectDataSet.ProjectResource.First(t => t.RES_NAME == inputTaskRow[_mapping.AssignmentMap["RES_NAME"]].ToString()).RES_UID;
                        }


                        resUids.Add(inputTaskRow[_mapping.AssignmentMap["RES_NAME"]].ToString(), assignmentRow.RES_UID);
                    }
                    assignmentRow.RES_NAME = inputTaskRow[_mapping.AssignmentMap["RES_NAME"]].ToString();
                    assignmentRow[targetColumn] = inputValue;
                    projectDataSet.Assignment.Rows.Add(assignmentRow);
                }
            }
            Console.WriteLine("build assignment done successfully");
        }

        private static void BuildTask(ProjectDataSet projectDataSet, DataRow inputTaskRow, ProjectDataSet.TaskRow taskRow, Guid projectGuid, bool isCreate)
        {
            Console.WriteLine("build task started for {0}", inputTaskRow[_mapping.TaskMap["TASK_NAME"]].ToString());
            
            foreach (var taskField in _mapping.Task.Fields)
            {
                var sourcecolumn = taskField.Source;
                var targetColumn = taskField.Target;
                if (!inputTaskRow.Table.HasColumn(sourcecolumn))
                    continue;

                Guid guidValue;
                var inputValue = inputTaskRow[sourcecolumn];

                if (_mapping.TaskMap.ContainsKey("TASK_ID") && sourcecolumn == _mapping.TaskMap["TASK_ID"].ToString())
                {
                    continue;
                }

                if (_mapping.TaskMap.ContainsKey("TASK_ID") && sourcecolumn == _mapping.TaskMap["TASK_OUTLINE_LEVEL"].ToString())
                {
                    if (isCreate)
                    {
                        
                            //taskRow.AddAfterTaskUID = parentGuid;
                            taskRow.AddPosition = (int)Microsoft.Office.Project.Server.Library.Task.AddPositionType.Last;
                        
                    }
                }

                if (!(inputValue == System.DBNull.Value || inputValue == null || string.IsNullOrEmpty(inputValue.ToString())))
                {
                    if (taskField.MapStringToGuid == true)
                    {
                        guidValue = Repository.GetResources().Resources.First(t => t.RES_NAME.Trim().ToUpper() == inputValue.ToString().Trim().ToUpper()).RES_UID;
                        taskRow[targetColumn] = guidValue;
                    }
                    else
                    {
                        if (taskRow.Table.Columns[targetColumn].DataType == typeof(bool))
                        {
                            taskRow[targetColumn] = inputValue.ToString().Trim().ToUpper() == "YES";
                        }
                        else
                        {
                            taskRow[targetColumn] = inputValue;
                        }
                    }

                    taskRow.PROJ_UID = projectGuid;
                    if (isCreate)
                    {
                        taskRow.TASK_UID = Guid.NewGuid();
                    }
                }

            }
            
            Console.WriteLine("build task done successfully for {0}", inputTaskRow[_mapping.TaskMap["TASK_NAME"]].ToString());
        }

        private static Guid GetParentUIDForTask(DataRow inputTaskRow, ProjectDataSet.TaskRow taskRow,ProjectDataSet projectDataSet)
        {
            if(Convert.ToInt32(inputTaskRow[_mapping.TaskMap["TASK_OUTLINE_LEVEL"]].ToString()) == taskRow.TASK_OUTLINE_LEVEL)
            {
                return Guid.Empty;
            }
            if (Convert.ToInt32(inputTaskRow[_mapping.TaskMap["TASK_OUTLINE_LEVEL"]].ToString()) == 1)
            {
                return projectDataSet.Task.First(t => t.TASK_ID == 0).TASK_UID;
            }
            if (Convert.ToInt32(inputTaskRow[_mapping.TaskMap["TASK_OUTLINE_LEVEL"]].ToString()) < taskRow.TASK_OUTLINE_LEVEL)
            {
                return projectDataSet.Task.First(t => t.TASK_OUTLINE_LEVEL == (Convert.ToInt32(inputTaskRow[_mapping.TaskMap["TASK_OUTLINE_LEVEL"]].ToString()) - 1)).TASK_UID;
            }
            else
            {
                return projectDataSet.Task.First(t => t.TASK_OUTLINE_LEVEL == Convert.ToInt32(inputTaskRow[_mapping.TaskMap["TASK_OUTLINE_LEVEL"]].ToString())).TASK_UID;
            }

        }

        

        private static void BuildProject(DataRow inputProjectRow, ProjectDataSet.ProjectRow projectRow,
            ProjectDataSet projectDataSet, Guid projectGuid,bool updateCustomFieldsnotreqd)
        {
            Console.WriteLine("build project started for {0}", inputProjectRow[_mapping.ProjectMap["PROJ_NAME"]].ToString());
            foreach (var projectField in _mapping.Project.Fields)
            {
                var sourcecolumn = projectField.Source;
                
                if (!inputProjectRow.Table.HasColumn(sourcecolumn))
                    continue; 
                var inputValue = inputProjectRow[sourcecolumn];
                var targetColumn = projectField.Target;

                if (targetColumn == "ENTERPRISE_PROJECT_TYPE_UID" || targetColumn == "ENTERPRISE_PROJECT_TYPE_NAME" || targetColumn == "PROJ_TYPE" || inputValue == System.DBNull.Value ||
                   inputValue == null || inputValue.ToString() == string.Empty
                    )
                {
                    continue;
                }
                Guid guidValue;
                 //row["Input column name in the exel file"]
                if (projectField.MapStringToGuid == true)
                {
                    guidValue = Repository.GetResources().Resources.First(t => t.RES_NAME.Trim().ToUpper() == inputValue.ToString().Trim().ToUpper()).RES_UID;
                    projectRow[targetColumn] = guidValue;
                }
                else
                {
                    projectRow[targetColumn] = inputValue; // projectROW["PROJ_UID"]
                }
                
                
            }
            if (!updateCustomFieldsnotreqd)
            {

                BuildProjectCustomFields(projectDataSet, inputProjectRow, projectGuid);
            }
            Console.WriteLine("build project done successfully for {0}", inputProjectRow[_mapping.ProjectMap["PROJ_NAME"]].ToString());
        }

        private static void BuildTaskCustomFields(ProjectDataSet projectDataSet,
            DataRow inputTaskRow, Guid projectGuid)
        {
            Console.WriteLine("build task custom fields started for {0}", inputTaskRow[_mapping.TaskMap["TASK_NAME"]].ToString());
            foreach (var customField in _mapping.Task.CustomFields)
            {
                var sourcecolumn = customField.Source;
                if (!inputTaskRow.Table.HasColumn(sourcecolumn))
                    continue;
                if (inputTaskRow[sourcecolumn] == System.DBNull.Value || inputTaskRow[sourcecolumn] == null || string.IsNullOrEmpty(inputTaskRow[sourcecolumn].ToString()))
                {
                    continue;
                }

                var targetColumn = customField.Target;
                var csField =
                    customDataSet.CustomFields.First(t => t.MD_PROP_NAME == targetColumn);
                var customfieldRow = projectDataSet.TaskCustomFields.NewTaskCustomFieldsRow();
                customfieldRow.MD_PROP_ID = csField.MD_PROP_ID;
                customfieldRow.CUSTOM_FIELD_UID = Guid.NewGuid();
                customfieldRow.FIELD_TYPE_ENUM = (byte)csField.MD_PROP_TYPE_ENUM;
                customfieldRow.MD_PROP_ID = csField.MD_PROP_ID;
                customfieldRow.MD_PROP_UID = csField.MD_PROP_UID;

                customfieldRow.PROJ_UID = projectGuid;
                customfieldRow.TASK_UID =
                    projectDataSet.Task.First(t => t.TASK_NAME == inputTaskRow[_mapping.TaskMap["TASK_NAME"]].ToString()).TASK_UID;
                //If it is a lookup table custom field
                if (!csField.IsMD_LOOKUP_TABLE_UIDNull())
                {
                    var lookup = Repository.lookupTableClient.ReadLookupTablesByUids(
                        new Guid[] { csField.MD_LOOKUP_TABLE_UID }, false, CultureInfo.CurrentCulture.LCID);
                    customfieldRow.CODE_VALUE =
                        lookup.LookupTableTrees.First(t => t.LT_VALUE_FULL == inputTaskRow[sourcecolumn].ToString())
                            .LT_STRUCT_UID;
                }
                else
                {
                    switch (customfieldRow.FIELD_TYPE_ENUM)
                    {
                        case 4:
                            customfieldRow.DATE_VALUE = Convert.ToDateTime(inputTaskRow[sourcecolumn]);
                            break;
                        case 9:
                            customfieldRow.NUM_VALUE = Convert.ToDecimal(inputTaskRow[sourcecolumn]);
                            break;
                        case 6:
                            customfieldRow.DUR_VALUE = Convert.ToInt16(inputTaskRow[sourcecolumn]);
                            break;
                        case 27:
                            customfieldRow.DATE_VALUE = Convert.ToDateTime(inputTaskRow[sourcecolumn]);
                            break;
                        case 17:
                            customfieldRow.FLAG_VALUE = inputTaskRow[sourcecolumn].ToString().Trim().ToUpper() == "YES";
                            break;
                        case 15:
                            customfieldRow.NUM_VALUE = Convert.ToDecimal(inputTaskRow[sourcecolumn]);
                            break;
                        case 21:
                            customfieldRow.TEXT_VALUE = Convert.ToString(inputTaskRow[sourcecolumn]);
                            break;
                    }
                }
                projectDataSet.TaskCustomFields.Rows.Add(customfieldRow);
            }
            Console.WriteLine("build task custom fields done successfully for {0}", inputTaskRow[_mapping.TaskMap["TASK_NAME"]].ToString());
        }

        private static void BuildProjectCustomFields(ProjectDataSet projectDataSet,
          DataRow inputProjectRow, Guid projectGuid)
        {
            Console.WriteLine("build project custom fields started for {0}", inputProjectRow[_mapping.ProjectMap["PROJ_NAME"]].ToString());
            foreach (var customField in _mapping.Project.CustomFields)
            {
                var sourcecolumn = customField.Source;
                if (!inputProjectRow.Table.HasColumn(sourcecolumn))
                    continue;
                bool customFieldExists = false; ;
                if (inputProjectRow[sourcecolumn] == System.DBNull.Value || inputProjectRow[sourcecolumn] == null || string.IsNullOrEmpty(inputProjectRow[sourcecolumn].ToString()))
                {
                    continue;
                }
               

                var targetColumn = customField.Target;
                var csField =
                    customDataSet.CustomFields.First(t => t.MD_PROP_NAME == targetColumn);
                SvcProject.ProjectDataSet.ProjectCustomFieldsRow customfieldRow;
                if (projectDataSet.ProjectCustomFields.Any(t => t.MD_PROP_UID == csField.MD_PROP_UID && t.PROJ_UID == projectGuid))
                {
                    customfieldRow = projectDataSet.ProjectCustomFields.First(t => t.MD_PROP_UID == csField.MD_PROP_UID && t.PROJ_UID == projectGuid);
                    customFieldExists = true;
                }
                else
                {
                    customfieldRow = projectDataSet.ProjectCustomFields.NewProjectCustomFieldsRow();
                    customfieldRow.MD_PROP_ID = csField.MD_PROP_ID;
                    customfieldRow.CUSTOM_FIELD_UID = Guid.NewGuid();
                    customfieldRow.FIELD_TYPE_ENUM = (byte)csField.MD_PROP_TYPE_ENUM;
                    customfieldRow.MD_PROP_ID = csField.MD_PROP_ID;
                    customfieldRow.MD_PROP_UID = csField.MD_PROP_UID;

                    customfieldRow.PROJ_UID = projectGuid;
                }


                //First(t => t.MD_PROP_UID == csField.MD_PROP_UID);
                if (!csField.IsMD_LOOKUP_TABLE_UIDNull())
                {
                    var lookup = Repository.lookupTableClient.ReadLookupTablesByUids(
                        new Guid[] { csField.MD_LOOKUP_TABLE_UID }, false, CultureInfo.CurrentCulture.LCID);
                    customfieldRow.CODE_VALUE =
                        lookup.LookupTableTrees.First(t => t.LT_VALUE_FULL == inputProjectRow[sourcecolumn].ToString())
                            .LT_STRUCT_UID;  //to do add a friendly message 
                }
                else
                {

                    switch (customfieldRow.FIELD_TYPE_ENUM)
                    {
                        case 4:
                            customfieldRow.DATE_VALUE = Convert.ToDateTime(inputProjectRow[sourcecolumn]);
                            break;
                        case 9:
                            customfieldRow.NUM_VALUE = Convert.ToDecimal(inputProjectRow[sourcecolumn]);
                            break;
                        case 6:
                            customfieldRow.DUR_VALUE = Convert.ToInt16(inputProjectRow[sourcecolumn]);
                            break;
                        case 27:
                            customfieldRow.DATE_VALUE = Convert.ToDateTime(inputProjectRow[sourcecolumn]);
                            break;
                        case 17:
                            customfieldRow.FLAG_VALUE = Convert.ToBoolean(inputProjectRow[sourcecolumn]);
                            break;
                        case 15:
                            customfieldRow.NUM_VALUE = Convert.ToDecimal(inputProjectRow[sourcecolumn]);
                            break;
                        case 21:
                            customfieldRow.TEXT_VALUE = Convert.ToString(inputProjectRow[sourcecolumn]);
                            break;
                    }
                }
                if (!customFieldExists)
                {
                    projectDataSet.ProjectCustomFields.Rows.Add(customfieldRow);
                }
            }
            Console.WriteLine("build project custom fields done successfully for {0}", inputProjectRow[_mapping.ProjectMap["PROJ_NAME"]].ToString());
        }
        public static void SetMappingUrl(string mappingFile)
        {
            if (string.IsNullOrEmpty(mappingFile.Trim()))
            {
                _mapping = Mapping.Load();
            }
            else
            {
                _mapping = Mapping.Load(mappingFile);
            }
        }

        public static string[] GetProjectNames(DataSet ds)
        {
            var projectDataSet = new SvcProject.ProjectDataSet();
            string[] projectNames = new string[ds.Tables.Count];
            int count = 0;
            foreach(DataTable table in ds.Tables)
            {
                projectNames[count] = (table.Rows[0][_mapping.ProjectMap[projectDataSet.Project.PROJ_NAMEColumn.ToString()]].ToString()); count++;
            }
            return projectNames;
        }
    }
}

