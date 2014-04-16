using System;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Security.Principal;
using System.ServiceModel;
using System.Xml;
using SvcProject;
using SvcResource;
using System.Collections.Generic;
using PSLib = Microsoft.Office.Project.Server.Library;

namespace PMSImporter
{
    public class Repository
    {
        public static SvcAdmin.AdminClient adminClient;
        public static SvcQueueSystem.QueueSystemClient queueSystemClient;
        public static SvcResource.ResourceClient resourceClient;
        public static SvcProject.ProjectClient projectClient;
        public static SvcLookupTable.LookupTableClient lookupTableClient;
        public static SvcCustomFields.CustomFieldsClient customFieldsClient;
        public static SvcCalendar.CalendarClient calendarClient;
        public static SvcArchive.ArchiveClient archiveClient;
        public static SvcStatusing.StatusingClient pwaClient;
        public static SvcTimeSheet.TimeSheetClient timesheetClient;
        public static SvcQueueSystem.QueueSystemClient queueClient;
        public static SvcWorkflow.WorkflowClient workFlowClient;
        private static string pwaUrl;
        private static ResourceDataSet _resourceList;
        private const int NO_QUEUE_MESSAGE = -1;


        public static void SaveProject(Guid guid)
        {
            projectClient.QueuePublish(Guid.NewGuid(), guid, true, "");
        }
        public Repository()
        {

        }

        static Repository()
        {
            
        }


        public static ProjectDataSet ReadProject(Guid projectGuid)
        {
            Console.WriteLine("Read project started for {0}", projectGuid);
            return projectClient.ReadProject(projectGuid, DataStoreEnum.WorkingStore);
            Console.WriteLine("Read project done successfully for {0}", projectGuid);
        }
        // Set the PSI client endpoints programmatically; don't use app.config.
        public static bool SetClientEndpointsProg(string pwaUrl)
        {
            const int MAXSIZE = int.MaxValue;
            const string SVC_ROUTER = "/_vti_bin/PSI/ProjectServer.svc";

            bool isHttps = pwaUrl.ToLower().StartsWith("https");
            bool result = true;
            BasicHttpBinding binding = null;

            try
            {
                if (isHttps)
                {
                    // Create a binding for HTTPS.TimesheetL
                    binding = new BasicHttpBinding(BasicHttpSecurityMode.Transport);
                }
                else
                {
                    // Create a binding for HTTP.
                    binding = new BasicHttpBinding(BasicHttpSecurityMode.TransportCredentialOnly);
                }

                binding.Name = "basicHttpConf";
                binding.MessageEncoding = WSMessageEncoding.Text;

                binding.CloseTimeout = new TimeSpan(00, 05, 00);
                binding.OpenTimeout = new TimeSpan(00, 05, 00);
                binding.ReceiveTimeout = new TimeSpan(00, 05, 00);
                binding.SendTimeout = new TimeSpan(00, 05, 00);
                binding.TextEncoding = System.Text.Encoding.UTF8;

                // If the TransferMode is buffered, the MaxBufferSize and 
                // MaxReceived MessageSize must be the same value.
                binding.TransferMode = TransferMode.Buffered;
                binding.MaxBufferSize = MAXSIZE;
                binding.MaxReceivedMessageSize = MAXSIZE;
                binding.MaxBufferPoolSize = MAXSIZE;


                binding.Security.Transport.ClientCredentialType = HttpClientCredentialType.Ntlm;
                binding.GetType().GetProperty("ReaderQuotas").SetValue(binding, XmlDictionaryReaderQuotas.Max, null);
                // The endpoint address is the ProjectServer.svc router for all public PSI calls.
                EndpointAddress address = new EndpointAddress(pwaUrl + SVC_ROUTER);



                adminClient = new SvcAdmin.AdminClient(binding, address);
                adminClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                adminClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;


                projectClient = new SvcProject.ProjectClient(binding, address);
                projectClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                projectClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;

                queueSystemClient = new SvcQueueSystem.QueueSystemClient(binding, address);
                queueSystemClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                queueSystemClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;

                resourceClient = new SvcResource.ResourceClient(binding, address);
                resourceClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                resourceClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;

                lookupTableClient = new SvcLookupTable.LookupTableClient(binding, address);
                lookupTableClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                lookupTableClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;


                customFieldsClient = new SvcCustomFields.CustomFieldsClient(binding, address);
                customFieldsClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                customFieldsClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;

                calendarClient = new SvcCalendar.CalendarClient(binding, address);
                calendarClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                calendarClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;

                archiveClient = new SvcArchive.ArchiveClient(binding, address);
                archiveClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                archiveClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;

                pwaClient = new SvcStatusing.StatusingClient(binding, address);
                pwaClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                pwaClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;

                timesheetClient = new SvcTimeSheet.TimeSheetClient(binding, address);
                timesheetClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                timesheetClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;

                queueClient = new SvcQueueSystem.QueueSystemClient(binding, address);
                queueClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                queueClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;

                workFlowClient = new SvcWorkflow.WorkflowClient(binding, address);
                workFlowClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                workFlowClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;
            }
            catch (Exception ex)
            {
                result = false;
            }
            return result;
        }

        public static Guid GetUIDFromEPTUID(Guid guid)
        {
            var enterpriseTypes = workFlowClient.ReadAvailableEnterpriseProjectTypes();
            if(enterpriseTypes.EnterpriseProjectType.Any(t=>t.ENTERPRISE_PROJECT_TYPE_UID == guid && !t.IsENTERPRISE_PROJECT_PLAN_TEMPLATE_UIDNull()))
            {
                var type = enterpriseTypes.EnterpriseProjectType.First(t => t.ENTERPRISE_PROJECT_TYPE_UID == guid && !t.IsENTERPRISE_PROJECT_PLAN_TEMPLATE_UIDNull());
                return enterpriseTypes.EnterpriseProjectType.First(t => t.ENTERPRISE_PROJECT_TYPE_UID == guid && !t.IsENTERPRISE_PROJECT_PLAN_TEMPLATE_UIDNull()).ENTERPRISE_PROJECT_PLAN_TEMPLATE_UID;
            }
            return Guid.Empty;
        }
        public Guid CreateProject(ProjectDataSet projectDs,bool hasPlan,Guid planUID)
        {
            string projectName = projectDs.Project[0].PROJ_NAME;

            Console.WriteLine("Create project started for {0}", projectName);
            Guid jobGuid = Guid.NewGuid();
            //if (hasPlan)
            //{
            //    var projID = projectClient.CreateProjectFromTemplate(planUID, projectName);
            //    projectClient.QueuePublish(jobGuid, projID, true, string.Empty);
            //    if (WaitForQueueJobCompletion(jobGuid, Guid.NewGuid(), (int)SvcQueueSystem.QueueMsgType.ProjectPublish
            //  ))
            //    {
            //        Console.WriteLine("Create project done successfully for {0}", projectName);
            //        projectDs = projectClient.ReadProject(projID,DataStoreEnum.PublishedStore);
            //    }
            //    else
            //    {
            //        Console.WriteLine("The project was not created due to queue error for : {0}", projectName);
            //    }
            //    return projID;
            //}
            //else
            //{
                projectClient.QueueCreateProject(jobGuid, projectDs, false);
                // Wait for the Project Server Queuing System to create the project.
                if (WaitForQueueJobCompletion(jobGuid, Guid.NewGuid(), (int)SvcQueueSystem.QueueMsgType.ProjectCreate
                    ))
                {

                    jobGuid = Guid.NewGuid();

                }
                else
                {
                    Console.WriteLine("The project was not created due to queue error for : {0}", projectName);
                }
                return projectDs.Project[0].PROJ_UID;
            //}
        }

        public void UpdateProjectTeam(ProjectDataSet projectDs)
        {
            Guid projUid = projectDs.Project[0].PROJ_UID;
            string projectName = projectDs.Project[0].PROJ_NAME;
            Guid sessionId = Guid.NewGuid();
            Console.WriteLine("update project team started for {0}", projectName);
            //Read from the server the team that is present on the server for the project
            var projectTeam = projectClient.ReadProjectTeam(projUid);
            //For every resource read from the input file add to thye team if not already existing within the team
            foreach (var resource in projectDs.ProjectResource)
            {
                if (!projectTeam.ProjectTeam.Any(t => t.RES_NAME == resource.RES_NAME))
                {
                    var projectTeamRow = projectTeam.ProjectTeam.NewProjectTeamRow();
                    projectTeamRow.RES_UID = resource.RES_UID;
                    projectTeamRow.RES_NAME = resource.RES_NAME;
                    projectTeamRow.PROJ_UID = projUid;
                    projectTeamRow.NEW_RES_UID = resource.RES_UID;
                    projectTeam.ProjectTeam.Rows.Add(projectTeamRow);
                }
            }
            ProjectTeamDataSet teamDelta = new ProjectTeamDataSet();
            if (projectTeam.GetChanges(DataRowState.Added) != null)
            {
                projectClient.CheckOutProject(projUid, sessionId, "");
                teamDelta.Merge(projectTeam.GetChanges(DataRowState.Added));
                Guid jobGuid = Guid.NewGuid();
                projectClient.QueueUpdateProjectTeam(jobGuid, sessionId, projUid, teamDelta);
                if (WaitForQueueJobCompletion(jobGuid, Guid.NewGuid(), (int)SvcQueueSystem.QueueMsgType.ProjectUpdateTeam))
                {
                    jobGuid = Guid.NewGuid();
                    projectClient.QueueCheckInProject(jobGuid, projUid, true, Guid.NewGuid(), "");
                    if (WaitForQueueJobCompletion(jobGuid, Guid.NewGuid(), (int)SvcQueueSystem.QueueMsgType.ProjectCheckIn))
                    {
                        Console.WriteLine("update project team done successfully for {0}", projectName);
                    }
                    else
                    {
                        Console.WriteLine(
                                       "update project team done queue error for {0}", projectName);
                    }

                }
                else
                {
                    jobGuid = Guid.NewGuid();
                    projectClient.QueueCheckInProject(jobGuid, projUid, true, Guid.NewGuid(), "");
                    if (WaitForQueueJobCompletion(jobGuid, Guid.NewGuid(), (int)SvcQueueSystem.QueueMsgType.ProjectCheckIn))
                    {
                        Console.WriteLine("update project team done successfully for {0}", projectName);
                    }
                    else
                    {
                        Console.WriteLine(
                                       "update project team done queue error for {0}", projectName);
                    }
                    Console.WriteLine(
                                        "update project team done queue error for {0}", projectName);
                }
            }
            else
            {
                Console.WriteLine("update project team done successfully for {0}", projectName);
            }

        }
        public void UpdateProject(ProjectDataSet projectDs)
        {
            Guid projUid = projectDs.Project[0].PROJ_UID;
            Guid sessionId = Guid.NewGuid();
            string projectName = projectDs.Project[0].PROJ_NAME;
            Console.WriteLine("update project team started for {0}", projectName);
            ProjectDataSet deltaDataSet = new ProjectDataSet();
            if (projectDs.GetChanges() != null)
            {
                
                if (projectDs.GetChanges(DataRowState.Added) != null)
                {
                    deltaDataSet.Merge(projectDs.GetChanges(DataRowState.Added));
                    projectClient.CheckOutProject(projUid, sessionId, "");
                    Guid jobGuid = Guid.NewGuid();
                    projectClient.QueueAddToProject(jobGuid, sessionId, deltaDataSet, false);
                    // Wait for the Project Server Queuing System to create the project.
                    if (WaitForQueueJobCompletion(jobGuid, Guid.NewGuid(), (int)SvcQueueSystem.QueueMsgType.ProjectUpdate))
                    {
                        jobGuid = Guid.NewGuid();
                        projectClient.QueueCheckInProject(jobGuid, projUid, true, Guid.NewGuid(), "");
                        if (WaitForQueueJobCompletion(jobGuid, projUid, (int)SvcQueueSystem.QueueMsgType.ProjectCheckIn))
                        {
                            projectClient.QueuePublish(Guid.NewGuid(), projUid, false, pwaUrl);
                            if (WaitForQueueJobCompletion(jobGuid, Guid.NewGuid(), (int)SvcQueueSystem.QueueMsgType.ProjectPublish))
                            {
                                Console.WriteLine("update project done successfully for {0}", projectName);
                            }
                            else
                            {
                                Console.WriteLine(
                                           "update project done queue error for {0}", projectName);
                            }
                        }
                        else
                        {
                            Console.WriteLine(
                                            "update project  done queue error for {0}", projectName);
                        }

                    }
                    else
                    {
                        Console.WriteLine(
                                           "update project done queue error for {0}", projectName);
                    }
                }
                else
                {
                    Console.WriteLine("update project done successfully for {0}", projectName);
                }
                
                 if (projectDs.GetChanges(DataRowState.Modified) != null)
                {
                    deltaDataSet = new ProjectDataSet();
                    deltaDataSet.Merge(projectDs.GetChanges(DataRowState.Modified));
                    projectClient.CheckOutProject(projUid, sessionId, "");
                    Guid jobGuid = Guid.NewGuid();
                    projectClient.QueueUpdateProject(jobGuid, sessionId, deltaDataSet, false);
                    // Wait for the Project Server Queuing System to create the project.
                    if (WaitForQueueJobCompletion(jobGuid, Guid.NewGuid(), (int)SvcQueueSystem.QueueMsgType.ProjectUpdate))
                    {
                        jobGuid = Guid.NewGuid();
                        projectClient.QueueCheckInProject(jobGuid, projUid, true, Guid.NewGuid(), "");
                        if (WaitForQueueJobCompletion(jobGuid, projUid, (int)SvcQueueSystem.QueueMsgType.ProjectCheckIn))
                        {

                            projectClient.QueuePublish(Guid.NewGuid(), projUid, false, pwaUrl);
                            if (WaitForQueueJobCompletion(jobGuid, Guid.NewGuid(), (int)SvcQueueSystem.QueueMsgType.ProjectPublish))
                            {
                                Console.WriteLine("update project done successfully for {0}", projectName);
                            }
                            else
                            {
                                Console.WriteLine(
                                           "update project done queue error for {0}", projectName);
                            }
                        }
                        else
                        {
                            Console.WriteLine(
                                            "update project  done queue error for {0}", projectName);
                        }

                    }
                    else
                    {
                        Console.WriteLine(
                                           "update project done queue error for {0}", projectName);
                    }
                }
                else
                {
                    Console.WriteLine("update project done successfully for {0}", projectName);
                }

            }
        }

        public static ResourceDataSet GetResources()
        {
            Console.WriteLine("Get all Resources called");
            if (_resourceList == null)
                _resourceList = resourceClient.ReadUserList(ResourceActiveFilter.All);
            Console.WriteLine("Get all Resources done successfully");
            return _resourceList;
        }

        public static Guid GetResourceUidFromNtAccount(String ntAccount, out bool isWindowsUser)
        {
            SvcResource.ResourceDataSet rds = new SvcResource.ResourceDataSet();

            Microsoft.Office.Project.Server.Library.Filter filter = new Microsoft.Office.Project.Server.Library.Filter();
            filter.FilterTableName = rds.Resources.TableName;


            Microsoft.Office.Project.Server.Library.Filter.Field ntAccountField1 = new Microsoft.Office.Project.Server.Library.Filter.Field(rds.Resources.TableName, rds.Resources.WRES_ACCOUNTColumn.ColumnName);
            filter.Fields.Add(ntAccountField1);

            Microsoft.Office.Project.Server.Library.Filter.Field ntAccountField2 = new Microsoft.Office.Project.Server.Library.Filter.Field(rds.Resources.TableName, rds.Resources.RES_IS_WINDOWS_USERColumn.ColumnName);
            filter.Fields.Add(ntAccountField2);

            Microsoft.Office.Project.Server.Library.Filter.FieldOperator op = new Microsoft.Office.Project.Server.Library.Filter.FieldOperator(Microsoft.Office.Project.Server.Library.Filter.FieldOperationType.Equal,
                rds.Resources.WRES_ACCOUNTColumn.ColumnName, ntAccount);
            filter.Criteria = op;



            rds = resourceClient.ReadResources(filter.GetXml(), false);

            isWindowsUser = rds.Resources[0].RES_IS_WINDOWS_USER;

            var obj = (Guid)rds.Resources.Rows[0]["RES_UID"];
            return obj;
        }

        private static void ShowErrorList(string errorType, int jobState, List<int> errorList)
        {
            string msg = "Errors in the " + errorType + ":";
            string enumName = string.Empty;

            foreach (int errorNum in errorList)
            {
                msg += "\n\t" + errorNum.ToString();
                enumName = Enum.GetName(typeof(PSLib.PSErrorID), errorNum);
                msg += ": " + enumName;
            }
            msg += "\n\nQueue job state: ";

            switch (jobState)
            {
                case (int)SvcQueueSystem.JobState.CorrelationBlocked:
                    msg += "Correlation blocked";
                    break;
                case (int)SvcQueueSystem.JobState.Failed:
                    msg += "Failed";
                    break;
                default:
                    msg += "OK";
                    break;
            }
            Console.Write(msg, errorType + " Errors");
        }
        /// <summary>
        /// Wait for the queue job completion of the specified message type.
        /// </summary>
        /// <param name="trackingGuid">Tracking GUID, used for getting the job group wait time. 
        ///    Not used in this implementation.</param>
        /// <param name="messageType">Type of queue message, specified by SvcQueueSystemQueueMsgType.
        ///    BugBug: using -1 for now.</param>
        /// <returns></returns>
        public static bool WaitForQueueJobCompletion(Guid jobGUID, Guid trackingGuid, int messageType)
        {
            SvcQueueSystem.QueueStatusDataSet queueStatusDataSet = new SvcQueueSystem.QueueStatusDataSet();
            SvcQueueSystem.QueueStatusRequestDataSet queueStatusRequestDataSet =
                new SvcQueueSystem.QueueStatusRequestDataSet();

            SvcQueueSystem.QueueStatusRequestDataSet.StatusRequestRow statusRequestRow =
                queueStatusRequestDataSet.StatusRequest.NewStatusRequestRow();
            statusRequestRow.JobGUID = jobGUID;
            statusRequestRow.JobGroupGUID = Guid.NewGuid();
            statusRequestRow.MessageType = messageType;
            queueStatusRequestDataSet.StatusRequest.AddStatusRequestRow(statusRequestRow);

            bool inProcess = true;
            bool result = false;
            DateTime startTime = DateTime.Now;
            int successState = (int)SvcQueueSystem.JobState.Success;
            int failedState = (int)SvcQueueSystem.JobState.Failed;
            int blockedState = (int)SvcQueueSystem.JobState.CorrelationBlocked;

            List<int> errorList = new List<int>();

            using (OperationContextScope scope = new OperationContextScope(queueSystemClient.InnerChannel))
            {

                while (inProcess)
                {
                    queueStatusDataSet = queueSystemClient.ReadJobStatus(queueStatusRequestDataSet, false,
                        SvcQueueSystem.SortColumn.Undefined, SvcQueueSystem.SortOrder.Undefined);

                    foreach (SvcQueueSystem.QueueStatusDataSet.StatusRow statusRow in queueStatusDataSet.Status)
                    {
                        if (statusRow["ErrorInfo"] != System.DBNull.Value)
                        {
                            errorList = CheckStatusRowErrors(statusRow["ErrorInfo"].ToString());

                            if (errorList.Count > 0
                                || statusRow.JobCompletionState == blockedState
                                || statusRow.JobCompletionState == failedState)
                            {
                                inProcess = false;
                                ShowErrorList("Queue", statusRow.JobCompletionState, errorList);
                            }
                        }
                        if (statusRow.JobCompletionState == successState)
                        {
                            inProcess = false;
                            result = true;
                        }
                        else
                        {
                            inProcess = true;
                            System.Threading.Thread.Sleep(500);  // Sleep 1/2 second.
                        }
                    }
                    DateTime endTime = DateTime.Now;
                    TimeSpan span = endTime.Subtract(startTime);

                    if (span.Seconds > 20) //Wait for only 20 secs - and then bail out.
                    {
                        Console.Write("Something is wrong in the queue. Check or clear the queue and then redo your action.");
                        inProcess = false;
                        result = false;
                    }
                }
            }
            return result;
        }

        public static List<int> CheckStatusRowErrors(string errorInfo)
        {
            List<int> errorList = new List<int>();
            bool containsError = false;

            XmlTextReader xReader = new XmlTextReader(new System.IO.StringReader(errorInfo));
            while (xReader.Read())
            {
                if (xReader.Name == "errinfo" && xReader.NodeType == XmlNodeType.Element)
                {
                    xReader.Read();
                    if (xReader.Value != string.Empty)
                    {
                        containsError = true;
                    }
                }
                if (containsError && xReader.Name == "error" && xReader.NodeType == XmlNodeType.Element)
                {
                    while (xReader.Read())
                    {
                        if (xReader.Name == "id" && xReader.NodeType == XmlNodeType.Attribute)
                        {
                            errorList.Add(Convert.ToInt32(xReader.Value));
                        }
                    }
                }
            }
            return errorList;
        }

        public static ProjectDataSet GetProjectList()
        {
            return projectClient.ReadProjectList();
        }


        public static bool CheckIfProjectExists(string projectName)
        {
            return GetProjectList().Project.Any(t => t.PROJ_NAME.Trim().ToUpper() == projectName.Trim().ToUpper());
        }

        public static void SetProjectServerUrl(string projectServerURL)
        {
            pwaUrl = projectServerURL;
            SetClientEndpointsProg(pwaUrl);
        }


       
    }


}
