using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Xml;
using PSLibrary = Microsoft.Office.Project.Server.Library;

namespace PMSImporter
{
    // Helper methods: WaitForQueue and GetPSClientError.
    class Helpers
    {
        // Wait for the queue jobs to complete.
        public static bool WaitForQueue(SvcQueueSystem.QueueMsgType jobType,
            SvcQueueSystem.QueueSystemClient queueSystemClient,
            DateTime startTime)
        {
            const int MAX_WAIT = 60;    // Maximum wait time, in seconds.
            int numJobs = 1;            // Number of jobs in the queue.
            bool completed = false;     // Queue job completed.
            SvcQueueSystem.QueueStatusDataSet queueStatusDs = 
                new SvcQueueSystem.QueueStatusDataSet();

            int timeout = 0;            // Number of seconds waited.
            Console.Write("Waiting for job: {0} ", jobType.ToString());

            SvcQueueSystem.QueueMsgType[] messageTypes = { jobType };
            SvcQueueSystem.JobState[] jobStates = { SvcQueueSystem.JobState.Success };

            while (timeout < MAX_WAIT)
            {
                System.Threading.Thread.Sleep(1000);    // Sleep one second.

                queueStatusDs = queueSystemClient.ReadMyJobStatus(
                    messageTypes,
                    jobStates,
                    startTime,
                    DateTime.Now,
                    numJobs,
                    true,
                    SvcQueueSystem.SortColumn.QueuePosition,
                    SvcQueueSystem.SortOrder.LastOrder);

                timeout++;
                Console.Write(".");
            }
            Console.WriteLine();

            if (queueStatusDs.Status.Count == numJobs)
                completed = true;
            return completed;
        }

        /// <summary>
        /// Extract a PSClientError object from the ServiceModel.FaultException,
        /// for use in output of the GetPSClientError stack of errors.
        /// </summary>
        /// <param name="e"></param>
        /// <param name="errOut">Shows that FaultException has more information 
        /// about the errors than PSClientError has. FaultException can also contain 
        /// other types of errors, such as failure to connect to the server.</param>
        /// <returns>PSClientError object, for enumerating errors.</returns>
        public static PSLibrary.PSClientError GetPSClientError(FaultException e, 
                                                               out string errOut)
        {
            const string PREFIX = "GetPSClientError() returns null: ";
            errOut = string.Empty;
            PSLibrary.PSClientError psClientError = null;

            if (e == null)
            {
                errOut = PREFIX + "Null parameter (FaultException e) passed in.";
                psClientError = null;
            }
            else
            {
                // Get a ServiceModel.MessageFault object.
                var messageFault = e.CreateMessageFault();

                if (messageFault.HasDetail)
                {
                    using (var xmlReader = messageFault.GetReaderAtDetailContents())
                    {
                        var xml = new XmlDocument();
                        xml.Load(xmlReader);

                        var serverExecutionFault = xml["ServerExecutionFault"];
                        if (serverExecutionFault != null)
                        {
                            var exceptionDetails = serverExecutionFault["ExceptionDetails"];
                            if (exceptionDetails != null)
                            {
                                try
                                {
                                    errOut = exceptionDetails.InnerXml + "\r\n";
                                    psClientError = 
                                        new PSLibrary.PSClientError(exceptionDetails.InnerXml);
                                }
                                catch (InvalidOperationException ex)
                                {
                                    errOut = PREFIX + "Unable to convert fault exception info ";
                                    errOut += "a valid Project Server error message. Message: \n\t";
                                    errOut += ex.Message;
                                    psClientError = null;
                                }
                            }
                            else
                            {
                                errOut = PREFIX + "The FaultException e is a ServerExecutionFault, "
                                    + "but does not have ExceptionDetails.";
                            }
                        }
                        else
                        {
                            errOut = PREFIX + "The FaultException e is not a ServerExecutionFault.";
                        }
                    }
                }
                else // No detail in the MessageFault.
                {
                    errOut = PREFIX + "The FaultException e does not have any detail.";
                }
            }
            errOut += "\r\n" + e.ToString() + "\r\n";
            return psClientError;
        }
    }
}

