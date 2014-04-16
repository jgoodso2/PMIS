using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using SvcProject;
using PMSImporter;

namespace PMISImport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            btnImport.Enabled = false;
            txtMapping.Text = new FileInfo(AppDomain.CurrentDomain.BaseDirectory + "/FieldMapping.xml").FullName;
            this.BackColor =  Color.FromArgb(10,57,100);
            this.pictureBox1.BackColor = Color.FromArgb(10, 57, 100); 
            this.txtError.BackColor = Color.FromArgb(10, 57, 100);
            this.ForeColor = Color.White;

            this.btnExit.BackColor = Color.White;
            this.btnImport.BackColor = Color.White;
            this.btnOpenFile.BackColor = Color.White;
            this.button1.BackColor = Color.White;
            this.btnExit.ForeColor = Color.Black;
            this.btnImport.ForeColor = Color.Black; 
            this.btnOpenFile.ForeColor = Color.Black;
            this.button1.ForeColor = Color.Black;
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;
                try
                {

                    txtFileName.Text = file;
                }
                catch (IOException)
                {
                }
            }
            btnImport.Enabled =  EnableImport();
            
        }

       

        private bool EnableImport()
        {
            Uri uri;
            


            return (!string.IsNullOrEmpty(txtMapping.Text) && !string.IsNullOrEmpty(txtFileName.Text)
                 && !string.IsNullOrEmpty(txtProjectServer.Text)
                 && Uri.TryCreate(txtProjectServer.Text, System.UriKind.Absolute, out uri));
            
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            btnExit.Enabled = false;
            txtError.Clear();
            bgw.DoWork -= new DoWorkEventHandler(bgw_DoWork);
            bgw.ProgressChanged -= new ProgressChangedEventHandler(bgw_ProgressChanged);
            bgw.RunWorkerCompleted -= new RunWorkerCompletedEventHandler(bgw_RunWorkerCompleted);
            bgw.DoWork += new DoWorkEventHandler(bgw_DoWork);
            bgw.ProgressChanged += new ProgressChangedEventHandler(bgw_ProgressChanged);
            bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgw_RunWorkerCompleted);
            bgw.WorkerReportsProgress = true;
            Repository.SetProjectServerUrl(txtProjectServer.Text);
            DataSetBuilder.SetMappingUrl(txtMapping.Text);
            bgw.RunWorkerAsync();
        }

        void bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                XLDataSource source = new XLDataSource();
                Repository repository = new Repository();
                List<string> successfulProjects = new List<string>();
                List<string> failedProjects = new List<string>();
                TextBox.CheckForIllegalCrossThreadCalls = false;

                DataSet ds = source.ReadData(txtFileName.Text);
                List<string> errors;
                if (ds.Tables.Count > 0 && !DataSetBuilder.Validate(ds.Tables[0], out errors))
                {

                    txtError.AppendText("Project Import aborted due to validation error" + Environment.NewLine);
                    foreach (string error in errors)
                    {
                        txtError.AppendText(error + Environment.NewLine);
                    }
                    return;
                }
                //For every project in the input file
                for (int i = 0; i < ds.Tables.Count; i++) //some number (total)
                {
                    //System.Threading.Thread.Sleep(100);
                    int percents = ((i) * 100) / ds.Tables.Count;
                    bgw.ReportProgress(percents, new ProjectStatus() { ProjectName = ds.Tables[i].Rows[0][DataSetBuilder.Mapping.ProjectMap["PROJ_NAME"]].ToString(), Status = "Start", SuccessCount = successfulProjects.Count, FailedCount = failedProjects.Count });
                    //2 arguments:
                    //1. procenteges (from 0 t0 100) - i do a calcumation 
                    //2. some current value!
                    ProjectDataSet projectDataSet;
                    ProjectDataSet.ProjectRow row = null;
                    try
                    {
                        DataTable table = ds.Tables[i];
                        bool iscreate, hasPlan;
                        DataSetBuilder.BuildProjectDataSetForCreate(table, out projectDataSet, out row, out iscreate, out hasPlan);
                        DataSetBuilder.BuildProjectDataSetForUpdate(table, projectDataSet, row, iscreate, hasPlan);
                        successfulProjects.Add(row.PROJ_NAME);
                        percents = ((i + 1) * 100) / ds.Tables.Count;
                        bgw.ReportProgress(percents, new ProjectStatus() { ProjectName = ds.Tables[i].Rows[0][DataSetBuilder.Mapping.ProjectMap["PROJ_NAME"]].ToString(), Status = "Complete", SuccessCount = successfulProjects.Count, FailedCount = failedProjects.Count });
                    }
                    catch (Exception ex)
                    {
                        if (row != null)
                        {
                            txtError.AppendText(string.Format("An error occured in import of project {0} .Skipping Project import. Failure reason = {1}", row.PROJ_NAME, ex.Message));
                            failedProjects.Add(row.PROJ_NAME);
                            percents = ((i + 1) * 100) / ds.Tables.Count;
                            bgw.ReportProgress(percents, new ProjectStatus() { ProjectName = ds.Tables[i].Rows[0][DataSetBuilder.Mapping.ProjectMap["PROJ_NAME"]].ToString(), Status = "Fail", SuccessCount = successfulProjects.Count, FailedCount = failedProjects.Count });
                        }
                        else
                        {
                            txtError.AppendText(string.Format("An error occured.Skipping Project import. Failure reason = {0}", ex.Message));
                            failedProjects.Add("");
                            percents = ((i + 1) * 100) / ds.Tables.Count;
                            bgw.ReportProgress(percents, new ProjectStatus() { ProjectName = ds.Tables[i].Rows[0][DataSetBuilder.Mapping.ProjectMap["PROJ_NAME"]].ToString(), Status = "Fail", SuccessCount = successfulProjects.Count, FailedCount = failedProjects.Count });
                        }
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                txtError.AppendText(string.Format("An error occured.Skipping Project import. Failure reason = {0}", ex.Message));
            }
            }

        void bgw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            label1.Text = String.Format("Progress: {0} %", e.ProgressPercentage);
            ProjectStatus status = e.UserState as ProjectStatus;
            if(status.Status == "Start")
            {
                label2.Text = string.Format("Starting Import of Project {0}",status.ProjectName);
            }
            else if(status.Status == "Complete")
            {
                label2.Text = string.Format("Completed Import of Project {0}",status.ProjectName);
            }
            else if(status.Status == "Fail")
            {
                label2.Text = string.Format("Failed Import of Project {0}",status.ProjectName);
            }

            label3.Text = String.Format("Total Projects successfully Imported: {0}",status.SuccessCount);
            label4.Text = String.Format("Total Projects Failed: {0}",status.FailedCount);
        }

        void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //do the code when bgv completes its work
            btnExit.Enabled = true;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog2.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog2.FileName;
                try
                {

                    txtMapping.Text = file;
                }
                catch (IOException)
                {
                }
            }
            btnImport.Enabled = EnableImport();
        }

       

        private void txtProjectServer_Leave(object sender, EventArgs e)
        {
            btnImport.Enabled = EnableImport();
            Uri uri;
            if (string.IsNullOrEmpty(txtProjectServer.Text)
                 || !Uri.TryCreate(txtProjectServer.Text, System.UriKind.Absolute, out uri))
            {
                lblError.Text = "Invalid project server URL format";
            }
            else
            {
                lblError.Text = "";
            }
        }
        }
    }

