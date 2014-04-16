namespace PMISImport
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btnOpenFile = new System.Windows.Forms.Button();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.btnImport = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.bgw = new System.ComponentModel.BackgroundWorker();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtError = new System.Windows.Forms.RichTextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.txtProjectServer = new System.Windows.Forms.TextBox();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.lblError = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label8 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.Location = new System.Drawing.Point(507, 118);
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(106, 23);
            this.btnOpenFile.TabIndex = 0;
            this.btnOpenFile.Text = "Browse";
            this.btnOpenFile.UseVisualStyleBackColor = true;
            this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
            // 
            // txtFileName
            // 
            this.txtFileName.Location = new System.Drawing.Point(129, 118);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.ReadOnly = true;
            this.txtFileName.Size = new System.Drawing.Size(362, 21);
            this.txtFileName.TabIndex = 1;
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(162, 163);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(168, 23);
            this.btnImport.TabIndex = 2;
            this.btnImport.Text = "&Import";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(370, 163);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(168, 23);
            this.btnExit.TabIndex = 3;
            this.btnExit.Text = "E&xit";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(1, 289);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(631, 23);
            this.label1.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(1, 237);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(631, 23);
            this.label2.TabIndex = 5;
            // 
            // progressBar1
            // 
            this.progressBar1.ForeColor = System.Drawing.Color.Green;
            this.progressBar1.Location = new System.Drawing.Point(-3, 354);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(979, 23);
            this.progressBar1.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(1, 212);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(631, 23);
            this.label3.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(1, 266);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(631, 23);
            this.label4.TabIndex = 8;
            // 
            // txtError
            // 
            this.txtError.BackColor = System.Drawing.Color.White;
            this.txtError.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtError.ForeColor = System.Drawing.Color.Gainsboro;
            this.txtError.Location = new System.Drawing.Point(621, 79);
            this.txtError.Name = "txtError";
            this.txtError.ReadOnly = true;
            this.txtError.Size = new System.Drawing.Size(355, 269);
            this.txtError.TabIndex = 9;
            this.txtError.Text = "";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(9, 121);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(60, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Input File";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(9, 35);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(116, 13);
            this.label7.TabIndex = 14;
            this.label7.Text = "Project Server URL";
            // 
            // txtProjectServer
            // 
            this.txtProjectServer.Location = new System.Drawing.Point(129, 32);
            this.txtProjectServer.Name = "txtProjectServer";
            this.txtProjectServer.Size = new System.Drawing.Size(441, 21);
            this.txtProjectServer.TabIndex = 15;
            this.txtProjectServer.Text = "https://perview.app.medcity.net/pwa";
            this.txtProjectServer.Leave += new System.EventHandler(this.txtProjectServer_Leave);
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "openFileDialog1";
            // 
            // lblError
            // 
            this.lblError.AutoSize = true;
            this.lblError.ForeColor = System.Drawing.Color.Red;
            this.lblError.Location = new System.Drawing.Point(642, 9);
            this.lblError.Name = "lblError";
            this.lblError.Size = new System.Drawing.Size(0, 13);
            this.lblError.TabIndex = 16;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::PMISImport.Properties.Resources.logo2;
            this.pictureBox1.Location = new System.Drawing.Point(621, -14);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(355, 71);
            this.pictureBox1.TabIndex = 17;
            this.pictureBox1.TabStop = false;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(8, 95);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(603, 13);
            this.label8.TabIndex = 18;
            this.label8.Text = "Use the mapping file to define how input file columns correspond to built-in and " +
    "custom fields in PerView";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.Moccasin;
            this.panel1.Location = new System.Drawing.Point(-2, 66);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(984, 10);
            this.panel1.TabIndex = 19;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(974, 377);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.lblError);
            this.Controls.Add(this.txtProjectServer);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtError);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.txtFileName);
            this.Controls.Add(this.btnOpenFile);
            this.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "PerView Bulk Edit Tool";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOpenFile;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.ComponentModel.BackgroundWorker bgw;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.RichTextBox txtError;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtProjectServer;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.Label lblError;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Panel panel1;
    }
}

