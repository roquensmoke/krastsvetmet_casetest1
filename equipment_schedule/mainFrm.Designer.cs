namespace equipment_schedule
{
    partial class frmMain
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
            this.btnReadFiles = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.btnCalc = new System.Windows.Forms.Button();
            this.pbReport = new System.Windows.Forms.PictureBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btnExcel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pbReport)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnReadFiles
            // 
            this.btnReadFiles.Location = new System.Drawing.Point(12, 12);
            this.btnReadFiles.Name = "btnReadFiles";
            this.btnReadFiles.Size = new System.Drawing.Size(140, 35);
            this.btnReadFiles.TabIndex = 0;
            this.btnReadFiles.Text = "Прочитать входные параметры...";
            this.btnReadFiles.UseVisualStyleBackColor = true;
            this.btnReadFiles.Click += new System.EventHandler(this.btnReadFiles_Click);
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.Description = "Выберите расположение входных файлов";
            this.folderBrowserDialog1.ShowNewFolderButton = false;
            // 
            // btnCalc
            // 
            this.btnCalc.Enabled = false;
            this.btnCalc.Location = new System.Drawing.Point(158, 12);
            this.btnCalc.Name = "btnCalc";
            this.btnCalc.Size = new System.Drawing.Size(121, 35);
            this.btnCalc.TabIndex = 2;
            this.btnCalc.Text = "Расчет";
            this.btnCalc.UseVisualStyleBackColor = true;
            this.btnCalc.Click += new System.EventHandler(this.btnCalc_Click);
            // 
            // pbReport
            // 
            this.pbReport.Location = new System.Drawing.Point(12, 53);
            this.pbReport.Name = "pbReport";
            this.pbReport.Size = new System.Drawing.Size(1033, 225);
            this.pbReport.TabIndex = 3;
            this.pbReport.TabStop = false;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 284);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(1033, 236);
            this.dataGridView1.TabIndex = 5;
            // 
            // btnExcel
            // 
            this.btnExcel.Enabled = false;
            this.btnExcel.Location = new System.Drawing.Point(285, 12);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(121, 35);
            this.btnExcel.TabIndex = 2;
            this.btnExcel.Text = "Показать файл Excel";
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1057, 289);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.pbReport);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.btnCalc);
            this.Controls.Add(this.btnReadFiles);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "frmMain";
            this.Text = "Расчет расписания работы оборудования";
            this.Load += new System.EventHandler(this.frmMain_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pbReport)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnReadFiles;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button btnCalc;
        private System.Windows.Forms.PictureBox pbReport;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btnExcel;
    }
}

