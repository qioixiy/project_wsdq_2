namespace ExportExcel
{
    partial class FormMain
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMain));
            this.buttonExportExcel = new System.Windows.Forms.Button();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.openFileDialogData = new System.Windows.Forms.OpenFileDialog();
            this.labelNumber = new System.Windows.Forms.Label();
            this.textBoxNumber = new System.Windows.Forms.TextBox();
            this.dataGridViewEnergy = new System.Windows.Forms.DataGridView();
            this.ColumnDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnPositivePower = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnNegativePower = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnTotalPower = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnDayConsumePower = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnDayFeedPower = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnConsumeTotalPower = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.labelTitle = new System.Windows.Forms.Label();
            this.buttonSelect = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.buttonExit = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewEnergy)).BeginInit();
            this.SuspendLayout();
            // 
            // buttonExportExcel
            // 
            this.buttonExportExcel.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            this.buttonExportExcel.ImageIndex = 0;
            this.buttonExportExcel.ImageList = this.imageList1;
            this.buttonExportExcel.Location = new System.Drawing.Point(952, 420);
            this.buttonExportExcel.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.buttonExportExcel.Name = "buttonExportExcel";
            this.buttonExportExcel.Size = new System.Drawing.Size(131, 34);
            this.buttonExportExcel.TabIndex = 0;
            this.buttonExportExcel.Text = "  导出为Excel";
            this.buttonExportExcel.UseVisualStyleBackColor = true;
            this.buttonExportExcel.Click += new System.EventHandler(this.buttonExportExcel_Click);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "Excel.ico");
            this.imageList1.Images.SetKeyName(1, "document.ico");
            this.imageList1.Images.SetKeyName(2, "Exit.ico");
            // 
            // labelNumber
            // 
            this.labelNumber.AutoSize = true;
            this.labelNumber.Location = new System.Drawing.Point(948, 46);
            this.labelNumber.Name = "labelNumber";
            this.labelNumber.Size = new System.Drawing.Size(35, 17);
            this.labelNumber.TabIndex = 1;
            this.labelNumber.Text = "车号:";
            // 
            // textBoxNumber
            // 
            this.textBoxNumber.Location = new System.Drawing.Point(988, 43);
            this.textBoxNumber.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.textBoxNumber.Name = "textBoxNumber";
            this.textBoxNumber.Size = new System.Drawing.Size(85, 23);
            this.textBoxNumber.TabIndex = 2;
            this.textBoxNumber.Text = "1571";
            // 
            // dataGridViewEnergy
            // 
            this.dataGridViewEnergy.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridViewEnergy.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewEnergy.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColumnDate,
            this.ColumnPositivePower,
            this.ColumnNegativePower,
            this.ColumnTotalPower,
            this.ColumnDayConsumePower,
            this.ColumnDayFeedPower,
            this.ColumnConsumeTotalPower});
            this.dataGridViewEnergy.Location = new System.Drawing.Point(0, 43);
            this.dataGridViewEnergy.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.dataGridViewEnergy.Name = "dataGridViewEnergy";
            this.dataGridViewEnergy.RowTemplate.Height = 23;
            this.dataGridViewEnergy.Size = new System.Drawing.Size(941, 531);
            this.dataGridViewEnergy.TabIndex = 3;
            this.dataGridViewEnergy.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dataGridViewEnergy_CellPainting);
            // 
            // ColumnDate
            // 
            this.ColumnDate.HeaderText = "日期";
            this.ColumnDate.Name = "ColumnDate";
            this.ColumnDate.ReadOnly = true;
            // 
            // ColumnPositivePower
            // 
            this.ColumnPositivePower.HeaderText = "正向电能";
            this.ColumnPositivePower.Name = "ColumnPositivePower";
            this.ColumnPositivePower.ReadOnly = true;
            // 
            // ColumnNegativePower
            // 
            this.ColumnNegativePower.HeaderText = "反向电能";
            this.ColumnNegativePower.Name = "ColumnNegativePower";
            this.ColumnNegativePower.ReadOnly = true;
            // 
            // ColumnTotalPower
            // 
            this.ColumnTotalPower.HeaderText = "总电能";
            this.ColumnTotalPower.Name = "ColumnTotalPower";
            this.ColumnTotalPower.ReadOnly = true;
            // 
            // ColumnDayConsumePower
            // 
            this.ColumnDayConsumePower.HeaderText = "日耗正向电能";
            this.ColumnDayConsumePower.Name = "ColumnDayConsumePower";
            this.ColumnDayConsumePower.ReadOnly = true;
            // 
            // ColumnDayFeedPower
            // 
            this.ColumnDayFeedPower.HeaderText = "日馈反向电能";
            this.ColumnDayFeedPower.Name = "ColumnDayFeedPower";
            this.ColumnDayFeedPower.ReadOnly = true;
            // 
            // ColumnConsumeTotalPower
            // 
            this.ColumnConsumeTotalPower.HeaderText = "单耗总电能";
            this.ColumnConsumeTotalPower.Name = "ColumnConsumeTotalPower";
            this.ColumnConsumeTotalPower.ReadOnly = true;
            // 
            // labelTitle
            // 
            this.labelTitle.AutoSize = true;
            this.labelTitle.Location = new System.Drawing.Point(476, 12);
            this.labelTitle.Name = "labelTitle";
            this.labelTitle.Size = new System.Drawing.Size(56, 17);
            this.labelTitle.TabIndex = 4;
            this.labelTitle.Text = "电能数据";
            this.labelTitle.Click += new System.EventHandler(this.labelTitle_Click);
            // 
            // buttonSelect
            // 
            this.buttonSelect.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            this.buttonSelect.ImageIndex = 1;
            this.buttonSelect.ImageList = this.imageList1;
            this.buttonSelect.Location = new System.Drawing.Point(951, 486);
            this.buttonSelect.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.buttonSelect.Name = "buttonSelect";
            this.buttonSelect.Size = new System.Drawing.Size(131, 33);
            this.buttonSelect.TabIndex = 5;
            this.buttonSelect.Text = "选择文件";
            this.buttonSelect.UseVisualStyleBackColor = true;
            this.buttonSelect.Click += new System.EventHandler(this.buttonSelect_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // buttonExit
            // 
            this.buttonExit.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            this.buttonExit.ImageIndex = 2;
            this.buttonExit.ImageList = this.imageList1;
            this.buttonExit.Location = new System.Drawing.Point(951, 541);
            this.buttonExit.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.buttonExit.Name = "buttonExit";
            this.buttonExit.Size = new System.Drawing.Size(131, 33);
            this.buttonExit.TabIndex = 6;
            this.buttonExit.Text = "退出";
            this.buttonExit.UseVisualStyleBackColor = true;
            this.buttonExit.Click += new System.EventHandler(this.buttonExit_Click);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(1085, 576);
            this.Controls.Add(this.buttonExit);
            this.Controls.Add(this.buttonSelect);
            this.Controls.Add(this.labelTitle);
            this.Controls.Add(this.dataGridViewEnergy);
            this.Controls.Add(this.textBoxNumber);
            this.Controls.Add(this.labelNumber);
            this.Controls.Add(this.buttonExportExcel);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.Name = "FormMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "苏州市万松电器有限公司";
            this.Shown += new System.EventHandler(this.FormMain_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewEnergy)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonExportExcel;
        private System.Windows.Forms.OpenFileDialog openFileDialogData;

        private System.Windows.Forms.Label labelNumber;
        private System.Windows.Forms.TextBox textBoxNumber;
        private System.Windows.Forms.DataGridView dataGridViewEnergy;
        private System.Windows.Forms.Label labelTitle;
        private System.Windows.Forms.Button buttonSelect;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button buttonExit;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnDate;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnPositivePower;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnNegativePower;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnTotalPower;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnDayConsumePower;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnDayFeedPower;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnConsumeTotalPower;
        private System.Windows.Forms.ImageList imageList1;
    }
}

