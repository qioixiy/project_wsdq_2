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
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.openFileDialogData = new System.Windows.Forms.OpenFileDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.labelTitle = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.comboBoxSelectDateTime = new System.Windows.Forms.ComboBox();
            this.textBox_V_type = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.buttonExit = new System.Windows.Forms.Button();
            this.buttonSelect = new System.Windows.Forms.Button();
            this.textBoxNumber = new System.Windows.Forms.TextBox();
            this.labelNumber = new System.Windows.Forms.Label();
            this.buttonExportExcel = new System.Windows.Forms.Button();
            this.dataGridViewEnergy = new System.Windows.Forms.DataGridView();
            this.ColumnDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnPositivePower = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnNegativePower = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnTotalPower = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnDayConsumePower = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnDayFeedPower = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnConsumeTotalPower = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnV = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnI = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnPower = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnRealTimePower = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnV3rd = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnV5rd = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnV7rd = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnV9rd = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnI3rd = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnI5rd = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnI7rd = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnI9rd = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewEnergy)).BeginInit();
            this.SuspendLayout();
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "Excel.ico");
            this.imageList1.Images.SetKeyName(1, "document.ico");
            this.imageList1.Images.SetKeyName(2, "Exit.ico");
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.labelTitle);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1278, 43);
            this.panel1.TabIndex = 7;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arial", 17F, System.Drawing.FontStyle.Bold);
            this.label3.Location = new System.Drawing.Point(807, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 27);
            this.label3.TabIndex = 7;
            this.label3.Text = "V1.2";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial", 17F, System.Drawing.FontStyle.Bold);
            this.label2.Location = new System.Drawing.Point(325, 10);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(76, 27);
            this.label2.TabIndex = 6;
            this.label2.Text = "EEMS";
            // 
            // labelTitle
            // 
            this.labelTitle.AutoSize = true;
            this.labelTitle.Font = new System.Drawing.Font("隶书", 23F, System.Drawing.FontStyle.Bold);
            this.labelTitle.Location = new System.Drawing.Point(392, 6);
            this.labelTitle.Name = "labelTitle";
            this.labelTitle.Size = new System.Drawing.Size(410, 31);
            this.labelTitle.TabIndex = 5;
            this.labelTitle.Text = "电能监控系统数据处理助手";
            this.labelTitle.Click += new System.EventHandler(this.labelTitle_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.panel3);
            this.panel2.Controls.Add(this.dataGridViewEnergy);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 43);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1278, 533);
            this.panel2.TabIndex = 8;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.comboBoxSelectDateTime);
            this.panel3.Controls.Add(this.textBox_V_type);
            this.panel3.Controls.Add(this.label1);
            this.panel3.Controls.Add(this.buttonExit);
            this.panel3.Controls.Add(this.buttonSelect);
            this.panel3.Controls.Add(this.textBoxNumber);
            this.panel3.Controls.Add(this.labelNumber);
            this.panel3.Controls.Add(this.buttonExportExcel);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Right;
            this.panel3.Location = new System.Drawing.Point(1078, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(200, 533);
            this.panel3.TabIndex = 12;
            // 
            // comboBoxSelectDateTime
            // 
            this.comboBoxSelectDateTime.FormattingEnabled = true;
            this.comboBoxSelectDateTime.Location = new System.Drawing.Point(39, 79);
            this.comboBoxSelectDateTime.Name = "comboBoxSelectDateTime";
            this.comboBoxSelectDateTime.Size = new System.Drawing.Size(131, 25);
            this.comboBoxSelectDateTime.TabIndex = 19;
            this.comboBoxSelectDateTime.Text = "请选择日期";
            this.comboBoxSelectDateTime.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // textBox_V_type
            // 
            this.textBox_V_type.Location = new System.Drawing.Point(76, 39);
            this.textBox_V_type.Name = "textBox_V_type";
            this.textBox_V_type.Size = new System.Drawing.Size(85, 23);
            this.textBox_V_type.TabIndex = 18;
            this.textBox_V_type.Text = "CRH400";
            this.textBox_V_type.Visible = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(36, 45);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(32, 17);
            this.label1.TabIndex = 17;
            this.label1.Text = "车型";
            this.label1.Visible = false;
            // 
            // buttonExit
            // 
            this.buttonExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.buttonExit.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            this.buttonExit.ImageIndex = 2;
            this.buttonExit.ImageList = this.imageList1;
            this.buttonExit.Location = new System.Drawing.Point(39, 486);
            this.buttonExit.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.buttonExit.Name = "buttonExit";
            this.buttonExit.Size = new System.Drawing.Size(131, 33);
            this.buttonExit.TabIndex = 16;
            this.buttonExit.Text = "退出";
            this.buttonExit.UseVisualStyleBackColor = true;
            this.buttonExit.Click += new System.EventHandler(this.buttonExit_Click);
            // 
            // buttonSelect
            // 
            this.buttonSelect.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.buttonSelect.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            this.buttonSelect.ImageIndex = 1;
            this.buttonSelect.ImageList = this.imageList1;
            this.buttonSelect.Location = new System.Drawing.Point(39, 430);
            this.buttonSelect.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.buttonSelect.Name = "buttonSelect";
            this.buttonSelect.Size = new System.Drawing.Size(131, 33);
            this.buttonSelect.TabIndex = 15;
            this.buttonSelect.Text = "选择文件";
            this.buttonSelect.UseVisualStyleBackColor = true;
            this.buttonSelect.Click += new System.EventHandler(this.buttonSelect_Click);
            // 
            // textBoxNumber
            // 
            this.textBoxNumber.Location = new System.Drawing.Point(76, 5);
            this.textBoxNumber.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.textBoxNumber.Name = "textBoxNumber";
            this.textBoxNumber.Size = new System.Drawing.Size(85, 23);
            this.textBoxNumber.TabIndex = 14;
            this.textBoxNumber.Text = "0000";
            this.textBoxNumber.Visible = false;
            // 
            // labelNumber
            // 
            this.labelNumber.AutoSize = true;
            this.labelNumber.Location = new System.Drawing.Point(36, 8);
            this.labelNumber.Name = "labelNumber";
            this.labelNumber.Size = new System.Drawing.Size(35, 17);
            this.labelNumber.TabIndex = 13;
            this.labelNumber.Text = "车号:";
            this.labelNumber.Visible = false;
            // 
            // buttonExportExcel
            // 
            this.buttonExportExcel.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.buttonExportExcel.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            this.buttonExportExcel.ImageIndex = 0;
            this.buttonExportExcel.ImageList = this.imageList1;
            this.buttonExportExcel.Location = new System.Drawing.Point(38, 372);
            this.buttonExportExcel.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.buttonExportExcel.Name = "buttonExportExcel";
            this.buttonExportExcel.Size = new System.Drawing.Size(131, 34);
            this.buttonExportExcel.TabIndex = 12;
            this.buttonExportExcel.Text = "  导出为Excel";
            this.buttonExportExcel.UseVisualStyleBackColor = true;
            this.buttonExportExcel.Click += new System.EventHandler(this.buttonExportExcel_Click);
            // 
            // dataGridViewEnergy
            // 
            this.dataGridViewEnergy.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewEnergy.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewEnergy.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColumnDate,
            this.ColumnPositivePower,
            this.ColumnNegativePower,
            this.ColumnTotalPower,
            this.ColumnDayConsumePower,
            this.ColumnDayFeedPower,
            this.ColumnConsumeTotalPower,
            this.ColumnV,
            this.ColumnI,
            this.ColumnPower,
            this.ColumnRealTimePower,
            this.ColumnV3rd,
            this.ColumnV5rd,
            this.ColumnV7rd,
            this.ColumnV9rd,
            this.ColumnI3rd,
            this.ColumnI5rd,
            this.ColumnI7rd,
            this.ColumnI9rd,
            this.Column1});
            this.dataGridViewEnergy.Location = new System.Drawing.Point(2, 0);
            this.dataGridViewEnergy.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.dataGridViewEnergy.Name = "dataGridViewEnergy";
            this.dataGridViewEnergy.RowTemplate.Height = 23;
            this.dataGridViewEnergy.Size = new System.Drawing.Size(1070, 523);
            this.dataGridViewEnergy.TabIndex = 4;
            this.dataGridViewEnergy.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dataGridViewEnergy_CellPainting);
            // 
            // ColumnDate
            // 
            this.ColumnDate.HeaderText = "时间";
            this.ColumnDate.Name = "ColumnDate";
            this.ColumnDate.ReadOnly = true;
            this.ColumnDate.Width = 107;
            // 
            // ColumnPositivePower
            // 
            this.ColumnPositivePower.HeaderText = "正向电能";
            this.ColumnPositivePower.Name = "ColumnPositivePower";
            this.ColumnPositivePower.ReadOnly = true;
            this.ColumnPositivePower.Width = 107;
            // 
            // ColumnNegativePower
            // 
            this.ColumnNegativePower.HeaderText = "反向电能";
            this.ColumnNegativePower.Name = "ColumnNegativePower";
            this.ColumnNegativePower.ReadOnly = true;
            this.ColumnNegativePower.Width = 107;
            // 
            // ColumnTotalPower
            // 
            this.ColumnTotalPower.HeaderText = "总电能";
            this.ColumnTotalPower.Name = "ColumnTotalPower";
            this.ColumnTotalPower.ReadOnly = true;
            this.ColumnTotalPower.Width = 107;
            // 
            // ColumnDayConsumePower
            // 
            this.ColumnDayConsumePower.HeaderText = "阶段正向电能";
            this.ColumnDayConsumePower.Name = "ColumnDayConsumePower";
            this.ColumnDayConsumePower.ReadOnly = true;
            this.ColumnDayConsumePower.Width = 107;
            // 
            // ColumnDayFeedPower
            // 
            this.ColumnDayFeedPower.HeaderText = "阶段反向电能";
            this.ColumnDayFeedPower.Name = "ColumnDayFeedPower";
            this.ColumnDayFeedPower.ReadOnly = true;
            this.ColumnDayFeedPower.Width = 107;
            // 
            // ColumnConsumeTotalPower
            // 
            this.ColumnConsumeTotalPower.HeaderText = "阶段总电能";
            this.ColumnConsumeTotalPower.Name = "ColumnConsumeTotalPower";
            this.ColumnConsumeTotalPower.ReadOnly = true;
            this.ColumnConsumeTotalPower.Width = 107;
            // 
            // ColumnV
            // 
            this.ColumnV.HeaderText = "电压";
            this.ColumnV.Name = "ColumnV";
            this.ColumnV.Width = 107;
            // 
            // ColumnI
            // 
            this.ColumnI.HeaderText = "电流";
            this.ColumnI.Name = "ColumnI";
            this.ColumnI.Width = 107;
            // 
            // ColumnPower
            // 
            this.ColumnPower.HeaderText = "功率因素";
            this.ColumnPower.Name = "ColumnPower";
            this.ColumnPower.Width = 108;
            // 
            // ColumnRealTimePower
            // 
            this.ColumnRealTimePower.HeaderText = "实时功率";
            this.ColumnRealTimePower.Name = "ColumnRealTimePower";
            this.ColumnRealTimePower.Width = 107;
            // 
            // ColumnV3rd
            // 
            this.ColumnV3rd.HeaderText = "电压3次谐波含有率";
            this.ColumnV3rd.Name = "ColumnV3rd";
            this.ColumnV3rd.Width = 107;
            // 
            // ColumnV5rd
            // 
            this.ColumnV5rd.HeaderText = "电压5次谐波含有率";
            this.ColumnV5rd.Name = "ColumnV5rd";
            this.ColumnV5rd.Width = 107;
            // 
            // ColumnV7rd
            // 
            this.ColumnV7rd.HeaderText = "电压7次谐波含有率";
            this.ColumnV7rd.Name = "ColumnV7rd";
            this.ColumnV7rd.Width = 107;
            // 
            // ColumnV9rd
            // 
            this.ColumnV9rd.HeaderText = "电压7次谐波含有率";
            this.ColumnV9rd.Name = "ColumnV9rd";
            this.ColumnV9rd.Width = 107;
            // 
            // ColumnI3rd
            // 
            this.ColumnI3rd.HeaderText = "电流3次谐波含有率";
            this.ColumnI3rd.Name = "ColumnI3rd";
            this.ColumnI3rd.Width = 107;
            // 
            // ColumnI5rd
            // 
            this.ColumnI5rd.HeaderText = "电流5次谐波含有率";
            this.ColumnI5rd.Name = "ColumnI5rd";
            this.ColumnI5rd.Width = 107;
            // 
            // ColumnI7rd
            // 
            this.ColumnI7rd.HeaderText = "电流7次谐波含有率";
            this.ColumnI7rd.Name = "ColumnI7rd";
            this.ColumnI7rd.Width = 107;
            // 
            // ColumnI9rd
            // 
            this.ColumnI9rd.HeaderText = "电流9次谐波含有率";
            this.ColumnI9rd.Name = "ColumnI9rd";
            this.ColumnI9rd.Width = 107;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "升弓车厢";
            this.Column1.Name = "Column1";
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(1278, 576);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.Name = "FormMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "苏州市万松电器有限公司";
            this.Shown += new System.EventHandler(this.FormMain_Shown);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewEnergy)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialogData;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label labelTitle;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button buttonExit;
        private System.Windows.Forms.Button buttonSelect;
        private System.Windows.Forms.TextBox textBoxNumber;
        private System.Windows.Forms.Label labelNumber;
        private System.Windows.Forms.Button buttonExportExcel;
        private System.Windows.Forms.DataGridView dataGridViewEnergy;
        private System.Windows.Forms.TextBox textBox_V_type;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox comboBoxSelectDateTime;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnDate;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnPositivePower;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnNegativePower;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnTotalPower;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnDayConsumePower;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnDayFeedPower;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnConsumeTotalPower;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnV;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnI;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnPower;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnRealTimePower;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnV3rd;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnV5rd;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnV7rd;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnV9rd;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnI3rd;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnI5rd;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnI7rd;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnI9rd;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
    }
}

