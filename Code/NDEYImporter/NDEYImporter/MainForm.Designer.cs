namespace NDEYImporter
{
    partial class MainForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.tsslHintText = new System.Windows.Forms.ToolStripStatusLabel();
            this.tsToolBar = new System.Windows.Forms.ToolStrip();
            this.btnImport = new System.Windows.Forms.ToolStripButton();
            this.btnSelectTotalDir = new System.Windows.Forms.ToolStripButton();
            this.btnImportAll = new System.Windows.Forms.ToolStripButton();
            this.btnImportWithSelectedList = new System.Windows.Forms.ToolStripButton();
            this.btnExportExcel = new System.Windows.Forms.ToolStripButton();
            this.btnExit = new System.Windows.Forms.ToolStripButton();
            this.plContent = new System.Windows.Forms.Panel();
            this.dgvCatalogs = new System.Windows.Forms.DataGridView();
            this.colIndex = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCreater = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCreaterUnit = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colRealUnit = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDel = new System.Windows.Forms.DataGridViewButtonColumn();
            this.ofdPackages = new System.Windows.Forms.OpenFileDialog();
            this.fbdTotalDirSelect = new System.Windows.Forms.FolderBrowserDialog();
            this.btnUnZipWithSelectedList = new System.Windows.Forms.ToolStripButton();
            this.btnSelectUnZipDir = new System.Windows.Forms.ToolStripButton();
            this.btnUnZipAll = new System.Windows.Forms.ToolStripButton();
            this.ssStatus.SuspendLayout();
            this.tsToolBar.SuspendLayout();
            this.plContent.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCatalogs)).BeginInit();
            this.SuspendLayout();
            // 
            // ssStatus
            // 
            this.ssStatus.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsslHintText});
            this.ssStatus.Location = new System.Drawing.Point(0, 498);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(1067, 22);
            this.ssStatus.TabIndex = 0;
            this.ssStatus.Text = "statusStrip1";
            // 
            // tsslHintText
            // 
            this.tsslHintText.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tsslHintText.Name = "tsslHintText";
            this.tsslHintText.Size = new System.Drawing.Size(0, 17);
            // 
            // tsToolBar
            // 
            this.tsToolBar.ImageScalingSize = new System.Drawing.Size(32, 32);
            this.tsToolBar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnImport,
            this.btnSelectTotalDir,
            this.btnImportAll,
            this.btnImportWithSelectedList,
            this.btnExportExcel,
            this.btnSelectUnZipDir,
            this.btnUnZipAll,
            this.btnUnZipWithSelectedList,
            this.btnExit});
            this.tsToolBar.Location = new System.Drawing.Point(0, 0);
            this.tsToolBar.Name = "tsToolBar";
            this.tsToolBar.Size = new System.Drawing.Size(1067, 39);
            this.tsToolBar.TabIndex = 1;
            // 
            // btnImport
            // 
            this.btnImport.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnImport.Image = ((System.Drawing.Image)(resources.GetObject("btnImport.Image")));
            this.btnImport.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(71, 36);
            this.btnImport.Text = "导入";
            this.btnImport.Visible = false;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // btnSelectTotalDir
            // 
            this.btnSelectTotalDir.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnSelectTotalDir.Image = ((System.Drawing.Image)(resources.GetObject("btnSelectTotalDir.Image")));
            this.btnSelectTotalDir.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnSelectTotalDir.Name = "btnSelectTotalDir";
            this.btnSelectTotalDir.Size = new System.Drawing.Size(155, 36);
            this.btnSelectTotalDir.Text = "设置申报包总目录";
            this.btnSelectTotalDir.Click += new System.EventHandler(this.btnSelectTotalDir_Click);
            // 
            // btnImportAll
            // 
            this.btnImportAll.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnImportAll.Image = ((System.Drawing.Image)(resources.GetObject("btnImportAll.Image")));
            this.btnImportAll.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnImportAll.Name = "btnImportAll";
            this.btnImportAll.Size = new System.Drawing.Size(99, 36);
            this.btnImportAll.Text = "全部导入";
            this.btnImportAll.Click += new System.EventHandler(this.btnImportAll_Click);
            // 
            // btnImportWithSelectedList
            // 
            this.btnImportWithSelectedList.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnImportWithSelectedList.Image = ((System.Drawing.Image)(resources.GetObject("btnImportWithSelectedList.Image")));
            this.btnImportWithSelectedList.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnImportWithSelectedList.Name = "btnImportWithSelectedList";
            this.btnImportWithSelectedList.Size = new System.Drawing.Size(113, 36);
            this.btnImportWithSelectedList.Text = "选择性导入";
            this.btnImportWithSelectedList.Click += new System.EventHandler(this.btnImportWithSelectedList_Click);
            // 
            // btnExportExcel
            // 
            this.btnExportExcel.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnExportExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExportExcel.Image")));
            this.btnExportExcel.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnExportExcel.Name = "btnExportExcel";
            this.btnExportExcel.Size = new System.Drawing.Size(148, 36);
            this.btnExportExcel.Text = "导出列表到Excel";
            this.btnExportExcel.Click += new System.EventHandler(this.btnExportExcel_Click);
            // 
            // btnExit
            // 
            this.btnExit.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnExit.Image = ((System.Drawing.Image)(resources.GetObject("btnExit.Image")));
            this.btnExit.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(71, 36);
            this.btnExit.Text = "退出";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // plContent
            // 
            this.plContent.Controls.Add(this.dgvCatalogs);
            this.plContent.Dock = System.Windows.Forms.DockStyle.Fill;
            this.plContent.Location = new System.Drawing.Point(0, 39);
            this.plContent.Name = "plContent";
            this.plContent.Size = new System.Drawing.Size(1067, 459);
            this.plContent.TabIndex = 2;
            // 
            // dgvCatalogs
            // 
            this.dgvCatalogs.AllowUserToAddRows = false;
            this.dgvCatalogs.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvCatalogs.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colIndex,
            this.colNumber,
            this.colName,
            this.colCreater,
            this.colCreaterUnit,
            this.colRealUnit,
            this.colDel});
            this.dgvCatalogs.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvCatalogs.Location = new System.Drawing.Point(0, 0);
            this.dgvCatalogs.MultiSelect = false;
            this.dgvCatalogs.Name = "dgvCatalogs";
            this.dgvCatalogs.ReadOnly = true;
            this.dgvCatalogs.RowHeadersVisible = false;
            this.dgvCatalogs.RowTemplate.Height = 23;
            this.dgvCatalogs.Size = new System.Drawing.Size(1067, 459);
            this.dgvCatalogs.TabIndex = 0;
            this.dgvCatalogs.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvCatalogs_CellContentClick);
            // 
            // colIndex
            // 
            this.colIndex.HeaderText = "序号";
            this.colIndex.Name = "colIndex";
            this.colIndex.ReadOnly = true;
            this.colIndex.Width = 60;
            // 
            // colNumber
            // 
            this.colNumber.HeaderText = "项目编号";
            this.colNumber.Name = "colNumber";
            this.colNumber.ReadOnly = true;
            // 
            // colName
            // 
            this.colName.HeaderText = "项目名称";
            this.colName.Name = "colName";
            this.colName.ReadOnly = true;
            this.colName.Width = 200;
            // 
            // colCreater
            // 
            this.colCreater.HeaderText = "申请人";
            this.colCreater.Name = "colCreater";
            this.colCreater.ReadOnly = true;
            // 
            // colCreaterUnit
            // 
            this.colCreaterUnit.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colCreaterUnit.HeaderText = "申请人单位";
            this.colCreaterUnit.Name = "colCreaterUnit";
            this.colCreaterUnit.ReadOnly = true;
            // 
            // colRealUnit
            // 
            this.colRealUnit.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colRealUnit.HeaderText = "申请单位名称(标准库)";
            this.colRealUnit.Name = "colRealUnit";
            this.colRealUnit.ReadOnly = true;
            // 
            // colDel
            // 
            this.colDel.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colDel.HeaderText = "";
            this.colDel.Name = "colDel";
            this.colDel.ReadOnly = true;
            this.colDel.Text = "删除";
            this.colDel.UseColumnTextForButtonValue = true;
            this.colDel.Width = 5;
            // 
            // ofdPackages
            // 
            this.ofdPackages.Filter = "*.zip|*.zip";
            // 
            // btnUnZipWithSelectedList
            // 
            this.btnUnZipWithSelectedList.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnUnZipWithSelectedList.Image = ((System.Drawing.Image)(resources.GetObject("btnUnZipWithSelectedList.Image")));
            this.btnUnZipWithSelectedList.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnUnZipWithSelectedList.Name = "btnUnZipWithSelectedList";
            this.btnUnZipWithSelectedList.Size = new System.Drawing.Size(113, 36);
            this.btnUnZipWithSelectedList.Text = "选择性解压";
            this.btnUnZipWithSelectedList.Click += new System.EventHandler(this.btnUnZipWithSelectedList_Click);
            // 
            // btnSelectUnZipDir
            // 
            this.btnSelectUnZipDir.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnSelectUnZipDir.Image = ((System.Drawing.Image)(resources.GetObject("btnSelectUnZipDir.Image")));
            this.btnSelectUnZipDir.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnSelectUnZipDir.Name = "btnSelectUnZipDir";
            this.btnSelectUnZipDir.Size = new System.Drawing.Size(127, 36);
            this.btnSelectUnZipDir.Text = "选择解包目录";
            this.btnSelectUnZipDir.Click += new System.EventHandler(this.btnSelectUnZipDir_Click);
            // 
            // btnUnZipAll
            // 
            this.btnUnZipAll.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnUnZipAll.Image = ((System.Drawing.Image)(resources.GetObject("btnUnZipAll.Image")));
            this.btnUnZipAll.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnUnZipAll.Name = "btnUnZipAll";
            this.btnUnZipAll.Size = new System.Drawing.Size(99, 36);
            this.btnUnZipAll.Text = "全部解压";
            this.btnUnZipAll.Click += new System.EventHandler(this.btnUnZipAll_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1067, 520);
            this.Controls.Add(this.plContent);
            this.Controls.Add(this.tsToolBar);
            this.Controls.Add(this.ssStatus);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "国防科技卓越青年科学基金申报系统汇总版";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ssStatus.ResumeLayout(false);
            this.ssStatus.PerformLayout();
            this.tsToolBar.ResumeLayout(false);
            this.tsToolBar.PerformLayout();
            this.plContent.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvCatalogs)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.ToolStrip tsToolBar;
        private System.Windows.Forms.Panel plContent;
        private System.Windows.Forms.DataGridView dgvCatalogs;
        private System.Windows.Forms.ToolStripButton btnImport;
        private System.Windows.Forms.ToolStripButton btnExit;
        private System.Windows.Forms.OpenFileDialog ofdPackages;
        private System.Windows.Forms.ToolStripButton btnExportExcel;
        private System.Windows.Forms.DataGridViewTextBoxColumn colIndex;
        private System.Windows.Forms.DataGridViewTextBoxColumn colNumber;
        private System.Windows.Forms.DataGridViewTextBoxColumn colName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCreater;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCreaterUnit;
        private System.Windows.Forms.DataGridViewTextBoxColumn colRealUnit;
        private System.Windows.Forms.DataGridViewButtonColumn colDel;
        private System.Windows.Forms.FolderBrowserDialog fbdTotalDirSelect;
        private System.Windows.Forms.ToolStripButton btnImportWithSelectedList;
        private System.Windows.Forms.ToolStripButton btnSelectTotalDir;
        private System.Windows.Forms.ToolStripButton btnImportAll;
        private System.Windows.Forms.ToolStripStatusLabel tsslHintText;
        private System.Windows.Forms.ToolStripButton btnUnZipWithSelectedList;
        private System.Windows.Forms.ToolStripButton btnSelectUnZipDir;
        private System.Windows.Forms.ToolStripButton btnUnZipAll;

    }
}