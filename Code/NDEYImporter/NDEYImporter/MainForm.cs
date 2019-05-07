using NDEYImporter.DB;
using Noear.Weed;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NDEYImporter
{
    public partial class MainForm : Form
    {
        /// <summary>
        /// NDEY上报包目录
        /// </summary>
        public static string PackageDir = System.IO.Path.Combine(Application.StartupPath, "PackageDir");

        /// <summary>
        /// NDEY上报包解压目录
        /// </summary>
        public static string DBTempDir = System.IO.Path.Combine(Application.StartupPath, "DBTempDir");

        public MainForm()
        {
            InitializeComponent();

            //创建NDEY上报包目录
            try
            {
                System.IO.Directory.CreateDirectory(PackageDir);
            }
            catch (Exception ex) { }

            //创建NDEY上报包解压目录
            try
            {
                System.IO.Directory.CreateDirectory(DBTempDir);
            }
            catch (Exception ex) { }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {

        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            reloadCatalogList();
        }

        /// <summary>
        /// 载入Catalog数据
        /// </summary>
        private void reloadCatalogList()
        {
            //清理GridView
            dgvCatalogs.Rows.Clear();

            //获得Catalog全部列表
            DataList dlCatalogs = ConnectionManager.Context.table("Catalog").select("*").getDataList();
            //显示序号
            int rIndex = 0;
            //向GridView添加可显示的列表
            foreach (DataItem di in dlCatalogs.getRows())
            {
                rIndex++;

                //行数据
                List<object> cells = new List<object>();
                cells.Add(rIndex);
                cells.Add(di.getString("ProjectNumber"));
                cells.Add(di.getString("ProjectName"));
                cells.Add(di.getString("ProjectCreater"));

                //提取并向GridView中添加Project中的单位名称
                string unitName = ConnectionManager.Context.table("Project").where("ID='" + di.getString("ProjectID") + "'").select("UnitName").getValue<string>(string.Empty);
                cells.Add(unitName);

                //提取并向GridView中添加Unit中的单位名称
                string realUnitName = ConnectionManager.Context.table("Unit").where("ID='" + di.getString("ProjectCreaterUnitID") + "'").select("UnitName").getValue<string>(string.Empty);
                cells.Add(realUnitName);

                //添加行数据，并获得行号
                int rowIndex = dgvCatalogs.Rows.Add(cells.ToArray());
                dgvCatalogs.Rows[rowIndex].Tag = di;
            }            
        }

        private void dgvCatalogs_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}