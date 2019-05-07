using NDEYImporter.DB;
using NDEYImporter.Util;
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

            //清空NDEY上报包解压目录
            try
            {
                System.IO.Directory.Delete(DBTempDir, true);
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
            //显示选择申报包的对话框
            if (ofdPackages.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //将压缩包解压到临时目录
                //生成一个临时路径
                string destDir = System.IO.Path.Combine(DBTempDir, DateTime.Now.Ticks.ToString());
                //解压ZIP包
                FileZipOpr fzo = new FileZipOpr();
                fzo.UnZipFile(ofdPackages.FileName, destDir, string.Empty, true);

                //判断申报包是否有效
                //生成DB文件路径
                string dbFile = System.IO.Path.Combine(destDir, @"Files\myData.db");
                //判断文件是否存在
                if (System.IO.File.Exists(dbFile))
                {
                    //有效的申报包
                    try
                    {
                        //导入数据并返回项目Id
                        string projectId = DBImporter.importDB(dbFile);

                        //获取项目编号
                        string projectNumber = ConnectionManager.Context.table("Catalog").where("ProjectID='" + projectId + "'").select("ProjectNumber").getValue<string>(string.Empty);

                        //清理申报包保存路径
                        if (System.IO.Directory.Exists(System.IO.Path.Combine(PackageDir, projectNumber)))
                        {
                            try
                            {
                                System.IO.Directory.Delete(System.IO.Path.Combine(PackageDir, projectNumber), true);
                            }
                            catch (Exception ex) { }
                        }

                        //创建申报包保存目录
                        try
                        {
                            System.IO.Directory.CreateDirectory(System.IO.Path.Combine(PackageDir, projectNumber));
                        }
                        catch (Exception ex) { }

                        //生成申报包保存路径
                        string packageDestPath = System.IO.Path.Combine(System.IO.Path.Combine(PackageDir, projectNumber), new System.IO.FileInfo(ofdPackages.FileName).Name);

                        //将申报包Copy到保存路径
                        System.IO.File.Copy(ofdPackages.FileName, packageDestPath, true);

                        MessageBox.Show("申报数据包导入完成！", "提示");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("对不起,申报数据包导入失败！", "提示");
                    }
                }
                else
                {
                    MessageBox.Show("对不起,无效的申报数据包！", "提示");
                }
            }
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
            //检查是否点击的是删除的那一列
            if (e.ColumnIndex == dgvCatalogs.Columns.Count - 1)
            {
                //获得要删除的项目ID
                string projectId = ((DataItem)dgvCatalogs.Rows[e.RowIndex].Tag).getString("ProjectID");

                //显示删除提示框
                if (MessageBox.Show("真的要删除吗？", "提示", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                {
                    //删除项目
                    DBImporter.deleteProject(projectId);

                    //刷新GridView
                    reloadCatalogList();
                }
            }
        }
    }
}