using NDEYImporter.DB;
using NDEYImporter.Util;
using Noear.Weed;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
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

                        //刷新列表
                        reloadCatalogList();

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
                //获得要删除的项目ID,项目编号
                string projectId = ((DataItem)dgvCatalogs.Rows[e.RowIndex].Tag).getString("ProjectID");
                string projectNumber = ((DataItem)dgvCatalogs.Rows[e.RowIndex].Tag).getString("ProjectNumber");

                //显示删除提示框
                if (MessageBox.Show("真的要删除吗？", "提示", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                {
                    //删除项目数据
                    DBImporter.deleteProject(projectId);

                    //删除申报包缓存
                    try
                    {
                        System.IO.Directory.Delete(System.IO.Path.Combine(PackageDir, projectNumber), true);
                    }
                    catch (Exception ex) { }

                    //刷新GridView
                    reloadCatalogList();
                }
            }
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            //输出的Excel路径
            string excelFile = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "output.xlsx");

            //Excel数据
            MemoryStream memoryStream = new MemoryStream();
            //创建Workbook
            NPOI.XSSF.UserModel.XSSFWorkbook workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
            //创建Sheet
            NPOI.SS.UserModel.ISheet sheet = workbook.CreateSheet();
            //创建单元格设置对象
            NPOI.SS.UserModel.ICellStyle cellStyle = workbook.CreateCellStyle();
            //设置水平、垂直居中
            cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            //设置边框
            cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;

            //创建设置字体对象
            NPOI.SS.UserModel.IFont font = workbook.CreateFont();
            font.FontHeightInPoints = 16;//设置字体大小
            font.Boldweight = short.MaxValue; //加粗
            font.FontName = "宋体";
            cellStyle.SetFont(font);

            //读取索引数据
            DataList dlCatalogs = ConnectionManager.Context.table("Catalog").select("*").getDataList();
            //行号
            int rowIndex = 0;
            //输出Excel数据
            foreach (DataItem diCatalogItem in dlCatalogs.getRows())
            {
                //项目ID
                string projectID = diCatalogItem.getString("ProjectID");
                //项目编号
                string projectNumber = diCatalogItem.getString("ProjectNumber");
                //项目申请单位ID
                string projectUnitID = diCatalogItem.getString("ProjectCreaterUnitID");

                //需要输出的数据
                List<KeyValuePair<string, object>> projectItem = new List<KeyValuePair<string, object>>();
                projectItem.Add(new KeyValuePair<string, object>("项目编号", projectNumber));

                //读取项目数据
                DataList dlProjects = ConnectionManager.Context.table("Project").where("ID='" + projectID + "'").select("*").getDataList();
                //判断是否可以填充数据恢复
                if (dlProjects.getRowCount() >= 1)
                {
                    projectItem.Add(new KeyValuePair<string, object>("申请人", dlProjects.getRow(0).get("UserName")));
                    projectItem.Add(new KeyValuePair<string, object>("性别", dlProjects.getRow(0).get("Sex")));
                    projectItem.Add(new KeyValuePair<string, object>("出生年月", dlProjects.getRow(0).get("Birthdate")));
                    projectItem.Add(new KeyValuePair<string, object>("学位", dlProjects.getRow(0).get("Degree")));
                    projectItem.Add(new KeyValuePair<string, object>("职称", dlProjects.getRow(0).get("JobTitle")));
                    projectItem.Add(new KeyValuePair<string, object>("职务", dlProjects.getRow(0).get("UnitPosition")));
                    projectItem.Add(new KeyValuePair<string, object>("主要研究领域", dlProjects.getRow(0).get("MainResearch")));
                    projectItem.Add(new KeyValuePair<string, object>("身份证号码", dlProjects.getRow(0).get("CardNo")));
                    projectItem.Add(new KeyValuePair<string, object>("办公电话", dlProjects.getRow(0).get("OfficePhones")));
                    projectItem.Add(new KeyValuePair<string, object>("手机", dlProjects.getRow(0).get("MobilePhone")));
                    projectItem.Add(new KeyValuePair<string, object>("电子邮箱", dlProjects.getRow(0).get("EMail")));
                    projectItem.Add(new KeyValuePair<string, object>("单位名称", dlProjects.getRow(0).get("UnitName")));

                    //读取开户帐号
                    string unitBankNo = ConnectionManager.Context.table("Unit").where("ID='" + projectUnitID + "'").select("UnitBankNo").getValue<string>(string.Empty);
                    projectItem.Add(new KeyValuePair<string, object>("开户账号", unitBankNo));

                    projectItem.Add(new KeyValuePair<string, object>("单位常用名", dlProjects.getRow(0).get("UnitNormal")));
                    projectItem.Add(new KeyValuePair<string, object>("单位联系人", dlProjects.getRow(0).get("UnitContacts")));
                    projectItem.Add(new KeyValuePair<string, object>("隶属部门", dlProjects.getRow(0).get("UnitForORG")));
                    projectItem.Add(new KeyValuePair<string, object>("单位性质", dlProjects.getRow(0).get("UnitProperties")));
                    projectItem.Add(new KeyValuePair<string, object>("单位联系电话", dlProjects.getRow(0).get("UnitContactsPhone")));
                    projectItem.Add(new KeyValuePair<string, object>("机要地址", dlProjects.getRow(0).get("UnitAddress")));
                    projectItem.Add(new KeyValuePair<string, object>("项目名称", dlProjects.getRow(0).get("ProjectName")));
                    projectItem.Add(new KeyValuePair<string, object>("技术方向", dlProjects.getRow(0).get("ProjectTD")));
                    projectItem.Add(new KeyValuePair<string, object>("主要研究方向", dlProjects.getRow(0).get("ProjectMRD")));
                    projectItem.Add(new KeyValuePair<string, object>("基地类别", dlProjects.getRow(0).get("ProjectBaseT")));
                    projectItem.Add(new KeyValuePair<string, object>("申报类别", dlProjects.getRow(0).get("ApplicationArea")));
                    projectItem.Add(new KeyValuePair<string, object>("研究周期", "五年"));
                    projectItem.Add(new KeyValuePair<string, object>("密级", dlProjects.getRow(0).get("ProjectSecret")));

                    //处理摘要
                    string projectBrief = dlProjects.getRow(0).getString("ProjectBrief");
                    if (projectBrief != string.Empty)
                    {
                        string[] array = projectBrief.Split(new char[]
					{
						'|'
					});
                        if (array.Length == 6)
                        {
                            projectItem.Add(new KeyValuePair<string, object>("入选人才计划情况", array[0]));
                            projectItem.Add(new KeyValuePair<string, object>("代表性成果或贡献", array[1]));
                            projectItem.Add(new KeyValuePair<string, object>("研究目标", array[2]));
                            projectItem.Add(new KeyValuePair<string, object>("研究内容", array[3]));
                            projectItem.Add(new KeyValuePair<string, object>("主要创新点", array[4]));
                            projectItem.Add(new KeyValuePair<string, object>("预期军事价值", array[5]));
                        }
                    }
                    
                    //处理成果形式
                    projectItem.Add(new KeyValuePair<string, object>("成果形式", new ResultConfigRecord(dlProjects.getRow(0).getString("ResultConfig")).getDescription()));

                    projectItem.Add(new KeyValuePair<string, object>("推荐方式", dlProjects.getRow(0).get("ApplicationType") == null ? "单位推荐" : dlProjects.getRow(0).get("ApplicationType")));

                    projectItem.Add(new KeyValuePair<string, object>("专家一姓名", dlProjects.getRow(0).get("ExpertName1")));
                    projectItem.Add(new KeyValuePair<string, object>("专家一研究领域", dlProjects.getRow(0).get("ExpertArea1")));
                    projectItem.Add(new KeyValuePair<string, object>("专家一职务职称", dlProjects.getRow(0).get("ExpertUnitPosition1")));
                    projectItem.Add(new KeyValuePair<string, object>("专家一工作单位", dlProjects.getRow(0).get("ExpertUnit1")));
                    projectItem.Add(new KeyValuePair<string, object>("专家二姓名", dlProjects.getRow(0).get("ExpertName2")));
                    projectItem.Add(new KeyValuePair<string, object>("专家二研究领域", dlProjects.getRow(0).get("ExpertArea2")));
                    projectItem.Add(new KeyValuePair<string, object>("专家二职务职称", dlProjects.getRow(0).get("ExpertUnitPosition2")));
                    projectItem.Add(new KeyValuePair<string, object>("专家二工作单位", dlProjects.getRow(0).get("ExpertUnit2")));
                    projectItem.Add(new KeyValuePair<string, object>("专家三姓名", dlProjects.getRow(0).get("ExpertName3")));
                    projectItem.Add(new KeyValuePair<string, object>("专家三研究领域", dlProjects.getRow(0).get("ExpertArea3")));
                    projectItem.Add(new KeyValuePair<string, object>("专家三职务职称", dlProjects.getRow(0).get("ExpertUnitPosition3")));
                    projectItem.Add(new KeyValuePair<string, object>("专家三工作单位", dlProjects.getRow(0).get("ExpertUnit3")));
                }

                //列号
                int colIndex = 0;

                //创建行
                NPOI.SS.UserModel.IRow row = sheet.CreateRow(rowIndex);

                //输出数据到Excel
                foreach (KeyValuePair<string, object> kvp in projectItem)
                {
                    //列名
                    //创建列
                    NPOI.SS.UserModel.ICell cell = row.CreateCell(colIndex);
                    //设置样式
                    cell.CellStyle = cellStyle;
                    //设置数据
                    cell.SetCellValue(kvp.Key);
                    colIndex++;

                    //列值
                    //创建列
                    cell = row.CreateCell(colIndex);
                    //设置样式
                    cell.CellStyle = cellStyle;
                    //设置数据
                    cell.SetCellValue(kvp.Value != null ? kvp.Value.ToString() : string.Empty);
                    colIndex++;
                }

                rowIndex++;
            }

            //输出到Excel文件
            workbook.Write(memoryStream);
            File.WriteAllBytes(excelFile, memoryStream.ToArray());

            MessageBox.Show("导出完成！","提示");
            //打开Excel文件
            System.Diagnostics.Process.Start(excelFile);
        }
    }

    /// <summary>
    /// 成果形式配置
    /// </summary>
    public class ResultConfigRecord
    {
        public ResultConfigRecord(string config)
        {
            string[] array = config.Split(new char[]
					{
						'|'
					});
            if (array.Length == 13)
            {
                isCBR1Checked = bool.Parse(array[0]);
                isCBR2Checked = bool.Parse(array[1]);
                isCBR3Checked = bool.Parse(array[2]);
                isCBR4Checked = bool.Parse(array[3]);
                isCBR5Checked = bool.Parse(array[4]);
                isCBR6Checked = bool.Parse(array[5]);
                isCBR7Checked = bool.Parse(array[6]);
                isCBR8Checked = bool.Parse(array[7]);
                isCBR9Checked = bool.Parse(array[8]);
                isCBR10Checked = bool.Parse(array[9]);
                isCBR11Checked = bool.Parse(array[10]);
                isCBR12Checked = bool.Parse(array[11]);
                cbrOtherText = array[12];
            }
        }

        public bool isCBR1Checked
        {
            get;
            set;
        }

        public bool isCBR2Checked
        {
            get;
            set;
        }

        public bool isCBR3Checked
        {
            get;
            set;
        }

        public bool isCBR4Checked
        {
            get;
            set;
        }

        public bool isCBR5Checked
        {
            get;
            set;
        }

        public bool isCBR6Checked
        {
            get;
            set;
        }

        public bool isCBR7Checked
        {
            get;
            set;
        }

        public bool isCBR8Checked
        {
            get;
            set;
        }

        public bool isCBR9Checked
        {
            get;
            set;
        }

        public bool isCBR10Checked
        {
            get;
            set;
        }

        public bool isCBR11Checked
        {
            get;
            set;
        }

        public bool isCBR12Checked
        {
            get;
            set;
        }

        public string cbrOtherText
        {
            get;
            set;
        }

        public string getDescription()
        {
            string text = "";
            if (this.isCBR1Checked)
            {
                text += "样品、";
            }
            if (this.isCBR2Checked)
            {
                text += "样机、";
            }
            if (this.isCBR3Checked)
            {
                text += "数据库、";
            }
            if (this.isCBR4Checked)
            {
                text += "试验（演示）系统、";
            }
            if (this.isCBR5Checked)
            {
                text += "软件、";
            }
            if (this.isCBR6Checked)
            {
                text += "仿真模型、";
            }
            if (this.isCBR7Checked)
            {
                text += "工程工艺、";
            }
            if (this.isCBR8Checked)
            {
                text += "标准（规范）、";
            }
            if (this.isCBR9Checked)
            {
                text += "论文、";
            }
            if (this.isCBR10Checked)
            {
                text += "专利、";
            }
            if (this.isCBR11Checked)
            {
                text += "研究报告、";
            }
            if (this.isCBR12Checked)
            {
                text += "试验方案（报告）、";
            }
            if (this.cbrOtherText != "")
            {
                text += this.cbrOtherText;
            }
            text = text.TrimEnd(new char[]
			{
				'、'
			});
            if (text != string.Empty)
            {
                text += "。";
            }
            return text;
        }
    }
}