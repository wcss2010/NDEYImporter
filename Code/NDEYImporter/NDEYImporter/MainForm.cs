using NDEYImporter.DB;
using NDEYImporter.Forms;
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
        /// 程序配置
        /// </summary>
        public static MainConfig Config { get; set; }

        /// <summary>
        /// NDEY上报包目录
        /// </summary>
        public static string PackageDir = System.IO.Path.Combine(Application.StartupPath, "PackageDir");

        /// <summary>
        /// NDEY上报包解压目录
        /// </summary>
        public static string DBTempDir = System.IO.Path.Combine(Application.StartupPath, "DBTempDir");

        /// <summary>
        /// 主窗体实例
        /// </summary>
        public static MainForm Instance { get; set; }

        public MainForm()
        {
            InitializeComponent();

            //设置运行实例
            MainForm.Instance = this;

            //载入配置
            MainForm.Instance.loadConfig();

            //创建NDEY上报包目录
            try
            {
                System.IO.Directory.CreateDirectory(PackageDir);
            }
            catch (Exception ex) { MainForm.writeLog(ex.ToString()); }

            //清空NDEY上报包解压目录
            try
            {
                System.IO.Directory.Delete(DBTempDir, true);
            }
            catch (Exception ex) { MainForm.writeLog(ex.ToString()); }

            //创建NDEY上报包解压目录
            try
            {
                System.IO.Directory.CreateDirectory(DBTempDir);
            }
            catch (Exception ex) { MainForm.writeLog(ex.ToString()); }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            //载入Catalog列表
            reloadCatalogList();

            //刷新状态栏中的总目录
            refreshTotalDir();
        }

        /// <summary>
        /// 刷新状态栏中的总目录
        /// </summary>
        private void refreshTotalDir()
        {
            tsslHintText.Text = "总目录：" + MainForm.Config.TotalDir;
        }

        /// <summary>
        /// 载入配置
        /// </summary>
        public void loadConfig()
        {
            string xmlFile = Path.Combine(Application.StartupPath, "config.xml");
            if (File.Exists(xmlFile))
            {
                try
                {
                    //读取配置
                    MainForm.Config = XmlSerializeTool.deserialize<MainConfig>(File.ReadAllText(xmlFile));
                }
                catch (Exception ex)
                {
                    //设置默认的配置项
                    MainForm.Config = new MainConfig();
                    MainForm.Config.TotalDir = PackageDir;
                    saveConfig();
                }
            }
            else
            {
                //设置默认的配置项
                MainForm.Config = new MainConfig();
                MainForm.Config.TotalDir = PackageDir;
                saveConfig();
            }
        }

        /// <summary>
        /// 保存配置
        /// </summary>
        public void saveConfig()
        {
            string xmlFile = Path.Combine(Application.StartupPath, "config.xml");
            if (MainForm.Config != null)
            {
                //写配置文件
                File.WriteAllText(xmlFile, XmlSerializeTool.serializer<MainConfig>(MainForm.Config));
            }
        }

        /// <summary>
        /// 载入Catalog数据
        /// </summary>
        public void reloadCatalogList()
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
                    catch (Exception ex) { MainForm.writeLog(ex.ToString()); }

                    //刷新GridView
                    reloadCatalogList();
                }
            }
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            try
            {
                //输出的Excel路径
                string excelFile = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "output.xlsx");

                //Excel数据
                MemoryStream memoryStream = new MemoryStream();
                //创建Workbook
                NPOI.XSSF.UserModel.XSSFWorkbook workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
                //创建Sheet
                NPOI.SS.UserModel.ISheet sheet = workbook.CreateSheet();

                //创建单元格设置对象(普通内容)
                NPOI.SS.UserModel.ICellStyle cellStyleA = workbook.CreateCellStyle();
                cellStyleA.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
                cellStyleA.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                cellStyleA.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                cellStyleA.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                cellStyleA.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                cellStyleA.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                cellStyleA.WrapText = true;

                //创建单元格设置对象(普通内容)
                NPOI.SS.UserModel.ICellStyle cellStyleB = workbook.CreateCellStyle();
                cellStyleB.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                cellStyleB.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                cellStyleB.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                cellStyleB.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                cellStyleB.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                cellStyleB.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                cellStyleB.WrapText = true;

                //创建设置字体对象(内容字体)
                NPOI.SS.UserModel.IFont fontA = workbook.CreateFont();
                fontA.FontHeightInPoints = 16;//设置字体大小
                fontA.FontName = "宋体";
                cellStyleA.SetFont(fontA);

                //创建设置字体对象(标题字体)
                NPOI.SS.UserModel.IFont fontB = workbook.CreateFont();
                fontB.FontHeightInPoints = 16;//设置字体大小
                fontB.FontName = "宋体";
                fontB.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                cellStyleB.SetFont(fontB);

                //是否需要输出表头
                bool isNeedCreateHeader = true;

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

                        //研究目标，研究内容，主要创建点，预期军事价值
                        string muBiaoNeiRongChuangXinDianJieZhi = string.Empty;
                        //资助对象简介
                        string ziZhuDuiXiang = string.Empty;

                        //处理摘要
                        string projectBrief = dlProjects.getRow(0).getString("ProjectBrief");
                        if (projectBrief != string.Empty)
                        {
                            //拆分摘要内容
                            string[] array = projectBrief.Split(new string[] { "|" }, StringSplitOptions.None);

                            //判断字符数组长度是否为6
                            if (array.Length == 6)
                            {
                                //处理某些字段没有值，只显示一个句号的情况
                                for (int k = 0; k < array.Length; k++)
                                {
                                    if (array[k] != null && array[k].Trim().Length < 2)
                                    {
                                        array[k] = string.Empty;
                                    }
                                }

                                //向列表中添加数据
                                projectItem.Add(new KeyValuePair<string, object>("入选人才计划情况", array[0].Trim()));
                                projectItem.Add(new KeyValuePair<string, object>("代表性成果或贡献", array[1].Trim()));
                                projectItem.Add(new KeyValuePair<string, object>("研究目标", array[2].Trim()));
                                projectItem.Add(new KeyValuePair<string, object>("研究内容", array[3].Trim()));
                                projectItem.Add(new KeyValuePair<string, object>("主要创新点", array[4].Trim()));
                                projectItem.Add(new KeyValuePair<string, object>("预期军事价值", array[5].Trim()));

                                //拼装研究目标，研究内容，主要创建点，预期军事价值
                                muBiaoNeiRongChuangXinDianJieZhi += "研究目标:" + array[2].Trim() + "\n";
                                muBiaoNeiRongChuangXinDianJieZhi += "研究内容:" + array[3].Trim() + "\n";
                                muBiaoNeiRongChuangXinDianJieZhi += "主要创新点:" + array[4].Trim() + "\n";
                                muBiaoNeiRongChuangXinDianJieZhi += "预期军事价值:" + array[5].Trim() + "\n";

                                //拼装资助对象简介
                                ziZhuDuiXiang += "基本信息:" + dlProjects.getRow(0).getString("UserName") + "," + dlProjects.getRow(0).getString("Sex") + "," + (DateTime.Now.Year - long.Parse(dlProjects.getRow(0).getString("Birthdate").Substring(0, 4))).ToString() + "," + dlProjects.getRow(0).getString("Degree") + "," + dlProjects.getRow(0).getString("UnitName") + "," + dlProjects.getRow(0).getString("JobTitle") + "," + dlProjects.getRow(0).getString("MainResearch") + "\n";
                                ziZhuDuiXiang += "入选人才计划情况:" + array[0].Trim() + "\n";
                                ziZhuDuiXiang += "代表性成果或贡献:" + array[1].Trim() + "\n";
                            }
                        }

                        //处理成果形式
                        projectItem.Add(new KeyValuePair<string, object>("成果形式", new ResultConfigRecord(dlProjects.getRow(0).getString("ResultConfig")).getDescription().Replace("。", string.Empty)));

                        //第一年任务
                        projectItem.Add(new KeyValuePair<string, object>("第一年任务", getFirstTaskText(projectNumber)));

                        //判断推荐方式
                        if (dlProjects.getRow(0).get("ApplicationType") == null || string.IsNullOrEmpty(dlProjects.getRow(0).get("ApplicationType").ToString()) || dlProjects.getRow(0).get("ApplicationType").Equals("单位推荐"))
                        {
                            //"单位推荐"
                            projectItem.Add(new KeyValuePair<string, object>("推荐方式", "单位推荐"));
                            projectItem.Add(new KeyValuePair<string, object>("推荐单位", dlProjects.getRow(0).get("UnitForORG")));
                        }
                        else
                        {
                            //"专家提名"
                            projectItem.Add(new KeyValuePair<string, object>("推荐方式", "专家提名"));
                            projectItem.Add(new KeyValuePair<string, object>("推荐单位", string.Empty));
                        }

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

                        //研究目标，研究内容，主要创新点，预期军事目标
                        projectItem.Add(new KeyValuePair<string, object>("研究目标，研究内容，主要创新点，预期军事目标", muBiaoNeiRongChuangXinDianJieZhi));

                        //资助对象简介
                        projectItem.Add(new KeyValuePair<string, object>("资助对象简介", ziZhuDuiXiang));
                    }

                    //列号
                    int colIndex = 0;

                    //Excel行
                    NPOI.SS.UserModel.IRow row = null;

                    //是否需要输入表头
                    if (isNeedCreateHeader)
                    {
                        isNeedCreateHeader = false;

                        //创建行
                        row = sheet.CreateRow(rowIndex);
                        //输出列名到Excel
                        colIndex = 0;
                        foreach (KeyValuePair<string, object> kvp in projectItem)
                        {
                            //列名
                            //创建列
                            NPOI.SS.UserModel.ICell cell = row.CreateCell(colIndex);
                            //设置样式
                            cell.CellStyle = cellStyleB;
                            //设置数据
                            cell.SetCellValue(kvp.Key);
                            colIndex++;
                        }
                        rowIndex++;
                    }

                    //创建行
                    row = sheet.CreateRow(rowIndex);
                    //输出列值到Excel
                    colIndex = 0;
                    foreach (KeyValuePair<string, object> kvp in projectItem)
                    {
                        //列值
                        //创建列
                        NPOI.SS.UserModel.ICell cell = row.CreateCell(colIndex);
                        //设置样式
                        cell.CellStyle = cellStyleA;
                        //设置数据
                        //判断是否为空
                        if (kvp.Value != null)
                        {
                            //不为空
                            //判断是否为RTF内容
                            if (kvp.Value.GetType().Name.Equals(typeof(NPOI.XSSF.UserModel.XSSFRichTextString).Name))
                            {
                                //RTF内容
                                cell.SetCellValue((NPOI.XSSF.UserModel.XSSFRichTextString)kvp.Value);
                            }
                            else
                            {
                                //文本内容
                                cell.SetCellValue(kvp.Value.ToString());
                            }
                        }
                        else
                        {
                            //为空
                            cell.SetCellValue(string.Empty);
                        }
                        colIndex++;
                    }
                    rowIndex++;
                }

                //Excel列宽自动适应
                for (int k = 0; k < sheet.GetRow(0).Cells.Count; k++)
                {
                    sheet.AutoSizeColumn(k);
                }

                //输出到Excel文件
                workbook.Write(memoryStream);
                File.WriteAllBytes(excelFile, memoryStream.ToArray());

                MessageBox.Show("导出完成！路径：" + excelFile, "提示");
                //打开Excel文件
                System.Diagnostics.Process.Start(excelFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show("对不起，导出失败！Ex:" + ex.ToString());
            }
        }

        /// <summary>
        /// 获得第一年任务的纯文本
        /// </summary>
        /// <param name="projectNumber">项目编号</param>
        /// <returns></returns>
        private object getFirstTaskText(string projectNumber)
        {
            //RTF文本
            string rtfText = string.Empty;

            //查找指定申报包
            string[] subDirs = Directory.GetDirectories(MainForm.Config.TotalDir);
            if (subDirs != null)
            {
                //循环子目录,查找与项目编号相符的目录名称
                foreach (string s in subDirs)
                {
                    //获得目录信息
                    DirectoryInfo di = new DirectoryInfo(s);

                    //判断目录是否符合条件
                    if (di.Name.Contains(projectNumber) && di.Name.Length >= 12)
                    {
                        //查找子文件
                        string[] subFiles = Directory.GetFiles(s);
                        if (subFiles != null && subFiles.Length >= 1)
                        {
                            //解压这个Zip包
                            string destDir = Path.Combine(DBTempDir, projectNumber);

                            //判断第一年研究任务.rtf这个文件是否存在,如果存在则说明之前解压过
                            if (File.Exists(Path.Combine(destDir, Path.Combine(new DirectoryInfo(Directory.GetDirectories(destDir)[0]).Name, "第一年研究任务.rtf"))))
                            {
                                //存在这个文件,直接读取
                                rtfText = File.ReadAllText(Path.Combine(destDir, Path.Combine(new DirectoryInfo(Directory.GetDirectories(destDir)[0]).Name, "第一年研究任务.rtf")));
                                break;
                            }
                            else
                            {
                                //不存在这个文件,先解压再读取

                                //创建临时目录
                                try
                                {
                                    Directory.CreateDirectory(destDir);
                                }
                                catch (Exception ex) { MainForm.writeLog(ex.ToString()); }

                                //解压这个包
                                new NdeyMyDataUnZip().UnZipFile(subFiles[0], destDir, string.Empty, true);

                                //判断第一年研究任务.rtf这个文件是否存在
                                if (File.Exists(Path.Combine(destDir, Path.Combine(new DirectoryInfo(Directory.GetDirectories(destDir)[0]).Name, "第一年研究任务.rtf"))))
                                {
                                    //存在这个文件,直接读取
                                    rtfText = File.ReadAllText(Path.Combine(destDir, Path.Combine(new DirectoryInfo(Directory.GetDirectories(destDir)[0]).Name, "第一年研究任务.rtf")));
                                    break;
                                }
                            }
                        }
                    }
                }

                //如果RTF内容非空,则需要获得纯文本
                if (!string.IsNullOrEmpty(rtfText))
                {
                    //将RTF文本设置到RichTextBox对象
                    RichTextBox rtb = new RichTextBox();
                    rtb.Rtf = rtfText;

                    //获得纯文本
                    rtfText = rtb.Text;
                }
            }

            return rtfText;
        }

        private void btnSelectTotalDir_Click(object sender, EventArgs e)
        {
            //显示目录选择对话框
            fbdTotalDirSelect.SelectedPath = MainForm.Config.TotalDir;
            if (fbdTotalDirSelect.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //设置总目录
                MainForm.Config.TotalDir = fbdTotalDirSelect.SelectedPath;
                //保存配置
                saveConfig();

                //显示总目录
                refreshTotalDir();
            }
        }

        private void btnImportAll_Click(object sender, EventArgs e)
        {
            //显示导入窗体
            ImporterForm form = new ImporterForm(true);
            form.ShowDialog();

            //刷新Catalog列表
            reloadCatalogList();
        }

        private void btnImportWithSelectedList_Click(object sender, EventArgs e)
        {
            //显示导入窗体
            ImporterForm form = new ImporterForm(false);
            form.ShowDialog();

            //刷新Catalog列表
            reloadCatalogList();
        }

        private void btnSelectUnZipDir_Click(object sender, EventArgs e)
        {
            //显示目录选择对话框
            fbdTotalDirSelect.SelectedPath = MainForm.Config.UnZipDir;
            if (fbdTotalDirSelect.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //设置解压目录
                MainForm.Config.UnZipDir = fbdTotalDirSelect.SelectedPath;

                //保存配置
                saveConfig();
            }
        }

        private void btnUnZipAll_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(MainForm.Config.UnZipDir))
            {
                MessageBox.Show("对不起,请先选择解压目录!");
                return;
            }

            UnZipDocsForm zdf = new UnZipDocsForm(true);
            zdf.ShowDialog();
        }

        private void btnUnZipWithSelectedList_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(MainForm.Config.UnZipDir))
            {
                MessageBox.Show("对不起,请先选择解压目录!");
                return;
            }

            UnZipDocsForm zdf = new UnZipDocsForm(false);
            zdf.ShowDialog();
        }

        /// <summary>
        /// 写入日志
        /// </summary>
        /// <param name="text"></param>
        public static void writeLog(string text)
        {
            //程序目录
            string path = AppDomain.CurrentDomain.BaseDirectory;

            //日志目录
            path = System.IO.Path.Combine(path, "Logs\\");

            //判断是否需要创建目录
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }

            //生成日志文件名称
            string fileFullName = System.IO.Path.Combine(path, string.Format("{0}.txt", DateTime.Now.ToString("yyyyMMdd-HH")));
            
            //写日志
            using (StreamWriter output = System.IO.File.AppendText(fileFullName))
            {
                //写日志
                output.WriteLine(DateTime.Now.ToString() + ":" + text);

                //关闭文件
                output.Close();
            }
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

    /// <summary>
    /// 程序配置文件
    /// </summary>
    public class MainConfig
    {
        /// <summary>
        /// 申报包总目录
        /// </summary>
        public string TotalDir { get; set; }

        /// <summary>
        /// 文档解压目录
        /// </summary>
        public string UnZipDir { get; set; }
    }
}