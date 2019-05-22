using NDEYImporter.Util;
using Noear.Weed;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace NDEYImporter.Forms
{
    public partial class UnZipDocsForm : Form
    {
        /// <summary>
        /// 缺少文件的日志
        /// </summary>
        private string missFileLogPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "缺少的文件.xlsx");

        /// <summary>
        /// 缺少文件列表
        /// </summary>
        private List<List<KeyValuePair<string, object>>> errorFileList = new List<List<KeyValuePair<string, object>>>();
        
        /// <summary>
        /// 是否导入所有申报包
        /// </summary>
        public bool IsImportAllPackage { get; private set; }

        public UnZipDocsForm(bool isImportAll)
        {
            InitializeComponent();

            //是否导入所有申报包
            IsImportAllPackage = isImportAll;

            //初始化
            init();
        }

        /// <summary>
        /// 初始化
        /// </summary>
        private void init()
        {
            //刷新子目录列表
            string[] dirs = System.IO.Directory.GetDirectories(MainForm.Config.TotalDir);
            //添加子目录到树
            foreach (string s in dirs)
            {
                //目录信息
                System.IO.DirectoryInfo fi = new System.IO.DirectoryInfo(s);

                //判断目录名是否符合要求(2019-ZQ-001_xxxxxxxxx)
                if (fi.Name.Contains("-ZQ-") && fi.Name.Length >= 13)
                {
                    //添加目录名称
                    tvSubDirs.Nodes.Add(fi.Name);
                }
            }
            //判断是否为全部导入，如果是则选择所有子目录
            if (IsImportAllPackage)
            {
                //选择所有子目录
                foreach (TreeNode tn in tvSubDirs.Nodes)
                {
                    tn.Checked = true;
                }
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = System.Windows.Forms.DialogResult.Cancel;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            //需要替换的申报包列表
            List<string> importList = new List<string>();

            #region 查的需要导入的路径
            foreach (TreeNode tn in tvSubDirs.Nodes)
            {
                //判断是否被选中
                if (tn.Checked)
                {
                    //读取目录名称中的项目编号
                    string projectNumber = tn.Text.Substring(0, 11);

                    //将需要处理的申报包添加到列表中
                    importList.Add(Path.Combine(MainForm.Config.TotalDir, tn.Text));

                }
            }
            #endregion

            //开始导入
            ProgressForm pf = new ProgressForm();
            pf.Show();
            pf.run(importList.Count, 0, new EventHandler(delegate(object senders, EventArgs ee)
            {
                //进度数值
                int progressVal = 0;
                //导入
                foreach (string pDir in importList)
                {
                    try
                    {
                        progressVal++;

                        //读取项目ID
                        string projectNumber = new DirectoryInfo(pDir).Name.Substring(0, 11);

                        //报告进度
                        pf.reportProgress(progressVal, string.Empty);

                        //报告进度
                        pf.reportProgress(progressVal, projectNumber + "_开始解压");

                        //解压ZIP包
                        unZipDocFiles(projectNumber);

                        //报告进度
                        pf.reportProgress(progressVal, projectNumber + "_结束解压");
                    }
                    catch (Exception ex)
                    {
                        MainForm.writeLog(ex.ToString());
                    }
                }

                //尝试输出Excel日志
                if (errorFileList != null && errorFileList.Count >= 1)
                {
                    writeErrorExcelFile();
                }

                //检查是否已创建句柄，并调用委托执行UI方法
                if (pf.IsHandleCreated)
                {
                    pf.Invoke(new MethodInvoker(delegate()
                    {
                        //打开缺少文件日志
                        pf.TopMost = false;
                        if (File.Exists(missFileLogPath))
                        {
                            MessageBox.Show("缺少文件列表生成完成!路径:" + missFileLogPath);

                            try
                            {
                                System.Diagnostics.Process.Start(missFileLogPath);
                            }
                            catch (Exception ex) { }
                        }

                        try
                        {
                            //刷新Catalog列表
                            MainForm.Instance.reloadCatalogList();

                            //关闭进度窗口
                            pf.Close();
                        }
                        catch (Exception ex)
                        {
                            MainForm.writeLog(ex.ToString());
                        }
                    }));
                }
            }));

            //关闭窗口
            Close();
        }

        /// <summary>
        /// 从申报包中解压文档
        /// </summary>
        /// <param name="projectNumber">项目编号</param>
        private void unZipDocFiles(string projectNumber)
        {
            try
            {
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
                            MainForm.writeLog("开始进行项目" + projectNumber + "的解包......");

                            //查找子文件
                            string[] subFiless = Directory.GetFiles(s);
                            if (subFiless != null && subFiless.Length >= 1)
                            {
                                //解压这个Zip包
                                string pkgDir = Path.Combine(MainForm.Config.UnZipDir, projectNumber);

                                //删除解压目录
                                try
                                {
                                    if (Directory.Exists(pkgDir))
                                    {
                                        Directory.Delete(pkgDir, true);
                                    }
                                }
                                catch (Exception ex) { MainForm.writeLog(ex.ToString()); }

                                //创建解压目录
                                try
                                {
                                    Directory.CreateDirectory(pkgDir);
                                }
                                catch (Exception ex) { MainForm.writeLog(ex.ToString()); }

                                //ZIP文件
                                string pkgZipFile = string.Empty;

                                //查找ZIP文件
                                foreach (string sssss in subFiless)
                                {
                                    if (sssss.EndsWith(".zip"))
                                    {
                                        pkgZipFile = sssss;
                                        break;
                                    }
                                }

                                //判断ZIP文件是否存在
                                if (pkgZipFile != null && pkgZipFile.EndsWith(".zip"))
                                {
                                    MainForm.writeLog("项目" + projectNumber + "的解包操作，开始ZIP文件解压");

                                    //解压这个包
                                    new NdeyDocFilesUnZip().UnZipFile(pkgZipFile, pkgDir, string.Empty, true);

                                    MainForm.writeLog("项目" + projectNumber + "的解包操作，结束ZIP文件解压");

                                    //文件目录
                                    string fileDir = pkgDir;

                                    //判断申报包是否有效
                                    //生成DB文件路径
                                    string dbFile = System.IO.Path.Combine(fileDir, "myData.db");
                                    //判断文件是否存在
                                    if (System.IO.File.Exists(dbFile))
                                    {
                                        //SQLite数据库工厂
                                        System.Data.SQLite.SQLiteFactory factory = new System.Data.SQLite.SQLiteFactory();
                                        //NDEY数据库连接
                                        Noear.Weed.DbContext context = new Noear.Weed.DbContext("main", "Data Source = " + dbFile, factory);

                                        try
                                        {
                                            MainForm.writeLog("项目" + projectNumber + "的解包操作，读取论文详细，科技奖项，专利情况信息...");

                                            //提取论文详细
                                            DataList dlRTreatises = context.table("RTreatises").orderBy("RTreatisesOrder desc").select("RTreatisesName,RTreatisesPDF").getDataList();
                                            //提取科技奖项
                                            DataList dlTechnologyAwards = context.table("TechnologyAwards").orderBy("TechnologyAwardsOrder desc").select("TechnologyAwardsPName,TechnologyAwardsPDF").getDataList();
                                            //提取专利情况
                                            DataList dlNDPatent = context.table("NDPatent").orderBy("NDPatentOrder desc").select("NDPatentName,NDPatentPDF").getDataList();
                                            //提取推荐意见
                                            DataList dlBaseInfor = context.table("BaseInfor").select("UnitRecommend,ExpertRecommend1,ExpertRecommend2,ExpertRecommend3").getDataList();

                                            //提取UnitID
                                            string unitID = context.table("BaseInfor").select("UnitID").getValue<string>(string.Empty);

                                            string personName = context.table("BaseInfor").select("UserName").getValue<string>(string.Empty);

                                            MainForm.writeLog("项目" + projectNumber + "的解包操作，开始替换附件名称...");
                                            //附件序号
                                            int fileIndex = 0;
                                            //整理附件名称
                                            foreach (DataItem item in dlRTreatises.getRows())
                                            {
                                                //获得文件扩展名
                                                string extName = new FileInfo(Path.Combine(fileDir, item.getString("RTreatisesPDF"))).Extension;

                                                //文件序号+1
                                                fileIndex++;

                                                //文件重命名
                                                try
                                                {
                                                    renameFile(Path.Combine(fileDir, item.getString("RTreatisesPDF")), fileIndex, item.getString("RTreatisesName"), extName);
                                                }
                                                catch (Exception ex)
                                                {
                                                    MainForm.writeLog(ex.ToString());

                                                    addErrorFile(projectNumber, item.getString("RTreatisesPDF"));
                                                }
                                            }
                                            foreach (DataItem item in dlTechnologyAwards.getRows())
                                            {
                                                //获得文件扩展名
                                                string extName = new FileInfo(Path.Combine(fileDir, item.getString("TechnologyAwardsPDF"))).Extension;

                                                //文件序号+1
                                                fileIndex++;

                                                //文件重命名
                                                try
                                                {
                                                    renameFile(Path.Combine(fileDir, item.getString("TechnologyAwardsPDF")), fileIndex, item.getString("TechnologyAwardsPName"), extName);
                                                }
                                                catch (Exception ex)
                                                {
                                                    MainForm.writeLog(ex.ToString());

                                                    addErrorFile(projectNumber, item.getString("TechnologyAwardsPDF"));
                                                }
                                            }
                                            foreach (DataItem item in dlNDPatent.getRows())
                                            {
                                                //获得文件扩展名
                                                string extName = new FileInfo(Path.Combine(fileDir, item.getString("NDPatentPDF"))).Extension;

                                                //文件序号+1
                                                fileIndex++;

                                                //文件重命名
                                                try
                                                {
                                                    renameFile(Path.Combine(fileDir, item.getString("NDPatentPDF")), fileIndex, item.getString("NDPatentName"), extName);
                                                }
                                                catch (Exception ex)
                                                {
                                                    MainForm.writeLog(ex.ToString());

                                                    addErrorFile(projectNumber, item.getString("NDPatentPDF"));
                                                }
                                            }

                                            //整理推荐意见附件
                                            string unitFileA = dlBaseInfor.getRow(0).get("UnitRecommend") != null ? dlBaseInfor.getRow(0).get("UnitRecommend").ToString() : string.Empty;
                                            string personFileA = dlBaseInfor.getRow(0).get("ExpertRecommend1") != null ? dlBaseInfor.getRow(0).get("ExpertRecommend1").ToString() : string.Empty;
                                            string personFileB = dlBaseInfor.getRow(0).get("ExpertRecommend2") != null ? dlBaseInfor.getRow(0).get("ExpertRecommend2").ToString() : string.Empty;
                                            string personFileC = dlBaseInfor.getRow(0).get("ExpertRecommend3") != null ? dlBaseInfor.getRow(0).get("ExpertRecommend3").ToString() : string.Empty;

                                            //查找单位推荐意见附件
                                            if (File.Exists(Path.Combine(fileDir, unitFileA)))
                                            {
                                                //文件序号+1
                                                fileIndex++;

                                                //文件重命名
                                                try
                                                {
                                                    renameFile(Path.Combine(fileDir, unitFileA), fileIndex, "单位推荐意见", new FileInfo(Path.Combine(fileDir, unitFileA)).Extension);
                                                }
                                                catch (Exception ex)
                                                {
                                                    MainForm.writeLog(ex.ToString());

                                                    addErrorFile(projectNumber, unitFileA);
                                                }
                                            }

                                            //查找专家提名附件
                                            if (File.Exists(Path.Combine(fileDir, personFileA)))
                                            {
                                                //文件序号+1
                                                fileIndex++;

                                                //文件重命名
                                                try
                                                {
                                                    renameFile(Path.Combine(fileDir, personFileA), fileIndex, "专家提名", new FileInfo(Path.Combine(fileDir, personFileA)).Extension);
                                                }
                                                catch (Exception ex)
                                                {
                                                    MainForm.writeLog(ex.ToString());

                                                    addErrorFile(projectNumber, personFileA);
                                                }
                                            }

                                            //查找专家提名附件
                                            if (File.Exists(Path.Combine(fileDir, personFileB)))
                                            {
                                                //文件序号+1
                                                fileIndex++;

                                                //文件重命名
                                                try
                                                {
                                                    renameFile(Path.Combine(fileDir, personFileB), fileIndex,"专家提名",new FileInfo(Path.Combine(fileDir, personFileB)).Extension);
                                                }
                                                catch (Exception ex)
                                                {
                                                    MainForm.writeLog(ex.ToString());

                                                    addErrorFile(projectNumber, personFileB);
                                                }
                                            }

                                            //查找专家提名附件
                                            if (File.Exists(Path.Combine(fileDir, personFileC)))
                                            {
                                                //文件序号+1
                                                fileIndex++;

                                                //文件重命名
                                                try
                                                {
                                                    renameFile(Path.Combine(fileDir, personFileC), fileIndex,"专家提名",new FileInfo(Path.Combine(fileDir, personFileC)).Extension);
                                                }
                                                catch (Exception ex)
                                                {
                                                    MainForm.writeLog(ex.ToString());

                                                    addErrorFile(projectNumber, personFileC);
                                                }
                                            }

                                            MainForm.writeLog("项目" + projectNumber + "的解包操作，结束替换附件名称...");

                                            MainForm.writeLog("项目" + projectNumber + "的解包操作，开始处理保密资质附件和申报书文档转PDF...");

                                            //是否已经找到Doc文件
                                            bool isFoundDocFile = false;

                                            //整理保密资质命名,查找Doc文件
                                            string[] extFiles = Directory.GetFiles(fileDir);
                                            foreach (string sss in extFiles)
                                            {
                                                FileInfo fii = new FileInfo(sss);
                                                if (fii.Name.StartsWith("extFile_"))
                                                {
                                                    try
                                                    {
                                                        //文件序号+1
                                                        fileIndex++;

                                                        //移动文件
                                                        renameFile(sss,fileIndex,"保密资质复印件",".png");
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        MainForm.writeLog(ex.ToString());

                                                        addErrorFile(projectNumber, fii.Name);
                                                    }
                                                }
                                                else if (fii.Name.EndsWith(".doc"))
                                                {
                                                    if (fii.Name.StartsWith(unitID) && fii.Name.Contains(personName))
                                                    {
                                                        //找到Doc文件
                                                        isFoundDocFile = true;

                                                        //真实的申报书路径
                                                        string destDocFile = Path.Combine(fii.DirectoryName, "项目申报书.doc");

                                                        //文件改名
                                                        try
                                                        {
                                                            File.Move(sss, destDocFile);
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            MainForm.writeLog(ex.ToString());

                                                            addErrorFile(projectNumber, fii.Name);
                                                        }

                                                        //转换成PDF
                                                        convertToPDF(destDocFile);
                                                    }
                                                    else
                                                    {
                                                        if (fii.Name.EndsWith("项目申报书.doc"))
                                                        {
                                                            try
                                                            {
                                                                File.Delete(sss);
                                                            }
                                                            catch (Exception ex) { MainForm.writeLog(ex.ToString()); }
                                                        }
                                                    }
                                                }
                                            }

                                            //判断是否找到Doc文件
                                            if (isFoundDocFile)
                                            {
                                                //找到DOC文件
                                                //MainForm.writeLog("项目申报书已找到！");
                                            }
                                            else
                                            {
                                                //没有找到Doc文件
                                                MainForm.writeLog("对不起，没有找到项目申报书！");
                                                //输出缺失的文件
                                                addErrorFile(projectNumber, "项目申报书.doc");
                                            }

                                            MainForm.writeLog("项目" + projectNumber + "的解包操作，结束处理保密资质附件和申报书文档转PDF...");
                                        }
                                        catch (Exception ex)
                                        {
                                            MainForm.writeLog(ex.ToString());
                                        }
                                        finally
                                        {
                                            factory.Dispose();
                                            context.Dispose();
                                            context = null;

                                            //**链接close（）和dispose（）之后任然不能释放与db文件的连接,原因是sqlite在执行SQLiteConnectionHandle.Dispose()操作时候，其实并没有真正的释放连接，只有显式调用 CLR垃圾回收之后才能真正释放连接
                                            GC.Collect();
                                            GC.WaitForPendingFinalizers();

                                            //删除数据库文件
                                            try
                                            {
                                                File.Delete(dbFile);
                                            }
                                            catch (Exception ex) { }
                                        }
                                    }
                                    else
                                    {
                                        MainForm.writeLog("没有找到DB文件__" + projectNumber);
                                    }
                                }
                                else
                                {
                                    MainForm.writeLog("没有找到ZIP文件__" + projectNumber);
                                }
                            }
                            else
                            {
                                MainForm.writeLog("没有找到ZIP文件__" + projectNumber);
                            }

                            MainForm.writeLog("结束进行项目" + projectNumber + "的解包......");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MainForm.writeLog("项目编号" + projectNumber + "解包错误，Ex:" + ex.ToString());
            }
        }

        /// <summary>
        /// 输出重命名错误的文件
        /// </summary>
        /// <param name="projectNumber">项目编号</param>
        /// <param name="fileName">文件名称</param>
        private void addErrorFile(string projectNumber, string fileName)
        {
            try
            {
                //错误文件的位置
                string destFile = Path.Combine(Path.Combine(MainForm.Config.UnZipDir, projectNumber), fileName);

                List<KeyValuePair<string, object>> row = new List<KeyValuePair<string, object>>();
                row.Add(new KeyValuePair<string, object>("项目编号", projectNumber));
                row.Add(new KeyValuePair<string, object>("缺少文件", fileName));
                row.Add(new KeyValuePair<string, object>("时间", DateTime.Now.ToString()));
                row.Add(new KeyValuePair<string, object>("类型", File.Exists(destFile) ? "Move错误" : "文件不存在"));
                errorFileList.Add(row);
            }
            catch (Exception ex)
            {
                MainForm.writeLog(ex.ToString());
            }
        }

        /// <summary>
        /// 写错误日志的Excel文件
        /// </summary>
        private void writeErrorExcelFile()
        {
            try
            {
                //输出的Excel路径
                string excelFile = missFileLogPath;

                //Excel数据
                MemoryStream memoryStream = new MemoryStream();
                //创建Workbook
                NPOI.XSSF.UserModel.XSSFWorkbook workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();

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

                //输出Sheet表格
                MainForm.writeSheet(workbook, cellStyleA, cellStyleB, errorFileList);

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
        /// 重命名附件
        /// </summary>
        /// <param name="source">源路径</param>
        /// <param name="fileIndex">文件序号</param>
        /// <param name="itemName">项目名称</param>
        /// <param name="extName">扩展名称</param>
        private void renameFile(string source, int fileIndex,string itemName,string extName)
        {
            //目标路径
            string destPath = string.Empty;

            //检查itemName是不是大于8个字符
            if (itemName != null && itemName.Length > 8)
            {
                //截取前8个字符
                itemName = itemName.Substring(0, 8);
            }

            //去掉itemName中的非法字符
            string[] errorString = new string[] { "\\", "/", ":", "*", "?", "\"", "<", ">", "|" };
            foreach (string essss in errorString)
            {
                itemName = itemName.Replace(essss, string.Empty);
            }

            //生成目标路径
            destPath = Path.Combine(new FileInfo(source).DirectoryName, "附件" + fileIndex + "_" + itemName + extName);

            //文件重命名
            File.Move(source, destPath);
        }

        /// <summary>
        /// 转换到PDF文件
        /// </summary>
        /// <param name="destDocFile"></param>
        private void convertToPDF(string destDocFile)
        {
            try
            {
                //Document文档对象
                Aspose.Words.Document xDoc = new Aspose.Words.Document(destDocFile);

                //获得附件表格
                Aspose.Words.Tables.Table tablee = (Aspose.Words.Tables.Table)xDoc.GetChild(Aspose.Words.NodeType.Table, xDoc.GetChildNodes(Aspose.Words.NodeType.Table, true).Count - 1, true);

                //行与
                int rowIndex = 0;

                //查找需要替换位置
                if (tablee != null && tablee.Rows != null)
                {
                    //表格有效
                    foreach (Aspose.Words.Tables.Row row in tablee.Rows)
                    {
                        if (row.Cells[0].GetText().Contains("序号"))
                        {
                            continue;
                        }
                        else
                        {
                            //行号+1
                            rowIndex++;

                            //清理单位格内容
                            row.Cells[0].Paragraphs.Clear();

                            //写行号
                            Aspose.Words.Paragraph p = new Aspose.Words.Paragraph(xDoc);
                            p.AppendChild(new Aspose.Words.Run(xDoc, rowIndex.ToString()));
                            row.Cells[0].AppendChild(p);
                        }
                    }
                }
                else
                {
                    //表格无效
                    MainForm.writeLog("没有找到附件清单，当前文件" + destDocFile + ",表格数量:" + xDoc.GetChildNodes(Aspose.Words.NodeType.Table, true).Count);
                }

                //保存为PDF
                xDoc.Save(Path.Combine(new FileInfo(destDocFile).DirectoryName, "项目申报书.pdf"), Aspose.Words.SaveFormat.Pdf);
            }
            catch (Exception ex)
            {
                MainForm.writeLog(ex.ToString());
            }

            //删除Doc文档
            try
            {
                File.Delete(destDocFile);
            }
            catch (Exception ex)
            {
                MainForm.writeLog(ex.ToString());
            }
        }
    }
}