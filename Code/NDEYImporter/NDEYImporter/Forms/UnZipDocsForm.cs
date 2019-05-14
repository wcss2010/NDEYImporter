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

                //检查是否已创建句柄，并调用委托执行UI方法
                if (pf.IsHandleCreated)
                {
                    pf.Invoke(new MethodInvoker(delegate()
                    {
                        //刷新Catalog列表
                        MainForm.Instance.reloadCatalogList();

                        //关闭进度窗口
                        pf.Close();
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
                        string[] subFiless = Directory.GetFiles(s);
                        if (subFiless != null && subFiless.Length >= 1)
                        {
                            //解压这个Zip包
                            string pkgDir = Path.Combine(MainForm.Config.UnZipDir, projectNumber);

                            //创建临时目录
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

                                //解压这个包
                                new NdeyDocFilesUnZip().UnZipFile(pkgZipFile, pkgDir, string.Empty, true);

                                //文件目录
                                string fileDir = System.IO.Path.Combine(pkgDir, new DirectoryInfo(Directory.GetDirectories(pkgDir)[0]).Name);

                                //判断申报包是否有效
                                //生成DB文件路径
                                string dbFile = System.IO.Path.Combine(fileDir, "myData.db");
                                //判断文件是否存在
                                if (System.IO.File.Exists(dbFile))
                                {
                                    try
                                    {
                                        //SQLite数据库工厂
                                        System.Data.SQLite.SQLiteFactory factory = new System.Data.SQLite.SQLiteFactory();

                                        //NDEY数据库连接
                                        Noear.Weed.DbContext context = new Noear.Weed.DbContext("main", "Data Source = " + dbFile, factory);

                                        //提取论文详细
                                        DataList dlRTreatises = context.table("RTreatises").select("RTreatisesName,RTreatisesPDF").getDataList();
                                        //提取科技奖项
                                        DataList dlTechnologyAwards = context.table("TechnologyAwards").select("TechnologyAwardsPName,TechnologyAwardsPDF").getDataList();
                                        //提取专利情况
                                        DataList dlNDPatent = context.table("NDPatent").select("NDPatentName,NDPatentPDF").getDataList();

                                        //提取UnitID
                                        string unitID = context.table("BaseInfor").select("UnitID").getValue<string>(string.Empty);

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
                                                File.Move(Path.Combine(fileDir, item.getString("RTreatisesPDF")), Path.Combine(fileDir, "附件" + fileIndex + "_" + item.getString("RTreatisesName") + extName));
                                            }
                                            catch (Exception ex) { MainForm.writeLog(ex.ToString()); }
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
                                                File.Move(Path.Combine(fileDir, item.getString("TechnologyAwardsPDF")), Path.Combine(fileDir, "附件" + fileIndex + "_" + item.getString("TechnologyAwardsPName") + extName));
                                            }
                                            catch (Exception ex) { MainForm.writeLog(ex.ToString()); }
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
                                                File.Move(Path.Combine(fileDir, item.getString("NDPatentPDF")), Path.Combine(fileDir, "附件" + fileIndex + "_" + item.getString("NDPatentName") + extName));
                                            }
                                            catch (Exception ex) { MainForm.writeLog(ex.ToString()); }
                                        }

                                        //整理保密资质命名,查找Doc文件
                                        string[] extFiles = Directory.GetFiles(fileDir);
                                        foreach (string sss in extFiles)
                                        {
                                            FileInfo fii = new FileInfo(sss);
                                            if (fii.Name.StartsWith("extFile_"))
                                            {
                                                try
                                                {
                                                    File.Move(sss, Path.Combine(fileDir, "保密资质复印件.png"));
                                                }
                                                catch (Exception ex)
                                                {
                                                    MainForm.writeLog(ex.ToString());
                                                }
                                            }
                                            else if (fii.Name.EndsWith(".doc"))
                                            {
                                                if (fii.Name.StartsWith(unitID))
                                                {
                                                    //真实的申报书路径
                                                    string destDocFile = Path.Combine(fii.DirectoryName, "项目申报书.doc");

                                                    //文件改名
                                                    try
                                                    {
                                                        File.Move(sss, destDocFile);
                                                    }
                                                    catch (Exception ex) { MainForm.writeLog(ex.ToString()); }

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
                                    }
                                    catch (Exception ex)
                                    {
                                        MainForm.writeLog(ex.ToString());
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 转换到PDF文件
        /// </summary>
        /// <param name="destDocFile"></param>
        private void convertToPDF(string destDocFile)
        {
            
        }
    }
}