using NDEYImporter.DB;
using NDEYImporter.Util;
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
    public partial class CheckNeedReplaceForm : Form
    {
        /// <summary>
        /// 是否需要更新替换字典在改变替换列表中项目状态时
        /// </summary>
        private bool isNeedUpdateDict = true;

        /// <summary>
        /// 替换字典，用于表示哪些项目可以替换(Key=项目ID,Value=是否替换)
        /// </summary>
        private Dictionary<string, bool> replaceDict = new Dictionary<string, bool>();

        /// <summary>
        /// 是否导入所有申报包
        /// </summary>
        public bool IsImportAllPackage { get; private set; }

        public CheckNeedReplaceForm(bool isImportAll)
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
            Close();
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

                    //判断是否需要替换
                    if (replaceDict.ContainsKey(projectNumber))
                    {
                        //需要替换

                        //判断是否替换这个项目
                        if (replaceDict[projectNumber])
                        {
                            //需要替换，添加到ImportList列表中
                            importList.Add(Path.Combine(MainForm.Config.TotalDir, tn.Text));
                        }
                        else
                        {
                            //不需要替换
                            continue;
                        }
                    }
                    else
                    {
                        //不需要替换
                        importList.Add(Path.Combine(MainForm.Config.TotalDir, tn.Text));
                    }
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

                        //获取文件列表
                        string[] subFiles = Directory.GetFiles(pDir);
                        //判断是否存在申报包
                        if (subFiles != null && subFiles.Length >= 1 && subFiles[0].EndsWith(".zip"))
                        {
                            //报告进度
                            pf.reportProgress(progressVal, string.Empty);

                            //生成一个临时路径
                            string destDir = System.IO.Path.Combine(MainForm.DBTempDir, DateTime.Now.Ticks.ToString());
                            //解压ZIP包
                            FileZipOpr fzo = new FileZipOpr();
                            fzo.UnZipFile(subFiles[0], destDir, string.Empty, true);

                            //判断申报包是否有效
                            //生成DB文件路径
                            string dbFile = System.IO.Path.Combine(destDir, @"Files\myData.db");
                            //判断文件是否存在
                            if (System.IO.File.Exists(dbFile))
                            {
                                //有效的申报包
                                try
                                {
                                    //报告进度
                                    pf.reportProgress(progressVal, projectNumber + "_开始导入");

                                    //导入数据并返回项目Id
                                    DBImporter.importDB(projectNumber, dbFile);

                                    //报告进度
                                    pf.reportProgress(progressVal, projectNumber + "_结束导入");
                                }
                                catch (Exception ex) { }                                
                            }
                        }
                    }

                    //检查是否已创建句柄，并调用委托执行UI方法
                    if (IsHandleCreated)
                    {
                        Invoke(new MethodInvoker(delegate()
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
        /// 刷新替换列表
        /// </summary>
        private void reloadReplaceList()
        {
            //锁定替换列表点击CheckBox时更新替换字典的功能
            isNeedUpdateDict = false;

            //清空替换列表
            lvErrorList.Items.Clear();
            //显示替换列表数
            foreach (KeyValuePair<string, bool> kvp in replaceDict)
            {
                //列表项
                ListViewItem lvi = new ListViewItem();
                lvi.Text = kvp.Key;
                lvi.Checked = kvp.Value;

                //向替换列表添加
                lvErrorList.Items.Add(lvi);
            }

            //解锁替换列表点击CheckBox时更新替换字典的功能
            isNeedUpdateDict = true;
        }

        private void tvSubDirs_AfterCheck(object sender, TreeViewEventArgs e)
        {
            //读取目录名称中的项目编号
            string projectNumber = e.Node.Text.Substring(0, 11);

            //根据项目编号查询项目数量
            long projectCount = ConnectionManager.Context.table("Catalog").where("ProjectNumber='" + projectNumber + "'").select("count(*)").getValue<long>(0);
            //判断这个项目是否被导入过
            if (projectCount >= 1)
            {
                replaceDict[projectNumber] = true;
            }

            //刷新替换列表
            reloadReplaceList();
        }

        private void lvErrorList_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            //改变选择状态
            if (isNeedUpdateDict)
            {
                if (replaceDict.ContainsKey(e.Item.Text))
                {
                    replaceDict[e.Item.Text] = e.Item.Checked;
                }
            }
        }
    }
}