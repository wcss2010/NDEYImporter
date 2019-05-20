using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace NDEYImporter.Forms
{
    /// <summary>
    /// 进度条窗体
    /// </summary>
    public partial class ProgressForm : Form
    {
        /// <summary>
        /// 工作线程
        /// </summary>
        private BackgroundWorker worker = new BackgroundWorker();

        /// <summary>
        /// 工作事件
        /// </summary>
        private EventHandler ehDynamicMethod = null;

        /// <summary>
        /// 错误计数
        /// </summary>
        public static int errorCount;

        /// <summary>
        /// 是否显示错误日志
        /// </summary>
        public static bool isNeedShowLog = true;

        public ProgressForm()
        {
            InitializeComponent();

            worker.WorkerSupportsCancellation = true;
            worker.DoWork += worker_DoWork;
        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            //运行工作事件
            if (ehDynamicMethod != null)
            {
                try
                {
                    ehDynamicMethod(this, new EventArgs());
                }
                catch (Exception ex)
                {
                    MainForm.writeLog(ex.ToString());
                }
            }

            //检查是否需要显示错误日志
            if (errorCount >= 1 && isNeedShowLog)
            {
                //错误数量清空
                errorCount = 0;

                //日志路径
                string logFile = System.IO.Path.Combine(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs"), string.Format("{0}.txt", DateTime.Now.ToString("yyyyMMdd")));

                //询问是否需要显示日志
                if (MessageBox.Show("是否需要打开错误日志文件?", "提示", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                {
                    //打开日志文件
                    try
                    {
                        System.Diagnostics.Process.Start(logFile);
                    }
                    catch (Exception ex) { }
                }
            }
        }

        /// <summary>
        /// 运行代码
        /// </summary>
        /// <param name="total">总进度</param>
        /// <param name="cur">当前进度</param>
        /// <param name="ehDynamic">工作事件</param>
        public void run(int total, int cur, EventHandler ehDynamic)
        {
            //设置进度
            pbProgress.Maximum = total;
            pbProgress.Value = cur;

            //设置工作事件
            ehDynamicMethod = ehDynamic;

            //检查工作线程是否正在运行
            if (worker.IsBusy)
            {
                return;
            }
            else
            {
                //运行工作线程
                worker.RunWorkerAsync();
            }
        }

        /// <summary>
        /// 设置当前进度
        /// </summary>
        /// <param name="current">进度值</param>
        public void reportProgress(int current,string text)
        {
            if (IsHandleCreated)
            {
                Invoke(new MethodInvoker(delegate()
                    {
                        //设置进度
                        pbProgress.Value = current;

                        //设置文本
                        lblProgressText.Text = text;
                    }));
            }
        }
    }
}