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
                ehDynamicMethod(this, new EventArgs());
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