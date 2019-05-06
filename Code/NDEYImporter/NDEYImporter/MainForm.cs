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

        private void btnTest_Click(object sender, EventArgs e)
        {
            NDEYImporter.Util.DBImporter.insertToLocalDB(@"C:\Users\wcss\Desktop\BB\myData.db");

            //NDEYImporter.Util.DBImporter.appendNdeyToLocal(@"C:\Users\wcss\Desktop\BB\myData.db");
        }
    }
}