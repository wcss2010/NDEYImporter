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
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            NDEYImporter.Util.DBImporter.appendNdeyToLocalDB(@"C:\Users\wcss\Desktop\AA\myData.db");

            NDEYImporter.Util.DBImporter.appendNdeyToLocalDB(@"C:\Users\wcss\Desktop\BB\myData.db");
        }
    }
}