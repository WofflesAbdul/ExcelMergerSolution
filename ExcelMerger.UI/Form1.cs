using System;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelMerger.UI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            toolStripLabel1.Text += " " + Assembly.GetExecutingAssembly().GetName().Version;
        }
    }
}
