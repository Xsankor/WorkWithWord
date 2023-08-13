using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form2 : Form
    {
        string savePlace;
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            savePlace = Directory.GetCurrentDirectory();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 form_01 = new Form1(savePlace);
            form_01.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form3 form_02 = new Form3(savePlace);
            form_02.ShowDialog();
        }

        private void задатьМестоСохраненияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.ShowDialog();
            savePlace = folderBrowserDialog.SelectedPath;
        }
    }
}
