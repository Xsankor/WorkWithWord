using System;
using System.IO;
using System.Windows.Forms;

namespace WorkWithWord
{
    public partial class Form2 : Form
    {
        string savePlace, jsonString, path;
        JsonBody jsonBody;

        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            jsonBody = new JsonBody();

            path = Directory.GetCurrentDirectory() + @"\setting.json";

            if (File.Exists(path))
            {
                jsonString = File.ReadAllText(path);
                File.WriteAllText(path, jsonString);
            }
            else 
            {
                jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(jsonBody);
            }
            
            jsonBody = Newtonsoft.Json.JsonConvert.DeserializeObject<JsonBody>(jsonString);
            savePlace = jsonBody.FileNamePath;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 form_01 = new Form1(savePlace);
            form_01.ShowDialog();
        }

        private void открытьМестоСохраненияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe", savePlace);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form3 form_02 = new Form3(savePlace);
            form_02.ShowDialog();
        }

        private void задатьМестоСохраненияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем выбранный файл
            savePlace = folderBrowserDialog1.SelectedPath;

            jsonBody.FileNamePath = savePlace;
            jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(jsonBody);
            File.WriteAllText(path, jsonString);
        }
    }
}
