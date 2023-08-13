using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        List<string> list;
        string[][] keys;
        string savePath;

        object missing = Type.Missing;
        object replace = 2;

        public Form1(string path)
        {
            InitializeComponent();
            list = new List<string>();
            keys = new string[22][];
            savePath = path;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CreateArrayKeys();

            dataBox4.CustomFormat = "dd.MM.yyyy";
            dataBox4.Format = DateTimePickerFormat.Custom;

            dataBox12.CustomFormat = "dd.MM.yyyy";
            dataBox12.Format = DateTimePickerFormat.Custom;

            dataBox20.CustomFormat = "dd.MM.yyyy";
            dataBox20.Format = DateTimePickerFormat.Custom;

            CultureInfo myClt = new CultureInfo("eu-EU", false);
            dataBox7.Culture = myClt;
        }

        private void DocumentLoad() 
        {
            //Создаём новый Word.Application
            Word.Application app = new Microsoft.Office.Interop.Word.Application();
            
            //Загружаем документ
            Word.Document doc = null;

            object fileName = Directory.GetCurrentDirectory() + @"\templates\template_01.docx";
            object newFileName = savePath + @"\Форма № 01-ФР.docx";
            object falseValue = false;
            object trueValue = true;
            
            doc = app.Documents.Open(ref fileName);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();
            CollectAllData();
            SetupData(ref app);
            
            app.ActiveDocument.SaveAs(ref newFileName, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                  ref missing, ref missing, ref missing, ref missing, ref missing,
                                                  ref missing, ref missing, ref missing, ref missing, ref missing);
            app.Documents.Close();
            app.Documents.Open(ref newFileName);
        }

        private void CollectAllData() 
        {
            List<Control> controls = new List<Control>();
            foreach (Control control in panel1.Controls) 
            {
                controls.Add(control);     
            }

            controls = controls.OrderBy(x => x.TabIndex).ToList();

            list.Clear();
            foreach (Control control in controls)
            {
                if (control is System.Windows.Forms.CheckBox)
                    list.Add(((System.Windows.Forms.CheckBox)control).Checked.ToString());
                else
                    list.Add(control.Text.Replace(".", string.Empty).Replace(",", "."));
            }
        }

        private void SetupData(ref Word.Application app) 
        {
            //Задаём параметры замены и выполняем замену.
            object findText, replaceWith; 

            for (int i = 0; i < keys.Length; ++i)
            {
                findText = "";
                replaceWith = "";
                
                for (int j = 0; j < keys[i].Length; ++j)
                {
                    if (keys[i].Length == 1 && !(i == 4 || i == 13 || i == 15))
                    {
                        replaceWith = list[i].ToString();
                    }
                    else if (i == 4 || i == 13 || i == 15)
                    {
                        replaceWith = "";
                        if (list[i].Equals("М") && j == 0) replaceWith = "V";
                        if (list[i].Equals("Ж") && j == 1) replaceWith = "V";

                        if (list[i].Equals("True") && i == 13) replaceWith = "V";
                        if (list[i].Equals("True") && i == 15) replaceWith = "V";
                    }
                    else if (list[i].Length <= j)
                    {
                        replaceWith = "";
                    }
                    else
                    {
                        replaceWith = list[i][j].ToString();
                    }

                    findText = keys[i][j].ToString();

                    app.Selection.Find.Execute(ref findText, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith,
                    ref replace, ref missing, ref missing, ref missing, ref missing);
                }
            }
        }

        private void CreateArrayKeys() 
        {
            keys[0]     = new string[5]  { "N0", "N1", "N2", "N3", "N4" };                                                                                                          // номер направления
            keys[1]     = new string[20] { "O0", "O1", "O2", "O3", "O4", "O5", "O6", "O7", "O8", "O9", "OA", "OB", "OC", "OD", "OE", "OF", "OG", "OH", "OK", "OL" };      // полис
            keys[2]     = new string[1]  { "FIO" };                                                                                                                                 // ФИО
            keys[3]     = new string[8]  { "r0", "r1", "r2", "r3", "r4", "r5", "r6", "r7" };                                                                                        // ДР
            keys[4]     = new string[2]  { "u0", "u1" };                                                                                                                            // Пол ?
            keys[5]     = new string[1]  { "ADRESS" };                                                                                                                              // Адрес МЖ
            keys[6]     = new string[1]  { "PRF" };                                                                                                                                // Место работы
            keys[7]     = new string[5]  { "C0", "C1", "C2", "C3", "C4" };                                                                                                          // Код МКБ-10
            keys[8]     = new string[1]  { "DU" };                                                                                                                                 // Документ
            keys[9]     = new string[4]  { "S0", "S1", "S2", "S3" };                                                                                                                // Серия
            keys[10]    = new string[6]  { "D0", "D1", "D2", "D3", "D4", "D5" };                                                                                                    // Номер
            keys[11]    = new string[1]  { "WHGV" };                                                                                                                             // Кем выдан
            keys[12]    = new string[8]  { "p0", "p1", "p2", "p3", "p4", "p5", "p6", "p7" };                                                                                        // Дата выдачи
            keys[13]    = new string[1]  { "sp" };                                                                                                                                  // Соц. помощь ?
            keys[14]    = new string[19] { "x0", "x1", "x2", "x3", "x4", "x5", "x6", "x7", "x8", "x9", "xA", "xB", "xC", "xD", "xE", "xF", "xG", "xH", "xI"};                                   // СНИЛС
            keys[15]    = new string[1]  { "L" };                                                                                                                                   // Льгота ?
            keys[16]    = new string[1]  { "NA" };                                                                                                                                // Обосновать направление
            keys[17]    = new string[1]  { "DCTR" };                                                                                                                              // Врач выдавший
            keys[18]    = new string[3]  { "k0", "k1", "k2" };                                                                                                                      // Код врача
            keys[19]    = new string[1]  { "ZAM" };                                                                                                                             // Заведующий
            keys[20]    = new string[1]  { "PRDSD" };                                                                                                                             // Председатель
            keys[21]    = new string[8]  { "w0", "w1", "w2", "w3", "w4", "w5", "w6", "w7" };                                                                                        // Дата
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DocumentLoad();
        }

        private void dataBox5_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            var list = sender as CheckedListBox;
            if (e.NewValue == CheckState.Checked)
                foreach (int index in list.CheckedIndices)
                    if (index != e.Index)
                        list.SetItemChecked(index, false);
        }
    }
}
