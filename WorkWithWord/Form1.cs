using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace WorkWithWord
{
    public partial class Form1 : Form
    {
        string savePath;
        object fileName, newFileName;
        string[][] keys;
        
        public Form1(string path)
        {
            InitializeComponent();
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

            fileName = Directory.GetCurrentDirectory() + @"\templates\template_01.docx";
            newFileName = savePath + @"\Форма № 01-ФР.docx";
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
            FillData();
            Utilitty.DocumentLoad(savePath, ref fileName, ref newFileName, ref keys);
        }

        private void dataBox5_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            Utilitty.CheckedOnlyOne(ref sender, ref e);
        }

        private void FillData() 
        {
            List<Control> controls = new List<Control>();
            foreach (Control control in panel1.Controls)
            {
                controls.Add(control);
            }
            controls = controls.OrderBy(x => x.TabIndex).ToList();
            Utilitty.CollectAllData(ref controls);
        }
    }
}
