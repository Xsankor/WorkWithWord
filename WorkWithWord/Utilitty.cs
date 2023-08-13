using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WorkWithWord
{
    public static class Utilitty
    {
        private static object missing = Type.Missing;
        private static object replace = 2;
        private static List<string> listData = new List<string>();

        public static void DocumentLoad(string savePath, ref object fileName, ref object newFileName, 
                                        ref string[][] keys)
        {
            //Создаём новый Word.Application
            Word.Application app = new Microsoft.Office.Interop.Word.Application();

            //Загружаем документ
            Word.Document doc = null;

            object falseValue = false;
            object trueValue = true;

            doc = app.Documents.Open(ref fileName);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();
            SetupData(ref app, ref keys);

            app.ActiveDocument.SaveAs(ref newFileName, ref missing, ref missing, ref missing, ref missing, ref missing,
                                      ref missing, ref missing, ref missing, ref missing, ref missing,
                                      ref missing, ref missing, ref missing, ref missing, ref missing);
            app.Documents.Close();
            app.Documents.Open(ref newFileName);
        }

        public static void CollectAllData(ref List<Control> controls)
        { 
            listData.Clear();
            foreach (Control control in controls)
            {
                if (control is System.Windows.Forms.CheckBox)
                    listData.Add(((System.Windows.Forms.CheckBox)control).Checked.ToString());
                else
                    listData.Add(control.Text.Replace(".", string.Empty).Replace(",", "."));
            }
        }

        public static void SetupData(ref Word.Application app, ref string[][] keys)
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
                        replaceWith = listData[i].ToString();
                    }
                    else if (i == 4 || i == 13 || i == 15)
                    {
                        replaceWith = "";
                        if (listData[i].Equals("М") && j == 0) replaceWith = "V";
                        if (listData[i].Equals("Ж") && j == 1) replaceWith = "V";

                        if (listData[i].Equals("True") && i == 13) replaceWith = "V";
                        if (listData[i].Equals("True") && i == 15) replaceWith = "V";
                    }
                    else if (listData[i].Length <= j)
                    {
                        replaceWith = "";
                    }
                    else
                    {
                        replaceWith = listData[i][j].ToString();
                    }

                    findText = keys[i][j].ToString();

                    app.Selection.Find.Execute(ref findText, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith,
                    ref replace, ref missing, ref missing, ref missing, ref missing);
                }
            }
        }

        public static void CheckedOnlyOne(ref object sender, ref ItemCheckEventArgs e)
        {
            var list = sender as CheckedListBox;
            if (e.NewValue == CheckState.Checked)
                foreach (int index in list.CheckedIndices)
                    if (index != e.Index)
                        list.SetItemChecked(index, false);
        }
    }
}
