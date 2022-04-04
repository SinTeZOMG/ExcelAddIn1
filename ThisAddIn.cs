using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.IO;
using System.Drawing;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
        private Excel.Application xlApp;
        private Excel.Workbook xlWorkBook;
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            var ribbon = new Ribbon1();
            ribbon.Button1Clicked += ribbon_Button1Clicked;
            ribbon.Button2Clicked += ribbon_Button2Clicked;
            ribbon.Button3Clicked += ribbon_Button3Clicked;
            ribbon.Button4Clicked += ribbon_Button4Clicked;
            ribbon.Button5Clicked += ribbon_Button5Clicked;
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new IRibbonExtension[] { ribbon });
        }
        int Flag;
        string StrFlag;
        private void ribbon_Button1Clicked()
        {
            // ListOfSheets(Application.ActiveWorkbook);
            //FindPath(Application.ActiveWorkbook.Worksheets["4027888"]);
            //OpenFile(@"C:\Users\kondakov_a_yu\Desktop\Вакунайский ЛУ 27 после экстракции шлюм\Вода 100","4027888");
            Dialog_Form Dialog = new Dialog_Form();
            if (Dialog.ShowDialog() == DialogResult.Yes)
            {
                Flag = 1;
                StrFlag = "[t1srf]";
                CreateAll(ListOfSheets(Application.ActiveWorkbook));

            }
            else if (Dialog.DialogResult == DialogResult.No)
            {
                Flag = 2;
                StrFlag = "[relax]";
                CreateAll(ListOfSheets(Application.ActiveWorkbook));
            }


            //  ChangeSource(OpenFile(@"C:\Users\kondakov_a_yu\Desktop\Вакунайский ЛУ 27 после экстракции шлюм\Вода 100", "4027888"), Application.ActiveWorkbook.Worksheets["4027888"]);           
            // Excel.Range range = getrange("4027888", 5); // Проверка метода getrange
            // range.Copy();

        }
        private void ribbon_Button2Clicked()
        {
            // SaveInXSL(@"C:\Users\kondakov_a_yu\Desktop\RelaxNMRv1.3\archive\Карты для теста\T2T1-SR_2020-07-23_05_23_43_Салымское 5 при насыщении деканом_30257-20.map");
            FolderBrowserDialog FBD = new FolderBrowserDialog();
            if (FBD.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show(FBD.SelectedPath);
                SaveInXslAll(FileList(FBD.SelectedPath), FBD.SelectedPath);

            }
        }
        private void ribbon_Button3Clicked()
        {
            FolderBrowserDialog FBD = new FolderBrowserDialog();
            if (FBD.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show(FBD.SelectedPath);
                InsertMap(ListOfSheets(Application.ActiveWorkbook), DirList(FBD.SelectedPath), FBD.SelectedPath);

            }

            //InsPic("38922-20", @"C:\Users\kondakov_a_yu\Desktop\Уренгойское 21\ALL\38922-20_Map3.tiff");


            //FolderBrowserDialog FBD = new FolderBrowserDialog();
            //if (FBD.ShowDialog() == DialogResult.OK)
            //{
            //    MessageBox.Show(FBD.SelectedPath);
            //    DirList(FBD.SelectedPath);

            //}
        }
        private void ribbon_Button4Clicked()
        {           
            FolderBrowserDialog FBD = new FolderBrowserDialog();
            if (FBD.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show(FBD.SelectedPath);
                InsertCoutoff(ListOfSheets(Application.ActiveWorkbook), FBD.SelectedPath);

            }
        }
        private void ribbon_Button5Clicked()
        {
            //MessageBox.Show("Привет я работаю");
            InsertCoutoff(ListOfSheets(Application.ActiveWorkbook));
        }
        private List<string> ListOfSheets(dynamic wBook)
        {
            int count = wBook.Worksheets.Count;
            //wBook.ActiveSheet.Cells[1, 2] = count;
            List<string> Sheets = new List<string> { };
            for (int i = 1; i <= count; i++)
            {
                Sheets.Add(wBook.Worksheets(i).Name);
            }
            // Проверка вывода списка
            //int j = 1;
            //foreach (string i in Sheets)
            //{                
            //    wBook.ActiveSheet.Cells[j, 1] = i;
            //    j++;
            //}          
            return Sheets;

        }
        private List<string> FindPath(dynamic wSheet)
        {
            string Path;
            Path = Application.ActiveWorkbook.Path;
            //wSheet.Cells[1, 3] = Path;
            //Path = Path + @"\" + wSheet.Cells[2,1].Value;
            //wSheet.Cells[1, 4] = Path;
            List<string> FullPath = new List<string> { };

            string x;
            for (int i = 1; i < 100; i++)
            {
                x = wSheet.Cells[2, i].Value;
                if (x != null)
                {
                    if (FullPath.Contains(Path + @"\" + x) == false)
                    {
                        FullPath.Add(Path + @"\" + x);
                    }

                }
            }

            //int j = 1;
            //foreach (string i in FullPath)
            //{
            //    wSheet.cells[1, 3 + j] = i;
            //    j++;
            //}
            return FullPath;

        }

        private string CreateFile(string path, string sheet, int count)//открывает и создаёт копию (возвращает путь к копии)
        {
            string FolderPath, FilePath, SourcePath;
            FolderPath = path + @"\Исход\Т" + Flag.ToString();
            FilePath = "";
            SourcePath = "";
            DirectoryInfo di = new DirectoryInfo(@FolderPath);
            foreach (var fi in di.GetFiles("*" + sheet + "*"))
            {

                FilePath = fi.Name;
            }
            SourcePath = FolderPath + @"\" + FilePath;
            FolderPath = FolderPath + @"(Изменённые)";
            DirectoryInfo di2 = new DirectoryInfo(@FolderPath);
            if (di2.Exists == false)
            {
                di2.Create();
            }
            if (!File.Exists(FolderPath + @"\" + FilePath))
            {
                ChangeSource(SourcePath, FolderPath + @"\" + FilePath, sheet, count);

            }
            else
            {
                File.Delete(FolderPath + @"\" + FilePath);
                ChangeSource(SourcePath, FolderPath + @"\" + FilePath, sheet, count);
            }

            return FolderPath + @"\" + FilePath;
        }

        private void CreateAll(List<string> Sheets)
        {
            List<string> FullPath = new List<string> { };
            foreach (string i in Sheets)
            {
                FullPath = FindPath(Application.ActiveWorkbook.Worksheets[i]);
                int count;
                count = 1;
                foreach (string j in FullPath)
                {
                    CreateFile(j, i, count);
                    count = count + 2;
                }
                FullPath.Clear();
            }
        }
        public void ChangeSource(string path, string path1, string sheetName, int count)
        {


            using (StreamReader sr = File.OpenText(path))
            {
                using (StreamWriter wr = File.CreateText(path1))
                {
                    var s = "";
                    while ((s = sr.ReadLine()) != StrFlag)
                    {
                        wr.WriteLine(s);
                    }
                    wr.WriteLine(StrFlag + "\n");
                    Excel.Range range = getrange(sheetName, count);
                    range.Copy();
                    wr.WriteLine(Clipboard.GetText());
                    Clipboard.Clear();
                }


            }
        }
        public Excel.Range getrange(string sheetName, int count)
        {
            Excel.Worksheet sheet1 = (Excel.Worksheet)Application.ActiveWorkbook.Worksheets[sheetName];
            List<string> FullPath = new List<string> { };
            FullPath = FindPath(sheet1);
            int j;
            j = 3;
            foreach (string i in FullPath)
            {
                while (sheet1.Cells[j, count].Value != null)
                {
                    j++;
                }
            }
            Excel.Range value_range = (Excel.Range)sheet1.get_Range((Excel.Range)sheet1.Cells[3, count], (Excel.Range)sheet1.Cells[j - 1, count + 1]);
            return value_range;
        }


        private List<string> FileList(string FolderPath) //создаёт список всех файлов в директории
        {


            List<string> MList = new List<string> { };
            DirectoryInfo di = new DirectoryInfo(@FolderPath);
            foreach (var fi in di.GetFiles())
            {
                MList.Add(fi.Name);

                // Проверка вывода списка

                //int j = 1;
                //foreach (string i in MList)
                //{
                //    wBook.ActiveSheet.Cells[j, 1] = i;
                //    j++;
                //}
            }
            return MList;
        }
        private List<string> DirList(string FolderPath) //создаёт список папок в директории
        {


            List<string> DList = new List<string> { };
            DirectoryInfo di = new DirectoryInfo(@FolderPath);
            foreach (var fi in di.GetDirectories())
            {
                DList.Add(fi.Name);

                //    // Проверка вывода списка

                //    int j = 1;
                //    foreach (string i in DList)
                //    {
                //        wBook.ActiveSheet.Cells[j, 1] = i;
                //        j++;
                //    }
            }
            return DList;
        }
        public void SaveInXSL(string Filename, string FolderPath)
        {      //сохраняет карту в xslx формате

            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            app.DisplayAlerts = false;

            //xlApp.Visible = true; //чтобы увидеть в том ли виде открылся документ                   
            object misValue = System.Reflection.Missing.Value;
            string Buffer;
            FileStream file1 = new FileStream(FolderPath + "\\" + Filename, FileMode.Open); //создаем файловый поток
            StreamReader reader = new StreamReader(file1); // создаем «потоковый читатель» и связываем его с файловым потоком

            Buffer = reader.ReadToEnd();
            reader.Close(); //закрываем поток
            File.WriteAllText(FolderPath + "\\" + Filename, Buffer, new UTF8Encoding(false));       // Сохраняем файл UTF-8

            app.ScreenUpdating = false; // фоновый режим работы макроса

            Microsoft.Office.Interop.Excel._Workbook excelWorkbook = app.Workbooks.Open(FolderPath + "\\" + Filename, Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited, Format: 1); //Открываем файл, читаем его с разделителями Tab
            string sheetname = Filename.Substring(Filename.Length - 12);
            excelWorkbook.Sheets[1].Name = sheetname.Substring(0, sheetname.Length - 4); //удаляем дату и время замера из названия листа

            DirectoryInfo di2 = new DirectoryInfo(@FolderPath + "\\" + "ЭксельКарты");
            if (di2.Exists == false)
            {
                di2.Create(); //Проверяем существует ли папка, если нет создаём
            }
            excelWorkbook.SaveAs(FolderPath + "\\" + "ЭксельКарты" + "\\" + Filename.Substring(28) + ".xlsx", Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            excelWorkbook.Close(false); // Сохраняем файл xlsx в новой папке 

            file1.Close();

            FileStream file2 = new FileStream(FolderPath + "\\" + Filename, FileMode.Open); //создаем файловый поток
            StreamReader reader2 = new StreamReader(file2); // создаем «потоковый читатель» и связываем его с файловым потоком

            Buffer = reader2.ReadToEnd(); // Записываем поток в буфер

            reader2.Close(); //закрываем поток
            reader2.Dispose();
            File.WriteAllText(FolderPath + "\\" + Filename, Buffer, new UTF8Encoding(true)); // Сохраняем файл UTF-8 c BOM
            file2.Close();
            app.ScreenUpdating = true; /// убераем фоновый режим
        }

        public void SaveInXslAll(List<string> MList, string FolderPath)
        {

            foreach (string i in MList)
            {
                SaveInXSL(i, FolderPath);
            }
            MessageBox.Show("Готово", "Поверх всех окон", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification); ;
        }

        public void InsPic(string sheetName, string PicPath, int flag) // вставляет картинку в лист на определённое место
        {
            Excel.Worksheet sheet1 = (Excel.Worksheet)Application.ActiveWorkbook.Worksheets[sheetName];
            Excel.Range xrange = (Excel.Range)sheet1.get_Range((Excel.Range)sheet1.Cells[flag, 1], (Excel.Range)sheet1.Cells[flag + 22, 11]);
            sheet1.Shapes.AddPicture(PicPath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, sheet1.Cells[flag, 1].Left, sheet1.Cells[flag, 1].top, xrange.Width, xrange.Height);

        }

        public void InsertMap(List<string> Sheets, List<string> DList, string FolderPath)  // вставляет все карты из всех папок в нужные листы друг за другом.
        {
            List<string> MList = new List<string> { };
            foreach (string i in Sheets)
            {

                int flag = 53;
                MList.Clear();

                foreach (string n in DList)
                {

                    MList = FileList(FolderPath + @"\" + n);
                    foreach (string j in MList)
                        if (i == j.Substring(0, j.Length - 9))
                        {
                            InsPic(i, FolderPath + @"\" + n + @"\" + j, flag);
                            flag = flag + 22;
                        }
                }




                //int j = 0;
                //foreach (string i in Sheets)
                //{   
                //    if (i == MList[j].Substring(0, MList[j].Length-9))
                //    {
                //        InsPic(i, FolderPath + @"\" + MList[j]);
                //        j++;
                //    }


            }
        }
        private string GetFileName(string path, string sheet)//открывает и создаёт копию (возвращает путь к копии)
        {
            string FilePath = "";
            DirectoryInfo di = new DirectoryInfo(@path);
            foreach (var fi in di.GetFiles("*" + sheet + "*"))
            {

                FilePath = fi.Name;
            }


            return path + @"\" + FilePath;


        }
        public void Coutoff(string sheetName, string FilePath) // вставляет картинку в лист на определённое место
        {
            Excel.Worksheet sheet1 = (Excel.Worksheet)Application.ActiveWorkbook.Worksheets[sheetName];
            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            app.DisplayAlerts = false;
            object misValue = System.Reflection.Missing.Value;
            app.ScreenUpdating = false; // фоновый режим работы макроса
            Microsoft.Office.Interop.Excel._Workbook excelWorkbook = app.Workbooks.Open(FilePath); //Открываем файл, читаем его с разделителями Tab          
            



            sheet1.Cells[7, 7] = excelWorkbook.Sheets[4].Cells[3, 3];
            sheet1.Cells[7, 8] = excelWorkbook.Sheets[4].Cells[4, 3];
            sheet1.Cells[7, 9] = excelWorkbook.Sheets[4].Cells[5, 3];
            sheet1.Cells[7, 11] = excelWorkbook.Sheets[5].Cells[3, 2];
            sheet1.Cells[7, 12] = excelWorkbook.Sheets[5].Cells[4, 2];

            //sheet1.Range["T10:T110"].Value = excelWorkbook.Sheets[3].Range["C13:C112"].Value;



            Excel.Chart Chart1 = (Excel.Chart)sheet1.ChartObjects(1).Chart;
            Excel.SeriesCollection SerColl = (Excel.SeriesCollection)Chart1.SeriesCollection();
            sheet1.Cells[7, 20].FormulaR1C1 = "=MAX(R[3]C[-4]:R[102]C[-4])";

            Excel.Series series1 = (Excel.Series)SerColl.NewSeries();        
            series1.Name = "Отсечка 1";
            series1.XValues = "='"+sheetName+"'!$O$5:$O$6";
            series1.Values = "='" + sheetName + "'!$P$5:$P$6";
            series1.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Red);
            series1.MarkerBackgroundColor = ColorTranslator.ToOle(Color.Red);
            series1.MarkerForegroundColor = ColorTranslator.ToOle(Color.Red);
            series1.Format.Line.DashStyle = Office.MsoLineDashStyle.msoLineDashDot;

           
            Excel.Series series2 = (Excel.Series)SerColl.NewSeries();
            series2.Name = "Отсечка 2";
            series2.XValues = "='" + sheetName + "'!$Q$5:$Q$6";
            series2.Values = "='" + sheetName + "'!$R$5:$R$6";

            series2.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Red);
            series2.MarkerBackgroundColor = ColorTranslator.ToOle(Color.Red);
            series2.MarkerForegroundColor = ColorTranslator.ToOle(Color.Red);
            series2.Format.Line.DashStyle = Office.MsoLineDashStyle.msoLineDashDot;

            double a; a = sheet1.Cells[7, 20].value;            
            Chart1.Axes(Excel.XlAxisType.xlValue).MaximumScale = a;
            //Chart1.FullSeriesCollection(3).ForeColor.RGB = Excel.XlRgbColor.rgbRed;


 



            excelWorkbook.Close(false); // 
            app.ScreenUpdating = true; /// убераем фоновый режим
        }
        public void Coutoff(string sheetName) // вставляет картинку в лист на определённое место
        {
            
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Worksheet sheet1 = (Excel.Worksheet)Application.ActiveWorkbook.Worksheets[sheetName];
            app.DisplayAlerts = false;
            object misValue = System.Reflection.Missing.Value;
      
            app.ScreenUpdating = false; // фоновый режим работы макроса

            int count = 1;

            for (int i = 1; i < 34; i++)
            {

                if (sheet1.Cells[9, i].Value == "После донасыщения")
                {
                    count = i;
                }

            }
            count = count - 19;
            Excel.ChartObjects xlCharts = (Excel.ChartObjects)sheet1.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Item(1);
            Excel.Chart Chart1 = (Excel.Chart)myChart.Chart;
            Excel.SeriesCollection SerColl = (Excel.SeriesCollection)Chart1.SeriesCollection(Type.Missing);
            sheet1.Cells[7, 20].FormulaR1C1 = "=MAX(R[3]C[" + count.ToString() + "]:R[102]C[" + count.ToString() + "])";       

            Excel.Series series1 = (Excel.Series)SerColl.NewSeries();
            series1.Name = "Отсечка 1";
            series1.XValues = "='" + sheetName + "'!$O$5:$O$6";
            series1.Values = "='" + sheetName + "'!$P$5:$P$6";


            series1.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Red);
            series1.MarkerBackgroundColor = ColorTranslator.ToOle(Color.Red);
            series1.MarkerForegroundColor = ColorTranslator.ToOle(Color.Red);
            series1.Format.Line.DashStyle = Office.MsoLineDashStyle.msoLineDashDot;
            


            Excel.Series series2 = (Excel.Series)SerColl.NewSeries();
            series2.Name = "Отсечка 2";
            series2.XValues = "='" + sheetName + "'!$Q$5:$Q$6";
            series2.Values = "='" + sheetName + "'!$R$5:$R$6";

            series2.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Red);
            series2.MarkerBackgroundColor = ColorTranslator.ToOle(Color.Red);
            series2.MarkerForegroundColor = ColorTranslator.ToOle(Color.Red);
            series2.Format.Line.DashStyle = Office.MsoLineDashStyle.msoLineDashDot;


            double a; a = sheet1.Cells[7, 20].value;
         
            Chart1.Axes(Excel.XlAxisType.xlValue).MaximumScale = a;
            //Chart1.FullSeriesCollection(3).Format.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Blue); 






            app.ScreenUpdating = true; /// убераем фоновый режим
        }
        public void InsertCoutoff(List<string> Sheets, string FolderPath)  // вставляет все отсечки из всех папок в нужные листы друг за другом.
        {
            List<string> MList = new List<string> { };
            foreach (string i in Sheets)
            {
                string FileName;
                FileName = GetFileName(FolderPath, i);
                Coutoff(i, FileName);

            };
        }
        public void InsertCoutoff(List<string> Sheets)  // вставляет все отсечки из всех папок в нужные листы друг за другом.
        {
            List<string> MList = new List<string> { };
            foreach (string i in Sheets)
            {
                Coutoff(i);
            };





        }
    }
}
