using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace Conversor_CSV
{
    public partial class Form1 : Form
    {
       
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Archivo CSV|*.CSV";
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                try
                {
                    string[] files = openFileDialog1.FileNames;

                    for (int i = 0; i < files.Length; i++)
                    {
                        string filePath = files[i];
                        ProcessFile(filePath);

                        //string baseDir = Path.GetDirectoryName(filePath);
                        //string outPath = baseDir + "\\out";
                        //string fileName = files[i].Split('\\').Last();

                        //string text = File.ReadAllText(filePath);
                        //text = HttpUtility.HtmlDecode(text);

                        //string line = string.Empty;
                        //List<string> lines = new List<string>();

                        //StreamReader file = new StreamReader(filePath);
                        //while ((line = file.ReadLine()) != null)
                        //{
                        //    lines.Add(HttpUtility.HtmlDecode(line));
                        //}

                        //string fileOut = $"{outPath}\\{fileName}";
                        //Directory.CreateDirectory(outPath);
                        ////File.AppendAllText(fileOut, text + '\n');
                        //lines.ForEach(x => File.AppendAllText(fileOut, x + '\n'));


                    }

                    result = MessageBox.Show("Proceso terminado");

                    
                }
                catch (IOException)
                {
                }
            }


            
        }

        public static void ProcessDirectory(string targetDirectory)
        {
            // Procesa la lista de ficheros encontrados en el directorio
            string extension = ".csv";
            string[] fileEntries = Directory.GetFiles(targetDirectory).Where(x => x.EndsWith(extension)).ToArray();

            foreach (string fileName in fileEntries)
                ProcessFile(fileName);

            // Busqueda recursiva en subdirectorios del directorio.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
                ProcessDirectory(subdirectory);
        }

        public static void ProcessFile(string filePath)
        {
            string fileName = filePath.Split('\\').Last();
            string baseDir = Path.GetDirectoryName(filePath);
            string outPath = baseDir + "\\out";

            string line = string.Empty;
            List<string> lines = new List<string>();

            StreamReader file = new StreamReader(filePath);
            while ((line = file.ReadLine()) != null)
            {
                lines.Add(HttpUtility.HtmlDecode(line));
            }

            string fileOut = $"{outPath}\\{fileName}";
            Directory.CreateDirectory(outPath);
            lines.ForEach(x => File.AppendAllText(fileOut, x + '\n'));

            file.Close();

            DisplayInExcel(lines);

            //ConvertToXlsx(filePath, fileOut);
        }

        static void DisplayInExcel(IEnumerable<string> lines)
        {
            var excelApp = new Excel.Application();
            // Make the object visible.
            excelApp.Visible = true;

            // Create a new, empty workbook and add it to the collection returned
            // by property Workbooks. The new workbook becomes the active workbook.
            // Add has an optional parameter for specifying a particular template.
            // Because no argument is sent in this example, Add creates a new workbook.
            excelApp.Workbooks.Add();

            // This example uses a single workSheet. The explicit type casting is
            // removed in a later procedure.
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            // Establish column headings in cells A1 and B1.
            //workSheet.Cells[1, "A"] = "ID Number";
            //workSheet.Cells[1, "B"] = "Current Balance";

            var row = 0;
            foreach (var line in lines)
            {
                row++;
                workSheet.Cells[row, "A"] = line;
            }

            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();
        }

        static void ConvertToXlsx(string sourcefile, string destfile)
        {
            int i, j;
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel._Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            string[] lines, cells;
            lines = File.ReadAllLines(sourcefile);
            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Add();
            xlWorkSheet = (Excel._Worksheet)xlWorkBook.ActiveSheet;
            for (i = 0; i < lines.Length; i++)
            {
                cells = lines[i].Split(new Char[] { '\t', ';' });
                for (j = 0; j < cells.Length; j++)
                    xlWorkSheet.Cells[i + 1, j + 1] = cells[j];
            }
            xlWorkBook.SaveAs(destfile, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
        }
    }
}
