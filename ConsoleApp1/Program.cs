using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;

namespace ConsoleApp1
{
    internal class Program
    {

        static void Main(string[] args)
        {
            string baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string inputPath = baseDir; // + "\\in";

            if (args.Length > 0)
            {
                inputPath += $"\\{args[0]}";
                Console.WriteLine($"Directorio donde están los ficheros CSV: {inputPath}");
            }

            try
            {
                ProcessDirectory(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
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
            var last = filePath.Split('\\').Last();
            string baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string outPath = baseDir + "\\out";

            string line = string.Empty;
            List<string> lines = new List<string>();

            StreamReader file = new StreamReader(filePath);
            while ((line = file.ReadLine()) != null)
            {
                lines.Add(HttpUtility.HtmlDecode(line));
            }

            string fileOut = $"{outPath}\\{last}";
            Directory.CreateDirectory(outPath);
            lines.ForEach(x => File.AppendAllText(fileOut, x + '\n'));

            file.Close();
        }
    }
}
