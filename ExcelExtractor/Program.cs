using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Threading;
using ClosedXML.Excel;

namespace ExcelExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            // ExcelExtractor [source] [target] [cells]
            //args = new string[3];
            //args[0] = @"C:\Git\ExcelExtractor\ExampleSources";
            //args[1] = @"C:\Git\ExcelExtractor\target.xlsx";
            //args[2] = "A1,A2,A3";

            EEmodel model = new EEmodel();
            
            ExitOnNoParameterGiven(args);
            model.source = args[0];
            model.target = args[1];
            model.cells = args[2];
            //Console.WriteLine(model.source);

            model  = readInCellDescriptions(model);

            List<string> endings = new List<string>{".xls", ".xlsx", ".xlsm"};
            model.sourcefiles = SearchFilesWithExtensions(model.source, endings);
            
            int i = 0;
            int countFiles = model.sourcefiles.Count;
            //Console.WriteLine(countFiles.ToString());
            string filename = "";
            foreach (var path in model.sourcefiles)
            {
                i++;
                ShowProgressBar(countFiles, i);
                filename = Path.GetFileName(path);
                model.cellValues.Add(filename,  ReadFieldsFromExcel(path, model.cellNames));
                //Console.WriteLine(path);
            }

            WriteDictionaryToExcel( model.target, model.cellValues, model.cellNames);
            Console.WriteLine("");  
            Console.WriteLine("Target file at: " +  model.target);
            waitDialogue();
            
        }
        private static void waitDialogue()
        {
            Console.WriteLine("Press andy key to proceed.");
            Console.ReadKey(true); 
            Environment.Exit(0);
        }
        private static void ExitOnNoParameterGiven(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("No parameter given! ExcelExtractor");
                Console.WriteLine("Format: [source] [target] [cells] ");
                Console.WriteLine("Example: \"C:\\sourcefolder\\\" \"c:\\target.xlsx\" \"A1,b2,G53\" ");
                waitDialogue();
            }
        }

        public static List<string> SearchFilesWithExtensions(string folderPath, List<string> extensions)
        {
            try
            {
                // SelectMany is used to create a flat list out of the deep list
                // The search is recursive due to the option AllDirectories
                Console.WriteLine(folderPath);
                var files = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories)
                    .Where(file => extensions.Contains(Path.GetExtension(file)))
                    .ToList();

                return files;
            }
            catch (Exception ex)
            {
                // Exception handling
                Console.WriteLine($"A failure was foubd: {ex.Message}");
                return new List<string>();
            }
        }
        private static EEmodel readInCellDescriptions(EEmodel model)
        {
            EEmodel result = model;
            result.cellNames = SplitStringIntoList(model.cells);
            return result;
        }
        public static List<string> SplitStringIntoList(string input)
        {
            return input.Split(',')
                        .Select(s => s.Replace(" ", "").ToLower())
                        .ToList();
        }
        public static void ShowProgressBar(int max, int current)
        {
    
            // Prozentualer Fortschritt berechnen
            double percent = (double)current / max * 100;

            // Fortschrittsbalken zeichnen
            Console.Write("\r[{0}{1}] {2:0.00}%", new string('#', current), new string(' ', max - current), percent);

        }
        public static List<string> ReadFieldsFromExcel(string filePath, List<string> cellReferences)
        {
            var results = new List<string>();
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1); // Nehmen Sie das erste Arbeitsblatt
                
                foreach (var cellReference in cellReferences)
                {
                    var cell = worksheet.Cell(cellReference);

                    if (cell != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                    {
                        results.Add(cell.Value.ToString());
                    }
                }
            }

            return results;
        }
        public static void WriteDictionaryToExcel(string filePath, Dictionary<string, List<string>> data, List<string> headers)
        {
            // Überprüfen, ob die Datei bereits existiert
            if (File.Exists(filePath))
            {
                string dir = Path.GetDirectoryName(filePath);
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string fileExt = Path.GetExtension(filePath);
                filePath = Path.Combine(dir, $"{fileName}(1){fileExt}");
            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");

                // Schreibe die Spaltenüberschriften
                worksheet.Cell(1, 1).Value = "Key";
                for (int i = 0; i < headers.Count; i++)
                {
                    worksheet.Cell(1, i + 2).Value = headers[i];
                }

                // Schreibe die Daten
                int row = 2;
                foreach (var kvp in data)
                {
                    worksheet.Cell(row, 1).Value = kvp.Key;
                    for (int i = 0; i < kvp.Value.Count; i++)
                    {
                        worksheet.Cell(row, i + 2).Value = kvp.Value[i];
                    }
                    row++;
                }

                workbook.SaveAs(filePath);
            }
        }
    }
}
