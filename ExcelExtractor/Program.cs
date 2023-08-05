using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Threading;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

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
            Console.WriteLine("Press andy key to proceed."); // Hinweis für den Benutzer
            Console.ReadKey(true); // Warten auf die Eingabe einer Taste
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
                // SelectMany wird verwendet, um eine flache Liste aus der verschachtelten Liste zu erstellen.
                // GetFiles gibt alle Dateien in dem angegebenen Pfad zurück.
                // Die Suche ist rekursiv, da die Option AllDirectories verwendet wird.
                Console.WriteLine(folderPath);
                var files = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories)
                    .Where(file => extensions.Contains(Path.GetExtension(file)))
                    .ToList();

                return files;
            }
            catch (Exception ex)
            {
                // Ausnahmehandhabung entsprechend Ihrer Anforderungen
                Console.WriteLine($"Ein Fehler ist aufgetreten: {ex.Message}");
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

            IWorkbook workbook = new XSSFWorkbook();

            ISheet sheet = workbook.CreateSheet("Sheet1");

            // Schreibe die Spaltenüberschriften
            IRow headerRow = sheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("Key");
            for (int i = 0; i < headers.Count; i++)
            {
                headerRow.CreateCell(i + 1).SetCellValue(headers[i]);
            }

            // Schreibe die Daten
            int row = 1;
            foreach (var kvp in data)
            {
                IRow dataRow = sheet.CreateRow(row);
                dataRow.CreateCell(0).SetCellValue(kvp.Key);
                for (int i = 0; i < kvp.Value.Count; i++)
                {
                    dataRow.CreateCell(i + 1).SetCellValue(kvp.Value[i]);
                }
                row++;
            }

            using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream);
            }
        }
        public static List<string> ReadFieldsFromExcel(string filePath, List<string> cellReferences)
        {
            var results = new List<string>();

            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook;

                string extension = Path.GetExtension(filePath);

                if (extension == ".xls")
                {
                    workbook = new HSSFWorkbook(fs);
                }
                else if (extension == ".xlsx" || extension == ".xlsm")
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else
                {
                    throw new ArgumentException("Unsupported file extension: " + extension);
                }

                ISheet sheet = workbook.GetSheetAt(0); // Nehmen Sie das erste Arbeitsblatt

                foreach (var cellReference in cellReferences)
                {
                    var cell = GetCell(sheet, cellReference);

                    if (cell != null && cell.CellType == CellType.String && !string.IsNullOrEmpty(cell.StringCellValue))
                    {
                        results.Add(cell.StringCellValue);
                    }
                }
            }

            return results;
        }

        // Hilfsmethode um die Zelle basierend auf einer Referenz wie "A1" zu holen
        private static ICell GetCell(ISheet sheet, string cellReference)
        {
            var cellRef = new NPOI.SS.Util.CellReference(cellReference);
            IRow row = sheet.GetRow(cellRef.Row);
            if (row != null)
            {
                return row.GetCell(cellRef.Col);
            }
            return null;
        }
    }
}
