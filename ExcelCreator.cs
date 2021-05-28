using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace WPFSWTry
{
    class ExcelCreator
    {
        public ExcelCreator (string path, string prodNum, List<OfficialBomItem> bomItems, List<RunTime> runTimes)
        {
            this.Path = path;
            this.ProdNum = prodNum;
            this.BomItems = bomItems;
            this.RunTimes = runTimes;
        }

        public string Path { get; set; }

        public string ProdNum { get; set; }

        public FileInfo File { get; set; }

        public List<OfficialBomItem> BomItems { get; set; }

        public List<RunTime> RunTimes { get; set; }

        public FileInfo CreateExcelFile(string path, string prodNum)
        {
            var excelFile = new FileInfo(fileName: @"" + path + $"\\PROD-{prodNum} BOM.xlsx");
            return excelFile;
        }

        public async Task SaveExcelFile(List<OfficialBomItem> bomItems, FileInfo file)
        {
            DeleteIfExists(file);

            using (var package = new ExcelPackage(file))
            {
               var ws = package.Workbook.Worksheets.Add(Name: "BOM");
               var range = ws.Cells[Address: "A2:C100"];
               range.AutoFitColumns();

                //Set up the main BOM 
                /*
                ws.Cells["A1"].Value = $"PROD-{this.ProdNum} Material Counts and Run Times";
                ws.Cells["A1:F1"].Merge = true;
                ws.Row(1).Style.Font.Size = 18;
                ws.Row(1).Style.Font.Color.SetColor(Color.DarkGreen);
                ws.Cells["A1:F2"].Style.Font.Bold = true;
                */


                int counter = 2;
                //Add in the CNC run time estimate
                for (int i = 0; i < bomItems.Count; i++)
                {
                    string name = bomItems[i].Description + " " + bomItems[i].Thickness;
                    string quantity = bomItems[i].Quantity.ToString();
                    string individualTime = this.RunTimes[i].IndividualTimeOutput(this.RunTimes[i].Seconds);
                    ws.Cells[Address: $"A{counter}"].Value = name;
                    ws.Cells[Address: $"B{counter}"].Value = quantity;

                    counter++;
                }


                int total = 0;
                foreach (var time in this.RunTimes)
                {
                    total += time.Seconds;
                }
                TimeSpan t = TimeSpan.FromSeconds(total);
                string totalRunTime = t.ToString(@"hh\:mm\:ss");

                ws.Cells[Address: $"A{counter + 2}"].Value = totalRunTime;

                range.AutoFitColumns();

                await package.SaveAsync();
            }
        }

        private void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }

        public void OpenExcelFile(FileInfo file)
        {
            Process excelDoc = new Process();

            try
            {
                excelDoc.StartInfo.FileName = file.FullName;
                excelDoc.Start();
            }
            catch
            {

            }
        }
    }
}
