using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;


namespace ConsoleApp1
{
    internal class Program
    {
        static void Main(string[] args)
        {          

            var tmList = new List<TestModel>();
            TestModel tm = new TestModel();
            tm.StockLength = 6000;
            tm.Waste = 4500;
            tm.Material = "Copper - A";
            tm.Type = "Standard";
            tm.Cuts = new Dictionary<CutData, float> { [new CutData { Cut = "100", PosNumber = 1 }] = 1000 };//   "Postion Number: 815 && Cut:1899";
            tm.Diameter = 100;
            tmList.Add(tm);

            TestModel tm1 = new TestModel();
            tm1.StockLength = 3000;
            tm1.Waste = 2500;
            tm1.Material = "Copper - A";
            tm1.Type = "Standard";
            tm1.Cuts = new Dictionary<CutData, float> { [new CutData { Cut = "200", PosNumber = 2 }] = 2000, [new CutData { Cut = "150", PosNumber = 15 }] = 1500 };//   "Postion Number: 815 && Cut:1899";
            tm1.Diameter = 100;
            tmList.Add(tm1);

            TestModel tm2 = new TestModel();
            tm2.StockLength = 5000;
            tm2.Waste = 4500;
            tm2.Material = "Copper - A";
            tm2.Type = "Standard";
            tm2.Cuts = new Dictionary<CutData, float> { [new CutData { Cut = "300", PosNumber = 3 }] = 3000 };//   "Postion Number: 815 && Cut:1899";
            tm2.Diameter = 50;
            tmList.Add(tm2);

            TestModel tm3 = new TestModel();
            tm3.StockLength = 4000;
            tm3.Waste = 3500;
            tm3.Material = "Copper - A";
            tm3.Type = "Standard";
            tm3.Cuts = new Dictionary<CutData, float> { [new CutData { Cut = "400", PosNumber = 4 }] = 4000 };//   "Postion Number: 815 && Cut:1899";
            tm3.Diameter = 25;
            tmList.Add(tm3);

            var groupedModelList = tmList.GroupBy(x => x.Diameter).Select(grp => grp.ToList()).ToList();
            //exportDocument("D:\\temp\\test.xlsx", groupedModelList);
            ExportExc p = new ExportExc();
            p.CreateExcelFile(groupedModelList, "d:\\");

        }
        /*var export = new ExportExc();
               var dt = ExportExc.ToDataTable(tmList.testData);            
               export.exportDocument("D:\\temp\\test.xlsx", dt, "Diameter");*/

        /*public static void exportDocument(string FilePath, List<List<TestModel>> GroupedList)
        {
            WorkbookPart wBookPart = null;
            using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Create(FilePath, SpreadsheetDocumentType.Workbook))
            {
                wBookPart = spreadsheetDoc.AddWorkbookPart();
                wBookPart.Workbook = new Workbook();
                uint sheetId = 1;
                spreadsheetDoc.WorkbookPart.Workbook.Sheets = new Sheets();
                Sheets sheets = spreadsheetDoc.WorkbookPart.Workbook.GetFirstChild<Sheets>();


                foreach (var testModels in GroupedList)
                {
                    var sheetName = testModels.First().Diameter.ToString();
                    WorksheetPart wSheetPart = wBookPart.AddNewPart<WorksheetPart>();
                    Sheet sheet = new Sheet() { Id = spreadsheetDoc.WorkbookPart.GetIdOfPart(wSheetPart), SheetId = sheetId, Name = "ND-" + sheetName };
                    sheets.Append(sheet);

                    SheetData sheetData = new SheetData();
                    wSheetPart.Worksheet = new Worksheet(sheetData);

                    List<string> properties = new List<string>();

                    Row headerRow = new Row();

                    Cell lastHeaderCell = new Cell();
                    lastHeaderCell.DataType = CellValues.String;
                    lastHeaderCell.CellValue = new CellValue("Cuts");

                    foreach (var property in testModels.First().GetType().GetProperties())
                    {
                        if (property.Name != "Diameter" && property.Name != "Cuts")
                        {
                            Cell cell = new Cell();
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(property.Name);
                            headerRow.AppendChild(cell);
                        }
                    }
                    headerRow.AppendChild(lastHeaderCell);
                    sheetData.AppendChild(headerRow);

                    foreach (var testModel in testModels)
                    {
                        Row row = new Row();
                        List<Cell> lastCells = new List<Cell>();
                        foreach (var property in testModel.GetType().GetProperties())
                        {
                            if (property.Name == "Cuts")
                            {
                                foreach (var keyValue in (property.GetValue(testModel, null) as Dictionary<CutData, float>))
                                {
                                    Cell cell = new Cell();
                                    cell.DataType = CellValues.String;
                                    cell.CellValue = new CellValue(keyValue.Key.Cut + ";" + keyValue.Key.PosNumber + "-" + keyValue.Value);
                                    lastCells.Add(cell);
                                    //row.AppendChild(cell);
                                }
                            }
                            else if (property.Name != "Diameter")
                            {
                                Cell cell = new Cell();
                                cell.DataType = CellValues.String;
                                cell.CellValue = new CellValue(Convert.ToString(property.GetValue(testModel, null)));
                                row.AppendChild(cell);
                            }

                        }
                        foreach (var cell in lastCells)
                        {
                            row.AppendChild(cell);
                        }
                        sheetData.AppendChild(row);
                    }
                    sheetId++;
                }
            }
        }

        public class TestModel
        {
            public int StockLength { get; set; }
            public int Waste { get; set; }
            public Dictionary<CutData, float> Cuts { get; set; }
            public string Material { get; set; }
            public string Type { get; set; }
            public int Diameter { get; set; }


        }*/

       
    }
}
