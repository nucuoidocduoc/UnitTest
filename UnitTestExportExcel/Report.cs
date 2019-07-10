using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTestExportExcel
{
    public class Report
    {
        public List<Data> DataCollector { get; set; }

        public Report()
        {
            DataCollector = new List<Data>();
            for (int i = 0; i < 3000; i++) {
                DataCollector.Add(new Data());
            }
        }

        public void CreateExcelDoc(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook)) {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = (DocumentFormat.OpenXml.Spreadsheet.Workbook)new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = (DocumentFormat.OpenXml.Spreadsheet.Worksheet)new Worksheet();

                #region region auto fit column

                //double w1 = GetContentLength(parameterSort, ParamProperty.Group);
                //double w2 = GetContentLength(parameterSort, ParamProperty.Name);
                //double w3 = GetContentLength(parameterSort, ParamProperty.Value);

                //Columns columns = new Columns();
                //if (w1 != 0)
                //{
                //    columns.Append(CreateColumnData(1, 1, w1));
                //}
                //if (w2 != 0)
                //{
                //    columns.Append(CreateColumnData(2, 2, w2));
                //}
                //if (w3 != 0)
                //{
                //    columns.Append(CreateColumnData(3, 3, w3));
                //}
                //worksheetPart.Worksheet.Append(columns);

                #endregion region auto fit column

                WorkbookStylesPart wbsp = workbookpart.AddNewPart<WorkbookStylesPart>();
                // add styles to sheet
                wbsp.Stylesheet = GenerateStylesheet();
                wbsp.Stylesheet.Save();
                SheetData sd = new SheetData();
                worksheetPart.Worksheet.Append(sd);

                OpenXmlWriter writer = null;
                Cell cell = null;
                Row row = null;
                // Add Sheets to the Workbook.
                //Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                //// Append a new worksheet and associate it with the workbook.
                //Sheet sheet = new Sheet() {
                //    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                //    SheetId = 1,
                //    Name = "Test"
                //};
                //sheets.Append(sheet);

                writer = OpenXmlWriter.Create(worksheetPart);
                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());

                int rowIndex = 1;

                foreach (Data excelRowData in DataCollector) {
                    row = new Row();
                    // Add Data
                    int columnIndex = 0;
                    bool isRowData = false;

                    if (excelRowData.Values.Count != 0) {
                        foreach (string value in excelRowData.Values) {
                            cell = CreateSpreadsheetCellIfNotExist(worksheetPart.Worksheet, columnIndex, rowIndex);
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(value);
                            cell.StyleIndex = 2;
                            if (columnIndex == (excelRowData.Values.Count - 1)) {
                                cell.StyleIndex = 0;
                            }
                            row.Append(cell);
                            columnIndex++;
                        }
                        if (isRowData) {
                            isRowData = false;
                        }
                    }
                    writer.WriteElement(row);
                    rowIndex++;
                }

                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.Close();
                workbookpart.Workbook.Save();
                spreadsheetDocument.Close();

                //abc
                // Close the document.
            }
        }

        public void CreateExcel(string fileName)
        {
            using (SpreadsheetDocument xl = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook)) {
                List<OpenXmlAttribute> oxa;
                OpenXmlWriter oxw;

                xl.AddWorkbookPart();
                WorksheetPart wsp = xl.WorkbookPart.AddNewPart<WorksheetPart>();
                WorkbookStylesPart wbsp = xl.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                // add styles to sheet
                wbsp.Stylesheet = GenerateStylesheet();
                wbsp.Stylesheet.Save();
                var childs = wbsp.Stylesheet.GetAttributes();

                oxw = OpenXmlWriter.Create(wsp);
                oxw.WriteStartElement(new Worksheet());
                oxw.WriteStartElement(new SheetData());
                Row row = null;
                for (int i = 1; i <= 50; ++i) {
                    oxa = new List<OpenXmlAttribute>();
                    // this is the row index
                    oxa.Add(new OpenXmlAttribute("r", null, i.ToString()));
                    row = new Row();
                    oxw.WriteStartElement(row, oxa);

                    for (int j = 1; j <= 10; ++j) {
                        oxa = new List<OpenXmlAttribute>();
                        // this is the data type ("t"), with CellValues.String ("str")
                        oxa.Add(new OpenXmlAttribute("t", null, "str"));

                        // it's suggested you also have the cell reference, but
                        // you'll have to calculate the correct cell reference yourself.
                        // Here's an example:
                        //oxa.Add(new OpenXmlAttribute("r", null, "A1"));

                        oxw.WriteStartElement(ConstructCell("abc", CellValues.String, 2), oxa);

                        oxw.WriteElement(new CellValue(string.Format("R{0}C{1}", i, j)));

                        // this is for Cell
                        oxw.WriteEndElement();
                    }

                    // this is for Row
                    oxw.WriteEndElement();
                }

                // this is for SheetData
                oxw.WriteEndElement();
                // this is for Worksheet
                oxw.WriteEndElement();
                oxw.Close();

                oxw = OpenXmlWriter.Create(xl.WorkbookPart);
                oxw.WriteStartElement(new Workbook());
                oxw.WriteStartElement(new Sheets());

                // you can use object initialisers like this only when the properties
                // are actual properties. SDK classes sometimes have property-like properties
                // but are actually classes. For example, the Cell class has the CellValue
                // "property" but is actually a child class internally.
                // If the properties correspond to actual XML attributes, then you're fine.
                oxw.WriteElement(new Sheet() {
                    Name = "Sheet1",
                    SheetId = 1,
                    Id = xl.WorkbookPart.GetIdOfPart(wsp)
                });

                // this is for Sheets
                oxw.WriteEndElement();
                // this is for Workbook
                oxw.WriteEndElement();
                oxw.Close();

                xl.Close();
            }
        }

        public void CreateExcelDoc(string fileName, bool a)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, true)) {
                WorkbookPart workbookpart = null;
                WorksheetPart worksheetpart = null;
                OpenXmlWriter writer = null;
                Cell cell = null;
                Row row = null;

                spreadsheetDocument.CompressionOption = CompressionOption.SuperFast;
                workbookpart = spreadsheetDocument.WorkbookPart;
                worksheetpart = workbookpart.WorksheetParts.First();

                WorkbookStylesPart wbsp = workbookpart.AddNewPart<WorkbookStylesPart>();
                // add styles to sheet
                wbsp.Stylesheet = GenerateStylesheet();
                wbsp.Stylesheet.Save();

                writer = OpenXmlWriter.Create(worksheetpart);
                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());

                foreach (Data excelRowData in DataCollector) {
                    row = new Row();

                    // Add Data
                    int columnIndex = 0;
                    bool isRowData = false;

                    if (excelRowData.Values.Count != 0) {
                        foreach (string value in excelRowData.Values) {
                            cell = new Cell();
                            row.Append(cell);

                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(value);
                            cell.StyleIndex = 2;
                            if (columnIndex == (excelRowData.Values.Count - 1)) {
                                cell.StyleIndex = 0;
                            }
                        }
                        if (isRowData) {
                            isRowData = false;
                        }
                    }

                    writer.WriteElement(row);
                }

                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.Close();
                workbookpart.Workbook.Save();
                spreadsheetDocument.Close();
            }
        }

        private static Cell CreateSpreadsheetCellIfNotExist(Worksheet worksheet, int columnIndex, int rowIndex)
        {
            string columnName = GetColumnName(columnIndex);
            string cellName = columnName + rowIndex;

            IEnumerable<Row> rows = worksheet.Descendants<Row>().Where(r => r.RowIndex.Value == rowIndex);
            Cell cell;

            // If the Worksheet does not contain the specified row, create the specified row.
            // Create the specified cell in that row, and insert the row into the Worksheet.
            if (!rows.Any()) {
                Row row = new Row() { RowIndex = new UInt32Value((uint)rowIndex) };
                cell = new Cell() { CellReference = new StringValue(cellName) };
                row.Append(cell);
                worksheet.Descendants<SheetData>().First().Append(row);
                worksheet.Save();
            }
            else {
                Row row = rows.First();

                IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellName);

                // If the row does not contain the specified cell, create the specified cell.
                if (!cells.Any()) {
                    cell = new Cell() { CellReference = new StringValue(cellName) };
                    row.Append(cell);
                    worksheet.Save();
                }
                else {
                    cell = cells.First();
                }
            }

            return cell;
        }

        public static string GetColumnName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            var value = "";

            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];

            value += letters[index % letters.Length];

            return value;
        }

        private Cell ConstructCell(string value, CellValues dataType, uint styleIndex = 0)
        {
            return new Cell() {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),
                StyleIndex = styleIndex
            };
        }

        private Stylesheet GenerateStylesheet()
        {
            Stylesheet styleSheet = null;

            Fonts fonts = new Fonts(
                new Font( // Index 0 - default
                    new FontSize() { Val = 10 }

                ),
                new Font( // Index 1 - header
                    new FontSize() { Val = 10 },
                    new Bold(),
                    new Color() { Rgb = "FFFFFF" }

                ));

            Fills fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "990000" } }) { PatternType = PatternValues.Solid }), // Index 2 - header
                     new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "666666" } }) { PatternType = PatternValues.Solid })

                );

            Borders borders = new Borders(
                    new Border(), // index 0 default
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true, Rgb = "FF0000" }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new Border(),
                        new DiagonalBorder())
                );

            CellFormats cellFormats = new CellFormats(
                    new CellFormat(), // default
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }, // body
                    new CellFormat { FontId = 1, FillId = 2, BorderId = 1, ApplyFill = true },// header
                      new CellFormat { FontId = 1, FillId = 3, BorderId = 1, ApplyFill = true } // header
                );

            styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);

            return styleSheet;
        }
    }
}