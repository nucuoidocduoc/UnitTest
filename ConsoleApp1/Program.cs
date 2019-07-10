using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            ProcessFillDataExcelFile();
        }

        public static void ProcessFillDataExcelFile()
        {
            using (var spreadsheet = SpreadsheetDocument.Create("D:\\output.xlsx", SpreadsheetDocumentType.Workbook)) {
                Console.WriteLine("Creating workbook");
                spreadsheet.AddWorkbookPart();
                spreadsheet.WorkbookPart.Workbook = new Workbook();
                Console.WriteLine("Creating worksheet");
                var wsPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
                wsPart.Worksheet = new Worksheet();

                var stylesPart = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet();

                Console.WriteLine("Creating styles");

                // blank font list
                stylesPart.Stylesheet.Fonts = new Fonts();
                stylesPart.Stylesheet.Fonts.Count = 1;
                stylesPart.Stylesheet.Fonts.AppendChild(new Font());

                // create fills
                stylesPart.Stylesheet.Fills = new Fills();

                // create a solid red fill
                var solidRed = new PatternFill() { PatternType = PatternValues.Solid };
                solidRed.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FFFF0000") }; // red fill
                solidRed.BackgroundColor = new BackgroundColor { Indexed = 64 };

                stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
                stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
                stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = solidRed });
                stylesPart.Stylesheet.Fills.Count = 3;

                // blank border list
                stylesPart.Stylesheet.Borders = GetBorder();
                stylesPart.Stylesheet.Borders.Count = 1;
                stylesPart.Stylesheet.Borders.AppendChild(GetBorder());

                // blank cell format list
                stylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();
                stylesPart.Stylesheet.CellStyleFormats.Count = 1;
                stylesPart.Stylesheet.CellStyleFormats.AppendChild(new CellFormat());

                // cell format list
                stylesPart.Stylesheet.CellFormats = new CellFormats();
                // empty one for index 0, seems to be required
                stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat());
                // cell format references style format 0, font 0, border 0, fill 2 and applies the fill
                stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 0, BorderId = 0, FillId = 2, ApplyFill = true }).AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Center });
                stylesPart.Stylesheet.CellFormats.Count = 2;

                stylesPart.Stylesheet.Save();

                Console.WriteLine("Creating sheet data");
                var sheetData = wsPart.Worksheet.AppendChild(new SheetData());

                Console.WriteLine("Adding rows / cells...");

                var row = sheetData.AppendChild(new Row());
                row.AppendChild(new Cell() { CellValue = new CellValue("This"), DataType = CellValues.String });
                row.AppendChild(new Cell() { CellValue = new CellValue("is"), DataType = CellValues.String });
                row.AppendChild(new Cell() { CellValue = new CellValue("a"), DataType = CellValues.String });
                row.AppendChild(new Cell() { CellValue = new CellValue("test."), DataType = CellValues.String });

                sheetData.AppendChild(new Row());

                row = sheetData.AppendChild(new Row());
                row.AppendChild(new Cell() { CellValue = new CellValue("Value:"), DataType = CellValues.String });
                row.AppendChild(new Cell() { CellValue = new CellValue("123"), DataType = CellValues.Number });
                row.AppendChild(new Cell() { CellValue = new CellValue("Formula:"), DataType = CellValues.String, StyleIndex = 1 });
                // style index = 1, i.e. point at our fill format
                row.AppendChild(new Cell() { CellFormula = new CellFormula("B3"), DataType = CellValues.Number, StyleIndex = 1 });

                wsPart.Worksheet.Save();

                var sheets = spreadsheet.WorkbookPart.Workbook.AppendChild(new Sheets());
                sheets.AppendChild(new Sheet() { Id = spreadsheet.WorkbookPart.GetIdOfPart(wsPart), SheetId = 1, Name = "Test" });

                spreadsheet.WorkbookPart.Workbook.Save();

                Console.WriteLine("Done.");
            }
        }

        private static Borders GetBorder()
        {
            Borders borders = new Borders() { Count = (UInt32Value)1U };

            Border border1 = new Border();

            LeftBorder leftBorder1 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color colorBorder1 = new Color() { Rgb = "#000000" };
            leftBorder1.Append(colorBorder1);

            RightBorder rightBorder1 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color colorBorder2 = new Color() { Rgb = "#000000" };
            rightBorder1.Append(colorBorder2);

            TopBorder topBorder1 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color colorBorder3 = new Color() { Rgb = "#000000" };
            topBorder1.Append(colorBorder3);

            BottomBorder bottomBorder1 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color colorBorder4 = new Color() { Rgb = "#000000" };
            bottomBorder1.Append(colorBorder4);
            DiagonalBorder diagonalBorder1 = new DiagonalBorder() { Style = BorderStyleValues.Thin };

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            borders.Append(border1);
            return borders;
        }

        public static void CreateValidator(Worksheet ws, string dataContainingSheet)
        {
            /***  DATA VALIDATION CODE ***/
            DataValidations dataValidations = new DataValidations();
            DataValidation dataValidation = new DataValidation {
                Type = DataValidationValues.List,
                AllowBlank = true,
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A2" }
            };

            dataValidation.Append(
                //new Formula1 { Text = "\"FirstChoice,SecondChoice,ThirdChoice\"" }
                new Formula1(string.Format("'{0}'!$A:$A", dataContainingSheet))
                );
            dataValidations.Append(dataValidation);

            var wsp = ws.WorksheetPart;
            wsp.Worksheet.AppendChild(dataValidations);
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
    }
}