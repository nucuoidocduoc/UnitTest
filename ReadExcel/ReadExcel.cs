using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    public static class ReadExcel
    {
        public static void ReadExcelFile()
        {
            using (FileStream fs = new FileStream("D:\\TestExcel.xlsx", FileMode.Open, FileAccess.ReadWrite, FileShare.Write)) {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, true)) {
                    WorkbookPart workbookPart = doc.WorkbookPart;

                    // add styles to sheet
                    //wbsp.Stylesheet = GenerateStylesheet();
                    //wbsp.Stylesheet.Save();

                    var relId = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name.Equals("Sheet1"));

                    WorksheetPart worksheetPart = workbookPart.GetPartById(relId.Id) as WorksheetPart;
                    Worksheet sheet = worksheetPart.Worksheet;

                    foreach (Row row in sheet.Descendants<Row>()) {
                        foreach (Cell cell in row.Descendants<Cell>()) {
                            cell.StyleIndex = 2;
                        }
                    }

                    workbookPart.Workbook.Save();
                    doc.Close();
                }
            }
        }

        private static Stylesheet GenerateStylesheet()
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