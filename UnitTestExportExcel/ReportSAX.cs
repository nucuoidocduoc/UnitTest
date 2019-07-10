using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTestExportExcel
{
    public class DataX
    {
        public static int MAXRowNumber = 10000;
        public static int MAXColNumber = 10;

        public static string[] Headers = new string[]
        {
            "Input1", "Input2", "Input3", "Input4", "Input5", "Input6", "Input7", "Input8", "Result1", "Result2"
        };

        public static string[] GetData(int index)
        {
            return new string[]
            {
                $"Value1_{index}", $"Value2_{index}",
                $"Value3_{index}", $"Value4_{index}",
                $"Value5_{index}", $"Value6_{index}",
                $"Value7_{index}", $"Value8_{index}",
                $"Result9_{index}", $"Result10_{index}"
            };
        }
    }

    public static class WriteDataToExcelSAX
    {
        private static readonly string FileName = @"D:\test.xlsx";

        public static void WriteDataSAX()
        {
            // (1) Create a file (Document)
            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(FileName, SpreadsheetDocumentType.Workbook)) {
                // (2) Add Workbook to Doc
                WorkbookPart workbookPart = doc.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // (3) Add a style sheet
                WorkbookStylesPart stylePart = AddStyleSheet(doc);

                // (4) Add Sheets to Workbook
                Sheets sheets = doc.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                // (5.1) Add the first Worksheet
                var workSheetPart1 = AddWorksheet(doc, 1, "My Page 1");
                WriteData(workSheetPart1);

                // (5.2) Add the second Worksheet
                var workSheetPart2 = AddWorksheet(doc, 2, "My Page 2");
                WriteData(workSheetPart2);

                // (5.3) Add the 3 Worksheet
                var workSheetPart3 = AddWorksheet(doc, 3, "My Page 3");
                WriteData(workSheetPart3);

                // (6) Save

                doc.Save();

                // (7) Close the document
                doc.Close();
            }
        }

        private static WorksheetPart AddWorksheet(SpreadsheetDocument doc, int sheetId, string sheetName)
        {
            // (1) Add WorksheetPart to WorkbookPart
            WorksheetPart worksheetPart = doc.WorkbookPart.AddNewPart<WorksheetPart>();

            // (2) Get a worksheetId
            var workSheetPartId = doc.WorkbookPart.GetIdOfPart(worksheetPart);

            // (3) Create a shee
            Sheet sheet = new Sheet() {
                Id = workSheetPartId,
                SheetId = Convert.ToUInt32(sheetId),
                Name = sheetName
            };
            doc.WorkbookPart.Workbook.Sheets.Append(sheet);

            return worksheetPart;
        }

        private static void WriteData(WorksheetPart workSheetPart)
        {
            OpenXmlWriter writer = OpenXmlWriter.Create(workSheetPart);
            {
                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());

                //// Write Headers
                //var headers = DataX.Headers;
                //Row rHeader = new Row() {
                //    RowIndex = Convert.ToUInt32(1)
                //};
                //writer.WriteStartElement(rHeader);
                //foreach (string headerValue in headers) {
                //    Cell c = new Cell() {
                //        StyleIndex =
                //            headerValue.StartsWith("Result")
                //            ? Convert.ToUInt32(2)
                //            : Convert.ToUInt32(1)
                //    };
                //    CellValue v = new CellValue(headerValue);
                //    c.DataType = new EnumValue<CellValues>(CellValues.String);
                //    c.Append(v);
                //    writer.WriteElement(c);
                //}
                //writer.WriteEndElement();

                // Write Values
                for (int row = 0; row < DataX.MAXRowNumber; row++) {
                    Row r = new Row() {
                        RowIndex = Convert.ToUInt32(row + 2)
                    }; ;
                    writer.WriteStartElement(r);

                    var values = DataX.GetData(row + 1);

                    for (int col = 0; col < DataX.MAXColNumber; col++) {
                        Cell c = new Cell() {
                            StyleIndex = values[col].StartsWith("Result")
                            ? Convert.ToUInt32(4)
                            : Convert.ToUInt32(3)
                        };
                        CellValue v = new CellValue(values[col]);
                        c.DataType = new EnumValue<CellValues>(CellValues.String);
                        c.Append(v);
                        writer.WriteElement(c);
                    }
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
                writer.WriteEndElement();

                writer.Close();
            }
        }

        private static WorkbookStylesPart AddStyleSheet(SpreadsheetDocument doc)
        {
            WorkbookStylesPart stylesheet = doc.WorkbookPart.AddNewPart<WorkbookStylesPart>();

            Stylesheet workbookstylesheet = new Stylesheet();

            // (1) Fonts
            Font defaultFont = new Font();

            Font boldFont = new Font();
            boldFont.Append(new Bold());

            Fonts fonts = new Fonts();
            fonts.Append(defaultFont); // index: 0
            fonts.Append(boldFont); // index: 1
            fonts.Count = Convert.ToUInt32(2);

            // (2) Fills
            Fill defaultFill = new Fill(); // required

            Fill greyPatternFill = new Fill(); // required
            greyPatternFill.Append(new PatternFill() { PatternType = PatternValues.Gray125 });

            Fill grayBGFill = new Fill();
            PatternFill grayBGPattern = new PatternFill {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor() { Rgb = HexBinaryValue.FromString("DDDDDD") },
                BackgroundColor = new BackgroundColor() { Indexed = (UInt32Value)64U }
            };
            grayBGFill.Append(grayBGPattern);

            Fills fills = new Fills();
            fills.Append(defaultFill); // index: 0
            fills.Append(greyPatternFill); // index: 1
            fills.Append(grayBGFill);    // index: 2
            fills.Count = Convert.ToUInt32(3);

            // (3) Borders
            Border defaultborder = new Border();

            Borders borders = new Borders();
            borders.Append(defaultborder);

            // (4) CellFormats
            CellFormat defaultCellformat = new CellFormat() {
                FontId = Convert.ToUInt32(0),
                FillId = Convert.ToUInt32(0),
                BorderId = Convert.ToUInt32(0)
            }; // Default style : Mandatory | Style Id =0

            CellFormat boldGrayBGFormat = new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center }) {
                FontId = Convert.ToUInt32(1),
                FillId = Convert.ToUInt32(2),
                ApplyFill = true
            };  // Style with Bold text + Gray BG ; Style Id = 1

            CellFormat boldFormat = new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center }) {
                FontId = Convert.ToUInt32(1)
            };  // Style with Bold text ; Style Id = 2

            CellFormat centerGrayBGFormat = new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center }) {
                FillId = Convert.ToUInt32(2),
                ApplyFill = true
            };
            // Default Style with Center Format ; Style Id = 3

            CellFormat centerFormat = new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center }) {
            };
            // Default Style with Center Format ; Style Id = 4

            CellFormats cellformats = new CellFormats();
            cellformats.Append(defaultCellformat); // index 0
            cellformats.Append(boldGrayBGFormat); // index 1
            cellformats.Append(boldFormat); // index 2
            cellformats.Append(centerGrayBGFormat); // index 3
            cellformats.Append(centerFormat); // index 4

            // Append FONTS, FILLS , BORDERS & CellFormats to stylesheet
            workbookstylesheet.Append(fonts);
            workbookstylesheet.Append(fills);
            workbookstylesheet.Append(borders);
            workbookstylesheet.Append(cellformats);

            stylesheet.Stylesheet = workbookstylesheet;
            stylesheet.Stylesheet.Save();

            return stylesheet;
        }

        private static void FreezeHeader(WorksheetPart workSheetPart)
        {
            SheetView sw = workSheetPart.Worksheet.SheetViews.FirstOrDefault() as SheetView;

            // the freeze pane
            Pane pane = new Pane() {
                VerticalSplit = 1D,
                TopLeftCell = "A2", // 1 row
                ActivePane = PaneValues.BottomLeft,
                State = PaneStateValues.Frozen
            };

            sw.Append(pane);

            workSheetPart.Worksheet.Save();
        }
    }
}