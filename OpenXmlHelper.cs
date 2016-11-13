using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Collections.Generic;
using System;

namespace ITV.Reporting.WorkTime.Code
{
    public class OpenXmlCreator : System.IDisposable
    {
        public OpenXmlWriter Writer;
        private SpreadsheetDocument _document;
        private uint _rowIndex = 0;
        private int _columnNumber = 0;

        private List<MergeCell> _mergeCells;

        public OpenXmlCreator(System.IO.Stream stream, string reportName)
        {
            if (reportName.Length > 31)
            {
                reportName = reportName.Remove(28);
                reportName += "...";
            }
            _document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

            // create the workbook
            var workbookPart = _document.AddWorkbookPart();

            

            var workbook = workbookPart.Workbook = new Workbook();
            workbook.Append(new BookViews(new WorkbookView()));

            AddStyles(_document);

            var sheets = workbook.AppendChild<Sheets>(new Sheets());

            _mergeCells = new List<MergeCell>();

            // create worksheet 1
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = reportName };
            sheets.Append(sheet);

            Writer = OpenXmlWriter.Create(worksheetPart);

            var attr = new List<OpenXmlAttribute> { new OpenXmlAttribute("mc:Ignorable", null, "x14ac") };
            var ns = new Dictionary<string, string>();
            ns["r"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            ns["mc"] = "http://schemas.openxmlformats.org/markup-compatibility/2006";
            ns["x14ac"] = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";

            Writer.WriteStartElement( new Worksheet(), attr, ns);

            Writer.WriteStartElement(new SheetData());

            _document.WorkbookPart.Workbook.Save();
                       
        }

        public void Dispose()
        {
            Writer.WriteEndElement();

            insertMergeInDocument();

            Writer.WriteEndElement();
            Writer.Close();
            
            _document.Dispose();
        }

        // Required opening and closing rows. You can not open a few rows in a row and close a few in a row. 
        // Follow the exact string Open - close the rows.
        internal void OpenRow()
        {
            Writer.WriteStartElement(new Row() { RowIndex = ++_rowIndex });
            _columnNumber = 0;
        }

        internal void CloseRow()
        {
            Writer.WriteEndElement();
        }

        //the addition of different types of cells, depending on the passed parameter.
        internal void AddCell(int cellValue, int styleID = 0, int colspan = 1)
        {
            AddCell(cellValue.ToString(), CellValues.Number, styleID, colspan);
        }

        internal void AddCell(string cellValue, int styleID = 0, int colspan = 1)
        {
            AddCell(cellValue, CellValues.String, styleID, colspan);
        }

        internal void AddHeaderCell(string cellValue, int styleID = 1, int colspan = 1)
        {
            AddCell(cellValue, CellValues.String, styleID, colspan);
        }

        internal void AddCell(bool cellValue, int styleID = 0, int colspan = 1)
        {
            AddCell(cellValue.ToString() == "True" ? "1" : "0", CellValues.Boolean, styleID, colspan);
        }
        internal void AddCell(DateTime cellValue, int styleID = 0, int colspan = 1)
        {
            AddCell(cellValue.ToString(), CellValues.String, styleID, colspan);
        }

        //Write the transmitted cells
        private void AddCell(string cellValue, CellValues dataType, int styleID = 0, int colspan = 1)
        {
            int startColumnNumber = _columnNumber;
            var attributes = new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, styleID.ToString()) }.ToList();
            attributes.Add(new OpenXmlAttribute("r", null, GetCellReference(_rowIndex, _columnNumber++)));
            switch (dataType)
            {
                case (CellValues.String):
                    {
                        attributes.Add(new OpenXmlAttribute("t", null, "inlineStr"));
                        Writer.WriteStartElement(new Cell() { DataType = dataType}, attributes);
                        Writer.WriteElement(new InlineString(new Text(cellValue)));
                        Writer.WriteEndElement();
                        break;
                    }
                case (CellValues.Boolean):
                    {
                        attributes.Add(new OpenXmlAttribute("t", null, "b"));
                        Writer.WriteStartElement(new Cell() { DataType = dataType }, attributes);
                        Writer.WriteElement(new CellValue(cellValue));
                        Writer.WriteEndElement();
                        break;
                    }
                default:
                    {
                        Writer.WriteStartElement(new Cell() { DataType = dataType }, attributes);
                        Writer.WriteElement(new CellValue(cellValue));
                        Writer.WriteEndElement();
                        break;
                    }

            }
            if (colspan > 1)
            {
                for (int i = 1; i < colspan; i++)
                {
                    var attr = new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, styleID.ToString()) }.ToList();
                    attr.Add(new OpenXmlAttribute("r", null, GetCellReference(_rowIndex, _columnNumber++)));

                    attr.Add(new OpenXmlAttribute("t", null, "inlineStr"));
                    Writer.WriteStartElement(new Cell() { DataType = dataType }, attr);
                    Writer.WriteElement(new InlineString(new Text("")));
                    Writer.WriteEndElement();
                }

                addMergeCellElement(_rowIndex, startColumnNumber, colspan);
            }

        }


        public void addMergeCellElement(uint rowNumber, int columnNumber, int colspan)
        {
            MergeCell NewMerge = new MergeCell() { Reference = GetCellReference(rowNumber, columnNumber) + ":" + 
                GetCellReference(rowNumber, columnNumber + colspan - 1)};
            _mergeCells.Add(NewMerge);

        }

        private void insertMergeInDocument ()
        {
            Writer.WriteStartElement(new MergeCells());
            foreach (MergeCell c in _mergeCells)
            {
                Writer.WriteElement(c);
            }
            Writer.WriteEndElement();
        }


        private static string GetCellReference(uint rowNumber, int columnNumber)
        {
            return GetExcelColumnName(columnNumber) + rowNumber;
        }

        private static string GetExcelColumnName(int columnIndex)
        {
            //  eg  (0) should return "A"
            //      (1) should return "B"
            //      (25) should return "Z"
            //      (26) should return "AA"
            //      (27) should return "AB"
            //      ..etc..
            char firstChar;
            char secondChar;
            char thirdChar;

            if (columnIndex < 26)
            {
                return ((char)('A' + columnIndex)).ToString();
            }

            if (columnIndex < 702)
            {
                firstChar = (char)('A' + (columnIndex / 26) - 1);
                secondChar = (char)('A' + (columnIndex % 26));

                return string.Format("{0}{1}", firstChar, secondChar);
            }

            int firstInt = columnIndex / 26 / 26;
            int secondInt = (columnIndex - firstInt * 26 * 26) / 26;
            if (secondInt == 0)
            {
                secondInt = 26;
                firstInt = firstInt - 1;
            }
            int thirdInt = (columnIndex - firstInt * 26 * 26 - secondInt * 26);

            firstChar = (char)('A' + firstInt - 1);
            secondChar = (char)('A' + secondInt - 1);
            thirdChar = (char)('A' + thirdInt);

            return string.Format("{0}{1}{2}", firstChar, secondChar, thirdChar);
        }

        //adding styles, one of which, on the index will be selected when you insert the cell.
        private static void AddStyles(SpreadsheetDocument doc)
        {
            WorkbookStylesPart workbookStylesPart = doc.WorkbookPart.AddNewPart<WorkbookStylesPart>("rIdStyles");
            Stylesheet stylesheet = new Stylesheet();
            stylesheet.Fonts = new Fonts(new Font(),
                new Font() { Color = new Color() { Rgb = new HexBinaryValue("FFFFFFFF") }, Bold = new Bold() });
            stylesheet.Fills = new Fills(

                new Fill(new PatternFill() { PatternType = PatternValues.None }),

                new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }),

                new Fill(new PatternFill()
                {
                    ForegroundColor = new ForegroundColor() { Rgb = new HexBinaryValue("FF2B85C8") },
                    PatternType = PatternValues.Solid
                }),
                new Fill(new PatternFill()
                {
                    ForegroundColor = new ForegroundColor() { Rgb = new HexBinaryValue("FFFF0000") },
                    PatternType = PatternValues.Solid
                })
                );

            stylesheet.Borders = new Borders(new Border(
                    new LeftBorder(),
                    new RightBorder(),
                    new TopBorder(),
                    new BottomBorder(),
                    new DiagonalBorder()),
                new Border(
                    new LeftBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Indexed = (UInt32Value)64U } },
                    new RightBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Indexed = (UInt32Value)64U } },
                    new TopBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Indexed = (UInt32Value)64U } },
                    new BottomBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Indexed = (UInt32Value)64U } },
                    new DiagonalBorder()));
            stylesheet.CellFormats = new CellFormats(
                new CellFormat(),
                new CellFormat() { FontId = 1, BorderId = 0, FillId = 2, ApplyFill = true, ApplyFont = true },
                new CellFormat() { BorderId = 1 },
                new CellFormat() { BorderId = 1, FillId = 3, ApplyFill = true, ApplyFont = true });
            workbookStylesPart.Stylesheet = stylesheet;
        }
    }
}
