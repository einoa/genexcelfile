using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;


namespace GenExcelFile1
{
    class Program
    {
        static void Main(string[] args)
        {
            /*
             * We will generate 4 XML documents.
             *  1) /xl/workbook.xml
             *  2) /xl/styles.xml
             *  3) /xl/worksheets/sheet1.xml
             *  4) /xl/sharedStrings.xml
             */

            const string nsSpreadsheetML = @"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            const string nsRelationships = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            const string contentTypeMain = @"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
            const string contentTypeWorksheet = @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
            const string contentTypeStyles = @"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
            const string contentTypeSharedStrings = @"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";

            System.IO.FileStream fs = new FileStream(@"C:\TMP\z.xlsx", FileMode.Create);
            Package pkg = Package.Open(fs, FileMode.Create);

            // ************************************************************************************************************************
            // ************************************************************************************************************************
            // styles.xml
            // ************************************************************************************************************************
            // ************************************************************************************************************************

            // Stylesheet document
            XmlDocument xmlStylesheetDoc = new XmlDocument();

            // Stylesheet root
            XmlElement xmlStylesheet = xmlStylesheetDoc.CreateElement("styleSheet", nsSpreadsheetML);
            xmlStylesheetDoc.AppendChild(xmlStylesheet);

            // ************************************************************************************************************************
            // Stylesheet - Fonts
            // ************************************************************************************************************************

            XmlElement xmlFonts = xmlStylesheetDoc.CreateElement("fonts", nsSpreadsheetML);
            xmlStylesheet.AppendChild(xmlFonts);

            // ------------------------------------------------------------------------------------------------------------------------
            // Font
            // ------------------------------------------------------------------------------------------------------------------------

            // Stylesheet - Fonts - Font
            XmlElement xFont = xmlStylesheetDoc.CreateElement("font", nsSpreadsheetML);
            xmlFonts.AppendChild(xFont);

            // Stylesheet - Fonts - Font - Sz
            XmlElement xSz = xmlStylesheetDoc.CreateElement("sz", nsSpreadsheetML);
            xFont.AppendChild(xSz);
            xSz.SetAttribute("val", "11");

            // Stylesheet - Fonts - Font - Val
            XmlElement xName = xmlStylesheetDoc.CreateElement("name", nsSpreadsheetML);
            xFont.AppendChild(xName);
            xName.SetAttribute("val", "Calibri");

            // ************************************************************************************************************************
            // Stylesheet - Fills
            // ************************************************************************************************************************

            XmlElement xFills = xmlStylesheetDoc.CreateElement("fills", nsSpreadsheetML);
            xmlStylesheet.AppendChild(xFills);

            // ------------------------------------------------------------------------------------------------------------------------
            // Stylesheet - Fills - Fill
            // ------------------------------------------------------------------------------------------------------------------------
            XmlElement xFill = xmlStylesheetDoc.CreateElement("fill", nsSpreadsheetML);
            xFills.AppendChild(xFill);

            // Stylesheet - Fills - Fill - PatternFill
            XmlElement xPatternFill = xmlStylesheetDoc.CreateElement("patternFill", nsSpreadsheetML);
            xFill.AppendChild(xPatternFill);
            xPatternFill.SetAttribute("patternType", "none");

            // ------------------------------------------------------------------------------------------------------------------------
            // Stylesheet - Fills - Fill
            // ------------------------------------------------------------------------------------------------------------------------
            xFill = xmlStylesheetDoc.CreateElement("fill", nsSpreadsheetML);
            xFills.AppendChild(xFill);

            // Stylesheet - Fills - Fill - PatternFill
            xPatternFill = xmlStylesheetDoc.CreateElement("patternFill", nsSpreadsheetML);
            xFill.AppendChild(xPatternFill);
            xPatternFill.SetAttribute("patternType", "gray125");

            // ************************************************************************************************************************
            // Stylesheet - Borders
            // ************************************************************************************************************************

            XmlElement xBorders = xmlStylesheetDoc.CreateElement("borders", nsSpreadsheetML);
            xmlStylesheet.AppendChild(xBorders);

            // ------------------------------------------------------------------------------------------------------------------------
            // Stylesheet - Borders - Border
            // ------------------------------------------------------------------------------------------------------------------------
            XmlElement xBorder = xmlStylesheetDoc.CreateElement("border", nsSpreadsheetML);
            xBorders.AppendChild(xBorder);

            // Stylesheet - Borders - Border - Left Right Top Bottom Diagonal
            xBorder.AppendChild(xmlStylesheetDoc.CreateElement("left", nsSpreadsheetML));
            xBorder.AppendChild(xmlStylesheetDoc.CreateElement("right", nsSpreadsheetML));
            xBorder.AppendChild(xmlStylesheetDoc.CreateElement("top", nsSpreadsheetML));
            xBorder.AppendChild(xmlStylesheetDoc.CreateElement("bottom", nsSpreadsheetML));
            xBorder.AppendChild(xmlStylesheetDoc.CreateElement("diagonal", nsSpreadsheetML));

            // ************************************************************************************************************************
            // Stylesheet - CellStyleXfs
            // ************************************************************************************************************************

            XmlElement xCellStyleXfs = xmlStylesheetDoc.CreateElement("cellStyleXfs", nsSpreadsheetML);
            xmlStylesheet.AppendChild(xCellStyleXfs);

            // Stylesheet - CellStyleXfs - Xf
            XmlElement xXf = xmlStylesheetDoc.CreateElement("xf", nsSpreadsheetML);
            xCellStyleXfs.AppendChild(xXf);
            xXf.SetAttribute("numFmtId", "0");
            xXf.SetAttribute("fontId", "0");
            xXf.SetAttribute("fillId", "0");
            xXf.SetAttribute("borderId", "0");

            // ************************************************************************************************************************
            // Stylesheet - CellXfs
            // ************************************************************************************************************************

            XmlElement xCellXfs = xmlStylesheetDoc.CreateElement("cellXfs", nsSpreadsheetML);
            xmlStylesheet.AppendChild(xCellXfs);

            // Stylesheet - CellXfs - Xf
            xXf = xmlStylesheetDoc.CreateElement("xf", nsSpreadsheetML);
            xCellXfs.AppendChild(xXf);
            xXf.SetAttribute("numFmtId", "0");
            xXf.SetAttribute("fontId", "0");
            xXf.SetAttribute("fillId", "0");
            xXf.SetAttribute("borderId", "0");
            xXf.SetAttribute("xfId", "0");

            // ************************************************************************************************************************
            // Write styles.xml
            // ************************************************************************************************************************

            Uri uriStylesheet = new Uri("/xl/styles.xml", UriKind.Relative);
            PackagePart ppStylesheet = pkg.CreatePart(uriStylesheet, contentTypeStyles);
            StreamWriter swStylesheet = new StreamWriter(ppStylesheet.GetStream(FileMode.Create, FileAccess.Write));
            xmlStylesheetDoc.Save(swStylesheet);
            swStylesheet.Close();
            pkg.Flush();

            // ************************************************************************************************************************
            // ************************************************************************************************************************
            // workbook.xml
            // ************************************************************************************************************************
            // ************************************************************************************************************************

            System.Collections.Generic.List<string> sharedStrings = new List<string>(); // *** sharedStrings

            // Workbook document
            XmlDocument xmlWorkbookDoc = new XmlDocument();

            // Workbook root
            XmlElement xmlWorkbookRoot = xmlWorkbookDoc.CreateElement("workbook", nsSpreadsheetML);
            xmlWorkbookDoc.AppendChild(xmlWorkbookRoot);
            XmlAttribute xmlAttr = xmlWorkbookDoc.CreateAttribute("xmlns", "r", @"http://www.w3.org/2000/xmlns/");
            xmlAttr.Value = nsRelationships;
            xmlWorkbookRoot.Attributes.Append(xmlAttr);

            // ************************************************************************************************************************
            // Workbook - Sheets
            // ************************************************************************************************************************

            XmlElement xSheets = xmlWorkbookDoc.CreateElement("sheets", nsSpreadsheetML);
            xmlWorkbookRoot.AppendChild(xSheets);

            // ************************************************************************************************************************
            // Workbook - Sheets - Sheet
            // ************************************************************************************************************************
            XmlElement xSheet = xmlWorkbookDoc.CreateElement("sheet", nsSpreadsheetML);
            xSheets.AppendChild(xSheet);
            xSheet.SetAttribute("name", "Sheet1");
            xSheet.SetAttribute("sheetId", "1");
            xSheet.SetAttribute("id", nsRelationships, "rId1");

            // Worksheet document
            XmlDocument xmlWorksheetDoc = new XmlDocument();

            // Worksheet node
            XmlElement xWorksheet = xmlWorksheetDoc.CreateElement("worksheet", nsSpreadsheetML);
            xmlWorksheetDoc.AppendChild(xWorksheet);
            xWorksheet.SetAttribute("xmlns:r", nsRelationships);

            // --------------------------------------------------------------------------------
            // Worksheet - SheetViews
            // --------------------------------------------------------------------------------
            XmlElement xSheetViews = xmlWorksheetDoc.CreateElement("sheetViews", nsSpreadsheetML);
            xWorksheet.AppendChild(xSheetViews);

            // Worksheet - SheetViews - SheetView
            XmlElement xSheetView = xmlWorksheetDoc.CreateElement("sheetView", nsSpreadsheetML);
            xSheetViews.AppendChild(xSheetView);
            xSheetView.SetAttribute("workbookViewId", "0");

            // ********************************************************************************
            // Worksheet - SheetData
            // ********************************************************************************
            XmlElement xSheetData = xmlWorksheetDoc.CreateElement("sheetData", nsSpreadsheetML);
            xWorksheet.AppendChild(xSheetData);


            string stringvalue;
            XmlElement xRow;
            XmlElement xCol;
            XmlElement value;
            int currentRow = 0;
            int currentCol = 0;


            // row
            currentRow = 1;
            xRow = xmlWorksheetDoc.CreateElement("row", nsSpreadsheetML);
            xSheetData.AppendChild(xRow);
            xRow.SetAttribute("r", (currentRow).ToString());
            // col
            currentCol = 1;
            stringvalue = "abc";
            xCol = xmlWorksheetDoc.CreateElement("c", nsSpreadsheetML);
            xRow.AppendChild(xCol);
            xCol.SetAttribute("r", ExcelColumnName(currentCol) + currentRow.ToString());
            xCol.SetAttribute("t", "s");
            xCol.SetAttribute("s", "0");
            value = xmlWorksheetDoc.CreateElement("v", nsSpreadsheetML);
            xCol.AppendChild(value);
            if (sharedStrings.Contains(stringvalue))
            {
                value.InnerText = sharedStrings.IndexOf(stringvalue).ToString();
            }
            else
            {
                value.InnerText = sharedStrings.Count().ToString();
                sharedStrings.Add(stringvalue);
            }
            // col
            currentCol = 2;
            xCol = xmlWorksheetDoc.CreateElement("c", nsSpreadsheetML);
            xRow.AppendChild(xCol);
            xCol.SetAttribute("r", ExcelColumnName(currentCol) + currentRow.ToString());
            value = xmlWorksheetDoc.CreateElement("v", nsSpreadsheetML);
            xCol.AppendChild(value);
            value.InnerText = "123.45";

            // row
            currentRow = 2;
            xRow = xmlWorksheetDoc.CreateElement("row", nsSpreadsheetML);
            xSheetData.AppendChild(xRow);
            xRow.SetAttribute("r", (currentRow).ToString());
            // col
            currentCol = 2;
            xCol = xmlWorksheetDoc.CreateElement("c", nsSpreadsheetML);
            xRow.AppendChild(xCol);
            xCol.SetAttribute("r", ExcelColumnName(currentCol) + currentRow.ToString());
            value = xmlWorksheetDoc.CreateElement("v", nsSpreadsheetML);
            xCol.AppendChild(value);
            value.InnerText = "99";


            // Write sheet1.xml
            Uri uriWorksheet = new Uri("/xl/worksheets/sheet1.xml", UriKind.Relative);
            PackagePart ppWorksheet = pkg.CreatePart(uriWorksheet, contentTypeWorksheet);
            StreamWriter swWorksheet = new StreamWriter(ppWorksheet.GetStream(FileMode.Create, FileAccess.Write));
            xmlWorksheetDoc.Save(swWorksheet);
            swWorksheet.Close();
            pkg.Flush();

            // ********************************************************************************
            // ********************************************************************************

            // SharedStrings document
            XmlDocument xmlSharedStringsDoc = new XmlDocument();

            // SharedStrings - Sst
            XmlElement xSst = xmlSharedStringsDoc.CreateElement("sst", nsSpreadsheetML);
            xSst.SetAttribute("count", sharedStrings.Count().ToString());
            xSst.SetAttribute("uniqueCount", sharedStrings.Count().ToString());
            xmlSharedStringsDoc.AppendChild(xSst);

            for (int i = 0; i < sharedStrings.Count; i++)
            {
                XmlElement xSi = xmlSharedStringsDoc.CreateElement("si", nsSpreadsheetML);
                XmlElement xt = xmlSharedStringsDoc.CreateElement("t", nsSpreadsheetML);
                xt.InnerText = sharedStrings[i];
                xSi.AppendChild(xt);
                xSst.AppendChild(xSi);
            }

            // Write sharedStrings.xml
            Uri uriStrings = new Uri("/xl/sharedStrings.xml", UriKind.Relative);
            PackagePart ppSharedStrings = pkg.CreatePart(uriStrings, contentTypeSharedStrings);
            StreamWriter swSharedStrings = new StreamWriter(ppSharedStrings.GetStream(FileMode.Create, FileAccess.Write));
            xmlSharedStringsDoc.Save(swSharedStrings);
            swSharedStrings.Close();
            pkg.Flush();


            // Write workbook.xml
            Uri uriWorkbook = new Uri("/xl/workbook.xml", UriKind.Relative);
            PackagePart ppWorkbook = pkg.CreatePart(uriWorkbook, contentTypeMain);
            StreamWriter swWorkbook = new StreamWriter(ppWorkbook.GetStream(FileMode.Create, FileAccess.Write));
            xmlWorkbookDoc.Save(swWorkbook);
            swWorkbook.Close();
            pkg.Flush();


            // Write relationships
            pkg.CreateRelationship(uriWorkbook, TargetMode.Internal, nsRelationships + "/officeDocument", "rId1");
            ppWorkbook.CreateRelationship(uriWorksheet, TargetMode.Internal, nsRelationships + "/worksheet", "rId1");
            ppWorkbook.CreateRelationship(uriStylesheet, TargetMode.Internal, nsRelationships + "/styles", "rId2");
            ppWorkbook.CreateRelationship(uriStrings, TargetMode.Internal, nsRelationships + "/sharedStrings", "rId3");
            pkg.Flush();

            // Close
            pkg.Close();

            fs.Close();

        }

        /*
         *  Function: ExcelColumnName
         *  Input: column number - 1 based - first column is 1
         *     1 returns A
         *     2 returns B
         *     3 returns C
         *    26 returns Z
         *    27 returns AA
         *    28 returns AB
         *    etc.
         */
        public static string ExcelColumnName(int columnNumber)
        {
            string a = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            if (columnNumber <= 26)
            {
                return (a[columnNumber - 1]).ToString();
            }
            else if (columnNumber <= 702)
            {
                int n1 = columnNumber - 27;
                int n2 = n1 / 26;
                int n3 = n1 - 26 * n2;

                char[] t = new char[2];
                t[0] = a[n2];
                t[1] = a[n3];
                return new string(t);
            }
            else
            {
                int n1 = columnNumber - 703;
                int n2 = n1 / 676;
                int n3 = n1 - 676 * n2;
                int n4 = n3 / 26;
                int n5 = n3 - 26 * n4;

                char[] t = new char[3];
                t[0] = a[n2];
                t[1] = a[n4];
                t[2] = a[n5];

                return new string(t);
            }
        }
    }
}
