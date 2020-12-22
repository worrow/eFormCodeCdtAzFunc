using System;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using System.Globalization;
using Microsoft.Rest.Serialization;

namespace GenerateBlobSASNSave2DB
{
    public static class ExportCodeCdt
    {
        public class SummaryModel
        {
            public string BusinessUnit { get; set; }
            public string Division { get; set; }
            public string Company { get; set; }
            public string Department { get; set; }
            public string Signed { get; set; }
            public string NotSigned { get; set; }
            public string Reviewing { get; set; }
            public string Reviewed { get; set; }
            public string Total { get; set; }
        }
        public class TestModel
        {
            public string Id { get; set; }
            public string EmployeeId { get; set; }
            public string StaffNo { get; set; }
            public string DisplayName { get; set; }
            public string IsConflictOfInterest { get; set; }
            public string signed_Date { get; set; }
            public string CompanyCode { get; set; }
            public string Company { get; set; }
            public string BusinessUnit { get; set; }
            public string DivisionName { get; set; }
            public string DepartmentName { get; set; }
            public string SectionName { get; set; }
            public string yearTermId { get; set; }
            public string LastUpdateDate { get; set; }
            public string OriginalFileName { get; set; }
            public string HRStatus { get; set; }

        }

        public class TestModelList
        {
            public List<TestModel> testData { get; set; }
        }
        public class SummaryModelList
        {
            public List<SummaryModel> testData { get; set; }
        }
        private static string reportType,orderby;
        private static int yeartermId;
        [FunctionName("ExportCodeCdt")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
            string result;
            int counter = 0;
            reportType = req.Headers["reporttype"];
            orderby = req.Headers["orderby"];
           // log.LogInformation(orderby + "*************************");
            if (String.IsNullOrEmpty(reportType))
                reportType = "A";
            if (String.IsNullOrEmpty(orderby))
                orderby = "";
           
            result = String.Empty;
            if (reportType != "S")
            {
                TestModelList tmList = new TestModelList();
                tmList.testData = new List<TestModel>();
                TestModel tm;
                try
                {
                    using (SqlConnection conn = new SqlConnection(System.Environment.GetEnvironmentVariable("SQLDB")))
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            cmd.CommandText = "sp_CodeCdt_StaffList";
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@order", orderby);
                            cmd.Connection = conn;


                            conn.Open();
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {


                                while (reader.Read())
                                {
                                    tm = new TestModel();
                                    if (reader.IsDBNull(0))
                                        tm.EmployeeId = "";
                                    else
                                        tm.EmployeeId = reader.GetString(0);
                                    if (tm.EmployeeId.StartsWith('0'))
                                        tm.EmployeeId = "'" + tm.EmployeeId;
                                    if (reader.IsDBNull(1))
                                        tm.StaffNo = "";
                                    else
                                        tm.StaffNo = reader.GetString(1);
                                    if (reader.IsDBNull(2))
                                        tm.DisplayName = "";
                                    else
                                        tm.DisplayName = reader.GetString(2);
                                   
                                    if (reader.IsDBNull(3))
                                        tm.IsConflictOfInterest = "";
                                    else
                                        tm.IsConflictOfInterest = reader.GetBoolean(3).ToString();
                                    if (reader.IsDBNull(4))
                                        tm.signed_Date = "";
                                    else
                                        tm.signed_Date = reader.GetString(4);
                                    if (reader.IsDBNull(5))
                                        tm.Company = "";
                                    else
                                        tm.Company = reader.GetString(5);
                                    if (reader.IsDBNull(6))
                                        tm.BusinessUnit = "";
                                    else
                                        tm.BusinessUnit = reader.GetString(6);
                                    if (reader.IsDBNull(7))
                                        tm.DivisionName = "";
                                    else
                                        tm.DivisionName = reader.GetString(7);
                                    if (reader.IsDBNull(8))
                                        tm.DepartmentName = "";
                                    else
                                        tm.DepartmentName = reader.GetString(8);
                                    if (reader.IsDBNull(9))
                                        tm.SectionName = "";
                                    else
                                        tm.SectionName = reader.GetString(9);
                                    if (reader.IsDBNull(10))
                                        tm.LastUpdateDate = "";
                                    else
                                        tm.LastUpdateDate = reader.GetString(10); //.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                                    if (reader.IsDBNull(14))
                                        tm.OriginalFileName = "";
                                    else
                                        tm.OriginalFileName = reader.GetString(14);
                                    if (reader.IsDBNull(15))
                                        tm.HRStatus = "Invalid";
                                    else
                                    {
                                        if (reader.GetBoolean(15) == true)
                                            tm.HRStatus = "Reviewed";
                                        else
                                            tm.HRStatus = "Reviewing";
                                    }
                                    tmList.testData.Add(tm);
                                    counter++;
                                }
                            }
                        }
                    }
                    MemoryStream ms = CreateExcelFile(tmList, @"D:\local\Temp");
                    byte[] data = ms.ToArray();
                    return new FileContentResult(data, "application/octet-stream");

                    //data = File.ReadAllBytes(file.toPath());
                }
                catch (Exception ex)
                {
                    log.LogError(ex.Message);
                    return new OkObjectResult(ex.Message);
                }
                //string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
                // parse query parameter
            }
            else
            {
                log.LogInformation("***********Summary*****************");
                SummaryModelList smList = new SummaryModelList();
                smList.testData = new List<SummaryModel>();
                SummaryModel sm;
                try
                {
                    log.LogInformation(yeartermId.ToString());
                    using (SqlConnection conn = new SqlConnection(System.Environment.GetEnvironmentVariable("SQLDB")))
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            cmd.CommandText = "sp_CodeCdt_Summary";
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@orderby", orderby);
                            cmd.Connection = conn;
                            conn.Open();
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
               
                                while (reader.Read())
                                {
                                    sm = new SummaryModel();
                                    if (reader.IsDBNull(0))
                                        sm.BusinessUnit = String.Empty;
                                    else
                                        sm.BusinessUnit = reader.GetString(0);
                                    if (reader.IsDBNull(1))
                                        sm.Division = String.Empty;
                                    else
                                        sm.Division = reader.GetString(1);
                                    if (reader.IsDBNull(2))
                                        sm.Company = String.Empty;
                                    else
                                        sm.Company = reader.GetString(2);
                                    if (reader.IsDBNull(3))
                                        sm.Department = String.Empty;
                                    else
                                        sm.Department = reader.GetString(3);
                                    sm.Signed = reader.GetInt32(4).ToString();
                                    sm.NotSigned = reader.GetInt32(5).ToString();
                                    sm.Reviewing = reader.GetInt32(6).ToString();
                                    sm.Reviewed = reader.GetInt32(7).ToString();
                                    sm.Total = (reader.GetInt32(4) + reader.GetInt32(5) + reader.GetInt32(6) + reader.GetInt32(7)).ToString();
                                    smList.testData.Add(sm);
                                    counter++;
                                }
                            }
                        }
                    }
                    log.LogInformation("start generate excel---------");
                    MemoryStream ms = CreateExcelFile(smList, @"D:\local\Temp");
                    byte[] data = ms.ToArray();
                    return new FileContentResult(data, "application/octet-stream");
                }
                catch (Exception ex)
                {
                    log.LogError(ex.Message);
                    return new OkObjectResult(ex.Message);
                }
            }

            //return new OkObjectResult("Import Success");
        }
        static public MemoryStream CreateExcelFile(SummaryModelList data, string OutPutFileDirectory)
        {
            var datetime = DateTime.Now.ToString().Replace("/", "_").Replace(":", "_");

            string fileFullname = String.Empty;

            fileFullname = Path.Combine(OutPutFileDirectory, "Outputz_" + datetime + ".xlsx");

            using (SpreadsheetDocument package = SpreadsheetDocument.Create(fileFullname, SpreadsheetDocumentType.Workbook))
            {
                CreatePartsForExcel(package, data);
            }
            MemoryStream ms = new MemoryStream();
            using (FileStream file = new FileStream(fileFullname, FileMode.Open, FileAccess.Read))
                file.CopyTo(ms);
            File.Delete(fileFullname);
            return ms;
        }
        static public MemoryStream CreateExcelFile(TestModelList data, string OutPutFileDirectory)
        {
            var datetime = DateTime.Now.ToString().Replace("/", "_").Replace(":", "_");

            string fileFullname = String.Empty;

            fileFullname = Path.Combine(OutPutFileDirectory, "Output_" + datetime + ".xlsx");

            using (SpreadsheetDocument package = SpreadsheetDocument.Create(fileFullname, SpreadsheetDocumentType.Workbook))
            {
                CreatePartsForExcel(package, data);
            }
            MemoryStream ms = new MemoryStream();
            using (FileStream file = new FileStream(fileFullname, FileMode.Open, FileAccess.Read))
                file.CopyTo(ms);
            File.Delete(fileFullname);
            return ms;
        }
        static private void CreatePartsForExcel(SpreadsheetDocument document, TestModelList data)
        {
            SheetData partSheetData = GenerateSheetdataForDetails(data);

            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPartContent(workbookPart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateWorkbookStylesPartContent(workbookStylesPart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetPartContent(worksheetPart1, partSheetData);
        }

        static private void CreatePartsForExcel(SpreadsheetDocument document, SummaryModelList data)
        {
            SheetData partSheetData = GenerateSheetdataForDetails(data);

            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPartContent(workbookPart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateWorkbookStylesPartContent(workbookStylesPart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetPartContent(worksheetPart1, partSheetData);
        }

        static private void GenerateWorksheetPartContent(WorksheetPart worksheetPart1, SheetData sheetData1)
        {
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1" };

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection1 = new Selection() { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(pageMargins1);
            worksheetPart1.Worksheet = worksheet1;
        }

        static private void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)2U, KnownFonts = true };

            DocumentFormat.OpenXml.Spreadsheet.Font font1 = new DocumentFormat.OpenXml.Spreadsheet.Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            DocumentFormat.OpenXml.Spreadsheet.Color color1 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            DocumentFormat.OpenXml.Spreadsheet.Font font2 = new DocumentFormat.OpenXml.Spreadsheet.Font();
            Bold bold1 = new Bold();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            DocumentFormat.OpenXml.Spreadsheet.Color color2 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Theme = (UInt32Value)1U };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font2.Append(bold1);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontScheme2);

            fonts1.Append(font1);
            fonts1.Append(font2);

            Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);

            Borders borders1 = new Borders() { Count = (UInt32Value)2U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color3 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color3);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color4 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color4);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color5 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color5);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            DocumentFormat.OpenXml.Spreadsheet.Color color6 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color6);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            borders1.Append(border1);
            borders1.Append(border2);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)3U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true };

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.TimelineStyles timelineStyles1 = new X15.TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

            stylesheetExtension2.Append(timelineStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);
            stylesheetExtensionList1.Append(stylesheetExtension2);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart1.Stylesheet = stylesheet1;
        }

        static private void GenerateWorkbookPartContent(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook();
            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet() { Name = "Sheet1", SheetId = (UInt32Value)1U, Id = "rId1" };
            sheets1.Append(sheet1);
            workbook1.Append(sheets1);
            workbookPart1.Workbook = workbook1;
        }
        static private Row CreateHeaderRowForExcel()
        {
            Row workRow = new Row();
            workRow.Append(CreateCell("EmployeeId", 2U));
            workRow.Append(CreateCell("StaffNo", 2U));
            workRow.Append(CreateCell("DisplayName", 2U));
            workRow.Append(CreateCell("IsConflictOfInterest", 2U));
            workRow.Append(CreateCell("Signed Date", 2U));
            workRow.Append(CreateCell("Attachment", 2U));
            workRow.Append(CreateCell("HR Status", 2U));
            workRow.Append(CreateCell("Company", 2U));
            workRow.Append(CreateCell("Business Unit", 2U));
            workRow.Append(CreateCell("Division", 2U));
            workRow.Append(CreateCell("Department", 2U));
            workRow.Append(CreateCell("Section", 2U));
            workRow.Append(CreateCell("Last Updated Date", 2U));
            return workRow;
        }
        static private Row CreateHeaderRowForExcelS()
        {
            Row workRow = new Row();
            workRow.Append(CreateCell("BusinessUnit", 2U));
            workRow.Append(CreateCell("Division", 2U));
            workRow.Append(CreateCell("Company", 2U));
            workRow.Append(CreateCell("Department", 2U));
            workRow.Append(CreateCell("Signed", 2U));
            workRow.Append(CreateCell("NotSigned", 2U));
            workRow.Append(CreateCell("Reviewing", 2U));
            workRow.Append(CreateCell("Reviewed", 2U));
            workRow.Append(CreateCell("Total", 2U));

            return workRow;
        }
        static private EnumValue<CellValues> ResolveCellDataTypeOnValue(string text)
        {
            int intVal;
            double doubleVal;
            if (int.TryParse(text, out intVal) || double.TryParse(text, out doubleVal))
            {
                return CellValues.String;
            }
            else
            {
                return CellValues.String;
            }
        }
        static private Cell CreateCell(string text, uint styleIndex)
        {
            Cell cell = new Cell();
            cell.StyleIndex = styleIndex;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);
            
            return cell;
        }
        static private Cell CreateCell(string text)
        {
            Cell cell = new Cell();
            cell.StyleIndex = 1U;
            cell.DataType = ResolveCellDataTypeOnValue(text);
           cell.CellValue = new CellValue(text);
           
            return cell;
        }
        static private Row GenerateRowForChildPartDetail(TestModel testmodel)
        {
            Row tRow = new Row();
            tRow.Append(CreateCell(testmodel.EmployeeId));
            tRow.Append(CreateCell(testmodel.StaffNo));
            tRow.Append(CreateCell(testmodel.DisplayName));
            tRow.Append(CreateCell(testmodel.IsConflictOfInterest));
            tRow.Append(CreateCell(testmodel.signed_Date));
            tRow.Append(CreateCell(testmodel.OriginalFileName));
            tRow.Append(CreateCell(testmodel.HRStatus));
            tRow.Append(CreateCell(testmodel.Company));
            tRow.Append(CreateCell(testmodel.BusinessUnit));
            tRow.Append(CreateCell(testmodel.DivisionName));
            tRow.Append(CreateCell(testmodel.DepartmentName));
            tRow.Append(CreateCell(testmodel.SectionName));
            tRow.Append(CreateCell(testmodel.LastUpdateDate));


            return tRow;
        }

        static private Row GenerateRowForChildPartDetail(SummaryModel summarymodel)
        {
            Row tRow = new Row();
            tRow.Append(CreateCell(summarymodel.BusinessUnit));
            tRow.Append(CreateCell(summarymodel.Division));
            tRow.Append(CreateCell(summarymodel.Company));            
            tRow.Append(CreateCell(summarymodel.Department));
            tRow.Append(CreateCell(summarymodel.Signed));
            tRow.Append(CreateCell(summarymodel.NotSigned));
            tRow.Append(CreateCell(summarymodel.Reviewing));
            tRow.Append(CreateCell(summarymodel.Reviewed));
            tRow.Append(CreateCell(summarymodel.Total));
            return tRow;
        }
        static private SheetData GenerateSheetdataForDetails(TestModelList data)
        {
            SheetData sheetData1 = new SheetData();
            sheetData1.Append(CreateHeaderRowForExcel());

            foreach (TestModel testmodel in data.testData)
            {
                Row partsRows = GenerateRowForChildPartDetail(testmodel);
                sheetData1.Append(partsRows);
            }
            return sheetData1;
        }

        static private SheetData GenerateSheetdataForDetails(SummaryModelList data)
        {
            SheetData sheetData1 = new SheetData();
            sheetData1.Append(CreateHeaderRowForExcelS());

            foreach (SummaryModel testmodel in data.testData)
            {
                Row partsRows = GenerateRowForChildPartDetail(testmodel);
                sheetData1.Append(partsRows);
            }
            return sheetData1;
        }
    }
}
