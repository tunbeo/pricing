using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using System.Text;

namespace PricingService
{
    public static class ExcelCalculator
    {

        public static string Calculate(string query)
        {
            string rootPath = "C:\\inetpub\\wwwroot\\Pricing\\BangGia\\";

            string path = Directory.GetCurrentDirectory() + "\\Excels\\_pp1_sale.xlsx";  //"C:\\inetpub\\wwwroot\\Pricing\\BangGia\\2021-12-21 10-07\\pp1_sale.xlsx";
            //Application xlApp = null;// = new Application();
            SpreadsheetDocument document = null;
            Sheet sheet = null;
            Worksheet worksheet = null;
            Workbook workbook = null;
            SheetData sheetData = null;
            WorkbookPart workbookPart = null;
            Models.PriceResponse priceResponse = new Models.PriceResponse();
            bool pp1 = true, pp2 = false, pp3 = false, vietstar = false;
            int rowToFillData = 12;
            int rowToGetColumnName = rowToFillData - 1;


            int plus_i = 0; //this value for "hàng thanh" has insert column C, so other value must plus 1
            bool flag_hangcuon = false;
            bool error = false;

            string sheetName = "n/a";
            try
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(path, false))
                {
                    workbookPart = doc.WorkbookPart;
                    workbook = doc.WorkbookPart.Workbook;

                    //int worksheetcount = workbook.Sheets.Count();
                    //sheet = workbook.Sheets[3].;
                    sheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(3);

                    worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;

                    //sheetName = sheet.Name;

                    sheetData = worksheet.GetFirstChild<SheetData>();


                    //SheetData Rows = (SheetData)worksheet.ChildElements.GetItem(wkschildno);

                    //Row currentrow = (Row)Rows.ChildElements.GetItem(0);

                    //Cell currentcell = (Cell)currentrow.ChildElements.GetItem(0);

                    //IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();

                }

            }
            catch (Exception ex)
            {
                priceResponse.Message = ex.Message;
            }

            try

            {
                if (!string.IsNullOrEmpty(query))
                {
                    var parsed = System.Web.HttpUtility.ParseQueryString(query);
                    foreach (string key in parsed)
                    {
                        string LoaiVatLieu = "";
                        if (key == "date")
                        {
                            var value = parsed[key];
                            path = rootPath + value + "\\";
                        }
                        else if (key == "file")
                        {
                            var value = parsed[key];
                            path = path + value;

                            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(path, false))
                            {
                                workbook = doc.WorkbookPart.Workbook;

                                if (value.Contains("pp2"))
                                {
                                    sheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(1);
                                    worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;

                                    sheetName = sheet.Name;
                                    pp2 = true;
                                    rowToFillData = 4;

                                    pp1 = false;
                                    pp3 = false;
                                    vietstar = false;
                                }
                                else if (value.Contains("pp1"))
                                {
                                    pp1 = true;
                                    rowToFillData = 12;
                                    sheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(1);
                                    worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                                    sheetName = sheet.Name;

                                    pp3 = false;
                                    pp2 = false;
                                    vietstar = false;
                                }
                                else if (value.Contains("pp3"))
                                {
                                    pp3 = true;
                                    sheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(1);
                                    worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                                    sheetName = sheet.Name;

                                    pp1 = false;
                                    pp3 = false;
                                    vietstar = false;
                                }
                                else if (value.StartsWith("vietstar"))
                                {
                                    vietstar = true;
                                    rowToFillData = 2;
                                    sheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(1);
                                    worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                                    sheetName = sheet.Name;

                                    pp1 = false;
                                    pp2 = false;
                                    pp3 = false;
                                }
                            }


                        }

                        //only for vietstar
                        else if (key.ToLower() == "from")
                        {
                            var value = parsed[key];
                            if (value.ToLower() == "hn" || value.ToLower() == "hy")
                            {
                                rowToFillData = 2;
                            }
                            else //hcm
                            {
                                rowToFillData = 3;
                                rowToGetColumnName = 1;
                            }

                        }

                        // only in pp1
                        else if (key == "sheet")
                        {
                            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(path, false))
                            {
                                var value = parsed[key];
                                if (value == "thep_tam_day" || value == "thep_tam_la")
                                {
                                    sheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(3);
                                    worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                                    if (value == "thep_tam_day")
                                        LoaiVatLieu = "Thép tấm dày";
                                    else
                                        LoaiVatLieu = "Thép tấm lá";
                                }

                                else if (value == "dong_bery" || value == "thep_dung_cu")
                                {
                                    sheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(4);
                                    worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                                    if (value == "dong_bery")
                                        LoaiVatLieu = "Đồng bery";
                                    else
                                        LoaiVatLieu = "Thép dụng cụ";
                                }

                                else if (value == "dong_tam_la" || value == "dong_hop_kim_tam_day" || value == "dong_tam_day" || value == "nhom_hop_kim_day" || value == "nhom_hop_kim_mong")
                                {
                                    sheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(5);
                                    worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                                    if (value == "dong_tam_la")
                                        LoaiVatLieu = "Đồng tấm lá";
                                    else if (value == "dong_hop_kim_tam_day")
                                        LoaiVatLieu = "Đồng hợp kim tấm dày ";
                                    else if (value == "dong_tam_day")
                                        LoaiVatLieu = "Đồng tấm dày";
                                    else if (value == "nhom_hop_kim_day")
                                        LoaiVatLieu = "Nhôm hợp kim dày";
                                    else
                                        LoaiVatLieu = "Nhôm hợp kim mỏng";
                                }
                                else if (value == "dong_thanh" || value == "dong_hop_kim_thanh" || value == "nhom_thanh" || value == "thep_thanh" || value == "dong_bery_thanh")
                                {
                                    sheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(6);
                                    worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                                    plus_i = 1;
                                    if (value == "dong_thanh")
                                        LoaiVatLieu = "Đồng thanh";
                                    else if (value == "dong_hop_kim_thanh")
                                        LoaiVatLieu = "Đồng hợp kim thanh";
                                    else if (value == "nhom_thanh")
                                        LoaiVatLieu = "Nhôm thanh";
                                    else if (value == "thep_thanh")
                                        LoaiVatLieu = "Thép thanh";
                                    else
                                        LoaiVatLieu = "Đồng bery thanh";
                                }
                                else if (value == "nhom_khong_hop_kim_cuon" || value == "nhom_hop_kim_cuon" || value == "dong_tinh_che_tam_la" || value == "dong_hop_kim_tam_la" || value == "thep_la_cuon")
                                {
                                    sheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(7);
                                    worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                                    flag_hangcuon = true;
                                    if (value == "nhom_khong_hop_kim_cuon")
                                        LoaiVatLieu = "Nhôm không hợp kim cuộn";
                                    else if (value == "nhom_hop_kim_cuon")
                                        LoaiVatLieu = "Nhôm hợp kim cuộn";
                                    else if (value == "dong_tinh_che_tam_la")
                                        LoaiVatLieu = "Đồng tinh chế tấm lá";
                                    else if (value == "dong_hop_kim_tam_la")
                                        LoaiVatLieu = "Đồng hợp kim tấm lá";
                                    else
                                        LoaiVatLieu = "Thép lá cuộn";
                                }
                                else
                                {
                                    error = true;
                                }
                                sheetName = sheet.Name;

                                //set b column to material type

                                sheetData = worksheet.GetFirstChild<SheetData>();

                                Row currentrow = (Row)sheetData.ChildElements.GetItem(12);

                                Cell currentcell = (Cell)currentrow.Elements<Cell>().Where(c => c.CellReference.Value == "B64");
                                currentcell.CellValue = new CellValue(LoaiVatLieu);

                                //worksheet.Cells[12, char.ToUpper(char.Parse("b")) - 64] = LoaiVatLieu;
                            }
                        }

                        else
                        {
                            var stt = char.ToUpper(char.Parse(key)) - 64;

                            var value = parsed[key];

                            Row currentrow = (Row)sheetData.ChildElements.GetItem(rowToFillData);

                            Cell currentcell = (Cell)currentrow.ChildElements.GetItem(stt);
                            currentcell.CellValue = new CellValue(LoaiVatLieu);


                            if (pp1)
                            {
                                if (key != "c")
                                {
                                    stt = stt + plus_i;
                                }
                                currentcell = (Cell)currentrow.ChildElements.GetItem(stt);
                                if (value == "circle")
                                {
                                    currentcell.CellValue = new CellValue("Tròn");
                                }
                                //wSheet.Cells[rowToFillData, stt] = "Tròn";
                                else if (value == "rectangle")
                                    currentcell.CellValue = new CellValue("Chữ nhật");
                                //wSheet.Cells[rowToFillData, stt] = "Chữ nhật";
                                else
                                    currentcell.CellValue = new CellValue(value);
                                //wSheet.Cells[rowToFillData, stt] = value;
                            }
                            if (pp2)
                            {
                                //wSheet.Cells[rowToFillData, stt] = value;
                                currentcell.CellValue = new CellValue(value);
                            }
                            if (vietstar)
                            {
                                //wSheet.Cells[rowToFillData, stt] = value;
                                currentcell.CellValue = new CellValue(value);
                            }


                            currentrow = (Row)sheetData.ChildElements.GetItem(rowToGetColumnName);

                            currentcell = (Cell)currentrow.ChildElements.GetItem(stt);

                            var tencolumn = currentcell.CellValue.ToString(); //wSheet.Cells[rowToGetColumnName, stt]).Value2.ToString();
                            priceResponse.Input += tencolumn.Replace("\n", "") + ": " + value + "; ";

                            //http://localhost:7250/Pricing?date=2022-01-20&file=vietstar.xlsx&from=hn&b=a&c=5&d=1
                        }
                    }
                }





                if (!error)
                {
                    priceResponse.Message = "OK";                  


                    priceResponse.SheetName = sheetName;

                    if (pp1)
                    {
                        if (!flag_hangcuon)
                        {
                            //var cellValue = GetCellValue(GetCell(sheetData, "B2"), wbPart);

                            priceResponse.LoaiVatLieu = GetCellValue(GetCell(sheetData, "" + 2), workbookPart); // ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 2]).Value2.ToString();
                            priceResponse.TheTich = GetCellValue(GetCell(sheetData, "" + 13 + plus_i), workbookPart); //((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 13 + plus_i]).Value2.ToString();
                            priceResponse.SoKg = GetCellValue(GetCell(sheetData, "" + 14 + plus_i), workbookPart); //((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 14 + plus_i]).Value2.ToString();
                            priceResponse.HaoPhiMachCat = GetCellValue(GetCell(sheetData, "" + 15 + plus_i), workbookPart);// ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 15 + plus_i]).Value2.ToString();
                            priceResponse.PhiGiaCong = GetCellValue(GetCell(sheetData, "" + 17 + plus_i), workbookPart);// ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 17 + plus_i]).Value2.ToString();
                            priceResponse.DonGia = GetCellValue(GetCell(sheetData, "" + 18 + plus_i), workbookPart);// ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 18 + plus_i]).Value2.ToString();
                            priceResponse.DonGiaTheoTSLN = GetCellValue(GetCell(sheetData, "" + 20 + plus_i), workbookPart); // ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 20 + plus_i]).Value2.ToString();
                            priceResponse.ThanhTienTheoTSLN = GetCellValue(GetCell(sheetData, "" + 21 + plus_i), workbookPart); // ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 21 + plus_i]).Value2.ToString();

                        }
                        //else // hang cuon
                        //{
                        //    priceResponse.LoaiVatLieu = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 2]).Value2.ToString();

                        //    priceResponse.PhiGiaCong = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 12]).Value2.ToString();
                        //    priceResponse.DonGia = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 13]).Value2.ToString();
                        //    priceResponse.DonGiaTheoTSLN = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 15]).Value2.ToString();
                        //    priceResponse.ThanhTienTheoTSLN = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 16]).Value2.ToString();

                        //}
                    }

                    //if (pp2)
                    //{
                    //    priceResponse.FramePrice = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[11, 11]).Value2.ToString();
                    //}   

                    //if (pp3)
                    //{

                    //}

                    //if (vietstar)
                    //{
                    //    priceResponse.ChiPhiVanChuyenThiXa = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[rowToFillData, 5]).Value2.ToString();
                    //    priceResponse.ChiPhiVanChuyenNgoaiThanh = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[rowToFillData, 6]).Value2.ToString();
                    //}




                    //var x = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[11, 4]).Value2.ToString();

                    //var a1 = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[11, 20]).Value2.ToString();
                    //var x1 = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 20]).Value2.ToString();




                }

                else
                {
                    priceResponse.Message = "Error: Sai tên vật liệu";

                }
                string json = JsonConvert.SerializeObject(priceResponse, Formatting.Indented, new JsonSerializerSettings
                {
                    NullValueHandling = NullValueHandling.Ignore
                });

                

                return json;

            }

            catch (Exception ex)
            {
                
                priceResponse.Message = ex.Message;

                string json = JsonConvert.SerializeObject(priceResponse, Formatting.Indented);
                //if (wb != null)
                //    wb.Close(false);
                //if (xlApp != null)
                //    xlApp.Quit();
                //xlApp = null;

                //Marshal.ReleaseComObject(wb);
                //Marshal.ReleaseComObject(wbs);
                //Marshal.ReleaseComObject(wSheet);
                

                return json;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetData"></param>
        /// <param name="cellAddress"></param>
        /// <returns></returns>

        public static Cell GetCell(SheetData sheetData, string cellAddress)
        {
            uint rowIndex = uint.Parse(Regex.Match(cellAddress, @"[0-9]+").Value);
            return sheetData.Descendants<Row>().FirstOrDefault(p => p.RowIndex == rowIndex).Descendants<Cell>().FirstOrDefault(p => p.CellReference == cellAddress);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="wbPart"></param>
        /// <returns></returns>
        public static string GetCellValue(Cell cell, WorkbookPart wbPart)
        {
            string value = cell.InnerText;
            if (cell.DataType != null)
            {
                switch (cell.DataType.Value)
                {
                    case CellValues.SharedString:
                        var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                        if (stringTable != null)
                        {
                            value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                        }
                        break;

                    case CellValues.Boolean:
                        switch (value)
                        {
                            case "0":
                                value = "FALSE";
                                break;
                            default:
                                value = "TRUE";
                                break;
                        }
                        break;
                }
            }
            return value;
        }


        public static string CalculatePrice(string request)
        {
            var returnValue = new Models.PriceResponse();
            returnValue.Message = "OK";
            
            string folder = "", fileName = "", filePath = "", rootFolder = "";
            int rowData = 12;
            var parsed = System.Web.HttpUtility.ParseQueryString(request);
            folder = parsed["date"];
            //folder = "Excels";
            rootFolder = "C:\\inetpub\\wwwroot\\pricing\\banggia\\";
            //rootFolder = "D:\\Z\\Excel\\pricing\\Excels\\";
            fileName = parsed["file"];
            if (!string.IsNullOrEmpty(folder) && !string.IsNullOrEmpty(fileName))
            {
                filePath = Path.Combine(rootFolder + folder, fileName);
                //filePath = Path.Combine(rootFolder, fileName);
            }
            Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();
            string A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U;
            try
            {
                if (System.IO.File.Exists(filePath))
                {
                    System.Text.CodePagesEncodingProvider.Instance.GetEncoding(437);
                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                    workbook.Open(filePath);
                    Aspose.Cells.Worksheet worksheet = null;
                    // Phuong phap 1
                    if (fileName.Contains("pp1"))
                    {
                        string sheetName = Models.ApiParamenter.GetApiParamenterSheet(parsed["sheet"]);
                        foreach (Aspose.Cells.Worksheet item in workbook.Worksheets)
                        {
                            // Tim vao sheet
                            if (item.Name.ToLower().Contains(sheetName.ToLower()))
                            {
                                if (sheetName.ToLower().Contains("hàng tấm"))
                                {
                                    for (int _i = 11; _i < item.Cells.Rows.Count; _i++)
                                    {
                                        A = "A" + _i; B = "B" + _i; C = "C" + _i; D = "D" + _i; E = "E" + _i; F = "F" + _i; G = "G" + _i;
                                        H = "H" + _i; I = "I" + _i; J = "J" + _i; K = "K" + _i; L = "L" + _i; M = "M" + _i; N = "N" + _i;
                                        O = "O" + _i; P = "P" + _i; Q = "Q" + _i; R = "R" + _i; S = "S" + _i; T = "T" + _i; U = "U" + _i;
                                        if (string.IsNullOrEmpty(item.Cells[B]?.Value?.ToString()))
                                        {
                                            rowData = 12;
                                            var s = parsed["sheet"];
                                            var w = Models.ApiParamenter.GetApiParamenterMaterial(s);
                                            item.Cells[B].PutValue(w);
                                            var d12 = parsed["d"];
                                            item.Cells[D].PutValue(Convert.ToDouble(parsed["d"]));
                                            item.Cells[E].PutValue(Convert.ToDouble(parsed["e"]));
                                            item.Cells[F].PutValue(Convert.ToDouble(parsed["f"]));
                                            item.Cells[G].PutValue(Convert.ToDouble(parsed["g"]));
                                            item.Cells[H].PutValue(Convert.ToDouble(parsed["h"]));
                                            item.Cells[I].PutValue(parsed["i"]);
                                            item.Cells[J].PutValue(parsed["j"]);
                                            item.Cells[K].PutValue(parsed["k"]);
                                            workbook.CalculateFormula();
                                            break;
                                        }
                                    }
                                    //workbook.Save(filePath);
                                    returnValue.Message = "OK";
                                    returnValue.Input = $"Dày (mm): {parsed["e"]}; Rộng (mm): {parsed["f"]}; Dài (mm): {parsed["g"]}; Số pcs: {parsed["h"]}; Giá vốn nguyên tấm: {parsed["d"]}; TSLN: {parsed["s"]};";
                                    returnValue.LoaiVatLieu = Models.ApiParamenter.GetApiParamenterMaterial(parsed["sheet"]);
                                    returnValue.SheetName = sheetName;
                                    returnValue.TheTich = item.Cells["M" + rowData].Value?.ToString();
                                    returnValue.SoKg = item.Cells["N" + rowData].Value?.ToString();
                                    returnValue.HaoPhiMachCat = item.Cells["O" + rowData].Value?.ToString();
                                    returnValue.PhiGiaCong = item.Cells["Q" + rowData].Value?.ToString();
                                    returnValue.DonGia = item.Cells["R" + rowData].Value?.ToString();
                                    returnValue.DonGiaTheoTSLN = item.Cells["T" + rowData].Value?.ToString();
                                    returnValue.ThanhTienTheoTSLN = item.Cells["U" + rowData].Value?.ToString();
                                }
                                else if (sheetName.ToLower().Contains("hàng thanh"))
                                {
                                    for (int _i = 11; _i < item.Cells.Rows.Count; _i++)
                                    {
                                        A = "A" + _i; B = "B" + _i; C = "C" + _i; D = "D" + _i; E = "E" + _i; F = "F" + _i; G = "G" + _i;
                                        H = "H" + _i; I = "I" + _i; J = "J" + _i; K = "K" + _i; L = "L" + _i; M = "M" + _i; N = "N" + _i;
                                        O = "O" + _i; P = "P" + _i; Q = "Q" + _i; R = "R" + _i; S = "S" + _i; T = "T" + _i; U = "U" + _i;
                                        if (string.IsNullOrEmpty(item.Cells[B]?.Value?.ToString()))
                                        {
                                            rowData = 12;
                                            var s = parsed["sheet"];
                                            var w = Models.ApiParamenter.GetApiParamenterMaterial(s);
                                            item.Cells[B].PutValue(w);
                                            if (parsed["c"] == "circle")
                                            {
                                                item.Cells[C].PutValue("Tròn");
                                            }
                                            else if (parsed["c"] == "rectangle")
                                            {
                                                item.Cells[C].PutValue("Chữ nhật");
                                            }
                                            else
                                            {
                                                item.Cells[C].PutValue(parsed["c"]);
                                            }
                                            item.Cells[E].PutValue(Convert.ToDouble(parsed["d"]));
                                            item.Cells[F].PutValue(Convert.ToDouble(parsed["e"]));
                                            item.Cells[G].PutValue(Convert.ToDouble(parsed["f"]));
                                            item.Cells[H].PutValue(Convert.ToDouble(parsed["g"]));
                                            item.Cells[I].PutValue(Convert.ToDouble(parsed["h"]));
                                            item.Cells[J].PutValue(parsed["i"]);
                                            item.Cells[K].PutValue(parsed["j"]);
                                            item.Cells[L].PutValue(parsed["k"]);
                                            workbook.CalculateFormula();
                                            break;
                                        }
                                    }

                                    //workbook.Save(filePath);
                                    returnValue.Message = "OK";
                                    returnValue.Input = $"Dày (mm): {parsed["e"]}; Rộng (mm): {parsed["f"]}; Dài (mm): {parsed["g"]}; Số pcs: {parsed["h"]}; Giá vốn nguyên tấm: {parsed["d"]}; TSLN: {parsed["s"]};";
                                    returnValue.LoaiVatLieu = Models.ApiParamenter.GetApiParamenterMaterial(parsed["sheet"]);
                                    returnValue.SheetName = sheetName;
                                    returnValue.SoKg = item.Cells["O" + rowData].Value.ToString();
                                    returnValue.HaoPhiMachCat = item.Cells["P" + rowData].Value.ToString();
                                    returnValue.PhiGiaCong = item.Cells["R" + rowData].Value.ToString();
                                    returnValue.DonGia = item.Cells["U" + rowData].Value.ToString();
                                    returnValue.DonGiaTheoTSLN = item.Cells["T" + rowData].Value.ToString();
                                    returnValue.ThanhTienTheoTSLN = item.Cells["V" + rowData].Value.ToString();
                                }
                                else if (sheetName.ToLower().Contains("hàng cuộn"))
                                {
                                    for (int _i = 11; _i < item.Cells.Rows.Count; _i++)
                                    {
                                        A = "A" + _i; B = "B" + _i; C = "C" + _i; D = "D" + _i; E = "E" + _i; F = "F" + _i; G = "G" + _i;
                                        H = "H" + _i; I = "I" + _i; J = "J" + _i; K = "K" + _i; L = "L" + _i; M = "M" + _i; N = "N" + _i;
                                        O = "O" + _i; P = "P" + _i; Q = "Q" + _i; R = "R" + _i; S = "S" + _i; T = "T" + _i; U = "U" + _i;
                                        if (string.IsNullOrEmpty(item.Cells[B]?.Value?.ToString()))
                                        {
                                            rowData = 12;
                                            var s = parsed["sheet"];
                                            var w = Models.ApiParamenter.GetApiParamenterMaterial(s);
                                            item.Cells[B].PutValue(w);
                                            item.Cells[C].PutValue(Convert.ToDouble(parsed["c"]));
                                            item.Cells[D].PutValue(Convert.ToDouble(parsed["d"]));
                                            item.Cells[E].PutValue(Convert.ToDouble(parsed["e"]));
                                            item.Cells[F].PutValue(Convert.ToDouble(parsed["f"]));
                                            item.Cells[G].PutValue(Convert.ToDouble(parsed["g"]));
                                            item.Cells[H].PutValue(parsed["h"]);
                                            item.Cells[I].PutValue(parsed["i"]);
                                            workbook.CalculateFormula();
                                            break;
                                        }
                                    }
                                    //workbook.Save(filePath);
                                    returnValue.Message = "OK";
                                    returnValue.Input = $"Dày (mm): {parsed["d"]}; Rộng (mm): {parsed["e"]}; Dài (mm): ; Số pcs: {parsed["f"]}; Giá vốn nguyên tấm: {parsed["c"]}; TSLN: {parsed["n"]};";
                                    returnValue.LoaiVatLieu = Models.ApiParamenter.GetApiParamenterMaterial(parsed["sheet"]);
                                    returnValue.SheetName = sheetName;
                                    returnValue.SoKg = item.Cells["F" + rowData]?.Value?.ToString();
                                    returnValue.PhiGiaCong = item.Cells["L" + rowData]?.Value?.ToString();
                                    returnValue.DonGia = item.Cells["O" + rowData]?.Value?.ToString();
                                    returnValue.ThanhTienTheoTSLN = item.Cells["P" + rowData]?.Value?.ToString();
                                }
                                break;
                            }
                        }
                    }
                    // _pp2
                    else if (fileName.Contains("pp2"))
                    {
                        worksheet = workbook.Worksheets[0];
                        for (int _i = 6; _i < worksheet.Cells.Rows.Count; _i++)
                        {
                            A = "A" + _i; B = "B" + _i; C = "C" + _i; D = "D" + _i; E = "E" + _i; F = "F" + _i; G = "G" + _i;
                            H = "H" + _i; I = "I" + _i; J = "J" + _i; K = "K" + _i; L = "L" + _i; M = "M" + _i; N = "N" + _i;
                            O = "O" + _i; P = "P" + _i; Q = "Q" + _i; R = "R" + _i; S = "S" + _i; T = "T" + _i; U = "U" + _i;
                            if (worksheet.Cells[A]?.Value?.ToString() == parsed["k"])
                            {
                                worksheet.Cells["K4"].PutValue(parsed["k"]);
                                worksheet.Cells["L4"].PutValue(parsed["o"]);
                                worksheet.Cells["M4"].PutValue(worksheet.Cells[C]?.Value);
                                worksheet.Cells["N4"].PutValue(worksheet.Cells[E]?.Value);
                                worksheet.Cells["O4"].PutValue(parsed["l"]);
                                workbook.CalculateFormula();
                                break;
                            }
                        }
                        //workbook.Save(filePath);
                        returnValue.Message = "OK";
                        returnValue.Input = $"Product: {parsed["k"]}; Square: {parsed["o"]}; Pcs: {parsed["l"]};";
                        returnValue.SheetName = worksheet.Name;
                        returnValue.FramePrice = worksheet.Cells["K11"]?.Value?.ToString();
                    }
                    // _pp3
                    else if (fileName.Contains("pp3"))
                    {
                        foreach (Aspose.Cells.Worksheet ws in workbook.Worksheets)
                        {
                            if (ws.Name.ToLower() == parsed["sheet"]?.ToLower())
                            {
                                ws.Cells["O12"].PutValue(Convert.ToDouble(parsed["l"]));
                                ws.Cells["O13"].PutValue(Convert.ToDouble(parsed["m"]));
                                ws.Cells["O14"].PutValue(Convert.ToDouble(parsed["n"]));
                                ws.Cells["O15"].PutValue(Convert.ToDouble(parsed["o"]));
                                ws.Cells["O16"].PutValue(Convert.ToDouble(parsed["p"]));
                                ws.Cells["O17"].PutValue(Convert.ToDouble(parsed["q"]));
                                ws.Cells["O18"].PutValue(Convert.ToDouble(parsed["r"]));
                                ws.Cells["O19"].PutValue(Convert.ToDouble(parsed["s"]));
                                // Ti gia mua vcb
                                ws.Cells["O24"].PutValue(Convert.ToDouble(parsed["l"]));
                                ws.Cells["O25"].PutValue(Convert.ToDouble(parsed["r"]));
                                workbook.CalculateFormula();
                                returnValue.DonGiaTheoTyGiaMuaVietCombank = ws.Cells["O27"]?.Value.ToString();
                                // Ti gia ban vcb
                                ws.Cells["O24"].PutValue(Convert.ToDouble(parsed["l"]));
                                ws.Cells["O25"].PutValue(Convert.ToDouble(parsed["s"]));
                                workbook.CalculateFormula();
                                returnValue.DonGiaTheoTyGiaBanVietCombank = ws.Cells["O27"]?.Value.ToString();
                            }
                        }
                        //workbook.Save(filePath);
                        returnValue.Message = "OK";
                        returnValue.Input = $"LME Thoi diem: {parsed["l"]}; LME Trung binh thang: {parsed["m"]}; LME Trung binh tuan: {parsed["n"]}; SMM Thoi diem: {parsed["o"]}; SMM Trung binh thang: {parsed["p"]}; SMM Trung binh tuan: {parsed["q"]}; Ty gia mua VCB: {parsed["r"]}; Ty gia ban VCB: {parsed["s"]};";
                        returnValue.SheetName = parsed["sheet"];
                    }
                    // _vietstar
                    else if (fileName.Contains("vietstar"))
                    {
                        worksheet = workbook.Worksheets[0];
                        for (int _i = 23; ; _i++)
                        {
                            A = "A" + _i; B = "B" + _i; C = "C" + _i; D = "D" + _i; E = "E" + _i; F = "F" + _i; G = "G" + _i;
                            H = "H" + _i; I = "I" + _i; J = "J" + _i; K = "K" + _i; L = "L" + _i; M = "M" + _i; N = "N" + _i;
                            O = "O" + _i; P = "P" + _i; Q = "Q" + _i; R = "R" + _i; S = "S" + _i; T = "T" + _i; U = "U" + _i;
                            if (string.IsNullOrEmpty(worksheet.Cells[A]?.Value?.ToString()))
                            {
                                rowData = 24;
                                var s = parsed["from"];
                                var w = Models.ApiParamenter.GetApiParamenterLocation(s);
                                var formula = "";
                                var sumFor = $"=IF(D{_i},120%,100%)*E{_i}";
                                if (w == "Formular HCM")
                                {
                                    formula = $"=INDEX($B$14:$J$20,MATCH(IF(C{_i}>2,2,C{_i}),$A$14:$A$20,1),MATCH(B{_i},$B$13:$J$13,1))+IF(C{_i}>2,ROUNDUP((C{_i}-2)/0.5,0)*HLOOKUP(B{_i},$B$13:$J$21,9,FALSE),0)";
                                }
                                else
                                {
                                    formula = $"=INDEX($B$3:$J$9,MATCH(IF(C{_i}>2,2,C{_i}),$A$3:$A$9,1),MATCH(B{_i},$B$2:$J$2,1))+IF(C{_i}>2,ROUNDUP((C{_i}-2)/0.5,0)*HLOOKUP(B{_i},$B$2:$J$10,9,FALSE),0)";
                                }
                                worksheet.Cells[A].PutValue(w);
                                worksheet.Cells[B].PutValue(parsed["b"]);
                                worksheet.Cells[C].PutValue(Convert.ToInt32(parsed["c"]));
                                worksheet.Cells[D].PutValue(Convert.ToInt32(parsed["d"]?.ToString()));
                                worksheet.Cells[E].Formula = formula;
                                workbook.CalculateFormula();
                                worksheet.Cells[F].Formula = sumFor;
                                workbook.CalculateFormula();
                                break;
                            }
                        }
                        //workbook.Save(filePath);
                        returnValue.Message = "OK";
                        returnValue.Input = $"Mã Vùng: {parsed["b"]}; KG: {parsed["c"]}; Trả hàng ngoại thành: {parsed["d"]};";
                        returnValue.SheetName = worksheet.Name;
                        returnValue.ChiPhiVanChuyenThiXa = worksheet.Cells["E" + rowData]?.Value?.ToString();
                        returnValue.ChiPhiVanChuyenNgoaiThanh = worksheet.Cells["F" + rowData]?.Value?.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                returnValue.Message = ex.Message;
            }
            

            string json = JsonConvert.SerializeObject(returnValue, Formatting.None, new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore
            });


            //json = json.Replace("\"{", "{").Replace("}\"", "}");
            //json = json.Replace("\\","");

            //this.Context.Response.ContentType = "application/json; charset=utf-8";
            //string jsonString = json.Replace(@"\", " ");
            json = json.Replace(@"\""", @"""");
            return json;
        }

        private static string SanitizeReceivedJson(string uglyJson)
        {
            var sb = new StringBuilder(uglyJson);
            sb.Replace("\\\t", "\t");
            sb.Replace("\\\n", "\n");
            sb.Replace("\\\r", "\r");
            sb.Replace("\\", "");
            sb.Replace("\"{", "{").Replace("}\"", "}");
            return sb.ToString();
        }

    }
}

