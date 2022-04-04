﻿using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
//using Microsoft.Office.Interop.Excel;
//using System.Runtime.InteropServices;

namespace PricingService
{
    public static class ExcelCalculator
    {
#pragma warning disable CS8600
#pragma warning disable CS8602
#pragma warning disable CS8604
#pragma warning disable CA1416
        //public static string Calculate(string query)
        //{
        //    string rootPath = "C:\\inetpub\\wwwroot\\Pricing\\BangGia\\";
        //    string path = "C:\\inetpub\\wwwroot\\Pricing\\BangGia\\2021-12-21 10-07\\pp1_sale.xlsx";
        //    Application xlApp = null;// = new Application();
        //    Microsoft.Office.Interop.Excel.Workbook wb = null;// = xlApp.Workbooks.Open(path);
        //    Worksheet wSheet = null;// = (Worksheet)wb.Worksheets[3];
        //    Workbooks wbs = null;
        //    Models.PriceResponse priceResponse = new Models.PriceResponse();
        //    bool pp1 = true, pp2 = false, pp3 = false, vietstar = false;
        //    int rowToFillData = 12;
        //    int rowToGetColumnName = rowToFillData - 1;


        //    int plus_i = 0; //this value for "hàng thanh" has insert column C, so other value must plus 1
        //    bool flag_hangcuon = false;
        //    bool error = false;

        //    string sheetName = "n/a";
        //    try
        //    {
        //        xlApp = new Application();
        //        wbs = xlApp.Workbooks;
        //        wb = wbs.Open(path);
        //        wSheet = (Worksheet)wb.Worksheets[3];

        //    }
        //    catch (Exception ex)
        //    {
        //        priceResponse.Message = ex.Message;
        //    }



        //    try

        //    {
        //        if (!string.IsNullOrEmpty(query))
        //        {
        //            var parsed = System.Web.HttpUtility.ParseQueryString(query);
        //            foreach (string key in parsed)
        //            {
        //                string LoaiVatLieu = "";
        //                if (key == "date")
        //                {
        //                    var value = parsed[key];
        //                    path = rootPath + value + "\\";
        //                }
        //                else if (key == "file")
        //                {
        //                    var value = parsed[key];
        //                    path = path + value;

        //                    //xlApp = new Application();
        //                    //wbs = xlApp.Workbooks;
        //                    wb = wbs.Open(path);
        //                    if (value.Contains("pp2"))
        //                    {
        //                        wSheet = (Worksheet)wb.Worksheets[1];
        //                        sheetName = wSheet.Name;
        //                        pp2 = true;
        //                        rowToFillData = 4;

        //                        pp1 = false;
        //                        pp3 = false;
        //                        vietstar = false;
        //                    }
        //                    else if (value.Contains("pp1"))
        //                    {
        //                        pp1 = true;
        //                        rowToFillData = 12;
        //                        wSheet = (Worksheet)wb.Worksheets[1];
        //                        sheetName = wSheet.Name;

        //                        pp3 = false;
        //                        pp2 = false;
        //                        vietstar = false;
        //                    }
        //                    else if (value.Contains("pp3"))
        //                    {
        //                        pp3 = true;
        //                        wSheet = (Worksheet)wb.Worksheets[1];
        //                        sheetName = wSheet.Name;

        //                        pp1 = false;
        //                        pp3 = false;
        //                        vietstar = false;
        //                    }
        //                    else if (value.StartsWith("vietstar"))
        //                    {
        //                        vietstar = true;
        //                        rowToFillData = 2;
        //                        wSheet = (Worksheet)wb.Worksheets[1];
        //                        sheetName = wSheet.Name;

        //                        pp1 = false;
        //                        pp2 = false;
        //                        pp3 = false;
        //                    }
        //                }

        //                //only for vietstar
        //                else if (key.ToLower() == "from")
        //                {
        //                    var value = parsed[key];
        //                    if (value.ToLower() == "hn" || value.ToLower() == "hy")
        //                    {
        //                        rowToFillData = 2;
        //                    }
        //                    else //hcm
        //                    {
        //                        rowToFillData = 3;
        //                        rowToGetColumnName = 1;
        //                    }

        //                }

        //                // only in pp1
        //                else if (key == "sheet")
        //                {
        //                    var value = parsed[key];
        //                    if (value == "thep_tam_day" || value == "thep_tam_la")
        //                    {
        //                        wSheet = (Worksheet)wb.Worksheets[3];
        //                        if (value == "thep_tam_day")
        //                            LoaiVatLieu = "Thép tấm dày";
        //                        else
        //                            LoaiVatLieu = "Thép tấm lá";
        //                    }

        //                    else if (value == "dong_bery" || value == "thep_dung_cu")
        //                    {
        //                        wSheet = (Worksheet)wb.Worksheets[4];
        //                        if (value == "dong_bery")
        //                            LoaiVatLieu = "Đồng bery";
        //                        else
        //                            LoaiVatLieu = "Thép dụng cụ";
        //                    }

        //                    else if (value == "dong_tam_la" || value == "dong_hop_kim_tam_day" || value == "dong_tam_day" || value == "nhom_hop_kim_day" || value == "nhom_hop_kim_mong")
        //                    {
        //                        wSheet = (Worksheet)wb.Worksheets[5];
        //                        if (value == "dong_tam_la")
        //                            LoaiVatLieu = "Đồng tấm lá";
        //                        else if (value == "dong_hop_kim_tam_day")
        //                            LoaiVatLieu = "Đồng hợp kim tấm dày ";
        //                        else if (value == "dong_tam_day")
        //                            LoaiVatLieu = "Đồng tấm dày";
        //                        else if (value == "nhom_hop_kim_day")
        //                            LoaiVatLieu = "Nhôm hợp kim dày";
        //                        else
        //                            LoaiVatLieu = "Nhôm hợp kim mỏng";
        //                    }
        //                    else if (value == "dong_thanh" || value == "dong_hop_kim_thanh" || value == "nhom_thanh" || value == "thep_thanh" || value == "dong_bery_thanh")
        //                    {
        //                        wSheet = (Worksheet)wb.Worksheets[6];
        //                        plus_i = 1;
        //                        if (value == "dong_thanh")
        //                            LoaiVatLieu = "Đồng thanh";
        //                        else if (value == "dong_hop_kim_thanh")
        //                            LoaiVatLieu = "Đồng hợp kim thanh";
        //                        else if (value == "nhom_thanh")
        //                            LoaiVatLieu = "Nhôm thanh";
        //                        else if (value == "thep_thanh")
        //                            LoaiVatLieu = "Thép thanh";
        //                        else
        //                            LoaiVatLieu = "Đồng bery thanh";
        //                    }
        //                    else if (value == "nhom_khong_hop_kim_cuon" || value == "nhom_hop_kim_cuon" || value == "dong_tinh_che_tam_la" || value == "dong_hop_kim_tam_la" || value == "thep_la_cuon")
        //                    {
        //                        wSheet = (Worksheet)wb.Worksheets[7];
        //                        flag_hangcuon = true;
        //                        if (value == "nhom_khong_hop_kim_cuon")
        //                            LoaiVatLieu = "Nhôm không hợp kim cuộn";
        //                        else if (value == "nhom_hop_kim_cuon")
        //                            LoaiVatLieu = "Nhôm hợp kim cuộn";
        //                        else if (value == "dong_tinh_che_tam_la")
        //                            LoaiVatLieu = "Đồng tinh chế tấm lá";
        //                        else if (value == "dong_hop_kim_tam_la")
        //                            LoaiVatLieu = "Đồng hợp kim tấm lá";
        //                        else
        //                            LoaiVatLieu = "Thép lá cuộn";
        //                    }
        //                    else
        //                    {
        //                        error = true;
        //                    }
        //                    sheetName = wSheet.Name;

        //                    //set b column to material type
        //                    wSheet.Cells[12, char.ToUpper(char.Parse("b")) - 64] = LoaiVatLieu;
        //                }

        //                else
        //                {
        //                    var stt = char.ToUpper(char.Parse(key)) - 64;

        //                    var value = parsed[key];

        //                    if (pp1)
        //                    {
        //                        if (key != "c")
        //                        {
        //                            stt = stt + plus_i;
        //                        }

        //                        if (value == "circle")
        //                            wSheet.Cells[rowToFillData, stt] = "Tròn";
        //                        else if (value == "rectangle")
        //                            wSheet.Cells[rowToFillData, stt] = "Chữ nhật";
        //                        else
        //                            wSheet.Cells[rowToFillData, stt] = value;
        //                    }
        //                    if (pp2)
        //                    {
        //                        wSheet.Cells[rowToFillData, stt] = value;
        //                    }
        //                    if (vietstar)
        //                    {
        //                        wSheet.Cells[rowToFillData, stt] = value;
        //                    }



        //                    var tencolumn = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[rowToGetColumnName, stt]).Value2.ToString();
        //                    priceResponse.Input += tencolumn.Replace("\n", "") + ": " + value + "; ";

        //                    //http://localhost:7250/Pricing?date=2022-01-20&file=vietstar.xlsx&from=hn&b=a&c=5&d=1
        //                }
        //            }
        //        }





        //        if (!error)
        //        {
        //            priceResponse.Message = "OK";

        //            xlApp.Calculate();


        //            priceResponse.SheetName = sheetName;

        //            if (pp1)
        //            {
        //                if (!flag_hangcuon)
        //                {
        //                    priceResponse.LoaiVatLieu = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 2]).Value2.ToString();
        //                    priceResponse.TheTich = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 13 + plus_i]).Value2.ToString();
        //                    priceResponse.SoKg = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 14 + plus_i]).Value2.ToString();
        //                    priceResponse.HaoPhiMachCat = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 15 + plus_i]).Value2.ToString();
        //                    priceResponse.PhiGiaCong = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 17 + plus_i]).Value2.ToString();
        //                    priceResponse.DonGia = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 18 + plus_i]).Value2.ToString();
        //                    priceResponse.DonGiaTheoTSLN = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 20 + plus_i]).Value2.ToString();
        //                    priceResponse.ThanhTienTheoTSLN = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 21 + plus_i]).Value2.ToString();

        //                }
        //                else // hang cuon
        //                {
        //                    priceResponse.LoaiVatLieu = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 2]).Value2.ToString();

        //                    priceResponse.PhiGiaCong = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 12]).Value2.ToString();
        //                    priceResponse.DonGia = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 13]).Value2.ToString();
        //                    priceResponse.DonGiaTheoTSLN = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 15]).Value2.ToString();
        //                    priceResponse.ThanhTienTheoTSLN = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 16]).Value2.ToString();

        //                }
        //            }

        //            if (pp2)
        //            {
        //                priceResponse.FramePrice = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[11, 11]).Value2.ToString();
        //            }

        //            if (pp3)
        //            {

        //            }

        //            if (vietstar)
        //            {
        //                priceResponse.ChiPhiVanChuyenThiXa = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[rowToFillData, 5]).Value2.ToString();
        //                priceResponse.ChiPhiVanChuyenNgoaiThanh = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[rowToFillData, 6]).Value2.ToString();
        //            }




        //            var x = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[11, 4]).Value2.ToString();

        //            //var a1 = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[11, 20]).Value2.ToString();
        //            //var x1 = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 20]).Value2.ToString();




        //        }

        //        else
        //        {
        //            priceResponse.Message = "Error: Sai tên vật liệu";

        //        }
        //        string json = JsonConvert.SerializeObject(priceResponse, Formatting.Indented, new JsonSerializerSettings
        //        {
        //            NullValueHandling = NullValueHandling.Ignore
        //        });

        //        if (wb != null)
        //        {
        //            GC.Collect();
        //            GC.WaitForPendingFinalizers();

        //            wb.Close(false);
        //            Marshal.ReleaseComObject(wb);
        //            Marshal.ReleaseComObject(wbs);
        //            Marshal.ReleaseComObject(wSheet);


        //        }


        //        if (xlApp != null)
        //            xlApp.Quit();


        //        wb = null;
        //        wbs = null;
        //        wSheet = null;
        //        xlApp = null;

        //        GC.Collect();
        //        GC.WaitForPendingFinalizers();

        //        return json;

        //    }

        //    catch (Exception ex)
        //    {
        //        GC.Collect();
        //        GC.WaitForPendingFinalizers();
        //        priceResponse.Message = ex.Message;

        //        string json = JsonConvert.SerializeObject(priceResponse, Formatting.Indented);
        //        //if (wb != null)
        //        //    wb.Close(false);
        //        //if (xlApp != null)
        //        //    xlApp.Quit();
        //        xlApp = null;

        //        //Marshal.ReleaseComObject(wb);
        //        //Marshal.ReleaseComObject(wbs);
        //        //Marshal.ReleaseComObject(wSheet);
        //        GC.Collect();
        //        GC.WaitForPendingFinalizers();

        //        return json;
        //    }
        //}

#pragma warning disable CS8600
#pragma warning disable CS8604
#pragma warning disable CS0168
        public static Models.PriceResponse CalculatePrice(string request)
        {
            var returnValue = new Models.PriceResponse
            {
                Message = "OK"
            };
            string folder = "", fileName = "", filePath = "", rootFolder = "";
            int rowData = 0;
            var parsed = System.Web.HttpUtility.ParseQueryString(request);
            folder = parsed["date"];
            //folder = "Excels";
            rootFolder = "C:\\inetpub\\wwwroot\\pricing\banggia\\";
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
                                            rowData = _i;
                                            var s = parsed["sheet"];
                                            var w = Models.ApiParamenter.GetApiParamenterMaterial(s);
                                            item.Cells[B].PutValue(w);
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
                                    workbook.Save(filePath);
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
                                            rowData = _i;
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

                                    workbook.Save(filePath);
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
                                            rowData = _i;
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
                                    workbook.Save(filePath);
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
                        workbook.Save(filePath);
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
                        workbook.Save(filePath);
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
                                rowData = _i;
                                var s = parsed["from"];
                                var w = Models.ApiParamenter.GetApiParamenterLocation(s);
                                var formula = "";
                                var sumFor = $"=IF(D{_i + 1},120%,100%)*E{_i + 1}";
                                if (w == "Formular HCM")
                                {
                                    formula = $"=INDEX($B$14:$J$20,MATCH(IF(C{_i + 1}>2,2,C{_i + 1}),$A$14:$A$20,1),MATCH(B{_i + 1},$B$13:$J$13,1))+IF(C{_i + 1}>2,ROUNDUP((C{_i + 1}-2)/0.5,0)*HLOOKUP(B{_i + 1},$B$13:$J$21,9,FALSE),0)";
                                }
                                else
                                {
                                    formula = $"=INDEX($B$3:$J$9,MATCH(IF(C{_i + 1}>2,2,C{_i + 1}),$A$3:$A$9,1),MATCH(B{_i + 1},$B$2:$J$2,1))+IF(C{_i + 1}>2,ROUNDUP((C{_i + 1}-2)/0.5,0)*HLOOKUP(B{_i + 1},$B$2:$J$10,9,FALSE),0)";
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
                        workbook.Save(filePath);
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
            return returnValue;
        }
    }
}

