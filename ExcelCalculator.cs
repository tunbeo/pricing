using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
//using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace PricingService
{
    public static class ExcelCalculator
    {
        
        public static Models.PriceResponse CalculatePrice(string request)
        {
            var returnValue = new Models.PriceResponse
            {
                Message = "Fail"
            };
            string folder = "", fileName = "", filePath = "";
            int rowData = 0;
            var parsed = System.Web.HttpUtility.ParseQueryString(request);
            //folder = parsed["date"];
            folder = "Excels";
            fileName = parsed["file"];
            filePath = Path.Combine("C:\\inetpub\\wwwroot\\Pricing\\BangGia" + folder, fileName);
            Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();
            string A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U;
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
            return returnValue;
        }
    }
}

