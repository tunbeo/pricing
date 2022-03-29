using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace PricingService
{
    public static class ExcelCalculator
    {        
        public static string Calculate(string query)
        {            
            string rootPath = "C:\\inetpub\\wwwroot\\Pricing\\BangGia\\";
            string path = "C:\\inetpub\\wwwroot\\Pricing\\BangGia\\2021-12-21 10-07\\pp1_sale.xlsx";
            Application xlApp = null;// = new Application();
            Workbook wb = null;// = xlApp.Workbooks.Open(path);
            Worksheet wSheet = null;// = (Worksheet)wb.Worksheets[3];
            Workbooks wbs = null;
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
                xlApp = new Application();
                wbs = xlApp.Workbooks;
                wb = wbs.Open(path);
                wSheet = (Worksheet)wb.Worksheets[3];

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

                            //xlApp = new Application();
                            //wbs = xlApp.Workbooks;
                            wb = wbs.Open(path);
                            if (value.Contains("pp2"))
                            {
                                wSheet = (Worksheet)wb.Worksheets[1];
                                sheetName = wSheet.Name;
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
                                wSheet = (Worksheet)wb.Worksheets[1];
                                sheetName = wSheet.Name;

                                pp3 = false;
                                pp2 = false;
                                vietstar = false;
                            }    
                            else if (value.Contains("pp3"))
                            {
                                pp3 = true;
                                wSheet = (Worksheet)wb.Worksheets[1];
                                sheetName = wSheet.Name;

                                pp1 = false;
                                pp3 = false;
                                vietstar = false;
                            }    
                            else if (value.StartsWith("vietstar"))
                            {
                                vietstar = true;
                                rowToFillData = 2;
                                wSheet = (Worksheet)wb.Worksheets[1];
                                sheetName = wSheet.Name;

                                pp1 = false;
                                pp2 = false;
                                pp3 = false;
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
                            var value = parsed[key];
                            if (value == "thep_tam_day" || value == "thep_tam_la")
                            {
                                wSheet = (Worksheet)wb.Worksheets[3];
                                if (value == "thep_tam_day")
                                    LoaiVatLieu = "Thép tấm dày";
                                else
                                    LoaiVatLieu = "Thép tấm lá";
                            }

                            else if (value == "dong_bery" || value == "thep_dung_cu")
                            {
                                wSheet = (Worksheet)wb.Worksheets[4];
                                if (value == "dong_bery")
                                    LoaiVatLieu = "Đồng bery";
                                else
                                    LoaiVatLieu = "Thép dụng cụ";
                            }

                            else if (value == "dong_tam_la" || value == "dong_hop_kim_tam_day" || value == "dong_tam_day" || value == "nhom_hop_kim_day" || value == "nhom_hop_kim_mong")
                            {
                                wSheet = (Worksheet)wb.Worksheets[5];
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
                                wSheet = (Worksheet)wb.Worksheets[6];
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
                                wSheet = (Worksheet)wb.Worksheets[7];
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
                            sheetName = wSheet.Name;

                            //set b column to material type
                            wSheet.Cells[12, char.ToUpper(char.Parse("b")) - 64] = LoaiVatLieu;
                        }
                            
                        else
                        {
                            var stt = char.ToUpper(char.Parse(key)) - 64;

                            var value = parsed[key];

                            if (pp1)
                            {
                                if (key != "c")
                                {
                                    stt = stt + plus_i;
                                }

                                if (value == "circle")
                                    wSheet.Cells[rowToFillData, stt] = "Tròn";
                                else if (value == "rectangle")
                                    wSheet.Cells[rowToFillData, stt] = "Chữ nhật";
                                else
                                    wSheet.Cells[rowToFillData, stt] = value;
                            }    
                            if (pp2)
                            {
                                wSheet.Cells[rowToFillData, stt] = value;
                            }    
                            if (vietstar)
                            {
                                wSheet.Cells[rowToFillData, stt] = value;
                            }

                            

                            var tencolumn = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[rowToGetColumnName, stt]).Value2.ToString();
                            priceResponse.Input += tencolumn.Replace("\n","") + ": " + value + "; ";

                            //http://localhost:7250/Pricing?date=2022-01-20&file=vietstar.xlsx&from=hn&b=a&c=5&d=1
                        }
                    }
                }
                 

               
                

                if (!error)
                {
                    priceResponse.Message = "OK";

                    xlApp.Calculate();

                    
                    priceResponse.SheetName = sheetName;

                    if (pp1)
                    {
                        if (!flag_hangcuon)
                        {
                            priceResponse.LoaiVatLieu = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 2]).Value2.ToString();
                            priceResponse.TheTich = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 13 + plus_i]).Value2.ToString();
                            priceResponse.SoKg = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 14 + plus_i]).Value2.ToString();
                            priceResponse.HaoPhiMachCat = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 15 + plus_i]).Value2.ToString();
                            priceResponse.PhiGiaCong = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 17 + plus_i]).Value2.ToString();
                            priceResponse.DonGia = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 18 + plus_i]).Value2.ToString();
                            priceResponse.DonGiaTheoTSLN = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 20 + plus_i]).Value2.ToString();
                            priceResponse.ThanhTienTheoTSLN = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 21 + plus_i]).Value2.ToString();

                        }
                        else // hang cuon
                        {
                            priceResponse.LoaiVatLieu = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 2]).Value2.ToString();
                            
                            priceResponse.PhiGiaCong = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 12]).Value2.ToString();
                            priceResponse.DonGia = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 13]).Value2.ToString();
                            priceResponse.DonGiaTheoTSLN = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 15]).Value2.ToString();
                            priceResponse.ThanhTienTheoTSLN = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[12, 16]).Value2.ToString();

                        }
                    }    
                    
                    if (pp2)
                    {
                        priceResponse.FramePrice = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[11, 11]).Value2.ToString();
                    }   
                    
                    if (pp3)
                    {

                    }

                    if (vietstar)
                    {
                        priceResponse.ChiPhiVanChuyenThiXa = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[rowToFillData, 5]).Value2.ToString();
                        priceResponse.ChiPhiVanChuyenNgoaiThanh = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[rowToFillData, 6]).Value2.ToString();
                    }




                    var x = ((Microsoft.Office.Interop.Excel.Range)wSheet.Cells[11, 4]).Value2.ToString();

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

                if (wb != null)
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    wb.Close(false);
                    Marshal.ReleaseComObject(wb);
                    Marshal.ReleaseComObject(wbs);
                    Marshal.ReleaseComObject(wSheet);
                    

                }
                    

                if (xlApp != null)
                    xlApp.Quit();
                

                wb = null;
                wbs = null;
                wSheet = null;
                xlApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();

                return json;

            }

            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                priceResponse.Message = ex.Message;

                string json = JsonConvert.SerializeObject(priceResponse, Formatting.Indented);
                //if (wb != null)
                //    wb.Close(false);
                //if (xlApp != null)
                //    xlApp.Quit();
                xlApp = null;

                //Marshal.ReleaseComObject(wb);
                //Marshal.ReleaseComObject(wbs);
                //Marshal.ReleaseComObject(wSheet);
                GC.Collect();
                GC.WaitForPendingFinalizers();

                return json;
            }        
        }
    }
}

