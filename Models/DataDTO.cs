using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PricingService.Models
{
    public class DataDTO
    {
        //pp1
        public string LoaiVatLieu { get; set; }
        public string SheetName { get; set; }
        public string TheTich { get; set; }
        public string SoKg { get; set; }
        public string HaoPhiMachCat { get; set; }
        public string PhiGiaCong { get; set; }
        public string DonGia { get; set; }
        public string DonGiaTheoTSLN { get; set; }
        public string ThanhTienTheoTSLN { get; set; }

        //pp2
        public string FramePrice { get; set; }

        //pp3
        public string ChiPhiVanChuyenThiXa { get; set; }
        public string ChiPhiVanChuyenNgoaiThanh { get; set; }

        //vietstar
    }
}