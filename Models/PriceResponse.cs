namespace PricingService.Models
{
    public class PriceResponse
    {
        //Input
        public string? Input { get; set; }
        

        //pp1
        public string? LoaiVatLieu { get; set; }
        public string? SheetName { get; set; }
        public string? TheTich { get; set; }
        public string? SoKg { get; set; }
        public string? HaoPhiMachCat { get; set; }
        public string? PhiGiaCong { get; set; }
        public string? DonGia { get; set; }
        public string? DonGiaTheoTSLN { get; set; }
        public string? ThanhTienTheoTSLN { get; set; }

        public string? PHAY { get; set; }
        public string? MatPhay { get; set; }
        public string? PhayStatus { get; set; }
        public string? PhiPhay { get; set; }


        //pp2
        public string? FramePrice { get; set; }

        //pp3
        public string? ChiPhiVanChuyenThiXa { get; set; }
        public string? ChiPhiVanChuyenNgoaiThanh { get; set; }

        //vietstar

        // _ppt3
        //public string? DonGia { get; set; }
        //public string? DonGiaTheoTyGiaBanVietCombank { get; set; }

        public string? Message { get; set; }
    }
}
