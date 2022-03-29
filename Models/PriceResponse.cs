namespace PricingService.Models
{
    public class PriceResponse
    {
        //Input
        public string? Input { get; set; }
        //public string? i_Day { get; set; }
        //public string? i_Rong { get; set; }
        //public string? i_Dai { get; set; }
        //public string? i_GiaVonNguyenTam { get; set; }
        //public string? i_SoPcs { get; set; }
        //public string? i_TSLN { get; set; }
        //public string? i_CongNo { get; set; }
        //public string? i_VanPhi { get; set; }
        //public string? i_PhiHQ { get; set; }
        //public string? i_HinhThai { get; set; }
        //public string? i_ThueXK { get; set; }
        //Output
        
        //pp1
        public string? LoaiVatLieu { get; set; }
        public string? SheetName { get; set; }
        public string? TheTich { get; set; }
        public string? SoKg { get; set; }
        public string? HaoPhiMachCat { get; set; }
        public string? PhiGiaCong { get; set; }
        public string? DonGia {get; set; }
        public string? DonGiaTheoTSLN { get; set; }
        public string? ThanhTienTheoTSLN { get; set; }

        //pp2
        public string? FramePrice { get; set; }

        //pp3
        public string? ChiPhiVanChuyenThiXa { get; set; }
        public string? ChiPhiVanChuyenNgoaiThanh { get; set; }

        //vietstar


        public string? Message { get; set; }
    }
}
