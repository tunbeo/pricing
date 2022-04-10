using System.Data;

namespace PricingService.Models
{
    public class MeciReponse
    {
        public string? Input { get; set; }
        public string? Message { get; set; }
        public DataTable GiaBan { get; set; }
        public VatTuTieuHao VatTu {get;set;}
    }
    public class VatTuTieuHao
    {
        public DataTable Nhom { get; set; }
        public DataTable Luoi { get; set; }
        public DataTable ChiTietNhua { get; set; }
        public DataTable Gioang { get; set; }
        public DataTable LoXoVongBi { get; set; }
        public DataTable DinhVitKeo { get; set; }
        public DataTable PhatSinh { get; set; }

    }
}
