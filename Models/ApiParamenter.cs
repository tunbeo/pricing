using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PricingService.Models
{
    public static class ApiParamenter
    {
        public static string GetApiParamenterSheet(string key)
        {
            try
            {
                if (!string.IsNullOrEmpty(key))
                {
                    var apiParamenterSheet = ApiParamenterSheet();
                    return apiParamenterSheet[key];
                }
                return "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        public static string GetApiParamenterMaterial(string key)
        {
            try
            {
                if (!string.IsNullOrEmpty(key))
                {
                    var apiParamenterMaterial = ApiParamenterMaterial();
                    return apiParamenterMaterial[key];
                }
                return "";
            }
            catch (Exception)
            {
                return "";
            }
        }
        public static string GetApiParamenterType(string key)
        {
            try
            {
                if (!string.IsNullOrEmpty(key))
                {
                    var apiParamenterType = ApiParamenterType();
                    return apiParamenterType[key];
                }
                return "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        public static string GetApiParamenterLocation(string key)
        {
            try
            {
                if (!string.IsNullOrEmpty(key))
                {
                    var apiParamenterLocation = ApiParamenterLocation();
                    return apiParamenterLocation[key];
                }
                return "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static Dictionary<string, string> ApiParamenterSheet()
        {
            return new Dictionary<string, string>
            {
                //Hàng tấm (thép ko gỉ)
                { "thep_tam_day", "Hàng tấm (thép ko gỉ)" },
                { "thep_tam_la", "Hàng tấm (thép ko gỉ)" },
                //Hàng tấm (bery)
                { "dong_bery", "Hàng tấm (bohler, bery)" },
                { "thep_dung_cu", "Hàng tấm (bohler, bery)" },
                //Hàng tấm (đồng nhôm)
                { "dong_tam_la", "Hàng tấm (đồng nhôm)" },
                { "dong_hop_kim_tam_day", "Hàng tấm (đồng nhôm)" },
                { "dong_tam_day", "Hàng tấm (đồng nhôm)" },
                { "nhom_hop_kim_day", "Hàng tấm (đồng nhôm)" },
                { "nhom_hop_kim_mong", "Hàng tấm (đồng nhôm)" },
                //Hàng thanh
                { "dong_thanh", "Hàng thanh" },
                { "dong_hop_kim_thanh", "Hàng thanh" },
                { "nhom_thanh", "Hàng thanh" },
                { "thep_thanh", "Hàng thanh" },
                { "dong_bery_thanh", "Hàng thanh" },
                //Hàng cuộn
                { "nhom_khong_hop_kim_cuon", "Hàng cuộn" },
                { "nhom_hop_kim_cuon", "Hàng cuộn" },
                { "dong_tinh_che_tam_la", "Hàng cuộn" },
                { "dong_hop_kim_tam_la", "Hàng cuộn" },
                { "thep_la_cuon", "Hàng cuộn" }
            };
        }

        private static Dictionary<string, string> ApiParamenterMaterial()
        {
            return new Dictionary<string, string>
            {
                //Hàng tấm (thép ko gỉ)
                { "thep_tam_day", "Thép tấm dày" },
                { "thep_tam_la", "Thép tấm lá" },
                //Hàng tấm (bery)
                { "dong_bery", "Đồng bery" },
                { "thep_dung_cu", "Thép dụng cụ" },
                //Hàng tấm (đồng nhôm)
                { "dong_tam_la", "Đồng tấm lá" },
                { "dong_hop_kim_tam_day", "Đồng hợp kim tấm dày" },
                { "dong_tam_day", "Đồng tấm dày" },
                { "nhom_hop_kim_day", "Nhôm hợp kim dày" },
                { "nhom_hop_kim_mong", "Nhôm hợp kim mỏng" },
                //Hàng thanh
                { "dong_thanh", "Đồng thanh" },
                { "dong_hop_kim_thanh", "Đồng hợp kim thanh" },
                { "nhom_thanh", "Nhôm thanh" },
                { "thep_thanh", "Thép thanh" },
                { "dong_bery_thanh", "Đồng bery thanh" },
                //Hàng cuộn
                { "nhom_khong_hop_kim_cuon", "Nhôm không hợp kim cuộn" },
                { "nhom_hop_kim_cuon", "Nhôm hợp kim cuộn" },
                { "dong_tinh_che_tam_la", "Đồng tinh chế tấm lá" },
                { "dong_hop_kim_tam_la", "Đồng hợp kim tấm lá" },
                { "thep_la_cuon", "Thép lá cuộn" }
            };
        }

        private static Dictionary<string, string> ApiParamenterType()
        {
            return new Dictionary<string, string>
            {
                //Hàng tấm
                { "thep_tam_day", "Hàng tấm" },
                { "thep_tam_la", "Hàng tấm" },
                { "dong_bery", "Hàng tấm" },
                { "thep_dung_cu", "Hàng tấm" },
                { "dong_tam_la", "Hàng tấm" },
                { "dong_hop_kim_tam_day", "Hàng tấm" },
                { "dong_tam_day", "Hàng tấm" },
                { "nhom_hop_kim_day", "Hàng tấm" },
                { "nhom_hop_kim_mong", "Hàng tấm" },
                //Hàng thanh
                { "dong_thanh", "Hàng thanh" },
                { "dong_hop_kim_thanh", "Hàng thanh" },
                { "nhom_thanh", "Hàng thanh" },
                { "thep_thanh", "Hàng thanh" },
                { "dong_bery_thanh", "Hàng thanh" },
                //Hàng cuộn
                { "nhom_khong_hop_kim_cuon", "Hàng cuộn" },
                { "nhom_hop_kim_cuon", "Hàng cuộn" },
                { "dong_tinh_che_tam_la", "Hàng cuộn" },
                { "dong_hop_kim_tam_la", "Hàng cuộn" },
                { "thep_la_cuon", "Hàng cuộn" }
            };
        }

        private static Dictionary<string, string> ApiParamenterLocation()
        {
            return new Dictionary<string, string>
            {
                { "hn", "Formular HN" },
                { "hy", "Formular HY" },
                { "hcm", "Formular HCM" },
            };
        }
    }
}