using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace UM.Models
{
    public class NhanVien
    {
        public int Id { get; set; }
        public string HoTen { get; set; }
        public string DienThoai { get; set; }

        public string ThoiGian1 { get; set; }
        public string Ca1 { get; set; }

        public string ThoiGian2 { get; set; }
        public string Ca2 { get; set; }

        public string ThoiGian3 { get; set; }
        public string Ca3 { get; set; }

        public string Ngay { get; set; }
    }
}