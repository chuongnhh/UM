using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace UM.Models
{
    public class NhanVienView
    {
        public NhanVienView()
        {
        }
        public int Id { get; set; }
        public string HoTen { get; set; }
        public string DienThoai { get; set; }

        public int Ca { get; set; }
        public int Goi { get; set; }
        public int HT { get; set; }
        public int TV { get; set; }
        public int  CF { get; set; }

        public int Ca50 { get; set; }
        public int Ca80 { get; set; }
        public int Ca100 { get; set; }

        public decimal Luong { get; set; }
    }
}