using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using UM.Models;

namespace UM.Controllers
{
    public class HomeController : Controller
    {
        //MyDbContext db = new MyDbContext();

        private static List<NhanVien> nhanViens;

        public ActionResult Index()
        {
            nhanViens = null;
            return View();
        }

        public JsonResult GetData()
        {
            return Json(new { data = nhanViens }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult Upload(FormCollection formCollection)
        {
            if (Request != null)
            {
                HttpPostedFileBase file = Request.Files["UploadedFile"];
                if ((file != null) && (file.ContentLength > 0) && !string.IsNullOrEmpty(file.FileName))
                {
                    try
                    {
                        nhanViens = new List<NhanVien>();
                        string fileName = file.FileName;
                        string fileContentType = file.ContentType;
                        byte[] fileBytes = new byte[file.ContentLength];
                        var data = file.InputStream.Read(fileBytes, 0, Convert.ToInt32(file.ContentLength));


                        using (var package = new ExcelPackage(file.InputStream))
                        {
                            var currentSheet = package.Workbook.Worksheets;
                            var workSheet = currentSheet.First();
                            var noOfCol = workSheet.Dimension.End.Column;
                            var noOfRow = workSheet.Dimension.End.Row;

                            for (int rowIterator = 1; rowIterator <= noOfRow; rowIterator++)
                            {
                                var nhanVien = new NhanVien
                                {
                                    HoTen = workSheet.Cells[rowIterator, 1].Text,
                                    DienThoai = workSheet.Cells[rowIterator, 2].Text,
                                    ThoiGian1 = workSheet.Cells[rowIterator, 3].Text,
                                    Ca1 = workSheet.Cells[rowIterator, 4].Text,
                                    ThoiGian2 = workSheet.Cells[rowIterator, 5].Text,
                                    Ca2 = workSheet.Cells[rowIterator, 6].Text,
                                    ThoiGian3 = workSheet.Cells[rowIterator, 7].Text,
                                    Ca3 = workSheet.Cells[rowIterator, 8].Text,
                                    KickSale = workSheet.Cells[rowIterator, 9].Text,
                                    Ngay = workSheet.Cells[rowIterator, 10].Text,
                                };
                                nhanViens.Add(nhanVien);
                            }
                        }
                        return Json(new { status = true, data = nhanViens }, JsonRequestBehavior.AllowGet);
                    }
                    catch (Exception ex)
                    {

                        return Json(new { status = false, data = ex.Message }, JsonRequestBehavior.AllowGet);
                    }
                }
            }
            return Json(new { status = false, data = new List<NhanVien>() }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult Download()
        {
            if (nhanViens == null)
            {
                return Json(new { status = false, data = new List<NhanVienView>() }, JsonRequestBehavior.AllowGet);
            }

            ExcelPackage excel = new ExcelPackage();
            var workSheet = excel.Workbook.Worksheets.Add("chuongnh");

            workSheet.Row(1).Style.Font.Bold = true;
            // Assign borders
            workSheet.Cells[1, 1, 1, 13].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            workSheet.Cells[1, 1, 1, 13].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            workSheet.Cells[1, 1, 1, 13].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            workSheet.Cells[1, 1, 1, 13].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            workSheet.Cells[1, 1].Value = "STT";
            workSheet.Cells[1, 2].Value = "Họ tên";
            workSheet.Cells[1, 3].Value = "Điện thoại";
            workSheet.Cells[1, 4].Value = "Ca";
            workSheet.Cells[1, 5].Value = "Gọi";
            workSheet.Cells[1, 6].Value = "HT";
            workSheet.Cells[1, 7].Value = "CF";
            workSheet.Cells[1, 8].Value = "TV";
            workSheet.Cells[1, 9].Value = "Ca50";
            workSheet.Cells[1, 10].Value = "Ca80";
            workSheet.Cells[1, 11].Value = "Ca100";
            workSheet.Cells[1, 12].Value = "Kick sales";
            workSheet.Cells[1, 13].Value = "Tiền lương";
            //Body of table  
            //  

            var nhanvienviews = nhanViens
                .GroupBy(z => z.HoTen)
                .Select(x => new NhanVienView
                {
                    HoTen = x.FirstOrDefault().HoTen,
                    DienThoai = x.FirstOrDefault().DienThoai,
                    Ca = x.Count(q => q.Ca1.Trim().Length > 0) + x.Count(q => q.Ca2.Trim().Length > 0) + x.Count(q => q.Ca3.Trim().Length > 0),

                    Goi = x.Where(z => z.Ca1.Trim().Length > 0 &&
                               GetDigit(z.Ca1.Trim()) >= 0).Count() +

                              x.Where(z => z.Ca2.Trim().Length > 0 &&
                              GetDigit(z.Ca2.Trim()) >= 0).Count() +

                              x.Where(z => z.Ca3.Trim().Length > 0 &&
                               GetDigit(z.Ca3.Trim()) >= 0).Count(),

                    HT = x.Count(q => q.Ca1.ToUpper().Contains("HT")) + x.Count(q => q.Ca2.ToUpper().Contains("HT")) + x.Count(q => q.Ca3.ToUpper().Contains("HT")),
                    CF = x.Count(q => q.Ca1.ToUpper().Contains("CF")) + x.Count(q => q.Ca2.ToUpper().Contains("CF")) + x.Count(q => q.Ca3.ToUpper().Contains("CF")),
                    TV = x.Count(q => q.ThoiGian1.ToUpper().Contains("TV")) + x.Count(q => q.ThoiGian2.ToUpper().Contains("TV")) + x.Count(q => q.ThoiGian3.ToUpper().Contains("TV")),

                    Ca50 = x.Where(z => z.Ca1.Trim().Length > 0 &&
                              z.ThoiGian1.ToUpper().Contains("TV") == false &&
                                z.Ca1.ToUpper().Contains("HT") == false &&
                             (GetDigit(z.Ca1.Trim()) >= 0 && GetDigit(z.Ca1.Trim()) < 8) || (z.Ca1.ToUpper().Contains("CF") && GetDigit(z.Ca1.Trim()) < 8)).Count() +

                              x.Where(z => z.Ca2.Trim().Length > 0 &&
                              z.ThoiGian2.ToUpper().Contains("TV") == false &&
                                z.Ca2.ToUpper().Contains("HT") == false &&
                              (GetDigit(z.Ca2.Trim()) >= 0 && GetDigit(z.Ca2.Trim()) < 8) || (z.Ca2.ToUpper().Contains("CF") && GetDigit(z.Ca2.Trim()) < 8)).Count() +

                              x.Where(z => z.Ca3.Trim().Length > 0 &&
                              z.ThoiGian3.ToUpper().Contains("TV") == false &&
                                z.Ca3.ToUpper().Contains("HT") == false &&
                               (GetDigit(z.Ca3.Trim()) >= 0 && GetDigit(z.Ca3.Trim()) < 8) || (z.Ca3.ToUpper().Contains("CF") && GetDigit(z.Ca3.Trim()) < 8)).Count(),


                    Ca80 = x.Where(z => z.Ca1.Trim().Length > 0 &&
                              z.ThoiGian1.ToUpper().Contains("TV") == false &&
                                z.Ca1.ToUpper().Contains("HT") == false &&
                              GetDigit(z.Ca1.Trim()) >= 8).Count() +

                              x.Where(z => z.Ca2.Trim().Length > 0 &&
                              z.ThoiGian2.ToUpper().Contains("TV") == false &&
                                z.Ca2.ToUpper().Contains("HT") == false &&
                              GetDigit(z.Ca2.Trim()) >= 8).Count() +

                              x.Where(z => z.Ca3.Trim().Length > 0 &&
                              z.ThoiGian3.ToUpper().Contains("TV") == false &&
                                z.Ca3.ToUpper().Contains("HT") == false &&
                               GetDigit(z.Ca3.Trim()) >= 8).Count(),
                    KickSale = x.Sum(q => GetDigit(q.KickSale) > -1 ? GetDigit(q.KickSale) : 0),
                    Ca100 = x.Count(q => q.Ca1.ToUpper().Contains("HT")) + x.Count(q => q.Ca2.ToUpper().Contains("HT")) + x.Count(q => q.Ca3.ToUpper().Contains("HT")),
                }).ToList();

            int recordIndex = 2;
            foreach (var nhanvien in nhanvienviews)
            {
                workSheet.Cells[recordIndex, 1].Value = (recordIndex - 1).ToString();
                workSheet.Cells[recordIndex, 2].Value = nhanvien.HoTen;
                workSheet.Cells[recordIndex, 3].Value = nhanvien.DienThoai;
                workSheet.Cells[recordIndex, 4].Value = nhanvien.Ca;
                workSheet.Cells[recordIndex, 5].Value = nhanvien.Goi;
                workSheet.Cells[recordIndex, 6].Value = nhanvien.HT;
                workSheet.Cells[recordIndex, 7].Value = nhanvien.CF;
                workSheet.Cells[recordIndex, 8].Value = nhanvien.TV;
                workSheet.Cells[recordIndex, 9].Value = nhanvien.Ca50;
                workSheet.Cells[recordIndex, 10].Value = nhanvien.Ca80;
                workSheet.Cells[recordIndex, 11].Value = nhanvien.Ca100;
                workSheet.Cells[recordIndex, 12].Value = nhanvien.KickSale;
                workSheet.Cells[recordIndex, 13].Value = nhanvien.Ca50 * 50000 + nhanvien.Ca80 * 80000 + nhanvien.Ca100 * 100000+nhanvien.KickSale*50000;
                //number with 2 decimal places and thousand separator and money symbol
                workSheet.Cells[recordIndex, 13].Style.Numberformat.Format = "#,##0";

                if (nhanvien.TV > 0)
                {
                    workSheet.Cells[recordIndex, 1, recordIndex, 13].Style.Font.Color.SetColor(Color.Red);
                }
                workSheet.Cells[recordIndex, 1, recordIndex, 13].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[recordIndex, 1, recordIndex, 13].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[recordIndex, 1, recordIndex, 13].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[recordIndex, 1, recordIndex, 13].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                recordIndex++;
            }
            for (int i = 1; i <= 13; i++)
            {
                workSheet.Column(i).AutoFit();
            }

            string excelName = "luong_chuongnh";
            using (var memoryStream = new MemoryStream())
            {
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment; filename=" + excelName + ".xlsx");
                excel.SaveAs(memoryStream);
                memoryStream.WriteTo(Response.OutputStream);
                Response.Flush();
                Response.End();
            }
            return Json(new { status = true, data = "" }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult Download1()
        {
            if (nhanViens == null)
            {
                return Json(new { status = false, data = new List<NhanVienView>() }, JsonRequestBehavior.AllowGet);
            }

            ExcelPackage excel = new ExcelPackage();
            var workSheet = excel.Workbook.Worksheets.Add("chuongnh");
            workSheet.Row(1).Style.Font.Bold = true;
            // Assign borders
            workSheet.Cells[1, 1, 1, 13].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            workSheet.Cells[1, 1, 1, 13].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            workSheet.Cells[1, 1, 1, 13].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            workSheet.Cells[1, 1, 1, 13].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            workSheet.Cells[1, 1].Value = "STT";
            workSheet.Cells[1, 2].Value = "Họ tên";
            workSheet.Cells[1, 3].Value = "Điện thoại";
            workSheet.Cells[1, 4].Value = "Ca";
            workSheet.Cells[1, 5].Value = "Gọi";
            workSheet.Cells[1, 6].Value = "HT";
            workSheet.Cells[1, 7].Value = "CF";
            workSheet.Cells[1, 8].Value = "TV";
            workSheet.Cells[1, 9].Value = "Ca50";
            workSheet.Cells[1, 10].Value = "Ca80";
            workSheet.Cells[1, 11].Value = "Ca100";
            workSheet.Cells[1, 12].Value = "Kick sales";
            workSheet.Cells[1, 13].Value = "Tiền lương";
            //Body of table  
            //  
            var nhanviengroup = nhanViens
                .GroupBy(x => x.Ngay)
                .Select(x => x.ToList()).ToList();

            int recordIndex = 2;
            foreach (var item in nhanviengroup)
            {

                var nhanvienviews = item
             .GroupBy(z => z.HoTen)
             .Select(x => new NhanVienView
             {
                 HoTen = x.FirstOrDefault().HoTen,
                 DienThoai = x.FirstOrDefault().DienThoai,
                 Ca = x.Count(q => q.Ca1.Trim().Length > 0) + x.Count(q => q.Ca2.Trim().Length > 0) + x.Count(q => q.Ca3.Trim().Length > 0),

                 Goi = x.Where(z => z.Ca1.Trim().Length > 0 &&
                            GetDigit(z.Ca1.Trim()) >= 0).Count() +

                           x.Where(z => z.Ca2.Trim().Length > 0 &&
                           GetDigit(z.Ca2.Trim()) >= 0).Count() +

                           x.Where(z => z.Ca3.Trim().Length > 0 &&
                            GetDigit(z.Ca3.Trim()) >= 0).Count(),

                 HT = x.Count(q => q.Ca1.ToUpper().Contains("HT")) + x.Count(q => q.Ca2.ToUpper().Contains("HT")) + x.Count(q => q.Ca3.ToUpper().Contains("HT")),
                 CF = x.Count(q => q.Ca1.ToUpper().Contains("CF")) + x.Count(q => q.Ca2.ToUpper().Contains("CF")) + x.Count(q => q.Ca3.ToUpper().Contains("CF")),
                 TV = x.Count(q => q.ThoiGian1.ToUpper().Contains("TV")) + x.Count(q => q.ThoiGian2.ToUpper().Contains("TV")) + x.Count(q => q.ThoiGian3.ToUpper().Contains("TV")),

                 Ca50 = x.Where(z => z.Ca1.Trim().Length > 0 &&
                           z.ThoiGian1.ToUpper().Contains("TV") == false &&
                             z.Ca1.ToUpper().Contains("HT") == false &&
                          (GetDigit(z.Ca1.Trim()) >= 0 && GetDigit(z.Ca1.Trim()) < 8) || (z.Ca1.ToUpper().Contains("CF") && GetDigit(z.Ca1.Trim()) < 8)).Count() +

                           x.Where(z => z.Ca2.Trim().Length > 0 &&
                           z.ThoiGian2.ToUpper().Contains("TV") == false &&
                             z.Ca2.ToUpper().Contains("HT") == false &&
                           (GetDigit(z.Ca2.Trim()) >= 0 && GetDigit(z.Ca2.Trim()) < 8) || (z.Ca2.ToUpper().Contains("CF") && GetDigit(z.Ca2.Trim()) < 8)).Count() +

                           x.Where(z => z.Ca3.Trim().Length > 0 &&
                           z.ThoiGian3.ToUpper().Contains("TV") == false &&
                             z.Ca3.ToUpper().Contains("HT") == false &&
                            (GetDigit(z.Ca3.Trim()) >= 0 && GetDigit(z.Ca3.Trim()) < 8) || (z.Ca3.ToUpper().Contains("CF") && GetDigit(z.Ca3.Trim()) < 8)).Count(),


                 Ca80 = x.Where(z => z.Ca1.Trim().Length > 0 &&
                           z.ThoiGian1.ToUpper().Contains("TV") == false &&
                             z.Ca1.ToUpper().Contains("HT") == false &&
                           GetDigit(z.Ca1.Trim()) >= 8).Count() +

                           x.Where(z => z.Ca2.Trim().Length > 0 &&
                           z.ThoiGian2.ToUpper().Contains("TV") == false &&
                             z.Ca2.ToUpper().Contains("HT") == false &&
                           GetDigit(z.Ca2.Trim()) >= 8).Count() +

                           x.Where(z => z.Ca3.Trim().Length > 0 &&
                           z.ThoiGian3.ToUpper().Contains("TV") == false &&
                             z.Ca3.ToUpper().Contains("HT") == false &&
                            GetDigit(z.Ca3.Trim()) >= 8).Count(),
                 KickSale = x.Sum(q => GetDigit(q.KickSale) > -1 ? GetDigit(q.KickSale) : 0),
                 Ca100 = x.Count(q => q.Ca1.ToUpper().Contains("HT")) + x.Count(q => q.Ca2.ToUpper().Contains("HT")) + x.Count(q => q.Ca3.ToUpper().Contains("HT")),
             }).ToList();
                workSheet.Cells[recordIndex, 1].Value = "Tuần " + item.FirstOrDefault().Ngay;
                workSheet.Cells[recordIndex, 1, recordIndex, 13].Merge = true;
                workSheet.Cells[recordIndex, 1, recordIndex, 13].Style.Font.Bold = true;
                workSheet.Cells[recordIndex, 1, recordIndex, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //workSheet.Cells[recordIndex, 1, recordIndex, 12].Style.Fill.BackgroundColor.SetColor(Color.Cyan);

                recordIndex++;

                int stt = 1;
                foreach (var nhanvien in nhanvienviews)
                {
                    workSheet.Cells[recordIndex, 1].Value = (stt++).ToString();
                    workSheet.Cells[recordIndex, 2].Value = nhanvien.HoTen;
                    workSheet.Cells[recordIndex, 3].Value = nhanvien.DienThoai;
                    workSheet.Cells[recordIndex, 4].Value = nhanvien.Ca;
                    workSheet.Cells[recordIndex, 5].Value = nhanvien.Goi;
                    workSheet.Cells[recordIndex, 6].Value = nhanvien.HT;
                    workSheet.Cells[recordIndex, 7].Value = nhanvien.CF;
                    workSheet.Cells[recordIndex, 8].Value = nhanvien.TV;
                    workSheet.Cells[recordIndex, 9].Value = nhanvien.Ca50;
                    workSheet.Cells[recordIndex, 10].Value = nhanvien.Ca80;
                    workSheet.Cells[recordIndex, 11].Value = nhanvien.Ca100;
                    workSheet.Cells[recordIndex, 12].Value = nhanvien.KickSale;
                    workSheet.Cells[recordIndex, 13].Value = nhanvien.Ca50 * 50000 + nhanvien.Ca80 * 80000 + nhanvien.Ca100 * 100000 + nhanvien.KickSale * 50000;
                    //number with 2 decimal places and thousand separator and money symbol
                    workSheet.Cells[recordIndex, 13].Style.Numberformat.Format = "#,##0";

                    if (nhanvien.TV > 0)
                    {
                        workSheet.Cells[recordIndex, 1, recordIndex, 13].Style.Font.Color.SetColor(Color.Red);
                    }
                    workSheet.Cells[recordIndex, 1, recordIndex, 13].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[recordIndex, 1, recordIndex, 13].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[recordIndex, 1, recordIndex, 13].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[recordIndex, 1, recordIndex, 13].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    recordIndex++;
                }
                for (int i = 1; i <= 13; i++)
                {
                    workSheet.Column(i).AutoFit();
                }
            }


            string excelName = "nhom_chuongnh";
            using (var memoryStream = new MemoryStream())
            {
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment; filename=" + excelName + ".xlsx");
                excel.SaveAs(memoryStream);
                memoryStream.WriteTo(Response.OutputStream);
                Response.Flush();
                Response.End();
            }
            return Json(new { status = true, data = "" }, JsonRequestBehavior.AllowGet);
        }

        public int GetDigit(string str)
        {
            int digit = -1;

            var d = Regex.Match(str, @"\d+").Value;
            if (d.Length > 0)
                int.TryParse(d, out digit);
            return digit;
        }
    }
}