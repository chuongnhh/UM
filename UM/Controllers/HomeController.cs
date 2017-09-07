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

            var d = Regex.Split(str, @"\D+");
            string x = "";
            foreach (var item in d)
            {
                x += item;
            }
            Console.WriteLine(x);
            if (x.Length > 0)
                int.TryParse(x, out digit);
            return digit;
        }


        public ActionResult PhoneNumber()
        {
            phoneNumbers = null;
            return View();
        }


        public static List<PhoneNumberView> phoneNumbers;

        public JsonResult UploadPhoneNumber(FormCollection formCollection)
        {
            if (Request != null)
            {
                HttpPostedFileBase file = Request.Files["UploadedFile"];
                if ((file != null) && (file.ContentLength > 0) && !string.IsNullOrEmpty(file.FileName))
                {
                    try
                    {
                        phoneNumbers = new List<PhoneNumberView>();
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
                                var phoneNumber = new PhoneNumberView
                                {
                                    HoTen = workSheet.Cells[rowIterator, 1].Text,
                                    DienThoai = workSheet.Cells[rowIterator, 2].Text,
                                    DiaChi = workSheet.Cells[rowIterator, 3].Text

                                };
                                phoneNumbers.Add(phoneNumber);
                            }
                        }
                        return Json(new { status = true, data = phoneNumbers }, JsonRequestBehavior.AllowGet);
                    }
                    catch (Exception ex)
                    {

                        return Json(new { status = false, data = ex.Message }, JsonRequestBehavior.AllowGet);
                    }
                }
            }
            return Json(new { status = false, data = new List<PhoneNumberView>() }, JsonRequestBehavior.AllowGet);
        }


        List<string> viettels2 = new List<string> {
            "86", "96", "97", "98"
        };
        List<string> viettels3 = new List<string> {
            "162", "163", "164", "165", "166", "167", "168", "169"
        };

        List<string> mobifones2 = new List<string> {
            "90", "93"
        };
        List<string> mobifones3 = new List<string> {
           "120", "121", "122", "126", "128"
        };

        List<string> vinaphones2 = new List<string> {
            "91", "94"
        };
        List<string> vinaphones3 = new List<string> {
            "123", "124", "125", "127", "129"
        };

        List<string> vietnamobiles2 = new List<string> {
            "92"
        };
        List<string> vietnamobiles3 = new List<string> {
            "188" , "186"
        };

        List<string> gmobiles2 = new List<string> {
            "99"
        };
        List<string> gmobiles3 = new List<string> {
            "199"
        };

        public JsonResult DownloadPhoneNumber()
        {
            if (phoneNumbers == null)
            {
                return Json(new { status = false, data = new List<PhoneNumberView>() }, JsonRequestBehavior.AllowGet);
            }

            var lstPhoneNumber = phoneNumbers.
                Where(x => GetDigit(x.DienThoai).ToString().Length > 8)
                .Select(x => new PhoneNumberView
                {
                    HoTen = x.HoTen,
                    DiaChi = x.DiaChi,
                    DienThoai = GetDigit(x.DienThoai).ToString()

                }).ToList();


            var lstViettel = lstPhoneNumber
                .Where(x => viettels3.Any(a => a.Contains(x.DienThoai.Substring(0, 3))) ||
                viettels2.Any(a => a.Contains(x.DienThoai.Substring(0, 2))))
                .ToList();

            var lstMobifone = lstPhoneNumber
               .Where(x => mobifones3.Any(a => a.Contains(x.DienThoai.Substring(0, 3))) ||
               mobifones2.Any(a => a.Contains(x.DienThoai.Substring(0, 2))))
               .ToList();

            var lstVinaphone = lstPhoneNumber
             .Where(x => vinaphones3.Any(a => a.Contains(x.DienThoai.Substring(0, 3))) ||
             vinaphones2.Any(a => a.Contains(x.DienThoai.Substring(0, 2))))
             .ToList();

            var lstVietnamobile = lstPhoneNumber
            .Where(x => vietnamobiles3.Any(a => a.Contains(x.DienThoai.Substring(0, 3))) ||
            vietnamobiles2.Any(a => a.Contains(x.DienThoai.Substring(0, 2))))
            .ToList();

            var lstGmobile = lstPhoneNumber
               .Where(x => gmobiles3.Any(a => a.Contains(x.DienThoai.Substring(0, 3))) ||
               gmobiles2.Any(a => a.Contains(x.DienThoai.Substring(0, 2))))
               .ToList();

            #region Viettel
            // Viettel
            ExcelPackage excel = new ExcelPackage();
            var workSheetAll = excel.Workbook.Worksheets.Add("Tất cả");

            workSheetAll.Row(1).Style.Font.Bold = true;
            // Assign borders
            workSheetAll.Cells[1, 1, 1, 4].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            workSheetAll.Cells[1, 1, 1, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            workSheetAll.Cells[1, 1, 1, 4].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            workSheetAll.Cells[1, 1, 1, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            workSheetAll.Cells[1, 1].Value = "STT";
            workSheetAll.Cells[1, 2].Value = "Họ tên";
            workSheetAll.Cells[1, 3].Value = "Điện thoại";
            workSheetAll.Cells[1, 4].Value = "Địa chỉ";
            //Body of table  
            //  

            int recordIndex = 2;
            foreach (var nhanvien in lstPhoneNumber)
            {

                workSheetAll.Cells[recordIndex, 1].Value = (recordIndex - 1).ToString();
                workSheetAll.Cells[recordIndex, 2].Value = nhanvien.HoTen;
                workSheetAll.Cells[recordIndex, 3].Value = nhanvien.DienThoai;
                workSheetAll.Cells[recordIndex, 4].Value = nhanvien.DiaChi;
                //number with 2 decimal places and thousand separator and money symbol

                workSheetAll.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                workSheetAll.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                workSheetAll.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                workSheetAll.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                recordIndex++;
            }
            for (int i = 1; i <= 4; i++)
            {
                workSheetAll.Column(i).AutoFit();
            }
            #endregion

            #region Viettel
            // Viettel
            var workSheetViettel = excel.Workbook.Worksheets.Add("Viettel");

            workSheetViettel.Row(1).Style.Font.Bold = true;
            // Assign borders
            workSheetViettel.Cells[1, 1, 1, 4].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            workSheetViettel.Cells[1, 1, 1, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            workSheetViettel.Cells[1, 1, 1, 4].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            workSheetViettel.Cells[1, 1, 1, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            workSheetViettel.Cells[1, 1].Value = "STT";
            workSheetViettel.Cells[1, 2].Value = "Họ tên";
            workSheetViettel.Cells[1, 3].Value = "Điện thoại";
            workSheetViettel.Cells[1, 4].Value = "Địa chỉ";
            //Body of table  
            //  

            recordIndex = 2;
            foreach (var nhanvien in lstViettel)
            {
                workSheetViettel.Cells[recordIndex, 1].Value = (recordIndex - 1).ToString();
                workSheetViettel.Cells[recordIndex, 2].Value = nhanvien.HoTen;
                workSheetViettel.Cells[recordIndex, 3].Value = nhanvien.DienThoai;
                workSheetViettel.Cells[recordIndex, 4].Value = nhanvien.DiaChi;
                //number with 2 decimal places and thousand separator and money symbol

                workSheetViettel.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                workSheetViettel.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                workSheetViettel.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                workSheetViettel.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                recordIndex++;
            }
            for (int i = 1; i <= 4; i++)
            {
                workSheetViettel.Column(i).AutoFit();
            }
            #endregion

            #region Mobifone
            // Mobifone
            var workSheetMobifone = excel.Workbook.Worksheets.Add("Mobifone");

            workSheetMobifone.Row(1).Style.Font.Bold = true;
            // Assign borders
            workSheetMobifone.Cells[1, 1, 1, 4].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            workSheetMobifone.Cells[1, 1, 1, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            workSheetMobifone.Cells[1, 1, 1, 4].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            workSheetMobifone.Cells[1, 1, 1, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            workSheetMobifone.Cells[1, 1].Value = "STT";
            workSheetMobifone.Cells[1, 2].Value = "Họ tên";
            workSheetMobifone.Cells[1, 3].Value = "Điện thoại";
            workSheetMobifone.Cells[1, 4].Value = "Địa chỉ";
            //Body of table  
            //  

            recordIndex = 2;
            foreach (var nhanvien in lstMobifone)
            {

                workSheetMobifone.Cells[recordIndex, 1].Value = (recordIndex - 1).ToString();
                workSheetMobifone.Cells[recordIndex, 2].Value = nhanvien.HoTen;
                workSheetMobifone.Cells[recordIndex, 3].Value = nhanvien.DienThoai;
                workSheetMobifone.Cells[recordIndex, 4].Value = nhanvien.DiaChi;
                //number with 2 decimal places and thousand separator and money symbol

                workSheetMobifone.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                workSheetMobifone.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                workSheetMobifone.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                workSheetMobifone.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                recordIndex++;
            }
            for (int i = 1; i <= 4; i++)
            {
                workSheetMobifone.Column(i).AutoFit();
            }
            #endregion

            #region Vinaphone
            // Vinaphone
            var workSheetVinaphone = excel.Workbook.Worksheets.Add("Vinaphone");

            workSheetVinaphone.Row(1).Style.Font.Bold = true;
            // Assign borders
            workSheetVinaphone.Cells[1, 1, 1, 4].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            workSheetVinaphone.Cells[1, 1, 1, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            workSheetVinaphone.Cells[1, 1, 1, 4].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            workSheetVinaphone.Cells[1, 1, 1, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            workSheetVinaphone.Cells[1, 1].Value = "STT";
            workSheetVinaphone.Cells[1, 2].Value = "Họ tên";
            workSheetVinaphone.Cells[1, 3].Value = "Điện thoại";
            workSheetVinaphone.Cells[1, 4].Value = "Địa chỉ";
            //Body of table  
            //  

            recordIndex = 2;
            foreach (var nhanvien in lstVinaphone)
            {

                workSheetVinaphone.Cells[recordIndex, 1].Value = (recordIndex - 1).ToString();
                workSheetVinaphone.Cells[recordIndex, 2].Value = nhanvien.HoTen;
                workSheetVinaphone.Cells[recordIndex, 3].Value = nhanvien.DienThoai;
                workSheetVinaphone.Cells[recordIndex, 4].Value = nhanvien.DiaChi;
                //number with 2 decimal places and thousand separator and money symbol

                workSheetVinaphone.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                workSheetVinaphone.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                workSheetVinaphone.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                workSheetVinaphone.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                recordIndex++;
            }
            for (int i = 1; i <= 4; i++)
            {
                workSheetVinaphone.Column(i).AutoFit();
            }
            #endregion

            #region Vietnamobile
            // Vietnamobile
            var workSheetVietnamobile = excel.Workbook.Worksheets.Add("Vietnamobile");

            workSheetVietnamobile.Row(1).Style.Font.Bold = true;
            // Assign borders
            workSheetVietnamobile.Cells[1, 1, 1, 4].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            workSheetVietnamobile.Cells[1, 1, 1, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            workSheetVietnamobile.Cells[1, 1, 1, 4].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            workSheetVietnamobile.Cells[1, 1, 1, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            workSheetVietnamobile.Cells[1, 1].Value = "STT";
            workSheetVietnamobile.Cells[1, 2].Value = "Họ tên";
            workSheetVietnamobile.Cells[1, 3].Value = "Điện thoại";
            workSheetVietnamobile.Cells[1, 4].Value = "Địa chỉ";
            //Body of table  
            //  

            recordIndex = 2;
            foreach (var nhanvien in lstVietnamobile)
            {

                workSheetVietnamobile.Cells[recordIndex, 1].Value = (recordIndex - 1).ToString();
                workSheetVietnamobile.Cells[recordIndex, 2].Value = nhanvien.HoTen;
                workSheetVietnamobile.Cells[recordIndex, 3].Value = nhanvien.DienThoai;
                workSheetVietnamobile.Cells[recordIndex, 4].Value = nhanvien.DiaChi;
                //number with 2 decimal places and thousand separator and money symbol

                workSheetVietnamobile.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                workSheetVietnamobile.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                workSheetVietnamobile.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                workSheetVietnamobile.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                recordIndex++;
            }
            for (int i = 1; i <= 4; i++)
            {
                workSheetVietnamobile.Column(i).AutoFit();
            }
            #endregion

            #region Gmobile
            // Gmobile
            var workSheetGmobile = excel.Workbook.Worksheets.Add("Gmobile");

            workSheetGmobile.Row(1).Style.Font.Bold = true;
            // Assign borders
            workSheetGmobile.Cells[1, 1, 1, 4].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            workSheetGmobile.Cells[1, 1, 1, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            workSheetGmobile.Cells[1, 1, 1, 4].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            workSheetGmobile.Cells[1, 1, 1, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            workSheetGmobile.Cells[1, 1].Value = "STT";
            workSheetGmobile.Cells[1, 2].Value = "Họ tên";
            workSheetGmobile.Cells[1, 3].Value = "Điện thoại";
            workSheetGmobile.Cells[1, 4].Value = "Địa chỉ";
            //Body of table  
            //  

            recordIndex = 2;
            foreach (var nhanvien in lstGmobile)
            {

                workSheetGmobile.Cells[recordIndex, 1].Value = (recordIndex - 1).ToString();
                workSheetGmobile.Cells[recordIndex, 2].Value = nhanvien.HoTen;
                workSheetGmobile.Cells[recordIndex, 3].Value = nhanvien.DienThoai;
                workSheetGmobile.Cells[recordIndex, 4].Value = nhanvien.DiaChi;
                //number with 2 decimal places and thousand separator and money symbol

                workSheetGmobile.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                workSheetGmobile.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                workSheetGmobile.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                workSheetGmobile.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                recordIndex++;
            }
            for (int i = 1; i <= 4; i++)
            {
                workSheetGmobile.Column(i).AutoFit();
            }
            #endregion


            #region Khac
            // Gmobile
            var workSheetKhac = excel.Workbook.Worksheets.Add("Khác");

            workSheetKhac.Row(1).Style.Font.Bold = true;
            // Assign borders
            workSheetKhac.Cells[1, 1, 1, 4].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            workSheetKhac.Cells[1, 1, 1, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            workSheetKhac.Cells[1, 1, 1, 4].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            workSheetKhac.Cells[1, 1, 1, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            workSheetKhac.Cells[1, 1].Value = "STT";
            workSheetKhac.Cells[1, 2].Value = "Họ tên";
            workSheetKhac.Cells[1, 3].Value = "Điện thoại";
            workSheetKhac.Cells[1, 4].Value = "Địa chỉ";
            //Body of table  
            //  
            var khacs = lstPhoneNumber
                .Except(lstViettel)
                .Except(lstMobifone)
                .Except(lstVinaphone)
                .Except(lstVietnamobile)
                .Except(lstGmobile);
            recordIndex = 2;
            foreach (var nhanvien in khacs)
            {

                workSheetKhac.Cells[recordIndex, 1].Value = (recordIndex - 1).ToString();
                workSheetKhac.Cells[recordIndex, 2].Value = nhanvien.HoTen;
                workSheetKhac.Cells[recordIndex, 3].Value = nhanvien.DienThoai;
                workSheetKhac.Cells[recordIndex, 4].Value = nhanvien.DiaChi;
                //number with 2 decimal places and thousand separator and money symbol

                workSheetKhac.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                workSheetKhac.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                workSheetKhac.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                workSheetKhac.Cells[recordIndex, 1, recordIndex, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                recordIndex++;
            }
            for (int i = 1; i <= 4; i++)
            {
                workSheetKhac.Column(i).AutoFit();
            }
            #endregion

            string excelName = "phone_chuongnh";
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

    }
}