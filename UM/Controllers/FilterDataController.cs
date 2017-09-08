using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using UM.Models;

namespace UM.Controllers
{
    public class FilterDataController : Controller
    {

        public static List<PhoneNumberView> phoneNumbers;

        // GET: FilterData
        public ActionResult Index()
        {
            phoneNumbers = null;
            return View();
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
                        if (phoneNumbers.Count() > 15000)
                        {
                            return Json(new { status = false, large = true, data = "Chúng tôi không hiện thị được dữ liệu vì tệp tin quá lớn." }, JsonRequestBehavior.AllowGet);
                        }
                        return Json(new { status = true, data = phoneNumbers }, JsonRequestBehavior.AllowGet);
                    }
                    catch (Exception ex)
                    {
                        return Json(new { status = false, large = false, data = ex.Message }, JsonRequestBehavior.AllowGet);
                    }
                }
            }
            return Json(new { status = false, large = false, data = new List<PhoneNumberView>() }, JsonRequestBehavior.AllowGet);
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

                }).Distinct().ToList();


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

            string excelName = "filter_data";
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