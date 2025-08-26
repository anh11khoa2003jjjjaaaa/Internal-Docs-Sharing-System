//using OfficeOpenXml;
//using System;
//using System.Data;
//using System.IO;
//using System.Net;
//using System.Threading.Tasks;
//using System.Windows.Forms;

//namespace ChiaseNoiBo
//{
//    internal class ReadExcel
//    {
//        public async Task<DataTable> ReadExcelFromUrlAsync(string url)
//        {
//            try
//            {
//                using (WebClient client = new WebClient())
//                {
//                    byte[] fileData = await client.DownloadDataTaskAsync(url);
//                    using (MemoryStream stream = new MemoryStream(fileData))
//                    {
//                        return ConvertWorksheetToDataTable(new ExcelPackage(stream).Workbook.Worksheets[0]);
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                MessageBox.Show($"Lỗi đọc file Excel: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
//                return null;
//            }
//        }

//        public static DataTable ConvertWorksheetToDataTable(ExcelWorksheet worksheet)
//        {
//            DataTable dt = new DataTable();

//            if (worksheet?.Dimension == null)
//            {
//                MessageBox.Show("File Excel trống hoặc không hợp lệ.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
//                return null;
//            }

//            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
//            {
//                dt.Columns.Add(worksheet.Cells[1, col].Text);
//            }

//            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
//            {
//                DataRow newRow = dt.NewRow();
//                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
//                {
//                    newRow[col - 1] = worksheet.Cells[row, col].Text;
//                }
//                dt.Rows.Add(newRow);
//            }

//            return dt;
//        }
//    }
//}
