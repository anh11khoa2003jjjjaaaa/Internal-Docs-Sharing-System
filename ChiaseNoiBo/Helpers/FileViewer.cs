using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using PdfiumViewer;

namespace ChiaseNoiBo.Helpers
{
    public static class FileViewer
    {
        /// <summary>
        /// Mở file Word bằng Microsoft Word (Interop)
        /// </summary>
        /// <param name="filePath">Đường dẫn file Word</param>
        public static void ViewWordFile(string filePath)
        {
            try
            {
                var wordApp = new Microsoft.Office.Interop.Word.Application();
                wordApp.Visible = true;
                wordApp.Documents.Open(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi mở file Word: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Mở file PDF bằng PdfiumViewer trong một Form mới
        /// </summary>
        /// <param name="filePath">Đường dẫn file PDF</param>
        public static void ViewPdfFile(string filePath)
        {
            try
            {
                // Tạo Form hiển thị PDF
                var pdfForm = new Form
                {
                    Text = $"Xem PDF: {Path.GetFileName(filePath)}",
                    Width = 900,
                    Height = 700,
                    StartPosition = FormStartPosition.CenterScreen
                };

                // Tạo control xem PDF
                var pdfViewer = new PdfViewer
                {
                    Dock = DockStyle.Fill,
                    Document = PdfDocument.Load(filePath)
                };

                pdfForm.Controls.Add(pdfViewer);
                pdfForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi mở file PDF: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
