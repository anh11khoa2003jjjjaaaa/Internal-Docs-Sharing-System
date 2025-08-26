//using System;
//using System.Collections.Generic;
//using System.Drawing;
//using System.Windows.Forms;
//using Google.Apis.Drive.v3.Data;

//namespace ChiaseNoiBo
//{
//    internal class FileListLoader
//    {
//        private Panel panel;
//        private Action<string, string> onFileClick;
//        public FileListLoader(Panel panel, Action<string, string> onFileClick)
//        {
//            this.panel = panel;
//            this.onFileClick = onFileClick;
//        }

//        public void LoadFiles()
//        {
//            panel.Controls.Clear();
//            List<File> files = GoogleDriveHelper.GetExcelFiles();
//            int y = 10;

//            foreach (var file in files)
//            {
//                Button btn = new Button
//                {
//                    Text = file.Name,
//                    Tag = file.Id,
//                    Size = new Size(panel.Width - 20, 40),
//                    Location = new Point(10, y),
//                    BackColor = Color.LightBlue
//                };
//                btn.Click += (s, e) => onFileClick(file.Id, file.Name);
//                panel.Controls.Add(btn);
//                y += 50;
//            }
//        }
//    }
//}
