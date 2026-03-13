using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace BatDongSan.Helper
{
    public class UploadHelper
    {
        public string PathSave { get; set; }
        public string FileName { get; set; }
        public string FileSave { get; set; }

        public void UploadFile(HttpPostedFileBase file)
        {
            if (file.ContentLength > 0)
            {
                //OK
                string dateNow = DateTime.Now.ToString("yyyyMMddhhmmss");
                FileName = Path.GetFileName(file.FileName);
                FileSave = Path.GetFileNameWithoutExtension(file.FileName).Replace(" ", "_") + "_" + dateNow + Path.GetExtension(file.FileName);

                if (!Directory.Exists(PathSave))
                {
                    Directory.CreateDirectory(PathSave);
                }
                file.SaveAs(Path.Combine(PathSave, FileSave));
            }
        }

        public void DeleteFileOfServer(string filePath)
        {
            FileInfo file = new FileInfo(filePath);
            file.Refresh();
            if (file.Exists)
                file.Delete();
        }
    }
}