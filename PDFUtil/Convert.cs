using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace PDFUtil
{
    class Convert
    {
        public enum DocType { 
            DOC,
            XLS,
            PPT
        }
        private FileInfo sourceFile, targetFile;
        

        public Convert(string source,string target) 
        {
            sourceFile = new FileInfo(source);
            targetFile = new FileInfo(target);

            CheckTrueFileName(sourceFile.FullName);
        }

        public string start() 
        {
            if (sourceFile.Exists)
            {
                string exceptionMsg = string.Empty;
                string ext = sourceFile.Extension.ToLower();
                if (ext.Contains("doc") || ext.Contains("wps"))
                {
                    Office2PDF.Converter.WordToPdf(sourceFile.FullName, targetFile.FullName.Replace(targetFile.Extension, ".pdf"), out exceptionMsg);
                }
                else if (ext.Contains("xls") || ext.Contains("et"))
                {
                    Office2PDF.Converter.ExcelToPdf(sourceFile.FullName, targetFile.FullName.Replace(targetFile.Extension, ".pdf"), out exceptionMsg);
                }
                else if (ext.Contains("ppt") || ext.Contains("dps") || ext.Contains("pps"))
                {
                    Office2PDF.Converter.PowerPointToPdf(sourceFile.FullName, targetFile.FullName.Replace(targetFile.Extension, ".pdf"), out exceptionMsg);
                }
                else {
                    Office2PDF.Converter.WordToPdf(sourceFile.FullName, targetFile.FullName.Replace(targetFile.Extension, ".pdf"), out exceptionMsg);
                }
                return exceptionMsg;
            }
            else
            {
                return "File Not Found.";
            }
        }

        public static string CheckTrueFileName(string path)
        {
            System.IO.FileStream fs = new System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            System.IO.BinaryReader r = new System.IO.BinaryReader(fs);
            string bx = "";
            byte buffer;
            try
            {
                buffer = r.ReadByte();
                bx = buffer.ToString();
                buffer = r.ReadByte();
                bx += buffer.ToString();
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc.Message);
                return "-1";
            }
            r.Close();
            fs.Close();
            return bx;
        }

    }
}
