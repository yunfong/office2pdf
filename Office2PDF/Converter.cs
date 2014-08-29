using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using WordApplicationClass = Microsoft.Office.Interop.Word.ApplicationClass;
using ExcelApplicationClass = Microsoft.Office.Interop.Excel.ApplicationClass;
using PPApplicationClass = Microsoft.Office.Interop.PowerPoint.ApplicationClass;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace Office2PDF
{
    public class Converter
    {
        /**/
        /// <summary>
        /// 把Word文件转换成pdf文件
        /// </summary>
        /// <param name="sourcePath">需要转换的文件路径和文件名称</param>
        /// <param name="targetPath">转换完成后的文件的路径和文件名名称</param>
        /// <returns>成功返回true,失败返回false</returns>
        public static bool WordToPdf(object sourcePath, string targetPath, out string exmsg)
        {

            exmsg = string.Empty;
            bool result = false;
            WdExportFormat wdExportFormatPDF = WdExportFormat.wdExportFormatPDF;
            object missing = Type.Missing;
            WordApplicationClass applicationClass = null;
            Document document = null;
            try
            {
                applicationClass = new WordApplicationClass();
                applicationClass.Visible = false;
                document = applicationClass.Documents.Open(sourcePath);
                if (document != null)
                {
                    //document.Unprotect();
                    //document.SaveAs();
                    document.ExportAsFixedFormat(targetPath, wdExportFormatPDF);
                }
                result = true;
            }
            catch (Exception ex)
            {
                exmsg = ex.Message;
                result = false;
            }
            finally
            {
                if (document != null)
                {
                    document.Close(ref missing, ref missing, ref missing);
                    document = null;
                }
                if (applicationClass != null)
                {
                    applicationClass.Quit(ref missing, ref missing, ref missing);
                    applicationClass = null;
                }
            }
            return result;
        }

        /**/
        /// <summary>
        /// 把Excel文件转换成pdf文件
        /// </summary>
        /// <param name="sourcePath">需要转换的文件路径和文件名称</param>
        /// <param name="targetPath">转换完成后的文件的路径和文件名名称</param>
        /// <returns>成功返回true,失败返回false</returns>
        public static bool ExcelToPdf(string sourcePath, string targetPath, out string exmsg)
        {

            exmsg = string.Empty;
            bool result = false;
            object missing = Type.Missing;

            ExcelApplicationClass applicationClass = null;
            Workbook workBook = null;
            try
            {
                applicationClass = new ExcelApplicationClass();
                applicationClass.Visible = false;
                workBook = applicationClass.Workbooks.Open(sourcePath);
                if (workBook != null)
                {
                    //document.Unprotect();
                    //document.SaveAs();
                    workBook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, targetPath, XlFixedFormatQuality.xlQualityMinimum, false, true);
                    //workBook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, targetPath, XlFixedFormatQuality.xlQualityMinimum, true, false, missing, missing, missing, missing);
                }
                result = true;
            }
            catch (Exception ex)
            {
                exmsg = ex.Message;
                result = false;
            }
            finally
            {
                if (workBook != null)
                {
                    workBook.Close(true,missing, missing);
                    workBook = null;
                }
                if (applicationClass != null)
                {
                    applicationClass.Quit();
                    applicationClass = null;
                }
            }
            return result;
        }

        public static bool PowerPointToPdf(string sourcePath, string targetPath, out string exmsg)
        {
            exmsg = string.Empty;
            bool result;
            object missing = Type.Missing;
            PPApplicationClass application = null;
            Presentation persentation = null;
            try
            {
                application = new PPApplicationClass();
                persentation = application.Presentations.Open(sourcePath);
                persentation.SaveAs(targetPath, PpSaveAsFileType.ppSaveAsPDF);

                result = true;
            }
            catch (Exception ex)
            {
                exmsg = ex.Message;
                result = false;
            }
            finally
            {
                if (persentation != null)
                {
                    persentation.Close();
                    persentation = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
             }
            return result;
        }
    }
}
