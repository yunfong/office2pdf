using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PDFUtil
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2) 
            {
                Console.WriteLine("\n使用前请确保安装了MS Office Word，并具备另存为PDF功能。\n建议最低版本为MS Office Word 2010。\n 使用方法：\n PDFUtil.exe   D:\\path\\myfile.docx   D:\\path\\myfile.pdf");
                Console.ReadKey();
                return;
            }
            Convert ct = new Convert(args[0], args[1]);
            string failureString = ct.start();
            Console.Write(failureString);
            
        }
    }
}
