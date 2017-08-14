using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pdf2Docx
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] files = GetFiles(args);

            foreach(var file in files)
            {
                if (Path.GetExtension(file).ToLower() == ".pdf")
                {
                    string pdfFile = file;
                    string wordFile = Path.GetDirectoryName(file) + "\\" + Path.GetFileNameWithoutExtension(file) + ".docx";

                    SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();
                    f.OpenPdf(pdfFile);
                    if (f.PageCount > 0)
                    {
                        f.WordOptions.Format = SautinSoft.PdfFocus.CWordOptions.eWordDocument.Docx;
                        int result = f.ToWord(wordFile);
                    }
                }
            }
        }

        private static string[] GetFiles(string[] args)
        {
            string folder = String.Empty;
            bool folderOk = false;
            string[] files = new string[]{};
            if (args.Length == 0)
            {
                Console.WriteLine("Please enter a folder location:");
                folder = Console.ReadLine();
            }

            while (!folderOk)
            {
                try
                {
                    files = Directory.GetFiles(folder);
                    folderOk = true;
                }
                catch
                {
                    Console.WriteLine("Please enter a valid folder location:");
                    folder = Console.ReadLine();
                }
            }
            return files;
        }
    }
}
