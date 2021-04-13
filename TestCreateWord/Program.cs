using System;
using System.IO;
using NPOI.XWPF.UserModel;

namespace TestCreateWord
{
    class Program
    {
        static void Main(string[] args)
        {
            var templateFile = "./../../../Doc1.docx";
            var destinationFile = "output.docx";
            using (FileStream file = new FileStream(templateFile, FileMode.Open, FileAccess.Read))
            {
                XWPFDocument wordDoc = new XWPFDocument(file);

                var fs = new FileStream(destinationFile, FileMode.OpenOrCreate, FileAccess.Write);

                wordDoc.Write(fs);
                fs.Close();

                wordDoc.Close();
            }
        }
    }
}
