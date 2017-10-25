using GdPicture;
using System;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace TecnoDimOcr.OcrGdpicture
{
    public class Ocr
    {
        public Ocr()
        {
        }

        public static string castTopdf(string file, GdPictureImaging oGdPictureImaging, GdPicturePDF oGdPicturePDF)
        {
            string str = "";
            oGdPictureImaging.TiffOpenMultiPageForWrite(false);
            int num = oGdPictureImaging.CreateGdPictureImageFromFile(file);
            if (num != 0)
            {
                oGdPicturePDF.NewPDF();
                if (oGdPictureImaging.TiffIsMultiPage(num))
                {
                    int num1 = oGdPictureImaging.TiffGetPageCount(num);
                    bool flag = true;
                    int num2 = 1;
                    while (num2 <= num1)
                    {
                        oGdPictureImaging.TiffSelectPage(num, num2);
                        oGdPicturePDF.AddImageFromGdPictureImage(num, false, true);
                        if (oGdPicturePDF.GetStat() == 0)
                        {
                            num2++;
                        }
                        else
                        {
                            flag = false;
                            break;
                        }
                    }
                    if (flag)
                    {
                        str = file.Replace(Path.GetExtension(file), ".pdf");
                        oGdPicturePDF.SaveToFile(file.Replace(Path.GetExtension(file), ".pdf"));
                        if (oGdPicturePDF.GetStat() != 0)
                        {
                        }
                    }
                    oGdPicturePDF.CloseDocument();
                    oGdPictureImaging.ReleaseGdPictureImage(num);
                }
                else
                {
                    oGdPicturePDF.AddImageFromGdPictureImage(num, false, true);
                    if (oGdPicturePDF.GetStat() == 0)
                    {
                        str = file.Replace(Path.GetExtension(file), ".pdf");
                        if (oGdPicturePDF.SaveToFile(file.Replace(Path.GetExtension(file), ".pdf")) != 0)
                        {
                        }
                    }
                    oGdPicturePDF.CloseDocument();
                    oGdPictureImaging.ReleaseGdPictureImage(num);
                }
            }
            File.Delete(file);
            return str;
        }

        public static string GerarDocumentoPesquisavelPdf(GdPictureImaging _gdPictureImaging, GdPicturePDF _gdPicturePDF, string documento, bool pdfa = true, string idioma = "por", string titulo = null, string autor = null, string assunto = null, string palavrasChaves = null, string criador = null, int dpi = 250)
        {
            if (Path.GetExtension(documento) != ".pdf")
            {
                documento = Ocr.castTopdf(documento, _gdPictureImaging, _gdPicturePDF);
            }
            int num = 0;
            var pasta = Guid.NewGuid().ToString();
            _gdPicturePDF.LoadFromFile(documento, true);
            string str = string.Concat(Ocr.GetCurrentDirectory(), "\\GdPicture\\Idiomas");


            using (FileStream fs = File.Create("c:\\lodg.txt"))
            {
                Byte[] info = new UTF8Encoding(true).GetBytes(str);
                // Add some information to the file.
                fs.Write(info, 0, info.Length);
            }
            //  Console.WriteLine(ex.Message);

            string str1 = ConfigurationManager.AppSettings["PastaDestinoTemp"].ToString();
            string str2 = string.Concat(str1, "\\", Path.GetFileName(documento));
            string folder = Guid.NewGuid().ToString();
            int pageCount = _gdPicturePDF.GetPageCount();
            for (int i = 1; i <= pageCount; i++)
            {
                Directory.CreateDirectory(str1 + "\\" + pasta);
                _gdPicturePDF.SelectPage(i);
                int gdPictureImageEx = _gdPicturePDF.RenderPageToGdPictureImageEx((float)dpi, true);
                if (gdPictureImageEx != 0)
                {
                    num = _gdPictureImaging.PdfOCRStart(str1 + "\\" + pasta + "\\" + i.ToString() + ".pdf", pdfa, titulo, autor, assunto, palavrasChaves, criador);
                    _gdPictureImaging.PdfAddGdPictureImageToPdfOCR(num, gdPictureImageEx, idioma, str, "");
                    _gdPictureImaging.ReleaseGdPictureImage(gdPictureImageEx);
                    _gdPictureImaging.PdfOCRStop(num);
                }
            }

            _gdPicturePDF.CloseDocument();
            File.Delete(documento);

            GdPictureStatus status = _gdPicturePDF.MergeDocuments(Directory.GetFiles(str1 + "\\" + pasta), str2);
          

            DirectoryInfo dir = new DirectoryInfo(str1 + "\\" + pasta);

            foreach (FileInfo fi in dir.GetFiles())
            {
                fi.Delete();
            }

            Directory.Delete(str1 + "\\" + pasta);
            return str2;
        }

        private static string GetCurrentDirectory()
        {
            string absolutePath = (new Uri(Assembly.GetExecutingAssembly().CodeBase)).AbsolutePath;
            string fullName = (new DirectoryInfo(Path.GetDirectoryName(absolutePath))).FullName;
            return Uri.UnescapeDataString(fullName);
        }

        public static bool IsFileLocked(string filePath)
        {
            bool flag;
            try
            {
                using (FileStream fileStream = File.Open(filePath, FileMode.Open))
                {
                }
            }
            catch (IOException oException)
            {
                int hRForException = Marshal.GetHRForException(oException) & 65535;
                flag = (hRForException == 32 ? true : hRForException == 33);
                return flag;
            }
            flag = false;
            return flag;
        }
    }
}