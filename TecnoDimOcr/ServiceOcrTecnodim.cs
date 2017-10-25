using Aspose.Pdf;
using GdPicture;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Reflection;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using TecnoDimOcr.OcrGdpicture;

namespace TecnoDimOcr
{
    partial class ServiceOcrTecnodim : ServiceBase
    {
        private System.Timers.Timer timer;
        public ServiceOcrTecnodim()
        {
            oGdPicturePDF.SetLicenseNumber("4118106456693265856441854");
            oGdPictureImaging.SetLicenseNumber("4118106456693265856441854");
            InitializeComponent();
            this.ServiceName = "ServiceOcrTecnodim";
        }
        private GdPictureImaging oGdPictureImaging = new GdPictureImaging();
        private GdPicturePDF oGdPicturePDF = new GdPicturePDF();

        public string executarOcr(string arquivo)
        {
            try
            {
                var saida = Ocr.GerarDocumentoPesquisavelPdf(oGdPictureImaging, oGdPicturePDF, arquivo, true, "por", null, null, null, null, null, 250);
                getPastaDestino(saida);
            }
            catch (Exception ex)
            {

                using (FileStream fs = File.Create("log.txt"))
                {
                    Byte[] info = new UTF8Encoding(true).GetBytes(ex.Message);
                    // Add some information to the file.
                    fs.Write(info, 0, info.Length);
                }
            }
            return "";
        }
        public void getPastaDestino(string arquivo)
        {
            string str = "";
            string str1 = Path.GetFileName(arquivo).ToString();
            string[] strArrays = Path.GetFileName(arquivo).ToString().Split(new char[] { '_' });
            if (strArrays.Length != 0)
            {
                str = strArrays[0];
            }
            string str2 = ConfigurationManager.AppSettings["PastaDestino"].ToString();
            str2 = string.Concat(str2, "\\", str);
            if (!Directory.Exists(str2))
            {
                str2 = ConfigurationManager.AppSettings["PastaDestinoNAOMAPEADO"].ToString();
            }
            File.Move(arquivo, string.Concat(str2, "\\", str1));
        }
        public void createDirectory(string diretorio)
        {

            if (!Directory.Exists(diretorio))
            {
                Directory.CreateDirectory(diretorio);
            }

        }
        public enum monitoramento
        {
            RTF, RTFEMAIL, OCR, OCREMAIL
        }

        public bool sendemail(string attached, string destinatario)
        {
            try
            {
                string smtp = ConfigurationManager.AppSettings["SMTP"].ToString();
                string USUARIO = ConfigurationManager.AppSettings["USUARIO"].ToString();
                string SENHA = ConfigurationManager.AppSettings["SENHA"].ToString();
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient(smtp);

                mail.From = new MailAddress(USUARIO);
                mail.To.Add(destinatario);
                mail.Subject = "Documento convertido";
                mail.Body = "documento convertido";
                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(attached);
                mail.Attachments.Add(attachment);
                SmtpServer.Port = 587;
                SmtpServer.UseDefaultCredentials = false;
                SmtpServer.Credentials = new System.Net.NetworkCredential(USUARIO, SENHA);
                SmtpServer.EnableSsl = true;

                SmtpServer.Send(mail);
                attachment.Dispose();
                return true;
            }
            catch (Exception ec)
            {
                return false;
            }
        }
        private static string GetCurrentDirectory()
        {
            string absolutePath = (new Uri(Assembly.GetExecutingAssembly().CodeBase)).AbsolutePath;
            string fullName = (new DirectoryInfo(Path.GetDirectoryName(absolutePath))).FullName;
            return Uri.UnescapeDataString(fullName);
        }
        public void monitorarPasta(monitoramento monitoramento, string caminho)
        {

            string AsposeWords = string.Concat(GetCurrentDirectory(), "\\Aspose\\Aspose.Words.lic");
            string AsposePdf = string.Concat(GetCurrentDirectory(), "\\Aspose\\Aspose.Pdf.lic");
            lockArquivo = true;
            switch (monitoramento)
            {
                

                case monitoramento.RTF:
                    string[] files = Directory.GetFiles(caminho);
                    for (int i = 0; i < (int)files.Length; i++)
                    {
                        string str = files[i];
                        string str1 = Ocr.GerarDocumentoPesquisavelPdf(this.oGdPictureImaging, this.oGdPicturePDF, str, true, "por", null, null, null, null, null, 250);
                        string str2 = AsposeWords;
                        using (FileStream fileStream = File.Open(str2, FileMode.Open))
                        {
                            (new Aspose.Words.License()).SetLicense(fileStream);
                        }
                        string str3 = AsposePdf;
                        using (FileStream fileStream1 = File.Open(str3, FileMode.Open))
                        {
                            (new Aspose.Pdf.License()).SetLicense(fileStream1);
                        }
                        Aspose.Pdf.Document document = new Aspose.Pdf.Document(str1);
                        document.Save(string.Concat(Path.GetDirectoryName(str1), "\\", Path.GetFileNameWithoutExtension(str1), ".docx"), SaveFormat.DocX);
                        Aspose.Words.Document document1 = new Aspose.Words.Document(string.Concat(Path.GetDirectoryName(str1), "\\", Path.GetFileNameWithoutExtension(str1), ".docx"));
                        document1.Save(string.Concat(Path.GetDirectoryName(str1), "\\", Path.GetFileNameWithoutExtension(str1), ".rtf"));
                        string str4 = ConfigurationManager.AppSettings["pastaSaidaRTF"].ToString();

                        var move = string.Concat(str4, "\\", Path.GetFileNameWithoutExtension(str) + ".rtf");
                        int count = 1;
                        while (File.Exists(move))
                        {
                            string tempFileName = string.Format("{0}({1})", (str4 + "\\" + Path.GetFileNameWithoutExtension(str)), count++);
                            move = tempFileName + ".rtf";
                        }
                        File.Move(string.Concat(Path.GetDirectoryName(str1), "\\", Path.GetFileNameWithoutExtension(str1), ".rtf"), move);
                        File.Delete(string.Concat(Path.GetDirectoryName(str1), "\\", Path.GetFileNameWithoutExtension(str1), ".docx"));
                        File.Delete(str1);
                    }
                    break;
                case monitoramento.RTFEMAIL:
                    ////captura de email
                    foreach (var item in Directory.GetFiles(caminho))
                    {

                        string str = item;
                        string str1 = Ocr.GerarDocumentoPesquisavelPdf(this.oGdPictureImaging, this.oGdPicturePDF, str, true, "por", null, null, null, null, null, 250);
                        string str2 = AsposeWords;
                        using (FileStream fileStream = File.Open(str2, FileMode.Open))
                        {
                            (new Aspose.Words.License()).SetLicense(fileStream);
                        }
                        string str3 = AsposePdf;
                        using (FileStream fileStream1 = File.Open(str3, FileMode.Open))
                        {
                            (new Aspose.Pdf.License()).SetLicense(fileStream1);
                        }
                        Aspose.Pdf.Document document = new Aspose.Pdf.Document(str1);
                        document.Save(string.Concat(Path.GetDirectoryName(str1), "\\", Path.GetFileNameWithoutExtension(str1), ".docx"), SaveFormat.DocX);
                        Aspose.Words.Document document1 = new Aspose.Words.Document(string.Concat(Path.GetDirectoryName(str1), "\\", Path.GetFileNameWithoutExtension(str1), ".docx"));
                        document1.Save(string.Concat(Path.GetDirectoryName(str1), "\\", Path.GetFileNameWithoutExtension(str1), ".rtf"));
                        string str4 = ConfigurationManager.AppSettings["pastaSaidaRTF"].ToString();

                        var move = string.Concat(str4, "\\", Path.GetFileNameWithoutExtension(str) + ".rtf");
                        int count = 1;
                        while (File.Exists(move))
                        {
                            string tempFileName = string.Format("{0}({1})", (str4 + "\\" + Path.GetFileNameWithoutExtension(str)), count++);
                            move = tempFileName + ".rtf";
                        }
                        File.Move(string.Concat(Path.GetDirectoryName(str1), "\\", Path.GetFileNameWithoutExtension(str1), ".rtf"), move);
                        File.Delete(string.Concat(Path.GetDirectoryName(str1), "\\", Path.GetFileNameWithoutExtension(str1), ".docx"));
                        File.Delete(str1);

                        string nome = Path.GetFileName(item);
                        string remetente = nome.Substring(0, nome.IndexOf('@') + 1);
                        string[] url = nome.Substring(nome.IndexOf('@') + 1).Split('_');

                        if (move.Contains(".rtf"))
                        {
                            if (sendemail(move, remetente + url[0]))
                            {
                                File.Delete(move);
                            }
                        }
                    }

                    break;
                case monitoramento.OCR:
                    foreach (var item in Directory.GetFiles(caminho))
                    {
                        string nome = Path.GetFileName(item);

                        var saida = Ocr.GerarDocumentoPesquisavelPdf(oGdPictureImaging, oGdPicturePDF, item, true, "por", null, null, null, null, null, 250);
                        if (saida.Contains(".pdf"))
                        {
                            string pastaSaidaOcr = ConfigurationManager.AppSettings["pastaSaidaOcr"].ToString();
                            if (File.Exists(pastaSaidaOcr + @"\" + Path.GetFileName(saida)))
                            {
                                File.Delete(pastaSaidaOcr + @"\" + Path.GetFileName(saida));
                            }
                            File.Move(saida, pastaSaidaOcr + @"\" + Path.GetFileName(saida));


                        }
                    }
                    break;
                case monitoramento.OCREMAIL:
                    foreach (var item in Directory.GetFiles(caminho))
                    {
                        string nome = Path.GetFileName(item);
                        string remetente = nome.Substring(0, nome.IndexOf('@') + 1);
                        string[] url = nome.Substring(nome.IndexOf('@') + 1).Split('_');
                        var saida = Ocr.GerarDocumentoPesquisavelPdf(oGdPictureImaging, oGdPicturePDF, item, true, "por", null, null, null, null, null, 300);

                        if (saida.Contains(".pdf"))
                        {
                            if (sendemail(saida, remetente + url[0]))
                            {
                                File.Delete(saida);

                            }
                        }
                    }
                    break;
                default:
                    break;
            }
            lockArquivo = false;
            var PastaDestinoTemp = ConfigurationManager.AppSettings["PastaDestinoTemp"];
            foreach (var item in Directory.GetFiles(PastaDestinoTemp))
            {
                File.Delete(item);
            }

        }
        private static bool lockArquivo = false;
        public void executar(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                if (!lockArquivo)
                {
                    string PastaRaizTOTAL = ConfigurationManager.AppSettings["PastaRaizTOTAL"].ToString();
                    var raiz = PastaRaizTOTAL + @"\";
                    createDirectory(raiz + "pastaEntradaOCR");
                    createDirectory(raiz + "pastaEntradaOCREMAIL");
                    createDirectory(raiz + "pastaEntradaRTF");
                    createDirectory(raiz + "pastaEntradaRTFEMAIL");
                    monitorarPasta(monitoramento.RTFEMAIL, raiz + "pastaEntradaRTFEMAIL");

                    monitorarPasta(monitoramento.RTF, raiz + "pastaEntradaRTF");
                    monitorarPasta(monitoramento.OCR, raiz + "pastaEntradaOCR");
                    monitorarPasta(monitoramento.OCREMAIL, raiz + "pastaEntradaOCREMAIL");
                }
            }
            catch (Exception ex)
            {
                using (FileStream fs = File.Create("c:\\log.txt"))
                {
                    Byte[] info = new UTF8Encoding(true).GetBytes(ex.Message);
                    // Add some information to the file.
                    fs.Write(info, 0, info.Length);
                }
            }

 
        }
        protected override void OnStart(string[] args)
        {

            this.timer = new System.Timers.Timer(10000);  // 30000 milliseconds = 30 seconds
            this.timer.AutoReset = true;
            this.timer.Elapsed += new System.Timers.ElapsedEventHandler(this.executar);
            this.timer.Start();
        }

        protected override void OnStop()
        {
            this.timer.Stop();
            this.timer = null;
        }




    }
}
