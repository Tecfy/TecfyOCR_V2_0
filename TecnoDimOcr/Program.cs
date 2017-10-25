using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using TecnoDimOcr.OcrGdpicture;

namespace TecnoDimOcr
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {

            try
            {
                //new ServiceOcrTecnodim().executar(null, null);
                ServiceBase.Run(new ServiceBase[] { new ServiceOcrTecnodim() });

            }
            catch (Exception ex)
            {
                using (FileStream fs = File.Create("c:\\log.txt"))
                {
                    Byte[] info = new UTF8Encoding(true).GetBytes(ex.Message);
                    // Add some information to the file.
                    fs.Write(info, 0, info.Length);
                }
                //  Console.WriteLine(ex.Message);

            }


        }
    }
}
