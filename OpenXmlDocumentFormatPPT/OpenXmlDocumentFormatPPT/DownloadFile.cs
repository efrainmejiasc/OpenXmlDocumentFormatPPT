using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace OpenXmlDocumentFormatPPT
{
    public class DownloadFile
    {
        public bool BufferedFileDownload(HttpResponse Response)
        {
            bool resultado = false;
            string filePath = System.Web.HttpContext.Current.Server.MapPath("~/TemplatePPT/Plantilla1.ppsx");
            string nombrePresentacion = "Plantilla1.ppsx";
            FileInfo file = new FileInfo(filePath);
            FileStream fromFile = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            BinaryReader binReader = new BinaryReader(fromFile);

            try
            {
                Response.Clear();
                Response.AddHeader("Content-Disposition", "attachment; filename =" + nombrePresentacion);
                Response.AddHeader("Content-Length", file.Length.ToString());
                Response.ContentType = "application/octet-stream";

                // Buffered download in blocks of 4096 bytes
                int bufsz = 0;      // size of buffer
                long block = 1;     // block of bytes
                long blocks = 0;    // number of blocks of bytes
                int rest = 0;       // number of bytes after the final block of bytes

                if (file.Length < 4096)
                {
                    bufsz = (int)file.Length;
                    blocks = 1;
                }
                else
                {
                    bufsz = 4096;
                    blocks = file.Length / bufsz;
                    rest = (int)file.Length % bufsz;
                }

                while ((block <= blocks) && (Response.IsClientConnected))
                {
                    Response.BinaryWrite(binReader.ReadBytes(bufsz));
                    Response.Flush();
                    block++;
                };

                Response.BinaryWrite(binReader.ReadBytes(rest));

                if (block * bufsz + rest == file.Length)
                    resultado = true;
            }
            catch (Exception ex)
            {
                string error = ex.ToString();
            }

            Response.End();
            binReader.Close();
            fromFile.Close();
            return resultado;

        }
    }
}