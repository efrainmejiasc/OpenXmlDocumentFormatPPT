using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Picture = DocumentFormat.OpenXml.Presentation.Picture;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Xml;


namespace OpenXmlDocumentFormatPPT
{
    public class OpenDocumentXmlFormatPPT
    {

        public bool WriteOnSlide()
        {
            string filePath = System.Web.HttpContext.Current.Server.MapPath("~/TemplatePPT/Plantilla1.ppsx");
            string imagen1 = System.Web.HttpContext.Current.Server.MapPath("~/Images/zima1.jpeg");
            string imagen2 = System.Web.HttpContext.Current.Server.MapPath("~/Images/zima2.jpeg");

            string tituloText = "Este es el titulo de la diapositiva: " + DateTime.Now.ToString();
            string subTituloText = "Este es el subtitulo de la diapositiva: " + Environment.NewLine + "Prueba de contexto del manejo de presentaciones PowerPoint con el Nuget \n DOCUMENT FORMAT OPEN XML " + Environment.NewLine + DateTime.Now.ToString();

            try
            {
                using (PresentationDocument documentoPP = PresentationDocument.Open(filePath, true))
                {
                    PresentationPart partesDocumento = documentoPP.PresentationPart;
                    Presentation presentacion = partesDocumento.Presentation;
                    OpenXmlElementList diapositivas = partesDocumento.Presentation.SlideIdList.ChildElements;
                    string idDiapositiva = (diapositivas[1] as SlideId).RelationshipId;//DIAPOSITIVA ACTUAL
                    SlidePart diapositiva = (SlidePart) partesDocumento.GetPartById(idDiapositiva);
                    List<OpenXmlElement> itemsdiapositiva = diapositiva.Slide.CommonSlideData.ShapeTree.ChildElements.ToList();
                 
                    DocumentFormat.OpenXml.Presentation.Shape titulo = (DocumentFormat.OpenXml.Presentation.Shape)itemsdiapositiva[2];
                    DocumentFormat.OpenXml.Presentation.Shape subTitulo = (DocumentFormat.OpenXml.Presentation.Shape)itemsdiapositiva[3];
                   // DocumentFormat.OpenXml.Presentation.Picture imagen = (DocumentFormat.OpenXml.Presentation.Picture)itemsdiapositiva[4];
                    titulo.InnerXml = GetXmlStringBody(titulo.InnerXml, tituloText);
                    subTitulo.InnerXml = GetXmlStringBody(subTitulo.InnerXml, subTituloText);
                    byte[] nuevaImagen = GetBytesImagen(imagen2);//CAMBIAR IMAGEN
                    ReplacePicture("", nuevaImagen, "image/jpeg", diapositiva);

                    partesDocumento.Presentation.Save();
                }

                ReadWriteTxt(filePath);
            }
            catch (Exception ex)
            {
                return false;
            }

            return true;
        }


        #region CambiarTitulo&Subtitulo

        private string  SetTextShape(string innerXml,string texto)
        {
            if (innerXml.Contains("<a:t>"))
                innerXml = System.Text.RegularExpressions.Regex.Replace(innerXml, @"<a:t(.*?)>(.*?)</a:t>", "<a:t>" + texto + "</a:t>");
            else
                innerXml = GetXmlStringBody(innerXml, texto);

             return innerXml;
        }


        private string GetXmlStringBody(string innerXml,string texto)
        {
            int indice = innerXml.IndexOf("<a:p xmlns:a=");
            string innerXml_1 = innerXml.Substring(0, indice);
            innerXml = innerXml_1 + StrMarcadoTextShape();
            innerXml = SetTextShape(innerXml, texto);

            return innerXml;
        }

        private string StrMarcadoTextShape()
        {
            StringBuilder str = new StringBuilder();
            str.Append("<a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" >");
            str.Append("<a:r>");
            str.Append("<a:rPr lang=\"en-US\" sz=\"2800\" smtClean=\"0\" >");
            str.Append("<a:latin typeface=\"Arial\" panose=\"020B0604020202020204\" pitchFamily=\"34\" charset=\"0\" />");
            str.Append("</a:rPr>");
            str.Append("<a:t></a:t>");
            str.Append("</a:r>");
            str.Append("<a:endParaRPr lang=\"en-US\" sz=\"2800\" >");
            str.Append("<a:latin typeface=\"Arial\" panose=\"020B0604020202020204\" pitchFamily=\"34\" charset=\"0\" />");
            str.Append("</a:endParaRPr>");
            str.Append("</a:p>");
            str.Append("</p:txBody>");

            return str.ToString();
        }

        #endregion


        #region CambiarImagen

        private byte [] GetBytesImagen(string pathImagen)
        {
            return File.ReadAllBytes(pathImagen);
        }

        /// <summary>
        /// Replaces a picture by another inside the slide.
        /// </summary>
        /// <param name="tag">The tag associated with the original picture so it can be found, if null or empty do nothing.</param>
        /// <param name="newPicture">The new picture (as a byte array) to replace the original picture with, if null do nothing.</param>
        /// <param name="contentType">The picture content type: image/png, image/jpeg...</param>
        /// <remarks>
        /// <see href="http://stackoverflow.com/questions/7070074/how-can-i-retrieve-images-from-a-pptx-file-using-ms-open-xml-sdk">How can I retrieve images from a .pptx file using MS Open XML SDK?</see>
        /// <see href="http://stackoverflow.com/questions/7137144/how-can-i-retrieve-some-image-data-and-format-using-ms-open-xml-sdk">How can I retrieve some image data and format using MS Open XML SDK?</see>
        /// <see href="http://msdn.microsoft.com/en-us/library/office/bb497430.aspx">How to: Insert a Picture into a Word Processing Document</see>
        /// </remarks>
        private void ReplacePicture(string tag, byte[] newPicture, string contentType,SlidePart diapositiva)
        {

            ImagePart imagePart = AddPicture(newPicture, contentType,diapositiva);

            foreach (Picture pic in diapositiva.Slide.Descendants<Picture>())
            {
                var cNvPr = pic.NonVisualPictureProperties.NonVisualDrawingProperties;
                if (cNvPr.Name != null)
                {
                    string title = cNvPr.Name.Value;
                    string rId = diapositiva.GetIdOfPart(imagePart);
                    pic.BlipFill.Blip.Embed.Value = rId;

                }
            }
        }


        private ImagePart AddPicture(byte[] picture, string contentType,SlidePart diapositiva)
        {
            ImagePartType type = 0;
            switch (contentType)
            {
                case "image/bmp":
                    type = ImagePartType.Bmp;
                    break;
                case "image/emf": 
                    type = ImagePartType.Emf;
                    break;
                case "image/gif": 
                    type = ImagePartType.Gif;
                    break;
                case "image/ico": 
                    type = ImagePartType.Icon;
                    break;
                case "image/jpeg":
                    type = ImagePartType.Jpeg;
                    break;
                case "image/pcx": 
                    type = ImagePartType.Pcx;
                    break;
                case "image/png":
                    type = ImagePartType.Png;
                    break;
                case "image/tiff":
                    type = ImagePartType.Tiff;
                    break;
                case "image/wmf":
                    type = ImagePartType.Wmf;
                    break;
            }

            ImagePart imagePart = diapositiva.AddImagePart(type);

            // FeedData() closes the stream and we cannot reuse it (ObjectDisposedException)
            // solution: copy the original stream to a MemoryStream
            using (MemoryStream stream = new MemoryStream(picture))
            {
                imagePart.FeedData(stream);
            }

            return imagePart;
        }

        private string GetIdOfImagePart(ImagePart imagePart , SlidePart diapositiva)
        {
            return diapositiva.GetIdOfPart(imagePart);
        }

        private void Save(SlidePart diapositiva)
        {
            diapositiva.Slide.Save();
        }

        #endregion

        private void ReadWriteTxt(string pathArchivo)
        {
            FileAttributes atr = File.GetAttributes(pathArchivo);
            File.SetAttributes(pathArchivo, atr & ~FileAttributes.ReadOnly);
        }

    }
}