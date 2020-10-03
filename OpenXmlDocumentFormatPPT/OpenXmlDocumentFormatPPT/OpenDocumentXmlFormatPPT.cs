using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
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
            string filePath = System.Web.HttpContext.Current.Server.MapPath("~/TemplatePPT/Plantilla1.pptx");
            string tituloText = "Este es el titulo de la diapositiva: " + DateTime.Now.ToString();
            string subTituloText = "Este es el subtitulo de la diapositiva: " + DateTime.Now.ToString();

            try
            {
                using (PresentationDocument documentoPP = PresentationDocument.Open(filePath, true))
                {
                    PresentationPart partesDocumento = documentoPP.PresentationPart;
                    Presentation presentacion = partesDocumento.Presentation;
                    OpenXmlElementList diapositivas = partesDocumento.Presentation.SlideIdList.ChildElements;
                    string idDiapositiva = (diapositivas[0] as SlideId).RelationshipId;
                    SlidePart diapositiva = (SlidePart) partesDocumento.GetPartById(idDiapositiva);
                    List<OpenXmlElement> itemsdiapositiva = diapositiva.Slide.CommonSlideData.ShapeTree.ChildElements.ToList();
                    DocumentFormat.OpenXml.Presentation.Shape titulo = (DocumentFormat.OpenXml.Presentation.Shape)itemsdiapositiva[2];
                    DocumentFormat.OpenXml.Presentation.Shape subTitulo = (DocumentFormat.OpenXml.Presentation.Shape)itemsdiapositiva[3];
                    titulo.InnerXml = GetXmlStringBody(titulo.InnerXml, tituloText);
                    subTitulo.InnerXml = GetXmlStringBody(subTitulo.InnerXml, subTituloText);
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
            str.Append ("<a:rPr lang=\"en-US\" sz=\"2800\" smtClean=\"0\" >");
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


        private void ReadWriteTxt(string pathArchivo)
        {
            FileAttributes atr = File.GetAttributes(pathArchivo);
            File.SetAttributes(pathArchivo, atr & ~FileAttributes.ReadOnly);
        }
    }
}