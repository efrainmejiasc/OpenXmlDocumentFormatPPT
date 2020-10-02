using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
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
                    string titulo = itemsdiapositiva[2].InnerText;
                    string subTitulo = itemsdiapositiva[3].InnerText;
                    itemsdiapositiva[2].InnerXml = itemsdiapositiva[2].InnerXml.Replace(titulo, "Este es el titulo de la diapositiva: " + DateTime.Now.ToString());
                    itemsdiapositiva[3].InnerXml = itemsdiapositiva[3].InnerXml.Replace(subTitulo, "Este es el subtitulo de la diapositiva: " + DateTime.Now.ToString());
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


        private void ReadWriteTxt(string pathArchivo)
        {
            FileAttributes atr = File.GetAttributes(pathArchivo);
            File.SetAttributes(pathArchivo, atr & ~FileAttributes.ReadOnly);
        }
    }
}