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




        private string  GetXmlStringBody(string innerXml,string texto)
        {
            if (innerXml.Contains("<a:t>"))
             innerXml = System.Text.RegularExpressions.Regex.Replace(innerXml, @"<a:t(.*?)>(.*?)</a:t>", "<a:t>" + texto + "</a:t>");

             return innerXml;
        }


        public Paragraph ReplaceText(Paragraph paragraph)
        {

          var parent = paragraph.Parent; //get parent element - to be used when removing placeholder
            var dataParam = new PowerPointParameter()
            {
                Name = "Titulo Set",
                Text = "Texto para escribir"
            };

      
                //insert text list
                if (dataParam.Name.Contains("string[]")) //check if param is a list
                {
                    var arrayText = dataParam.Text.Split(Environment.NewLine.ToCharArray()); //in our case we split it into lines

                    if (arrayText is IEnumerable) //enumerate if we can
                    {
                        foreach (var itemData in arrayText)
                        {
                            Paragraph bullet = CloneParaGraphWithStyles(paragraph, dataParam.Name, itemData);// create new param - preserve styles
                            parent.InsertBefore(bullet, paragraph); //insert new element
                        }
                    }
                    paragraph.Remove();//delete placeholder
                }
                else
                {
                    //insert text line
                    var param = CloneParaGraphWithStyles(paragraph, dataParam.Name, dataParam.Text); // create new param - preserve styles
                    parent.InsertBefore(param, paragraph);//insert new element

                    paragraph.Remove();//delete placeholder
                }

            return paragraph;
        }

        public class PowerPointParameter
        {
            public string Name { get; set; }
            public string Text { get; set; }
            public FileInfo Image { get; set; }
        }

        public static Paragraph CloneParaGraphWithStyles(Paragraph sourceParagraph, string paramKey, string text)
        {
            var xmlSource = sourceParagraph.OuterXml;

            xmlSource = xmlSource.Replace(paramKey.Trim(), text.Trim());

            return new Paragraph(xmlSource);
        }


        private void ReadWriteTxt(string pathArchivo)
        {
            FileAttributes atr = File.GetAttributes(pathArchivo);
            File.SetAttributes(pathArchivo, atr & ~FileAttributes.ReadOnly);
        }
    }
}