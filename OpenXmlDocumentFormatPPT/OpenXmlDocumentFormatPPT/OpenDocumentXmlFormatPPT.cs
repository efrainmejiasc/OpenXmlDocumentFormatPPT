using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;

namespace OpenXmlDocumentFormatPPT
{
    public class OpenDocumentXmlFormatPPT
    {
        public bool WriteOnSlide()
        {
            string filePath = System.Web.HttpContext.Current.Server.MapPath("~/TemplatePPT/Plantilla1.pptx");
            try
            {
                using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, true))
                {
                    PresentationPart presentationPart = presentationDocument.PresentationPart;
                    Presentation presentation = presentationPart.Presentation;
                    List<Shape> shapes = new List<Shape>();
                    foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
                    {
                        SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                        shapes = (from shape in slidePart.Slide.Descendants<Shape>()
                                     where IsTitleShape(shape)
                                     select shape).ToList();
                   
                    }
                    StringBuilder paragraphText = new StringBuilder();
                    var shape0 = shapes[0];

                    string paragraphSeparator = null;
                    foreach (var shape in shapes)
                    {
                        // Get the text in each paragraph in this shape.
                        foreach (var paragraph in shape.TextBody.Descendants())
                        {
                            var a = paragraph;
                            // Add a line break.
                            paragraphText.Append(paragraphSeparator);

                            foreach (var text in paragraph.Descendants())
                            {
                                var b = text;
                                paragraphText.Append(text);
                            }

                            paragraphSeparator = "\n";
                        }
                    }
                    int n = 0;

                }
                ReadWriteTxt(filePath);
            }
            catch (Exception ex)
            {
                return false;
            }

            return true;
        }

        private static bool IsTitleShape(Shape shape)
        {
            // t.InnerText = "Escribiendo en Titulo: " + DateTime.Now.ToString();
            var placeholderShape = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();
            if (placeholderShape != null && placeholderShape.Type != null && placeholderShape.Type.HasValue)
            {
                switch ((PlaceholderValues)placeholderShape.Type)
                {
                    // Any title shape.
                    case PlaceholderValues.Title:

                    // A centered title.
                    case PlaceholderValues.CenteredTitle:
                        return true;

                    default:
                        return false;
                }
            }
            return false;
        }

        private void ReadWriteTxt(string pathArchivo)
        {
            FileAttributes atr = File.GetAttributes(pathArchivo);
            File.SetAttributes(pathArchivo, atr & ~FileAttributes.ReadOnly);
        }
    }
}