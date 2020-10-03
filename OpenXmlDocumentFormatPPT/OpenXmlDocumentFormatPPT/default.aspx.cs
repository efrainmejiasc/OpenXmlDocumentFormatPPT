﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace OpenXmlDocumentFormatPPT
{
    public partial class _default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            OpenDocumentXmlFormatPPT obj = new OpenDocumentXmlFormatPPT();
            var result = obj.WriteOnSlide();
            if (result)
                Label1.Text = "Escrituta en titulo y subtitulo de la diapositiva (0) exitosa";
            else
                Label1.Text = "Escrituta en titulo y subtitulo de la diapositiva (0) fallida";

            Label2.Text = "NUGET: DOCUMENT FORMAT OPEN XML";
        }
    }
}