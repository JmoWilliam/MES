using System;
using System.Collections.Generic;
using System.Web;
using System.Web.Mvc;
using System.Web.WebPages;

namespace Business_Manager.Helpers
{
    public static partial class HtmlRenderHelper
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="HtmlHelper"></param>
        /// <param name="Type"></param>
        /// <param name="Template"></param>
        /// <returns></returns>
        public static IHtmlString Resource(this HtmlHelper HtmlHelper, string Type, Func<object, HelperResult> Template)
        {
            if (HtmlHelper.ViewContext.HttpContext.Items[Type] != null) ((List<Func<object, HelperResult>>)HtmlHelper.ViewContext.HttpContext.Items[Type]).Add(Template);
            else HtmlHelper.ViewContext.HttpContext.Items[Type] = new List<Func<object, HelperResult>>() { Template };

            return new HtmlString(string.Empty);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="HtmlHelper"></param>
        /// <param name="Type"></param>
        /// <returns></returns>
        public static IHtmlString RenderResources(this HtmlHelper HtmlHelper, string Type)
        {
            if (HtmlHelper.ViewContext.HttpContext.Items[Type] != null)
            {
                List<Func<object, HelperResult>> Resources = (List<Func<object, HelperResult>>)HtmlHelper.ViewContext.HttpContext.Items[Type];

                foreach (var Resource in Resources)
                {
                    if (Resource != null) HtmlHelper.ViewContext.Writer.Write(Resource(null));
                }
            }

            return new HtmlString(string.Empty);
        }
    }
}