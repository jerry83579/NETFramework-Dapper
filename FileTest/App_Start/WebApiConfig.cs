using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;

namespace FileTest
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            // 啟用網域存取
            config.EnableCors();
            // Web API 設定和服務
            GlobalConfiguration.Configuration.Formatters.XmlFormatter.SupportedMediaTypes.Clear();
            // Web API 路由
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );
        }
    }
}
