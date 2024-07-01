using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace WebApplication2
{
    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);


            //لایسنس aspose
            new Aspose.Words.License().SetLicense(LicenseHelper.License.LStreamAsposeKey);
            new Aspose.Cells.License().SetLicense(LicenseHelper.License.LStreamCellAsposeKey);
        }
    }
}
