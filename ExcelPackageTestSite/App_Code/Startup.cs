using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ExcelPackageTestSite.Startup))]
namespace ExcelPackageTestSite
{
    public partial class Startup {
        public void Configuration(IAppBuilder app) {
            ConfigureAuth(app);
        }
    }
}
