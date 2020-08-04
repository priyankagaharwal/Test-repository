using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ExcelImport.Startup))]
namespace ExcelImport
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            //ConfigureAuth(app);
        }
    }
}
