using Autofac;
using Autofac.Integration.Mvc;
using System.Web.Mvc;

namespace EmailStart.App_Start
{
    public class AutofacConfig
    {
        public static void ConfigureContainer()
        {
            var builder = new ContainerBuilder();
            builder.RegisterControllers(typeof(Global).Assembly);
            builder.RegisterFilterProvider();
            builder.RegisterSource(new ViewRegistrationSource());
            builder.RegisterModule(new BusinessModule());
            var container = builder.Build();

            DependencyResolver.SetResolver(new AutofacDependencyResolver(container));
        }
    }
}