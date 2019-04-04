using Autofac;
using EmailStart.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmailStart
{
    public class BusinessModule : Autofac.Module
    {
        protected override void Load(ContainerBuilder builder)
        {
            
            builder.RegisterType<HeadersRepository>().As<IHeadersRepository>().InstancePerRequest();

            base.Load(builder);
        }
    }
}