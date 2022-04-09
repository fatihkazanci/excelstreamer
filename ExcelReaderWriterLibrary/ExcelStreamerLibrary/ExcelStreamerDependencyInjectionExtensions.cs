using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelStreamerLibrary
{
    public static class ExcelStreamerDependencyInjectionExtensions
    {
        public static IServiceCollection AddExcelStreamer(this IServiceCollection services, ServiceLifetime serviceLifetime = ServiceLifetime.Scoped)
        {
            switch (serviceLifetime)
            {
                case ServiceLifetime.Scoped:
                    services.AddScoped<ExcelStreamer>();
                    break;
                case ServiceLifetime.Transient:
                    services.AddTransient<ExcelStreamer>();
                    break;
                case ServiceLifetime.Singleton:
                    services.AddSingleton<ExcelStreamer>();
                    break;
            }
            return services;
        }
    }
}
