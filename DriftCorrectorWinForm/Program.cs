using AsposeFormAdjustment;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace DriftCorrectorWinForm
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();

            // 1. Build the Configuration directly (Zero background hosting baggage)
            IConfiguration configuration = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            // 2. Set up Dependency Injection directly
            var services = new ServiceCollection();

            // Register the configuration so it can be injected
            services.AddSingleton<IConfiguration>(configuration);

            // Register your form
            services.AddTransient<TestForm>();

            // 3. Build the Service Provider
            using (ServiceProvider serviceProvider = services.BuildServiceProvider())
            {
                // 4. Resolve the form (with IConfiguration injected) and run
                var mainForm = serviceProvider.GetRequiredService<TestForm>();
                Application.Run(mainForm);
            }
        }
    }
}
