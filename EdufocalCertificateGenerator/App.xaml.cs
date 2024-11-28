
using System.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System.Windows;
using EdufocalCertificateGenerator.ViewModels;
using NISInspectorApp.Core;

namespace EdufocalCertificateGenerator;

public partial class App : Application
{
    private readonly ServiceProvider _serviceProvider;

    public App()
    {
        IServiceCollection services = new ServiceCollection();

        // Register configuration
        var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        services.AddSingleton(config);

        // Window Views
        services.AddSingleton<MainWindow>();

        // View Models
        services.AddSingleton<MainViewModel>();

        // Services
        services.AddSingleton<Func<Type, ViewModel>>(serviceProvider =>
            viewModelType => (ViewModel)serviceProvider.GetRequiredService(viewModelType));

        _serviceProvider = services.BuildServiceProvider();
        Services = _serviceProvider;
    }

    public static ServiceProvider Services { get; private set; }

    protected override void OnStartup(StartupEventArgs e)
    {
        var mainWindow = _serviceProvider.GetRequiredService<MainWindow>();
        mainWindow.Show();
    }
}