using Microsoft.Extensions.DependencyInjection;
using Simply.ClipboardMonitor.Services;
using Simply.ClipboardMonitor.Services.Impl;
using System.Windows;

namespace Simply.ClipboardMonitor;

/// <summary>
/// Application entry point.  Builds the DI container, then creates and shows
/// <see cref="MainWindow"/> through it so that constructor injection works.
/// </summary>
public partial class App : Application
{
    private ServiceProvider? _serviceProvider;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        var services = new ServiceCollection();
        RegisterServices(services);
        _serviceProvider = services.BuildServiceProvider();

        var mainWindow = _serviceProvider.GetRequiredService<MainWindow>();
        mainWindow.Show();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _serviceProvider?.Dispose();
        base.OnExit(e);
    }

    // ── Service registrations ────────────────────────────────────────────────

    private static void RegisterServices(IServiceCollection services)
    {
        // Domain services — all stateless or cheaply shared; registered as singletons.
        services.AddSingleton<IClipboardReader,         ClipboardReaderService>();
        services.AddSingleton<IClipboardWriter,         ClipboardWriterService>();
        services.AddSingleton<ITextDecodingService,     TextDecodingService>();
        services.AddSingleton<IImagePreviewService,     ImagePreviewService>();
        services.AddSingleton<IFormatClassifier,        FormatClassifierService>();
        services.AddSingleton<IPreferencesService,      PreferencesService>();
        services.AddSingleton<IHistoryRepository,       HistoryRepository>();
        services.AddSingleton<IClipboardFileRepository, ClipboardFileRepository>();

        // Main window — singleton because only one instance is ever shown.
        services.AddSingleton<MainWindow>();
    }
}
