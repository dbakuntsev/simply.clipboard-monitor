using Microsoft.Extensions.DependencyInjection;
using Simply.ClipboardMonitor.Services;
using Simply.ClipboardMonitor.Services.Impl;
using Simply.ClipboardMonitor.Services.Impl.Strategies;
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
        // Handle-read strategies — injected as IEnumerable<IHandleReadStrategy> into ClipboardReaderService.
        services.AddSingleton<IHandleReadStrategy, NoneHandleReadStrategy>();
        services.AddSingleton<IHandleReadStrategy, HGlobalHandleReadStrategy>();
        services.AddSingleton<IHandleReadStrategy, HBitmapHandleReadStrategy>();
        services.AddSingleton<IHandleReadStrategy, HEnhMetaFileHandleReadStrategy>();

        // Handle-write strategies — injected as IEnumerable<IHandleWriteStrategy> into ClipboardWriterService.
        services.AddSingleton<IHandleWriteStrategy, HGlobalHandleWriteStrategy>();
        services.AddSingleton<IHandleWriteStrategy, HBitmapHandleWriteStrategy>();
        services.AddSingleton<IHandleWriteStrategy, HEnhMetaFileHandleWriteStrategy>();

        // Format exporters — injected as IEnumerable<IFormatExporter> into MainWindow.
        // Order determines the filter list position in the Save dialog.
        services.AddSingleton<IFormatExporter, TextFormatExporter>();
        services.AddSingleton<IFormatExporter, PngFormatExporter>();
        services.AddSingleton<IFormatExporter, JpegFormatExporter>();
        services.AddSingleton<IFormatExporter, BinaryFormatExporter>();

        // Domain services — all stateless or cheaply shared; registered as singletons.
        services.AddSingleton<IClipboardReader,         ClipboardReaderService>();
        services.AddSingleton<IClipboardWriter,         ClipboardWriterService>();
        services.AddSingleton<ITextDecodingService,     TextDecodingService>();
        services.AddSingleton<IImagePreviewService,     ImagePreviewService>();
        services.AddSingleton<IFormatClassifier,        FormatClassifierService>();
        services.AddSingleton<IPreferencesService,      PreferencesService>();
        services.AddSingleton<IClipboardFileRepository, ClipboardFileRepository>();

        // HistoryRepository implements both IHistoryRepository and IHistoryMaintenance;
        // register the concrete class once and alias both interfaces to the same instance.
        services.AddSingleton<HistoryRepository>();
        services.AddSingleton<IHistoryRepository>(sp  => sp.GetRequiredService<HistoryRepository>());
        services.AddSingleton<IHistoryMaintenance>(sp => sp.GetRequiredService<HistoryRepository>());

        // Main window — singleton because only one instance is ever shown.
        services.AddSingleton<MainWindow>();
    }
}
