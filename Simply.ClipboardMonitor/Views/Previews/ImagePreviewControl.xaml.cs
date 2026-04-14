using Simply.ClipboardMonitor.Services;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace Simply.ClipboardMonitor.Views.Previews;

public sealed partial class ImagePreviewControl : UserControl, IPreviewTab
{
    private readonly IImagePreviewService _imagePreviews;

    private double _fitScale = 1.0;
    private bool   _ignoreZoomChanges;

    // Pan state
    private bool   _isPanning;
    private Point  _panStartMouse;
    private double _panHorizontalOffset;
    private double _panVerticalOffset;

    public TabItem TabItem { get; }
    public int Priority => 1;

    /// <summary>The image currently shown in the preview, or <see langword="null"/> if none.</summary>
    public BitmapSource? PreviewImageSource => ImagePreview.Source as BitmapSource;

    public ImagePreviewControl(IImagePreviewService imagePreviews)
    {
        _imagePreviews = imagePreviews;
        InitializeComponent();
        TabItem = new TabItem { Header = "Image", Content = this };
    }

    // ── IPreviewTab implementation ───────────────────────────────────────────

    public void Update(uint formatId, string name, byte[]? bytes)
    {
        TabItem.IsEnabled = _imagePreviews.IsImageCompatible(formatId, name);

        if (!TabItem.IsEnabled)
        {
            SetUnavailable("Image preview unavailable for this format.");
            return;
        }

        if (bytes == null)
        {
            SetUnavailable("Image preview requires byte-addressable clipboard data.");
            return;
        }

        if (_imagePreviews.TryCreatePreview(formatId, name, bytes,
                out var imagePreview, out var failureMessage) && imagePreview != null)
            SetImagePreview(imagePreview);
        else
            SetUnavailable(failureMessage);
    }

    public void Reset()
    {
        TabItem.IsEnabled = true;
        SetUnavailable("Select a clipboard format to preview.");
    }

    // ── Loaded hook ──────────────────────────────────────────────────────────

    private void ImagePreviewControl_Loaded(object sender, RoutedEventArgs e)
    {
        var window = Window.GetWindow(this);
        if (window == null) return;
        window.PreviewKeyDown += (_, _) => UpdateImageScrollViewerCursor();
        window.PreviewKeyUp   += (_, _) => UpdateImageScrollViewerCursor();
    }

    // ── Private helpers ──────────────────────────────────────────────────────

    private void SetImagePreview(BitmapSource image)
    {
        ImagePreview.Source = image;
        ImageStatusTextBlock.Text = string.Empty;
        ZoomSlider.IsEnabled = true;
        ImageDimensionsWidthTextBlock.Text   = $"{image.PixelWidth}px";
        ImageDimensionsHeightTextBlock.Text  = $"{image.PixelHeight}px";
        ImageDimensionsStackPanel.Visibility = Visibility.Visible;
        SetZoomValue(1);
        UpdateFitScale();
    }

    private void SetUnavailable(string message)
    {
        ImagePreview.Source = null;
        ImageStatusTextBlock.Text = message;
        ZoomSlider.IsEnabled = false;
        ImageDimensionsStackPanel.Visibility = Visibility.Collapsed;
        SetZoomValue(1);
        ApplyImageScale(1);
    }

    // ── Zoom ─────────────────────────────────────────────────────────────────

    private void ZoomSlider_OnValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (_ignoreZoomChanges || !IsLoaded)
            return;

        ApplyImageScale();
    }

    private void ResetZoomButton_OnClick(object sender, RoutedEventArgs e)
    {
        SetZoomValue(1);
        ApplyImageScale();
    }

    private void ZoomUpButton_Click(object sender, RoutedEventArgs e)   => StepZoom(+1);
    private void ZoomDownButton_Click(object sender, RoutedEventArgs e) => StepZoom(-1);

    private void StepZoom(int direction)
    {
        var currentPct = Math.Round(_fitScale * ZoomSlider.Value * 100);
        ApplyZoomFromPercentage(currentPct + direction);
    }

    private void ZoomTextBox_LostFocus(object sender, RoutedEventArgs e) => CommitZoomTextBox();

    private void ZoomTextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            CommitZoomTextBox();
            e.Handled = true;
        }
        else if (e.Key == Key.Escape)
        {
            ZoomTextBox.Text = FormatZoomPercentage(_fitScale * ZoomSlider.Value);
            ZoomTextBox.SelectAll();
            e.Handled = true;
        }
    }

    private void CommitZoomTextBox()
    {
        var text = ZoomTextBox.Text.Trim().TrimEnd('%').TrimEnd();
        if (double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out var pct) && pct > 0)
            ApplyZoomFromPercentage(pct);
        else
            ZoomTextBox.Text = FormatZoomPercentage(_fitScale * ZoomSlider.Value);
    }

    private void ApplyZoomFromPercentage(double pct)
    {
        var newSliderValue = _fitScale > 0 ? pct / 100.0 / _fitScale : pct / 100.0;
        newSliderValue   = Math.Clamp(newSliderValue, ZoomSlider.Minimum, ZoomSlider.Maximum);
        ZoomSlider.Value = newSliderValue;
    }

    private void ImageScrollViewer_OnSizeChanged(object sender, SizeChangedEventArgs e)
    {
        UpdateFitScale();
    }

    private void UpdateFitScale()
    {
        if (ImagePreview.Source is not BitmapSource source || source.PixelWidth <= 0 || source.PixelHeight <= 0)
        {
            _fitScale = 1;
            ApplyImageScale();
            return;
        }

        var viewportWidth  = ImageScrollViewer.ViewportWidth;
        var viewportHeight = ImageScrollViewer.ViewportHeight;
        if (viewportWidth <= 0 || viewportHeight <= 0)
        {
            _fitScale = 1;
            ApplyImageScale();
            return;
        }

        var fitX  = viewportWidth  / source.PixelWidth;
        var fitY  = viewportHeight / source.PixelHeight;
        _fitScale = Math.Min(1.0, Math.Min(fitX, fitY));
        ApplyImageScale();
    }

    private void ApplyImageScale(double? explicitScale = null)
    {
        if (ImageScaleTransform == null || ZoomTextBox == null || ZoomSlider == null)
            return;

        var scale = explicitScale ?? (_fitScale * ZoomSlider.Value);
        if (scale <= 0)
            return;

        ImageScaleTransform.ScaleX = scale;
        ImageScaleTransform.ScaleY = scale;
        ZoomTextBox.Text = FormatZoomPercentage(scale);
    }

    private static string FormatZoomPercentage(double scale) => $"{scale * 100:0}%";

    private void SetZoomValue(double value)
    {
        _ignoreZoomChanges = true;
        ZoomSlider.Value   = value;
        _ignoreZoomChanges = false;
    }

    // ── Pan & scroll ─────────────────────────────────────────────────────────

    private void ImageScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if ((Keyboard.Modifiers & ModifierKeys.Control) == 0)
            return;

        if (ImagePreview.Source is not BitmapSource source)
            return;

        e.Handled = true;

        var mousePos     = e.GetPosition(ImageScrollViewer);
        var currentScale = _fitScale * ZoomSlider.Value;
        var contentX     = ImageScrollViewer.HorizontalOffset + mousePos.X;
        var contentY     = ImageScrollViewer.VerticalOffset   + mousePos.Y;
        var scaledW      = source.PixelWidth  * currentScale;
        var scaledH      = source.PixelHeight * currentScale;
        var imageLeft    = Math.Max(0.0, (ImageScrollViewer.ViewportWidth  - scaledW) / 2.0);
        var imageTop     = Math.Max(0.0, (ImageScrollViewer.ViewportHeight - scaledH) / 2.0);
        var imgLocalX    = (contentX - imageLeft) / currentScale;
        var imgLocalY    = (contentY - imageTop)  / currentScale;

        StepZoom(Math.Sign(e.Delta));

        Dispatcher.BeginInvoke(DispatcherPriority.Background, () =>
        {
            var newScale   = _fitScale * ZoomSlider.Value;
            var newScaledW = source.PixelWidth  * newScale;
            var newScaledH = source.PixelHeight * newScale;
            var newImgLeft = Math.Max(0.0, (ImageScrollViewer.ViewportWidth  - newScaledW) / 2.0);
            var newImgTop  = Math.Max(0.0, (ImageScrollViewer.ViewportHeight - newScaledH) / 2.0);
            ImageScrollViewer.ScrollToHorizontalOffset(imgLocalX * newScale + newImgLeft - mousePos.X);
            ImageScrollViewer.ScrollToVerticalOffset  (imgLocalY * newScale + newImgTop  - mousePos.Y);
        });
    }

    private void ImageScrollViewer_MouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Middle)
            return;

        _isPanning           = true;
        _panStartMouse       = e.GetPosition(ImageScrollViewer);
        _panHorizontalOffset = ImageScrollViewer.HorizontalOffset;
        _panVerticalOffset   = ImageScrollViewer.VerticalOffset;
        ImageScrollViewer.Cursor = Cursors.ScrollAll;
        Mouse.Capture(ImageScrollViewer);
        e.Handled = true;
    }

    private void ImageScrollViewer_MouseUp(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Middle || !_isPanning)
            return;

        _isPanning = false;
        Mouse.Capture(null);
        UpdateImageScrollViewerCursor();
        e.Handled = true;
    }

    private void ImageScrollViewer_MouseMove(object sender, MouseEventArgs e)
    {
        if (_isPanning)
        {
            var pos = e.GetPosition(ImageScrollViewer);
            ImageScrollViewer.ScrollToHorizontalOffset(_panHorizontalOffset + (_panStartMouse.X - pos.X));
            ImageScrollViewer.ScrollToVerticalOffset  (_panVerticalOffset   + (_panStartMouse.Y - pos.Y));
            return;
        }

        UpdateImageScrollViewerCursor();
    }

    private void ImageScrollViewer_MouseLeave(object sender, MouseEventArgs e)
    {
        if (_isPanning)
            return;

        ImageScrollViewer.Cursor = null;
    }

    private void UpdateImageScrollViewerCursor()
    {
        if (_isPanning)
            return;

        if (!ImageScrollViewer.IsMouseOver)
            return;

        ImageScrollViewer.Cursor = (Keyboard.Modifiers & ModifierKeys.Control) != 0
            ? Cursors.Hand
            : null;
    }
}
