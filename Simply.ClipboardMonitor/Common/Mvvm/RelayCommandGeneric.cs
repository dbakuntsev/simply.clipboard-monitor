using System.Windows.Input;

namespace Simply.ClipboardMonitor.Common.Mvvm;

public sealed class RelayCommand<T> : ICommand
{
    private readonly Action<T?>       _execute;
    private readonly Func<T?, bool>?  _canExecute;

    public RelayCommand(Action<T?> execute, Func<T?, bool>? canExecute = null)
    {
        _execute    = execute;
        _canExecute = canExecute;
    }

    public event EventHandler? CanExecuteChanged
    {
        add    => CommandManager.RequerySuggested += value;
        remove => CommandManager.RequerySuggested -= value;
    }

    public bool CanExecute(object? parameter) =>
        _canExecute?.Invoke(parameter is T t ? t : default) ?? true;

    public void Execute(object? parameter) =>
        _execute(parameter is T t ? t : default);

    public void RaiseCanExecuteChanged() => CommandManager.InvalidateRequerySuggested();
}
