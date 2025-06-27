using System;
using System.Windows.Input;

namespace ExcelMergerWpf02.Commands
{
    public class RelayCommand : ICommand
    {
        private readonly Action<object> _execute;
        private readonly Predicate<object> _canExecute;

        // CanExecuteChanged 事件的声明方式保持不变
        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public RelayCommand(Action<object> execute, Predicate<object> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter) => _canExecute == null || _canExecute(parameter);

        public void Execute(object parameter) => _execute(parameter);

        // +++ 新增方法：允许从外部手动触发 CanExecuteChanged 事件 +++
        /// <summary>
        /// Manually raises the CanExecuteChanged event to force the command's state to be re-evaluated.
        /// </summary>
        public void RaiseCanExecuteChanged()
        {
            // CommandManager.InvalidateRequerySuggested() 是在UI线程上触发全局刷新的一种方式
            CommandManager.InvalidateRequerySuggested();
        }
    }
}