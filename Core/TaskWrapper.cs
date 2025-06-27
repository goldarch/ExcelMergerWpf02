using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelMergerWpf02.Core
{
    public enum TaskExecutionState
    {
        Idle, Starting, Running, Cancelling, Completed, Faulted, Cancelled
    }

    public class TaskStateChangedEventArgs : EventArgs
    {
        public TaskExecutionState PreviousState { get; }
        public TaskExecutionState NewState { get; }
        public Exception Exception { get; }

        public TaskStateChangedEventArgs(TaskExecutionState previousState, TaskExecutionState newState, Exception exception = null)
        {
            PreviousState = previousState;
            NewState = newState;
            Exception = exception;
        }
    }

    public class TaskWrapper : INotifyPropertyChanged
    {
        private TaskExecutionState _currentState = TaskExecutionState.Idle;
        private readonly object _stateLock = new object();

        public Func<CancellationToken, IProgress<TaskProgressInfo>, Task<string>> DoWorkFuncAsync { get; set; }
        public CancellationTokenSource CancellationTokenSource { get; private set; }
        public Progress<TaskProgressInfo> ProgressReporter { get; } = new Progress<TaskProgressInfo>();

        public TaskExecutionState CurrentState
        {
            get { lock (_stateLock) return _currentState; }
            private set
            {
                if (_currentState != value)
                {
                    _currentState = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(IsBusy));
                }
            }
        }

        public bool IsBusy => CurrentState == TaskExecutionState.Starting || CurrentState == TaskExecutionState.Running;

        public event EventHandler<TaskStateChangedEventArgs> StateChanged;
        public event PropertyChangedEventHandler PropertyChanged;

        public TaskWrapper()
        {
            CancellationTokenSource = new CancellationTokenSource();
        }

        public void RequestCancel()
        {
            if (IsBusy)
            {
                CancellationTokenSource?.Cancel();
            }
        }

        public async Task StartTaskAsync()
        {
            if (IsBusy)
            {
                ((IProgress<TaskProgressInfo>)ProgressReporter).Report(new TaskProgressInfo("任务已在运行中。", ReportLevel.Warning));
                return;
            }

            CancellationTokenSource?.Dispose();
            CancellationTokenSource = new CancellationTokenSource();
            var token = CancellationTokenSource.Token;
            var stopwatch = Stopwatch.StartNew();

            try
            {
                SetState(TaskExecutionState.Starting);
                if (DoWorkFuncAsync == null)
                {
                    throw new InvalidOperationException($"{nameof(DoWorkFuncAsync)} 委托未设置。");
                }

                SetState(TaskExecutionState.Running);
                string taskResultMessage = await DoWorkFuncAsync(token, ProgressReporter).ConfigureAwait(false);
                token.ThrowIfCancellationRequested();

                if (!string.IsNullOrWhiteSpace(taskResultMessage))
                {
                    throw new Exception(taskResultMessage);
                }

                SetState(TaskExecutionState.Completed);
            }
            catch (OperationCanceledException)
            {
                SetState(TaskExecutionState.Cancelled);
            }
            catch (Exception ex)
            {
                SetState(TaskExecutionState.Faulted, ex);
            }
        }

        private void SetState(TaskExecutionState newState, Exception ex = null)
        {
            var previousState = CurrentState;
            CurrentState = newState;
            StateChanged?.Invoke(this, new TaskStateChangedEventArgs(previousState, newState, ex));
        }

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}