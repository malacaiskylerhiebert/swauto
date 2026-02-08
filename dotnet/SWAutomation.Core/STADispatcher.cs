using System;
using System.Collections.Concurrent;
using System.Threading;

namespace SWAutomation.Core
{
    internal sealed class STADispatcher : IDisposable
    {
        private readonly BlockingCollection<Action> _queue;
        private readonly Thread _thread;
        private int _staThreadID;
        private bool _disposed;

        public STADispatcher()
        {
            _queue = new BlockingCollection<Action>();

            _thread = new Thread(Run)
            {
                IsBackground = true
            };

            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }

        private void Run()
        {
            _staThreadID = Thread.CurrentThread.ManagedThreadId;

            foreach (var action in _queue.GetConsumingEnumerable())
            {
                action();
            }
        }
        public T Invoke<T>(Func<T> func)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(STADispatcher));

            if (Thread.CurrentThread.ManagedThreadId == _staThreadID)
                return func();

            var done = new ManualResetEventSlim(false);
            T result = default;
            Exception error = null;

            _queue.Add(() =>
            {
                try
                {
                    result = func();
                }
                catch (Exception ex)
                {
                    error = ex;
                }
                finally
                {
                    done.Set();
                }
            });

            done.Wait();

            if (error != null)
                throw error;

            return result;
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            _disposed = true;
            _queue.CompleteAdding();
            _thread.Join();
            _queue.Dispose();
        }
    }
}