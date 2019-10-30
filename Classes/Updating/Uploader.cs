using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Diagnostics.Contracts;

namespace ASTA.Classes.Updating
{
    public class Uploader : IDisposable
    {
        public delegate void InfoStatus(object sender, EventTextArgs e);
        public event InfoStatus Info;

        public delegate void MarkerOfUploading(object sender, EventColorArgs e);
        public event MarkerOfUploading ColorOfStatus;
        string messageOfErrorUploading = null;

        public delegate void Uploaded(object sender, EventBoolArgs e);
        public event Uploaded Status;

        private string[] _source;
        private string[] _target;
        private bool uploadingError = false;
        UpdatingParameters _parameters { get; set; }

        public Uploader(UpdatingParameters parameters, string[] source, string[] target)
        {
            Info?.Invoke(this, new EventTextArgs(""));
            _parameters = parameters;
            _source = source;
            _target = target;
        }

        public async void Upload()
        {
            Info?.Invoke(this, new EventTextArgs("Начало отправки обновлений..."));

            ColorOfStatus?.Invoke(this, new EventColorArgs( System.Drawing.SystemColors.Control));
            uploadingError = false;

            Contract.Requires(_parameters != null);
            Contract.Requires(_parameters.localFolderUpdatingURL != null);
            Contract.Requires(_parameters.appUpdateFolderURI != null);
            Contract.Requires(_parameters.appFileZip != null);
            Contract.Requires(_parameters.appFileXml != null);

            Contract.Requires(_source != null);
            Contract.Requires(_target.Length == _source.Length);

            Func<Task>[] tasks = MakeFuncTask();
            await InvokeAsync(tasks, maxDegreeOfParallelism: 2);


            if (!uploadingError)
            {
                Info?.Invoke(this, new EventTextArgs("Обновление отправлено -> " + _parameters.remoteFolderUpdatingURL));
                ColorOfStatus?.Invoke(this, new EventColorArgs( System.Drawing.Color.PaleGreen));
            }
            else
            {
                Info?.Invoke(this, new EventTextArgs("Ошибки отправки обновления -> " + _parameters.remoteFolderUpdatingURL + "\r\n" + messageOfErrorUploading));
                ColorOfStatus?.Invoke(this, new EventColorArgs( System.Drawing.Color.DarkOrange));
            }
        }

        Func<Task>[] MakeFuncTask()
        {
            int len = _source.Length;
            Func<Task>[] tasks = new Func<Task>[len];

            for (int index = 0; index < len; index++)
            {
                int i = index;
                tasks[i] = (() => UploadApplicationToShare(_source[i], _target[i]));
            }

            return tasks;
        }

        private async Task UploadApplicationToShare(string source, string target)
        {
            Contract.Requires(!source.Equals(target));
            Info?.Invoke(this, new EventTextArgs("Идет отправка файла " + source + " -> " + target));
            Status?.Invoke(this, new EventBoolArgs(false));
            ColorOfStatus?.Invoke(this, new EventColorArgs(System.Drawing.SystemColors.Control));

            try
            {
                // var fileByte = System.IO.File.ReadAllBytes(source);
                // System.IO.File.WriteAllBytes(target, fileByte);
                try { System.IO.File.Delete(target); }
                catch { Info?.Invoke(this, new EventTextArgs("Файл на сервере: " + target + " удалить не удалось")); } //@"\\server\folder\Myfile.txt"

                System.IO.File.Copy(source, target, true); //@"\\server\folder\Myfile.txt"
                
                Info?.Invoke(this, new EventTextArgs("Отправка файла на сервер выполнена " + target));
                ColorOfStatus?.Invoke(this, new EventColorArgs(System.Drawing.Color.LightGreen));
                Status?.Invoke(this, new EventBoolArgs(true));
            }
            catch (Exception err)
            {
                Info?.Invoke(this, new EventTextArgs("Отправка файла на сервер " + target + " завершена с ошибкой: " + err.ToString()));
                uploadingError = true;

                if (string.IsNullOrEmpty(messageOfErrorUploading))
                    messageOfErrorUploading = err.Message;
                else
                    messageOfErrorUploading += "|" + err.Message;
                ColorOfStatus?.Invoke(this, new EventColorArgs(System.Drawing.Color.LightYellow));
            }
        }
        private static async Task InvokeAsync(IEnumerable<Func<Task>> taskFactories, int maxDegreeOfParallelism)
        {
            Queue<Func<Task>> queue = new Queue<Func<Task>>(taskFactories);

            if (queue.Count == 0)
            {
                return;
            }

            List<Task> tasksInFlight = new List<Task>(maxDegreeOfParallelism);

            do
            {
                while (tasksInFlight.Count < maxDegreeOfParallelism && queue.Count != 0)
                {
                    Func<Task> taskFactory = queue.Dequeue();

                    tasksInFlight.Add(taskFactory());
                }

                Task completedTask = await Task.WhenAny(tasksInFlight).ConfigureAwait(false);

                // Propagate exceptions. In-flight tasks will be abandoned if this throws.
                await completedTask.ConfigureAwait(false);

                tasksInFlight.Remove(completedTask);
            }
            while (queue.Count != 0 || tasksInFlight.Count != 0);
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    _parameters = null;
                    _source = _target = null;
                    messageOfErrorUploading = null;
                    ColorOfStatus = null;
                    Info = null;
                    Status = null;
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            GC.SuppressFinalize(this);
        }
        #endregion
    }
}

