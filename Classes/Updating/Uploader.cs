using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Diagnostics.Contracts;
using System.IO.Abstractions;

namespace ASTA.Classes.Updating
{
    public class Uploader : IDisposable
    {
        readonly IFileSystem fileSystem;
        public delegate void Info<TextEventArgs>(object sender, TextEventArgs e);
        public event Info<TextEventArgs> StatusText;

        public delegate void MarkerOfUploading(object sender, ColorEventArgs e);
        public event MarkerOfUploading StatusColor;
        string messageOfErrorUploading = null;

        public delegate void Uploaded(object sender, BoolEventArgs e);
        public event Uploaded StatusFinishedUploading;

       // private string[] _source;
        System.IO.Abstractions.IFileInfo[] _source;
        //private string[] _target;
        System.IO.Abstractions.IFileInfo[] _target;

        private bool uploadingError = false;
        UpdatingParameters _parameters { get; set; }



        public Uploader(IFileSystem fileSystem) { this.fileSystem = fileSystem; }

        //public Uploader(UpdatingParameters parameters, string[] source, string[] target) : this(fileSystem: new FileSystem())     //use default implementation which calls System.IO
       public Uploader(UpdatingParameters parameters, List<System.IO.Abstractions.IFileInfo> source, List<System.IO.Abstractions.IFileInfo> target): this(fileSystem: new FileSystem())     //use default implementation which calls System.IO
        {
            StatusText?.Invoke(this, new TextEventArgs(""));
            _parameters = parameters;
            _source = source.ToArray();
            _target = target.ToArray();
        }

        public async void Upload()
        {
            StatusText?.Invoke(this, new TextEventArgs("Начало отправки обновлений..."));
            StatusColor?.Invoke(this, new ColorEventArgs( System.Drawing.SystemColors.Control));
            uploadingError = false;

            Contract.Requires(!string.IsNullOrWhiteSpace(_parameters?.localFolderUpdatingURL));
            Contract.Requires(!string.IsNullOrWhiteSpace(_parameters?.appUpdateFolderURI));
            Contract.Requires(!string.IsNullOrWhiteSpace(_parameters?.appFileZip));
            Contract.Requires(!string.IsNullOrWhiteSpace(_parameters?.appFileXml));

            Contract.Requires(_source?.Length > 0);
            Contract.Requires(_target.Length == _source.Length);

            Func<Task>[] tasks = MakeFuncTask();
            
            await InvokeAsync(tasks, maxDegreeOfParallelism: 2);
            
            if (!uploadingError)
            {
                StatusText?.Invoke(this, new TextEventArgs($"Обновление отправлено -> {_parameters.remoteFolderUpdatingURL}"));
                StatusColor?.Invoke(this, new ColorEventArgs( System.Drawing.Color.PaleGreen));
            }
            else
            {
                StatusText?.Invoke(this, new TextEventArgs($"Ошибки отправки обновления -> {_parameters.remoteFolderUpdatingURL}\r\n{messageOfErrorUploading}"));
                StatusColor?.Invoke(this, new ColorEventArgs( System.Drawing.Color.DarkOrange));
            }
        }

        private Func<Task>[] MakeFuncTask()
        {
            int len = _source.Length;
            Func<Task>[] tasks = new Func<Task>[len];

            for (int index = 0; index < len; index++)
            {
                int i = index;
                tasks[i] = () => UploadApplicationToShare(_source[i].FullName, _target[i].FullName);
            }

            return tasks;
        }


        public async Task UploadApplicationToShare(string source, string target)
        {
            Contract.Requires(!source.Equals(target));
            StatusText?.Invoke(this, new TextEventArgs($"Идет отправка файла {source} -> {target}"));
            StatusFinishedUploading?.Invoke(this, new BoolEventArgs(false));
            StatusColor?.Invoke(this, new ColorEventArgs(System.Drawing.SystemColors.Control));

            try
            {
                // var fileByte = System.IO.File.ReadAllBytes(source);
                // System.IO.File.WriteAllBytes(target, fileByte);
                try { System.IO.File.Delete(target); }
                catch { StatusText?.Invoke(this, new TextEventArgs($"Файл на сервере: {target} удалить не удалось")); } //@"\\server\folder\Myfile.txt"

                System.IO.File.Copy(source, target, true); //@"\\server\folder\Myfile.txt"
                
                StatusText?.Invoke(this, new TextEventArgs($"Отправка файла на сервер выполнена " + target));
                StatusColor?.Invoke(this, new ColorEventArgs(System.Drawing.Color.LightGreen));
                StatusFinishedUploading?.Invoke(this, new BoolEventArgs(true));
            }
            catch (Exception err)
            {
                StatusText?.Invoke(this, new TextEventArgs($"Отправка файла на сервер {target} завершена с ошибкой: {err.ToString()}"));
                uploadingError = true;

                if (string.IsNullOrEmpty(messageOfErrorUploading))
                    messageOfErrorUploading = err.Message;
                else
                    messageOfErrorUploading += "|" + err.Message;
                StatusColor?.Invoke(this, new ColorEventArgs(System.Drawing.Color.LightYellow));
            }
        }
        private static async Task InvokeAsync(IEnumerable<Func<Task>> taskFactories, int maxDegreeOfParallelism)
        {
            var queue = new Queue<Func<Task>>(taskFactories);

            if (queue.Count == 0)
            {
                return;
            }

            var tasksInFlight = new List<Task>(maxDegreeOfParallelism);

            do
            {
                while (tasksInFlight.Count < maxDegreeOfParallelism && queue.Count != 0)
                {
                    var taskFactory = queue.Dequeue();

                    tasksInFlight.Add(taskFactory());
                }

                var completedTask = await Task.WhenAny(tasksInFlight).ConfigureAwait(false);

                // Propagate exceptions. In-flight tasks will be abandoned if this throws.
                await completedTask.ConfigureAwait(false);

                tasksInFlight.Remove(completedTask);
            }
            while (queue.Count != 0 || tasksInFlight.Count != 0);
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        private void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    _parameters = null;
                    _source = _target = null;
                    messageOfErrorUploading = null;
                    StatusColor = null;
                    StatusText = null;
                    StatusFinishedUploading = null;
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

