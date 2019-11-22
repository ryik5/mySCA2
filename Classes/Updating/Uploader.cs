using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Threading.Tasks;
using System.Diagnostics.Contracts;
using System.IO.Abstractions;

namespace ASTA.Classes.Updating
{
    public class Uploader : IDisposable
    {
        public delegate void Info<TextEventArgs>(object sender, TextEventArgs e);
        public event Info<TextEventArgs> StatusText;

        public delegate void MarkerOfUploading(object sender, ColorEventArgs e);
        public event MarkerOfUploading StatusColor;
        string messageOfErrorUploading = null;

        public delegate void Uploaded(object sender, BoolEventArgs e);
        public event Uploaded StatusFinishedUploading;

        readonly IFileSystem fileSystem;
        List<System.IO.Abstractions.IFileInfo> _sourceList;
        List<System.IO.Abstractions.IFileInfo> _targetList;

        FilePathSourceAndTarget[] _couples;

        private bool uploadingError = false;
        UpdatingParameters _parameters { get; set; }



        public Uploader(IFileSystem fileSystem) { this.fileSystem = fileSystem; }

        //public Uploader(UpdatingParameters parameters, string[] source, string[] target) : this(fileSystem: new FileSystem())     //use default implementation which calls System.IO
        public Uploader(UpdatingParameters parameters, List<System.IO.Abstractions.IFileInfo> source, List<System.IO.Abstractions.IFileInfo> target) : this(fileSystem: new FileSystem())     //use default implementation which calls System.IO
        {
            StatusText?.Invoke(this, new TextEventArgs(""));
            _parameters = parameters;

            _sourceList = source;
            _targetList = target;
        }

        public async void Upload()
        {
            StatusText?.Invoke(this, new TextEventArgs("Начало отправки обновлений..."));
            StatusColor?.Invoke(this, new ColorEventArgs(System.Drawing.SystemColors.Control));
            uploadingError = false;

            Contract.Requires(!string.IsNullOrWhiteSpace(_parameters?.localFolderUpdatingURL));
            Contract.Requires(!string.IsNullOrWhiteSpace(_parameters?.appUpdateFolderURI));
            Contract.Requires(!string.IsNullOrWhiteSpace(_parameters?.appFileZip));
            Contract.Requires(!string.IsNullOrWhiteSpace(_parameters?.appFileXml));

            Contract.Requires(_sourceList?.Count > 0);
            Contract.Requires(_targetList.Count == _sourceList.Count);

            _couples = MakeArrayFilePathesFromTwoListsOfFilePathes(_sourceList, _targetList);

            Func<Task>[] tasks = MakeFuncTask(_couples);

            await InvokeAsync(tasks, maxDegreeOfParallelism: 2);

            if (!uploadingError)
            {
                StatusText?.Invoke(this, new TextEventArgs($"Обновление отправлено -> {_parameters.remoteFolderUpdatingURL}"));
                StatusColor?.Invoke(this, new ColorEventArgs(System.Drawing.Color.PaleGreen));
            }
            else
            {
                StatusText?.Invoke(this, new TextEventArgs($"Ошибки отправки обновления -> {_parameters.remoteFolderUpdatingURL}\r\n{messageOfErrorUploading}"));
                StatusColor?.Invoke(this, new ColorEventArgs(System.Drawing.Color.DarkOrange));
            }
        }

        private Func<Task>[] MakeFuncTask(FilePathSourceAndTarget[] couples)
        {
            int len = couples.Length;
            Func<Task>[] tasks = new Func<Task>[len];

            for (int i = 0; i < len; i++)
            {
                int index = i;
                tasks[index] = () => UploadFileToShare(couples[index]);
            }

            return tasks;
        }

        private FilePathSourceAndTarget[] MakeArrayFilePathesFromTwoListsOfFilePathes(List<System.IO.Abstractions.IFileInfo> source, List<System.IO.Abstractions.IFileInfo> target)
        {
            FilePathSourceAndTarget[] couples = new FilePathSourceAndTarget[source.Count];
            for (int i = 0; i < source.Count; i++)
            {
                couples[i] = new FilePathSourceAndTarget(source[i], target[i]);
            }
            return couples;
        }

        public async Task UploadFileToShare(FilePathSourceAndTarget pathes)
        {
            var source = pathes.Get()._sourcePath;
            var target = pathes.Get()._targetPath;
            Contract.Requires(!source.Equals(target));
            StatusText?.Invoke(this, new TextEventArgs($"Идет отправка файла {source.FullName} -> {target.FullName}"));
            StatusFinishedUploading?.Invoke(this, new BoolEventArgs(false));
            StatusColor?.Invoke(this, new ColorEventArgs(System.Drawing.SystemColors.Control));

            try
            {
                // var fileByte = System.IO.File.ReadAllBytes(source);
                // System.IO.File.WriteAllBytes(target, fileByte);
                try { target.Delete(); }
                catch { StatusText?.Invoke(this, new TextEventArgs($"Файл на сервере: {target.FullName} удалить не удалось")); } //@"\\server\folder\Myfile.txt"

                fileSystem.File.Copy(source.FullName, target.FullName, true); //@"\\server\folder\Myfile.txt"

                StatusText?.Invoke(this, new TextEventArgs($"Отправка файла на сервер выполнена " + target));
                StatusColor?.Invoke(this, new ColorEventArgs(System.Drawing.Color.LightGreen));
                StatusFinishedUploading?.Invoke(this, new BoolEventArgs(true));
            }
            catch (Exception err)
            {
                StatusText?.Invoke(this, new TextEventArgs($"Отправка файла на сервер {target.FullName} завершена с ошибкой: {err.ToString()}"));
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
                    _sourceList = null;
                    _targetList = null;
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

    public class FilePathSourceAndTarget
    {
        public System.IO.Abstractions.IFileInfo _sourcePath { get; set; }
        public System.IO.Abstractions.IFileInfo _targetPath { get; set; }
        public FilePathSourceAndTarget Get()
        {
            return new FilePathSourceAndTarget(_sourcePath, _targetPath);
        }
        public FilePathSourceAndTarget(System.IO.Abstractions.IFileInfo source, System.IO.Abstractions.IFileInfo target)
        {
            _sourcePath = source;
            _targetPath = target;
        }
    }

}

