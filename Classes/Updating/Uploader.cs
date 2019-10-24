using ASTA.Classes.AutoUpdating;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ASTA.Classes.Updating
{
   public class Uploader
    {
        public delegate void Status(object sender, AccountEventArgs e);
        public event Status status;

        UpdatingParameters _parameters { get; set; }

        public Uploader(UpdatingParameters parameters)
        {
            _parameters = parameters;
        }

        public async void Upload()
        {
            if (_parameters == null)
            {
                throw new ArgumentNullException("UpdatingParameters", "UpdatingParameters cannot be empty");
            }

            if (string.IsNullOrWhiteSpace(_parameters.appUpdateFolderURI ))
            {
                throw new ArgumentNullException("appUpdateFolderURI", "appUpdateFolderURI cannot be empty");
            }
            if (string.IsNullOrWhiteSpace( _parameters.appFileZip))
            {
                throw new ArgumentNullException("appFileZip", "appFileZip cannot be empty");
            }

            Func<Task>[] tasks =
            {
                () => UploadApplicationToShare(
                    _parameters.localFolderUpdatingURL + @"\" + _parameters.appFileXml,
                    _parameters.appUpdateFolderURI +_parameters.appFileXml),  //Send app.xml file to server
                () => UploadApplicationToShare(
                    _parameters.localFolderUpdatingURL + @"\" + _parameters.appFileZip,
                    _parameters.appUpdateFolderURI + _parameters.appFileZip)                        //Send app.zip file to server
            };

            await InvokeAsync(tasks, maxDegreeOfParallelism: 2);
            
            status?.Invoke(this, new AccountEventArgs("Обновление на сервер загружено -> " + _parameters.remoteFolderUpdatingURL));
        }

        private async Task UploadApplicationToShare(string source, string target)
        {
                status?.Invoke(this, new AccountEventArgs("Идет отправка файла " + source + " -> " + target));

                try
                {
                    // var fileByte = System.IO.File.ReadAllBytes(source);
                    // System.IO.File.WriteAllBytes(target, fileByte);
                    try { System.IO.File.Delete(target); }
                    catch { status?.Invoke(this, new AccountEventArgs("Файл на сервере: " + target + " удалить не удалось")); } //@"\\server\folder\Myfile.txt"

                    System.IO.File.Copy(source, target, true); //@"\\server\folder\Myfile.txt"
                    status?.Invoke(this, new AccountEventArgs("Отправка файла на сервер выполнена " + target));

                    try { System.IO.File.Delete(source); } catch { }
                }
                catch (Exception err)
                {
                    status?.Invoke(this, new AccountEventArgs("Отправка файла на сервер " + target + " завершена с ошибкой: " + err.ToString()));
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

    }
}

