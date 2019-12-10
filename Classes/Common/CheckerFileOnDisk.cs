using System;

namespace ASTA.Classes
{
    /// <summary>
    /// Used testablbe 'System.IO.Abstractions'
    /// </summary>
    public class CheckerFileOnDisk : IDisposable
    {
        public delegate void Status(object sender, TextEventArgs e);

        public event Status status;

        //public CheckerFileOnDisk() : this(new FileSystem()) { }
        //public CheckerFileOnDisk(IFileSystem fileSystem)
        //{
        //    _fileSystem = fileSystem;
        //}

        private string _inputFileName, _extension, _inputFileNameWithExtention;
        //        private  IFileSystem _fileSystem;

        public CheckerFileOnDisk(string inputFileName, string extension)
        {
            _inputFileName = inputFileName;
            _extension = extension;
        }

        public void Set(string inputFileNameWithExtention)
        {
            _inputFileNameWithExtention = inputFileNameWithExtention;
        }

        public void Set(string inputFileName, string extension)
        {
            _inputFileName = inputFileName;
            _extension = extension;
        }

        public CheckerFileOnDisk(string inputFileNameWithExtention)
        {
            _inputFileNameWithExtention = inputFileNameWithExtention;
        }

        public bool Exists()
        {
            System.Diagnostics.Contracts.Contract.Requires(_inputFileName != null || _inputFileNameWithExtention != null);
            status?.Invoke(this, new TextEventArgs(
                $"{nameof(_inputFileName)}: {_inputFileName} |{nameof(_extension)}: {_extension} |{nameof(_inputFileNameWithExtention)}: {_inputFileNameWithExtention}"
                ));

            if (_inputFileName != null && _extension != null)
            {
                return System.IO.File.Exists($"{_inputFileName}.{_extension}");
            }
            else if (_inputFileName != null)
            {
                return System.IO.File.Exists($"{_inputFileName}");
                //return _fileSystem.File.Exists($"{_inputFileName}");
            }
            else
            {
                return System.IO.File.Exists($"{_inputFileNameWithExtention}");
                //return _fileSystem.File.Exists($"{_inputFileNameWithExtention}");
            }
        }

        #region IDisposable Support

        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).
                    if (_inputFileName != null) _inputFileName = null;
                    if (_extension != null) _extension = null;
                    if (_inputFileNameWithExtention != null) _inputFileNameWithExtention = null;
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~CheckerFileOnDisk()
        // {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }

        #endregion IDisposable Support
    }
}