using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA.Classes
{
  public  class BuilderFileName
    {
        private readonly IFileSystem _fileSystem;
        private string _inputFileName;
        private string _extension;
        public string OutputFileName { get; set; };

        public BuilderFileName() : this(new FileSystem()) { }

        public BuilderFileName(IFileSystem fileSystem)
        {
            _fileSystem = fileSystem;
        }

        public BuilderFileName(string inputFileName, string extension)
        {
            _inputFileName = inputFileName;
            _extension = extension;
        }
        public void BuildPath()
        {
            OutputFileName= MakeNewFileNameIfItHereExists(_inputFileName, _extension);
        }

        private string MakeNewFileNameIfItHereExists(string inputFileName, string extension)
        {
            string newNameOfFile = inputFileName;
            if (_fileSystem.File.Exists($"{inputFileName}.{extension}"))
            {
                newNameOfFile = $"{inputFileName}_1";
                MakeNewFileNameIfItHereExists(newNameOfFile, extension);
            }

            return newNameOfFile;
        }
    }
}
