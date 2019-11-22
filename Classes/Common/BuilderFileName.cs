using System.IO.Abstractions;

namespace ASTA.Classes
{
    public class BuilderFileName
    {
        private readonly IFileSystem _fileSystem;
        private string _inputFileName;
        private string _extension;
        private string _outputFileNameWithExtention;

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

        public string BuildPath()
        {
            _outputFileNameWithExtention = ReturnFileNameWithExtention(_inputFileName, _extension);
            return _outputFileNameWithExtention;
        }

        /// <summary>
        /// If File 'inputFileName.extension' on the that place exists it returns 'inputFileName_1.extension' else return 'inputFileName.extension'
        /// </summary>
        /// <param name="inputFileName"></param>
        /// <param name="extension"></param>
        /// <returns></returns>
        private string ReturnFileNameWithExtention(string inputFileName, string extension)
        {
            string newNameOfFile = inputFileName;
            using (CheckerFileOnDisk fileExists = new CheckerFileOnDisk(inputFileName, extension))
            {
                if (fileExists.Exists())
                {
                    newNameOfFile = $"{inputFileName}_1";
                    ReturnFileNameWithExtention(newNameOfFile, extension);
                }
            }
            return $"{newNameOfFile}.{extension}";
        }
    }
}
    

