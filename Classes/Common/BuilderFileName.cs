using System.IO.Abstractions;

namespace ASTA.Classes
{
    public static class BuilderFileName
    {

        public static string BuildPath(string inputFileName, string extension)
        {
            return ReturnFileNameWithExtention(inputFileName, extension); 
        }

        /// <summary>
        /// If File 'inputFileName.extension' on the that place exists it returns 'inputFileName_1.extension' else return 'inputFileName.extension'
        /// </summary>
        /// <param name="inputFileName"></param>
        /// <param name="extension"></param>
        /// <returns></returns>
        static string ReturnFileNameWithExtention(string inputFileName, string extension)
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
    

