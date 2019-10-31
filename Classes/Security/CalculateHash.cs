using System;
using System.Security.Cryptography;

namespace ASTA.Classes.Security
{
    internal class CalculatingHash
    {
        private  string FilName{ get; set; }
        public CalculatingHash(string filename)
        {
            FilName = filename;
        }

        public  string Calculate(string algorithm = "MD5") //MD5, SHA1, SHA256, SHA384, SHA512
        {
            string fileChecksum = null;
            using (var hashAlgorithm = HashAlgorithm.Create(algorithm))
            {
                using (var stream = System.IO.File.OpenRead(FilName))
                {
                    if (hashAlgorithm != null)
                    {
                        var hash = hashAlgorithm.ComputeHash(stream);
                        fileChecksum = BitConverter.ToString(hash).Replace("-", string.Empty).ToLowerInvariant();
                    }

                    return fileChecksum;
                }
            }
        }

        //todo
        //check it
        /*
                static string CalculateHash(string filename)
                {
                    using (var fileStream = new System.IO.FileStream(filename, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                    {
                        SHA1CryptoServiceProvider unmanaged = new SHA1CryptoServiceProvider();
                        byte[] retVal = unmanaged.ComputeHash(fileStream);
                        return string.Join("", retVal.Select(x => x.ToString("x2")));
                    }
                }*/

        /*
                //Test algorithm to evaluate with the autoupdater's algorithm
                private static bool CompareChecksum(string fileName, string checksum, string algorythm = "MD5") //MD5, SHA1, SHA256, SHA384, SHA512
                {
                    using (var hashAlgorithm = HashAlgorithm.Create(algorythm))
                    {
                        using (var stream =System.IO.File.OpenRead(fileName))
                        {
                            if (hashAlgorithm != null)
                            {
                                var hash = hashAlgorithm.ComputeHash(stream);
                                var fileChecksum = BitConverter.ToString(hash).Replace("-", string.Empty).ToLowerInvariant();

                                if (fileChecksum == checksum.ToLower()) return true;
                            }

                            return false;
                        }
                    }
                }

                //Test algorithm to evaluate with the autoupdater's algorithm
                private void TestHash()
                {
                    string filePath = null;
                    filePath = SelectFileOpenFileDialog("Выберите первый файл для вычисления хэша");
                    string myFileHash = CalculateFileHash(filePath);

                    filePath = SelectFileOpenFileDialog("Выберите второй файл для вычисления хэша");
                    bool result = CompareChecksum(filePath, myFileHash);
                    MessageBox.Show("Result of evaluation of checking\n"+ result);
                }*/

    }
}
