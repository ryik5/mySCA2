using System;
using System.Security.Cryptography;

namespace ASTA.Classes.Security
{
//    internal class EncryptDecrypt
//    {
//        /*
//            string plainText = "Hello, World!";            
//            string TestKey = "gFlMfLZu4unBGfBh9weIuLwlcXHSa59vxyUQWM3yh1M=";
//            string TestIV = "rRNgGypocle9VG9bth6kxg==";

//            var ed = new EncDec(TestKey, TestIV);
//            var cypherText = ed.Encrypt(plainText);  //4U8SPwju5aH9pHwmiYd1jQ==
//            var plainText2 = ed.Decrypt(cypherText); //4U8SPwju5aH9pHwmiYd1jQ==
//*/
//        private byte[] key;
//        private byte[] IV;

//        public void EncDec(string keyText, string ivText)
//        {
//            key = Convert.FromBase64String(keyText);
//            IV = Convert.FromBase64String(ivText);
//        }

//        public void EncDecKeys(byte[] key, byte[] IV)
//        {
//            this.key = key;
//            this.IV = IV;
//        }

//        public string Encrypt(string plainText)
//        {
//            var bytes = EncryptStringToBytes(plainText, key, IV);
//            return Convert.ToBase64String(bytes);
//        }

//        private static byte[] EncryptStringToBytes(string plainText, byte[] Key, byte[] IV)
//        {
//            // Check arguments. 
//            if (plainText == null || plainText.Length <= 0)
//                throw new ArgumentNullException("plainText");
//            if (Key == null || Key.Length <= 0)
//                throw new ArgumentNullException("Key");
//            if (IV == null || IV.Length <= 0)
//                throw new ArgumentNullException("IV");
//            byte[] encrypted;
//            // Create an RijndaelManaged object 
//            // with the specified key and IV. 
//            using (RijndaelManaged rijAlg = new RijndaelManaged())
//            {
//                rijAlg.Key = Key;
//                rijAlg.IV = IV;

//                // Create a decrytor to perform the stream transform.
//                ICryptoTransform encryptor = rijAlg.CreateEncryptor(rijAlg.Key, rijAlg.IV);

//                // Create the streams used for encryption. 
//                using (System.IO.MemoryStream msEncrypt = new System.IO.MemoryStream())
//                {
//                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
//                    {
//                        using (System.IO.StreamWriter swEncrypt = new System.IO.StreamWriter(csEncrypt))
//                        {
//                            swEncrypt.Write(plainText);  //Write all data to the stream.
//                        }
//                        encrypted = msEncrypt.ToArray();
//                    }
//                }
//            }

//            // Return the encrypted bytes from the memory stream. 
//            return encrypted;
//        }

//        public string Decrypt(string cipherText)
//        {
//            var bytes = Convert.FromBase64String(cipherText);
//            return DecryptStringFromBytes(bytes, key, IV);
//        }

//        private static string DecryptStringFromBytes(byte[] cipherText, byte[] Key, byte[] IV)
//        {
//            // Check arguments. 
//            if (cipherText == null || cipherText.Length <= 0)
//                throw new ArgumentNullException("cipherText");
//            if (Key == null || Key.Length <= 0)
//                throw new ArgumentNullException("Key");
//            if (IV == null || IV.Length <= 0)
//                throw new ArgumentNullException("IV");

//            // Declare the string used to hold  the decrypted text. 
//            string plaintext = null;

//            // Create an RijndaelManaged object with the specified key and IV. 
//            using (RijndaelManaged rijAlg = new RijndaelManaged())
//            {
//                rijAlg.Key = Key;
//                rijAlg.IV = IV;

//                // Create a decrytor to perform the stream transform.
//                ICryptoTransform decryptor = rijAlg.CreateDecryptor(rijAlg.Key, rijAlg.IV);

//                // Create the streams used for decryption. 
//                using (System.IO.MemoryStream msDecrypt = new System.IO.MemoryStream(cipherText))
//                {
//                    using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
//                    {
//                        using (System.IO.StreamReader srDecrypt = new System.IO.StreamReader(csDecrypt))
//                        {

//                            // Read the decrypted bytes from the decrypting stream 
//                            // and place them in a string.
//                            plaintext = srDecrypt.ReadToEnd();
//                        }
//                    }
//                }
//            }

//            return plaintext;
//        }
//    }

    internal static class EncryptionDecryptionCriticalData
    {
        public static string EncryptStringToBase64Text(string plainText, byte[] Key, byte[] IV) //Encrypt variables PlainText Data
        {
            string sBase64Test;
            sBase64Test = Convert.ToBase64String(EncryptStringToBytes(plainText, Key, IV));
            return sBase64Test;
        }

        public static byte[] EncryptStringToBytes(string plainText, byte[] Key, byte[] IV) //Encrypt variables PlainText Data
        {
            // Check arguments. 
            if (plainText == null || plainText.Length <= 0)
                throw new ArgumentNullException("plainText");
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("IV");
            byte[] encrypted;
            // Create an RijndaelManaged object with the specified key and IV. 
            using (RijndaelManaged rijAlg = new RijndaelManaged())
            {
                rijAlg.Key = Key;
                rijAlg.IV = IV;

                // Create a decrytor to perform the stream transform.
                ICryptoTransform encryptor = rijAlg.CreateEncryptor(rijAlg.Key, rijAlg.IV);

                // Create the streams used for encryption. 
                using (System.IO.MemoryStream msEncrypt = new System.IO.MemoryStream())
                {
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        using (System.IO.StreamWriter swEncrypt = new System.IO.StreamWriter(csEncrypt))
                        {

                            //Write all data to the stream.
                            swEncrypt.Write(plainText);
                        }
                        encrypted = msEncrypt.ToArray();
                    }
                }
            }
            // Return the encrypted bytes from the memory stream. 
            return encrypted;
        }

        public static string DecryptBase64ToString(string sBase64Text, byte[] Key, byte[] IV) //Encrypt variables PlainText Data
        {
            byte[] bBase64Test;
            bBase64Test = Convert.FromBase64String(sBase64Text);
            return DecryptStringFromBytes(bBase64Test, Key, IV);
        }

        public static string DecryptStringFromBytes(byte[] cipherText, byte[] Key, byte[] IV) //Decrypt PlainText Data to variables
        {
            // Check arguments. 
            if (cipherText == null || cipherText.Length <= 0)
                throw new ArgumentNullException("cipherText");
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("IV");

            // Declare the string used to hold the decrypted text.
            string plaintext = null;

            // Create an RijndaelManaged object  with the specified key and IV.
            using (RijndaelManaged rijAlg = new RijndaelManaged())
            {
                rijAlg.Key = Key;
                rijAlg.IV = IV;

                // Create a decrytor to perform the stream transform.
                ICryptoTransform decryptor = rijAlg.CreateDecryptor(rijAlg.Key, rijAlg.IV);

                // Create the streams used for decryption. 
                using (System.IO.MemoryStream msDecrypt = new System.IO.MemoryStream(cipherText))
                {
                    using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                    {
                        using (System.IO.StreamReader srDecrypt = new System.IO.StreamReader(csDecrypt))
                        {
                            // Read the decrypted bytes from the decrypting stream and place them in a string. 
                            plaintext = srDecrypt.ReadToEnd();
                        }
                    }
                }
            }

            return plaintext;
        }
    }


    /*
           //TestCryptionItem_Click
           
            string original = "Here is some data to encrypt!";
            MessageBox.Show("Original:   " + original);

            using (RijndaelManaged myRijndael = new RijndaelManaged())
            {
                myRijndael.Key = btsMess1;
                myRijndael.IV = btsMess2;
                // Encrypt the string to an array of bytes.
                byte[] encrypted = EncryptionDecryptionCriticalData.EncryptStringToBytes(original, myRijndael.Key, myRijndael.IV);

                StringBuilder s = new StringBuilder();
                foreach (byte item in encrypted)
                {
                    s.Append(item.ToString("X2") + " ");
                }
                MessageBox.Show("Encrypted:   " + s);

                // Decrypt the bytes to a string.
                string decrypted = EncryptionDecryptionCriticalData.DecryptStringFromBytes(encrypted, btsMess1, btsMess2);

                //Display the original data and the decrypted data.
                MessageBox.Show("Decrypted:    " + decrypted);
            }
 
      */
}
