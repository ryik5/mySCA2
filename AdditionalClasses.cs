﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Cryptography;
using System.Windows.Forms;

namespace ASTA
{

    public class Person
    {
        public int id;
        public int idCard = 0;
        public string FIO = "";
        public string NAV = "";
        public string Department = "";
        public string PositionInDepartment = "";

        public string GroupPerson = "Office";
        public decimal ControlInDecimal = 9;
        public decimal ControlOutDecimal = 18;
        public string ControlInHHMM = "09:00";
        public string ControlOutHHMM = "18:00";

        public string RealInHHMM = "09:00";
        public string RealOutHHMM = "18:00";

        public string RealWorkedDayHoursHHMM = "09:00";
        public string RealDate = "";
        public string RealDayOfWeek = "";

        public bool Late = false;
        public bool Early = false;

        public string serverSKD = "";
        public string namePassPoint = "";
        public string directionPass = "";
        public string Shift = "";
        public string Comment = "";
    }

    public class DataGridViewSeekValuesInSelectedRow
    {
        public string[] values = new string[10];
        public bool correctData { get; set; }

        public void FindValueInCells(DataGridView DGV, params string[] columnsName)
        {
            int IndexCurrentRow = DGV.CurrentRow.Index;
            correctData = IndexCurrentRow > -1 ? true : false;

            if (correctData)
            {
                for (int i = 0; i < DGV.ColumnCount; i++)
                {
                    if (DGV.Columns[i].HeaderText == columnsName[0])
                    {
                        values[0] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 1 && DGV.Columns[i].HeaderText == columnsName[1])
                    {
                        values[1] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 2 && DGV.Columns[i].HeaderText == columnsName[2])
                    {
                        values[2] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 3 && DGV.Columns[i].HeaderText == columnsName[3])
                    {
                        values[3] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 4 && DGV.Columns[i].HeaderText == columnsName[4])
                    {
                        values[4] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 5 && DGV.Columns[i].HeaderText == columnsName[5])
                    {
                        values[5] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 6 && DGV.Columns[i].HeaderText == columnsName[6])
                    {
                        values[6] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 7 && DGV.Columns[i].HeaderText == columnsName[7])
                    {
                        values[7] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 8 && DGV.Columns[i].HeaderText == columnsName[8])
                    {
                        values[8] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 9 && DGV.Columns[i].HeaderText == columnsName[9])
                    {
                        values[9] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                }
            }

        }
    }

    public static class DataTableExtensions
    {
        public static void SetColumnsOrder(this DataTable table, params string[] columnNames)
        {
            List<string> listColNames = columnNames.ToList();

            //Remove invalid column names.
            foreach (string colName in columnNames)
            {
                if (!table.Columns.Contains(colName))
                {
                    listColNames.Remove(colName);
                }
            }

            int columnIndex = 0;
            foreach (var columnName in listColNames)
            {
                table.Columns[columnName].SetOrdinal(columnIndex);
                columnIndex++;
            }
        }
    }

    public class MailingStructure
    {
        public string _sender = "";
        public string _recipient = "";
        public string _groupsReport = "";
        public string _nameReport = "";
        public string _descriptionReport = "";
        public string _period = "";
        public string _status = "";
    }

    public class OutReasons
    {
        public string _id = "";
        public string _name = "";
        public string _visibleName = "";
        public int _hourly = 0;
    }

    public class OutPerson
    {
        public string _reason_id = "0";
        public string _nav = "";
        public string _date = "";
        public int _from = 0;
        public int _to = 0;
        public int _hourly = 0;
    }

    public struct PeopleShift
    {
        public string _fio;
        public string _nav;
        public string _dayStartShift;
        public int _MoStart;
        public int _MoEnd;
        public int _TuStart;
        public int _TuEnd;
        public int _WeStart;
        public int _WeEnd;
        public int _ThStart;
        public int _ThEnd;
        public int _FrStart;
        public int _FrEnd;
        public int _SaStart;
        public int _SaEnd;
        public int _SuStart;
        public int _SuEnd;
        public string _Status;
        public string _Comment;
        public string _DayInputed;
        public string _Reserv1;
        public string _Reserv2;
    }

    public class PeopleDepartment
    {
        public string _departmentName;
        public string _departmentDescription;
        public string _parent_id;
        public string _id;

    }

    public class EncryptDecrypt
    {
        /*
            string plainText = "Hello, World!";            
            string TestKey = "gFlMfLZu4unBGfBh9weIuLwlcXHSa59vxyUQWM3yh1M=";
            string TestIV = "rRNgGypocle9VG9bth6kxg==";

            var ed = new EncDec(TestKey, TestIV);
            var cypherText = ed.Encrypt(plainText);  //4U8SPwju5aH9pHwmiYd1jQ==
            var plainText2 = ed.Decrypt(cypherText); //4U8SPwju5aH9pHwmiYd1jQ==
*/
        private byte[] key;
        private byte[] IV;

        public void EncDec(string keyText, string ivText)
        {
            key = Convert.FromBase64String(keyText);
            IV = Convert.FromBase64String(ivText);
        }

        public void EncDecKeys(byte[] key, byte[] IV)
        {
            this.key = key;
            this.IV = IV;
        }

        public string Encrypt(string plainText)
        {
            var bytes = EncryptStringToBytes(plainText, key, IV);
            return Convert.ToBase64String(bytes);
        }

        private static byte[] EncryptStringToBytes(string plainText, byte[] Key, byte[] IV)
        {
            // Check arguments. 
            if (plainText == null || plainText.Length <= 0)
                throw new ArgumentNullException("plainText");
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("IV");
            byte[] encrypted;
            // Create an RijndaelManaged object 
            // with the specified key and IV. 
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
                            swEncrypt.Write(plainText);  //Write all data to the stream.
                        }
                        encrypted = msEncrypt.ToArray();
                    }
                }
            }

            // Return the encrypted bytes from the memory stream. 
            return encrypted;
        }

        public string Decrypt(string cipherText)
        {
            var bytes = Convert.FromBase64String(cipherText);
            return DecryptStringFromBytes(bytes, key, IV);
        }

        private static string DecryptStringFromBytes(byte[] cipherText, byte[] Key, byte[] IV)
        {
            // Check arguments. 
            if (cipherText == null || cipherText.Length <= 0)
                throw new ArgumentNullException("cipherText");
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("IV");

            // Declare the string used to hold  the decrypted text. 
            string plaintext = null;

            // Create an RijndaelManaged object with the specified key and IV. 
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

                            // Read the decrypted bytes from the decrypting stream 
                            // and place them in a string.
                            plaintext = srDecrypt.ReadToEnd();
                        }
                    }
                }
            }

            return plaintext;
        }
    }
}