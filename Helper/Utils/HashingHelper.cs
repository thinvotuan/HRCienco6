using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web;

namespace BatDongSan.Utils
{
    public static class HashingHelper
    {

        //private string EnCrypt(string strEnCrypt, string key)
        //{
        //    try
        //    {
        //        byte[] keyArr;
        //        byte[] EnCryptArr = UTF8Encoding.UTF8.GetBytes(strEnCrypt);
        //        MD5CryptoServiceProvider MD5Hash = new MD5CryptoServiceProvider();
        //        keyArr = MD5Hash.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
        //        TripleDESCryptoServiceProvider tripDes = new TripleDESCryptoServiceProvider();
        //        tripDes.Key = keyArr;
        //        tripDes.Mode = CipherMode.ECB;
        //        tripDes.Padding = PaddingMode.PKCS7;
        //        ICryptoTransform transform = tripDes.CreateEncryptor();
        //        byte[] arrResult = transform.TransformFinalBlock(EnCryptArr, 0, EnCryptArr.Length);
        //        return Convert.ToBase64String(arrResult, 0, arrResult.Length);
        //    }
        //    catch (Exception ex) { }
        //    return "";
        //}

        //private string DeCrypt(string strDecypt, string key)
        //{
        //    try
        //    {
        //        byte[] keyArr;
        //        byte[] DeCryptArr = Convert.FromBase64String(strDecypt);
        //        MD5CryptoServiceProvider MD5Hash = new MD5CryptoServiceProvider();
        //        keyArr = MD5Hash.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
        //        TripleDESCryptoServiceProvider tripDes = new TripleDESCryptoServiceProvider();
        //        tripDes.Key = keyArr;
        //        tripDes.Mode = CipherMode.ECB;
        //        tripDes.Padding = PaddingMode.PKCS7;
        //        ICryptoTransform transform = tripDes.CreateDecryptor();
        //        byte[] arrResult = transform.TransformFinalBlock(DeCryptArr, 0, DeCryptArr.Length);
        //        return UTF8Encoding.UTF8.GetString(arrResult);
        //    }
        //    catch (Exception ex) { }
        //    return "";
        //}

        public static string Encrypt(string toEncrypt, bool useHashing)
        {
            byte[] keyArray;
            byte[] toEncryptArray = UTF8Encoding.UTF8.GetBytes(toEncrypt);

            // Get the key from config file
            string key = "thuanvietmvc4";
            //System.Windows.Forms.MessageBox.Show(key);
            if (useHashing)
            {
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
                hashmd5.Clear();
            }
            else
                keyArray = UTF8Encoding.UTF8.GetBytes(key);

            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            tdes.Key = keyArray;
            tdes.Mode = CipherMode.ECB;
            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateEncryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(toEncryptArray, 0, toEncryptArray.Length);
            tdes.Clear();
            return Convert.ToBase64String(resultArray, 0, resultArray.Length);
        }
        /// <summary>
        /// DeCrypt a string using dual encryption method. Return a DeCrypted clear string
        /// </summary>
        /// <param name="cipherString">encrypted string</param>
        /// <param name="useHashing">Did you use hashing to encrypt this data? pass true is yes</param>
        /// <returns></returns>
        public static string Decrypt(string cipherString, bool useHashing)
        {
            byte[] keyArray;
            byte[] toEncryptArray = Convert.FromBase64String(cipherString);

            //Get your key from config file to open the lock!
            string key = "thuanvietmvc4";

            if (useHashing)
            {
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
                hashmd5.Clear();
            }
            else
                keyArray = UTF8Encoding.UTF8.GetBytes(key);

            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            tdes.Key = keyArray;
            tdes.Mode = CipherMode.ECB;
            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateDecryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(toEncryptArray, 0, toEncryptArray.Length);

            tdes.Clear();
            return UTF8Encoding.UTF8.GetString(resultArray);
        }
    }
}
