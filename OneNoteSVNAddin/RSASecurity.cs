using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography;

namespace OneNoteSVNAddin
{
    /// <summary>
    /// RSA加解密法的静态操作类
    /// </summary>
    public static class RSASecurity
    {
        static RSASecurity()
        {
            param = new CspParameters();
            param.KeyContainerName = KeyContainerName;
        }

        private static CspParameters param;
        private static string KeyContainerName = "SVNAddin";//RSA加解密的密匙容器名称

        /// <summary>
        /// 加密操作
        /// </summary>
        /// <param name="encryptString">要加密的明文字符串</param>
        /// <returns></returns>
        public static string EncryptRSA(string encryptString)
        {
            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider(param))
            {
                byte[] plaindata = Encoding.Default.GetBytes(encryptString);//将要加密的字符串转换为字节数组
                byte[] encryptdata = rsa.Encrypt(plaindata, false);//将加密后的字节数据转换为新的加密字节数组
                return Convert.ToBase64String(encryptdata);//将加密后的字节数组转换为字符串
            }
        }

        /// <summary>
        /// 解密操作
        /// </summary>
        /// <param name="decryptString">要解密的密文字符串</param>
        /// <returns></returns>
        public static string DecryptRSA(string decryptString)
        {
            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider(param))
            {
                byte[] encryptdata = Convert.FromBase64String(decryptString);//字符串转字节数组
                byte[] decryptdata = rsa.Decrypt(encryptdata, false);//解密
                return Encoding.Default.GetString(decryptdata);//转字符串
            }
        }
    }
}