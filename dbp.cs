using OpenMcdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace CompoundFormat1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //VBAのパスワードクラッキングで学ぶ共通鍵暗号とハッシュ関数
            //https://note.com/mizuki_arioka/n/n6b4c0aa5b7f8

            String filename = null;
            CFStream foundStream;
            byte[] byteBuf = null;
            CFStream myStream;
            string str;
            byte seed;

            string strDPB = "94963888C84FE54FE5B01B50E59251526FE67A1CC76C84ED0DAD653FD058F324BFD9D38DED37";
            strDPB = "D4D678CD8801A501A5FE5B02A509433EBAB0E6407491D89B60E50AEF0053158339B9ABBDCE7E"; // pass:MVG

            seed = Convert.ToByte(strDPB.Substring(0, 2), 16);

            // Check decrypt->encrypt will produce the original string
            byte[] y = decryptDPB(strDPB); // Returned first byte is a seed. Last byte is 0x00
            byte[] z1 = encryptDPB(BitConverter.ToString(y.ToArray()).Replace("-", "").Substring(2), seed); // Specify seed
            str = BitConverter.ToString(z1.ToArray()).Replace("-", "");
            if (str == strDPB) { Console.WriteLine("Two strings are equal"); }

            // Extract Key for the Password Hash Algorithm
            byte[] passKey = GetPassKeyWithNull(strDPB);

            bool blRet;
            blRet = IsPasswordCorrect(strDPB, "MVG");
            string passwd = "a%#&-)bc";
            str = GenerateDBP(passwd, seed, passKey); // To replace current DBP= in vbaProject.bin, use the same seed as current so that the length of DBP won't change.
            blRet = IsPasswordCorrect(str, passwd);

            return;

            // Replace one file in the archived file(overwrite)
            // Clear VBA Project password
            try
            {
                filename = @"..\..\vbaProject2.bin";
                CompoundFile cf4 = new CompoundFile(filename, CFSUpdateMode.Update, CFSConfiguration.Default);
                foundStream = cf4.RootStorage.GetStream("PROJECT");
                byteBuf = foundStream.GetData();
                var x = Encoding.UTF8.GetString(byteBuf.ToArray()).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                cf4.RootStorage.Delete("PROJECT");
                // modify byteBuf 
                using (var ms = new MemoryStream())
                {
                    for (int i = 0; i < x.Length; i++)
                    {
                        if (x[i].StartsWith("DPB=") == true) { x[i] = "DPB=\"0E0CD1ECDFF4E7F5E7F5E7\""; } // clear VBA project password
                        ms.Write(Encoding.GetEncoding("UTF-8").GetBytes(x[i]), 0, x[i].Length);
                        if (i != x.Length - 1) { ms.WriteByte(0x0D); ms.WriteByte(0x0A); }
                    }
                    myStream = cf4.RootStorage.AddStream("PROJECT");
                    myStream.SetData(ms.ToArray());
                }
                cf4.Commit(); cf4.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
            return;

            //create a new compound file
            byte[] b = { 0x60, 0x62 };
            CompoundFile cf = new CompoundFile();
            myStream = cf.RootStorage.AddStream("MyStream");
            myStream.SetData(b);
            cf.SaveAs("MyCompoundFile.cfs");
            cf.Close();

            //create a new compound file with subdirectory
            CompoundFile cf2 = new CompoundFile();
            CFStorage st = cf2.RootStorage.AddStorage("MyStorage").AddStorage("bin");
            st.AddStorage("bin2");
            CFStream sm = st.AddStream("MyStream");
            sm.SetData(b);
            cf2.SaveAs("MyCompoundFile2.cfs");
            cf2.Close();

            //open an existing one
            filename = @"..\..\vbaProject.bin";
            CompoundFile cf3 = new CompoundFile(filename);
            foundStream = cf3.RootStorage.GetStorage("VBA").GetStream("modchecksumif"); //case insensitive
            byteBuf = foundStream.GetData();
            File.WriteAllBytes(@".\outfile", byteBuf);
            cf3.Close();



            //If you need to compress a compound file, you can purge its unused space
            //CompoundFile.ShrinkCompoundFile("MultipleStorage_Deleted_Compress.cfs");

        }

        private static byte[] GetPassKeyWithNull(string strDPB)
        {
            byte seed = Convert.ToByte(strDPB.Substring(0, 2), 16);
            byte[] y = decryptDPB(strDPB); // Returned first byte is a seed. Last byte is 0x00

            int IgnoredLength = (seed & 6) / 2;

            byte[] datalength = new byte[4]; // Always {0x1d, 0x00, 0x00, 0x00 }?
            Buffer.BlockCopy(y, 1 + 2 + IgnoredLength, datalength, 0, 4);

            int headerLength = 1 + 2 + IgnoredLength + 4; // Header: Seed,ver,projkey, ignored bytes
            string dat1 = BitConverter.ToString(y.ToArray()).Replace("-", "").Substring(2 * headerLength); // DATA
            //Console.WriteLine(dat1);

            //var byteDat1 = Enumerable.Range(0, s1.Length / 2).Select(x => Convert.ToByte(s1.Substring(x * 2, 2), 16)).ToArray(); // Hex string to byte[]
            var byteDat1 = y.Skip(headerLength).ToArray();
            //byteDat1[0] must be 0xFF.  xx[28] must be Terminator 0x00
            if (byteDat1[0] != 0xFF) { Console.WriteLine("corrupt data"); } //MUST be 0xFF. MUST be ignored
            byte[] Grbits = new byte[3]; //Each bit specifies a corresponding null byte of Key or Password Hash
            Buffer.BlockCopy(byteDat1, 1, Grbits, 0, 3);
            //Replace byte[] data with 0x00 based on Grbits

            string bitString = Convert.ToString(Grbits[0] * 256 * 256 + Grbits[1] * 256 + Grbits[2], 2);
            for (int i = 0; i < bitString.Length; i++)
            {
                if (bitString[i] == '0') { byteDat1[i + 4] = 0x00; }  // Decode NULL
            }

            byte[] passKeyWithNulls = new byte[4];
            Buffer.BlockCopy(byteDat1, 4, passKeyWithNulls, 0, 4);
            byte[] hashedPassWithNulls = new byte[20];
            Buffer.BlockCopy(byteDat1, 8, hashedPassWithNulls, 0, 20);
            return passKeyWithNulls;
        }

        private static bool IsPasswordCorrect(string strDPB, string password)
        {
            //2.4.4.1 Password Hash Data Structure
            byte seed = Convert.ToByte(strDPB.Substring(0, 2), 16); // First two characters is HEX seed
            byte[] decrypted_DPB = decryptDPB(strDPB);
            int IgnoredLength = (seed & 6) / 2;

            //Grbits: Each bit specifies a corresponding null byte of Key or Password Hash
            byte[] Grbits = decrypted_DPB.Skip(1 + 2 + IgnoredLength + 4 + 1).Take(3).ToArray();
            //Convert.ToString(Convert.ToInt64("0xffffff", 16), 2);
            string bitString = Convert.ToString(Grbits[0] * 256 * 256 + Grbits[1] * 256 + Grbits[2], 2);
            var KeyNoNulls_PasswordHashNoNulls = decrypted_DPB.Skip(1 + 2 + IgnoredLength + 4 + 4).Take(24).ToArray();
            for (int i = 0; i < bitString.Length; i++)
            {
                if (bitString[i] == '0') { KeyNoNulls_PasswordHashNoNulls[i] = 0x00; }  // Decode NULL
            }

            var KeyWithNulls = KeyNoNulls_PasswordHashNoNulls.Take(4).ToArray();
            var PasswordHashWithNulls = KeyNoNulls_PasswordHashNoNulls.Skip(4).Take(20).ToArray();

            //PasswordHash is the 160-bit cryptographic digest of a password combined with Key as specified by Password Hash Algorithm
            var passBytes = Encoding.GetEncoding("UTF-8").GetBytes(password).ToArray();
            var passBytesPlusKey = passBytes.Concat(KeyWithNulls).ToArray();  // append key after password
            byte[] passHash = SHA1CryptoServiceProvider.Create().ComputeHash(passBytesPlusKey);
            bool ret = PasswordHashWithNulls.SequenceEqual(passHash); // Compare byte arrays
            return ret;
        }

        private static string GenerateDBP(string password, byte seed, byte[] passKey)
        {
            //string password = "MVG";
            byte[] passByte = new byte[password.Length + 4];
            //byte[] tmp = Encoding.GetEncoding("UTF-8").GetBytes(password);
            Buffer.BlockCopy(Encoding.GetEncoding("UTF-8").GetBytes(password), 0, passByte, 0, password.Length);
            Buffer.BlockCopy(passKey, 0, passByte, password.Length, 4); // Add 4 bytes key after password
            byte[] passHash = SHA1CryptoServiceProvider.Create().ComputeHash(passByte);
            string s2 = BitConverter.ToString(passHash).Replace("-", "");
            //string s4=(new string('1',24)).ToCharArray();
            char[] cc = (new string('1', 24)).ToCharArray();
            byte[] keyAndpassByte = new byte[24];
            var passkey_plus_passHash = passKey.Concat(passHash).ToArray();
            for (int i = 0; i < cc.Length; i++)
            {
                if (passkey_plus_passHash[i] == 0) { passkey_plus_passHash[i] = 0xFF; cc[i] = '0'; }
            }
            string s4 = new string(cc);
            byte[] x0 = new byte[1] { 0xff };
            byte[] xx = BitConverter.GetBytes(Convert.ToInt32(s4, 2));
            keyAndpassByte = x0.Concat(xx.Take(3).Reverse().ToArray()).Concat(passkey_plus_passHash).Concat(new byte[1]).ToArray();  // need to reverse Grbitkey

            //2.4.3.1 Encrypted Data Structure
            // seed,

            int IgnoredLength = (seed & 6) / 2;
            byte[] header = Enumerable.Repeat((byte)0xBB, 6 + IgnoredLength).ToArray(); // 0xBB is arbitary. Ignored bytes
            header[0] = 0x02; header[1] = 0xAC; header[IgnoredLength + 2] = 0x1d; header[IgnoredLength + 3] = 0x00; header[IgnoredLength + 4] = 0x00; header[IgnoredLength + 5] = 0x00;
            //byte[] header = { 0x02, 0xAC, 0xFF, 0xFF, 0x1d, 0x00, 0x00, 0x00 }; // version,projkey,ignored,DataLength  [IgnoredLength = (seed & 6) / 2]
            //byte[] DBPbeforeXOR = (new byte[] { seed }).Concat(header).Concat(keyAndpassByte).ToArray(); // first byte is arbitary seed
            byte[] DBPbeforeXOR_without_seed = header.Concat(keyAndpassByte).ToArray();
            string s5 = BitConverter.ToString(DBPbeforeXOR_without_seed).Replace("-", "");
            byte[] y1 = encryptDPB(s5, seed); // y1[0]=seed
            string str = BitConverter.ToString(y1.ToArray()).Replace("-", "");
            return str;
        }

        private static byte[] decryptDPB(string strDPB)
        {
            return DecryptEncryptDPB(false, strDPB, 0x00);
        }
        private static byte[] encryptDPB(string strDPB, byte seed)
        {
            return DecryptEncryptDPB(true, strDPB, seed);
        }
        private static byte[] DecryptEncryptDPB(bool Enc_or_Decrypt_Flag, string strDPB, byte EncSeed)
        {
            // Enc_or_Decrypt_Flag==True then Encryption
            //2.4.3.1 Encrypted Data Structure
            //string strDPB = "94963888C84FE54FE5B01B50E59251526FE67A1CC76C84ED0DAD653FD058F324BFD9D38DED37";

            // Return: [0]=Encryption seed [37]=0x00

            var xx = Enumerable.Range(0, strDPB.Length / 2).Select(x => Convert.ToByte(strDPB.Substring(x * 2, 2), 16)).ToArray(); // Hex string to byte[]
            if (Enc_or_Decrypt_Flag == true) { xx = (new byte[1] { EncSeed }).Concat(xx).ToArray(); }

            List<byte> y = new List<byte>();

            // xx[0]: Encryption seed
            y.Add(xx[0]);
            y.Add(Convert.ToByte(xx[1] ^ xx[0])); //version Must be 2
            y.Add(Convert.ToByte(xx[2] ^ xx[0])); // project key

            for (int i = 3; i < xx.Length; i++)
            {
                // 2.4.3 Data Encryption
                // All operations resulting in integer overflow MUST only store low-order bits, resulting inhigh-order bit truncation
                if (Enc_or_Decrypt_Flag == true)
                {
                    y.Add(Convert.ToByte(xx[i] ^ ((y[i - 2] + xx[i - 1]) & 0xFF))); // 0xFF is not to overflow
                }
                else
                {
                    y.Add(Convert.ToByte(xx[i] ^ ((xx[i - 2] + y[i - 1]) & 0xFF)));
                }
            }
            return y.ToArray();
        }
        //private static byte[] encryptDPB(string strDPB)
        //{
        //    var xx = Enumerable.Range(0, strDPB.Length / 2).Select(x => Convert.ToByte(strDPB.Substring(x * 2, 2), 16)).ToArray(); // Hex string to byte[]
        //    List<byte> y = new List<byte>();

        //    // xx[0]: Encryption seed
        //    y.Add(xx[0]);
        //    y.Add(Convert.ToByte(xx[1] ^ xx[0]));
        //    y.Add(Convert.ToByte(xx[2] ^ xx[0]));

        //    for (int i = 3; i < xx.Length; i++)
        //    {
        //        y.Add(Convert.ToByte(xx[i] ^ ((y[i - 2] + xx[i - 1]) & 0xFF))); // 0xFF is not to overflow
        //    }
        //    return y.ToArray();
        //}

        //private static void IgnoredEnc(byte seed, byte versionenc, byte projkey, byte projkeyenc)
        //{
        //    byte UnencryptedByte1 = projkey;
        //    byte EncryptedByte1 = projkeyenc;
        //    byte EncryptedByte2 = versionenc;

        //    int IgnoredLength = (seed & 6) / 2;
        //    for (int i = 1; i <= IgnoredLength; i++)
        //    {
        //        byte TempValue = 9; // any value.

        //        byte ByteEnc = (byte)(TempValue ^ ((EncryptedByte2 + UnencryptedByte1) & 0xFF)); // is this correct???
        //        //APPEND IgnoredEnc WITH ByteEnc.
        //        EncryptedByte2 = EncryptedByte1;
        //        EncryptedByte1 = ByteEnc;
        //        UnencryptedByte1 = TempValue;
        //    }

        //}

    }
}
