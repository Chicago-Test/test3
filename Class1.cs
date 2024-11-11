# save as UTF-8 with BOM
# Need to restart powershell session when changing c#.
# powershell.exe -noprofile -executionpolicy bypass -file GenerateDBP.ps1

$passwd = '!@#$%^&*()+=[]{};:<>?'

$excelFilePath = 'C:\Users\Administrator\source\repos\GenerateDBP_powershell\MVG-change-password_v1.4(64bit).xlsm'
$excelFilePath = 'C:\Users\Administrator\source\repos\GenerateDBP_powershell\a.xls'
$excelFilePath = 'C:\Users\Administrator\source\repos\GenerateDBP_powershell\no_VBA.xls'
$excelFilePath = 'C:\Users\Administrator\source\repos\GenerateDBP_powershell\no_VBA.xlsm'
$excelFilePath = 'C:\Users\Administrator\source\repos\GenerateDBP_powershell\a_pass$$.xls'

##############################################################
Remove-Variable * -ErrorAction SilentlyContinue
cd $PSScriptRoot

Remove-Item "vbaProject.bin" -Force -ErrorAction SilentlyContinue
Remove-Item "PROJECT" -Force -ErrorAction SilentlyContinue
$7zipPath = "C:\Program Files\7-Zip\7z.exe"


#if ($Args.Count -gt 1) { $excelFilePath = $Args[0];$passwd=$Args[1] }else {}

if ([System.IO.File]::Exists($excelFilePath) -eq $false) {
  Write-Host "File does not exist";
  exit;
}

$x = $excelFilePath.Split(".")[-1]
if ($x -eq "xlsm" -or $x -eq "xlsb") {
  $cmdArgs1 = " e " + """" + $excelFilePath + """ " + """" + "xl\vbaProject.bin" + """"
  $cmdArgs2 = " e " + "vbaProject.bin PROJECT"
  Start-Process -FilePath $7zipPath -ArgumentList $cmdArgs1 -Wait
  Start-Process -FilePath $7zipPath -ArgumentList $cmdArgs2 -Wait
}
elseif ($x -eq "xls") {
  $cmdArgs1 = " e " + """" + $excelFilePath + """ " + "_VBA_PROJECT_CUR\PROJECT" + """"
  Start-Process -FilePath $7zipPath -ArgumentList $cmdArgs1 -Wait
}
else {
  Write-Host "Input file is not xlsm/xlsb/xls"
  exit;
}

if ([System.IO.File]::Exists(".\PROJECT") -eq $false) {
  Write-Host "no VBA (no PROJECT file)";
  exit;
}

$DPB = ""
# Extract "DPB=" line
foreach ($line in Get-Content .\PROJECT) {
  #There are differences between foreach and foreach-object
  if ($line -match "^(DPB=)") {
    #Write-Host $line.ToString()
    $DPB = $line.Substring(5, $line.Length - 6)
    break;
  }
}

if ($DPB.Length -lt 50) { Write-Host "VBA not locked"; exit; }

$source = @"
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace GenerateDBP
{
    public static class Program
    {
        public static string Main1(string strDPB,string passwd)
        {
            string str;
            byte seed;

            //strDPB = "94963888C84FE54FE5B01B50E59251526FE67A1CC76C84ED0DAD653FD058F324BFD9D38DED37";

            seed = Convert.ToByte(strDPB.Substring(0, 2), 16);

            // Check decrypt->encrypt will produce the original string
            byte[] y = decryptDPB(strDPB); // Returned first byte is a seed. Last byte is 0x00
            byte[] z1 = encryptDPB(BitConverter.ToString(y.ToArray()).Replace("-", "").Substring(2), seed); // Specify seed
            str = BitConverter.ToString(z1.ToArray()).Replace("-", "");
            //if (str == strDPB) { Console.WriteLine("Two strings are equal"); }

            // Extract Key for the Password Hash Algorithm

            byte[] passKey;
            try
            {
                passKey = GetPassKeyWithNull(strDPB);
            }
            catch (Exception)
            {
                //Console.WriteLine("no password");
                return "no password";
            }

            bool blRet;
            blRet = IsPasswordCorrect(strDPB, passwd);
            //if (blRet == true) { Console.WriteLine("passwd correct"); } else { Console.WriteLine("passwd wrong"); }
            if (blRet == true) { str="passwd correct"; } else { str="passwd wrong"; }

            return str;
        }

        private static byte[] GetPassKeyWithNull(string strDPB)
        {
            byte seed = Convert.ToByte(strDPB.Substring(0, 2), 16);
            byte[] y = decryptDPB(strDPB); // Returned first byte is a seed. Last byte is 0x00

            int IgnoredLength = (seed & 6) / 2;

            byte[] datalength = new byte[4];
            Buffer.BlockCopy(y, 1 + 2 + IgnoredLength, datalength, 0, 4); // Always {0x1d, 0x00, 0x00, 0x00 }? If no password {0x01,0x00,0x00,0x00}?

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
    }
}
"@

Add-Type -TypeDefinition $source
$ret = [GenerateDBP.Program]::Main1($DPB, $passwd)
Write-Host $ret

exit
