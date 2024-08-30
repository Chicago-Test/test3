//4644EA4B1EBDC4DAC4DA3B26C5DAE779DE1F3BFC000294541C86D549B3AF2491BF629398A3018D
//0604AA0BDE7D849A849A7B66859A8C4CB3A3D4A7FF426093A566A0D1FA943B2965E1FAC477737F
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Security.Cryptography;

using SevenZipExtractor; //SevenZipExtractor.1.0.17 //Source code in this repo is licensed under The MIT License

//https://github.com/bontchev/pcodedmp/tree/master
//https://github.com/decalage2/oletools/wiki/olevba
//https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/575462ba-bf67-4190-9fac-c275523c75fc

namespace VBAStreamDecompress
{
    public class MVGvbaDecompress
    {
        static void Main(string[] args)
        {
            string filePath = @"..\..\check-formula-protection_v1.88 - testing.xlsm";
            if (args.Length > 0) { filePath = args[0]; Console.WriteLine(args[0]); }

            try
            {
                byte[] hash = archiveVBAcodes(filePath);
                Console.WriteLine("hash:" + Convert.ToString(hash[0], 16).ToUpper() + Convert.ToString(hash[1], 16).ToUpper() + Convert.ToString(hash[2], 16).ToUpper());
                Console.WriteLine("Done");

                //archiveVBAcodes2(new MemoryStream,""); // no VBA zip file output, just returns hash.

            }
            catch (Exception)
            {
                throw;
            }

            return;

            filePath = "../../modUnHideAllSheets";
            byte[] buf = File.ReadAllBytes(filePath);
            byte[] srcCode = MVG_decompress_stream(buf);
            File.WriteAllBytes("$$$out.txt", srcCode);
        }
        static private byte[] MVG_decompress_stream(byte[] buf)
        {
            // Search "00 00 FF FF FF FF 00 00 01" in Emeditor
            byte[] patternByte = { 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x01 };

            int indx = 0;
            int i = 0;

            // This will get first occurence of the pattern
            //indx = Boyer_Moore.indexOf(buf, patternByte);

            // Need to get the last occurence of the pattern
            i = Boyer_Moore.indexOf(buf.Reverse().ToArray(), patternByte.Reverse().ToArray());
            indx = buf.Length - (i+patternByte.Length);
            
            var buf2 = buf.Skip(indx + 8).ToArray();
            byte[] ret = decompress_stream(buf2);
            return ret;
        }
        // Based on oledump python
        static private (int, int, int, int) CopyTokenHelp(int decompressedCurrent, int decompressedChunkStart)
        //public void CopyTokenHelp(int decompressedCurrent, int decompressedChunkStart)
        {
            int difference = decompressedCurrent - decompressedChunkStart;
            int bitCount = (int)Math.Ceiling(Math.Log(difference, 2));
            bitCount = Math.Max(bitCount, 4);
            int lengthMask = 0xFFFF >> bitCount;
            int offsetMask = ~lengthMask;
            int maximumLength = (0xFFFF >> bitCount) + 3;
            return (lengthMask, offsetMask, bitCount, maximumLength);
            //return;
        }
        static private byte[] decompress_stream(byte[] compressedContainer)
        {
            //"""
            //Decompress a stream according to MS-OVBA section 2.4.1

            //:param compressed_container bytearray: bytearray or bytes compressed according to the MS-OVBA 2.4.1.3.6 Compression algorithm
            //:return: the decompressed container as a bytes string
            //:rtype: bytes
            //"""

            //# Check the input is a bytearray, otherwise convert it (assuming it's bytes):
            if (!(compressedContainer is byte[]))
            {
                compressedContainer = new byte[compressedContainer.Length];
                // throw new ArgumentException("decompress_stream requires a byte[] as input");
            }
            // log.Debug("decompress_stream: compressed size = {0} bytes", compressedContainer.Length);

            List<byte> decompressedContainer = new List<byte>(); //byte[] decompressedContainer = new byte[0]; // result

            int compressedCurrent = 0;

            // data starts from 0x01
            byte sigByte = compressedContainer[compressedCurrent];
            if (sigByte != 0x01)
            {
                throw new ArgumentException(string.Format("invalid signature byte {0:X2}", sigByte));
            }

            compressedCurrent++;

            while (compressedCurrent < compressedContainer.Length)
            {
                //# 2.4.1.1.5
                int compressed_chunk_start = compressedCurrent;
                // chunk header = first 16 bits
                int compressed_chunk_header = compressedContainer[compressed_chunk_start] + compressedContainer[compressed_chunk_start + 1] * 256;// struct.unpack("<H", compressed_container[compressed_chunk_start:compressed_chunk_start + 2])[0]  //little endian, unsighned short
                                                                                                                                                  // chunk size = 12 first bits of header + 3

                //YI chunk_header: 2 bytes. Little endian. In binary, it starts with 1011.... Remaining 12bit shows chunk size. Chunk size includes its own header.
                //YI 2^12=4096. This can express [0,4095]. Because there is no zero size, add one and [1,4096]

                int chunk_size = (compressed_chunk_header & 0x0FFF) + 3;
                // chunk signature = 3 next bits - should always be 0b011
                int chunk_signature = (compressed_chunk_header >> 12) & 0x07;
                if (chunk_signature != 0b011)
                {
                    Debug.Print("Invalid CompressedChunkSignature in VBA compressed stream");
                }

                // chunk flag = next bit - 1 == compressed, 0 == uncompressed
                int chunk_flag = (compressed_chunk_header >> 15) & 0x01;
                //log.debug("chunk size = {}, offset = {}, compressed flag = {}".format(chunk_size, compressed_chunk_start, chunk_flag))
                /*
                # MS-OVBA 2.4.1.3.12: the maximum size of a chunk including its header is 4098 bytes (header 2 + data 4096)
                # The minimum size is 3 bytes
                # NOTE: there seems to be a typo in MS-OVBA, the check should be with 4098, not 4095 (which is the max value
                # in chunk header before adding 3.
                # Also the first test is not useful since a 12 bits value cannot be larger than 4095.
                */
                if (chunk_flag == 1 && chunk_size > 4098) { Debug.Print("CompressedChunkSize=%d > 4098 but CompressedChunkFlag == 1' % chunk_size"); }
                if (chunk_flag == 0 && chunk_size != 4098) { Debug.Print("CompressedChunkSize=%d != 4098 but CompressedChunkFlag == 0' % chunk_size"); }

                //# check if chunk_size goes beyond the compressed data, instead of silently cutting it:
                //# TODO: raise an exception?
                if (compressed_chunk_start + chunk_size > compressedContainer.Length)
                {
                    Console.WriteLine("Chunk size is larger than remaining compressed data");
                }
                int compressedEnd = Math.Min(compressedContainer.Length, compressed_chunk_start + chunk_size);
                // read after chunk header:
                compressedCurrent = compressed_chunk_start + 2;


                if (chunk_flag == 0)
                {
                    // MS-OVBA 2.4.1.3.3 Decompressing a RawChunk
                    // uncompressed chunk: read the next 4096 bytes as-is
                    //TODO: check if there are at least 4096 bytes left

                    ////////decompressedContainer.AddRange(compressedContainer.GetRange(compressedCurrent, 4096));
                    compressedCurrent += 4096;
                }
                else
                {
                    // MS-OVBA 2.4.1.3.2 Decompressing a CompressedChunk
                    // compressed chunk
                    int decompressedChunkStart = decompressedContainer.Count();
                    while (compressedCurrent < compressedEnd)
                    {
                        byte flagByte = compressedContainer[compressedCurrent];
                        compressedCurrent++;
                        for (int bitIndex = 0; bitIndex < 8; bitIndex++)
                        {
                            // code logic here
                            if (compressedCurrent >= compressedEnd) break;
                            int flagBit = (flagByte >> bitIndex) & 1;
                            //log.Debug("bitIndex={0}: flagBit={1}", bitIndex, flagBit);
                            if (flagBit == 0) // LiteralToken
                            {
                                // copy one byte directly to output
                                decompressedContainer.Add(compressedContainer[compressedCurrent]);
                                compressedCurrent++;
                            }
                            else
                            {
                                //var copy_token = 999;//=struct.unpack("<H", compressed_container[compressed_current:compressed_current + 2])[0];
                                var copy_token = compressedContainer[compressedCurrent] + compressedContainer[compressedCurrent + 1] * 256;

                                // TODO: check this
                                int lengthMask, offsetMask, bitCount;
                                int length;
                                int temp1, temp2, offset;
                                int copySource;

                                //return (lengthMask, offsetMask, bitCount, maximumLength);
                                var result = CopyTokenHelp(decompressedContainer.Count, decompressedChunkStart);
                                lengthMask = result.Item1;
                                offsetMask = result.Item2;
                                bitCount = result.Item3;

                                length = (copy_token & lengthMask) + 3;
                                temp1 = copy_token & offsetMask;
                                temp2 = 16 - bitCount;
                                offset = (temp1 >> temp2) + 1;

                                // log.Debug($"offset={offset} length={length}");

                                copySource = decompressedContainer.Count - offset;

                                for (int index = copySource; index < copySource + length; index++)
                                {
                                    decompressedContainer.Add(decompressedContainer[index]);
                                }
                                compressedCurrent += 2;
                            }
                        }
                    }
                }


            }
            return decompressedContainer.ToArray();
        }
        private static int PatternAt(byte[] source, byte[] pattern)
        {
            // Too slow..... Don't use
            int ret = 0;
            for (int i = 0; i < source.Length; i++)
            {
                if (source.Skip(i).Take(pattern.Length).SequenceEqual(pattern))
                {
                    ret = i;
                }
            }
            return ret;
        }
        public static byte[] archiveVBAcodes(string filenameFullPath)
        {
            byte[] file_dat = File.ReadAllBytes(filenameFullPath);
            byte[] hash = archiveVBAcodes2(new MemoryStream(file_dat), filenameFullPath);

            return hash;
        }
        public static byte[] archiveVBAcodes2(MemoryStream fileStream, string filenameFullPath = "")
        {
            // Archive VBA codes into zip and also returns hash of vba code
            byte[] vbaCodeHash = null;

            using (ArchiveFile archiveFile = new ArchiveFile(fileStream))
            {
                foreach (Entry entry in archiveFile.Entries)
                {
                    //Console.WriteLine(entry.FileName);

                    // extract to file
                    //entry.Extract(entry.FileName);

                    if (entry.FileName == @"xl\vbaProject.bin")
                    {
                        // extract to stream
                        MemoryStream memoryStream = new MemoryStream();
                        entry.Extract(memoryStream);
                        memoryStream.Seek(0, SeekOrigin.Begin); // Need to set position to the start!!!

                        List<string> ProjectVBAItems = new List<string>();
                        List<string> moduleItems = new List<string>();
                        string[] PROJECT_file = new string[1];
                        //Compound File Binary Format
                        using (ArchiveFile archiveFile2 = new ArchiveFile(memoryStream, SevenZipFormat.Compound))
                        {
                            foreach (Entry entry2 in archiveFile2.Entries)
                            {
                                //Console.WriteLine(entry2.FileName);
                                ProjectVBAItems.Add(entry2.FileName);

                                MemoryStream memoryStream2 = new MemoryStream();
                                if (entry2.FileName == "PROJECT")
                                {
                                    entry2.Extract(memoryStream2);
                                    //memoryStream2.Seek(0, SeekOrigin.Begin); // Need to set position to the start!!!
                                    PROJECT_file = Encoding.UTF8.GetString(memoryStream2.ToArray()).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                                }
                            }
                            for (int i = 0; i < PROJECT_file.Length; i++)
                            {
                                string s = PROJECT_file[i];
                                if (s.StartsWith("Document=")) { moduleItems.Add(s.Substring(9, s.Length - 20)); } // e.g. Document=ThisWorkbook/&H00000000
                                if (s.StartsWith("Module=")) { moduleItems.Add(s.Substring(7, s.Length - 7)); }
                                if (s.StartsWith("Class=")) { moduleItems.Add(s.Substring(6, s.Length - 6)); }
                                if (s.StartsWith("BaseClass=")) { moduleItems.Add(s.Substring(10, s.Length - 10)); }
                                //if (s == "Name=\"VBAProject\"") { break; } // VBA project name can be changed from the default "VBAProject"
                                if (s.StartsWith("Name=")) { break; }
                            }
                            ////////////////////////////
                            using (var ms = new System.IO.MemoryStream())
                            {
                                using (var msAllCodes = new MemoryStream())
                                using (var zipArchive = new ZipArchive(ms, ZipArchiveMode.Create, true))
                                {
                                    foreach (var c in moduleItems)
                                    {
                                        MemoryStream memoryStream3 = new MemoryStream();
                                        // There is a case lowercase characters changed to uppercase automatically...
                                        int i = ProjectVBAItems.FindIndex(x => x.Equals("VBA\\" + c, StringComparison.OrdinalIgnoreCase));
                                        (archiveFile2.Entries[i]).Extract(memoryStream3);
                                        byte[] srcCode = MVG_decompress_stream(memoryStream3.ToArray());
                                        //System.IO.File.WriteAllBytes(c, xxx);
                                        byte[] srcWithOutAttributes = truncateAttributes(srcCode);
                                        msAllCodes.Write(srcWithOutAttributes, 0, srcWithOutAttributes.Length);
                                        //var tmpHash = new SHA256CryptoServiceProvider().ComputeHash(srcCode);

                                        var entry2 = zipArchive.CreateEntry(c);
                                        using (var es = entry2.Open())
                                        {
                                            es.Write(srcCode, 0, srcCode.Length);
                                        }
                                    }
                                    vbaCodeHash = new SHA256CryptoServiceProvider().ComputeHash(msAllCodes.ToArray());
                                }

                                if (filenameFullPath.Length > 0)
                                {
                                    string str = filenameFullPath;
                                    int n = str.LastIndexOf("\\");
                                    if (n >= 0) { str = str.Substring(n + 1); }
                                    string outFileFullPath = System.AppDomain.CurrentDomain.BaseDirectory + "(VBAcode)" + str + ".zip";
                                    System.IO.File.WriteAllBytes(outFileFullPath, ms.ToArray());
                                }
                            }
                            ////////////////////////////
                        }
                    }
                }
            }
            return vbaCodeHash;
        }
        /// <summary>
        /// Remove Attribute/VERSION, Begin/End headers
        /// </summary>
        /// <returns></returns>
        private static byte[] truncateAttributes(byte[] inBytes)
        {
            using (var ms1 = new MemoryStream())
            {
                var x = Encoding.UTF8.GetString(inBytes.ToArray()).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                for (int i = 0; i < x.Length; i++)
                {
                    if (x[i].IndexOf("Attribute ") == 0) continue;
                    if (x[i].IndexOf("VERSION ") == 0) continue;

                    // Skip BEGIN-END lines
                    if (x[i].Trim().ToUpper().IndexOf("BEGIN") == 0)
                    { //# class: "BEGIN","END"  frm:"Begin","End"
                        do
                        {
                            i++;
                        } while (x[i].Trim().ToUpper().IndexOf("END") == 0);
                        continue;
                    }


                    ms1.Write(Encoding.GetEncoding("UTF-8").GetBytes(x[i]), 0, x[i].Length);
                    ms1.WriteByte(0x0D); ms1.WriteByte(0x0A);
                }
                //var tmp = Encoding.UTF8.GetString(ms1.ToArray()).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                return ms1.ToArray();
            }
        }
    }
}
