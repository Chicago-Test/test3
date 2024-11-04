using OpenMcdf; //https://github.com/ironfede/openmcdf
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.IO.Pipes;
using System.Text;
using static System.Net.WebRequestMethods;


namespace unlockVBA
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string filenameFullPath = null;


            //filenameFullPath = @"..\..\..\Book1.xls";
            filenameFullPath = @"C:\Users\Administrator\source\repos\unlockVBA\VBA_password_abc.xlsm";

            if (Debugger.IsAttached == false)
            {
                if (args.Length > 0) { filenameFullPath = args[0]; }
                else
                {
                    Console.WriteLine("Example: >" + "unlockVBA" + " \"c:\\abc\\def.xlsm\"" + "   (xlsm,xlsb,xlam,xls)");
                    Environment.Exit(0);
                }
            }

            string fname = Path.GetFileName(filenameFullPath);
            string fileExtension_lowercase = Path.GetExtension(filenameFullPath).ToLower();
            string outFileName = "(unprotected VBA)" + fname;

            byte[] file_dat = System.IO.File.ReadAllBytes(filenameFullPath);

            if (fileExtension_lowercase.Equals(".xls"))
            {
                using (var memStream = testXLS(new MemoryStream(file_dat)))
                {
                    if (memStream.Length == 0)
                    {
                        Console.WriteLine("Error: VBA does not exist: " + fname);
                        Environment.Exit(-1);
                    }
                    // write memorystream to file
                    using (FileStream xlsToCreate = new FileStream(outFileName, FileMode.Create))
                    {
                        memStream.Position = 0;
                        memStream.WriteTo(xlsToCreate);
                    }
                }
            }
            else if (fileExtension_lowercase.Equals(".xlsm") || fileExtension_lowercase.Equals(".xlsb") || fileExtension_lowercase.Equals(".xlam"))
            {
                clearVBAprotection_XLSM(new MemoryStream(file_dat), outFileName);
            }
            Console.WriteLine("Done:" + fname);
            return;

            string filePathOfvbaProject = @"..\..\..\vbaProject.bin";
            file_dat = System.IO.File.ReadAllBytes(filePathOfvbaProject);
            generateUnprotected_vbaProject(new MemoryStream(file_dat));

        }
        //7za.exe x archive.zip -o outputdir a.xml -r
        // 7z e 1.xlsm xl\vbaproject.bin -aoa  (Overwrite All existing files without prompt)

        public static void clearVBAprotection_XLSM(MemoryStream fileStream, string outFileName)
        {
            // Input filestream -> filestream of *.xlsm
            byte[] bytes_new_vbaProject;
            using (var srcXLSM = new ZipArchive(fileStream, ZipArchiveMode.Update, false))
            {
                //foreach (var entry in zip.Entries) { var ss = entry.FullName; Console.WriteLine(ss); }
                var vbaProject_Entry = srcXLSM.GetEntry("xl/vbaProject.bin");
                if (vbaProject_Entry == null)
                {
                    Console.WriteLine("Error: xl/vbaProject.bin does not exist. Didn't output: " + outFileName);
                    Environment.Exit(-1);
                }

                //using (var memStream_vbaProject = new MemoryStream())
                using (var memStream = new MemoryStream())
                using (var wrapped_stream = vbaProject_Entry.Open())
                {
                    wrapped_stream.CopyTo(memStream);
                    var memStream_new_vbaProject = generateUnprotected_vbaProject(memStream);
                    bytes_new_vbaProject = memStream_new_vbaProject.ToArray();
                    //tmp.CopyTo(memStream_vbaProject);
                }

                vbaProject_Entry.Delete(); // Delete "xl/vbaProject.bin"

                using (var ms = new System.IO.MemoryStream())
                {
                    using (var zipArchive = new ZipArchive(ms, ZipArchiveMode.Create, true))
                    {
                        foreach (var file in srcXLSM.Entries)
                        {
                            // add files
                            var entry = zipArchive.CreateEntry(file.FullName);
                            using (var es = entry.Open())
                            {
                                var tmp_stream = file.Open();
                                tmp_stream.CopyTo(es);
                            }
                        }
                        // Add unprotected "xl/vbaProject.bin"
                        var entry2 = zipArchive.CreateEntry("xl/vbaProject.bin");
                        using (var es = entry2.Open())
                        {
                            es.Write(bytes_new_vbaProject, 0, bytes_new_vbaProject.Length);
                        }
                    }

                    // write memorystream to file
                    using (FileStream zipToCreate = new FileStream(outFileName, FileMode.Create))
                    {
                        //fileStream.Seek(0, SeekOrigin.Begin);
                        //fileStream.CopyTo(zipToCreate);
                        ms.Position = 0;
                        ms.WriteTo(zipToCreate);
                    }
                }
            }

            return;
        }

        static MemoryStream testXLS(MemoryStream memStream_XLS)
        {
            CompoundFile cf = new CompoundFile(memStream_XLS);
            MemoryStream retMemStream = new MemoryStream();

            //CFStorage storage_VBA_PROJECT_CUR;
            bool x = cf.RootStorage.TryGetStorage("_VBA_PROJECT_CUR", out CFStorage storage_VBA_PROJECT_CUR);
            if (x == false)
            {
                return retMemStream;
            }
            CFStream projectStream = storage_VBA_PROJECT_CUR.GetStream("PROJECT");
            //CFStream projectStream = cf.RootStorage.GetStorage("_VBA_PROJECT_CUR").GetStream("PROJECT");

            byte[] bytePROJECT = projectStream.GetData();
            byte[] bytesNewProject = updateProjectFile(bytePROJECT);
            // Replace "PROJECT" file
            cf.RootStorage.GetStorage("_VBA_PROJECT_CUR").Delete("PROJECT");
            CFStream myStream = cf.RootStorage.GetStorage("_VBA_PROJECT_CUR").AddStream("PROJECT");
            myStream.SetData(bytesNewProject);
            cf.Save(retMemStream);
            cf.Close();
            return retMemStream;
        }

        //static void createNew_vbaProject(string filePathOfvbaProject)
        //static void createNew_vbaProject(MemoryStream filePathOfvbaProject)
        static MemoryStream generateUnprotected_vbaProject(MemoryStream memStream_vbaProject)
        {
            //string filePathOfvbaProject = @"..\..\..\vbaProject.bin";

            // Extract "PROJECT" from vbaProject.bin as bytes
            CompoundFile cf = new CompoundFile(memStream_vbaProject);
            CFStream projectStream = cf.RootStorage.GetStream("PROJECT");
            byte[] bytePROJECT = projectStream.GetData();

            byte[] bytesNewProject = updateProjectFile(bytePROJECT);

            // Replace "PROJECT" file in vbaProject.bin
            cf.RootStorage.Delete("PROJECT");
            CFStream myStream = cf.RootStorage.AddStream("PROJECT");
            myStream.SetData(bytesNewProject);
            //cf.Commit();

            //cf.SaveAs("out_vbaProject.bin");
            MemoryStream retMemStream = new MemoryStream();
            cf.Save(retMemStream);

            cf.Close();
            return retMemStream;
        }
        static byte[] updateProjectFile(byte[] bytePROJECT)
        {
            // Convert bytes[] to line by line text
            string[] PROJECT_file = Encoding.UTF8.GetString(bytePROJECT).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

            // Replace properties. Unlock, no pass
            for (int i = 0; i < PROJECT_file.Length; i++)
            {
                string s = PROJECT_file[i];
                //https://github.com/outflanknl/EvilClippy/blob/master/evilclippy.cs
                if (s.StartsWith("ID=")) { PROJECT_file[i] = "ID=\"{595FFAAA-903C-4C82-8C80-B10F9F008836}\""; }
                if (s.StartsWith("CMG=")) { PROJECT_file[i] = "CMG=\"AFAD4D2053600664066406640664\""; }
                if (s.StartsWith("DPB=")) { PROJECT_file[i] = "DPB=\"5E5CBC91C4EF72F072F072\""; }
                if (s.StartsWith("GC=")) { PROJECT_file[i] = "GC=\"0D0FEF5E310C320C32F3\""; }
            }

            // Reconstruct "PROJECT"
            StringBuilder outPROJECT = new StringBuilder();
            for (int i = 0; i < PROJECT_file.Length; i++)
            {
                outPROJECT.AppendLine(PROJECT_file[i]);
            }
            var bytesNewProject = Encoding.GetEncoding("UTF-8").GetBytes(outPROJECT.ToString().ToArray());
            return bytesNewProject;
        }
    }
}
