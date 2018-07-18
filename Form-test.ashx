<%@ WebHandler Language="C#" Class="Handler" %>

using System;
using System.Web;
using System.IO;
using Newtonsoft.Json.Linq;
using TB_Check;
using System.Collections.Generic;
using System.Linq;
using System.Collections;
using System.Configuration;
using ICSharpCode.SharpZipLib;
using ICSharpCode.SharpZipLib.Checksums;
using ICSharpCode.SharpZipLib.Zip;

    public class Handler : IHttpHandler
    {
        string GUID = "";
        string UserName = "";
        string DBName = "";
        string CurrentTime = "";
        string InputFilePath = "";
        string OutFilePath = "";
        string outputExt = "";
        List<string> FileList = new List<string>();
        public void ProcessRequest(HttpContext context)
        {
            
            UserName = System.Web.HttpContext.Current.User.Identity.Name.ToString();
            CurrentTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            GUID = Guid.NewGuid().ToString("D");
            outputExt = context.Request["File_Ext"];
            try
            {
                uploadfileSave(context);
                For_OCR.Log.line = InputFilePath + "," + OutFilePath;
                For_OCR.Log.Write();
                OCRRun();
                OutputPackage();
                context.Response.Write("[{\"result\":\"success\",\"data\": \"\"}]");
                context.Response.End();
            }
            catch (Exception ex)
            {
                if (!(ex is System.Threading.ThreadAbortException))
                {
                    For_OCR.Log.line = ex.Message + ex.StackTrace;
                    For_OCR.Log.Write();
                    context.Response.Write("[{\"result\":\"fail\",\"data\": \"存在系统错误，请联系IDNET管理员寻求帮助！\"}]");
                    context.Response.End();
                }
                
            }
        }
        public void OCRRun()
        {
            For_OCR.OCR OCR = new For_OCR.OCR();
            OCR.inputPath = InputFilePath;
            OCR.outputPath = OutFilePath;
            OCR.outputExtension = outputExt;
            OCR.Files = FileList;
            OCR.Process();
        }
        public void uploadfileSave(HttpContext context)
        {
            HttpFileCollection hfc = context.Request.Files;
            string filepath = GetWebConfigValueByKey("OCRRequestFilePath");
            if(!Directory.Exists(Path.Combine(filepath,GUID)))
            {
                Directory.CreateDirectory(Path.Combine(filepath, GUID));
                InputFilePath = Path.Combine(filepath, GUID, "Input");
                Directory.CreateDirectory(InputFilePath);
                OutFilePath = Path.Combine(filepath, GUID, "Output");
                Directory.CreateDirectory(OutFilePath);
            }
            for (int i = 0; i < hfc.Count; i++)
            {
                HttpPostedFile file = hfc[i];
                string fullfilePath = Path.Combine(filepath, GUID, "Input", file.FileName);
                file.SaveAs(fullfilePath);
                if (Path.GetExtension(fullfilePath).ToLower() == ".zip")
                {
                    FileList = UnZip(fullfilePath, Path.Combine(filepath, GUID, "Input"));
                    File.Delete(fullfilePath);
                }
                else
                {
                    FileList.Add(fullfilePath);
                }
            }
        }
        public void OutputPackage()
        {
            string filepath = GetWebConfigValueByKey("OCRRequestFilePath");
            ZipDirectory(Path.Combine(filepath, GUID, @"\Output\"), Path.Combine(filepath, GUID, @"\Result.zip"), "");
        }
        public List<string> UnZip(string zipFilePath, string unZipDir)
        {
            List<string> FileList = new List<string>();
            if (zipFilePath == string.Empty)
            {
                throw new Exception("压缩文件不能为空！");
            }
            if (!File.Exists(zipFilePath))
            {
                throw new FileNotFoundException("压缩文件不存在！");
            }
            //解压文件夹为空时默认与压缩文件同一级目录下，跟压缩文件同名的文件夹  
            if (unZipDir == string.Empty)
                unZipDir = zipFilePath.Replace(Path.GetFileName(zipFilePath), Path.GetFileNameWithoutExtension(zipFilePath));
            if (!unZipDir.EndsWith("/"))
                unZipDir += "/";
            if (!Directory.Exists(unZipDir))
                Directory.CreateDirectory(unZipDir);

            using (var s = new ZipInputStream(File.OpenRead(zipFilePath)))
            {

                ZipEntry theEntry;
                while ((theEntry = s.GetNextEntry()) != null)
                {
                    string directoryName = Path.GetDirectoryName(theEntry.Name);
                    string fileName = Path.GetFileName(theEntry.Name);
                    if (!string.IsNullOrEmpty(directoryName))
                    {
                        Directory.CreateDirectory(unZipDir + directoryName);
                    }
                    if (directoryName != null && !directoryName.EndsWith("/"))
                    {
                    }
                    if (fileName != String.Empty)
                    {
                        using (FileStream streamWriter = File.Create(unZipDir + theEntry.Name))
                        {

                            int size;
                            byte[] data = new byte[2048];
                            while (true)
                            {
                                size = s.Read(data, 0, data.Length);
                                if (size > 0)
                                {
                                    streamWriter.Write(data, 0, size);
                                }
                                else
                                {
                                    break;
                                }
                            }
                            FileList.Add(Path.Combine(unZipDir, theEntry.Name));
                        }
                    }
                }
            }
            return FileList;
        }
        private static bool ZipDirectory(string folderToZip, ZipOutputStream zipStream, string parentFolderName)
        {
            bool result = true;
            string[] folders, files;
            ZipEntry ent = null;
            FileStream fs = null;
            try
            {
                ent = new ZipEntry(Path.Combine(Path.GetFileName(folderToZip) + "\\"));
                ent.IsUnicodeText = true;
                zipStream.PutNextEntry(ent);
                zipStream.Flush();

                files = Directory.GetFiles(folderToZip);
                foreach (string file in files)
                {
                    fs = File.OpenRead(file);

                    byte[] buffer = new byte[fs.Length];
                    fs.Read(buffer, 0, buffer.Length);
                    ent = new ZipEntry(Path.Combine(Path.GetFileName(folderToZip) + "\\" + Path.GetFileName(file)));
                    ent.IsUnicodeText = true;
                    ent.DateTime = DateTime.Now;
                    ent.Size = fs.Length;

                    fs.Close();
                    zipStream.PutNextEntry(ent);
                    zipStream.Write(buffer, 0, buffer.Length);
                }

            }
            catch
            {
                result = false;
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                    fs.Dispose();
                }
                if (ent != null)
                {
                    ent = null;
                }
                GC.Collect();
                GC.Collect(1);
            }

            folders = Directory.GetDirectories(folderToZip);
            foreach (string folder in folders)
                if (!ZipDirectory(folder, zipStream, folderToZip))
                    return false;

            return result;
        }
        public static bool ZipDirectory(string folderToZip, string zipedFile)
        {
            bool result = ZipDirectory(folderToZip, zipedFile, null);
            return result;
        }
        public static bool ZipDirectory(string folderToZip, string zipedFile, string password)
        {
            bool result = false;
            if (!Directory.Exists(folderToZip))
                return result;

            ZipOutputStream zipStream = new ZipOutputStream(File.Create(zipedFile));
            zipStream.SetLevel(6);
            if (!string.IsNullOrEmpty(password)) zipStream.Password = password;
            result = ZipDirectory(folderToZip, zipStream, "");
            zipStream.Finish();
            zipStream.Close();

            return result;
        }
        public static string GetWebConfigValueByKey(string key)
        {
            string value = string.Empty;
            Configuration config = System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration(HttpContext.Current.Request.ApplicationPath);
            AppSettingsSection appSetting = (AppSettingsSection)config.GetSection("appSettings");
            if (appSetting.Settings[key] != null)//如果不存在此节点，则添加  
            {
                value = appSetting.Settings[key].Value;
            }
            config = null;
            return value;
        }
        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
    public class CallBack
    {
        public string Result;
        public string Text;
    }

