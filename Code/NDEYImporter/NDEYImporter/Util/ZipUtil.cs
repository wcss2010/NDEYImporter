using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using ICSharpCode.SharpZipLib.Checksums;
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.GZip;

namespace NDEYImporter.Util
{
    /// <summary>
    /// 压缩文件格式
    /// </summary>
    public enum fileType
    {
        Zip = 1,
        RAR = 2,
        Tar = 3
    }
    /// <summary>
    /// 文件解压缩操作类库
    /// </summary>
    public class ZipTool
    {
        /// <summary>
        /// 压缩单个文件
        /// </summary>
        /// <param name="fileToZip">要压缩的文件</param>
        /// <param name="zipedFile">压缩后的文件</param>
        /// <param name="compressionLevel">压缩等级</param>
        /// <param name="blockSize">每次写入大小</param>
        public void ZipFile(string fileToZip, string zipedFile, int compressionLevel, int blockSize)
        {
            //如果文件没有找到，则报错
            if (!File.Exists(fileToZip))
            {
                throw new System.IO.FileNotFoundException("指定要压缩的文件: " + fileToZip + " 不存在!");
            }

            using (FileStream ZipFile = File.Create(zipedFile))
            {
                using (ZipOutputStream ZipStream = new ZipOutputStream(ZipFile))
                {
                    using (FileStream StreamToZip = new FileStream(fileToZip, FileMode.Open, FileAccess.Read))
                    {
                        string fileName = fileToZip.Substring(fileToZip.LastIndexOf("\\") + 1);
                        ZipEntry ZipEntry = new ZipEntry(fileName);
                        ZipStream.PutNextEntry(ZipEntry);
                        ZipStream.SetLevel(compressionLevel);
                        byte[] buffer = new byte[blockSize];
                        int sizeRead = 0;
                        try
                        {
                            do
                            {
                                sizeRead = StreamToZip.Read(buffer, 0, buffer.Length);
                                ZipStream.Write(buffer, 0, sizeRead);
                            }
                            while (sizeRead > 0);
                        }
                        catch (System.Exception ex)
                        {
                            throw ex;
                        }
                        StreamToZip.Close();
                    }

                    ZipStream.Finish();
                    ZipStream.Close();
                }
                ZipFile.Close();
            }
        }

        /// <summary>
        /// 压缩单个文件
        /// </summary>
        /// <param name="fileToZip">要进行压缩的文件名</param>
        /// <param name="zipedFile">压缩后生成的压缩文件名</param>
        public void ZipFile(string fileToZip, string zipedFile)
        {
            //如果文件没有找到，则报错
            if (!File.Exists(fileToZip))
            {
                throw new System.IO.FileNotFoundException("指定要压缩的文件: " + fileToZip + " 不存在!");
            }

            using (FileStream fs = File.OpenRead(fileToZip))
            {
                byte[] buffer = new byte[fs.Length];
                fs.Read(buffer, 0, buffer.Length);
                fs.Close();
                using (FileStream ZipFile = File.Create(zipedFile))
                {
                    using (ZipOutputStream ZipStream = new ZipOutputStream(ZipFile))
                    {
                        string fileName = fileToZip.Substring(fileToZip.LastIndexOf("\\") + 1);
                        ZipEntry ZipEntry = new ZipEntry(fileName);
                        ZipStream.PutNextEntry(ZipEntry);
                        ZipStream.SetLevel(5);

                        ZipStream.Write(buffer, 0, buffer.Length);
                        ZipStream.Finish();
                        ZipStream.Close();
                    }
                }
            }
        }

        /// <summary>
        /// 压缩多层目录
        /// </summary>
        /// <param name="strDirectory">要进行压缩的文件夹</param>
        /// <param name="zipedFile">压缩后生成的压缩文件名</param>
        public void ZipFileDirectory(string strDirectory, string zipedFile)
        {
            using (System.IO.FileStream ZipFile = System.IO.File.Create(zipedFile))
            {
                using (ZipOutputStream s = new ZipOutputStream(ZipFile))
                {
                    ZipSetp(strDirectory, s, "");
                }
            }
        }

        /// <summary>
        /// 递归遍历目录
        /// </summary>
        /// <param name="strDirectory">文件夹名称</param>
        /// <param name="s">The ZipOutputStream Object.</param>
        /// <param name="parentPath"></param>
        private static void ZipSetp(string strDirectory, ZipOutputStream s, string parentPath)
        {
            if (strDirectory[strDirectory.Length - 1] != Path.DirectorySeparatorChar)
            {
                strDirectory += Path.DirectorySeparatorChar;
            }
            Crc32 crc = new Crc32();

            string[] filenames = Directory.GetFileSystemEntries(strDirectory);

            foreach (string file in filenames)// 遍历所有的文件和目录
            {

                if (Directory.Exists(file))// 先当作目录处理如果存在这个目录就递归Copy该目录下面的文件
                {
                    string pPath = parentPath;
                    pPath += file.Substring(file.LastIndexOf("\\") + 1);
                    pPath += "\\";
                    ZipSetp(file, s, pPath);
                }

                else // 否则直接压缩文件
                {
                    //打开压缩文件
                    using (FileStream fs = File.OpenRead(file))
                    {

                        byte[] buffer = new byte[fs.Length];
                        fs.Read(buffer, 0, buffer.Length);

                        string fileName = parentPath + file.Substring(file.LastIndexOf("\\") + 1);
                        ZipEntry entry = new ZipEntry(fileName);

                        entry.DateTime = DateTime.Now;
                        entry.Size = fs.Length;

                        fs.Close();

                        crc.Reset();
                        crc.Update(buffer);
                        entry.Crc = crc.Value;
                        s.PutNextEntry(entry);

                        s.Write(buffer, 0, buffer.Length);
                    }
                }
            }
        }

        /// <summary>
        /// 解压缩一个 zip 文件。
        /// </summary>
        /// <param name="zipedFile">The ziped file.</param>
        /// <param name="strDirectory">The STR directory.</param>
        /// <param name="password">zip 文件的密码。</param>
        /// <param name="overWrite">是否覆盖已存在的文件。</param>
        public void UnZipFile(string zipedFile, string strDirectory, string password, bool overWrite)
        {

            if (strDirectory == "")
                strDirectory = Directory.GetCurrentDirectory();
            if (!strDirectory.EndsWith("\\"))
                strDirectory = strDirectory + "\\";

            using (ZipInputStream s = new ZipInputStream(File.OpenRead(zipedFile)))
            {
                s.Password = password;
                ZipEntry theEntry;

                while ((theEntry = s.GetNextEntry()) != null)
                {
                    string directoryName = "";
                    string pathToZip = "";
                    pathToZip = theEntry.Name;

                    if (pathToZip != "")
                        directoryName = Path.GetDirectoryName(pathToZip) + "\\";

                    string fileName = Path.GetFileName(pathToZip);

                    Directory.CreateDirectory(strDirectory + directoryName);

                    if (fileName != "")
                    {
                        if ((File.Exists(strDirectory + directoryName + fileName) && overWrite) || (!File.Exists(strDirectory + directoryName + fileName)))
                        {
                            using (FileStream streamWriter = File.Create(strDirectory + directoryName + fileName))
                            {
                                int size = 2048;
                                byte[] data = new byte[2048];
                                while (true)
                                {
                                    size = s.Read(data, 0, data.Length);

                                    if (size > 0)
                                        streamWriter.Write(data, 0, size);
                                    else
                                        break;
                                }
                                streamWriter.Close();
                            }
                        }
                    }
                }

                s.Close();
            }
        }
    }

    /// <summary>
    /// NDEY数据包释放工具（只释放*.db和*.rtf文件其它文件不管）
    /// </summary>
    public class NdeyMyDataUnZip
    {
        /// <summary>
        /// 解压缩一个 NDEY数据包(ZIP) 文件（只释放*.db文件其它文件不管）
        /// </summary>
        /// <param name="zipedFile">The ziped file.</param>
        /// <param name="strDirectory">The STR directory.</param>
        /// <param name="password">zip 文件的密码。</param>
        /// <param name="overWrite">是否覆盖已存在的文件。</param>
        public void UnZipFile(string zipedFile, string strDirectory, string password, bool overWrite)
        {
            //判断目标是否为空
            if (string.IsNullOrEmpty(strDirectory))
            {
                return;
            }

            //判断压缩文件是否存在
            if (File.Exists(zipedFile))
            {
                //载入ZIP文件
                ZipFile zip = new ZipFile(zipedFile);

                //解压myData.db文件
                try
                {
                    foreach (ZipEntry ze in zip)
                    {
                        //目标路径
                        string destPath = Path.Combine(strDirectory, ze.Name);

                        //检查当前项是否为文件
                        if (ze.IsFile)
                        {
                            //检查后缀名是否带有.db或.rtf
                            if (ze.Name.EndsWith(".db") || ze.Name.EndsWith(".rtf"))
                            {
                                //获得该文件输入流
                                Stream s = zip.GetInputStream(ze);

                                try
                                {
                                    //写文件
                                    using (FileStream streamWriter = File.Create(destPath))
                                    {
                                        int size = 2048;
                                        byte[] data = new byte[2048];
                                        try
                                        {
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
                                        }
                                        finally
                                        {
                                            streamWriter.Close();
                                        }
                                    }
                                }
                                finally
                                {
                                    s.Close();
                                }
                            }
                        }
                        else
                        {
                            try
                            {
                                Directory.CreateDirectory(Path.Combine(strDirectory, ze.Name));
                            }
                            catch (Exception ex) { }
                        }
                    }
                }
                catch (Exception ex)
                {
                    //Error
                }
                finally
                {
                    zip.Close();
                }
            }
        }
    }

    /// <summary>
    /// NDEY数据包释放工具（只释放*.db,*.doc和附件文件其它文件不管）
    /// </summary>
    public class NdeyDocFilesUnZip
    {
        /// <summary>
        /// 解压缩一个 NDEY数据包(ZIP) 文件（只释放*.db文件其它文件不管）
        /// </summary>
        /// <param name="zipedFile">The ziped file.</param>
        /// <param name="strDirectory">The STR directory.</param>
        /// <param name="password">zip 文件的密码。</param>
        /// <param name="overWrite">是否覆盖已存在的文件。</param>
        public void UnZipFile(string zipedFile, string strDirectory, string password, bool overWrite)
        {
            //判断目标是否为空
            if (string.IsNullOrEmpty(strDirectory))
            {
                return;
            }

            //判断压缩文件是否存在
            if (File.Exists(zipedFile))
            {
                //载入ZIP文件
                ZipFile zip = new ZipFile(zipedFile);

                //解压myData.db文件
                try
                {
                    foreach (ZipEntry ze in zip)
                    {
                        //目标路径
                        string destPath = Path.Combine(strDirectory, ze.Name);

                        //检查当前项是否为文件
                        if (ze.IsFile)
                        {
                            //检查后缀名是否带有.db或.rtf,如果是则跳过
                            if (!ze.Name.EndsWith(".rtf"))
                            {
                                //获得该文件输入流
                                Stream s = zip.GetInputStream(ze);

                                try
                                {
                                    //写文件
                                    using (FileStream streamWriter = File.Create(destPath))
                                    {
                                        int size = 2048;
                                        byte[] data = new byte[2048];
                                        try
                                        {
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
                                        }
                                        finally
                                        {
                                            streamWriter.Close();
                                        }
                                    }
                                }
                                finally
                                {
                                    s.Close();
                                }
                            }
                        }
                        else
                        {
                            try
                            {
                                Directory.CreateDirectory(Path.Combine(strDirectory, ze.Name));
                            }
                            catch (Exception ex) { }
                        }
                    }
                }
                catch (Exception ex)
                {
                    //Error
                }
                finally
                {
                    zip.Close();
                }
            }
        }
    }
}