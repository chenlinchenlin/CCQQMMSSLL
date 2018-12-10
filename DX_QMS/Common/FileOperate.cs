using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;

namespace DX_QMS.Common
{
    class FileOperate
    {
        /// <summary>
        /// 获取得文件后缀名
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public static string GetPostfixStr(string filename)
        {
            int start = filename.LastIndexOf(".");
            int length = filename.Length;
            string postfix = filename.Substring(start, length - start);
            return postfix;
        }

        /// <summary>
        /// 写文件
        /// </summary>
        /// <param name="Path"></param>
        /// <param name="Strings"></param>
        public static void WriteFile(string Path, string Strings)
        {

            if (!System.IO.File.Exists(Path))
            {
                System.IO.FileStream f = System.IO.File.Create(Path);
                f.Close();
                f.Dispose();
            }
            System.IO.StreamWriter f2 = new System.IO.StreamWriter(Path, true, System.Text.Encoding.UTF8);
            f2.WriteLine(Strings);
            f2.Close();
            f2.Dispose();

        }

        /// <summary>
        /// 拷贝文件
        /// </summary>
        /// <param name="OrignFile"></param>
        /// <param name="NewFile"></param>
        public static void FileCoppy(string OrignFile, string NewFile)
        {
            File.Copy(OrignFile, NewFile, true);
        }

        /// <summary>
        /// 删除文件
        /// </summary>
        /// <param name="Path"></param>
        public static void FileDel(string Path)
        {
            File.Delete(Path);
        }

        /// <summary>
        /// 移动文件
        /// </summary>
        /// <param name="OrignFile"></param>
        /// <param name="NewFile"></param>
        public static void FileMove(string OrignFile, string NewFile)
        {
            File.Move(OrignFile, NewFile);
        }

        /// <summary>
        /// 在当前目录下创建目录
        /// </summary>
        /// <param name="OrignFolder"></param>
        /// <param name="NewFloder"></param>
        public static void FolderCreate(string OrignFolder, string NewFloder)
        {
            Directory.SetCurrentDirectory(OrignFolder);
            Directory.CreateDirectory(NewFloder);
        }


        /// <summary>
        /// 递归删除文件夹目录及文件
        /// </summary>
        /// <param name="dir"></param>
        public static void DeleteFolder(string dir)
        {
            if (Directory.Exists(dir)) //如果存在这个文件夹删除之 
            {
                foreach (string d in Directory.GetFileSystemEntries(dir))
                {
                    if (File.Exists(d))
                        File.Delete(d); //直接删除其中的文件                        
                    else
                        DeleteFolder(d); //递归删除子文件夹 
                }
                Directory.Delete(dir, true); //删除已空文件夹                 
            }
        }


        /// <summary>
        /// 将指定文件夹下面的所有内容copy到目标文件夹下面 果目标文件夹为只读属性就会报错。
        /// </summary>
        /// <param name="srcPath"></param>
        /// <param name="aimPath"></param>
        public static void CopyDir(string srcPath, string aimPath)
        {
            try
            {
                // 检查目标目录是否以目录分割字符结束如果不是则添加之
                if (aimPath[aimPath.Length - 1] != Path.DirectorySeparatorChar)
                    aimPath += Path.DirectorySeparatorChar;
                // 判断目标目录是否存在如果不存在则新建之
                if (!Directory.Exists(aimPath))
                    Directory.CreateDirectory(aimPath);
                // 得到源目录的文件列表，该里面是包含文件以及目录路径的一个数组
                //如果你指向copy目标文件下面的文件而不包含目录请使用下面的方法
                //string[] fileList = Directory.GetFiles(srcPath);
                string[] fileList = Directory.GetFileSystemEntries(srcPath);
                //遍历所有的文件和目录
                foreach (string file in fileList)
                {
                    //先当作目录处理如果存在这个目录就递归Copy该目录下面的文件

                    if (Directory.Exists(file))
                        CopyDir(file, aimPath + Path.GetFileName(file));
                    //否则直接Copy文件
                    else
                        File.Copy(file, aimPath + Path.GetFileName(file), true);
                }
            }
            catch (Exception ee)
            {
                throw new Exception(ee.ToString());
            }
        }

        /// <summary>
        /// 获取指定文件夹下所有子目录及文件
        /// </summary>
        /// <param name="Path"></param>
        /// <returns></returns>
        public static string GetFoldAll(string Path)
        {

            string str = "";
            DirectoryInfo thisOne = new DirectoryInfo(Path);
            str = ListTreeShow(thisOne, 0, str);
            return str;
        }
        /// <summary>
        /// 获取指定文件夹下所有子目录及文件函数
        /// </summary>
        /// <param name="theDir"></param>
        /// <param name="nLevel"></param>
        /// <param name="Rn"></param>
        /// <returns></returns>
        public static string ListTreeShow(DirectoryInfo theDir, int nLevel, string Rn)//递归目录 文件
        {
            DirectoryInfo[] subDirectories = theDir.GetDirectories();//获得目录
            foreach (DirectoryInfo dirinfo in subDirectories)
            {

                if (nLevel == 0)
                {
                    Rn += "├";
                }
                else
                {
                    string _s = "";
                    for (int i = 1; i <= nLevel; i++)
                    {
                        _s += "│&nbsp;";
                    }
                    Rn += _s + "├";
                }
                Rn += "<b>" + dirinfo.Name.ToString() + "</b><br />";
                FileInfo[] fileInfo = dirinfo.GetFiles();   //目录下的文件
                foreach (FileInfo fInfo in fileInfo)
                {
                    if (nLevel == 0)
                    {
                        Rn += "│&nbsp;├";
                    }
                    else
                    {
                        string _f = "";
                        for (int i = 1; i <= nLevel; i++)
                        {
                            _f += "│&nbsp;";
                        }
                        Rn += _f + "│&nbsp;├";
                    }
                    Rn += fInfo.Name.ToString() + " <br />";
                }
                Rn = ListTreeShow(dirinfo, nLevel + 1, Rn);


            }
            return Rn;
        }

        /// <summary>
        ///  获取指定文件夹下所有子目录及文件(下拉框形)
        /// </summary>
        /// <param name="Path"></param>
        /// <param name="DropName"></param>
        /// <param name="tplPath"></param>
        /// <returns></returns>
        public static string GetFoldAll(string Path, string DropName, string tplPath)
        {
            string strDrop = "<select name=\"" + DropName + "\" id=\"" + DropName + "\"><option value=\"\">--请选择详细模板--</option>";
            string str = "";
            DirectoryInfo thisOne = new DirectoryInfo(Path);
            str = ListTreeShow(thisOne, 0, str, tplPath);
            return strDrop + str + "</select>";

        }

        /// <summary>
        /// 获取指定文件夹下所有子目录及文件函数
        /// </summary>
        /// <param name="theDir"></param>
        /// <param name="nLevel"></param>
        /// <param name="Rn"></param>
        /// <param name="tplPath"></param>
        /// <returns></returns>
        public static string ListTreeShow(DirectoryInfo theDir, int nLevel, string Rn, string tplPath)//递归目录 文件
        {
            DirectoryInfo[] subDirectories = theDir.GetDirectories();//获得目录

            foreach (DirectoryInfo dirinfo in subDirectories)
            {

                Rn += "<option value=\"" + dirinfo.Name.ToString() + "\"";
                if (tplPath.ToLower() == dirinfo.Name.ToString().ToLower())
                {
                    Rn += " selected ";
                }
                Rn += ">";

                if (nLevel == 0)
                {
                    Rn += "┣";
                }
                else
                {
                    string _s = "";
                    for (int i = 1; i <= nLevel; i++)
                    {
                        _s += "│&nbsp;";
                    }
                    Rn += _s + "┣";
                }
                Rn += "" + dirinfo.Name.ToString() + "</option>";


                FileInfo[] fileInfo = dirinfo.GetFiles();   //目录下的文件
                foreach (FileInfo fInfo in fileInfo)
                {
                    Rn += "<option value=\"" + dirinfo.Name.ToString() + "/" + fInfo.Name.ToString() + "\"";
                    if (tplPath.ToLower() == fInfo.Name.ToString().ToLower())
                    {
                        Rn += " selected ";
                    }
                    Rn += ">";

                    if (nLevel == 0)
                    {
                        Rn += "│&nbsp;├";
                    }
                    else
                    {
                        string _f = "";
                        for (int i = 1; i <= nLevel; i++)
                        {
                            _f += "│&nbsp;";
                        }
                        Rn += _f + "│&nbsp;├";
                    }
                    Rn += fInfo.Name.ToString() + "</option>";
                }
                Rn = ListTreeShow(dirinfo, nLevel + 1, Rn, tplPath);


            }
            return Rn;
        }



        public class filetoServer
        {
            // logon types  
            const int LOGON32_LOGON_INTERACTIVE = 2;
            const int LOGON32_LOGON_NETWORK = 3;
            const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
            // logon providers  
            const int LOGON32_PROVIDER_DEFAULT = 0;
            const int LOGON32_PROVIDER_WINNT50 = 3;
            const int LOGON32_PROVIDER_WINNT40 = 2;
            const int LOGON32_PROVIDER_WINNT35 = 1;

            [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern int LogonUser(String lpszUserName,
                String lpszDomain,
                String lpszPassword,
                int dwLogonType,
                int dwLogonProvider,
                ref IntPtr phToken);

            [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern int DuplicateToken(IntPtr hToken,
                int impersonationLevel,
                ref IntPtr hNewToken);

            [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern bool RevertToSelf();

            [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
            public static extern bool CloseHandle(IntPtr handle);

            private static WindowsImpersonationContext impersonationContext;

            public static bool impersonateValidUser(String userName, String domain, String password)
            {
                WindowsIdentity tempWindowsIdentity;
                IntPtr token = IntPtr.Zero;
                IntPtr tokenDuplicate = IntPtr.Zero;

                if (RevertToSelf())
                {
                    // 这里使用LOGON32_LOGON_NEW_CREDENTIALS来访问远程资源。  
                    // 如果要（通过模拟用户获得权限）实现服务器程序，访问本地授权数据库可  
                    // 以用LOGON32_LOGON_INTERACTIVE  
                    if (LogonUser(userName, domain, password, LOGON32_LOGON_NEW_CREDENTIALS,
                        LOGON32_PROVIDER_DEFAULT, ref token) != 0)
                    {
                        if (DuplicateToken(token, 2, ref tokenDuplicate) != 0)
                        {
                            tempWindowsIdentity = new WindowsIdentity(tokenDuplicate);
                            impersonationContext = tempWindowsIdentity.Impersonate();
                            if (impersonationContext != null)
                            {
                                System.AppDomain.CurrentDomain.SetPrincipalPolicy(PrincipalPolicy.WindowsPrincipal);
                                IPrincipal pr = System.Threading.Thread.CurrentPrincipal;
                                IIdentity id = pr.Identity;
                                CloseHandle(token);
                                CloseHandle(tokenDuplicate);
                                return true;
                            }
                        }
                    }
                }

                if (token != IntPtr.Zero)
                    CloseHandle(token);

                if (tokenDuplicate != IntPtr.Zero)
                    CloseHandle(tokenDuplicate);

                return false;
            }

            public static void undoImpersonation()
            {
                impersonationContext.Undo();
            }

            bool isImpersonated = false;

            public bool TestFunctionLogin()  //登陆  
            {
                isImpersonated = impersonateValidUser("waveaccess", @"\\10.0.xx.xxx", "waveaccessxxx");
                return isImpersonated;
            }

            public void TestFunctionCopyFile(string path, string spath, bool istrue)
            {
                File.Copy(path, spath, istrue);
            }

            public void TestFunctionQuit()
            {
                if (isImpersonated)
                {
                    undoImpersonation();
                }
            }

        }
    }
}
