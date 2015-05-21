using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.OneNote;
using System.Windows.Forms;
using System.Runtime.InteropServices.ComTypes;
using System.IO;
using OneNoteSVNAddin.Properties;
using System.Drawing.Imaging;
using System.Reflection;
using System.Drawing;
using SharpSvn;
using SharpSvn.Security;
using System.Xml.Linq;
using Microsoft.Win32;

namespace OneNoteSVNAddin
{
    [GuidAttribute("C1AA601C-9FBD-4DDD-9B2D-01CF11DBD56D"), ProgId("OneNoteSVNAddin.SVNAddin")]
    public class SVNAddin : IDTExtensibility2, IRibbonExtensibility
    {
        #region IDTExtensibility2 成员

        ApplicationClass onApp = new ApplicationClass();

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            /*
                For debugging, it is useful to have a MessageBox.Show() here, so that execution is paused while you have a chance to get VS to 'Attach to Process' 
            */
            onApp = (ApplicationClass)Application;
        }
        public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
        {
            //Clean up. Application is closing
            onApp = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        public void OnBeginShutdown(ref System.Array custom)
        {
            if (onApp != null)
                onApp = null;
        }
        public void OnStartupComplete(ref Array custom) { }
        public void OnAddInsUpdate(ref Array custom) { }

        #endregion

        #region IRibbonExtensibility 成员

        /// <summary>
        /// Called at the start of the running of the add-in. Loads the ribbon
        /// </summary>
        public string GetCustomUI(string RibbonID)
        {
            return OneNoteSVNAddin.Properties.Resources.ribbon;
        }

        #endregion

        private SvnClient _client;//SVN客户端实例
        private string _serverURL;//SVN服务端地址
        private string _userName;//本地连接用户名
        private string _password;//本地连接密码（明文）
        private string _localPath;//本地工作路径

        public SVNAddin()
        {
            _client = new SvnClient();
            _serverURL = "";
            _userName = "";
            _password = "";
            _localPath = "";
            GetMsg();
        }

        /// <summary>
        /// 同步设置
        /// Called from the onAction="" parameter in ribbon.xml. This is effectivley the onClick() function
        /// </summary>
        /// <param name="control">The control that was just clicked. control.Id will give you its ID</param>
        public void ShowSetting(IRibbonControl control)
        {
            Microsoft.Office.Interop.OneNote.Window context = control.Context as Microsoft.Office.Interop.OneNote.Window;
            CWin32WindowWrapper owner =
                new CWin32WindowWrapper((IntPtr)context.WindowHandle);

            FrmSetting frm = new FrmSetting(_client,_serverURL,_localPath,_userName,_password);
            if (frm.ShowDialog(owner) == DialogResult.OK)
            {
                _serverURL = frm.ServerURL;
                _userName = frm.UserName;
                _password = frm.Password;
                _localPath = frm.LocalPath;
            }
        }

        /// <summary>
        /// 更新操作
        /// </summary>
        /// <param name="control"></param>
        public void Update(IRibbonControl control)
        {
            try
            {
                System.Collections.ObjectModel.Collection<SvnStatusEventArgs> list = null;
                _client.CleanUp(_localPath);
                SvnStatusArgs args = new SvnStatusArgs();
                args.Depth = SvnDepth.Infinity;
                //args.RetrieveAllEntries = true;
                args.RetrieveRemoteStatus = true;
                _client.GetStatus(_localPath, args, out list);

                foreach (var statusItem in list)
                {
                    if (statusItem.LocalContentStatus == SvnStatus.Normal)
                    {
                        continue;
                    }
                    else if (statusItem.LocalContentStatus == SvnStatus.Modified)
                    {
                        try
                        {
                            try
                            {
                                File.Delete(statusItem.FullPath);
                            }
                            catch
                            {
                                MessageBox.Show(string.Format("文件{0}被占用，请关闭相关程序再重试。", statusItem.FullPath));
                                return;
                            }
                            _client.Update(statusItem.FullPath);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(string.Format("更新错误，错误原因:\n{0}", ex.Message));
                        }
                    }
                    else if (statusItem.LocalContentStatus == SvnStatus.Missing)
                    {
                        try
                        {
                            _client.Update(statusItem.FullPath);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(string.Format("更新错误，错误原因:\n{0}", ex.Message));
                        }
                    }
                }

                MessageBox.Show("更新成功！");
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("更新错误，错误原因:\n{0}", ex.Message));
            }
        }

        /// <summary>
        /// 同步至服务器操作
        /// </summary>
        /// <param name="control"></param>
        public void Commit(IRibbonControl control)
        {
            DirectoryInfo dirInfo = new DirectoryInfo(_localPath);
            if (dirInfo.GetDirectories(".svn").Count() == 0)
            {
                MessageBox.Show(string.Format("当前目录{0}没有配置文件，请先服务器的笔记本导出至该文件夹（记得先备份当前目录的笔记本）。", _localPath));
                return;
            }

            try
            {
                //查找状态信息
                System.Collections.ObjectModel.Collection<SvnStatusEventArgs> list = null;
                _client.CleanUp(_localPath);
                SvnStatusArgs args = new SvnStatusArgs();
                args.Depth = SvnDepth.Infinity;
                //args.RetrieveAllEntries = true;
                args.RetrieveRemoteStatus = true;
                _client.GetStatus(_localPath, args, out list);

                SvnImportArgs importArgs = new SvnImportArgs();
                SvnCommitArgs commitArgs = new SvnCommitArgs();
                SvnDeleteArgs deleteArgs = new SvnDeleteArgs();

                foreach (var statusItem in list)
                {
                    if (statusItem.LocalContentStatus == SvnStatus.Normal)
                    {
                        continue;
                    }
                    else if (statusItem.LocalContentStatus == SvnStatus.NotVersioned)
                    {
                        string addToUrl = statusItem.FullPath.Substring(_localPath.Length);
                        importArgs.LogMessage = "导入时间：" + DateTime.Now.ToString();
                        try
                        {
                            _client.RemoteImport(statusItem.FullPath, new Uri(_serverURL + addToUrl), importArgs);
                            //_client.Import(statusItem.FullPath, new Uri(URL + addToUrl), importArgs);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(string.Format("导入错误，错误原因:\n{0}", ex.Message));
                        }
                    }
                    else if (statusItem.LocalContentStatus == SvnStatus.Modified)
                    {
                        commitArgs.LogMessage = "提交时间：" + DateTime.Now.ToString();
                        try
                        {
                            _client.Commit(statusItem.FullPath, commitArgs);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(string.Format("修改错误，错误原因:\n{0}", ex.Message));
                        }
                    }
                    else if (statusItem.LocalContentStatus == SvnStatus.Missing)
                    {
                        deleteArgs.LogMessage = "删除时间：" + DateTime.Now.ToString();
                        try
                        {
                            _client.RemoteDelete(statusItem.Uri, deleteArgs);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(string.Format("删除错误，错误原因:\n{0}", ex.Message));
                        }
                    }
                }

                MessageBox.Show("同步成功！");
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("同步失败，错误原因:\n{0}", ex.Message));
            }
        }

        //public void Import(IRibbonControl control)
        //{
        //    Dictionary<string, string> notebookPaths = new Dictionary<string, string>();

        //    ApplicationClass onApp = new ApplicationClass();
        //    string notebookXml;
        //    onApp.GetHierarchy(null, HierarchyScope.hsPages, out notebookXml);
        //    var doc = XDocument.Parse(notebookXml);//doc.Descendants().FirstOrDefault().Elements()
        //    if (doc.Descendants().FirstOrDefault() != null)
        //    {
        //        var notebooks = doc.Descendants().FirstOrDefault().Elements();
        //        foreach (var notebook in notebooks)
        //        {
        //            if (notebook.Attributes().Count() > 1)//排除未归档笔记
        //            {
        //                string name = notebook.Attribute("name").Value;
        //                string path = notebook.Attribute("path").Value;
        //                if (notebookPaths.ContainsKey(name))
        //                {
        //                    MessageBox.Show("当前笔记本中部分笔记本名称重复，请先关闭这些笔记本。");
        //                    return;
        //                }
        //                else
        //                {
        //                    notebookPaths.Add(name, path);
        //                }
        //            }
        //        }
        //    }
        //    //
        //    System.Collections.ObjectModel.Collection<SvnListEventArgs> list;
        //    SvnPropertyListArgs args = new SvnPropertyListArgs();
        //    SvnTarget target = SvnTarget.FromString(_serverURL);
        //    _client.GetList(target, out list);
        //    var lists = list.Select(item => item.Name);

        //    SvnImportArgs importArgs = new SvnImportArgs();
        //    importArgs.LogMessage = "导入时间：" + DateTime.Now.ToString();
        //    foreach (KeyValuePair<string, string> pair in notebookPaths)
        //    {
        //        if (lists.Contains(pair.Key))
        //        {
        //            MessageBox.Show(string.Format("服务器上已存在{0}笔记本，将跳过导入该笔记本。", pair.Key));
        //            continue;
        //        }
        //        try
        //        {
        //            _client.Import(pair.Value, new Uri("https://user-PC/svn/OneNoteRepositoryTest/" + pair.Key), importArgs);
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(string.Format("提交笔记本{0}出错，错误信息\n{1}", pair.Key, ex.Message));
        //        }
        //    }
        //    MessageBox.Show("导入至服务器完成！");
        //}

        //public void Export(IRibbonControl control)
        //{
        //    Microsoft.Office.Interop.OneNote.Window context = control.Context as Microsoft.Office.Interop.OneNote.Window;
        //    CWin32WindowWrapper owner =
        //        new CWin32WindowWrapper((IntPtr)context.WindowHandle);
        //    FolderBrowserDialog dialog = new FolderBrowserDialog();
        //    if (dialog.ShowDialog(owner) == DialogResult.OK)
        //    {
        //        try
        //        {
        //            _client.CheckOut(SvnUriTarget.FromString(_serverURL), dialog.SelectedPath);
        //            MessageBox.Show("导出至本地成功！");
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(string.Format("导出服务器的笔记本出错，错误信息\n{0}", ex.InnerException.Message));
        //        }
        //    }
        //}


        /// <summary>
        /// 关于操作
        /// </summary>
        /// <param name="control"></param>
        public void About(IRibbonControl control)
        {
            Microsoft.Office.Interop.OneNote.Window context = control.Context as Microsoft.Office.Interop.OneNote.Window;
            CWin32WindowWrapper owner =
                new CWin32WindowWrapper((IntPtr)context.WindowHandle);

            FrmAbout frm = new FrmAbout();
            frm.ShowDialog(owner);
        }

        /// <summary>
        /// 根据图片名获取IStream流
        /// Called from the loadImage="" parameter in ribbon.xml. Converts the images into IStreams
        /// </summary>
        /// <param name="imageName">The image="" parameter in ribbon.xml, i.e. the image name</param>
        public IStream GetImage(string imageName)
        {
            MemoryStream mem = new MemoryStream();
            //BindingFlags flags = BindingFlags.Static | BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

            //var b = typeof(Properties.Resources).GetProperty(imageName.Substring(0, imageName.IndexOf('.')), flags).GetValue(null, null) as Bitmap;
            //b.Save(mem, ImageFormat.Png);
            switch (imageName)
            {
                case "Settings.png":
                    Resources.Settings.Save(mem, ImageFormat.Png);
                    break;
                case "Update.png":
                    Resources.Update.Save(mem, ImageFormat.Png);
                    break;
                case "Commit.png":
                    Resources.Commit.Save(mem, ImageFormat.Png);
                    break;
                case "Import.png":
                    Resources.Import.Save(mem, ImageFormat.Png);
                    break;
                case "Export.png":
                    Resources.Export.Save(mem, ImageFormat.Png);
                    break;
                case "About.png":
                    Resources.About.Save(mem, ImageFormat.Png);
                    break;
                default:
                    break;
            }

            return new CCOMStreamWrapper(mem);
        }

        #region 在注册表中读取和存储数据
        private const string _Register_Key = @"software\GISwilson\OneNoteSVNAddin";
        private bool SaveMsg()
        {
            bool b = true;
            RegistryKey key = Registry.LocalMachine;
            RegistryKey software = null;
            software = key.OpenSubKey(_Register_Key, true);
            if (software == null)
            {
                software = key.CreateSubKey(_Register_Key, RegistryKeyPermissionCheck.ReadWriteSubTree);
            }

            try
            {
                software.SetValue("ServerURL", _serverURL);
                software.SetValue("UserName", _userName);
                software.SetValue("Password", RSASecurity.EncryptRSA(_password));
                software.SetValue("LocalPath", _localPath);
            }
            catch
            {
                b = false;
            }
            finally
            {
                key.Close();
            }

            return b;
        }

        private void GetMsg()
        {
            RegistryKey key = Registry.LocalMachine;
            RegistryKey software = key.OpenSubKey(_Register_Key, true);
            try
            {
                _userName = software.GetValue("UserName").ToString();
                _password = software.GetValue("Password").ToString();
                try
                {
                    //如果未存储密文，则可能报错
                    _password = RSASecurity.DecryptRSA(_password);
                }
                catch { }
                _serverURL = software.GetValue("ServerURL").ToString();
                _localPath = software.GetValue("LocalPath").ToString();
            }
            catch
            {
            }
        }
        #endregion
    }
}
