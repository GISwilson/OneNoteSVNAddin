using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SharpSvn;
using SharpSvn.Security;
using Microsoft.Win32;

namespace OneNoteSVNAddin
{
    public partial class FrmSetting : Form
    {
        private SvnClient _client;

        public string ServerURL { get; private set; }
        public string UserName { get; private set; }
        public string Password { get; private set; }
        public string LocalPath { get; private set; }

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
                software.SetValue("ServerURL", ServerURL);
                software.SetValue("UserName", UserName);
                software.SetValue("Password", Password);
                software.SetValue("LocalPath", LocalPath);
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

        public FrmSetting(SvnClient client,string serverURL,string localPath,string userName,string password)
        {
            InitializeComponent();
            this._client = client;
            ServerURL = serverURL;
            LocalPath = localPath;
            UserName = userName;
            Password = password;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtURL.Text) || string.IsNullOrEmpty(txtUserName.Text)
                || string.IsNullOrEmpty(txtPassword.Text))
            {
                MessageBox.Show("请填写完整。");
                return;
            }
            ServerURL = txtURL.Text;
            UserName = txtUserName.Text;
            Password = txtPassword.Text;
            LocalPath = txtLocalPath.Text;
            //测试是否能连接成功
            try
            {
                SvnInfoEventArgs args;
                _client.GetInfo(SvnTarget.FromString(ServerURL), out args);
                MessageBox.Show("连接成功！");
            }
            catch (SvnAuthenticationException)
            {
                MessageBox.Show("用户名或密码错误");
                return;
            }
            catch(Exception)
            {
                MessageBox.Show("连接失败，请检测网络情况。");
                ServerURL = "";
                UserName = "";
                Password = "";
                LocalPath = "";
                return;
            }
            SaveMsg();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void FrmSetting_Load(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(LocalPath))
            {
                txtLocalPath.Text = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\OneNote 笔记本";
            }
            else
            {
                txtLocalPath.Text = LocalPath;
            }
            txtURL.Text = ServerURL;
            txtUserName.Text = UserName;
            txtPassword.Text = Password;
        }

       
    }
}
