using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using Microsoft.Win32;
using System.Diagnostics;
using System.Windows.Forms;

namespace ZhuoMian
{
    public partial class Form1 : Form
    {
        // 桌面文件夹设置注册表路径（相对于CurrentUser）
        const string REG_DESKTOP = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders";
        // 文件夹上添加鼠标右键菜单（相对于ClassesRoot）
        const string REG_DIRECTORY = "Directory\\shell";
        // 添加右键菜单复选框设置文件名
        // 文件存在表示复选框未选择
        // 文件不存在表示复选框选中
        // 因为默认为需要添加右键菜单
        private string CHECKED_FILE_NAME = Environment.CurrentDirectory+Path.DirectorySeparatorChar+".zhuomian_menu_cancel";
        // 启动时右键菜单复选框选择情况
        private bool initChecked = false;
        public Form1()
        {
            InitializeComponent();
        }
        /**
         * 生效桌面文件夹
         * 并根据添加右键菜单设置情况添加或者删除鼠标右键菜单（以修改注册表的方式）
         */
        private void button1_Click(object sender, EventArgs e)
        {
            // 桌面文件夹切换
            string directory = this.textBox1.Text;
            this.changeDektop(directory);
            // 右键菜单处理，与初始状态有变化时才处理
            if (this.checkBox1.Checked != this.initChecked)
            {
                if(this.checkBox1.Checked)
                {
                    this.addMenu();
                    this.deleteMenuChecked();
                }
                else
                {
                    this.removeMenu();
                    this.setMenuChecked();
                }
                this.initChecked = this.checkBox1.Checked;
            }
        }
            
        /**
         * 窗体加载成功后
         * 1. 需要给文件夹输入框赋初始值
         * 2. 给右键菜单复选框赋初始值
         * 3. 根据右键菜单初始值设置或者删除鼠标右键菜单
         */
        private void Form1_Load(object sender, EventArgs e)
        {
            // 桌面文件夹输入框的默认值为当前桌面所在文件夹
            string directory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            this.textBox1.Text = directory;
            this.checkBox1.Checked = this.addMenuChecked();
            // 右键菜单复选框初值
            this.initChecked = this.checkBox1.Checked;
            // 添加或者删除右键菜单
            if (this.initChecked)
            {
                this.addMenu();
            }
            else
            {
                this.removeMenu();
            }
        }
        /**
         * 恢复到系统默认桌面文件夹
         * 即当前用户文件夹下Desktop文件夹
         */
        private void button3_Click(object sender, EventArgs e)
        {
            // TODO: 当前用户文件夹不在C盘的情况兼容
            string directory = "C:"+ Path.DirectorySeparatorChar+"Users"+Path.DirectorySeparatorChar+Environment.UserName+Path.DirectorySeparatorChar+"Desktop";
            this.textBox1.Text = directory;
            this.changeDektop(directory);
        }
        /**
         * 选择文件夹
         */
        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择桌面文件夹";
            // 默认为当前选择的文件夹
            dialog.SelectedPath = this.textBox1.Text;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                this.textBox1.Text = dialog.SelectedPath;
            }
        }

        /**
         * 通过修改系统环境变量的方式修改桌面对应文件夹
         */
        private void changeDektop(string directory)
        {
            RegistryKey key = Registry.CurrentUser;
            RegistryKey desktop = key.OpenSubKey(REG_DESKTOP, true);
            desktop.SetValue("Desktop", directory);
            // 重启资源管理器，以刷新桌面
            Process process = Process.GetProcessesByName("explorer")[0];
            process.Kill();
        }
        /**
         * 判断添加右键菜单复选框是否选中
         * 判断依据是当前文件夹下是否存在指定文件
         * 存在则视为未选中
         * 其他情况都是为选中
         */
        private bool addMenuChecked()
        {
            if (File.Exists(CHECKED_FILE_NAME))
            {
                return false;
            }
            return true;
        }
        private void setMenuChecked()
        {
            try
            {
                File.CreateText(CHECKED_FILE_NAME);
                // 为了不被误删，设置文件为隐藏文件
                FileInfo fileInfo = new FileInfo(CHECKED_FILE_NAME);
                if (fileInfo.Exists)
                {
                    fileInfo.Attributes = FileAttributes.Hidden;
                }
            }catch(Exception e)
            {
                ;
            }
        }
        private void deleteMenuChecked()
        {
            try
            {
                File.Delete(CHECKED_FILE_NAME);
            }catch(Exception e)
            {
                ;
            }
        }
        private void addMenu()
        {
            // 不管当前有没有添加过右键菜单，先删除一次
            // 因为可能可执行文件的路径发生了改变
            this.removeMenu();
            // 开始添加右键菜单
            // 菜单名称，&Z表示设置运行快捷键为Z
            string menuName = "设置为桌面(&Z)";
            // 当前可执行文件路径
            string exe = Application.ExecutablePath;
            RegistryKey key = Registry.ClassesRoot;

            // 在文件夹上右键添加菜单
            RegistryKey reg = key.OpenSubKey(REG_DIRECTORY, true);
            reg.CreateSubKey("ZhuoMian");
            reg = reg.OpenSubKey("ZhuoMian",true);
            // 设置菜单名称
            reg.SetValue("", menuName);
            // 设置图标
            reg.SetValue("Icon", "\"" + exe + "\"");
            reg.CreateSubKey("command");
            // 设置可执行文件
            reg = reg.OpenSubKey("command", true);
            reg.SetValue("", "\""+exe+"\" %1");
        }

        /**
         * 比较暴力但是方便的移除注册表项的方式
         * 注册表项不存在，移除时会抛异常，此时直接忽略
         */
        private void removeMenu()
        {
            try
            {
                RegistryKey key = Registry.ClassesRoot;
                RegistryKey reg = key.OpenSubKey(REG_DIRECTORY, true);
                reg.DeleteSubKeyTree("ZhuoMian");
            }
            catch (Exception e)
            {
                ;
            }
        }
    }
}
