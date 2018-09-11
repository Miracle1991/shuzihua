using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;
using MSEXCL = Microsoft.Office.Interop.Excel;
using System.Data;
using System.IO;
using System.Reflection;

namespace shuzihua
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.文件ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.设置默认数据文件路径ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.设置默认模板文件路径ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.帮助ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.module_file_btn = new System.Windows.Forms.Button();
            this.ofd = new System.Windows.Forms.OpenFileDialog();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.module_path_label = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.progress_bar_label = new System.Windows.Forms.Label();
            this.checkbg_fbd = new System.Windows.Forms.FolderBrowserDialog();
            this.savebg_folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.check_bg_ofd = new System.Windows.Forms.OpenFileDialog();
            this.save_bg_sfd = new System.Windows.Forms.SaveFileDialog();
            this.gen_btn = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.dataSet_file_btn = new System.Windows.Forms.Button();
            this.fbd = new System.Windows.Forms.FolderBrowserDialog();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.huizong_rbtn = new System.Windows.Forms.RadioButton();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.rbtn_cmm = new System.Windows.Forms.RadioButton();
            this.rbtn_word = new System.Windows.Forms.RadioButton();
            this.fbd1 = new System.Windows.Forms.FolderBrowserDialog();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.word_rbtn = new System.Windows.Forms.RadioButton();
            this.excel_rbtn_heng = new System.Windows.Forms.RadioButton();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.chaocha_rbtn = new System.Windows.Forms.RadioButton();
            this.gongcha_rbtn = new System.Windows.Forms.RadioButton();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.ref_rbtn = new System.Windows.Forms.CheckBox();
            this.RTF_CHAOCHA = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.excel_rbtn_shu = new System.Windows.Forms.RadioButton();
            this.menuStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.panel2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.文件ToolStripMenuItem,
            this.帮助ToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(909, 25);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // 文件ToolStripMenuItem
            // 
            this.文件ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.设置默认数据文件路径ToolStripMenuItem,
            this.设置默认模板文件路径ToolStripMenuItem});
            this.文件ToolStripMenuItem.Name = "文件ToolStripMenuItem";
            this.文件ToolStripMenuItem.Size = new System.Drawing.Size(44, 21);
            this.文件ToolStripMenuItem.Text = "设置";
            // 
            // 设置默认数据文件路径ToolStripMenuItem
            // 
            this.设置默认数据文件路径ToolStripMenuItem.Name = "设置默认数据文件路径ToolStripMenuItem";
            this.设置默认数据文件路径ToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.设置默认数据文件路径ToolStripMenuItem.Text = "设置默认数据文件路径";
            // 
            // 设置默认模板文件路径ToolStripMenuItem
            // 
            this.设置默认模板文件路径ToolStripMenuItem.Name = "设置默认模板文件路径ToolStripMenuItem";
            this.设置默认模板文件路径ToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.设置默认模板文件路径ToolStripMenuItem.Text = "设置默认模板文件路径";
            // 
            // 帮助ToolStripMenuItem
            // 
            this.帮助ToolStripMenuItem.Name = "帮助ToolStripMenuItem";
            this.帮助ToolStripMenuItem.Size = new System.Drawing.Size(44, 21);
            this.帮助ToolStripMenuItem.Text = "帮助";
            this.帮助ToolStripMenuItem.Click += new System.EventHandler(this.帮助ToolStripMenuItem_Click);
            // 
            // module_file_btn
            // 
            this.module_file_btn.Location = new System.Drawing.Point(316, 67);
            this.module_file_btn.Name = "module_file_btn";
            this.module_file_btn.Size = new System.Drawing.Size(98, 23);
            this.module_file_btn.TabIndex = 2;
            this.module_file_btn.Text = "选择";
            this.module_file_btn.UseVisualStyleBackColor = true;
            this.module_file_btn.Click += new System.EventHandler(this.button1_Click);
            // 
            // ofd
            // 
            this.ofd.InitialDirectory = "F:\\王东伟\\shuzihua_data";
            this.ofd.Title = "请选择要打开的文件";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.module_path_label);
            this.groupBox1.Controls.Add(this.module_file_btn);
            this.groupBox1.Location = new System.Drawing.Point(91, 228);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(730, 96);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "选择模板";
            // 
            // module_path_label
            // 
            this.module_path_label.AutoSize = true;
            this.module_path_label.Location = new System.Drawing.Point(20, 32);
            this.module_path_label.Name = "module_path_label";
            this.module_path_label.Size = new System.Drawing.Size(65, 12);
            this.module_path_label.TabIndex = 3;
            this.module_path_label.Text = "模板路径：";
            // 
            // progressBar1
            // 
            this.progressBar1.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.progressBar1.Location = new System.Drawing.Point(92, 509);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(721, 23);
            this.progressBar1.TabIndex = 4;
            // 
            // progress_bar_label
            // 
            this.progress_bar_label.AutoSize = true;
            this.progress_bar_label.Location = new System.Drawing.Point(448, 494);
            this.progress_bar_label.Name = "progress_bar_label";
            this.progress_bar_label.Size = new System.Drawing.Size(17, 12);
            this.progress_bar_label.TabIndex = 5;
            this.progress_bar_label.Text = "0%";
            // 
            // save_bg_sfd
            // 
            this.save_bg_sfd.Filter = "txt files(*.txt*)|*.txt|Excel工作簿(*.xlsx)|*.xlsx";
            this.save_bg_sfd.FileOk += new System.ComponentModel.CancelEventHandler(this.save_bg_sfd_FileOk);
            // 
            // gen_btn
            // 
            this.gen_btn.Location = new System.Drawing.Point(407, 547);
            this.gen_btn.Name = "gen_btn";
            this.gen_btn.Size = new System.Drawing.Size(98, 23);
            this.gen_btn.TabIndex = 8;
            this.gen_btn.Text = "生成报告";
            this.gen_btn.UseVisualStyleBackColor = true;
            this.gen_btn.Click += new System.EventHandler(this.gen_btn_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 589);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(909, 22);
            this.statusStrip1.TabIndex = 9;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(32, 17);
            this.toolStripStatusLabel1.Text = "状态";
            // 
            // dataSet_file_btn
            // 
            this.dataSet_file_btn.Location = new System.Drawing.Point(316, 62);
            this.dataSet_file_btn.Name = "dataSet_file_btn";
            this.dataSet_file_btn.Size = new System.Drawing.Size(98, 23);
            this.dataSet_file_btn.TabIndex = 10;
            this.dataSet_file_btn.Text = "选择";
            this.dataSet_file_btn.UseVisualStyleBackColor = true;
            this.dataSet_file_btn.Click += new System.EventHandler(this.button1_Click_2);
            // 
            // fbd
            // 
            this.fbd.RootFolder = System.Environment.SpecialFolder.MyComputer;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.dataSet_file_btn);
            this.groupBox2.Location = new System.Drawing.Point(91, 360);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(730, 100);
            this.groupBox2.TabIndex = 11;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "选择数据路径";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 11;
            this.label1.Text = "数据路径：";
            // 
            // huizong_rbtn
            // 
            this.huizong_rbtn.AutoSize = true;
            this.huizong_rbtn.Location = new System.Drawing.Point(11, 56);
            this.huizong_rbtn.Name = "huizong_rbtn";
            this.huizong_rbtn.Size = new System.Drawing.Size(47, 16);
            this.huizong_rbtn.TabIndex = 16;
            this.huizong_rbtn.TabStop = true;
            this.huizong_rbtn.Text = "汇总";
            this.huizong_rbtn.UseVisualStyleBackColor = true;
            this.huizong_rbtn.CheckedChanged += new System.EventHandler(this.huizong_rbtn_CheckedChanged);
            this.huizong_rbtn.Click += new System.EventHandler(this.huizong_rbtn_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(164, 44);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(657, 21);
            this.textBox1.TabIndex = 15;
            // 
            // rbtn_cmm
            // 
            this.rbtn_cmm.AutoSize = true;
            this.rbtn_cmm.Location = new System.Drawing.Point(10, 33);
            this.rbtn_cmm.Name = "rbtn_cmm";
            this.rbtn_cmm.Size = new System.Drawing.Size(65, 16);
            this.rbtn_cmm.TabIndex = 14;
            this.rbtn_cmm.TabStop = true;
            this.rbtn_cmm.Text = "读取CMM";
            this.rbtn_cmm.UseVisualStyleBackColor = true;
            this.rbtn_cmm.CheckedChanged += new System.EventHandler(this.rbtn_cmm_CheckedChanged);
            this.rbtn_cmm.Click += new System.EventHandler(this.rbtn_cmm_Click);
            // 
            // rbtn_word
            // 
            this.rbtn_word.AutoSize = true;
            this.rbtn_word.Checked = true;
            this.rbtn_word.Location = new System.Drawing.Point(10, 11);
            this.rbtn_word.Name = "rbtn_word";
            this.rbtn_word.Size = new System.Drawing.Size(65, 16);
            this.rbtn_word.TabIndex = 13;
            this.rbtn_word.TabStop = true;
            this.rbtn_word.Text = "读取RTF";
            this.rbtn_word.UseVisualStyleBackColor = true;
            this.rbtn_word.CheckedChanged += new System.EventHandler(this.rbtn_word_CheckedChanged);
            this.rbtn_word.Click += new System.EventHandler(this.rbtn_word_Click);
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 500;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.rbtn_word);
            this.panel1.Controls.Add(this.huizong_rbtn);
            this.panel1.Controls.Add(this.rbtn_cmm);
            this.panel1.Location = new System.Drawing.Point(6, 18);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(79, 82);
            this.panel1.TabIndex = 17;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(97, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 18;
            this.label2.Text = "报告名称：";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.panel1);
            this.groupBox4.Location = new System.Drawing.Point(307, 93);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(98, 109);
            this.groupBox4.TabIndex = 22;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "数据格式";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.word_rbtn);
            this.panel2.Controls.Add(this.excel_rbtn_shu);
            this.panel2.Controls.Add(this.excel_rbtn_heng);
            this.panel2.Location = new System.Drawing.Point(16, 20);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(135, 82);
            this.panel2.TabIndex = 20;
            // 
            // word_rbtn
            // 
            this.word_rbtn.AutoSize = true;
            this.word_rbtn.Checked = true;
            this.word_rbtn.Location = new System.Drawing.Point(10, 9);
            this.word_rbtn.Name = "word_rbtn";
            this.word_rbtn.Size = new System.Drawing.Size(71, 16);
            this.word_rbtn.TabIndex = 13;
            this.word_rbtn.TabStop = true;
            this.word_rbtn.Text = "Word文档";
            this.word_rbtn.UseVisualStyleBackColor = true;
            this.word_rbtn.CheckedChanged += new System.EventHandler(this.word_rbtn_CheckedChanged);
            // 
            // excel_rbtn_heng
            // 
            this.excel_rbtn_heng.AutoSize = true;
            this.excel_rbtn_heng.Location = new System.Drawing.Point(10, 31);
            this.excel_rbtn_heng.Name = "excel_rbtn_heng";
            this.excel_rbtn_heng.Size = new System.Drawing.Size(101, 16);
            this.excel_rbtn_heng.TabIndex = 14;
            this.excel_rbtn_heng.TabStop = true;
            this.excel_rbtn_heng.Text = "横版Excel文档";
            this.excel_rbtn_heng.UseVisualStyleBackColor = true;
            this.excel_rbtn_heng.CheckedChanged += new System.EventHandler(this.excel_rbtn_heng_CheckedChanged);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.panel2);
            this.groupBox3.Location = new System.Drawing.Point(92, 93);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(168, 109);
            this.groupBox3.TabIndex = 23;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "报告格式";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.chaocha_rbtn);
            this.groupBox5.Controls.Add(this.gongcha_rbtn);
            this.groupBox5.Enabled = false;
            this.groupBox5.Location = new System.Drawing.Point(460, 93);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(89, 109);
            this.groupBox5.TabIndex = 24;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "CMM输出";
            // 
            // chaocha_rbtn
            // 
            this.chaocha_rbtn.AutoSize = true;
            this.chaocha_rbtn.Checked = true;
            this.chaocha_rbtn.Location = new System.Drawing.Point(21, 30);
            this.chaocha_rbtn.Name = "chaocha_rbtn";
            this.chaocha_rbtn.Size = new System.Drawing.Size(47, 16);
            this.chaocha_rbtn.TabIndex = 17;
            this.chaocha_rbtn.TabStop = true;
            this.chaocha_rbtn.Text = "超差";
            this.chaocha_rbtn.UseVisualStyleBackColor = true;
            // 
            // gongcha_rbtn
            // 
            this.gongcha_rbtn.AutoSize = true;
            this.gongcha_rbtn.Location = new System.Drawing.Point(22, 66);
            this.gongcha_rbtn.Name = "gongcha_rbtn";
            this.gongcha_rbtn.Size = new System.Drawing.Size(47, 16);
            this.gongcha_rbtn.TabIndex = 18;
            this.gongcha_rbtn.TabStop = true;
            this.gongcha_rbtn.Text = "偏差";
            this.gongcha_rbtn.UseVisualStyleBackColor = true;
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.ref_rbtn);
            this.groupBox6.Controls.Add(this.RTF_CHAOCHA);
            this.groupBox6.Controls.Add(this.radioButton2);
            this.groupBox6.Location = new System.Drawing.Point(610, 93);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(211, 109);
            this.groupBox6.TabIndex = 25;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "RTF输出";
            // 
            // ref_rbtn
            // 
            this.ref_rbtn.AutoSize = true;
            this.ref_rbtn.Location = new System.Drawing.Point(107, 29);
            this.ref_rbtn.Name = "ref_rbtn";
            this.ref_rbtn.Size = new System.Drawing.Size(90, 16);
            this.ref_rbtn.TabIndex = 19;
            this.ref_rbtn.Text = "符号参考T值";
            this.ref_rbtn.UseVisualStyleBackColor = true;
            // 
            // RTF_CHAOCHA
            // 
            this.RTF_CHAOCHA.AutoSize = true;
            this.RTF_CHAOCHA.Checked = true;
            this.RTF_CHAOCHA.Location = new System.Drawing.Point(21, 30);
            this.RTF_CHAOCHA.Name = "RTF_CHAOCHA";
            this.RTF_CHAOCHA.Size = new System.Drawing.Size(47, 16);
            this.RTF_CHAOCHA.TabIndex = 17;
            this.RTF_CHAOCHA.TabStop = true;
            this.RTF_CHAOCHA.Text = "超差";
            this.RTF_CHAOCHA.UseVisualStyleBackColor = true;
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(22, 66);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(47, 16);
            this.radioButton2.TabIndex = 18;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "偏差";
            this.radioButton2.UseVisualStyleBackColor = true;
            // 
            // excel_rbtn_shu
            // 
            this.excel_rbtn_shu.AutoSize = true;
            this.excel_rbtn_shu.Location = new System.Drawing.Point(11, 54);
            this.excel_rbtn_shu.Name = "excel_rbtn_shu";
            this.excel_rbtn_shu.Size = new System.Drawing.Size(101, 16);
            this.excel_rbtn_shu.TabIndex = 16;
            this.excel_rbtn_shu.TabStop = true;
            this.excel_rbtn_shu.Text = "竖版Excel文档";
            this.excel_rbtn_shu.UseVisualStyleBackColor = true;
            this.excel_rbtn_shu.Visible = false;
            this.excel_rbtn_shu.CheckedChanged += new System.EventHandler(this.excel_rbtn_shu_CheckedChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(909, 611);
            this.Controls.Add(this.groupBox6);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.gen_btn);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.progress_bar_label);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "数字化软件";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion


        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 文件ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 帮助ToolStripMenuItem;
        private System.Windows.Forms.Button module_file_btn;
        private System.Windows.Forms.OpenFileDialog ofd;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label module_path_label;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label progress_bar_label;
        private System.Windows.Forms.FolderBrowserDialog checkbg_fbd;
        private System.Windows.Forms.FolderBrowserDialog savebg_folderBrowserDialog;
        private System.Windows.Forms.OpenFileDialog check_bg_ofd;
        private System.Windows.Forms.SaveFileDialog save_bg_sfd;
        private System.Windows.Forms.Button gen_btn;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.Button dataSet_file_btn;

        int data_line = 0;
        //缓存叶片信息
        private List<myBuffer> myPoints1 = new List<myBuffer>();
        //缓存文件列表
        private List<FileInfo> myFileList = new List<FileInfo>();
        //缓存文件头
        private List<myPoint> myPoints = new List<myPoint>();

        //word相关定义
        MSWord.Document doc = null;

        //DEV项在 .RTF中的位置 1为起始位
        private int rtf_dev_loc = 0;
        private int rtf_out_loc = 0;

        //
        private int num_rows = 0;

        //
        private int col_num = 1;
        //
        private string root;

        private bool contour_min_valid;
        private bool contour_max_valid;
        private bool cv_cont_min_valid;
        private bool cv_cont_max_valid;
        private bool cc_cont_min_valid;
        private bool cc_cont_max_valid;
        private bool le_contr_min_valid;
        private bool le_contr_max_valid;
        private bool te_contr_min_valid;
        private bool te_contr_max_valid;
        private bool stack_x_valid;
        private bool stack_y_valid;
        private bool twist_ang_valid;
        private bool te_redius_valid;
        private bool le_redius_valid;
        private bool le_aq_valid;
        private bool chord_wid_valid;
        private bool b_valid;
        private bool b1_valid;
        private bool extreme_valid;
        private bool le_1_5_valid;
        private bool te_1_5_valid;
        private bool qiexian_valid;

        private Form2 f2 = new Form2();
        //    
        private int file_num = 0;
        private double result = 0;

        private System.Windows.Forms.FolderBrowserDialog fbd;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ToolStripMenuItem 设置默认数据文件路径ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 设置默认模板文件路径ToolStripMenuItem;
        private System.Windows.Forms.FolderBrowserDialog fbd1;
        private Timer timer1;
        private RadioButton rbtn_cmm;
        private RadioButton rbtn_word;
        private TextBox textBox1;
        private RadioButton huizong_rbtn;
        private Label label2;
        private Panel panel1;
        private GroupBox groupBox4;
        private Panel panel2;
        private RadioButton word_rbtn;
        private RadioButton excel_rbtn_heng;
        private GroupBox groupBox3;
        private GroupBox groupBox5;
        private RadioButton chaocha_rbtn;
        private RadioButton gongcha_rbtn;
        private GroupBox groupBox6;
        private RadioButton RTF_CHAOCHA;
        private RadioButton radioButton2;
        private CheckBox ref_rbtn;
        private RadioButton excel_rbtn_shu;

       

    }
}

