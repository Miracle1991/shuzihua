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
using MSExcel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;


namespace shuzihua
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.timer1.Start();
            this.root = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
        }

       ~Form1()
        {
            //wordApp.Quit();
        }

        private void openFileDialog2_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.ofd.ShowDialog() == DialogResult.OK)
            {
                this.module_path_label.Text = "模板路径：" + this.ofd.FileName;
                int l = this.ofd.FileName.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries).Length;
                this.textBox1.Text = this.ofd.FileName.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries)[l - 1];
                this.textBox1.Text = this.textBox1.Text.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries)[0];
                this.toolStripStatusLabel1.Text = "成功加载模板";
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (this.save_bg_sfd.ShowDialog() == DialogResult.OK)
            {
                StreamWriter sw = new StreamWriter(save_bg_sfd.FileName, false, Encoding.Default);
                sw.Write("wdw\n wangdongwei");
                sw.Close();
            }
        }

        private void check_bg_Click(object sender, EventArgs e)
        {
            if (this.check_bg_ofd.ShowDialog() == DialogResult.OK)
            {
                MSWord.Application app = new MSWord.Application();
                MSWord.Document doc = null;
                string file_name = this.check_bg_ofd.FileName;
                object file = file_name;
                doc = app.Documents.Open(ref file);
                app.Visible = false;
            }
        }

        private void save_bg_sfd_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void gen_btn_Click(object sender, EventArgs e)
        {
            try
            {
                this.progressBar1.Value = 0;
                file_num = 0;

                try
                {
                    //默认名称
                    if (this.textBox1.Text == "")
                    {
                        try
                        {
                            int l = this.ofd.FileName.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries).Length;
                            this.textBox1.Text = this.ofd.FileName.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries)[l - 1];
                            this.textBox1.Text = this.textBox1.Text.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries)[0];
                        }
                        catch
                        {
                            MessageBox.Show("warnning in line 139 ,请选择模板");
                        }
                        //写入的报告是word格式的
                        if (this.word_rbtn.Checked)
                        {
                            //读取 各文件夹中的 数据文件
                            if (this.huizong_rbtn.Checked == true)
                            {
                                this.parse_report1();
                                f2.getpath(this.root + "cache", this.fbd.SelectedPath,this.ofd.FileName);
                            }
                            //读取 word 数据文件
                            else if (this.rbtn_word.Checked == true)
                            {
                                this.parse_report();
                                this.read_tdd();
                            }
                            //读取 cmm 数据文件
                            else if (this.rbtn_cmm.Checked == true)
                            {
                                this.parse_report2();
                                this.read_xmlkd();
                            }
                        }
                        //写入的报告是横版excel格式的
                        else if (this.excel_rbtn_heng.Checked)
                        {
                            this.parse_excel_heng();
                            this.rd_database();
                        }
                        //写入的报告是竖版excel格式的
                        else
                        {
                            this.parse_excel_heng();
                            this.rd_database();
                        }


                        //开始生成报告
                        if (this.huizong_rbtn.Checked == false)
                            this.gen_report();
                    }
                    //手动写入名称
                    else 
                    {
                        if (this.word_rbtn.Checked)
                        {
                            //读取 通道点记录/弯曲扭转记录 数据文件
                            if (this.huizong_rbtn.Checked == true)
                            {
                                this.parse_report1();
                                f2.getpath(this.root + "cache", this.fbd.SelectedPath,this.ofd.FileName);
                            }
                            else if (this.rbtn_word.Checked == true)
                            {
                                this.parse_report();
                                this.read_tdd();
                            }
                            //读取 型面轮廓度(定公差)
                            else if (this.rbtn_cmm.Checked == true)
                            {
                                this.parse_report2();
                                this.read_xmlkd();
                            }
                        }
                        //写入的报告是横版excel格式的
                        else if (this.excel_rbtn_heng.Checked)
                        {
                            this.parse_excel_heng();
                            this.rd_database();

                        }
                        //写入的报告是竖版excel格式的
                        else
                        {
                            this.parse_excel_heng();
                            this.rd_database();
                        }


                        //开始生成报告
                        if (this.huizong_rbtn.Checked == false)
                            this.gen_report();
                    }
                }
                catch
                { }
            }
            catch
            {}
        }

        private void parse_excel_heng()
        {
            //分析模板文件
            this.toolStripStatusLabel1.Text = "正在分析模板...";
            object missing = System.Reflection.Missing.Value;
            MSExcel.Application excelapp = new MSExcel.Application();
            MSExcel.Workbooks wbs = excelapp.Workbooks;
            wbs.Open(this.ofd.FileName, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            MSExcel.Worksheet ws = excelapp.Worksheets.get_Item(1);
            int rows = ws.UsedRange.Rows.Count;
            int cols = ws.UsedRange.Columns.Count;
            
            this.myPoints1.Clear();
            myBuffer temp = new myBuffer();
            this.myPoints1.Add(temp);

            this.num_rows = rows + 1;

            //根据项目创建字典
            string cur_key = "";
            int index = 3;
            for (int i = 3; i <= rows; i++)
            {
                //确保只有在首列非空的情况下赋值
                if (excelapp.Cells[i, 1].Text != "")
                {
                    cur_key = excelapp.Cells[i, 1].Text;
                    Dictionary<string, string> temp_dic = new Dictionary<string, string>();
                    this.myPoints1[0].dic.Add(cur_key, temp_dic);
                }

                if (cur_key == "通道")
                {
                    Regex r = new Regex(@"[a-zA-Z]+");
                    Match m = r.Match(excelapp.Cells[i, 2].Text);
                    this.myPoints1[0].dic[cur_key].Add(m.Value + "_" + excelapp.Cells[i, 3].Text,"");
                }
                else
                {
                    this.myPoints1[0].dic[cur_key].Add(excelapp.Cells[i, 3].Text,"");
                }
            }
            wbs.Close();
            excelapp.Quit();

            //读取word文档，确定dev和超差的横坐标
            int index_of_data = 0;
            if (this.RTF_CHAOCHA.Checked == true)
            {
                index_of_data = 6;
            }
            else
            {
                index_of_data = 5;
            }

            MSWord.Document doc = null;
            MSWord.Application wordApp = new MSWord.Application();
            MSWord.Table table = null;
            string[] c;
            List<string> myFileList1 = new List<string>();
            string my_dir = "";
            DirectoryInfo myFolder;
            try
            {
                myFolder = new DirectoryInfo(this.fbd.SelectedPath);
                my_dir = myFolder.FullName;
                //扫描文件夹，找到第一个.RTF文件
                foreach (FileInfo myNestFile in myFolder.GetFiles())
                {
                    if (Regex.IsMatch(myNestFile.Name, ".RTF"))
                    {
                        myFileList1.Add(myNestFile.FullName);
                        break;
                    }
                }
            }
            catch
            { MessageBox.Show("请先选择模板路径和数据路径！"); }



            //打开文件
            try
            {
                if (File.Exists(myFileList1[0]))
                {
                    object isread = true;
                    object isvisible = false;
                    doc = wordApp.Documents.Open(myFileList1[0], ref missing, ref isread, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                }
            }
            catch
            {
                MessageBox.Show("未能成功打开模板文件！");
                return;
            }

            //读取数据
            try
            {
                for (int i = 1; i < doc.Paragraphs.Count; i++)
                {
                    if (Regex.IsMatch(doc.Paragraphs[i].Range.Text.ToLower(), "ax"))
                    {
                        string[] b = doc.Paragraphs[i].Range.Text.Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        for (int j = 0; j < b.Length; j++)
                        {
                            if (Regex.IsMatch(b[j].ToLower(), "dev"))
                            {
                                this.rtf_dev_loc = j;
                            }

                            if (Regex.IsMatch(b[j].ToLower(), "outtol"))
                            {
                                this.rtf_out_loc = j;
                            }
                        }
                        break;
                    }
                }
            }
            catch
            { }

            doc.Close(ref missing, ref missing, ref missing);
            wordApp.Quit();

        }

        private void rd_database()
        {
            myFileList.Clear();
            //超差和偏差
            int valid_index = 0;
            if (this.chaocha_rbtn.Checked)
            {
                valid_index = 6;
            }
            else
            {
                valid_index = 3;
            }
            
            //以下读取CMM
            object missing = System.Reflection.Missing.Value;
            string a;
            string[] b;
            int id_flag = 0;                                //id_flag的作用是标志此文档的id是否已经读取，1有效
            string current_Section = "";

            this.toolStripStatusLabel1.Text = "";

            DirectoryInfo myFolder;
            string my_dir = "";
            //扫描文件夹，记住各个.CMM文档名字，缓存在 myFileList 列表中

            myFolder = new DirectoryInfo(this.fbd.SelectedPath);

            FileInfo[] FILES = myFolder.GetFiles("*.CMM");

            //文件名排序

            myFileList = FILES.OrderBy(y => y.Name, new FileComparer()).ToList();


            //进度条清零
            this.progressBar1.Value = 0;
            this.progressBar1.Maximum = myFileList.Count*2;

            //依次打开myFileList下各个文件，并读取数据
            this.toolStripStatusLabel1.Text = "正在读取.CMM 数据...";
            int file_index = 0;
            StreamReader sr = null;
            foreach (FileInfo myFile in myFileList)
            {
                id_flag = 0;
                try
                {
                    if (File.Exists(myFile.FullName))
                    {
                        sr = new StreamReader(myFile.FullName);
                    }
                }
                catch
                {
                    MessageBox.Show("未能成功打开数据文件！");
                    this.toolStripStatusLabel1.Text = "";
                    myFileList.Clear();
                    return;
                }

                //列表中增加一个叶片信息
                myBuffer myPoint_temp = new myBuffer();
                Dictionary<string, Dictionary<string, string>> level1 = new Dictionary<string, Dictionary<string, string>>(this.myPoints1[0].dic);
                myPoint_temp.dic = level1;
                foreach (string key in this.myPoints1[0].dic.Keys)
                {
                    Dictionary<string, string> level2 = new Dictionary<string, string>(this.myPoints1[0].dic[key]);
                    myPoint_temp.dic[key] = level2;
                }

                this.myPoints1.Add(myPoint_temp);
                


                bool start = false;
                string line = "";
                //按行遍历文件内容
                while((line = sr.ReadLine())!= null)
                {
                    //判断是否是空行，标准是字符串长度小于2
                    a = line;
                    //自动排除空行
                    if (a.Length > 2)
                    {
                        //应首先查找叶片ID，id_flag保证此部分只执行一次
                        if (id_flag == 0)
                        {
                            b = a.Split(new char[] { '=' }, StringSplitOptions.RemoveEmptyEntries);
                            if (b[0].ToString().Trim() == "SERNO")
                            {
                                //增加对应叶片的名字
                                myPoints1[file_index].id.Add(b[1].ToString().Trim());
                                id_flag = 1;
                            }
                        }
                        //查找各部分的超差值
                        else
                        {
                            b = a.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            //读到section 开始匹配 myPoint1的 section
                            if (b[0].ToString().Trim().ToLower() == "section")
                            {
                                //获取当前section编号6-6
                                current_Section = b[1].ToString().Trim();
                                start = true;
                            }
                            //写入数据
                            else if(start)
                            {
                                try
                                {
                                    if (b.Length == 7)
                                    {
                                        string value = "";
                                        try
                                        {
                                            if (Convert.ToDouble(b[valid_index]) > 0.0001 || Convert.ToDouble(b[valid_index]) < -0.0001)
                                                value = Convert.ToDouble(b[valid_index]).ToString("f3");
                                            else
                                                value = "—";
                                        }
                                        catch
                                        {
                                            value = "—";
                                        }
                                        if (this.myPoints1[file_index].dic[current_Section].ContainsKey(b[0]))
                                            this.myPoints1[file_index].dic[current_Section][b[0]] = value;
                                    }
                                    else if (b.Length == 8)
                                    {
                                        string value = "";
                                        try
                                        {
                                            if (Convert.ToDouble(b[valid_index+1]) > 0.0001 || Convert.ToDouble(b[valid_index+1]) < -0.0001)
                                                value = Convert.ToDouble(b[valid_index+1]).ToString("f3");
                                            else
                                                value = "—";
                                        }
                                        catch
                                        {
                                            value = "—";
                                        }
                                        if (this.myPoints1[file_index].dic[current_Section].ContainsKey(b[0] + " " + b[1]))
                                            this.myPoints1[file_index].dic[current_Section][b[0] + " " + b[1]] = value;
                                    }
                                    else
                                    {
                                        string value = "";
                                        try
                                        {
                                            if (Convert.ToDouble(b[valid_index+2]) > 0.0001 || Convert.ToDouble(b[valid_index+2]) < -0.0001)
                                                value = Convert.ToDouble(b[valid_index+2]).ToString("f3");
                                            else
                                                value = "—";
                                        }
                                        catch
                                        {
                                            value = "—";
                                        }
                                        if (this.myPoints1[file_index].dic[current_Section].ContainsKey(b[0] + " " + b[1] + " " + b[2]))
                                            this.myPoints1[file_index].dic[current_Section][b[0] + " " + b[1] + " " + b[2]] = value;
                                    }
                                }
                                catch { }
                            }
                        }
                    }
                }
                this.progressBar1.Value++;
                result = Convert.ToDouble(this.progressBar1.Value) / Convert.ToDouble(this.progressBar1.Maximum + 1) * 100;
                string d = result.ToString("f1") + "%";
                this.progress_bar_label.Text = d;
                this.progress_bar_label.Refresh();
                file_index++;
            }
            sr.Close();

            //以下读取WORD
            //查看超差、公差的rbtn，根据相应的选中结果赋值
            //如果模板中并未要求收集word信息，则跳过读取word步骤
            bool no_word = false;
            foreach(string key in this.myPoints1[0].dic.Keys)
            {
                if (Regex.IsMatch(key, "通道"))
                {
                    no_word = true;
                }
            }

            if (!no_word)
            {
                return;
            }

            int index_of_data = 0;
            if (this.RTF_CHAOCHA.Checked == true)
            {
                index_of_data = this.rtf_out_loc;
            }
            else
            {
                index_of_data = this.rtf_dev_loc;
            }

            MSWord.Document doc = null;
            MSWord.Application wordApp = new MSWord.Application();
            MSWord.Table table = null;
            string[] c;
            myFileList.Clear();

            //扫描文件夹，记住各个.CMM文档名字，缓存在 myFileList 列表中

            myFolder = new DirectoryInfo(this.fbd.SelectedPath);

            FILES = myFolder.GetFiles("*.RTF");

            //文件名排序

            myFileList = FILES.OrderBy(y => y.Name, new FileComparer()).ToList();


            //依次打开各个数据文件文件读取数据
            this.toolStripStatusLabel1.Text = "正在读取.RTF 数据...";
            foreach (FileInfo myFile in myFileList)
            {
                bool name_is_found = false;
                int blade_id = -1;
                //查找叶片ID是否与文件名对应，未找到就 break;
                for(int i = 0; i < this.myPoints1.Count; i++)
                {
                    if (Regex.IsMatch(myFile.Name, myPoints1[i].id[0]))
                    {
                        blade_id = i;
                        name_is_found = true;
                        break;
                    }
                }

                //如果找到了对应的名字
                bool break_flag = false;
                if (name_is_found)
                {
                    //打开文件
                    try
                    {
                        if (File.Exists(myFile.FullName))
                        {
                            doc = wordApp.Documents.Open(myFile.FullName);
                        }
                    }
                    catch
                    {
                        MessageBox.Show("未能成功打开模板文件！");
                        return;
                    }

                    //读取数据
                    try
                    {
                        int cnt = 0;
                        //从第一行开始遍历
                        for (int i = 1; i < doc.Paragraphs.Count; i++)
                        {
                            //掐头去尾
                            a = doc.Paragraphs[i].Range.Text.Trim();
                            if (a.Length > 2)
                            {
                                //解析文档
                                //如果查到‘DIM’开头的一行，则解析出具体点名字
                                if (a[0].ToString() == "D" && a[1].ToString() == "I" && a[2].ToString() == "M")
                                {
                                    while (true)
                                    {
                                        a = doc.Paragraphs[i].Range.Text.Trim();
                                        //按‘=’分为三段，取第二段
                                        b = a.Split(new char[] { '=' }, StringSplitOptions.RemoveEmptyEntries);
                                        //按‘ ’分为四段，取第四段
                                        c = b[1].Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                        i++;
                                        //循环查找对应的点的名字
                                        foreach(string key1 in this.myPoints1[blade_id].dic.Keys)
                                        {
                                            foreach (string key2 in this.myPoints1[blade_id].dic[key1].Keys)
                                            {

                                                //找到了该叶片对应的point的名字
                                                if (Regex.IsMatch(c[3].ToString().ToLower(), key2.ToLower().Split(new char[] { '_' },StringSplitOptions.RemoveEmptyEntries)[0]))
                                                {
                                                    while (true)
                                                    {
                                                        //以DIM行为开始，找到该point相应的x,y,z,t值
                                                        i++;
                                                        a = doc.Paragraphs[i].Range.Text.Trim();
                                                        try
                                                        {
                                                            if (a[0].ToString() == "D" && a[1].ToString() == "I" && a[2].ToString() == "M")
                                                            {
                                                                i--;
                                                                break_flag = true;
                                                                break;
                                                                
                                                            }
                                                        }
                                                        catch
                                                        { }
                                                        b = a.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                                        try
                                                        {
                                                            if(Regex.IsMatch(key2.Split(new char[] { '_' },StringSplitOptions.RemoveEmptyEntries)[1].ToLower(),b[0].ToString().ToLower()))
                                                            {
                                                                //是负数
                                                                if (Regex.IsMatch( key2.ToLower(),"-"))
                                                                {
                                                                    if ( Convert.ToDouble(b[index_of_data]) > -0.00001 && Convert.ToDouble(b[index_of_data]) < 0.00001)
                                                                        this.myPoints1[blade_id].dic[key1][key2] = "—";
                                                                    else
                                                                        this.myPoints1[blade_id].dic[key1][key2] = (-Convert.ToDouble(b[index_of_data])).ToString("f3");

                                                                }
                                                                //是正数
                                                                else
                                                                {
                                                                    if ( Convert.ToDouble(b[index_of_data]) > -0.00001 && Convert.ToDouble(b[index_of_data]) < 0.00001)
                                                                        this.myPoints1[blade_id].dic[key1][key2] = "—";
                                                                    else
                                                                        this.myPoints1[blade_id].dic[key1][key2] = (Convert.ToDouble(b[index_of_data])).ToString("f3");
                                                                }
                                                            }
                                                        }
                                                        catch
                                                        {}
                                                    }
                                                }
                                                if (break_flag)
                                                {
                                                    break;
                                                }
                                            }
                                            if (break_flag)
                                            {
                                                break_flag = false;
                                                break;
                                            }
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    catch
                    { }
                    doc.Close(ref missing, ref missing, ref missing);

                    this.progressBar1.Value++;
                    result = Convert.ToDouble(this.progressBar1.Value) / Convert.ToDouble(this.progressBar1.Maximum + 1) * 100;
                    string d = result.ToString("f1") + "%";
                    this.progress_bar_label.Text = d;
                    this.progress_bar_label.Refresh();
                }
                else
                {
                    MessageBox.Show("no .CMM file named with" + myFile );
                }
            }

        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            if (this.fbd.ShowDialog() == DialogResult.OK && Directory.Exists(this.fbd.SelectedPath) )
            {
                this.label1.Text = "数据路径：" + this.fbd.SelectedPath;
                this.toolStripStatusLabel1.Text = "已选择数据路径：" + this.fbd.SelectedPath;
            }
            else
            {
                MessageBox.Show("无效的数据路径！");
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int flag = 0;
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        /*
         * 读取word文档类原始数据
         */
        private void read_tdd()
        {
            //查看超差、公差的rbtn，根据相应的选中结果赋值
            int index_of_data = 0;
            if (this.RTF_CHAOCHA.Checked == true)
            {
                index_of_data = this.rtf_out_loc;
            }
            else
            {
                index_of_data = this.rtf_dev_loc;
            }

            MSWord.Document doc = null;
            MSWord.Application wordApp = new MSWord.Application();
            object missing = System.Reflection.Missing.Value;
            MSWord.Table table = null;
            string a;
            string[] b;
            string[] c;
            DirectoryInfo myFolder;
            string my_dir = "";

            myFolder = new DirectoryInfo(this.fbd.SelectedPath);

            FileInfo[] FILES = myFolder.GetFiles("*.RTF");

            //文件名排序
            myFileList = FILES.OrderBy(y => y.Name, new FileComparer()).ToList();


            this.progressBar1.Maximum = myFileList.Count;

            //依次打开各个数据文件文件读取数据
            this.toolStripStatusLabel1.Text = "正在读取.RTF数据...";
            foreach (FileInfo myFile in myFileList)
            {
                try
                {
                    if (File.Exists(myFile.FullName))
                    {
                        object isread = true;
                        object isvisible = false;
                        doc = wordApp.Documents.Open(myFile.FullName, ref missing, ref isread, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                    }
                }
                catch
                {
                    MessageBox.Show("未能成功打开模板文件！");
                    return;
                }

                //列表中增加一个叶片信息
                myBuffer temp_buffur = new myBuffer();

                //记录叶片ID
                try 
                {
                    b = myFile.FullName.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                    c = b[b.Length - 1].Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
                    string[] d1 = null;
                    try
                    {
                        d1 = c[0].Split(new char[] { 't' }, StringSplitOptions.RemoveEmptyEntries);
                        this.myPoints[1].ID.Add(d1[1]);
                    }
                    catch
                    {
                        this.myPoints[1].ID.Add("none");
                    }
                    
                }
                catch 
                {
                    MessageBox.Show("未能在数据文件中找到零件编号！\n 请检查文件“" + myFile + "”格式");
                    continue;
                }

                //记录每个ID的内容
                try
                {
                    int cnt = 0;
                    for (int i = 1; i < doc.Paragraphs.Count; i++)
                    {
                        //读取一行
                        a = doc.Paragraphs[i].Range.Text.Trim();
                        if (a.Length > 2)
                        {
                            //解析文档
                            //如果查到‘DIM’开头的一行，则解析出具体点名字
                            if (a[0].ToString() == "D" && a[1].ToString() == "I" && a[2].ToString() == "M")
                            {
                                while (true)
                                {
                                    a = doc.Paragraphs[i].Range.Text.Trim();
                                    //按‘=’分为三段，取第二段
                                    b = a.Split(new char[] { '=' }, StringSplitOptions.RemoveEmptyEntries);
                                    //按‘ ’分为四段，取第四段
                                    c = b[1].Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                    i++;
                                    //循环查找对应的点的名字
                                    for (int j = 0; j < this.myPoints.Count; j++)
                                    {
                                        try
                                        {
                                            if (c[3].ToString() == this.myPoints[j].Name || c[3].ToString() == ("PNT_" + this.myPoints[j].Name))
                                            {
                                                try
                                                {
                                                    while (true)
                                                    {
                                                        i++;
                                                        a = doc.Paragraphs[i].Range.Text.Trim();
                                                        b = a.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                                        if (b[0].ToString().ToLower() == "ax")
                                                        {
                                                            continue;
                                                        }
                                                        else if (b[0].ToString().ToLower() == "x")
                                                        {

                                                            if (Convert.ToDouble(b[index_of_data].ToString()) < 0.00001 && Convert.ToDouble(b[index_of_data].ToString()) > -0.00001)
                                                            {
                                                                this.myPoints[j].X.Add("—");
                                                                if (this.myPoints[j].X_valid)
                                                                {
                                                                    this.myPoints[j].Is_ok.Add(true);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                
                                                                //如果是取反，则增加负号
                                                                if(this.myPoints[j].Neg)
                                                                    this.myPoints[j].X.Add((-Convert.ToDouble(b[index_of_data])).ToString("f3"));
                                                                else
                                                                    this.myPoints[j].X.Add((Convert.ToDouble(b[index_of_data])).ToString("f3"));

                                                                if (this.myPoints[j].X_valid)
                                                                {
                                                                    this.myPoints[j].Is_ok.Add(false);
                                                                }
                                                            }
                                                            continue;
                                                        }
                                                        else if (b[0].ToString().ToLower() == "y")
                                                        {
                                                            if (Convert.ToDouble(b[index_of_data].ToString()) < 0.00001 && Convert.ToDouble(b[index_of_data].ToString()) > -0.00001)
                                                            {
                                                                this.myPoints[j].Y.Add("—");
                                                                if (this.myPoints[j].Y_valid)
                                                                {
                                                                    this.myPoints[j].Is_ok.Add(true);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (this.myPoints[j].Neg)
                                                                    this.myPoints[j].Y.Add((-Convert.ToDouble(b[index_of_data])).ToString("f3"));
                                                                else
                                                                    this.myPoints[j].Y.Add((Convert.ToDouble(b[index_of_data])).ToString("f3"));
                                                                if (this.myPoints[j].Y_valid)
                                                                {
                                                                    this.myPoints[j].Is_ok.Add(false);
                                                                }
                                                            }
                                                            continue;
                                                        }
                                                        else if (b[0].ToString().ToLower() == "z")
                                                        {
                                                            if (Convert.ToDouble(b[index_of_data].ToString()) < 0.00001 && Convert.ToDouble(b[index_of_data].ToString()) > -0.00001)
                                                            {
                                                                this.myPoints[j].Z.Add("—");
                                                                if (this.myPoints[j].Z_valid)
                                                                {
                                                                    this.myPoints[j].Is_ok.Add(true);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (this.myPoints[j].Neg)
                                                                    this.myPoints[j].Z.Add((-Convert.ToDouble(b[index_of_data])).ToString("f3"));
                                                                else
                                                                    this.myPoints[j].Z.Add((Convert.ToDouble(b[index_of_data])).ToString("f3"));

                                                                if (this.myPoints[j].Z_valid)
                                                                {
                                                                    this.myPoints[j].Is_ok.Add(false);
                                                                }
                                                            }
                                                            continue;
                                                        }
                                                        else if (b[0].ToString().ToLower() == "t")
                                                        {

                                                            if (Convert.ToDouble(b[index_of_data].ToString()) < 0.00001 && Convert.ToDouble(b[index_of_data].ToString()) > -0.00001)
                                                            {
                                                                this.myPoints[j].T.Add("—");
                                                                if (this.myPoints[j].T_valid)
                                                                {
                                                                    this.myPoints[j].Is_ok.Add(true);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (this.myPoints[j].Neg)
                                                                    this.myPoints[j].T.Add((-Convert.ToDouble(b[index_of_data])).ToString("f3"));
                                                                else
                                                                    this.myPoints[j].T.Add((Convert.ToDouble(b[index_of_data])).ToString("f3"));

                                                                if (this.myPoints[j].T_valid)
                                                                {
                                                                    this.myPoints[j].Is_ok.Add(false);
                                                                }
                                                            }
                                                            continue;
                                                        }
                                                        else
                                                        {
                                                            break;
                                                        }
                                                    }
                                                   
                                                    break;
                                                }
                                                catch
                                                {
                                                    break;
                                                }
                                            }
                                        }
                                        catch
                                        { }
                                    }
                                    i--;
                                    break;

                                }
                            }
                            else
                            {
                                //i++;
                            }
                        }
                    }
                }
                catch
                { }
                file_num++;     //已读取的文件数量
                this.progressBar1.Value++;
                result = Convert.ToDouble(this.progressBar1.Value) / Convert.ToDouble(this.progressBar1.Maximum + 1) * 100;
                string d = result.ToString("f1") + "%";
                this.progress_bar_label.Text = d;
                this.progress_bar_label.Refresh();
                doc.Close(ref missing, ref missing, ref missing);
            }

            try
            {
                for (int i = 0; i < this.myPoints[2].Is_ok.Count; i++)
                {
                    this.myPoints[1].Is_ok.Add(true);
                    for (int j = 2; j < this.myPoints.Count; j++)
                    {
                        if (this.myPoints[j].Is_ok[i] == false)
                        {
                            this.myPoints[1].Is_ok[i] = false;
                            break;
                        }
                    }
                }

                //拷贝文件
                //首先创建子文件夹
                myFolder = new DirectoryInfo(this.fbd.SelectedPath);
                string newPath = System.IO.Path.Combine(this.root,"cache" + "\\" + this.textBox1.Text + "合格");
                System.IO.Directory.CreateDirectory(newPath);
                newPath = System.IO.Path.Combine(this.root, "cache" + "\\" + this.textBox1.Text + "不合格");
                System.IO.Directory.CreateDirectory(newPath);

                foreach (FileInfo myNestFile in myFolder.GetFiles())
                {
                    //判断文件名中是否包含合格品名字，是则复制到 合格 文件夹
                    for (int i = 0; i < this.myPoints[2].Is_ok.Count; i++)
                    {
                        if (Regex.IsMatch(myNestFile.FullName.ToString(), this.myPoints[1].ID[i].ToString()) && this.myPoints[1].Is_ok[i])
                        {
                            try
                            {
                                myNestFile.CopyTo(this.root + "cache" + "\\" + this.textBox1.Text + "合格\\" + myNestFile.Name);
                                break;
                            }
                            catch
                            { break; }
                        }

                        if (Regex.IsMatch(myNestFile.FullName.ToString(), this.myPoints[1].ID[i].ToString()) && !this.myPoints[1].Is_ok[i])
                        {
                            try
                            {
                                myNestFile.CopyTo(this.root + "cache" + "\\" + this.textBox1.Text + "不合格\\" + myNestFile.Name);
                                break;
                            }
                            catch
                            {
                                break;
                            }
                        }

                    }
                }
            }
            catch { }
            wordApp.Quit();
        }

        private void SortAsFileName(ref FileInfo[] arrFi)
        {
            Array.Sort(arrFi, delegate(FileInfo x, FileInfo y) { return x.Name.CompareTo(y.Name); });
        }

        //读取 CMM文件
        private void read_xmlkd()
        {
            int valid_index = 0;
            if (this.word_rbtn.Checked)
            {
                MSWord.Document doc = null;
                MSWord.Application wordApp = new MSWord.Application();
                object missing = System.Reflection.Missing.Value;
                string a;
                string[] b;
                int id_flag = 0;                                //id_flag的作用是标志此文档的id是否已经读取，1有效
                string current_Section = "";

                this.toolStripStatusLabel1.Text = "";
                try
                {
                    doc.Close(ref missing, ref missing, ref missing);
                }
                catch
                { }

                DirectoryInfo myFolder;
                myFolder = new DirectoryInfo(this.fbd.SelectedPath);

                FileInfo[] FILES = myFolder.GetFiles("*.CMM");
  
                //文件名排序
                
                myFileList = FILES.OrderBy(y => y.Name, new FileComparer()).ToList();

                //进度条清零
                this.progressBar1.Value = 0;
                this.progressBar1.Maximum = myFileList.Count();

                //依次打开myFileList下各个文件，并读取数据
                this.toolStripStatusLabel1.Text = "正在读取实验数据...";

                foreach (FileInfo myFile in myFileList)
                {
                    id_flag = 0;
                    try
                    {
                        if (File.Exists(myFile.FullName))
                        {
                            object isread = true;
                            object isvisible = false;
                            doc = wordApp.Documents.Open(myFile.FullName, ref missing, ref isread, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                        }
                    }
                    catch
                    {
                        MessageBox.Show("未能成功打开数据文件！");
                        this.toolStripStatusLabel1.Text = "";
                        return;
                    }

                    //按行遍历文件内容
                    int cnt = 0;
                    bool section_matched = false;
                    List<int> index = new List<int>();
                    for (int i = 1; i < doc.Paragraphs.Count; i++)
                    {
                        //判断是否是空行，标准是字符串长度小于2
                        a = doc.Paragraphs[i].Range.Text.Trim();
                        if (a.Length > 2)
                        {
                            //应首先查找叶片ID
                            if (id_flag == 0)
                            {
                                b = a.Split(new char[] { '=' }, StringSplitOptions.RemoveEmptyEntries);
                                if (b[0].ToString().Trim() == "SERNO")
                                {
                                    // myPoints1[0] 是一个列表
                                    // myPoints1[0] 中存放着每个文件中的叶片ID，和叶片是否有效的标志 valid
                                    myPoints1[0].id.Add(b[1].ToString().Trim());
                                    this.myPoints1[0].valid.Add(true);
                                    id_flag = 1;
                                }
                            }
                            //之后查找各部分的超差值
                            else
                            {
                                b = a.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                //由于每行的长度不同，所以使用try，防止访问不存在的位置
                                try
                                {
                                    if (this.gongcha_rbtn.Checked)
                                    {
                                        if (b[0].ToString().ToLower() == "actual")
                                        {
                                            int ii = 0;
                                            foreach (string temp in b)
                                            {

                                                if (temp.Contains("dev"))
                                                {
                                                    valid_index = ii;
                                                }
                                                ii += 1;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        valid_index = 5;
                                    }
                                    //读到section 开始匹配 myPoint1的 section
                                    if (b[0].ToString().Trim() == "Section")
                                    {
                                        index.Clear();
                                        //获取当前section编号6-6
                                        current_Section = b[1].ToString().Trim();
                                        for (int m = 1; m < this.myPoints1.Count; m++)
                                        {
                                            if (this.myPoints1[m].section == current_Section.Split(new char[] { '-' }, StringSplitOptions.RemoveEmptyEntries)[0])
                                            {
                                                //这行数据找到了对应myPoints1的对应的section
                                                section_matched = true;
                                                //记录下对应的myPoints1是哪几个，好几个myPoints1会有相同的section
                                                index.Add(m);
                                                if (this.myPoints1[m].item != "" && this.myPoints1[m].item1 != "")
                                                    cnt = index.Count * 2;
                                                else
                                                    cnt = index.Count;
                                            }
                                        }
                                    }

                                    //既然已经对应上了，那就在接下来的数据中寻找对应section的相应数据
                                    if (section_matched)
                                    {
                                        //遍历上面找到的几个具有相同section 的 myPoints1
                                        for (int m = 0; m < index.Count; m++)
                                        {
                                            //如果找到了对应了参数，就写入该 myPoints1 的value中
                                            if ((b[0]).ToString().Trim() == this.myPoints1[index[m]].item)
                                            {
                                                //如果是没超差，此位置为数字，不会报错；否则就是没有超差
                                                try
                                                {
                                                    this.myPoints1[index[m]].value.Add(Convert.ToDouble(b[valid_index + 1].ToString().Trim()).ToString("f3"));
                                                    //if (Convert.ToDouble(b[valid_index + 1].ToString().Trim()) < 0.0001 && Convert.ToDouble(b[valid_index + 1].ToString().Trim()) > -0.0001)
                                                        this.myPoints1[0].valid[this.myPoints1[0].id.Count - 1] = false;
                                                }
                                                catch
                                                {
                                                    this.myPoints1[index[m]].value.Add("0.000");
                                                }

                                                cnt--;
                                                if (cnt == 0)
                                                {
                                                    section_matched = false;
                                                }
                                            }
                                            else if ((b[0] + " " + b[1]).ToString().Trim() == this.myPoints1[index[m]].item)
                                            {
                                                //如果是没超差，此位置为数字，不会报错；否则就是没有超差
                                                try
                                                {
                                                    this.myPoints1[index[m]].value.Add(Convert.ToDouble(b[valid_index+2].ToString().Trim()).ToString("f3"));
                                                    //if (Convert.ToDouble(b[valid_index + 2].ToString().Trim()) < 0.0001 && Convert.ToDouble(b[valid_index + 2].ToString().Trim()) > -0.0001)
                                                        this.myPoints1[0].valid[this.myPoints1[0].id.Count - 1] = false;
                                                }
                                                catch
                                                {
                                                    this.myPoints1[index[m]].value.Add("0.000");
                                                }

                                                cnt--;
                                                if (cnt == 0)
                                                {
                                                    section_matched = false;
                                                }
                                            }
                                            else if ((b[0] + " " + b[1] + " " + b[2]).ToString().Trim() == this.myPoints1[index[m]].item)
                                            {
                                                //如果是没超差，此位置为数字，不会报错；否则就是没有超差
                                                try
                                                {
                                                    this.myPoints1[index[m]].value.Add(Convert.ToDouble(b[valid_index+3].ToString().Trim()).ToString("f3"));
                                                    //if (Convert.ToDouble(b[valid_index + 3].ToString().Trim()) < 0.0001 && Convert.ToDouble(b[valid_index + 3].ToString().Trim()) > -0.0001)
                                                        this.myPoints1[0].valid[this.myPoints1[0].id.Count - 1] = false;
                                                }
                                                catch
                                                {
                                                    this.myPoints1[index[m]].value.Add("0.000");
                                                }

                                                cnt--;
                                                if (cnt == 0)
                                                {
                                                    section_matched = false;
                                                }
                                            }
                                        }

                                        for (int m = 0; m < index.Count; m++)
                                        {
                                            if (this.myPoints1[index[m]].item1 != "")
                                            {
                                                if ((b[0]).ToString().Trim() == this.myPoints1[index[m]].item1)
                                                {
                                                    //如果是没超差，此位置为数字，不会报错；否则就是没有超差
                                                    try
                                                    {
                                                        this.myPoints1[index[m]].value.Add(Convert.ToDouble(b[6].ToString().Trim()).ToString("f3"));
                                                        //if (Convert.ToDouble(b[6].ToString().Trim()) < 0.0001 && Convert.ToDouble(b[6].ToString().Trim()) > -0.0001)
                                                            this.myPoints1[0].valid[this.myPoints1[0].id.Count - 1] = false;
                                                    }
                                                    catch
                                                    {
                                                        this.myPoints1[index[m]].value.Add("0.000");
                                                    }

                                                    cnt--;
                                                    if (cnt == 0)
                                                    {
                                                        section_matched = false;
                                                    }
                                                }
                                                else if ((b[0] + " " + b[1]).ToString().Trim() == this.myPoints1[index[m]].item1)
                                                {
                                                    //如果是没超差，此位置为数字，不会报错；否则就是没有超差
                                                    try
                                                    {
                                                        this.myPoints1[index[m]].value.Add(Convert.ToDouble(b[7].ToString().Trim()).ToString("f3"));
                                                        //if (Convert.ToDouble(b[7].ToString().Trim()) < 0.0001 && Convert.ToDouble(b[7].ToString().Trim()) > -0.0001)
                                                            this.myPoints1[0].valid[this.myPoints1[0].id.Count - 1] = false;
                                                    }
                                                    catch
                                                    {
                                                        this.myPoints1[index[m]].value.Add("0.000");
                                                    }

                                                    cnt--;
                                                    if (cnt == 0)
                                                    {
                                                        section_matched = false;
                                                    }
                                                }
                                                else if ((b[0] + " " + b[1] + " " + b[2]).ToString().Trim() == this.myPoints1[index[m]].item1)
                                                {
                                                    //如果是没超差，此位置为数字，不会报错；否则就是没有超差
                                                    try
                                                    {
                                                        this.myPoints1[index[m]].value.Add(Convert.ToDouble(b[8].ToString().Trim()).ToString("f3"));
                                                        //if (Convert.ToDouble(b[8].ToString().Trim()) < 0.0001 && Convert.ToDouble(b[8].ToString().Trim()) > -0.0001)
                                                            this.myPoints1[0].valid[this.myPoints1[0].id.Count - 1] = false;
                                                    }
                                                    catch
                                                    {
                                                        this.myPoints1[index[m]].value.Add("0.000");
                                                    }

                                                    cnt--;
                                                    if (cnt == 0)
                                                    {
                                                        section_matched = false;
                                                    }
                                                }
                                            }
                                        }

                                    }
                                }
                                catch
                                {
                                }
                            }
                        }
                    }

                    file_num++;
                    this.progressBar1.Value++;
                    result = Convert.ToDouble(this.progressBar1.Value) / Convert.ToDouble(this.progressBar1.Maximum + 1) * 100;
                    string d = result.ToString("f1") + "%";
                    this.progress_bar_label.Text = d;
                    this.progress_bar_label.Refresh();
                    doc.Close(ref missing, ref missing, ref missing);
                }

                //拷贝文件
                //首先创建子文件夹
                myFolder = new DirectoryInfo(this.fbd.SelectedPath);
                string newPath = System.IO.Path.Combine(this.root + "cache", this.textBox1.Text + "合格");
                System.IO.Directory.CreateDirectory(newPath);
                newPath = System.IO.Path.Combine(this.root + "cache", this.textBox1.Text + "不合格");
                System.IO.Directory.CreateDirectory(newPath);

                foreach (FileInfo myNestFile in myFolder.GetFiles())
                {
                    //判断文件名中是否包含合格品名字，是则复制到 合格 文件夹
                    for (int i = 0; i < this.myPoints1[0].id.Count; i++)
                    {
                        try
                        {
                            if (Regex.IsMatch(myNestFile.FullName.ToString(), this.myPoints1[0].id[i].ToString()) && this.myPoints1[0].valid[i])
                            {
                                try
                                {
                                    myNestFile.CopyTo(this.root + "cache" + this.textBox1.Text + "合格\\" + myNestFile.Name);
                                    break;
                                }
                                catch
                                { break; }
                            }

                            if (Regex.IsMatch(myNestFile.FullName.ToString(), this.myPoints1[0].id[i].ToString()) && !this.myPoints1[0].valid[i])
                            {
                                try
                                {
                                    myNestFile.CopyTo(this.root + "cache" + this.textBox1.Text + "不合格\\" + myNestFile.Name);
                                    break;
                                }
                                catch
                                {
                                    break;
                                }
                            }
                        }
                        catch { }
                    }
                }

                myFileList.Clear();
                wordApp.Quit();
            }
            else if (this.excel_rbtn_heng.Checked)
            {
                object missing = System.Reflection.Missing.Value;
                MSExcel.Application excelapp = new MSExcel.Application();
                MSExcel.Workbooks wbs = excelapp.Workbooks;
                wbs.Open(this.ofd.FileName, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                MSExcel.Worksheet ws = excelapp.Worksheets.get_Item(1);
                int rows = ws.UsedRange.Rows.Count;
                int cols = ws.UsedRange.Columns.Count;
                this.myPoints1.Clear();
                myBuffer temp = new myBuffer();
                this.myPoints1.Add(temp);

                //根据项目创建字典
                string cur_key = "";
                int index = 3;
                for (int i = 3; i <= rows; i++)
                {
                    //确保只有在首列非空的情况下赋值
                    if (excelapp.Cells[i, 1].Text != "")
                    {
                        cur_key = excelapp.Cells[i, 1].Text;
                        Dictionary<string, string> temp_dic = new Dictionary<string, string>();
                        this.myPoints1[0].dic.Add(cur_key, temp_dic);
                    }

                    if (cur_key == "通道")
                    {
                        Regex r = new Regex(@"[a-zA-Z]+");
                        Match m = r.Match(excelapp.Cells[i, 2].Text);
                        this.myPoints1[0].dic[cur_key].Add(m.Value + "-" + excelapp.Cells[i, 3].Text, "");
                    }
                    else
                    {
                        this.myPoints1[0].dic[cur_key].Add(excelapp.Cells[i, 3].Text, "");
                    }
                }
                wbs.Close();
                excelapp.Quit();
            }
        }

        private void parse_report()
        {
            MSWord.Document doc = null;
            MSWord.Application wordApp = new MSWord.Application();
            object missing = System.Reflection.Missing.Value;
            MSWord.Table table = null;
            this.toolStripStatusLabel1.Text = "正在分析报告...";

            //如果报告模板存在
            if (File.Exists(this.ofd.FileName))
            {
                try
                {
                    object isread = true;
                    object isvisible = false;
                    //复制文件到另外一个文件夹
                    FileInfo file = new FileInfo(this.ofd.FileName);
                    file.CopyTo(this.fbd.SelectedPath + "//" + this.textBox1.Text + ".doc", true);
                    doc = wordApp.Documents.Open(this.fbd.SelectedPath + "//" + this.textBox1.Text + ".doc", ref missing, ref isread, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                }
                //System.Reflection.Missing.Value
                catch
                {
                    MessageBox.Show("无法打开模板!");
                }

                int tabel_num = doc.Tables.Count;
                table = doc.Tables[1];

                //遍历表格内部各个Cell的信息
                //确定表格中需要填写内容的起始行数
                for (int i = 1; i <= table.Rows.Count; i++)
                {
                    for (int j = 1; j <= table.Columns.Count; j++)
                    {
                        try
                        {
                            if (table.Cell(i, j).Range.Text.Trim().ToString() == "\r\a" || table.Cell(i, j).Range.Text.Trim().ToString() == "\r" || table.Cell(i, j).Range.Text.Trim().ToString() == "\a")
                            {
                                //找到起始行
                                if (table.Cell(i, j + 1).Range.Text.Trim().ToString() == "\r\a" || table.Cell(i, j + 1).Range.Text.Trim().ToString() == "\r" || table.Cell(i, j + 1).Range.Text.Trim().ToString() == "\a")
                                {
                                    data_line = i-1;
                                    break;
                                }
                            }
                        }
                        catch { }
                    }
                }	

                myPoint temp0 = new myPoint();
                temp0.Col = 1;
                temp0.Name = "序号";
                this.myPoints.Add(temp0);

                myPoint temp1 = new myPoint();
                temp1.Col = 1;
                temp1.Name = "叶片编号";
                this.myPoints.Add(temp1);


                //在起始行中分析需要填入哪些东西
                string[] a;
                for (int j = 1; j <= table.Columns.Count; j++)
                {
                    try
                    {
                        if (j >= 3)
                        {
                            a = table.Cell(data_line, j).Range.Text.ToString().Split(new string[] { "\r\a" }, StringSplitOptions.RemoveEmptyEntries);
                            myPoint temp = new myPoint();
                            temp.Col = j;
                            temp.Name = a[0];
                            this.myPoints.Add(temp);
                        }
                    }
                    catch { }
                }

                //在起始行的下一行中分析x,y,z,t哪个有效
                for (int j = 1; j <= table.Columns.Count; j++)
                {
                    try
                    {
                        if (table.Cell(data_line + 1, j).Range.Text.ToString().ToLower() == "x\r\a" || table.Cell(data_line + 1, j).Range.Text.ToString().ToLower() == "-x\r\a")
                        {
                            this.myPoints[j-1].X_valid = true;
                            if (table.Cell(data_line + 1, j).Range.Text.ToString().ToLower() == "-x\r\a")
                            {
                                this.myPoints[j - 1].Neg = true;
                            }
                            table.Cell(data_line + 1, j).Range.Text = "\r\a";
                            continue;
                        }
                        else if (table.Cell(data_line + 1, j).Range.Text.ToString().ToLower() == "y\r\a" || table.Cell(data_line + 1, j).Range.Text.ToString().ToLower() == "-y\r\a")
                        {
                            this.myPoints[j-1].Y_valid = true;
                            if (table.Cell(data_line + 1, j).Range.Text.ToString().ToLower() == "-y\r\a")
                            {
                                this.myPoints[j - 1].Neg = true;
                            }
                            table.Cell(data_line + 1, j).Range.Text = "\r\a";
                            continue;
                        }
                        else if (table.Cell(data_line + 1, j).Range.Text.ToString().ToLower() == "z\r\a" || table.Cell(data_line + 1, j).Range.Text.ToString().ToLower() == "-z\r\a")
                        {
                            this.myPoints[j-1].Z_valid = true;
                            if (table.Cell(data_line + 1, j).Range.Text.ToString().ToLower() == "-z\r\a")
                            {
                                this.myPoints[j - 1].Neg = true;
                            }
                            table.Cell(data_line + 1, j).Range.Text = "\r\a";
                            continue;
                        }
                        else if (table.Cell(data_line + 1, j).Range.Text.ToString().ToLower() == "t\r\a" || table.Cell(data_line + 1, j).Range.Text.ToString().ToLower() == "-t\r\a")
                        {
                            this.myPoints[j-1].T_valid = true;
                            if (table.Cell(data_line + 1, j).Range.Text.ToString().ToLower() == "-t\r\a")
                            {
                                this.myPoints[j - 1].Neg = true;
                            }
                            table.Cell(data_line + 1, j).Range.Text = "\r\a";
                            continue;
                        }
                    }
                    catch { }
                }
            }
            else
            {
                MessageBox.Show("模板不存在!");
            }

            //读取word文档，确定dev和超差的横坐标
            List<string> myFileList1 = new List<string>();
            string my_dir = "";
            DirectoryInfo myFolder;
            try
            {
                myFolder = new DirectoryInfo(this.fbd.SelectedPath);
                my_dir = myFolder.FullName;
                //扫描文件夹，找到第一个.RTF文件
                foreach (FileInfo myNestFile in myFolder.GetFiles())
                {
                    if (Regex.IsMatch(myNestFile.Name, ".RTF"))
                    {
                        myFileList1.Add(myNestFile.FullName);
                        break;
                    }
                }
            }
            catch
            { MessageBox.Show("请先选择模板路径和数据路径！"); }



            //打开文件
            try
            {
                if (File.Exists(myFileList1[0]))
                {
                    doc = wordApp.Documents.Open(myFileList1[0]);
                }
            }
            catch
            {
                MessageBox.Show("未能成功打开模板文件！");
                return;
            }

            //读取数据
            try
            {
                for (int i = 1; i < doc.Paragraphs.Count; i++)
                {
                    if (Regex.IsMatch(doc.Paragraphs[i].Range.Text.ToLower(), "ax"))
                    {
                        string[] b = doc.Paragraphs[i].Range.Text.Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        for (int j = 0; j < b.Length; j++)
                        {
                            if (Regex.IsMatch(b[j].ToLower(), "dev"))
                            {
                                this.rtf_dev_loc = j;
                            }

                            if (Regex.IsMatch(b[j].ToLower(), "outtol"))
                            {
                                this.rtf_out_loc = j;
                            }
                        }
                        break;
                    }

                }
            }
            catch
            { }

            doc.Close(ref missing, ref missing, ref missing);
            string bb = "";
            for (int i = 0; i < this.myPoints.Count; i++)
            {
                bb += this.myPoints[i].Name + " ";
            }
                
            this.toolStripStatusLabel1.Text = "分析完成，需要检测:" + bb;
            wordApp.Quit();
        }

        private void parse_report2()
        {
            MSWord.Document doc = null;
            MSWord.Application wordApp = new MSWord.Application();
            //wordApp.Visible = true;
            object missing = System.Reflection.Missing.Value;
            MSWord.Table table = null;

            int start_line = 0;

            try
            {
                object isread = true;
                object isvisible = false;
                //复制文件到另外一个文件夹
                FileInfo file = new FileInfo(this.ofd.FileName);
                file.CopyTo(this.fbd.SelectedPath + "//" + this.textBox1.Text + ".doc", true);
                doc = wordApp.Documents.Open(this.fbd.SelectedPath + "//" + this.textBox1.Text + ".doc", ref missing, ref isread, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            }
            catch
            {
                MessageBox.Show("无法解析模板!");
            }

            int tabel_num = doc.Tables.Count;
            table = doc.Tables[1];

            //遍历 确定写入文件开始位置
            for (int i = this.data_line + 1; i <= table.Rows.Count; i++)
            {
                for (int j = 1; j <= table.Columns.Count; j++)
                {
                    try
                    {
                        if (table.Cell(i, j).Range.Text.ToString() == "\r\a" || table.Cell(i, j).Range.Text.ToString() == "\a" || table.Cell(i, j).Range.Text.ToString() == "\r")
                        {
                            if(table.Cell(i, j + 1).Range.Text.ToString() == "\r\a" || table.Cell(i, j + 1).Range.Text.ToString() == "\a" || table.Cell(i, j + 1).Range.Text.ToString() == "\r")
                            {
                                start_line = i;
                                break;
                            }
                        }
                    }
                    catch
                    { }
                }
            }

            this.col_num = 1;

            //确定要写入的内容
            myBuffer temp11 = new myBuffer();
            this.myPoints1.Add(temp11);

            for (int j = 3; j <= table.Columns.Count; j++)
            {
                try
                {
                    if (table.Cell(start_line, j).Range.Text.ToString() != "\r\a")
                    {
                        myBuffer temp = new myBuffer();
                        if (table.Cell(start_line, j).Range.Text.ToString().Split('\r').Length == 3)
                        {
                            temp.section = table.Cell(start_line, j).Range.Text.ToString().Split('\r')[0];
                            temp.item = table.Cell(start_line, j).Range.Text.ToString().Split('\r')[1];
                        }
                        else if (table.Cell(start_line, j).Range.Text.ToString().Split('\r').Length == 4)
                        {
                            temp.section = table.Cell(start_line, j).Range.Text.ToString().Split('\r')[0];
                            temp.item = table.Cell(start_line, j).Range.Text.ToString().Split('\r')[1];
                            temp.item1 = table.Cell(start_line, j).Range.Text.ToString().Split('\r')[2];
                        }
                        this.myPoints1.Add(temp);
                        this.col_num++;
                    }
                }
                catch
                { }
            }
            doc.Close(ref missing, ref missing, ref missing);
            wordApp.Quit();
        }

        private void parse_report1()
        {
            
            DirectoryInfo myFolder;
            try
            {
                f2.checkedListBox1.Items.Clear();
                myFolder = new DirectoryInfo(this.root + "cache");
                //扫描文件夹
                foreach (DirectoryInfo myNestFile in myFolder.GetDirectories())
                {
                    string temp = Regex.Replace(myNestFile.Name, "[\u4300-\u9fa5]", "", RegexOptions.IgnoreCase);

                    if (Regex.IsMatch(myNestFile.Name, "不合格") && this.ofd.FileName.Contains(temp.Trim()))
                    {
                        string[] a = myNestFile.Name.Split(new string[] { "不合格" }, StringSplitOptions.RemoveEmptyEntries);
                        f2.checkedListBox1.Items.Add(a[0]);
                    }
                }
            }
            catch
            {
                MessageBox.Show("请先生成其他报告！");
            }

            f2.Visible = false;
            f2.Show();
        }

        private void gen_report()
        {
            //选择了word文档按钮
            if (this.word_rbtn.Checked)
            {
                MSWord.Document doc = null;
                MSWord.Application wordApp = new MSWord.Application();
                wordApp.Visible = false;
                object missing = System.Reflection.Missing.Value;
                MSWord.Table table = null;
                this.toolStripStatusLabel1.Text = "正在生成报告...";
                if (File.Exists(this.ofd.FileName))
                {
                    try
                    {
                        doc = wordApp.Documents.Open(this.fbd.SelectedPath + "//" + this.textBox1.Text +".doc");
                    }
                    //System.Reflection.Missing.Value
                    catch
                    { return; }

                    int tabel_num = doc.Tables.Count;
                    table = doc.Tables[1];
                    int start_line = 0;
                    int current_line = 0;
                    int myIndex = 0;
                    int curIndex = 0;
                    //数据列数
                    //wordApp.Visible = true;

                    List<int> index_col = new List<int>();              //序号栏起始列坐标
                    List<int> ID_col = new List<int>();                 //叶片编号栏偏移列坐标
                    List<int> jiemian_start_col = new List<int>();      //截面栏偏移列坐标
                    List<int> qy_col = new List<int>();                 //前缘栏偏移列坐标
                    List<int> wy_col = new List<int>();                 //尾缘栏偏移列坐标
                    List<int> yp_col = new List<int>();                 //叶盆栏偏移列坐标
                    List<int> yb_col = new List<int>();                 //叶背栏偏移列坐标

                    //计算报告页数，提供换页
                    MSWord.WdStatistic stat = MSWord.WdStatistic.wdStatisticPages;
                    int pages = doc.ComputeStatistics(stat, Missing.Value);
                    bool flag = false;
                    //1 生成 .RTF 报告
                    if (this.rbtn_word.Checked)
                    {
                        int maxindex = Math.Max(Math.Max(Math.Max(this.myPoints[3].X.Count, this.myPoints[3].Y.Count), this.myPoints[3].Z.Count), this.myPoints[3].T.Count);
                        //遍历 通道记录点 模板文件中的所有表格
                        for (int tabel_cnt = 1; tabel_cnt <= doc.Tables.Count; tabel_cnt++)
                        {
                            //定位i
                            int i = this.data_line + 1;

                            //循环写入表格
                            while (curIndex < maxindex)
                            {
                                bool temp = false;
                                //只写入不合格的数据,temp是是否合格的标志位
                                for (int iii = 0; iii < this.myPoints.Count; iii++)
                                {
                                    try
                                    {
                                        if (!this.myPoints[iii].Is_ok[curIndex])
                                        {
                                            temp = true;
                                            break;
                                        }
                                    }
                                    catch
                                    {

                                    }
                                }


                                if (temp)
                                {
                                    temp = false;
                                    if (table.Rows.Count < 3)
                                    {
                                        this.data_line = 1;
                                        i = this.data_line + 1;
                                        doc.Content.Tables[tabel_cnt].Rows.Add(table.Cell(i, 1));
                                    }
                                    else
                                    {
                                        doc.Content.Tables[tabel_cnt].Rows.Add(table.Cell(i, 1));
                                    }


                                    
                                    try
                                    {
                                        table.Cell(i, 1).Range.Text = (myIndex + 1).ToString();
                                    }
                                    catch
                                    {
                                        MessageBox.Show("写入第" + (myIndex + 1).ToString() + "个叶片序号失败！");
                                        find_word_error(doc, wordApp, missing);
                                        return;
                                    }
                                    try
                                    {
                                        table.Cell(i, 2).Range.Text = this.myPoints[1].ID[curIndex].ToString();
                                    }
                                    catch
                                    {
                                        MessageBox.Show("写入第" + (myIndex + 1).ToString() + "个叶片ID失败！");
                                        find_word_error(doc, wordApp, missing);
                                        return;
                                    }

                                    try
                                    {
                                        int j = 3;
                                        for (int n = 0; n < this.myPoints.Count; n++)
                                        {
                                            try
                                            {
                                                if (j == this.myPoints[n].Col)
                                                {
                                                    if (this.myPoints[n].T_valid)
                                                    {
                                                        table.Cell(i, j).Range.Text = this.myPoints[n].T[curIndex].ToString();
                                                    }
                                                    else if (this.myPoints[n].X_valid)
                                                    {
                                                        double x = 0;
                                                        try
                                                        {
                                                            x = Convert.ToDouble(this.myPoints[n].X[curIndex]);
                                                        }
                                                        catch
                                                        {
                                                            x = 0.0;
                                                        }

                                                        double t = 0;
                                                        try
                                                        {
                                                            t = Convert.ToDouble(this.myPoints[n].T[curIndex]);
                                                        }
                                                        catch
                                                        {
                                                            t = 0.0;
                                                        }

                                                        if (this.ref_rbtn.Checked && x * t <= 0)
                                                        {
                                                            x = -x;
                                                            if (x > 0.0001 || x < -0.0001)
                                                                table.Cell(i, j).Range.Text = x.ToString("f3");
                                                            else
                                                                table.Cell(i, j).Range.Text = "—";
                                                        }
                                                        else
                                                        {
                                                            table.Cell(i, j).Range.Text = this.myPoints[n].X[curIndex].ToString();
                                                        }
                                                    }
                                                    else if (this.myPoints[n].Y_valid)
                                                    {
                                                        double y = 0;
                                                        try
                                                        {
                                                            y = Convert.ToDouble(this.myPoints[n].Y[curIndex]);
                                                        }
                                                        catch
                                                        {
                                                            y = 0.0;
                                                        }

                                                        double t = 0;
                                                        try
                                                        {
                                                            t = Convert.ToDouble(this.myPoints[n].T[curIndex]);
                                                        }
                                                        catch
                                                        {
                                                            t = 0.0;
                                                        }

                                                        if (this.ref_rbtn.Checked && y * t <= 0)
                                                        {
                                                            y = -y;
                                                            if (y > 0.0001 || y < -0.0001)
                                                                table.Cell(i, j).Range.Text = y.ToString("f3");
                                                            else
                                                                table.Cell(i, j).Range.Text = "—";
                                                        }
                                                        else
                                                        {
                                                            table.Cell(i, j).Range.Text = this.myPoints[n].Y[curIndex].ToString();
                                                        }
                                                    }
                                                    else if (this.myPoints[n].Z_valid)
                                                    {
                                                        double z = 0;
                                                        try
                                                        {
                                                            z = Convert.ToDouble(this.myPoints[n].Z[curIndex]);
                                                        }
                                                        catch
                                                        {
                                                            z = 0.0;
                                                        }

                                                        double t = 0;
                                                        try
                                                        {
                                                            t = Convert.ToDouble(this.myPoints[n].T[curIndex]);
                                                        }
                                                        catch
                                                        {
                                                            t = 0.0;
                                                        }

                                                        if (this.ref_rbtn.Checked && z * t <= 0)
                                                        {
                                                            z = -z;
                                                            if (z > 0.0001 || z < -0.0001)
                                                                table.Cell(i, j).Range.Text = z.ToString("f3");
                                                            else
                                                                table.Cell(i, j).Range.Text = "—";
                                                        }
                                                        else
                                                        {
                                                            table.Cell(i, j).Range.Text = this.myPoints[n].Z[curIndex].ToString();
                                                        }
                                                    }
                                                    j++;
                                                }
                                            }
                                            catch
                                            { }
                                        }
                                    }
                                    catch
                                    {
                                        MessageBox.Show("写入第" + (myIndex + 1).ToString() + "个叶片数据失败！");
                                        find_word_error(doc, wordApp, missing);
                                        return;                                    
                                    }

                                    i++;
                                    myIndex++;
                                }
                                curIndex++;
                            }

                            //最后一行清空
                            try
                            {
                                for (int ii = 1; ii <= this.myPoints.Count; ii++)
                                {
                                    table.Cell(i, ii).Range.Text = "";
                                }
                            }
                            catch
                            {
                                //MessageBox.Show("删除标记时遇到错误！");
                                //find_word_error(doc, wordApp, missing);
                                //return;
                            }
                        }
                    }

                    //2  .CMM 报告
                    else if (this.rbtn_cmm.Checked)
                    {
                        //遍历 确定写入文件开始位置
                        for (int i = 1; i <= table.Rows.Count; i++)
                        {
                            for (int j = 1; j <= table.Columns.Count; j++)
                            {
                                try
                                {
                                    if (table.Cell(i, j).Range.Text.ToString() == "\r\a" && table.Cell(i, j + 1).Range.Text.ToString() == "\r\a")
                                    {
                                        start_line = i;
                                        break;
                                    }
                                }
                                catch
                                { }
                            }
                        }

                        int maxindex = this.myPoints1[0].id.Count;
                        myIndex = 0;
                        int index = 0;
                        int iiii = 0;
                        //遍历,填入信息
                        try
                        {
                            for (int i = start_line; i <= table.Rows.Count; )
                            {
                                if (myIndex < maxindex)
                                {
                                    if (!this.myPoints1[0].valid[myIndex])
                                    {
                                        //增加一行
                                        try
                                        {
                                            doc.Content.Tables[1].Rows.Add(table.Cell(i, 1));
                                        }
                                        catch
                                        {
                                            MessageBox.Show("已经写入第" + myIndex.ToString() + "个叶片的数据，但是在给word表格增加下一行时发生错误！");
                                            find_word_error(doc, wordApp, missing);
                                            return;
                                        }

                                        index++;
                                        for (int j = 1; j <= this.col_num; )
                                        {
                                            if (j == 1)
                                            {
                                                try
                                                {
                                                    table.Cell(i, j).Range.Text = (index).ToString();
                                                    j++;
                                                }
                                                catch
                                                {
                                                    MessageBox.Show("写入第" + myIndex.ToString() +  "个叶片的序号失败！");
                                                    find_word_error(doc, wordApp, missing);
                                                    return;
                                                }
                                            }
                                            else if (j == 2)
                                            {
                                                try
                                                {
                                                    table.Cell(i, j).Range.Text = this.myPoints1[0].id[myIndex];
                                                    j++;
                                                }
                                                catch
                                                {
                                                    MessageBox.Show("写入第" + myIndex.ToString() + "个叶片的ID失败！");
                                                    find_word_error(doc, wordApp, missing);
                                                    return;
                                                }
                                            }
                                            else
                                            {
                                                //循环写入各个点
                                                for (int m = 1; m < this.myPoints1.Count; m++)
                                                {
                                                    //有两排数据的情况
                                                    if (this.myPoints1[m].item1 != "")
                                                    {
                                                        //写第一排
                                                        try
                                                        {
                                                            if (Convert.ToDouble(this.myPoints1[m].value[2 * myIndex]) > -0.0001 && Convert.ToDouble(this.myPoints1[m].value[2 * myIndex]) < 0.0001)
                                                            {
                                                                //如果第一行是‘-’,第二行不是‘-’,则第不写
                                                                if (Convert.ToDouble(this.myPoints1[m].value[2 * myIndex + 1]) > -0.0001 && Convert.ToDouble(this.myPoints1[m].value[2 * myIndex + 1]) < 0.0001)
                                                                    table.Cell(i, j).Range.Text = "—";
                                                            }
                                                            else
                                                            {
                                                                table.Cell(i, j).Range.Text = this.myPoints1[m].value[2 * myIndex];
                                                            }
                                                        }
                                                        catch
                                                        {
                                                            MessageBox.Show("写入第" + myIndex.ToString() + "个叶片第1排数据错误！");
                                                            find_word_error(doc, wordApp, missing);
                                                            return;
                                                        }

                                                        //写第二排
                                                        try
                                                        {
                                                            if (Convert.ToDouble(this.myPoints1[m].value[2 * myIndex + 1]) > -0.0001 && Convert.ToDouble(this.myPoints1[m].value[2 * myIndex + 1]) < 0.0001)
                                                            {
                                                                ////如果第一行是‘-’,第二行也是‘-’,则第不写
                                                                //if (!(Convert.ToDouble(this.myPoints1[m].value[2 * myIndex]) > -0.0001 && Convert.ToDouble(this.myPoints1[m].value[2 * myIndex]) < 0.0001))
                                                                //    table.Cell(i, j).Range.InsertAfter("\r—");
                                                            }
                                                            else
                                                            {
                                                                if (Convert.ToDouble(this.myPoints1[m].value[2 * myIndex]) > -0.0001 && Convert.ToDouble(this.myPoints1[m].value[2 * myIndex]) < 0.0001)
                                                                    table.Cell(i, j).Range.Text =  this.myPoints1[m].value[2 * myIndex + 1];
                                                                else
                                                                    table.Cell(i, j).Range.InsertAfter("\r" + this.myPoints1[m].value[2 * myIndex + 1]);
                                                            }
                                                        }
                                                        catch
                                                        {
                                                            MessageBox.Show("写入第" + myIndex.ToString() + "个叶片第2排数据错误！");
                                                            find_word_error(doc, wordApp, missing);
                                                            return;
                                                        }
                                                    }
                                                    //有一排数据的情况
                                                    else
                                                    {
                                                        try
                                                        {
                                                            if (Convert.ToDouble(this.myPoints1[m].value[myIndex]) > -0.0001 && Convert.ToDouble(this.myPoints1[m].value[myIndex]) < 0.0001)
                                                            {
                                                                table.Cell(i, j).Range.Text = "—";
                                                            }
                                                            else
                                                            {
                                                                table.Cell(i, j).Range.Text = this.myPoints1[m].value[myIndex];
                                                            }
                                                        }
                                                        catch
                                                        {
                                                            MessageBox.Show("写入第" + myIndex.ToString() + "个叶片数据错误！");
                                                            find_word_error(doc, wordApp, missing);
                                                            return;
                                                        }
                                                    }

                                                    j++;
                                                }

                                            }
                                        }
                                        i++;
                                        iiii = i;
                                    }
                                    myIndex++;
                                }
                                else
                                { break; }
                            }

                            //重写最后一行，删除模板标记
                            for (int j = table.Columns.Count; j >= 1; j--)
                            {
                                try
                                {
                                    table.Cell(iiii, j).Range.Text = "";
                                }
                                catch
                                {
                                    //MessageBox.Show("删除标记时遇到错误！");
                                    //find_word_error(doc, wordApp, missing);
                                    //return;
                                }
                            }
                        }
                        catch
                        {
                            MessageBox.Show("error in line 1449, when writing data into target doc");
                        }

                    }
                    else if (this.huizong_rbtn.Checked)
                    {
                        doc.Close(ref missing, ref missing, ref missing);
                        wordApp.Quit();
                    }



                    string d = "100%";
                    this.progress_bar_label.Text = d;
                    this.progress_bar_label.Refresh();

                    //文档编辑完以后必须要保存，否则更改无效！                  
                    doc.Save();
                    doc.Close(ref missing, ref missing, ref missing);
                    this.toolStripStatusLabel1.Text = "共" + file_num.ToString() + "个数据文件" + "，成功生成报告！";

                    //清理缓存
                    myPoints1.Clear();
                    myPoints.Clear();
                    myFileList.Clear();

                }
                else
                {
                    MessageBox.Show("请先选择模板！");
                }
                wordApp.Quit();
            }
            else if (this.excel_rbtn_heng.Checked || this.excel_rbtn_shu.Checked)
            {
                int col_length = 0;
                if (this.excel_rbtn_heng.Checked)
                {
                    col_length = 9;
                }
                else
                {
                    col_length = 5;
                }
                object missing = System.Reflection.Missing.Value;
                MSExcel.Application excelapp = new MSExcel.Application();
                MSExcel.Workbooks wbs = excelapp.Workbooks;
                wbs.Open(this.ofd.FileName, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                MSExcel.Worksheet ws = excelapp.Worksheets.get_Item(1);

                excelapp.Visible = false;

                int row = 3;
                int col = 3;
                int save_name = 1;
                //根据字典写入数据
                int cnt = 0;
                this.toolStripStatusLabel1.Text = "正在写入数据，请等待...";
                try
                {
                    foreach (myBuffer blade in myPoints1)
                    {
                        cnt++;
                        if (col < col_length)
                        {
                            excelapp.Cells[2, col] = blade.id[0];
                            foreach (string key1 in blade.dic.Keys)
                            {
                                foreach (string key2 in blade.dic[key1].Keys)
                                {
                                    excelapp.Cells[row, col] = blade.dic[key1][key2];
                                    row++;
                                }
                            }
                            col++;
                            row = 3;
                            //如果写到一半，需要把其他的覆盖掉
                            if (cnt == myPoints1.Count - 1)
                            {
                                row = 2;
                                for (int col1 = col; col1 <= col_length; col1++)
                                {
                                    for (row = 2; row < this.num_rows; row++)
                                    {
                                        excelapp.Cells[row, col1] = "";
                                    }
                                }
                                ws.SaveAs(this.fbd.SelectedPath + "\\" + save_name.ToString());
                                break;
                            }
                        }
                        else
                        {
                            excelapp.Cells[2, col] = blade.id[0];
                            foreach (string key1 in blade.dic.Keys)
                            {
                                foreach (string key2 in blade.dic[key1].Keys)
                                {
                                    excelapp.Cells[row, col] = blade.dic[key1][key2];
                                    row++;
                                }
                            }

                            ws.SaveAs(this.fbd.SelectedPath + "\\" + save_name.ToString());
                            col = 3;
                            row = 3;
                            save_name++;
                        }


                    }

                    myPoints1.Clear();
                }
                catch
                {
                    myPoints1.Clear();
                }



                col = 0;
                row = 0;

                string d = "100%";
                this.progress_bar_label.Text = d;
                this.progressBar1.Value = this.progressBar1.Maximum;
                this.progressBar1.Refresh();
                this.progress_bar_label.Refresh();
                string path = this.fbd.SelectedPath + "\\" + save_name.ToString();
                ws.SaveAs(@path);
                wbs.Close();
                excelapp.Quit();
                this.toolStripStatusLabel1.Text = "成功生成excel报告！";
            }
        }

        private void 帮助ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        //如果发现错误，清缓存，关word
        private void find_word_error(MSWord.Document doc, MSWord.Application wordApp, object missing)
        {
            doc.Close(ref missing, ref missing, ref missing);
            //清理缓存
            myPoints1.Clear();
            myPoints.Clear();
            myFileList.Clear();
            wordApp.Quit();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if(f2.message == true)
            {
                string d = "100%";
                this.progress_bar_label.Text = d;
                this.progress_bar_label.Refresh();
                this.progressBar1.Value = this.progressBar1.Maximum;

                this.toolStripStatusLabel1.Text = "成功生成汇总报告！";
            }
        }

        private void rbtn_cmm_CheckedChanged(object sender, EventArgs e)
        {
            if(this.word_rbtn.Checked)
                if (this.rbtn_cmm.Checked)
                {
                    this.groupBox5.Enabled = true;
                    this.groupBox6.Enabled = false;
                }
                else
                {
                    this.groupBox5.Enabled = false;
                    this.groupBox6.Enabled = true;
                }
            else
                this.groupBox6.Enabled = true;
        }

        private void rbtn_word_CheckedChanged(object sender, EventArgs e)
        {
            if (this.word_rbtn.Checked)
                if (this.rbtn_word.Checked)
                    this.groupBox5.Enabled = false;
                else
                    this.groupBox5.Enabled = true;
            else
                this.groupBox5.Enabled = true;
        }

        private void rbtn_word_Click(object sender, EventArgs e)
        {
            this.rbtn_cmm.Checked = false;
            this.rbtn_word.Checked = true;
            this.huizong_rbtn.Checked = false;
        }

        private void rbtn_cmm_Click(object sender, EventArgs e)
        {
            this.rbtn_word.Checked = false;
            this.rbtn_cmm.Checked = true;
            this.huizong_rbtn.Checked = false;
        }

        private void huizong_rbtn_Click(object sender, EventArgs e)
        {
            this.rbtn_cmm.Checked = false;
            this.rbtn_word.Checked = false;
            this.huizong_rbtn.Checked = true;
        }

        private void word_rbtn_CheckedChanged(object sender, EventArgs e)
        {
            if (this.word_rbtn.Checked)
            {
                this.groupBox4.Enabled = true;
                if (this.rbtn_word.Checked)
                {
                    this.groupBox5.Enabled = false;
                    this.groupBox6.Enabled = true;
                }
                else if (this.huizong_rbtn.Checked)
                {
                    this.groupBox5.Enabled = false;
                    this.groupBox6.Enabled = false;
                }
                else
                {
                    this.groupBox6.Enabled = false;
                    this.groupBox5.Enabled = true;
                }

            }
            else
            {
                this.groupBox4.Enabled = false;
                this.groupBox5.Enabled = true;
                this.groupBox6.Enabled = true;
            }
        }

        private void excel_rbtn_heng_CheckedChanged(object sender, EventArgs e)
        {
            if (this.word_rbtn.Checked)
            {
                this.groupBox4.Enabled = true;
            }
            else
            {
                this.groupBox5.Enabled = true;
                this.groupBox6.Enabled = true;
                this.groupBox4.Enabled = false;
            }
        }

        private void excel_rbtn_shu_CheckedChanged(object sender, EventArgs e)
        {
            if (this.word_rbtn.Checked)
            {
                this.groupBox4.Enabled = true;
            }
            else
            {
                this.groupBox5.Enabled = true;
                this.groupBox6.Enabled = true;
                this.groupBox4.Enabled = false;
            }
        }

        private void huizong_rbtn_CheckedChanged(object sender, EventArgs e)
        {
            if (this.huizong_rbtn.Checked)
            {
                this.groupBox5.Enabled = false;
                this.groupBox6.Enabled = false;
            }
        }

    }
}

