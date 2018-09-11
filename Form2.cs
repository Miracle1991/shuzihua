using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using MSWord = Microsoft.Office.Interop.Word;

namespace shuzihua
{
    
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            this.clicked = false;
            this.item = new List<string>();
        }
        public bool done = false;
        public bool clicked ;
        public List<string> item = new List<string>();
        public string f2_muban_path = "";
        public string f2_data_path = "";
        public string f2_save_path = "";
        List<string> buhege = new List<string>();
        List<string> hege = new List<string>();
        List<string> hege1 = new List<string>();
        public bool f2_ok = false;
        public bool message = false;
        private List<string> myitems = new List<string>();
        private void button1_Click(object sender, EventArgs e)
        {

            this.Visible = false;
            for (int i = 0; i < this.checkedListBox1.Items.Count; i++)
            {
                try
                {
                    if (this.checkedListBox1.GetItemChecked(i))
                        myitems.Add(this.checkedListBox1.GetItemText(this.checkedListBox1.Items[i]));
                }
                catch
                {
                    MessageBox.Show("error in line 47, file Form2.cs");
                }
            }

            this.buhege = this.gen_huizong();
            this.message = this.gen_report();
            this.f2_ok = true;
        }

        //f2.getpath(this.root + "cache", this.fbd.SelectedPath, this.ofd.FileName);
        public void getpath(string path0, string path1,string path2)
        {
            this.f2_muban_path = path2;
            this.f2_data_path = path0;
            this.f2_save_path = path1;
        }

        private List<string> gen_huizong()
        {
            
            DirectoryInfo myFolder;
            int m = 0;
            int n = 0;

            try
            {
                myFolder = new DirectoryInfo(this.f2_data_path);
                foreach (string temp in myitems)
                {
                    //扫描文件夹，寻找子文件夹“不合格”
                    foreach (DirectoryInfo myNestFile in myFolder.GetDirectories())
                    {
                        //按照名字寻找文件夹
                        if (Regex.IsMatch(myNestFile.Name, temp))
                        {
                            //进入不合格文件夹
                            if (Regex.IsMatch(myNestFile.Name, "不合格"))
                            {
                                //进入子文件夹，扫描文件
                                foreach (FileInfo myfile in myNestFile.GetFiles())
                                {
                                    string a;
                                    a = myfile.Name.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries)[0];
                                    if (buhege.Count == 0)
                                    {
                                        try
                                        {
                                            buhege.Add(a.Split(new string[] { "report" }, StringSplitOptions.RemoveEmptyEntries)[0]);
                                        }
                                        catch
                                        {
                                            buhege.Add(a);
                                        }
                                    }


                                    //判断此文件名是否重复
                                    for (int i = 0; i < buhege.Count; i++)
                                    {
                                        if (Regex.IsMatch(myfile.Name, buhege[i]))
                                        {
                                            n++;
                                            break;
                                        }
                                    }

                                    //n == 0 不重复  n > 0 重复
                                    if (n == 0)
                                    {
                                        try
                                        {
                                            buhege.Add(a.Split(new string[] { "report" }, StringSplitOptions.RemoveEmptyEntries)[0]);
                                        }
                                        catch
                                        {
                                            buhege.Add(a);
                                        }
                                    }

                                    n = 0;
                                }
                            }
                            //进入合格文件夹
                            else
                            {
                                //进入子文件夹，扫描文件
                                foreach (FileInfo myfile in myNestFile.GetFiles())
                                {
                                    string a;
                                    a = myfile.Name.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries)[0];
                                    if (hege.Count == 0)
                                    {
                                        try
                                        {
                                            hege.Add(a.Split(new string[] { "report" }, StringSplitOptions.RemoveEmptyEntries)[0]);
                                        }
                                        catch
                                        {
                                            hege.Add(a);
                                        }
                                    }


                                    //判断此此叶片编号在合格列表中是否重复
                                    for (int i = 0; i < hege.Count; i++)
                                    {
                                        if (Regex.IsMatch(myfile.Name, hege[i]))
                                        {
                                            n++;
                                            break;
                                        }
                                    }

                                    //n == 0 不重复  n > 0 重复
                                    if (n == 0)
                                    {
                                        try
                                        {
                                            hege.Add(a.Split(new string[] { "report" }, StringSplitOptions.RemoveEmptyEntries)[0]);
                                        }
                                        catch
                                        {
                                            hege.Add(a);
                                        }
                                    }

                                    n = 0;
                                }
                            }

                        }
                    }
                }
            }
            catch
            { }
            return buhege;
        }

        public bool gen_report()
        {
            //初始
            MSWord.Document doc = null;
            MSWord.Application wordApp = new MSWord.Application();
            object missing = null;
            MSWord.Table table = null;
            MSWord.Table sub_table = null;
            string tuhao = "";
            string jyxm = "";
            try
            {
                doc = wordApp.Documents.Open(this.f2_muban_path);
            }
            //System.Reflection.Missing.Value
            catch
            {  }

            table = doc.Tables[1];

            //判断此叶片编号在不合格列表中是否存在
            for (int i = 0; i < hege.Count; i++)
            {
                for (int j = 0; j < buhege.Count; j++)
                {
                    if (Regex.IsMatch(hege[i], buhege[j]))
                    {
                        break;
                    }
                    if (j == buhege.Count - 1)
                    {
                        hege1.Add(hege[i]);
                    }
                }
            }

            //基表格的遍历
            for (int tabel_cnt = 1; tabel_cnt <= doc.Tables.Count; tabel_cnt++)
            {
                //表格内容遍历
                for (int i = 1; i <= table.Rows.Count; i++)
                {
                    for (int j = 1; j <= table.Columns.Count; j++)
                    {
                        //因为每一行的列数并不确定，当访问不存在的cell时会报错，所以使用try-catch
                        try
                        {
                            //填入 零件数量
                            if (table.Cell(i, j).Range.Text.ToString() == "零件数量\r\a")
                            {
                                table.Cell(i, j + 1).Range.Text = (buhege.Count + hege1.Count).ToString() + "件.";
                            }

                            //读取 零件图号
                            if (table.Cell(i, j).Range.Text.ToString() == "零件图号\r\a")
                            {
                                tuhao = table.Cell(i, j + 1).Range.Text;
                                tuhao = tuhao.Split(new string[] { "\r\a" }, StringSplitOptions.RemoveEmptyEntries)[0];
                                try
                                {
                                    string[] a = tuhao.Split(new string[] { "\r" }, StringSplitOptions.RemoveEmptyEntries);
                                    tuhao = "";
                                    foreach(string b in a)
                                    {
                                        tuhao += b;
                                    }
                                }
                                catch
                                { }
                            }

                            //读取 检验项目
                            if (table.Cell(i, j).Range.Text.ToString() == "检验项目\r\a")
                            {
                                jyxm = table.Cell(i, j + 1).Range.Text.Split(new string[] { "\r\a" }, StringSplitOptions.RemoveEmptyEntries)[0]; 
                            }

                            //寻找 报告内容
                            if (Regex.IsMatch(table.Cell(i, j).Range.Text.ToString(), "报告内容："))
                            {
                                int add_rows = 0;
                                //首先写入 合格报告
                                if (hege1.Count > 0)
                                {
                                    //换行
                                    table.Cell(i, j).Range.InsertAfter("\r\n");
                                    //
                                    table.Cell(i, j).Range.InsertAfter("经三坐标检测,共" + hege1.Count.ToString() + "件精铸件" + jyxm + "符合图样" + tuhao + "的要求，为合格件\r\n");

                                    object what = MSWord.WdUnits.wdLine;
                                    object count = 14;
                                    object dummy = System.Reflection.Missing.Value;

                                    wordApp.Selection.MoveDown(what, count);

                                    sub_table = doc.Tables.Add(wordApp.Selection.Range, 2, 8);
                                    
                                    sub_table.AutoFitBehavior(MSWord.WdAutoFitBehavior.wdAutoFitWindow);
                                    //写入表头
                                    sub_table.Cell(1, 1).Range.Text = "序号";
                                    sub_table.Cell(1, 2).Range.Text = "编号";
                                    sub_table.Cell(1, 3).Range.Text = "序号";
                                    sub_table.Cell(1, 4).Range.Text = "编号";
                                    sub_table.Cell(1, 5).Range.Text = "序号";
                                    sub_table.Cell(1, 6).Range.Text = "编号";
                                    sub_table.Cell(1, 7).Range.Text = "序号";
                                    sub_table.Cell(1, 8).Range.Text = "编号";
                                    sub_table.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                                    sub_table.Borders.OutsideLineWidth = table.Borders.OutsideLineWidth;
                                    sub_table.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                                    sub_table.Borders.InsideLineWidth = table.Borders.InsideLineWidth;

                                    sub_table.Cell(1, 1).Width = sub_table.Cell(1, 1).Width * 2f / 3f;
                                    sub_table.Cell(1, 2).Width = sub_table.Cell(1, 2).Width * 4f / 3f;
                                    sub_table.Cell(1, 3).Width = sub_table.Cell(1, 3).Width * 2f / 3f;
                                    sub_table.Cell(1, 4).Width = sub_table.Cell(1, 4).Width * 4f / 3f;
                                    sub_table.Cell(1, 5).Width = sub_table.Cell(1, 5).Width * 2f / 3f;
                                    sub_table.Cell(1, 6).Width = sub_table.Cell(1, 6).Width * 4f / 3f;
                                    sub_table.Cell(1, 7).Width = sub_table.Cell(1, 7).Width * 2f / 3f;
                                    sub_table.Cell(1, 8).Width = sub_table.Cell(1, 8).Width * 4f / 3f;
                                    sub_table.Cell(2, 1).Width = sub_table.Cell(2, 1).Width * 2f / 3f;
                                    sub_table.Cell(2, 2).Width = sub_table.Cell(2, 2).Width * 4f / 3f;
                                    sub_table.Cell(2, 3).Width = sub_table.Cell(2, 3).Width * 2f / 3f;
                                    sub_table.Cell(2, 4).Width = sub_table.Cell(2, 4).Width * 4f / 3f;
                                    sub_table.Cell(2, 5).Width = sub_table.Cell(2, 5).Width * 2f / 3f;
                                    sub_table.Cell(2, 6).Width = sub_table.Cell(2, 6).Width * 4f / 3f;
                                    sub_table.Cell(2, 7).Width = sub_table.Cell(2, 7).Width * 2f / 3f;
                                    sub_table.Cell(2, 8).Width = sub_table.Cell(2, 8).Width * 4f / 3f;

                                    int rows = hege1.Count / 4;
                                    add_rows = rows + 1;
                                    for (int m = 0; m < rows; m++ )
                                    {
                                        sub_table.Rows.Add(sub_table.Cell(table.Rows.Count, 1));
                                    }

                                    int jj = 1;
                                    int ii = 2;
                                    for (int m = 0; m < hege1.Count; m++)
                                    {
                                        sub_table.Cell(ii, jj).Range.Text = (m+1).ToString();
                                        sub_table.Cell(ii, jj + 1).Range.Text = hege1[m].ToString();
                                        if (ii == rows + 2)
                                        {
                                            jj += 2;
                                            ii = 2;
                                        }
                                        else
                                        {
                                            ii++;
                                        }
                                    }
                                }

                                //之后写入 不合格
                                if (buhege.Count > 0)
                                {
                                    table.Cell(i, j).Range.InsertAfter("经三坐标检测,共" + buhege.Count.ToString() + "件精铸件" + jyxm + "不符合图样" + tuhao + "的要求，为不合格件\r\n");
                                    //table.Cell(i, j).Range.Font.Spacing = table.Cell(i-1, 1).Range.Font.Spacing;

                                    object what = MSWord.WdUnits.wdLine;
                                    object count = 4 + add_rows;
                                    object dummy = System.Reflection.Missing.Value;

                                    wordApp.Selection.MoveDown(what, count);

                                    sub_table = doc.Tables.Add(wordApp.Selection.Range, 2, 8);

                                    sub_table.AutoFitBehavior(MSWord.WdAutoFitBehavior.wdAutoFitWindow);
                                    //写入表头
                                    sub_table.Cell(1, 1).Range.Text = "序号";
                                    sub_table.Cell(1, 2).Range.Text = "编号";
                                    sub_table.Cell(1, 3).Range.Text = "序号";
                                    sub_table.Cell(1, 4).Range.Text = "编号";
                                    sub_table.Cell(1, 5).Range.Text = "序号";
                                    sub_table.Cell(1, 6).Range.Text = "编号";
                                    sub_table.Cell(1, 7).Range.Text = "序号";
                                    sub_table.Cell(1, 8).Range.Text = "编号";

                                    sub_table.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                                    sub_table.Borders.OutsideLineWidth = table.Borders.OutsideLineWidth;
                                    sub_table.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                                    sub_table.Borders.InsideLineWidth = table.Borders.InsideLineWidth;

                                    sub_table.Cell(1, 1).Width = sub_table.Cell(1, 1).Width * 2f / 3f;
                                    sub_table.Cell(1, 2).Width = sub_table.Cell(1, 2).Width * 4f / 3f;
                                    sub_table.Cell(1, 3).Width = sub_table.Cell(1, 3).Width * 2f / 3f;
                                    sub_table.Cell(1, 4).Width = sub_table.Cell(1, 4).Width * 4f / 3f;
                                    sub_table.Cell(1, 5).Width = sub_table.Cell(1, 5).Width * 2f / 3f;
                                    sub_table.Cell(1, 6).Width = sub_table.Cell(1, 6).Width * 4f / 3f;
                                    sub_table.Cell(1, 7).Width = sub_table.Cell(1, 7).Width * 2f / 3f;
                                    sub_table.Cell(1, 8).Width = sub_table.Cell(1, 8).Width * 4f / 3f;
                                    sub_table.Cell(2, 1).Width = sub_table.Cell(2, 1).Width * 2f / 3f;
                                    sub_table.Cell(2, 2).Width = sub_table.Cell(2, 2).Width * 4f / 3f;
                                    sub_table.Cell(2, 3).Width = sub_table.Cell(2, 3).Width * 2f / 3f;
                                    sub_table.Cell(2, 4).Width = sub_table.Cell(2, 4).Width * 4f / 3f;
                                    sub_table.Cell(2, 5).Width = sub_table.Cell(2, 5).Width * 2f / 3f;
                                    sub_table.Cell(2, 6).Width = sub_table.Cell(2, 6).Width * 4f / 3f;
                                    sub_table.Cell(2, 7).Width = sub_table.Cell(2, 7).Width * 2f / 3f;
                                    sub_table.Cell(2, 8).Width = sub_table.Cell(2, 8).Width * 4f / 3f;

                                    int rows = buhege.Count / 4;
                                    for (int m = 0; m < rows; m++)
                                    {
                                        sub_table.Rows.Add(sub_table.Cell(table.Rows.Count, 1));
                                    }

                                    int jj = 1;
                                    int ii = 2;
                                    for (int m = 0; m < buhege.Count; m++)
                                    {
                                        sub_table.Cell(ii, jj).Range.Text = (m+1).ToString();
                                        sub_table.Cell(ii, jj+1).Range.Text = buhege[m].ToString();
                                        if (ii == rows + 2)
                                        {
                                            jj += 2;
                                            ii = 2;
                                        }
                                        else
                                        {
                                            ii++;
                                        }
                                    }
                                }

                            }
                        }
                        catch
                        { }
                    }
                }
            }

            //文档编辑完以后必须要保存，否则更改无效！
            string path = f2_save_path + "\\" + "检验报告" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString() + "-" + DateTime.Now.Hour.ToString() + "-" + DateTime.Now.Minute.ToString();
            doc.SaveAs(@path);
            doc.Close(ref missing, ref missing, ref missing);
            this.done = true;
            return true;
        }





    }


}
