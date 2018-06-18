using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Threading;
using System.Xml.Linq;
using Aspose.Cells;
using Aspose.Pdf.Text;
using Aspose.Words;

namespace 我的工具箱
{


    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        // TMX
        string file_fullname;
        string filename;
        string filename_noext;
        string path_name;
        string filename_ext;
        // Excel
        string file_fullname_excel;
        string filename_excel;
        string filename_noext_excel;
        string path_name_excel;
        string filename_ext_excel;

        // Txt
        string file_fullname_txt;
        string filename_txt;
        string filename_noext_txt;
        string path_name_txt;
        string filename_ext_txt;
        string file_fullname_txtdir;
        List<string> file_fullnames;




        // Access
        string file_fullname_access;
        string filename_access;
        string filename_noext_access;
        string path_name_access;
        string filename_ext_access;



        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // prepare  output Excel
            string file_output = @"C:\Users\wey\Desktop\Part ZZ 每日\语料库\90 中文语料库\40 中文写作\文学描写辞典散文诗歌卷.xlsx";
            Workbook workbook_output = new Workbook();
            // Adding a new worksheet to the Workbook object
            int num = workbook_output.Worksheets.Add();
            // Obtaining the reference of the newly added worksheet by passing its sheet index
            Worksheet worksheet = workbook_output.Worksheets[num];
            worksheet.Name = "从Word读入段落";



            // read pdf document
            string file_name = @"C:\Users\wey\Desktop\Part ZZ 每日\语料库\90 中文语料库\40 中文写作\写作参考书籍\10 文学描写诗歌散文戏剧卷(上下)标记.docx";
            Document oDoc = new Aspose.Words.Document(file_name);

            Aspose.Pdf.Document doc = new Aspose.Pdf.Document(file_name);
            ParagraphAbsorber absorber = new ParagraphAbsorber();
            absorber.Visit(doc);






            // Save the workbook in output format
            workbook_output.Save(file_output, Aspose.Cells.SaveFormat.Xlsx);
            MessageBox.Show("OK!!!");

        }

        private void ListBox_Drop(object sender, DragEventArgs e)
        {

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                file_fullname = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
                path_name = Path.GetDirectoryName(file_fullname);
                filename = Path.GetFileName(file_fullname);
                filename_noext = Path.GetFileNameWithoutExtension(file_fullname);
                filename_ext = Path.GetExtension(file_fullname);
            }

            lbFile.Items.Add(file_fullname);


        }

        // convert tmx to excel
        private void TMX_To_Excel(object sender, RoutedEventArgs e)
        {

            //           string strS = sourceL.Text;
            //           string strT = targetL.Text;
            int i = 0;
            // Prepere output XLS
            string file_output = path_name + @"\" + filename_noext + "." + "xlsx";
            Workbook workbook_output = new Workbook();
            Worksheet ws_output = workbook_output.Worksheets[0];

            // 把TXT文件读入到字符串  
            StreamReader str = new StreamReader(file_fullname);
            string ss = str.ReadToEnd();


            string s1 = "<seg>";
            string s2 = "</seg>";
            Regex rg = new Regex("(?<=(" + s1 + "))[.\\s\\S]*?(?=(" + s2 + "))", RegexOptions.Multiline | RegexOptions.Singleline);
            MatchCollection myMatches = rg.Matches(ss);

            foreach (Match nextMatch in myMatches)
            {
                if (i % 2 == 0)
                {
                    ws_output.Cells[i / 2, 0].PutValue(nextMatch.Value);
                }
                else
                {
                    ws_output.Cells[i / 2, 1].PutValue(nextMatch.Value);

                }
                i++;
            }

            //// Save the workbook in output format
            workbook_output.Save(file_output, Aspose.Cells.SaveFormat.Xlsx);





        }


        private void Excel_to_TMX(object sender, RoutedEventArgs e)
        {
            Workbook workbook_input = new Workbook(file_fullname);
            Worksheet ws_input = workbook_input.Worksheets[0];
            // Prepere output XLS
            string file_output = path_name + @"\" + filename_noext + "." + "tmx";
            // 创建XML文档和根节点
            // 定义一个name space
            XNamespace ns = "http://www.w3.org/XML/1998/namespace";

            XDocument xDoc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), //创建版本<?xml version="1.0" encoding="utf-8" standalone="yes" ?> );

                 new XComment("Created:" + DateTime.Now.ToString()),//创建注释节点
                  new XElement("tmx", new XAttribute("version", "1.4"),
                  new XAttribute(XNamespace.Xmlns + "xml", ns)
                  ));                     // 创建根节点


            //获取根节点
            XElement root = xDoc.Root;

            // 在根节点下面添加body 和header子节点
            XElement xheader = new XElement("header",
                new XAttribute("creationtool", "OffAna"),
                new XAttribute("segtype", "sentence"),
                   new XAttribute("adminlang", "en-US"),
                   new XAttribute("srclang", sourceL.Text),
                              new XAttribute("datatype", "xml"),
                              new XAttribute("creationdate", DateTime.Now.ToString()),
                              new XAttribute("creationid", "OffAna"));


            XElement xbody = new XElement("body");
            root.Add(xheader);
            root.Add(xbody);


            // 在body节点下面添加各子节点
            for (int i = 0; i < ws_input.Cells.Rows.Count; i++)
            {
                // tu 节点
                XElement xtu = new XElement("tu",
                 new XAttribute("creationdate", DateTime.Now.ToString()),
                 new XAttribute("creationid", "OffAna"));

                // 第一个tuv节点
                XElement xtuv1 = new XElement("tuv", new XAttribute(ns + "lang", sourceL.Text));
                XElement seg1 = new XElement("seg", ws_input.Cells[i, 0].Value);
                xtuv1.Add(seg1);
                // 第二个tuv节点
                XElement xtuv2 = new XElement("tuv", new XAttribute(ns + "lang", targetL.Text));
                XElement seg2 = new XElement("seg", ws_input.Cells[i, 1].Value);
                xtuv1.Add(seg2);

                // tuv --》 tu; ti ---> body
                xtu.Add(xtuv1);
                xtu.Add(xtuv2);
                xbody.Add(xtu);

            }

            // 替换   “tuv lang=” 到 "tuv xml:lang="

            xDoc.Save(file_output);


        }
        private void Convert_TMX_EXCEL(object sender, RoutedEventArgs e)
        {
            if (filename_ext == ".tmx")
            {
                TMX_To_Excel(sender, e);
            }
            else
            {

                Excel_to_TMX(sender, e);
            }
            MessageBox.Show("OK!!!");
        }



        private void ExcelInput_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                file_fullname_excel = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
                path_name_excel = Path.GetDirectoryName(file_fullname_excel);
                filename_excel = Path.GetFileName(file_fullname_excel);
                filename_noext_excel = Path.GetFileNameWithoutExtension(file_fullname_excel);
                filename_ext_excel = Path.GetExtension(file_fullname_excel);
            }

            ExcelInput.Items.Add(file_fullname_excel);

        }



        private void Color_Worksheets(object sender, RoutedEventArgs e)
        {

            //    int key_col  = Convert.ToInt32(col_key.Text)-1;
            //    Workbook workbook = new Workbook(file_fullname_excel);
            //    Worksheet ws = workbook.Worksheets[1];
            //    string temp = ws.Cells[0,2].Value.ToString();
            //    // 辅助列
            //    int count = 0;
            //    for (int i = 0; i < ws.Cells.Rows.Count; i++)
            //    {
            //        if (ws.Cells[i, 2].Value.ToString() != temp)
            //        {
            //            count++;
            //            temp = ws.Cells[i, 2].Value.ToString();
            //        }

            //        ws.Cells[i, 7].PutValue(count);


            //    }
            //}


            //    // background color
            //    for (int i = 0; i < ws.Cells.Rows.Count; i++)
            //    {
            //        for (int j = 0; j < ws.Cells.Columns.Count; j++)
            //        {
            //            Aspose.Cells.Style s = ws.Cells[i, j].GetStyle();
            //            s.ForegroundColor = Color.Cyan;

            //            s.Pattern = BackgroundType.Solid;
            //            ws.Cells[i, j].SetStyle(s);

            //        }

            //    }


            //    //// Save the workbook in output format
            //    //workbook.Save(file_fullname_excel, Aspose.Cells.SaveFormat.Xlsx);
            //    //MessageBox.Show("OK!!!");
        }

        private void Clear_Input(object sender, RoutedEventArgs e)
        {
            file_fullname_excel = string.Empty;
            filename_excel = string.Empty; ;
            filename_noext_excel = string.Empty; ;
            path_name_excel = string.Empty; ;
            filename_ext_excel = string.Empty;
            ExcelInput.Items.Clear(); ;
        }

        private void Clear_InputExcel(object sender, RoutedEventArgs e)
        {
            file_fullname = string.Empty;
            filename = string.Empty; ;
            filename_noext = string.Empty; ;
            path_name = string.Empty; ;
            filename_ext = string.Empty;
            lbFile.Items.Clear(); ;

        }

        private void Odd_Even(object sender, RoutedEventArgs e)
        {
            // excel 工作簿至少有2表 1.源 2.目标
            Workbook workbook = new Workbook(file_fullname_excel);
            // 增加一个合并表
            Worksheet ws_co = workbook.Worksheets.Add("Arranged");
            Worksheet ws_input = workbook.Worksheets[0];
            Worksheet ws_output = workbook.Worksheets[1];

            // 遍历拷贝Cell
            int j = 0;
            for (int i = 0; i < ws_input.Cells.Rows.Count; i = i + 2)
            {
                ws_output.Cells[j, 1].PutValue(ws_input.Cells[i, 0].Value);
                ws_output.Cells[j, 0].PutValue(ws_input.Cells[i + 1, 0].Value);
                //              ws_output.Cells[j, 2].PutValue(ws_input.Cells[i, 1].Value);
                j++;

            }

            // Save the workbook in output format
            workbook.Save(file_fullname_excel, Aspose.Cells.SaveFormat.Xlsx);

            MessageBox.Show("OK!!!");


        }

        private void TXT_TOEXCEL(object sender, RoutedEventArgs e)
        {

            string[] orig_lines;

            string file_fullname_txt;
            // 1.  打开文件
            // Read each line of the file into a string array. Each element
            // of the array is one line of the file.
            file_fullname_txt = file_fullname_excel;
            string file_output = path_name_excel + @"\" + filename_noext_excel + "." + "xlsx";
            Workbook workbook_output = new Workbook();
            Worksheet ws_output = workbook_output.Worksheets[0];

            orig_lines = System.IO.File.ReadAllLines(file_fullname_txt);

            // 遍历拷贝Cell
            int i = 0;
            for (i = 0; i < orig_lines.Length; i = i + 1)
            {
                ws_output.Cells[i, 0].PutValue(orig_lines[i]);


            }

            // Save the workbook in output format
            workbook_output.Save(file_output, Aspose.Cells.SaveFormat.Xlsx);

            MessageBox.Show("OK!!!");
            //2. 使用lambda表达式过滤掉空字符串
            //           new_lines = orig_lines.Where(s => !string.IsNullOrEmpty(s)).ToArray();

            // 3. 写回文件
            //
            //            以下情况可能会导致异常：
            //               文件已存在并且为只读。
            //              路径名可能太长。
            //              磁盘可能已满。

            //try
            //{
            //    File.WriteAllLines(FullName, new_lines);

            //}
            //catch (Exception e)
            //{
            //    MessageBox.Show(e.ToString());
            //}
        }

        private void Clear_InputAccess(object sender, RoutedEventArgs e)
        {
            file_fullname_access = string.Empty;
            filename_access = string.Empty; ;
            filename_noext_access = string.Empty; ;
            path_name_access = string.Empty; ;
            filename_ext_access = string.Empty;
            AccessInput.Items.Clear(); ;

        }

        private void AccessInput_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                file_fullname_access = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
                path_name_access = Path.GetDirectoryName(file_fullname_access);
                filename_access = Path.GetFileName(file_fullname_access);
                filename_noext_access = Path.GetFileNameWithoutExtension(file_fullname_access);
                filename_ext_access = Path.GetExtension(file_fullname_access);
            }

            AccessInput.Items.Add(file_fullname_access);

        }

        private void ExtractAccessToExcel(object sender, RoutedEventArgs e)
        {

            // 准备excel 文件，工作簿，工作表
            string file_output = path_name_access + @"\" + filename_noext_access + "." + "xlsx";
            Workbook workbook_output = new Workbook();
            Worksheet ws_output = workbook_output.Worksheets[0];


            // connection 有两个attributes： connectionstring + providename.  
      
            String s0 = "Provider= Microsoft.ACE.OLEDB.12.0;Data Source=";

            var dbConnectionString = s0 + file_fullname_access;
            var conn = DbProviderFactories.GetFactory("JetEntityFrameworkProvider").CreateConnection();
            conn.ConnectionString = dbConnectionString;


            List<Phrase> phrases_total = new List<Phrase>();
            List<Phrase> phrases_tmp = new List<Phrase>();



            // read Access： out of memory issue

            using (var context = new MyDbContext(conn))
            {
                //               var query = context.Langs.Where(lang => lang.num < 20000 );
                          //     var query = context.Langs.Where(lang => lang.num >= 25000 && lang.num < 50000);

                //   var query = context.Langs.Where(lang => lang.num >= 50000 && lang.num < 75000);
                var query = context.Langs;
                foreach (var item in query)

                {

                    //                  Word word = new Word(item.word, item.Exp,item.ID);
                    Word word = new Word(item.word, item.Exp);
                    phrases_tmp = word.ExtractPhrases();
                   phrases_total.AddRange(phrases_tmp);

                }

            }

            // 写入Excel文件

            for (int i = 0; i < phrases_total.Count; i++)
            {
                ws_output.Cells[i, 0].PutValue(phrases_total[i].DE_CN);
                ws_output.Cells[i, 1].PutValue(phrases_total[i].Keyword);
                ws_output.Cells[i, 2].PutValue(phrases_total[i].Type);
                ws_output.Cells[i, 3].PutValue(phrases_total[i].Source);
                ws_output.Cells[i, 3].PutValue(phrases_total[i].ID);
            }


            // Save the workbook in output format
            workbook_output.Save(file_output, Aspose.Cells.SaveFormat.Xlsx);

            MessageBox.Show("OK!!!");

        }








        private void TXT_Drop(object sender, DragEventArgs e)
        {
            //if (e.Data.GetDataPresent(DataFormats.FileDrop))
            //{
            //    file_fullname_excel = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            //    path_name_excel = Path.GetDirectoryName(file_fullname_excel);
            //    filename_excel = Path.GetFileName(file_fullname_excel);
            //    filename_noext_excel = Path.GetFileNameWithoutExtension(file_fullname_excel);
            //    filename_ext_excel = Path.GetExtension(file_fullname_excel);
            //}

            //ExcelInput.Items.Add(file_fullname_excel);
        }

        // 清理德语助手，法语助手
        private void ZS_TOEXCEL(object sender, RoutedEventArgs e)
        {
            // excel input
            Workbook workbook_input = new Workbook(file_fullname_excel);
            Worksheet ws_input = workbook_input.Worksheets[0];

            // excel 文件，工作簿，工作表
            string file_output = path_name_excel + @"\" + filename_noext_excel + "_p02" + "." + "xlsx";
            Workbook workbook_output = new Workbook();
            Worksheet ws_output = workbook_output.Worksheets[0];

            string s1;
            string s2;
            string str1;
            string str2;
            string str3;
            string str4;

            for (int i = 0; i < ws_input.Cells.Rows.Count; i++)
            {
                //       string ss = ws_input.Cells[i, 0].Value.ToString();

                //  Möchten Sie Möhrrübe <span class="key">kaufen</span>？
                //  Anna: Man kann sich vieles <span class="key">kaufen</span>.

                string ss = ws_input.Cells[i, 4].Value.ToString();
                //                 str1 = ss.Replace("<SPAN class=key>", "");
                //                 str2 = str1.Replace("</SPAN>", "");

                str3 = ss.Replace("class=key>", "");

                ws_output.Cells[i, 0].PutValue(str3);
                //    if (ws_input.Cells[i, 2].Value.ToString() == "德语原声例句")
                //    {
                //        //    // p01: 取得原声例句的source
                //        //    //                  string str = ws_input.Cells[i, 0].Value.ToString();
                //        //    //                  string[] strArray = str.Split('<');
                //        //    //                  ws_output.Cells[i, 6].PutValue(strArray[0]);

                //        // p02: 提取德语句子
                //        s1 = "=line>";
                //        s2 = "<IM";


                //        //    //p03 提取中文
                //        //    // 返回在此字符串中最右边出现的指定子字符串的索引
                //        //    string s3 = ">";
                //        //    int index = ss.LastIndexOf(s3);
                //        //    string s4;
                //        //    if (index > 0)
                //        //    { s4 = ss.Substring(index+1);
                //        //        ws_output.Cells[i, 0].PutValue(s4);
                //        //    }
                //        //    else
                //        //    { ws_output.Cells[i, 0].PutValue("  "); }
                //        if (ss != "abbuchen" && ss != "abbrechen" && ss != "abbauen")
                //        {
                //            Regex rg = new Regex("(?<=(" + s1 + "))[.\\s\\S]*?(?=(" + s2 + "))", RegexOptions.Multiline | RegexOptions.Singleline);
                //            MatchCollection matchlist = rg.Matches(ss);
                //            if (matchlist.Count > 0)
                //            { ws_output.Cells[i, 0].PutValue(matchlist[0].Value); }

                //            else
                //            { ws_output.Cells[i, 0].PutValue("  "); }
                //        }

                //        else
                //        {
                //              ws_output.Cells[i, 0].PutValue(ws_input.Cells[i, 4].Value);
                //        }
                //    }




                //    else {
                //            // p02: 提取德语句子
                //                      s1 = "=line>";
                //                      s2 = "<IM";
                ////            s1 = "xp>";
                ////            s2 = "<S";

                //            Regex rg = new Regex("(?<=(" + s1 + "))[.\\s\\S]*?(?=(" + s2 + "))", RegexOptions.Multiline | RegexOptions.Singleline);
                //            MatchCollection matchlist = rg.Matches(ss);
                //            if (matchlist.Count > 0)
                //            { ws_output.Cells[i, 0].PutValue(matchlist[0].Value); }
                //            else
                //            {
                //            s1 = "=line>";
                //            s2 = "</SPAN><B";
                //            rg = new Regex("(?<=(" + s1 + "))[.\\s\\S]*?(?=(" + s2 + "))", RegexOptions.Multiline | RegexOptions.Singleline);
                //            matchlist = rg.Matches(ss);
                //            if (matchlist.Count > 0)
                //            { ws_output.Cells[i, 0].PutValue(matchlist[0].Value); }
                //            else
                //            { ws_output.Cells[i, 0].PutValue("  "); }

                //        }

                //        }



            }

            // Save the workbook in output format
            workbook_output.Save(file_output, Aspose.Cells.SaveFormat.Xlsx);

            MessageBox.Show("OK!!!");

        }

        private void MERGE_EXCEL(object sender, RoutedEventArgs e)
        {
            //Worksheet ws0 = workbook_input.Worksheets[0];
            //Worksheet ws1 = workbook_input.Worksheets[1];
            //Worksheet ws2 = workbook_input.Worksheets[2];
            //int num_0 = ws0.Cells.Rows.Count;
            //int num_1 = ws1.Cells.Rows.Count;
            //int num_2 = ws2.Cells.Rows.Count;
            // excel 文件，工作簿，工作表
            string file_output = path_name_excel + @"\" + filename_noext_excel + "_p02" + "." + "xlsx";
            Workbook workbook_output = new Workbook();
            Worksheet ws_output = workbook_output.Worksheets[0];


            // excel input
            Workbook workbook_input = new Workbook(file_fullname_excel);
            int num_sheets = workbook_input.Worksheets.Count;

            int m;
            m = 0;

            for (int i = 0; i < num_sheets; i++)
            {

                // copy one sheet    
                for (int j = 0; j < workbook_input.Worksheets[i].Cells.Rows.Count; j++)

                {
                    // copy one row
                    for (int k = 0; k < 20; k++)
                    {
                        ws_output.Cells[m, k].PutValue(workbook_input.Worksheets[i].Cells[j, k].Value);

                    }
                    m++;
                }



            }







            // Save the workbook in output format
            workbook_output.Save(file_output, Aspose.Cells.SaveFormat.Xlsx);

            MessageBox.Show("OK!!!");
        }

        private void ZS_FR_TOEXCEL(object sender, RoutedEventArgs e)
        {
            // excel input
            Workbook workbook_input = new Workbook(file_fullname_excel);
            Worksheet ws_input = workbook_input.Worksheets[0];

            // excel 文件，工作簿，工作表
            string file_output = path_name_excel + @"\" + filename_noext_excel + "_p02" + "." + "xlsx";
            Workbook workbook_output = new Workbook();
            Worksheet ws_output = workbook_output.Worksheets[0];


            string str1;
            string str2;
            string str3;
            string str4;

            for (int i = 0; i < ws_input.Cells.Rows.Count; i++)
            {


                // 1  取得原声例句的source
                //if (ws_input.Cells[i, 2].Value.ToString() == "法语原声例句")
                //{
                //    string str = ws_input.Cells[i, 0].Value.ToString();
                //    string[] strArray = str.Split('<');
                //    ws_output.Cells[i, 6].PutValue(strArray[0]);
                //}
                //else
                //{
                //    ws_output.Cells[i, 6].PutValue("  ");

                //}


                // 2 提取中文
                //   string ss = ws_input.Cells[i, 0].Value.ToString();
                //    string s1 = "exp>";
                //    string s2 = "<SPAN";
                //    Regex rg = new Regex("(?<=(" + s1 + "))[.\\s\\S]*?(?=(" + s2 + "))", RegexOptions.Multiline | RegexOptions.Singleline);

                //    if (ws_input.Cells[i, 2].Value.ToString() == "法语原声例句")
                //    {
                //        返回在此字符串中最右边出现的指定子字符串的索引
                //        string s3 = ">";
                //        int index = ss.LastIndexOf(s3);
                //        string s4;
                //        if (index > 0)
                //        {
                //            s4 = ss.Substring(index + 1);
                //            ws_output.Cells[i, 0].PutValue(s4);
                //        }
                //        else
                //        { ws_output.Cells[i, 0].PutValue("  "); }
                //    }
                //    else
                //    {
                //        MatchCollection matchlist = rg.Matches(ss);
                //        if (matchlist.Count > 0) { ws_output.Cells[i, 0].PutValue(matchlist[0].Value); }
                //        else
                //        {
                //            matchlist = rg.Matches(ss);
                //            if (matchlist.Count > 0) { ws_output.Cells[i, 0].PutValue(matchlist[0].Value); }
                //            else { ws_output.Cells[i, 0].PutValue("  "); }
                //        }
                //    }
                //}

                // 3. 提取法文
                //string ss = ws_input.Cells[i, 0].Value.ToString();
                //string s1 = "=line>";
                //string s2 = "<IMG";
                //Regex rg = new Regex("(?<=(" + s1 + "))[.\\s\\S]*?(?=(" + s2 + "))", RegexOptions.Multiline | RegexOptions.Singleline);
                //MatchCollection matchlist = rg.Matches(ss);
                //if (matchlist.Count > 0)
                //{ ws_output.Cells[i, 0].PutValue(matchlist[0].Value); }
                //else
                //{
                //    s1 = "=line>";
                //    s2 = "</SPAN><B";
                //    rg = new Regex("(?<=(" + s1 + "))[.\\s\\S]*?(?=(" + s2 + "))", RegexOptions.Multiline | RegexOptions.Singleline);
                //    matchlist = rg.Matches(ss);
                //    if (matchlist.Count > 0)
                //    { ws_output.Cells[i, 0].PutValue(matchlist[0].Value); }
                //    else
                //    { ws_output.Cells[i, 0].PutValue("  "); }


                //}

                // 4. 清理法文

                string ss = ws_input.Cells[i, 4].Value.ToString();
                //                str1 = ss.Replace("<SPAN class=key>", "");
                //               str2 = str1.Replace("</SPAN>", "");
                 str2 = ss.Replace("<SPAN", "");
                 str3 = str2.Replace("class=key>", "");
                ws_output.Cells[i, 0].PutValue(str3);

            }

            //  Anna: Man kann sich vieles <span class="key">kaufen</span>.

            //string ss = ws_input.Cells[i, 4].Value.ToString();
            ////                 str1 = ss.Replace("<SPAN class=key>", "");
            ////                 str2 = str1.Replace("</SPAN>", "");

            //str3 = ss.Replace("class=key>", "");

            //ws_output.Cells[i, 0].PutValue(str3);
            // Save the workbook in output format
            workbook_output.Save(file_output, Aspose.Cells.SaveFormat.Xlsx);

                MessageBox.Show("OK!!!");
         }

        private void ZS_ES_TOEXCEL(object sender, RoutedEventArgs e)
        {
            // excel input
            Workbook workbook_input = new Workbook(file_fullname_excel);
            Worksheet ws_input = workbook_input.Worksheets[0];

            // excel 文件，工作簿，工作表
            string file_output = path_name_excel + @"\" + filename_noext_excel + "_p02" + "." + "xlsx";
            Workbook workbook_output = new Workbook();
            Worksheet ws_output = workbook_output.Worksheets[0];


            string str1;
            string str2;
            string str3;
            string str4;

            for (int i = 0; i < ws_input.Cells.Rows.Count; i++)
            {


                // 1  取得原声例句的source
                //if (ws_input.Cells[i, 2].Value.ToString() == "西语原声例句")
                //{
                //    string str = ws_input.Cells[i, 0].Value.ToString();
                //    string[] strArray = str.Split('<');
                //    ws_output.Cells[i, 6].PutValue(strArray[0]);
                //}
                //else
                //{
                //    ws_output.Cells[i, 6].PutValue("  ");

                //}


                ////               2 提取中文
                //                  string ss = ws_input.Cells[i, 0].Value.ToString();


                // //                  返回在此字符串中最右边出现的指定子字符串的索引
                //                       string s3 = ">";
                //                   int index = ss.LastIndexOf(s3);
                //                   string s4;
                //                   if (index > 0)
                //                   {
                //                       s4 = ss.Substring(index + 1);
                //                       ws_output.Cells[i, 0].PutValue(s4);
                //                   }
                //                   else
                //                   { ws_output.Cells[i, 0].PutValue("  "); }



                // 3. 提取西班牙文
                //string ss = ws_input.Cells[i, 0].Value.ToString();
                //string s1 = "=line>";
                //string s2 = "<IMG";
                //Regex rg = new Regex("(?<=(" + s1 + "))[.\\s\\S]*?(?=(" + s2 + "))", RegexOptions.Multiline | RegexOptions.Singleline);
                //MatchCollection matchlist = rg.Matches(ss);
                //if (matchlist.Count > 0)
                //{ ws_output.Cells[i, 0].PutValue(matchlist[0].Value); }
                //else
                //{
                //    s1 = "=line>";
                //    s2 = "</SPAN><B";
                //    rg = new Regex("(?<=(" + s1 + "))[.\\s\\S]*?(?=(" + s2 + "))", RegexOptions.Multiline | RegexOptions.Singleline);
                //    matchlist = rg.Matches(ss);
                //    if (matchlist.Count > 0)
                //    { ws_output.Cells[i, 0].PutValue(matchlist[0].Value); }
                //    else
                //    { ws_output.Cells[i, 0].PutValue("  "); }


                //}

                // 4. 清理西班牙文

                string ss = ws_input.Cells[i, 3].Value.ToString();
                               str1 = ss.Replace("<SPAN class=key>", "");
                               str2 = str1.Replace("</SPAN>", "");
                str3 = str2.Replace("<SPAN", "");
                str4 = str3.Replace("class=key>", "");
                ws_output.Cells[i, 0].PutValue(str4);

            }

            workbook_output.Save(file_output, Aspose.Cells.SaveFormat.Xlsx);

            MessageBox.Show("OK!!!");
        }

        private void ExcelInput_DropDir(object sender, DragEventArgs e)
        {

        }

        private void TXT_DropDir(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                file_fullname_txtdir = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
                var files = Directory.GetFiles(file_fullname_txtdir, "*.csv");

                //path_name_access = Path.GetDirectoryName(file_fullname_access);
                //filename_access = Path.GetFileName(file_fullname_access);
                //filename_noext_access = Path.GetFileNameWithoutExtension(file_fullname_access);
                //filename_ext_access = Path.GetExtension(file_fullname_access);
            }
            string dir_disp = "Dir: " + file_fullname_txtdir + @"\";
            TXTInputDir.Items.Add(dir_disp);
        }
    }
}

