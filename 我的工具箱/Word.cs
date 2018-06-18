using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace 我的工具箱
{
 public   class Word
    {
        public string ID { get; set; }
        public string word { get; set; }
        public string text { get; set; }
        public  List<String> get_match(string s1, string s2, string ss)
        {
            List<string> list_tmp = new List<string>();
            Regex rg = new Regex("(?<=(" + s1 + "))[.\\s\\S]*?(?=(" + s2 + "))", RegexOptions.Multiline | RegexOptions.Singleline);
            MatchCollection matchlist = rg.Matches(ss);
            foreach (Match nextMatch in matchlist)
            {
                list_tmp.Add(nextMatch.Value);
            }
            return list_tmp;
        }

        //       public Word(string wd,string Exp,string id)
        public Word(string wd, string Exp)
        {
            word = wd;
            text = Exp;
     //       ID = id;

        }
        public List<Phrase> ExtractPhrases()
        {
            List<Phrase> phrases = new List<Phrase>();
            // 正则提取

            // 西班牙语常用例句库
            //DE_CN
            string s1 = "<LI>";
            string s2 = "<SPAN class=ljid ";

            // 西语原声例句
            // DE_CN
            string s3 = "channel_title>";
            string s4 = "</P><!";
            // Source 
            //          string s9 = "=channel_title>";
            //          string s10 = "</SPAN>";

            // 提取例句  
            List<string> list_common_decn = get_match(s1, s2, text);
            List<string> list_orig_decn = get_match(s3, s4, text);
            //           List<string> list_orig_source = get_match(s9, s10, text);

            for (int i = 0; i < list_orig_decn.Count; i++)
            {
                Phrase phrase_tmp = new Phrase();

                phrase_tmp.DE_CN = list_orig_decn[i];
                // 清除De中的HTML符号
                //            phrase_tmp.DE = clear_String(phrase_tmp.DE);

                phrase_tmp.Keyword = word;
                phrase_tmp.ID = ID;
                phrase_tmp.Type = "西语原声例句";
                phrase_tmp.Source = "   ";
                phrases.Add(phrase_tmp);

            }


            // construct 常用例句 object list
            for (int i = 0; i < list_common_decn.Count; i++)
            {
                Phrase phrase_tmp = new Phrase();

                phrase_tmp.DE_CN = list_common_decn[i];
                // 清除De中的HTML符号
                phrase_tmp.Keyword = word;
                phrase_tmp.ID = ID;
                phrase_tmp.Type = "西班牙语常用例句库";
                phrase_tmp.Source = "   ";
                phrases.Add(phrase_tmp);

            }
            if (phrases.Count == 0)
            {
                Phrase phrase_tmp = new Phrase();
                phrase_tmp.Keyword = word;
                phrase_tmp.ID = ID;
                phrases.Add(phrase_tmp);
            }
            return phrases;
        }

        private string clear_String(string str)
        {

            //  Möchten Sie Möhrrübe <span class="key">kaufen</span>？
            //  Anna: Man kann sich vieles <span class="key">kaufen</span>.
            str = str.Replace("<SPAN class=key>", "");
            str = str.Replace("</SPAN>", "");

            return str;
        }



    }




}
