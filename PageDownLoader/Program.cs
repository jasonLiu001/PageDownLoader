using HtmlAgilityPack;
using Jurassic;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Xml;

namespace PageDownLoader
{
    class Program
    {
        //the page count.
        public const int pageCount = 2;
        //存储页面内容集合
        public static Dictionary<string, string> pageContent = new Dictionary<string, string>();

        public static XmlDocument xml=new XmlDocument();

        //最终的处理结果路径
        public static string xmlFilePath = System.Environment.CurrentDirectory + "\\data.xml";

        //搜索关键字文件路径
        public static string keyWordsFilePath = System.Environment.CurrentDirectory + "\\keywords.txt";

        static void Main(string[] args)
        {
            initXmlDoc();
            List<string> keyWords = GetKeyWords();

            if (keyWords.Count == 0)
            {
                Console.WriteLine("搜索词不能为空！请检查程序运行目录下是否存在keywords.txt文件");
                Console.ReadKey();
                return;
            }

            int keyWordIndex = 1;
            foreach (string word in keyWords)
            {
                Console.WriteLine("正在抓取第" + keyWordIndex + "个关键词的内容，请稍等...");               
                DownloadPageContent(keyWordIndex,word);
                Console.WriteLine("第" + keyWordIndex + "个关键词内容抓取完毕！");
                keyWordIndex++;
            }           
            AnalysisPageContent();
            Console.WriteLine("数据全部处理完毕！结果已保存在程序目录下的data.xml文件中");
            Console.ReadKey();
        }

        static void initXmlDoc()
        {
            XmlDeclaration xmlDecalaration = xml.CreateXmlDeclaration("1.0", "gb2312", null);
            xml.AppendChild(xmlDecalaration);//追加文档声明

            //加入根元素
            XmlElement rootElement = xml.CreateElement("", "Root", "");
            xml.AppendChild(rootElement);
        }

        static void SaveXmlDoc()
        {
            xml.Save(xmlFilePath);
        }

        static void DownloadPageContent(int keyWordIndex, string queryString)
        {
            queryString = HttpUtility.UrlEncode(queryString, Encoding.UTF8).ToLower();
            string pageinfo = string.Empty;          
            for (int i = 1; i <= pageCount; i++)
            {
                Thread.Sleep(1000);
                try
                {
                    //queryString = "%e6%88%91%e4%bb%ac";
                    string url = "http://weixin.sogou.com/weixin?type=2&query=" + queryString + "&ie=utf8&_ast=1410346624&_asf=null&w=01029901&p=40040100&dp=1&cid=null&page=" + i;

                    HttpWebRequest myReq = (HttpWebRequest)HttpWebRequest.Create(url);
                    myReq.Accept = "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/x-shockwave-flash, application/vnd.ms-excel, application/vnd.ms-powerpoint, application/msword, */*";
                    myReq.UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; .NET CLR 2.0.50727)";
                    myReq.Timeout = 12000;

                    HttpWebResponse myRep = (HttpWebResponse)myReq.GetResponse();
                    Stream myStream = myRep.GetResponseStream();
                    StreamReader sr = new StreamReader(myStream, Encoding.UTF8);
                    pageinfo = sr.ReadToEnd().ToString();
                    if (!string.IsNullOrEmpty(pageinfo))
                    {
                        string contentKey = keyWordIndex.ToString() + "_" + i.ToString();
                        pageContent.Add(contentKey, pageinfo);
                        Console.WriteLine("第"+keyWordIndex+"个搜索词,第" + i + "页数据读取成功！");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }          

        }

        static void AnalysisPageContent()
        {
            if (pageContent.Count > 0)
            {
                Console.WriteLine("页面内容解析操作开始.....");
                foreach (KeyValuePair<string, string> keyValue in pageContent)
                {
                    string[] keyStringArr = keyValue.Key.Split('_');
                    Console.WriteLine("正在解析第"+keyStringArr[0]+"个搜索词，第" + keyStringArr[1]+ "页数据....");
                    string pageContentString = keyValue.Value;
                    HtmlDocument doc = new HtmlDocument();
                    doc.LoadHtml(pageContentString);
                    //HtmlNodeCollection nodes=doc.DocumentNode.SelectNodes("/html[1]/body[1]/div[1]/div[2]");
                    HtmlNode mainDiv=doc.GetElementbyId("main");
                    HtmlNode resultDiv = mainDiv.SelectNodes("./div[1]/div[3]")[0].ChildNodes[0];
                    HtmlNodeCollection resultNodes = resultDiv.ChildNodes;
                    foreach (HtmlNode node in resultNodes)
                    {
                        if (node.Name == "div")
                        {   //标题
                            HtmlNode title= node.SelectNodes("./div[2]/h4[1]/a[1]")[0];
                            string titleString = RemoveHtmlElements(title.InnerHtml);
                            titleString = HttpUtility.HtmlDecode(titleString);

                            //链接
                            string titleLink = title.Attributes["href"].Value;
                            titleLink=HttpUtility.HtmlDecode(titleLink);

                            //内容
                            HtmlNode summary = node.SelectNodes("./div[2]/p[1]")[0];
                            string summaryString = RemoveHtmlElements(summary.InnerHtml);
                            summaryString = HttpUtility.HtmlDecode(summaryString);

                            //时间戳
                            HtmlNode postTime = node.SelectNodes("./div[2]/p[2]")[0];
                            string postTimeString = postTime.InnerHtml;
                            postTimeString = postTimeString.Substring((postTimeString.IndexOf('\'')+1), (postTimeString.LastIndexOf('\'') - postTimeString.IndexOf('\'')-1));
                            postTimeString = GetPostTime(postTimeString);

                            SaveXmlNodes(titleString, titleLink, summaryString, postTimeString);
                        }
                    }

                    Console.WriteLine("第"+keyStringArr[0]+"个关键词，第"+keyStringArr[1]+"页数据解析完毕！");
                }

                SaveXmlDoc();//保存结果               
            }
        }

        /// <summary>
        /// 将页面上的脚本转换成时间
        /// </summary>
        /// <param name="postTime">格式是：1409710649 </param>
        /// <returns></returns>
        static string GetPostTime(string postTime)
        {
            string timeSpan = string.Empty;
            var JSEngine = new ScriptEngine();
            JSEngine.Evaluate("function vrTimeHandle552(time){ if (time) {var type = [\"1分钟前\", \"分钟前\", \"小时前\", \"天前\", \"周前\", \"个月前\", \"年前\"];"
                +" var secs = (new Date().getTime())/1000 - time; if(secs < 60){ return type[0]; }else if(secs < 3600){	return Math.floor(secs/60) + type[1]; }else if(secs < 24*3600){"
                +" return Math.floor(secs/3600) + type[2]; }else if(secs < 24*3600 *7){	return Math.floor(secs/(24*3600)) + type[3]; }else if(secs < 24*3600*31){"
                +" return Math.round(secs/(24*3600*7)) + type[4]; }else if(secs < 24*3600*365){	return Math.round(secs/(24*3600*31)) + type[5];	}else if(secs >= 24*3600*365){"
                +" return Math.round(secs/(24*3600*365)) + type[6]; }else {	return ''; } } else { return ''; }}");           
            timeSpan = JSEngine.CallGlobalFunction<string>("vrTimeHandle552", int.Parse(postTime));
            return timeSpan;
        }

        static void SaveXmlNodes(string title, string titleLink, string summary, string postTime)
        {    
            XmlNode root = xml.SelectSingleNode("Root");

            XmlElement eData = xml.CreateElement("Data");//创建一个<Data>节点 
            eData.SetAttribute("source", "weixin");//设置该节点source属性  
            eData.SetAttribute("timePoint",DateTime.Now.ToString());

            XmlElement eTitle = xml.CreateElement("title");
            eTitle.InnerText = title;
            eData.AppendChild(eTitle);

            XmlElement etitleLink = xml.CreateElement("titleLink");
            etitleLink.InnerText = titleLink;
            eData.AppendChild(etitleLink);

            XmlElement esummary = xml.CreateElement("summary");
            esummary.InnerText = summary;
            eData.AppendChild(esummary);

            XmlElement epostTime = xml.CreateElement("postTime");
            epostTime.InnerText = postTime;
            eData.AppendChild(epostTime);

            root.AppendChild(eData);
        }

        static List<string> GetKeyWords()
        {
            List<string> list = new List<string>();

            if (File.Exists(keyWordsFilePath))
            {
                string[] keyWords = File.ReadAllLines(keyWordsFilePath,Encoding.Default);
                foreach (string word in keyWords)
                {
                    list.Add(word);
                }
                
            }             

            return list;
        }

        ///<summary>   
        ///去除HTML标记 
        ///</summary>   
        ///<param name="NoHTML">包括HTML的源码</param>   
        ///<returns>已经去除后的文字</returns>
        static string RemoveHtmlElements(string Htmlstring)
        {

            //删除脚本   

            Htmlstring =

                Regex.Replace(Htmlstring, @"<script[^>]*?>.*?</script>",

                "", RegexOptions.IgnoreCase);

            //删除HTML   

            Htmlstring = Regex.Replace(Htmlstring, @"<(.[^>]*)>", "", RegexOptions.IgnoreCase);

            Htmlstring = Regex.Replace(Htmlstring, @"([\r\n])[\s]+", "", RegexOptions.IgnoreCase);

            Htmlstring = Regex.Replace(Htmlstring, @"-->", "", RegexOptions.IgnoreCase);

            Htmlstring = Regex.Replace(Htmlstring, @"<!--.*", "", RegexOptions.IgnoreCase);



            Htmlstring = Regex.Replace(Htmlstring, @"&(quot|#34);", "\"", RegexOptions.IgnoreCase);

            Htmlstring = Regex.Replace(Htmlstring, @"&(amp|#38);", "&", RegexOptions.IgnoreCase);

            Htmlstring = Regex.Replace(Htmlstring, @"&(lt|#60);", "<", RegexOptions.IgnoreCase);

            Htmlstring = Regex.Replace(Htmlstring, @"&(gt|#62);", ">", RegexOptions.IgnoreCase);

            Htmlstring = Regex.Replace(Htmlstring, @"&(nbsp|#160);", "   ", RegexOptions.IgnoreCase);

            Htmlstring = Regex.Replace(Htmlstring, @"&(iexcl|#161);", "\xa1", RegexOptions.IgnoreCase);

            Htmlstring = Regex.Replace(Htmlstring, @"&(cent|#162);", "\xa2", RegexOptions.IgnoreCase);

            Htmlstring = Regex.Replace(Htmlstring, @"&(pound|#163);", "\xa3", RegexOptions.IgnoreCase);

            Htmlstring = Regex.Replace(Htmlstring, @"&(copy|#169);", "\xa9", RegexOptions.IgnoreCase);

            Htmlstring = Regex.Replace(Htmlstring, @"&#(\d+);", "", RegexOptions.IgnoreCase);



            Htmlstring.Replace("<", "");

            Htmlstring.Replace(">", "");

            Htmlstring.Replace("\r\n", "");

            Htmlstring.Trim();

            return Htmlstring;

        }
    }
}
