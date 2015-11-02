using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Diagnostics;
using System.IO;
using System.Text;

public partial class yz : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void Button1_Click(object sender, EventArgs e)
    {
        //    System.Diagnostics.Process.Start(@".\..\..\keyfile\key.txt");
        //string path = Server.MapPath("..\\keyfile\\2.bat");
        //   string path1 = Server.MapPath("2.bat");

        // System.Diagnostics.Process.Start(path);
        //    string str="D:\\sarah\\WebSite3\\keyfile\\2.bat";
        // Process.Start(str);
        // System.Diagnostics.Process.Start(str);
        //  System.Diagnostics.Process.Start(path1);
        string sPath = "d://sarah//WebSite3//keyfile//2.bat";
        string sPath1 = Server.MapPath("~/keyfile/2.bat");
        string sDict = Server.MapPath("~/keyfile/");
        ProcessStartInfo psi = new ProcessStartInfo(sPath1);


        psi.UseShellExecute = false;

        psi.RedirectStandardOutput = true;

        psi.RedirectStandardInput = true;

        psi.RedirectStandardError = true;

        psi.Arguments = sPath;

        psi.WorkingDirectory = sDict;

        // Start the process

        System.Diagnostics.Process proc = System.Diagnostics.Process.Start(psi);

        // Attach the output for reading

        System.IO.StreamReader sOut = proc.StandardOutput;

        proc.Close();

        // Read the sOut to a string.
        string results = sOut.ReadToEnd().Trim();

        sOut.Close();

        // Write out the results.
        //string fmtStdOut = "<font face=courier size=0>{0}</font>";

        //this.Response.Write(String.Format(fmtStdOut, results.Replace(System.Environment.NewLine, "<br />")));


    }
    /*----------------点击校验------------------------------*/
    protected void Button3_Click(object sender, EventArgs e)
    {
        string mc = FileUpload2.FileName;
        string sPath4 = Server.MapPath("~/keyfile/jymy.jpg");
        FileUpload2.SaveAs(sPath4);


        string sPath2 = Server.MapPath("~/keyfile/jymy.jpg");
        readFile(sPath2);
    }
    /*------------------读文件--------------------------------*/
    protected void readFile(string path)
    {
        FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Read);
        StreamReader reader = new StreamReader(fs);

        //添加查找方法
        Encoding myEncoding = Encoding.GetEncoding("utf-8");


        string sData = "www.baidu.com";



        if (File.ReadAllText(path, myEncoding).Contains(sData))
        {
            lb_jy.Text = "通过";
        }
        else
        {
            lb_jy.Text = "不通过";
        }

        reader.Close();
        fs.Close();
    }


    /*--------------------确定上传需加密的图片----------------------------------------------*/
    protected void Button2_Click(object sender, EventArgs e)
    {
        string mc = FileUpload1.FileName;
        string sPath3 = Server.MapPath("~/keyfile/1.png");
        FileUpload1.SaveAs(sPath3);
        //   File.Move(FileUpload1.FileName, "1.png");

    }
    /*--------------------下载方法----------------------------------------------*/
    protected void Button4_Click(object sender, EventArgs e)
    {
        string downLoadFileName = "new83.jpg";
        this.Response.ContentType = "application/x-zip-compressed";
        string downLoadPath = this.Server.MapPath("~/keyfile/new83.jpg");
        this.Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", this.Server.UrlPathEncode(downLoadFileName)));
        this.Response.TransmitFile(downLoadPath);
    }
    /*--------------拆分.doc格式文档为.htm+文件夹--------------------------------------------------------*/
    protected void Button5_Click(object sender, EventArgs e)
    {
        //string mc = FileUpload3.FileName;
        //string sPath5 = Server.MapPath("~/keyfile/a.html");
        //FileUpload3.SaveAs(sPath5);
    }

    /*-------------------------word转化----------------------------------------------------------*/
    private static void WordToHtmlFile(string WordFilePath)
    {

        Microsoft.Office.Interop.Word.Application newApp = new Microsoft.Office.Interop.Word.Application();
        // 指定原文件和目标文件 
        object Source = WordFilePath;
        string SaveHtmlPath = WordFilePath.Substring(0, WordFilePath.Length - 3) + "html";
        object Target = SaveHtmlPath;

        // 缺省参数 
        object Unknown = Type.Missing;

        //为了保险,只读方式打开 
        object readOnly = true;

        // 打开doc文件 
        Microsoft.Office.Interop.Word.Document doc = newApp.Documents.Open(ref Source, ref Unknown,
        ref readOnly, ref Unknown, ref Unknown,
        ref Unknown, ref Unknown, ref Unknown,
        ref Unknown, ref Unknown, ref Unknown,
        ref Unknown, ref Unknown, ref Unknown,
        ref Unknown, ref Unknown);

        // 指定另存为格式(rtf) 
        object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatHTML;
        // 转换格式 
        doc.SaveAs(ref Target, ref format,
        ref Unknown, ref Unknown, ref Unknown,
        ref Unknown, ref Unknown, ref Unknown,
        ref Unknown, ref Unknown, ref Unknown,
        ref Unknown, ref Unknown, ref Unknown,
        ref Unknown, ref Unknown);

        // 关闭文档和Word程序 
        doc.Close(ref Unknown, ref Unknown, ref Unknown);
        newApp.Quit(ref Unknown, ref Unknown, ref Unknown);
    }

    public string wordToHtml(System.Web.UI.HtmlControls.HtmlInputFile wordFilePath)
    {
        Microsoft.Office.Interop.Word.ApplicationClass word = new Microsoft.Office.Interop.Word.ApplicationClass();
        Type wordType = word.GetType();
        Microsoft.Office.Interop.Word.Documents docs = word.Documents;
        // 打开文件
        Type docsType = docs.GetType();
        //应当先把文件上传至服务器然后再解析文件为html
        string filePath = uploadWord(wordFilePath);
        //判断是否上传文件成功
        if (filePath == "0")
            return "0";
        //判断是否为word文件
        if (filePath == "1")
            return "1";
        object fileName = filePath;
        Microsoft.Office.Interop.Word.Document doc = (Microsoft.Office.Interop.Word.Document)docsType.InvokeMember("Open",
        System.Reflection.BindingFlags.InvokeMethod, null, docs, new Object[] { fileName, true, true });
        // 转换格式，另存为html
        Type docType = doc.GetType();
        string filename = System.DateTime.Now.Year.ToString() + System.DateTime.Now.Month.ToString() + System.DateTime.Now.Day.ToString() +
        System.DateTime.Now.Hour.ToString() + System.DateTime.Now.Minute.ToString() + System.DateTime.Now.Second.ToString();
        //被转换的html文档保存的位置
        string ConfigPath = HttpContext.Current.Server.MapPath("~/keyfile/a.html");
        object saveFileName = ConfigPath;


        docType.InvokeMember("SaveAs", System.Reflection.BindingFlags.InvokeMethod,
        null, doc, new object[] { saveFileName, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatHTML });
        //关闭文档
        docType.InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod,
        null, doc, new object[] { null, null, null });
        // 退出 Word
        wordType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, word, null);
        //转到新生成的页面
        return ("/" + filename + ".html");
    }
    public string uploadWord(System.Web.UI.HtmlControls.HtmlInputFile uploadFiles)
    {
        if (uploadFiles.PostedFile != null)
        {
            string fileName = uploadFiles.PostedFile.FileName;
            int extendNameIndex = fileName.LastIndexOf(".");
            string extendName = fileName.Substring(extendNameIndex);
            string newName = "";
            try
            {
                //验证是否为word格式
                if (extendName == ".doc")
                {
                    DateTime now = DateTime.Now;
                    newName = now.DayOfYear.ToString() + uploadFiles.PostedFile.ContentLength.ToString();
                    //上传路径 指当前上传页面的同一级的目录下面的wordTmp路径
                    uploadFiles.PostedFile.SaveAs(System.Web.HttpContext.Current.Server.MapPath("wordTmp/" + newName + extendName));
                }
                else
                {
                    return "1";
                }
            }
            catch
            {
                return "0";
            }
            //return "http://" + HttpContext.Current.Request.Url.Host + HttpContext.Current.Request.ApplicationPath + "/wordTmp/" + newName + extendName;
            return System.Web.HttpContext.Current.Server.MapPath("wordTmp/" + newName + extendName);
        }
        else
        {
            return "0";
        }
    }
    protected void btnUpload_Click(object sender, EventArgs e)
    {
        try
        {
            //上传
            //uploadWord(File1);
            //转换
            string fileph = Server.MapPath("~/keyfile/a.doc");
            wordToHtml(File1);
        }
        catch (Exception ex)
        {
            throw ex;
        }
        finally
        {
            Response.Write("恭喜，转换成功！");
        }
    }

    /*------------------提取图片进行校验-------------------------------------------------*/
    protected void Button6_Click(object sender, EventArgs e)
    {
        string dPath = this.Server.MapPath("~/keyfile/a.files/");
        List<string> filenameList = new List<string>();
        DirectoryInfo dirInfo = new DirectoryInfo(dPath);
       
        foreach (FileInfo fileInfo in dirInfo.GetFiles())
        {
            filenameList.Add(fileInfo.Name);        
        }
        for (int n = 0; n < filenameList.Count; n++)
        {
            string sp = filenameList[n];

            readFile(dPath + sp);
            if (lb_jy.Text == "通过")
            {
                break;
            }
        }
       

    }
}