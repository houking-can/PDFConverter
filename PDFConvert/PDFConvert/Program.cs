using System;
using System.IO;
using Acrobat;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.InteropServices;

static class Program
{

    public class PDFConvert
    {
        private AcroAVDoc g_AVDoc = null;
        public string inputFile = null;
        public string saveFile = null;
        public string outputDir = null;
        public string format = null;

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", EntryPoint = "SendMessage")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, string lParam);

        [DllImport("user32.dll ")]
        public static extern IntPtr FindWindowEx(IntPtr parent, IntPtr childe, string strclass, string FrmText);


        public PDFConvert(string inputFile, string outputDir, string format = "html")
        {
            this.inputFile = inputFile;
            this.outputDir = outputDir;
            if (format == null)
            {
                format = "html";
            }
            this.format = format;
        }


        public void OpenPDF()
        {
            g_AVDoc.Open(inputFile, "");
        }


        public string GetConvID()
        {
            switch (format)
            {
                case "html": return "com.adobe.acrobat.html";
                case "xml": return "com.adobe.acrobat.xml-1-00";
                case "jpeg": return "com.adobe.acrobat.jpeg";
                case "jpe": return "com.adobe.acrobat.jpeg";
                case "jpg": return "com.adobe.acrobat.jpeg";
                case "png": return "com.adobe.acrobat.png";
                case "doc": return "com.adobe.acrobat.doc";
                case "docx": return "com.adobe.acrobat.docx";
                case "jpf": return "com.adobe.acrobat.jp2k";
                case "jpx": return "com.adobe.acrobat.jp2k";
                case "jp2": return "com.adobe.acrobat.jp2k";
                case "j2k": return "com.adobe.acrobat.jp2k";
                case "j2c": return "com.adobe.acrobat.jp2k";
                case "jpc": return "com.adobe.acrobat.jp2k";
                case "ps": return "com.adobe.acrobat.ps";
                case "txt": return "com.adobe.acrobat.plain-text";
                case "accesstext": return "com.adobe.acrobat.accesstext	";
                case "rtf": return "com.adobe.acrobat.rtf";
                case "tiff": return "com.adobe.acrobat.tiff";
                case "tif": return "com.adobe.acrobat.tiff";
                case "xlsx": return "com.adobe.acrobat.xlsx";
                default: return "com.adobe.acrobat.html";
            }
        }


        public string SetSaveFileName()
        {
            if (Directory.Exists(outputDir))
            {
                saveFile = outputDir + '\\' + Path.GetFileNameWithoutExtension(inputFile);
            }
            else
            {
                saveFile = Path.GetDirectoryName(outputDir) + Path.GetFileNameWithoutExtension(outputDir);
            }
            saveFile = saveFile + "." + format;
            return saveFile;
        }

        public void Convert()
        {

            FileInfo fileinfo = new FileInfo(inputFile);
            //file exits
            if (fileinfo.Exists)
            {
                if (g_AVDoc != null)
                {
                    g_AVDoc.Close(0);
                }
                g_AVDoc = new AcroAVDoc();
                /*
                 * open pdf without blocking, to avoid adobe open automatic when pdf file is invalid
                 */
                MethodInvoker invoker = new MethodInvoker(OpenPDF);
                invoker.BeginInvoke(null, null);

            }//file not exist
            else
            {
                Console.WriteLine("{0} not exist!", inputFile);
            }
            /* 
             * if open pdf success
             * wait 200 Milliseconds to open file, to process the case which pdf is invalid, 
             * lead to process blocking.
             */
            Thread.Sleep(200);

            /* even if the g_AVDoc is valid, it may be the txt or html format, 
             * which acrobat can also process, but it will lead blocking, should be killed.
             */
            if (g_AVDoc.IsValid())
            {
                SetSaveFileName();
                CAcroPDDoc pdDoc = (CAcroPDDoc)g_AVDoc.GetPDDoc();
                //Acquire the PDFConvert JavaScript Object interface from the PDDoc object
                Object jsObj = pdDoc.GetJSObject();
                Type T = jsObj.GetType();
                string cConvID = GetConvID();
                object[] saveAsParam = { saveFile, cConvID };

                T.InvokeMember(
                                "saveAs",
                                BindingFlags.InvokeMethod |
                                BindingFlags.Public |
                                BindingFlags.Instance,
                                null, jsObj, saveAsParam);

                try
                {
                    g_AVDoc.Close(0);
                }
                catch
                {
                    Console.WriteLine("PDF has been closed, no need to close again!");
                }
            }
            else
            {
                //Open file error
                Console.WriteLine("Open {0} failed, kill PDFConvert!", inputFile);
                try
                {
                    g_AVDoc.Close(0);
                }
                catch
                {
                    Console.WriteLine("PDF has been closed, no need to close again!");
                }
            }

        }

    }


    static void Main(string[] args)
    {
        string inputFile = null;
        string outputDir = Application.StartupPath;
        string format = null;
        // get args
        var arguments = CommandLineArgumentParser.Parse(args);
        if (!arguments.Has("-i"))
        {
            Console.WriteLine("Usage: PDFConvert.exe \n" +
                "-i inputfile \n" +
                "-o outputdir\tdefault: run dictionary)\n" +
                "-f format\tdefault: html\n\t\t" +
                "Support: xml,txt,doc,docx,\n\t\t" +
                "ps,jpeg,jpe,jpg,\n\t\t" +
                "jpf,jpx,j2k,j2c,jpc,rtf,\n\t\t" +
                "accesstext,tif,tiff)\n" +
                "-h help\n");
            Environment.Exit(0);
        }
        else
        {
            inputFile = arguments.Get("-i").Next;
        }
        if (arguments.Has("-o"))
        {
            outputDir = arguments.Get("-o").Next;
        }
        if (arguments.Has("-f"))
        {
            format = arguments.Get("-f").Next;
        }

        if (arguments.Has("-h"))
        {
            Console.WriteLine("Usage: PDFConvert.exe \n" +
                "-i inputfile \n" +
                "-o outputdir\tdefault: run dictionary)\n" +
                "support: xml,txt,doc,docx,\n\t\t" +
                "ps,jpeg,jpe,jpg,\n\t\t" +
                "jpf,jpx,j2k,j2c,jpc,rtf,\n\t\t" +
                "accesstext,tif,tiff)\n" +
                "-h help\n");
            Environment.Exit(0);
        }

        PDFConvert PDF = new PDFConvert(inputFile, outputDir, format);
        string saveFile = PDF.SetSaveFileName();
        if (File.Exists(saveFile))
        {
            Environment.Exit(0);
        }
        Thread convert = new Thread(PDF.Convert);
        convert.Start();
        convert.Join();
    }
}

