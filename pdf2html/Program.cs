using System;
using System.IO;
using Acrobat;
using System.Reflection;
using System.Windows.Forms;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using ThreadState = System.Diagnostics.ThreadState;



static class Program
{
    public class PDF2Html
    {
        private AcroAVDoc g_AVDoc = null;
        private String FileName = null;
        private String SavefileName = null;
        private String SaveDir = null;
        private int Timeout = 60000;




        public PDF2Html(String fileName,String saveDir,int timeout)
        {
            FileName = fileName;
            SaveDir = saveDir;
            Timeout = timeout;
        }

        public void OpenPDF()
        {
            g_AVDoc.Open(FileName, "");
        }


        public void KillProcess(String processName)
        {
            try{
                //get processes of Acrobat
                Process[] processIdAry = Process.GetProcessesByName(processName);
                if (processIdAry.Count() > 0){
                    for (int i = 0; i < processIdAry.Count(); i++){
                        processIdAry[i].Kill();
                        Console.WriteLine("Kill {0} sucessfully!", processName);
                    }
                }
                else{
                    Console.WriteLine("Not found {0} by name", processName);
                }
            }
            catch{
                Console.WriteLine("Kill {0} failed!", processName);
            }
        }

        public void Convert()
        {
            
            FileInfo fileinfo = new FileInfo(FileName);
            //file exits
            if (fileinfo.Exists)
            {
                if (g_AVDoc != null){
                    g_AVDoc.Close(0);
                }
                g_AVDoc = new AcroAVDoc();
                /*
                 * open pdf without blocking, to avoid adobe open automatic when pdf file is invalid
                 */
                MethodInvoker invoker = new MethodInvoker(OpenPDF);
                invoker.BeginInvoke(null, null);

            }//file not exist
            else{
                Console.WriteLine("{0} not exist!", FileName);
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
                    if (Directory.Exists(SaveDir))
                    {
                        SavefileName = SaveDir + '\\' + Path.GetFileNameWithoutExtension(FileName);
                    }
                    else
                    {
                        SavefileName = Path.GetDirectoryName(SaveDir) + Path.GetFileNameWithoutExtension(SaveDir);
                    }
                    SavefileName = SavefileName + ".html";
                    if (File.Exists(SavefileName))
                    {
                        File.Delete(SavefileName);
                    }

                    CAcroPDDoc pdDoc = (CAcroPDDoc)g_AVDoc.GetPDDoc();
                    //Acquire the Acrobat JavaScript Object interface from the PDDoc object
                    Object jsObj = pdDoc.GetJSObject();
                    Type T = jsObj.GetType();
                    object[] saveAsParam1 = { SavefileName, "com.adobe.acrobat.html-3-20" };
                    //object[] saveAsParam2 = { SavefileName + ".xml", "com.adobe.acrobat.xml-1-00" };
                    //object[] saveAsParam3 = { SavefileName + ".docx", "com.adobe.acrobat.docx" };
                    T.InvokeMember(
                                    "saveAs",
                                    BindingFlags.InvokeMethod |
                                    BindingFlags.Public |
                                    BindingFlags.Instance,
                                    null, jsObj, saveAsParam1);

 
                while (!File.Exists(SavefileName) && Timeout > 0){
                    Thread.Sleep(1000);
                    Timeout -= 1000;
                }

                if (!File.Exists(SavefileName)){
                    Console.WriteLine("{0} is invalid pdf or too big to convert, kill Acrobat!", FileName);
                    KillProcess("Acrobat");
                }
                else{
                    g_AVDoc.Close(0);
                    KillProcess("Acrobat");
                    Console.WriteLine("Convert {0} to HTML success!", Path.GetFileName(FileName));
                }

            }
            else
            {
                //Open file error
                Console.WriteLine("Open {0} failed, kill Acrobat!", FileName);
                KillProcess("Acrobat");
            }

        }

    }


    static void Main(string[] args)
    {
        String fileName=null;
        // set the default output dictionary
        String SaveDir = Application.StartupPath;
        // check the args
        if (args.Length == 1){
            fileName = args[0];
        }
        else if (args.Length == 2){
            fileName = args[0];
            SaveDir = args[1];
        }
        else{
            Console.WriteLine("Usage: PDF2XML.exe inputfile outputdir(optional)!");
        }

        PDF2Html Converter = new PDF2Html(fileName,SaveDir,60000);

        Thread t = new Thread(Converter.Convert);
        t.Start();
        try
        {
            t.Abort();
        }
        catch
        {

        }
       
        //Process pro = Process.GetCurrentProcess();
        //Console.WriteLine(pro.Id.ToString());

        

    }
}
