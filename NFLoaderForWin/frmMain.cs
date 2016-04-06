using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace NFLoaderForWin
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
            txtConfigPath.Text = Application.StartupPath + "\\NFLoaderConfig.xml";
            if (System.IO.File.Exists(txtConfigPath.Text))
            {
                NFRoot = ReadNFRootFromXML(txtConfigPath.Text);
            }
            CheckConfig();
        }

        private string NFRoot;
        private bool isConfigChecked = false;
        private Dictionary<int, string> ProcessIDs = new Dictionary<int, string>();
        private List<string> servers = new List<string>();

        private void btnOpenConfig_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "NFLoaderConfig|*.xml";
                ofd.FileName = txtConfigPath.Text;
                ofd.ShowDialog();
                if (ofd.FileName != txtConfigPath.Text)
                {
                    txtConfigPath.Text = ofd.FileName;
                }

                NFRoot = ReadNFRootFromXML(txtConfigPath.Text);
                CheckConfig();
            }


        }

        private void CheckConfig()
        {
            if (NFRoot != "")
            {
                if (System.IO.Directory.Exists(NFRoot))
                {
                    txtLog.AppendText("NFRoot:" + NFRoot + "\r\n");
                    isConfigChecked = true;
                }
                else
                {
                    txtLog.AppendText("NFRoot:Error, Please Check Your NFRoot:" + NFRoot + "\r\n");
                }
            }
            txtLog.SelectionStart = txtLog.TextLength;
            txtLog.ScrollToCaret();
        }

        private string ReadNFRootFromXML(string text)
        {
            if (System.IO.File.Exists(text))
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(txtConfigPath.Text);    //加载Xml文件  
                XmlElement rootElem = doc.DocumentElement;   //获取根节点  
                XmlNodeList xmlNode = rootElem.GetElementsByTagName("NFRoot");
                foreach (XmlNode node in xmlNode)
                {
                    string strName = ((XmlElement)node).GetAttribute("Path");
                    return strName;
                }
            }
            return "";
        }

        private void ReadServerConfigFromLogicClass(string logicClass)
        {

            if (System.IO.File.Exists(logicClass))
            {
                servers.Clear();
                XmlDocument doc = new XmlDocument();
                doc.Load(logicClass);    //加载Xml文件  
                XmlElement rootElem = doc.DocumentElement;   //获取根节点  
                XmlNodeList xmlNode = rootElem.GetElementsByTagName("Class");
                foreach (XmlNode node in xmlNode)
                {
                    XmlNodeList sub2Nodes = ((XmlElement)node).GetElementsByTagName("Class");  //获取age子XmlElement集合  
                    foreach (XmlNode node2 in sub2Nodes)
                    {
                        string strName = ((XmlElement)node2).GetAttribute("Id");
                        if (strName == "Server")
                        {
                            string serverConfigPaht = ((XmlElement)node2).GetAttribute("InstancePath");
                            XmlDocument docServer = new XmlDocument();
                            docServer.Load(NFRoot + "\\_Out\\Server\\" + serverConfigPaht);
                            XmlElement rootElemServ = docServer.DocumentElement;   //获取根节点  
                            XmlNodeList xmlNodeServ = rootElemServ.GetElementsByTagName("Object");
                            foreach (XmlNode serNode in xmlNodeServ)
                            {
                                string servName = ((XmlElement)serNode).GetAttribute("ID");
                                servers.Add(servName);
                            }
                        }

                    }
                }
            }
        }

        private void btnStartAll_Click(object sender, EventArgs e)
        {
            if (!isConfigChecked) { return; }
            ProcessIDs.Clear();
            CopyDllFile();
            int ServCount = 0;
            string faildList = "";
            //1.启动NFMasterServer
            if (StartServer("NFMasterServer"))
                ServCount++;
            else
                faildList += " NFMasterServer ";
            Thread.Sleep(1000);

            //2.启动NFWorldServer
            if (StartServer("NFWorldServer"))
                ServCount++;
            else
                faildList += " NFWorldServer ";
            Thread.Sleep(1000);

            //3.启动NFLoginServer
            if (StartServer("NFLoginServer"))
                ServCount++;
            else
                faildList += " NFLoginServer ";
            Thread.Sleep(1000);

            //4.启动NFGameServer
            if (StartServer("NFGameServer1"))
                ServCount++;
            else
                faildList += " NFGameServer1 ";
            Thread.Sleep(1000);

            //5.启动NFProxyServer
            if (StartServer("NFProxyServer1"))
                ServCount++;
            else
                faildList += " NFProxyServer1 ";

            if (ServCount == 5)
            {
                txtLog.Text += "All Servers have been started!\r\n";
            }
            else
            {
                txtLog.Text += "Some server haven't been started!" + faildList + "\r\n";
            }
        }

        private void CopyDllFile()
        {
            System.IO.File.Copy(NFRoot + "Dependencies\\lib\\Debug\\libmysql.dll", NFRoot + "\\_Out\\Server\\Debug\\NFGameServer1\\libmysql.dll", true);
            System.IO.File.Copy(NFRoot + "Dependencies\\lib\\Debug\\libmysql.dll", NFRoot + "\\_Out\\Server\\Debug\\NFLoginServer\\libmysql.dll", true);
            System.IO.File.Copy(NFRoot + "Dependencies\\lib\\Debug\\libmysql.dll", NFRoot + "\\_Out\\Server\\Debug\\NFMasterServer\\libmysql.dll", true);
            System.IO.File.Copy(NFRoot + "Dependencies\\lib\\Debug\\libmysql.dll", NFRoot + "\\_Out\\Server\\Debug\\NFProxyServer1\\libmysql.dll", true);
            System.IO.File.Copy(NFRoot + "Dependencies\\lib\\Debug\\libmysql.dll", NFRoot + "\\_Out\\Server\\Debug\\NFWorldServer\\libmysql.dll", true);

            System.IO.File.Copy(NFRoot + "Dependencies\\lib\\Debug\\mysqlpp_d.dll", NFRoot + "\\_Out\\Server\\Debug\\NFGameServer1\\mysqlpp_d.dll", true);
            System.IO.File.Copy(NFRoot + "Dependencies\\lib\\Debug\\mysqlpp_d.dll", NFRoot + "\\_Out\\Server\\Debug\\NFLoginServer\\mysqlpp_d.dll", true);
            System.IO.File.Copy(NFRoot + "Dependencies\\lib\\Debug\\mysqlpp_d.dll", NFRoot + "\\_Out\\Server\\Debug\\NFMasterServer\\mysqlpp_d.dll", true);
            System.IO.File.Copy(NFRoot + "Dependencies\\lib\\Debug\\mysqlpp_d.dll", NFRoot + "\\_Out\\Server\\Debug\\NFProxyServer1\\mysqlpp_d.dll", true);
            System.IO.File.Copy(NFRoot + "Dependencies\\lib\\Debug\\mysqlpp_d.dll", NFRoot + "\\_Out\\Server\\Debug\\NFWorldServer\\mysqlpp_d.dll", true);
        }

        private bool StartServer(string v)
        {
            string serverFileName = NFRoot + "\\_Out\\Server\\Debug\\" + v + "\\NFPluginLoader_d.exe";
            string serverFilePath = NFRoot + "\\_Out\\Server\\Debug\\" + v;
            if (!System.IO.File.Exists(serverFileName))
            {
                return false;
            }
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo();
                psi.FileName = serverFileName;
                //psi.UseShellExecute = false;
                psi.WorkingDirectory = serverFilePath;
                //psi.CreateNoWindow = true;

                var sPro = Process.Start(psi);
                ProcessIDs.Add(sPro.Id, v);

                txtLog.Text += v + " [" + sPro.Id + "] started!\r\n";
                txtLog.SelectionStart = txtLog.TextLength;
                txtLog.ScrollToCaret();
                sPro.EnableRaisingEvents = true;//设置进程终止时触发的时间
                sPro.Exited += new EventHandler(myprocess_Exited);//发现外部程序关闭即触发方法myprocess_Exited
                return true;
            }
            catch
            {
                return false;
            }


        }

        private void ReadServerConfig()
        {
            string logicClass = "\\_Out\\Server\\NFDataCfg\\Struct\\LogicClass.xml";
            ReadServerConfigFromLogicClass(NFRoot + logicClass);
        }

        private void myprocess_Exited(object sender, EventArgs e)//被触发的程序
        {
            var pid = ((System.Diagnostics.Process)sender).Id;
            lock (txtLog)
            {
                try
                {
                    txtLog.Text += ProcessIDs[pid] + " [" + pid + "] close\r\n";
                    txtLog.SelectionStart = txtLog.TextLength;
                    txtLog.ScrollToCaret();
                    ProcessIDs.Remove(pid);
                }
                catch
                {
                    return;
                }
            }
        }

        private void btnStopAll_Click(object sender, EventArgs e)
        {
            foreach (var serv in ProcessIDs.Keys)
            {
                Process.GetProcessById(serv).Kill();
            }
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (var serv in ProcessIDs.Keys)
            {
                Process.GetProcessById(serv).Kill();
            }
            ProcessIDs.Clear();
        }
    }
}
