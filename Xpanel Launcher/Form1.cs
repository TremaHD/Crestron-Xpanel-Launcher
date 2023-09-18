using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Management;
using System.IO.Compression;
using System.Text;
using System.Linq;
using System.Xml;
using Newtonsoft.Json;
using System.Xml.Linq;

namespace Xpanel_Launcher
{
    public partial class Form1 : Form
    {
        public RootObject myJsonData { get; set; }
        const int MAXITEM = 8;
        const string JSONFILE = "Recently Opened.json";

        string sourceDirectory = @"C:\Program Files (x86)\Crestron\XPanel\CrestronXPanel";
        string targetDirectory = @".\Crestron Xpanels";

        TextBox[] txtIPArray = new TextBox[MAXITEM];
        TextBox[] txtIpIdArray = new TextBox[MAXITEM];
        TextBox[] txtRoomIdArray = new TextBox[MAXITEM];
        ToolTip[] toolTipArray = new ToolTip[MAXITEM];
        Label[] txtFilenameArray = new Label[MAXITEM];
        TextBox[] txtWindowTitleArray = new TextBox[MAXITEM];
        Button[] btnBrowseArray = new Button[MAXITEM];
        Button[] btnLaunchArray = new Button[MAXITEM];
        CheckBox[] chkSSLArray = new CheckBox[MAXITEM];
        Button[] btnClearArray = new Button[MAXITEM];

        public Form1()
        {
            InitializeComponent();

            #region "Define GUI object arrays"
            for (int x = 0; x< MAXITEM; x++)
            {
                toolTipArray[x] = new ToolTip();
                toolTipArray[x].InitialDelay = 100;
                toolTipArray[x].ShowAlways = true;
            }
            txtIPArray[0] = txtAdresseIP1;
            txtIPArray[1] = txtAdresseIP2;
            txtIPArray[2] = txtAdresseIP3;
            txtIPArray[3] = txtAdresseIP4;
            txtIPArray[4] = txtAdresseIP5;
            txtIPArray[5] = txtAdresseIP6;
            txtIPArray[6] = txtAdresseIP7;
            txtIPArray[7] = txtAdresseIP8;

            txtIpIdArray[0] = txtIpId1;
            txtIpIdArray[1] = txtIpId2;
            txtIpIdArray[2] = txtIpId3;
            txtIpIdArray[3] = txtIpId4;
            txtIpIdArray[4] = txtIpId5;
            txtIpIdArray[5] = txtIpId6;
            txtIpIdArray[6] = txtIpId7;
            txtIpIdArray[7] = txtIpId8;

            txtRoomIdArray[0] = txtRoomId1;
            txtRoomIdArray[1] = txtRoomId2;
            txtRoomIdArray[2] = txtRoomId3;
            txtRoomIdArray[3] = txtRoomId4;
            txtRoomIdArray[4] = txtRoomId5;
            txtRoomIdArray[5] = txtRoomId6;
            txtRoomIdArray[6] = txtRoomId7;
            txtRoomIdArray[7] = txtRoomId8;

            txtFilenameArray[0] = txtPath1;
            txtFilenameArray[1] = txtPath2;
            txtFilenameArray[2] = txtPath3;
            txtFilenameArray[3] = txtPath4;
            txtFilenameArray[4] = txtPath5;
            txtFilenameArray[5] = txtPath6;
            txtFilenameArray[6] = txtPath7;
            txtFilenameArray[7] = txtPath8;

            txtWindowTitleArray[0] = txtShortName1;
            txtWindowTitleArray[1] = txtShortName2;
            txtWindowTitleArray[2] = txtShortName3;
            txtWindowTitleArray[3] = txtShortName4;
            txtWindowTitleArray[4] = txtShortName5;
            txtWindowTitleArray[5] = txtShortName6;
            txtWindowTitleArray[6] = txtShortName7;
            txtWindowTitleArray[7] = txtShortName8;

            btnBrowseArray[0] = btnBrowse1;
            btnBrowseArray[1] = btnBrowse2;
            btnBrowseArray[2] = btnBrowse3;
            btnBrowseArray[3] = btnBrowse4;
            btnBrowseArray[4] = btnBrowse5;
            btnBrowseArray[5] = btnBrowse6;
            btnBrowseArray[6] = btnBrowse7;
            btnBrowseArray[7] = btnBrowse8;

            btnLaunchArray[0] = btnLaunch1;
            btnLaunchArray[1] = btnLaunch2;
            btnLaunchArray[2] = btnLaunch3;
            btnLaunchArray[3] = btnLaunch4;
            btnLaunchArray[4] = btnLaunch5;
            btnLaunchArray[5] = btnLaunch6;
            btnLaunchArray[6] = btnLaunch7;
            btnLaunchArray[7] = btnLaunch8;

            chkSSLArray[0] = chkSSL1;
            chkSSLArray[1] = chkSSL2;
            chkSSLArray[2] = chkSSL3;
            chkSSLArray[3] = chkSSL4;
            chkSSLArray[4] = chkSSL5;
            chkSSLArray[5] = chkSSL6;
            chkSSLArray[6] = chkSSL7;
            chkSSLArray[7] = chkSSL8;

            btnClearArray[0] = btnClear1;
            btnClearArray[1] = btnClear2;
            btnClearArray[2] = btnClear3;
            btnClearArray[3] = btnClear4;
            btnClearArray[4] = btnClear5;
            btnClearArray[5] = btnClear6;
            btnClearArray[6] = btnClear7;
            btnClearArray[7] = btnClear8;

            for (int x=0; x < MAXITEM; x++)
            {
                btnBrowseArray[x].Click += new System.EventHandler(ClickBtnBrowse);
                btnClearArray[x].Click += new System.EventHandler(ClickBtnClear);
                btnLaunchArray[x].Click += new System.EventHandler(ClickBtnLaunch);
                btnBrowseArray[x].Tag = x;
                btnClearArray[x].Tag = x;
                btnLaunchArray[x].Tag = x;
            }
            #endregion
           
            CreateFolders();
            LoadJsonFile(JSONFILE);
        }

        public void ClickBtnBrowse(Object sender, System.EventArgs e)
        {
            Button btn = (Button)sender;
            int idx = (int)btn.Tag;
            int index;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtFilenameArray[idx].Text = openFileDialog1.FileName;
                toolTipArray[idx].SetToolTip(txtFilenameArray[idx], openFileDialog1.FileName);
                index = openFileDialog1.FileName.LastIndexOf("\\");
                txtFilenameArray[idx].Text = openFileDialog1.FileName.Substring(index + 1);
            }
        }

        public void ClickBtnClear(Object sender, System.EventArgs e)
        {
            Button btn = (Button)sender;
            int idx = (int)btn.Tag;

            txtFilenameArray[idx].Text = "";
            toolTipArray[idx].SetToolTip(txtFilenameArray[idx], "");
            txtFilenameArray[idx].Text = "";
            txtIPArray[idx].Text = "";
            txtIpIdArray[idx].Text = "";
            txtRoomIdArray[idx].Text = "";
            chkSSLArray[idx].Checked = false; 
            txtWindowTitleArray[idx].Text = "";
        }

        private void ClickBtnLaunch(Object sender, System.EventArgs e)
        {
            Button btn = (Button)sender;
            int idx = (int)btn.Tag;
            var FileName = toolTipArray[idx].GetToolTip(txtFilenameArray[idx]);
            var IP = txtIPArray[idx].Text;
            var IPID = txtIpIdArray[idx].Text;
            var RoomID = txtRoomIdArray[idx].Text;
            var WindowTitle = txtWindowTitleArray[idx].Text;
            var useSSL = chkSSLArray[idx].Checked;
            if(File.Exists(FileName))
                GetNextAvailableXpanel(FileName, IP, IPID, RoomID, WindowTitle, useSSL);
        }

        private void Form1_Closing(object sender, FormClosingEventArgs e)
        {
            SaveJsonFile();
        }
        
        public int LoadJsonFile(string JsonFile)
        {
            int response = 0;
            int index = 0;
            myJsonData = new RootObject();
            try
            {
                if (File.Exists(JsonFile))
                {
                    using (StreamReader file = new StreamReader(JsonFile))
                    {
                        string data = file.ReadToEnd();
                        file.Close();
                        myJsonData = JsonConvert.DeserializeObject<RootObject>(data);
                        if (myJsonData != null)
                        {
                            for (int idx = 0; idx < myJsonData.HistoryList.Data.Count; idx++)
                            {
                                txtIPArray[idx].Text = myJsonData.HistoryList.Data[idx].IPAddress;
                                txtIpIdArray[idx].Text = myJsonData.HistoryList.Data[idx].IpId;
                                txtRoomIdArray[idx].Text = myJsonData.HistoryList.Data[idx].RoomId;
                                toolTipArray[idx].SetToolTip(txtFilenameArray[idx], myJsonData.HistoryList.Data[idx].Path);
                                index = myJsonData.HistoryList.Data[idx].Path.LastIndexOf("\\");
                                txtFilenameArray[idx].Text = myJsonData.HistoryList.Data[idx].Path.Substring(index + 1);
                                txtWindowTitleArray[idx].Text = myJsonData.HistoryList.Data[idx].WindowTitle;
                                chkSSLArray[idx].Checked = myJsonData.HistoryList.Data[idx].EnableSSL;
                            }
                        }
                        response = 1;
                    }
                }
                else
                {
                    response = 0;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Load JsonFile error: {0}", e.Message);
            }
            return response;
        }
      
        private void SaveJsonFile()
        {
            RootObject newRootObject = new RootObject();
            List<Data> listData = new List<Data>();

            for (int idx = 0; idx < MAXITEM; idx++)
            {
                Data newData = new Data
                {
                    IPAddress = txtIPArray[idx].Text,
                    IpId = txtIpIdArray[idx].Text,
                    RoomId = txtRoomIdArray[idx].Text.ToUpper(),
                    Path = toolTipArray[idx].GetToolTip(txtFilenameArray[idx]),
                    WindowTitle = txtWindowTitleArray[idx].Text,
                    EnableSSL = chkSSLArray[idx].Checked
                };
                listData.Add(newData);
            }
            newRootObject.HistoryList = new HistoryList();
            newRootObject.HistoryList.Data = listData;

            using (StreamWriter file = File.CreateText(JSONFILE))
            {
                JsonSerializer serializer = new JsonSerializer();
                //serialize object directly into file stream
                serializer.Serialize(file, newRootObject);
            }
        }
        
        private void ModifyEnvironmentXml(string Filename, string IPAddress, string IpID, string RoomID, string WindowTitle, bool EnableSSL)
        {
            using (FileStream zipToOpen = new FileStream(@Filename, FileMode.Open))
            {
                using (ZipArchive myArchive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                {
                    ZipArchiveEntry myEntry = myArchive.Entries.FirstOrDefault(entry => entry.Name == "Environment.xml");
                    var xmlDoc = new XmlDocument();
                    using (StreamReader reader = new StreamReader(myEntry.Open()))
                    {
                        xmlDoc.LoadXml(reader.ReadToEnd());
                        XmlNode xnodeTitle = xmlDoc.SelectSingleNode("Crestron/Properties/XPanel/WindowTitle");
                        xnodeTitle.InnerXml = "<![CDATA[" + WindowTitle + "]]>";

                        XmlNode xnodeHost = xmlDoc.SelectSingleNode("Crestron/Properties/CNXConnection/Host");
                        xnodeHost.InnerXml = "<![CDATA[" + IPAddress + "]]>";
                        
                        XmlNode xnodeIpid = xmlDoc.SelectSingleNode("Crestron/Properties/CNXConnection/IPID");
                        xnodeIpid.InnerXml = "<![CDATA[" + IpID + "]]>";

                        XmlNode xnodeRoomid = xmlDoc.SelectSingleNode("Crestron/Properties/CNXConnection/ProgramInstanceId");
                        if(xno != null)
                            xnodeRoomid.InnerXml = "<![CDATA[" + RoomID.ToUpper() + "]]>";

                        XmlNode xnodePort = xmlDoc.SelectSingleNode("Crestron/Properties/CNXConnection/Port");
                        if (EnableSSL)
                            xnodePort.InnerXml = "41796";
                        else
                            xnodePort.InnerXml = "";
                        XmlNode xnodeSSL = xmlDoc.SelectSingleNode("Crestron/Properties/CNXConnection/EnableSSL");
                        if (EnableSSL)
                            xnodeSSL.InnerXml = "true";
                        else
                            xnodeSSL.InnerXml = "false";
                    }
                    myArchive.Entries.FirstOrDefault(entry => entry.Name == "Environment.xml").Delete();
                    ZipArchiveEntry writeEntry = myArchive.CreateEntry("swf/Environment.xml");
                    using (StreamWriter writer = new StreamWriter(writeEntry.Open()))
                        xmlDoc.Save(writer);
                }
            }
        }

        public static void CopyAll(DirectoryInfo source, DirectoryInfo target)
        {
            if (source.FullName.ToLower() == target.FullName.ToLower())
                return;

            if (Directory.Exists(target.FullName) == false)
                Directory.CreateDirectory(target.FullName);

            // Copy each file into it's new directory.
            foreach (FileInfo fi in source.GetFiles())
                fi.CopyTo(Path.Combine(target.ToString(), fi.Name), true);

            // Copy each subdirectory using recursion.
            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {
                DirectoryInfo nextTargetSubDir =
                    target.CreateSubdirectory(diSourceSubDir.Name);
                CopyAll(diSourceSubDir, nextTargetSubDir);
            }
        }

        public void CreateFolders()
        {
            if (!Directory.Exists(sourceDirectory))
            {
                System.Windows.Forms.MessageBox.Show("Crestron Xpanel doesn't seem to be installed." + Environment.NewLine + "Program will stop", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(1);
            }
            else
            {
                if (!Directory.Exists(targetDirectory))
                {
                    DirectoryInfo diSource = new DirectoryInfo(sourceDirectory);
                    for (int x = 1; x <= MAXITEM; x++)
                    {
                        DirectoryInfo diTarget = new DirectoryInfo(targetDirectory + @"\CrestronXpanel-" + x.ToString());
                        CopyAll(diSource, diTarget);

                        //Modify xml file 
                        string path = targetDirectory + @"\CrestronXpanel-" + x.ToString() + @"\META-INF\AIR\application.xml";
                        var doc = XDocument.Load(path);
                        var elementsToUpdate = doc.Descendants();
                        //update elements value
                        foreach (XElement element in elementsToUpdate)
                        {
                            if (element.Name.LocalName == "id")
                                element.Value = "CrestronXpanel-" + x.ToString();
                        }
                        doc.Save(path);
                    }
                }
            }
        }
       
        public static string GetProcessOwner(int processId)
        {
            string query = "Select * From Win32_Process Where ProcessID = " + processId;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);
            ManagementObjectCollection processList = searcher.Get();

            foreach (ManagementObject obj in processList)
            {
                string[] argList = new string[] { string.Empty, string.Empty };
                int returnVal = Convert.ToInt32(obj.InvokeMethod("GetOwner", argList));
                if (returnVal == 0)
                {
                    return argList[1] + "\\" + argList[0];
                }
            }
            return "NO OWNER";
        }
        
        private void GetNextAvailableXpanel(string filename, string IPAddress, string IpID, string RoomID, string ShortName, bool EnableSSL)
        {
            bool foundSlot;
            ModifyEnvironmentXml(filename, IPAddress, IpID, RoomID, ShortName, EnableSSL);

            Process[] listXPanel = Process.GetProcessesByName("CrestronXPanel");
            if (listXPanel.Length == 0)
            {
                RunXPanel(1, filename);
            }
            else
            {
                List<int> listID = new List<int>();
                foreach (Process proc in listXPanel)
                {
                    if (GetProcessOwner(proc.Id) == Environment.UserDomainName + "\\" + Environment.UserName)
                    {
                        try
                        {
                            string sPath = proc.MainModule.FileName.ToString();
                            int sValue = sPath.Length - 20;
                            char cPath = sPath[sValue];

                            int iPath;
                            if (int.TryParse(cPath.ToString(), out iPath))
                                listID.Add(iPath);
                        }
                        catch (Exception)
                        {
                            //looks like it is an Xpanel instance initiated by VTPro.  Skip to the next one.
                            //No need to worry, as he doesn't use one of our available instance.
                        }
                    }
                }
                listID.Sort();
                foundSlot = false;
                for (int i = 2; i <= MAXITEM; i++)
                {
                    if (!listID.Contains(i))
                    {
                        RunXPanel(i, filename);
                        foundSlot = true;
                        break;
                    }
                }
                if (!foundSlot)
                {
                    MessageBox.Show("All executables are in use.  Please close one of them and retry.", Application.ProductName);
                }
            }
        }

        public void RunXPanel(int ID, string filename)
        {
            string sXPanelPath = (targetDirectory + @"\CrestronXpanel-" + ID.ToString() + "\\");
            Process p = new Process();
            p.StartInfo.FileName = sXPanelPath + "CrestronXPanel.exe";
            string ARG = "\"" + filename + "\"";
            p.StartInfo.Arguments = ARG;
            p.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
            p.Start();
        }
    }

}
