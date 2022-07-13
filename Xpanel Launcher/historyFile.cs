using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Xpanel_Launcher
{
    public class Data
    {
        public string IPAddress { get; set; }
        public string IpId { get; set; }
        public string RoomId { get; set; }
        public string Path { get; set; }
        public string WindowTitle { get; set; }
        public bool EnableSSL { get; set; }
    }

    public class HistoryList
    {
        public List<Data> Data { get; set; }
    }

    public class RootObject
    {
        public HistoryList HistoryList { get; set; }
    }
}