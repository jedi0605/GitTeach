using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1
{
    class Class1
    {
        using System.Linq;
using System.ServiceProcess;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Text;
using System.Timers;
using System.Data.Sql;
using System.Data.SqlClient;
using DataAccessLayer;
using System.Configuration;
using System.Net.Mail;
using System.Management;
using System.Management.Instrumentation;
using System.Collections;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Windows.Forms;
using Microsoft.VisualBasic.Devices;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions; // System.Collections 命名空間
using Cjwdev; 
using Cjwdev.WindowsApi;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using assign_ip.AutoProvision_WS;
using System.Runtime.Serialization.Json;
using System.Runtime.Serialization;
using Newtonsoft.Json.Linq;
using System.Net.NetworkInformation;
namespace assign_ip
{
    public partial class Service1 : ServiceBase
    {
        [DllImport("kernel32.dll")]
        static extern bool SetComputerName(string lpComputerName);
        [DllImport("kernel32.dll")]
        static extern bool SetComputerNameEx(_COMPUTER_NAME_FORMAT iType, string lpComputerName);
        enum _COMPUTER_NAME_FORMAT
        {
            ComputerNameNetBIOS,
            ComputerNameDnsHostname,
            ComputerNameDnsDomain,
            ComputerNameDnsFullyQualified,
            ComputerNamePhysicalNetBIOS,
            ComputerNamePhysicalDnsHostname,
            ComputerNamePhysicalDnsDomain,
            ComputerNamePhysicalDnsFullyQualified,
            ComputerNameMax
        };
        public Service1()
        {
            InitializeComponent();
        }


        class Ctrl_String
        {

            #region 字串分割

            public static string[] Split(string separator, string Text)
            {

                string[] sp = new string[] { separator };

                return Text.Split(sp, StringSplitOptions.RemoveEmptyEntries);

            }

            public static string[] Split(char separator, string Text)
            {

                char[] sp = new char[] { separator };

                return Text.Split(sp, StringSplitOptions.RemoveEmptyEntries);

            }

            #endregion

        }
        //        private static void VM_config_log(string vmname, string log_class ,string vm_level ,string info)
        //        {
        //            string sql;
        //            DBManager dbManager = new DBManager(DataProvider.SqlServer);
        //            dbManager.ConnectionString = ConfigurationManager.AppSettings["SSM"].ToString();
        //            dbManager.Open();
        //            dbManager.CreateParameters(4);
        //            dbManager.AddParameters(0, "@order_id", vmname);
        //            dbManager.AddParameters(1, "@log_class", log_class);
        //            dbManager.AddParameters(2, "@vm_level", vm_level);
        //            dbManager.AddParameters(3, "@info", info);
        //            sql = @"INSERT VM_config_log(udp_datetime , order_id , class , VM_level , info)
        //                    VALUES (GETDATE(),@order_id,@log_class,@vm_level,@info)";
        //            int write_log = dbManager.ExecuteNonQuery(CommandType.Text, sql);

        //        }
        private static string getIP_State(int IP_State)
        {
            string result;
            switch (IP_State)
            {
                case 0: result = "Successful completion, no reboot required."; break;

                case 1: result = "Successful completion, reboot required."; break;

                case 64: result = "Method not supported when the NIC is in DHCP mode."; break;

                case 65: result = "Unknown failure."; break;

                case 66: result = "Invalid subnet mask."; break;

                case 67: result = "An error occurred while processing an instance that was returned."; break;

                case 68: result = "Invalid input parameter."; break;

                case 69: result = "More than five gateways specified."; break;

                case 70: result = "Invalid IP address."; break;

                case 71: result = "Invalid gateway IP address."; break;

                case 72: result = "An error occurred while accessing the registry for the requested information."; break;

                case 73: result = "Invalid domain name."; break;

                case 74: result = "Invalid host name."; break;

                case 75: result = "No primary or secondary WINS server defined."; break;

                case 76: result = "Invalid file."; break;

                case 77: result = "Invalid system path."; break;

                case 78: result = "File copy failed."; break;

                case 79: result = "Invalid security parameter."; break;

                case 80: result = "Unable to configure TCP/IP service."; break;

                case 81: result = "Unable to configure DHCP service."; break;

                case 82: result = "Unable to renew DHCP lease."; break;

                case 83: result = "Unable to release DHCP lease."; break;

                case 84: result = "IP not enabled on adapter."; break;

                case 85: result = "IPX not enabled on adapter."; break;

                case 86: result = "Frame or network number bounds error."; break;

                case 87: result = "Invalid frame type."; break;

                case 88: result = "Invalid network number."; break;

                case 89: result = "Duplicate network number."; break;

                case 90: result = "Parameter out of bounds."; break;

                case 91: result = "Access denied."; break;

                case 92: result = "Out of memory."; break;

                case 93: result = "Already exists."; break;

                case 94: result = "Path, file, or object not found."; break;

                case 95: result = "Unable to notify service."; break;

                case 96: result = "Unable to notify DNS service."; break;

                case 97: result = "Interface not configurable."; break;

                case 98: result = "Not all DHCP leases can be released or renewed."; break;

                case 100: result = "DHCP not enabled on adapter."; break;

                default: result = ""; break;

            }
            return result;
        }  //回傳錯誤狀態
        //        private static void clean_IP(string vmname)
        //        {
        //            string sql;
        //            DBManager dbManager = new DBManager(DataProvider.SqlServer);
        //            dbManager.ConnectionString = ConfigurationManager.AppSettings["SSM"].ToString();
        //            dbManager.Open();
        //            dbManager.CreateParameters(1);
        //            dbManager.AddParameters(0, "@order_id", vmname);
        //            sql = @"update c_ip_list 
        //                    set ip_address=(select substring(ip_address,0,case charindex(':',ip_address) when ''  then Len(ip_address)+1 else charindex(':',ip_address) end) from c_ip_list a where c_ip_list.row_id=a.row_id),order_id=NULL,used=NULL,upd_userid='Service',used_mac=NULL,upd_datetime=GETDATE()
        //                    where order_id =@order_id";

        //            int updateOderId = dbManager.ExecuteNonQuery(CommandType.Text, sql);
        //        } //clean 該order所有分配的IP
        //                                                  "10.10.1.222"        "255.255.255.0"        " 10.10.1.10"      
        private static int SetIP(string mac, string newIPAddress, string newSubnetMask, string newGateway, string[] newDNS)
        {
            ManagementClass objMC = new ManagementClass("Win32_NetworkAdapterConfiguration");
            ManagementObjectCollection objMOC = objMC.GetInstances();
            int IP_State = 0;

            foreach (ManagementObject objMO in objMOC)
            {
                if (!(bool)objMO["IPEnabled"]) continue;

                try
                {

                    //Only change for device specified
                    if (string.Compare((string)objMO["MACAddress"], mac, true) == 0)  //網卡 == Web_servive傳進來的網卡 (忽略大小寫)
                    {
                        string result = "";
                        ManagementBaseObject objNewIP = null;
                        ManagementBaseObject objSetIP = null;
                        ManagementBaseObject objNewGate = null;
                        ManagementBaseObject objDNSServerSearchOrder = null;
                        objNewIP = objMO.GetMethodParameters("EnableStatic"); //新增IP的方法 包含兩個參數
                        objNewIP["IPAddress"] = new object[] { newIPAddress };//參數1
                        objNewIP["SubnetMask"] = new object[] { newSubnetMask };//參數2
                        objSetIP = objMO.InvokeMethod("EnableStatic", objNewIP, null); //將參數寫入方法
                        IP_State = IP_State + (Convert.ToInt32(objSetIP["returnValue"]));
                        result = getIP_State(IP_State);

                        objNewGate = objMO.GetMethodParameters("SetGateways");//設定GATEWAY
                        objNewGate["DefaultIPGateway"] = new object[] { newGateway };//參數1
                        objSetIP = objMO.InvokeMethod("SetGateways", objNewGate, null);//將參數寫入方法
                        IP_State = IP_State + (Convert.ToInt32(objSetIP["returnValue"]));

                        objDNSServerSearchOrder = objMO.GetMethodParameters("SetDNSServerSearchOrder");//設定DNSServerSearchOrder
                        objDNSServerSearchOrder["DNSServerSearchOrder"] = newDNS;//參數
                        objDNSServerSearchOrder = objMO.InvokeMethod("SetDNSServerSearchOrder", objDNSServerSearchOrder, null);
                        IP_State = IP_State + (Convert.ToInt32(objSetIP["returnValue"]));

                        //System.IO.File.WriteAllText(@"C:\AutoProvision\ip_state.txt", "ip state : " + result);
                        System.IO.File.AppendAllText(@"C:\AutoProvision\ip_state.txt", "ip state : " + result + "\n");
                        Console.WriteLine("ip state : " + result);
                        System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "ip state : " + result + Environment.NewLine);
                    }


                }

                catch (Exception ex)
                {
                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "Exception setting IP" + Environment.NewLine);

                    Console.WriteLine("Exception setting IP: " + ex.Message);
                }
            }
            return IP_State;
        }

        public static bool change_computername(string vmname)
        {
            string MachineName = vmname;

            bool succeeded = SetComputerName(MachineName);
            System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "SetComputerName :" + (succeeded?"true":"false") + Environment.NewLine);
            bool succeeded2 = SetComputerNameEx(_COMPUTER_NAME_FORMAT.ComputerNamePhysicalDnsHostname, MachineName);
            System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "SetComputerNameEx :" + (succeeded2 ? "true" : "false") + Environment.NewLine);
            if (succeeded && succeeded2)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        //public static String rdp_port_radom(int digitOfSet, int set)
        //{
        //    String newPwd = "";
        //    String pwdSet = "";
        //    System.Random rand = new System.Random();	// for C#
        //    int num = 0;
        //    for (int i = 0; i < set; i++)
        //    {
        //        if (i > 0)
        //            newPwd += "-";	//各組英數之間的分隔號
        //        pwdSet = "";
        //        while (pwdSet.Length < digitOfSet)
        //        {
        //            // Java 中 pwdSet.Length 改為 pwdSet.length()
        //            num = rand.Next(48, 57);	// for C#

        //            pwdSet += (char)num;	//將數字轉換為字元
        //        }
        //        newPwd += pwdSet;
        //    }
        //    return newPwd;
        //}

        public static String pwdGenerator(int digitOfSet, int set)
        {
            String newPwd = "";
            String pwdSet = "";
            System.Random rand = new System.Random();	// for C#
            int num = 0;
            for (int i = 0; i < set; i++)
            {
                if (i > 0)
                    newPwd += "-";	//各組英數之間的分隔號
                pwdSet = "";
                while (pwdSet.Length < digitOfSet)
                {
                    // Java 中 pwdSet.Length 改為 pwdSet.length()
                    num = rand.Next(50, 90);	// for C#
                    if (num > 57 && num < 65)
                        continue;	//排除 58~64 這區間的非英數符號
                    else if (num == 79 || num == 73)
                        continue;	//排除 I 和 O
                    pwdSet += (char)num;	//將數字轉換為字元
                }
                newPwd += pwdSet;
            }
            return newPwd;
        }
        /// <summary>
        /// 抓取本機所有的MAC 回傳到陣列上
        /// </summary>
        /// <returns> mac <LIST></returns>
        public static List<string> getLocalMac()
        {
            PhysicalAddress mac;
            List<string> mac_name = new List<string>(); //save mac name
            NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces(); // get all network interface 
            foreach (NetworkInterface adapter in nics)
            {
                if (adapter.NetworkInterfaceType.ToString().Equals("Ethernet"))
                {
                    //取得IPInterfaceProperties(可提供網路介面相關資訊)
                    IPInterfaceProperties ipProperties = adapter.GetIPProperties();
                    if (ipProperties.UnicastAddresses.Count > 0)
                    {
                        mac = adapter.GetPhysicalAddress();
                        //取得Mac Address
                        string t = mac.ToString();
                        for (int i = 10; i > 0; i = i - 2)
                        {
                            t = t.Insert(i, ":");
                        }


                        string name = adapter.Name;                                             //網路介面名稱
                        string description = adapter.Description;                               //網路介面描述
                        string ip = ipProperties.UnicastAddresses[0].Address.ToString();        //取得IP
                        string netmask = ipProperties.UnicastAddresses[0].IPv4Mask.ToString();  //取得遮罩
                        //Console.WriteLine(mac);
                        mac_name.Add(t);
                    }
                }

            }
            System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " we get local mac"+mac_name + Environment.NewLine);

            return mac_name;
        }
        public static Dictionary<string, string> JsonParser(string Jedi_IP)
        {

            string[] jsonParser = Jedi_IP.Remove(Jedi_IP.Length - 2).Remove(0, 2).Split(',');
            Dictionary<string, string> dic = new Dictionary<string, string>();
            foreach (string jsonSession in jsonParser)
            {
                dic.Add(jsonSession.Split(':')[0].ToLower().Replace("\"", ""), jsonSession.Split(':')[1].Replace("\"", ""));
            }
            return dic;
        }
        public static void Do_Power_Shell(string script)
        {
            Runspace Runspace_do_power_shell = RunspaceFactory.CreateRunspace();
            Runspace_do_power_shell.Open();
            Pipeline do_power_shell_pipe = Runspace_do_power_shell.CreatePipeline();
            do_power_shell_pipe.Commands.AddScript(script);
            do_power_shell_pipe.Invoke();
            do_power_shell_pipe.Dispose();
        }

        private string checkFile_ip_reboot() {
            try
            {
                bool result = System.IO.File.Exists("c:\\AutoProvision\\check_ip_reboot.txt");
                System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "checkFile_ip_reboot : " + (result ? "true" : "false") + Environment.NewLine);
                if (result)
                    return "True";
                return "False";
                //string check_ip_reboot = @"Test-Path c:\\AutoProvision\\check_ip_reboot.txt";
                //Runspace check_ip_runspace = RunspaceFactory.CreateRunspace();
                //check_ip_runspace.Open();
                //Pipeline check_ip = check_ip_runspace.CreatePipeline();
                //check_ip.Commands.AddScript(check_ip_reboot);
                //check_ip.Commands.Add("Out-String");
                //var check_ip2 = check_ip.Invoke();
                //string[] check_ip_output0 = check_ip2[0].ToString().Split('\r');
                //string check_ip_output = check_ip_output0[0];
                //check_ip.Dispose();
                //check_ip_runspace.Close();
                //System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "checkFile_ip_reboot : " + check_ip_output + Environment.NewLine);
                //return check_ip_output;
            }
            catch (Exception ex)
            {
                System.IO.File.AppendAllText(@"C:\AutoProvision\error.txt", "checkFile_ip_reboot have ex" + ex.Message + Environment.NewLine);
                return "";
            }
        }
        private string checkFile_step1()
        {
            try
            {
                bool result = System.IO.File.Exists("c:\\AutoProvision\\step1.txt");
                System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "checkFile_step1 : " + (result ? "true" : "false") + Environment.NewLine);
                if (result)
                    return "True";
                return "False";
                //string step_1 = @"Test-Path c:\\AutoProvision\\step1.txt";
                //Runspace step_1_runspace = RunspaceFactory.CreateRunspace();
                //step_1_runspace.Open();
                //Pipeline step1_pipe = step_1_runspace.CreatePipeline();
                //step1_pipe.Commands.AddScript(step_1);
                //step1_pipe.Commands.Add("Out-String");
                //var step1_pipe2 = step1_pipe.Invoke();
                //string[] step1_pipe_output0 = step1_pipe2[0].ToString().Split('\r');
                //string step1_output = step1_pipe_output0[0];
                //step1_pipe.Dispose();
                //step_1_runspace.Close();
                //System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "checkFile_step1 : " + step1_output + Environment.NewLine);
                //return step1_output;
            }
            catch (Exception ex)
            {
                System.IO.File.AppendAllText(@"C:\AutoProvision\error.txt", "checkFile_step1 have ex" + ex.Message + Environment.NewLine);
                return "";
            }
        }
        protected override void OnStart(string[] args)
        {
            string vmname = "";
            string group_id="";
            AutoProvision_WS.AutoProvision_WS ws = new AutoProvision_WS.AutoProvision_WS();
            List<string> mac = getLocalMac();
            //////檢查 兩個檔案是否存在 利用powershell
            string check_ip_output = checkFile_ip_reboot();
            if (check_ip_output == "")
                return;
            string step1_output = checkFile_step1();
            if (step1_output == "")
                return;
            string line = "";
            try
            {
                System.IO.StreamReader vmname_load = new System.IO.StreamReader(@"c:\AutoProvision\vmname.txt");
                string[] get_orderID_groupID = vmname_load.ReadToEnd().Split(' ');
                vmname = get_orderID_groupID[0];
                group_id = get_orderID_groupID[1];
                System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "Read vmname.txt vmname:" + vmname +",group_id:"+group_id+ Environment.NewLine);
            }
            catch (Exception ex)
            {
                System.IO.File.AppendAllText(@"C:\AutoProvision\error.txt", "Read vmname.txt have ex" + ex.Message + Environment.NewLine);
                return;
            }
            /////end

            ///兩者皆否  撰寫完成
            if (step1_output == "False" && check_ip_output == "False")
            {
                //AutoProvision_WS.AutoProvision_WS ws = new AutoProvision_WS.AutoProvision_WS();
                try
                {
                    ws.Inset_Percent(vmname, "60", ""); //寫入進度60%
                    //string[] info_result = ws.Get_Order_Info(vmname).Split('"');
                    //string company_id = info_result[3];
                    //string area = info_result[7];
                    //string member_id = info_result[11];

                    string info_result = ws.Get_Order_Info(vmname);

                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "Get_Order_Info info_result:" + info_result+ Environment.NewLine);

                    JToken info = JObject.Parse(info_result);
                    string company_id = info["company_id"].ToString();
                    string area = info["area"].ToString();
                    string member_id = info["member_id"].ToString();
                    string Get_ComputerName_result = ws.Get_ComputerName(vmname);
                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "Get_ComputerName static_ip_result:" + Get_ComputerName_result + Environment.NewLine);
                    string[] static_ip_result = Get_ComputerName_result.Split('"');
                    
                    string FQDN = static_ip_result[5];
                    change_computername(FQDN);

                    IntPtr userTokenHandle = IntPtr.Zero;
                    ApiDefinitions.WTSQueryUserToken(ApiDefinitions.WTSGetActiveConsoleSessionId(), ref userTokenHandle);
                    ApiDefinitions.PROCESS_INFORMATION procInfo = new ApiDefinitions.PROCESS_INFORMATION();
                    ApiDefinitions.STARTUPINFO startInfo = new ApiDefinitions.STARTUPINFO();
                    startInfo.cb = (uint)Marshal.SizeOf(startInfo);
                    string restart = "restart-computer -force";
                    System.IO.File.WriteAllText(@"C:\AutoProvision\del_item_sc2.ps1", restart);
                    System.IO.File.WriteAllText(@"C:\AutoProvision\step1.txt", "");
                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "CreateFile step1.txt" + Environment.NewLine);
                    Process p = new Process();
                    p.StartInfo.FileName = @"C:\AutoProvision\creboot2.exe";
                    p.StartInfo.UseShellExecute = true;
                    p.Start();

                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "Run creboot.exe" + Environment.NewLine);
                    p.WaitForExit();

                    System.Threading.Thread.Sleep(2000);

                    RunspaceInvoke invoker = new RunspaceInvoke();
                    invoker.Invoke("restart-computer -force");
                    string RebootPath = @"C:\AutoProvision\creboot2.exe";

                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "Reboot.exe" + Environment.NewLine);
                    ApiDefinitions.CreateProcessAsUser(userTokenHandle,
                        RebootPath,
                        "",
                        IntPtr.Zero,
                        IntPtr.Zero,
                        false,
                        0,
                        IntPtr.Zero,
                        null,
                        ref startInfo,
                        out procInfo);
                    
                }
                catch (Exception ex)
                {
                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "set computer name have ex" + ex.Message + Environment.NewLine);

                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "Line=" + line + "*********" + ex.Message);
                }
                finally
                {
                    ws.Dispose();
                }
            }
            /////end

            ///// 修改 ip domain rdp  未完成
            if (step1_output == "True" && check_ip_output == "False")
            {
                System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " Set computer name ok" + Environment.NewLine);

                //string v = ws.test();
                try
                {

                    RunspaceInvoke invoker = new RunspaceInvoke();
                    invoker.Invoke("Set-ExecutionPolicy Unrestricted");
                    invoker.Dispose();
                    //公司ID及地區
                    string ip = "";
                    string netmask = "";
                    string d_gateway = "";
                    string d_dns = "";
                    string o_dns = "";
                    string rdp_port = "";
                    string company_id = "";
                    string area = "";
                    List<string> macaddress = getLocalMac();  //temp remark
                    List<JToken> IP_detail = new List<JToken>();
                    string upwd = "!QAZ2wsx"; // pwdGenerator(12, 1);
                    string lines2 = "net user User " + upwd + "";
                    string domain_ip = "";
                    string domain_pwd = "";
                    string domain_account = "";
                    string domain_name = "";
                    string member_id = "";

                    Runspace runspace4 = RunspaceFactory.CreateRunspace();
                    runspace4.Open();
                    Pipeline cpassword = runspace4.CreatePipeline();
                    cpassword.Commands.AddScript(lines2);
                    cpassword.Commands.Add("Out-String");
                    var cpassword2 = cpassword.Invoke();
                    string cpassword2output = cpassword2[0].ToString();
                    cpassword.Dispose();
                    runspace4.Dispose();

                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " cpassword2 : " + cpassword2output + Environment.NewLine);
                    
                    string info_result = ws.Get_Order_Info(vmname);

                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " Get_Order_Info : " + info_result + Environment.NewLine);
                    
                    JToken info = JObject.Parse(info_result);
                    company_id = info["company_id"].ToString();
                    area = info["area"].ToString();
                    member_id = info["member_id"].ToString();

                    //company_id = info_result[3];
                    //area = info_result[7];
                    //member_id = info_result[11];


                    ///////////////////////////////  設定 IP   START

                    string Jedi_sql;
                    string temp_vlan = "";
                    int chech_adapter_num = 0;

                    //                    Dictionary<string, string> dic = new Dictionary<string, string>();
                    //                    DBManager dbManager = new DBManager(DataProvider.SqlServer);
                    //                    dbManager.ConnectionString = ConfigurationManager.AppSettings["SSM"].ToString();
                    //                    dbManager.Open();
                    //                    dbManager.CreateParameters(1);
                    //                    dbManager.AddParameters(0, "@order_id", vmname);
                    //                    Jedi_sql = @"SELECT COUNT(order_id)  
                    //                                from order_nic_mac_list
                    //                                where order_id=@order_id";
                    //                    chech_adapter_num = System.Convert.ToInt32(dbManager.ExecuteScalar(CommandType.Text, Jedi_sql));
                    //                    if (macaddress.Count != chech_adapter_num)  //檢查網卡數量 是否等於 DB數量
                    //                    {
                    //                        ws.Inset_VM_config_log(vmname, "SET IP", "ERROR", "Network adapter have some problams."); 
                    //                        //VM_config_log(vmname, "SET IP", "ERROR", "Network adapter have some problams.");
                    //                        System.IO.File.WriteAllText(@"C:\AutoProvision\logs.txt", "adapter number != order adapter number");
                    //                        return;
                    //                    }
                    try
                    {
                        //get all network configuration
                        string[] dns_t;
                        int ip_result = 0;
                        for (int i = 0; i < macaddress.Count; i++) // get vlanID
                        {

                            //                            dbManager.CreateParameters(2);
                            //                            dbManager.AddParameters(0, "@order_id", vmname);
                            //                            dbManager.AddParameters(1, "@nic_mac", macaddress[i]);
                            //                            Jedi_sql = @"select vlan_id
                            //                                from order_nic_mac_list
                            //                                where order_id=@order_id and nic_mac=@nic_mac";
                            //                            temp_vlan = System.Convert.ToString(dbManager.ExecuteScalar(CommandType.Text, Jedi_sql));
                            temp_vlan = ws.Get_VLAN_ID_Info(vmname, macaddress[i]);


                            System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " Get_VLAN_ID_Info : " + temp_vlan + Environment.NewLine);

                            JToken temp_parser = JObject.Parse(temp_vlan);

                            string Jedi_IP = ws.Assign_Network_Configuration(macaddress[i], group_id.Replace("\r\n", ""), vmname, area, company_id, temp_parser["vlan_id"].ToString());
                            System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " Assign_Network_Configuration: " + Jedi_IP + Environment.NewLine);

                            if (Jedi_IP == "no free ip")
                            {
                                ws.Insert_VM_config_log(vmname, "Static_ip", "error", Jedi_IP);
                                System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " Insert_VM_config_log"+ Environment.NewLine);
                                //VM_config_log(vmname, "Static_ip", "error", Jedi_IP);
                                return;
                            }
                            JToken token;
                            token = JObject.Parse(Jedi_IP.Remove(Jedi_IP.Length - 1).Remove(0, 1).Replace("\\", "\\\\"));
                            IP_detail.Add(token);
                        }
                        //for (int i = 0; i < macaddress.Count; i++) //偵測分配IP錯誤 如有錯誤 將以分配的IP ORDER初始化
                        //{
                        //    if (((Dictionary<string, string>)h[i])["ip"] == "" || h.Count != macaddress.Count)
                        //    {
                        //        System.IO.File.WriteAllText(@"C:\AutoProvision\IP_logs.txt", "static_IP web_service have some problam.");
                        //        ws.Inset_Percent(vmname, "69", "static_IP web_service have some problam.");
                        //    }
                        //    //clean_IP(vmname);  暫時關閉
                        //    //return;
                        //}
                        for (int i = 0; i < macaddress.Count; i++)
                        {
                            //if (i > 1) { break; } // debug用
                            List<string> dns = new List<string>();
                            dns.Add(IP_detail[i].SelectToken("d_dns").ToString());
                            dns.Add(IP_detail[i].SelectToken("o_dns").ToString());
                            dns_t = dns.ToArray();

                            ip_result = ip_result + SetIP(IP_detail[i].SelectToken("used_mac").ToString(), IP_detail[i].SelectToken("ip").ToString(), IP_detail[i].SelectToken("netmask").ToString(), IP_detail[i].SelectToken("d_gateway").ToString(), dns_t);
                        }
                        if (ip_result == 0) //如果更改都成功，更改DB的IP使用狀態
                        {
                            System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " Ip set is ok! " + Environment.NewLine);

                            string change_nim = ws.Change_IP_Status(vmname);
                        }
                        else
                        {
                            System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " Ip set is ERROR " + Environment.NewLine);

                            ws.Insert_VM_config_log(vmname, "SET IP", "ERROR", "Network configuration fail.");
                            //VM_config_log(vmname, "SET IP", "ERROR", "Network configuration fail.");

                            return;
                        }
                        /////////////////////////////////////設定 IP結束
                    }
                    catch (Exception ex)
                    {
                        System.IO.File.AppendAllText(@"C:\AutoProvision\error.txt", " Ip set is ERROR " +ex.Message+ Environment.NewLine);
                    }
                    finally
                    {
                        //dbManager.Dispose();
                    }
                    ////////////////////////////////
                    ws.Inset_Percent(vmname, "70", "");
                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " Inset_Percent" + Environment.NewLine);
                    //                    string sql = @"select vlan_id
                    //                                from user_vm_order
                    //                                where order_id=@order_id";
                    //                    DataSet ds = dbManager.ExecuteDataSet(CommandType.Text, sql);
                    //                    int nic_num = ds.Tables[0].Rows[0]["vlan_id"].ToString().Split(',').Count();
                    //for( int i = 0 ; i < nic_num ; i++ )
                    //ws.Set_VM_pwd(vmname, upwd);
                    string Assign_Network_Configuration = ws.Assign_Network_Configuration(macaddress[0], group_id, vmname, area, company_id, temp_vlan);
                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " Assign_Network_Configuration" + Assign_Network_Configuration + Environment.NewLine);
                    string[] static_ip_result = Assign_Network_Configuration.Split('"');
                    //ip = static_ip_result[5];
                    //netmask = static_ip_result[9];
                    //d_gateway = static_ip_result[13];
                    //d_dns = static_ip_result[17];
                    //o_dns = static_ip_result[21];
                    rdp_port = static_ip_result[25];
                    domain_name = static_ip_result[29];
                    domain_ip = static_ip_result[33];
                    domain_account = static_ip_result[37];
                    domain_pwd = CryptoAES.decrypt(static_ip_result[41], "GccA@stanchengGg");
                    if (rdp_port != "3389")
                    {
                        System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "Ready Change RDP Port" + Environment.NewLine);
                        string rdp_port_ps1 = "Set-ItemProperty -path 'HKLM:\\System\\CurrentControlSet\\Control\\Terminal Server\\WinStations\\RDP-Tcp' -name PortNumber -value " + rdp_port;
                        Runspace runspace_rdp_port = RunspaceFactory.CreateRunspace();
                        runspace_rdp_port.Open();
                        Pipeline rdp_port_pipe = runspace_rdp_port.CreatePipeline();
                        rdp_port_pipe.Commands.AddScript(rdp_port_ps1);
                        rdp_port_pipe.Commands.Add("Out-String");
                        var rdp_port_pipe2 = rdp_port_pipe.Invoke();
                        string rdp_port_out = rdp_port_pipe2[0].ToString();
                        runspace_rdp_port.Dispose();
                        System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " RDP PORD have change! " + Environment.NewLine);

                    }
                    ws.Inset_Percent(vmname, "80", "");
                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " Inset_Percent" + Environment.NewLine);
                    System.Threading.Thread.Sleep(1000);
                    //IP-Setting & Join Domain
                    //@"$NICs = Get-WMIObject Win32_NetworkAdapterConfiguration | where{$_.IPEnabled -eq ""TRUE""}" + "\n" +
                    //            "$NIC=\"Lan\"\n" +
                    //            "Foreach($NIC in $NICs) {\n" +
                    //            "$NIC.EnableStatic(\"" + ip + "\", \"" + netmask + "\")\n" +
                    //            "$NIC.SetGateways(\"" + d_gateway + "\")\n" +
                    //            "$DNSServers = \"" + d_dns + "\",\"" + o_dns + "\"\n" +
                    //            "$NIC.SetDNSServerSearchOrder($DNSServers)\n" +
                    //            "$NIC.SetDynamicDNSRegistration(\"TRUE\")\n" +
                    //            "} \n" +
                    string ip_set =
                                "Start-Sleep -s 20 \n" +
                                "$domain = " + "\"" + domain_name + "\"" + "\n" +
                                "$password = " + "\"" + domain_pwd + "\"" + " | ConvertTo-SecureString -asPlainText -Force\n" +
                                "$username = " + "\"" + domain_account + "\"" + "\n" +
                                "$credential = New-Object System.Management.Automation.PSCredential($username,$password)\n" +
                                "Add-Computer -DomainName $domain -Cred $credential\n" +
                                "Start-Sleep -s 20\n" +
                                "restart-computer -force\n";
                    System.IO.File.WriteAllText(@"C:\AutoProvision\set_ip.ps1", ip_set);
                    System.IO.File.WriteAllText(@"C:\AutoProvision\check_ip_reboot.txt", domain_name + " " + domain_pwd + " " + domain_account + " " + ip + " " + member_id + " " + vmname);
                    ws.Inset_Percent(vmname, "85", "");

                    IntPtr userTokenHandle = IntPtr.Zero;
                    ApiDefinitions.WTSQueryUserToken(ApiDefinitions.WTSGetActiveConsoleSessionId(), ref userTokenHandle);
                    ApiDefinitions.PROCESS_INFORMATION procInfo = new ApiDefinitions.PROCESS_INFORMATION();
                    ApiDefinitions.STARTUPINFO startInfo = new ApiDefinitions.STARTUPINFO();
                    startInfo.cb = (uint)Marshal.SizeOf(startInfo);
                    string RebootPath = @"C:\AutoProvision\set_ip.exe";
                    ApiDefinitions.CreateProcessAsUser(userTokenHandle,
                        RebootPath,
                        "",
                        IntPtr.Zero,
                        IntPtr.Zero,
                        false,
                        0,
                        IntPtr.Zero,
                        null,
                        ref startInfo,
                        out procInfo);
                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " Add to domain! " + Environment.NewLine);

                }
                catch (Exception ex)
                {
                    //VM_config_log(vmname, "", "ERROR", "Network adapter have some problams.");
                    System.IO.File.AppendAllText(@"C:\AutoProvision\error.txt", ex.Message);
                }
                finally
                {
                    ws.Dispose();
                }
            }
            if (step1_output == "True" && check_ip_output == "True")
            {
                //AutoProvision_WS.AutoProvision_WS ws = new AutoProvision_WS.AutoProvision_WS();
                try
                {
                    System.IO.StreamReader check_ip_reboot_f = new System.IO.StreamReader(@"c:\AutoProvision\check_ip_reboot.txt");
                    string[] domain_name0 = check_ip_reboot_f.ReadToEnd().Split(' ');
                    string domain_name = domain_name0[0];
                    string ip = domain_name0[3];
                    string member_id = domain_name + "/" + domain_name0[4];
                    vmname = domain_name0[5];
                    check_ip_reboot_f.Dispose();
                    string join_account = "$user=[ADSI]\"WinNT://" + member_id + "\"\n" +
                                         "$group=[ADSI]\"WinNT://./Remote Desktop Users\"\n" +
                                         "$group.Psbase.Invoke(\"Add\",$user.Psbase.path)";
                    //Runspace Runspace_join_account = RunspaceFactory.CreateRunspace();
                    //Runspace_join_account.Open();
                    //Pipeline join_account_pipe = Runspace_join_account.CreatePipeline();
                    //join_account_pipe.Commands.AddScript(join_account);
                    //join_account_pipe.Invoke();
                    //Runspace_join_account.Dispose();
                    Do_Power_Shell(join_account);
                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " join account  " + member_id + Environment.NewLine);

                    ws.Inset_Percent(vmname, "90", "");


                    string remove_DomainAdminsAccount = "$user=[ADSI]\"WinNT://Domain Admins" + "\"\n" +
                                                        "$group=[ADSI]\"WinNT://./Administrators\"\n" +
                                                        "$group.Psbase.Invoke(\"Remove\",$user.Psbase.path)";
                    Do_Power_Shell(remove_DomainAdminsAccount);

                    string remove_LocalUserAccount = "$ComputerName = $env:COMPUTERNAME" + "\n" +
                                                     "[ADSI]$server=\"WinNT://$ComputerName\"" + "\n" +
                                                     "$removeName=\"user\"" + "\n" +
                                                     "$server.Delete(\"user\",$removeName)";
                    Do_Power_Shell(remove_LocalUserAccount);

                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " remove account !" + Environment.NewLine);

                    //Runspace Runspace_remove_account = RunspaceFactory.CreateRunspace();
                    //Runspace_remove_account.Open();
                    //Pipeline remove_account_pipe = Runspace_remove_account.CreatePipeline();
                    //remove_account_pipe.Commands.AddScript(remove_DomainAdminsAccount);
                    //remove_account_pipe.Invoke();
                    ////remove_account_pipe.Commands.AddScript(remove_LocalUserAccount);
                    ////remove_account_pipe.Invoke();
                    //Runspace_remove_account.Dispose();

                    ws.Inset_Percent(vmname, "95", "");

                    string app_id = @"Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like ""VMconfig""}|Format-list IdentifyingNumber >c:\\AutoProvision\\app_guid.txt";
                    Runspace Runspace_app_guid = RunspaceFactory.CreateRunspace();
                    Runspace_app_guid.Open();
                    Pipeline app_guid_pipe = Runspace_app_guid.CreatePipeline();
                    app_guid_pipe.Commands.AddScript(app_id);
                    var app_guid_pipe2 = app_guid_pipe.Invoke();
                    Runspace_app_guid.Dispose();
                    System.IO.StreamReader app_guid_pipe_out = new System.IO.StreamReader(@"c:\AutoProvision\app_guid.txt");
                    string[] app = app_guid_pipe_out.ReadToEnd().Split(':');
                    string[] app2 = app[1].Split('\r');
                    string[] app3 = app2[0].Split(' ');
                    string app_guid = app3[1];
                    app_guid_pipe_out.Dispose();


                    string remove_it = @"Remove-Item c:\\AutoProvision -Recurse";
                    Runspace Runspace_remove = RunspaceFactory.CreateRunspace();
                    Runspace_remove.Open();
                    Pipeline remove_it_pipe = Runspace_remove.CreatePipeline();
                    remove_it_pipe.Commands.AddScript(remove_it);
                    var remove_it_pipe2 = remove_it_pipe.Invoke();
                    Runspace_remove.Dispose();

                    ws.Inset_Percent(vmname, "100", "");


                    ws.Change_Order_Status(vmname, "5", false);
                    string del_item_sc = @"MsiExec.exe /norestart /q/x""" + app_guid + "\" REMOVE=ALL";
                    Runspace Runspace_del_item = RunspaceFactory.CreateRunspace();
                    Runspace_del_item.Open();
                    Pipeline del_item_pipe = Runspace_del_item.CreatePipeline();
                    del_item_pipe.Commands.AddScript(del_item_sc);
                    var del_item_pipe2 = del_item_pipe.Invoke();
                    Runspace_del_item.Dispose();
                }
                catch (Exception ex)
                {
                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", ex.ToString());
                }
                finally
                {
                    ws.Dispose();
                }
            }
            
        
        }
        protected override void OnStop()
        {

        }
    }
}


    }
}
