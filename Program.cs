using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Timers;
using System.Data.Sql;
using System.Data.SqlClient;
using DataAccessLayer;
using System.Configuration;
using System.Net.Mail;
using Create_VM_Service;
using System.DirectoryServices;
using System.Management;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text.RegularExpressions;
using System.Threading;
using Create_VM_Service.VMAPI;
using CPU_Memory_Usage_API;
using GCCA.AutoProv.SSHClient;
using System.Xml;
using Newtonsoft.Json.Linq;
using Create_VM_Service.AutoProvision_WS;


namespace GCCA.AutoProv.Agent
{
    static class CreateVMServiceTest
    {

        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        /// 
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            //ServicesToRun = new ServiceBase[] 
            //{ 
            //    new Service1() 
            //};
            //ServiceBase.Run(ServicesToRun);
            de_bug();
            //a();
        }
        static void a() //test area
        {
            DBManager dbManager = new DBManager(DataProvider.SqlServer);
            dbManager.ConnectionString = ConfigurationManager.AppSettings["SSM"].ToString();

            dbManager.Open();
            dbManager.CreateParameters(4);
            dbManager.AddParameters(0, "@order_area", "TP");
            dbManager.AddParameters(1, "@company_id", "GCCA");
            dbManager.AddParameters(2, "@vmtype", "VMware");
            dbManager.AddParameters(3, "@temp_id", "Temp_20130606_001");
            string sql = @"SELECT vmware_apiurl,vmware_datacenter_name,vmware_datastore_name,vmware_host_account,vmware_host_name,vmware_host_pwd,b.hostname as create_on_hostname,resource_pool_name,b.datacenter_name as create_on_datacentername,b.temp_id
                    FROM config_vm_host a left outer join vm_temp b on a.vmtype=b.vm_type 
                    WHERE area=@order_area and a.company_id=@company_id and vmtype=@vmtype and b.temp_id=@temp_id";
            DataSet ds = dbManager.ExecuteDataSet(CommandType.Text, sql);
            int nCount = ds.Tables[0].Rows.Count;
            string vmware_apiurl = ds.Tables[0].Rows[0]["vmware_apiurl"].ToString();
            string vmware_vcenter_ip = vmware_apiurl.Split('/')[2];
            //vmware_datacenter_name = ds.Tables[0].Rows[0]["vmware_datacenter_name"].ToString();
            //vmware_datastore_name = ds.Tables[0].Rows[0]["vmware_datastore_name"].ToString();
            //vmware_host_account = ds.Tables[0].Rows[0]["vmware_host_account"].ToString();
            //vmware_host_name = ds.Tables[0].Rows[0]["vmware_host_name"].ToString();
            //vmware_host_pwd = CryptoAES.decrypt(ds.Tables[0].Rows[0]["vmware_host_pwd"].ToString(), "GccA@stanchengGg");
            //vmware_host_encryp_pwd = ds.Tables[0].Rows[0]["vmware_host_pwd"].ToString();
            //create_on_hostname = ds.Tables[0].Rows[0]["create_on_hostname"].ToString();
            //create_on_datacentername = ds.Tables[0].Rows[0]["create_on_datacentername"].ToString();
            //resource_pool_name = Convert.ToString(ds.Tables[0].Rows[0]["resource_pool_name"]);
            dbManager.Dispose();
            Create_VM_Service.VMAPI.GCCA_HypervisorAPI vmapi = new GCCA_HypervisorAPI();
            //string[] guidf0 = vmapi.init(vmware_apiurl, "administrator", "Passw0rd", "GCCA_8F_Lab", true).Split(':'); //連線至VMWARE HOST
            //string[] guidf1 = guidf0[2].Split('\'');
            //string guid = guidf1[1];//取得GUID
            string vm_network_name = vmapi.getVMNICNetworkList("f048f491-05c1-44aa-9aff-55e4129b4c1d", "GCCA_8F_Lab", "172.16.10.45", "20130606_006");
            return;
        }
        static void de_bug()
        {
            DBManager dbManager = new DBManager(DataProvider.SqlServer);
            Console.WriteLine("debug Start");
            string AutoProvision_WS_url = Create_VM_Service.Properties.Settings.Default.Create_VM_Service_AutoProvision_WS_AutoProvision_WS;

            string IsDelImg = ConfigurationManager.AppSettings["IsDelImg"].ToString();//是否刪除VM
            string db_server = ConfigurationManager.AppSettings["SSM"].ToString();
            string EmailAccount = ConfigurationManager.AppSettings["EmailAccount"].ToString();
            string EmailPassword = ConfigurationManager.AppSettings["EmailPassword"].ToString();
            string smtphost = ConfigurationManager.AppSettings["smtphost"].ToString();
            dbManager.ConnectionString = ConfigurationManager.AppSettings["SSM"].ToString();
            Create_VM_Service.AutoProvision_WS.AutoProvision_WS ws = new Create_VM_Service.AutoProvision_WS.AutoProvision_WS();
            //CPU_Memory_Usage_API.David_API vm_useage = new CPU_Memory_Usage_API.David_API();
            string ftp = ConfigurationManager.AppSettings["FTP_IP"].Replace("ftp://", "").Trim();
            string ftp_user = ConfigurationManager.AppSettings["ftpUsername"].Trim().ToString();
            string ftp_pwd = ConfigurationManager.AppSettings["ftpPassword"].Trim().ToString();
            string ftp_folder = ConfigurationManager.AppSettings["FTP_IP"] + "/" + ConfigurationManager.AppSettings["agentFtpPath"];
            int Max_Create_VM_Num = Convert.ToInt32(ConfigurationManager.AppSettings["Max_Create_VM_Num"].Trim().ToString());
            string vm_account = ConfigurationManager.AppSettings["vm_account"].Trim().ToString();
            string vm_password = ConfigurationManager.AppSettings["vm_password"].Trim().ToString();

            #region //是否有訂單需要刪除
            try//是否有訂單需要刪除
            {
                Create_VM_Service.VMAPI.GCCA_HypervisorAPI vmapi = new GCCA_HypervisorAPI();
                dbManager.Open();
                string sql = "";
                Int32 nCount = 0;
                sql = @"select TOP 1 a.order_audit,a.order_vm_type,a.order_id,order_area,a.company_id,a.temp_id,b.*,c.*
                        from user_vm_order a 
                        left outer join config_vm_host b on a.company_id=b.company_id and a.order_vm_type=b.vmtype left outer join order_audit c on a.order_id=c.order_id
                        where c.vm_del='1' and a.order_audit='6'
                        order by a.order_id";
                DataSet ds = dbManager.ExecuteDataSet(CommandType.Text, sql);
                nCount = ds.Tables[0].Rows.Count;
                System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " 是否有訂單需要刪除 : " + nCount + Environment.NewLine);

                if (nCount > 0)
                {
                    string order_id = ds.Tables[0].Rows[0]["order_id"].ToString();
                    string aa = ws.Change_Order_Status(order_id, "8", false);
                    string vmtype = ds.Tables[0].Rows[0]["order_vm_type"].ToString();
                    if (vmtype == "VMware")
                    {
                        string vcenter_url = ds.Tables[0].Rows[0]["vmware_apiurl"].ToString();
                        string vmware_host_account = ds.Tables[0].Rows[0]["vmware_host_account"].ToString();
                        string vmware_host_pwd = CryptoAES.decrypt(ds.Tables[0].Rows[0]["vmware_host_pwd"].ToString(), "GccA@stanchengGg");
                        string vmware_datacenter_name = ds.Tables[0].Rows[0]["vmware_datacenter_name"].ToString();
                        string vmware_host_name = ds.Tables[0].Rows[0]["vmware_host_name"].ToString();
                        string vmware_datastore_name = ds.Tables[0].Rows[0]["vmware_datastore_name"].ToString();
                        string[] guidf0 = vmapi.init(vcenter_url, vmware_host_account, vmware_host_pwd, vmware_datacenter_name, true).Split(':');
                        string[] guidf1 = guidf0[2].Split('\'');
                        string guid = guidf1[1];

                        string VMWare_Power_off_f = vmapi.powerOffVM(guid, vmware_datacenter_name, vmware_host_name, order_id).Split(':')[1].Split(',')[0];
                        try
                        {
                            if (IsDelImg == "true" && VMWare_Power_off_f == "true")
                            {
                                vmapi.removeVM(guid, vmware_datacenter_name, vmware_host_name, order_id, vmware_datastore_name, true);
                                ws.Change_Order_Status(order_id, "7", false);
                            }
                            else if (IsDelImg == "true" && VMWare_Power_off_f == "false")
                            {
                                ws.Inset_Log(order_id, "Power Off Error");
                            }
                            else if (IsDelImg == "false" && VMWare_Power_off_f == "false")
                            {
                                ws.Inset_Log(order_id, "Power Off Error");
                            }
                            else if (IsDelImg == "false" && VMWare_Power_off_f == "true")
                            {
                                ws.Change_Order_Status(order_id, "7", false);
                            }
                            System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " delete order is : " + order_id + Environment.NewLine);

                        }
                        catch (Exception ex)
                        {
                            dbManager.Dispose();
                            ws.Dispose();
                            ws.Inset_Log(order_id, ex.Message);
                            System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " delete order tppe is VMware,and fail order_id is : " + order_id + Environment.NewLine);

                        }
                    }
                    else if (vmtype == "KVM")
                    {
                        string kvm_hostname = ds.Tables[0].Rows[0]["kvm_hostname"].ToString();
                        string kvm_dsname = ds.Tables[0].Rows[0]["kvm_dsname"].ToString();
                        string kvm_account = ds.Tables[0].Rows[0]["kvm_account"].ToString();
                        string kvm_pwd = CryptoAES.decrypt(ds.Tables[0].Rows[0]["kvm_pwd"].ToString(), "GccA@stanchengGg");
                        string[] guidf0 = vmapi.init(kvm_hostname, kvm_dsname, kvm_account, kvm_pwd).Split(':');
                        string[] guidf1 = guidf0[2].Split('\'');
                        string guid = guidf1[1];
                        string VMWare_Power_off_f = vmapi.powerOffVM(guid, "", kvm_hostname, order_id).Split(':')[1].Split(',')[0];
                        try
                        {
                            if (IsDelImg == "true" && VMWare_Power_off_f == "true")
                            {
                                vmapi.removeVM(guid, "", kvm_hostname, order_id, kvm_dsname, true);
                                ws.Change_Order_Status(order_id, "7", false);
                            }
                            else if (IsDelImg == "true" && VMWare_Power_off_f == "false")
                            {
                                ws.Inset_Log(order_id, "Power Off Error");
                            }
                            else if (IsDelImg == "false" && VMWare_Power_off_f == "false")
                            {
                                ws.Inset_Log(order_id, "Power Off Error");
                            }
                            else if (IsDelImg == "false" && VMWare_Power_off_f == "true")
                            {
                                ws.Change_Order_Status(order_id, "7", false);
                            }
                            System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " delete order is : " + order_id + Environment.NewLine);

                        }
                        catch (Exception ex)
                        {
                            dbManager.Dispose();
                            ws.Dispose();
                            ws.Inset_Log(order_id, ex.Message);
                            System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " delete order tppe is KVM,and fail order_id is : " + order_id + Environment.NewLine);

                        }
                    }

                }
            }
            catch (Exception)
            {
                dbManager.Dispose();
                ws.Dispose();
            }
            finally
            {
                dbManager.Dispose();
            }
            #endregion

            #region 是否有訂單完成  vm create ready
            try//是否有訂單完成  vm create ready 
            {
                dbManager.Open();
                string sql = "";
                Int32 nCount = 0;
                //
                sql = @"select TOP 1 e.os_type as os,a.order_audit,a.FQDN+'@'+f.domain_name as FQDN,a.order_vm_type,order_id,order_area,a.order_cpu,a.order_ram,a.temp_id,order_vm_type,c.vpath,a.company_id,a.temp_id,a.group_id
                        from user_vm_order a 
                        left outer join vm_temp b on a.temp_id=b.temp_id 
                        left outer join vm_temp_virus_r c on a.temp_id=c.temp_id and a.order_virus=c.virus 
                        left outer join Param e on b.os=e.para_id
                        left outer join c_domain f on a.company_id=f.company_id and a.order_area=f.area_id
                        where a.order_audit='5'
                        order by order_id";
                DataSet ds = dbManager.ExecuteDataSet(CommandType.Text, sql);
                nCount = ds.Tables[0].Rows.Count;
                System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "訂單完成數 : " + nCount + Environment.NewLine);

                if (nCount > 0)
                {
                    System.Threading.Thread.Sleep(18000);
                    string order_id = ds.Tables[0].Rows[0]["order_id"].ToString();
                    string FQDN = ds.Tables[0].Rows[0]["FQDN"].ToString();
                    ws.Send_mail(order_id, FQDN);
                    ws.Inset_Percent(order_id, "100", "");
                    ws.Change_Order_Status(order_id, "1", true);
                    dbManager.CreateParameters(1);
                    dbManager.AddParameters(0, "order_id", order_id);
                    string sql2 = @"update user_vm_order
                                set finish_time=getdate()
                                where order_id=@order_id";
                    DataSet ds2 = dbManager.ExecuteDataSet(CommandType.Text, sql2);
                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "send mail, change order status ok. "  + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                dbManager.Dispose();
                ws.Dispose();
                System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "訂單完成產生錯誤 : " + ex.Message + Environment.NewLine);
            }
            finally
            {
                dbManager.Dispose();
            }
            #endregion

            try//是否有訂單需建立 need create 
            {
                dbManager.Open();
                string check_order = @"select order_audit
                                   from user_vm_order
                                   where order_audit = '3' ";
                DataSet ds0 = dbManager.ExecuteDataSet(CommandType.Text, check_order);
                Int32 n1 = ds0.Tables[0].Rows.Count;
                System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "craeting VM number is : " + n1 + Environment.NewLine);
                if (n1 <= Max_Create_VM_Num)                                //一次只能建立?筆訂單
                {
                    try
                    {
                        string sql = "";
                        Int32 nCount = 0;
                        Int32 result = 0;
                        sql = @"SELECT TOP 1 e.os_type as os,a.order_audit,a.order_vm_type,order_id,order_area,a.order_cpu,
                                        a.order_ram,a.temp_id,order_vm_type,c.vpath,a.company_id,a.temp_id,a.group_id,a.vlan_id,a.order_nhd
                                FROM user_vm_order a 
                                        left outer join vm_temp b on a.temp_id=b.temp_id 
                                        left outer join vm_temp_virus_r c on a.temp_id=c.temp_id and a.order_virus=c.virus 
                                        left outer join Param e on b.os=e.para_id
                                WHERE a.order_audit='2'
                                ORDER BY order_id";
                        DataSet ds = dbManager.ExecuteDataSet(CommandType.Text, sql);
                        nCount = ds.Tables[0].Rows.Count;
                        Create_VM_Service.VMAPI.GCCA_HypervisorAPI vmapi = new GCCA_HypervisorAPI();
                        string group_id = "";
                        string company_id = "";
                        string order_id = "";
                        string os = "";
                        string vmtype = "";
                        string order_area = "";
                        string temp_id = "";
                        string order_vm_type = "";
                        string vpath = "";
                        string cpu = "";
                        string ram = "";
                        string hdSize = "";
                        Int32 add_nic_num = 0;
                        string[] vlan_id = { };
                        string virNetworkName = "";
                        string[] virNetworkName_m = { };
                        string kvm_account = "";
                        string kvm_dsname = "";
                        string kvm_hostname = "";
                        string kvm_pwd = "";
                        string vmware_apiurl = "";
                        string vmware_datacenter_name = "";
                        string vmware_host_encryp_pwd = "";
                        string vmware_datastore_name = "";
                        string vmware_host_account = "";
                        string vmware_host_name = "";
                        string vmware_host_pwd = "";
                        string create_on_hostname = "";
                        string create_on_datacentername = "";
                        string resource_pool_name = "";
                        string guid = "";
                        bool exist_group_flag = false;
                        bool create_flag = false;

                        #region 是否有訂單需建立 need create vm have group(?)

                        if (nCount == 1)   // 根本一次只能建立一筆xddd
                        {
                            group_id = ds.Tables[0].Rows[0]["group_id"].ToString();
                            company_id = ds.Tables[0].Rows[0]["company_id"].ToString();
                            order_id = ds.Tables[0].Rows[0]["order_id"].ToString();
                            os = ds.Tables[0].Rows[0]["os"].ToString();
                            vmtype = ds.Tables[0].Rows[0]["order_vm_type"].ToString();
                            order_area = ds.Tables[0].Rows[0]["order_area"].ToString();
                            temp_id = ds.Tables[0].Rows[0]["temp_id"].ToString();
                            order_vm_type = ds.Tables[0].Rows[0]["order_vm_type"].ToString();
                            vpath = ds.Tables[0].Rows[0]["vpath"].ToString();
                            cpu = ds.Tables[0].Rows[0]["order_cpu"].ToString();
                            ram = Convert.ToString(Convert.ToInt16(ds.Tables[0].Rows[0]["order_ram"].ToString()) * 1024);
                            add_nic_num = ds.Tables[0].Rows[0]["vlan_id"].ToString().Split(',').Count() - 1;
                            vlan_id = ds.Tables[0].Rows[0]["vlan_id"].ToString().Split(',');
                            hdSize = ds.Tables[0].Rows[0]["order_nhd"].ToString();
                            dbManager.CreateParameters(1);
                            dbManager.AddParameters(0, "order_id", order_id);
                            sql = @"UPDATE user_vm_order  
                                         SET order_audit='3',upd_datetime=getdate()  
                                         WHERE order_id=@order_id
                                        update user_vm_order
                                        set create_time=getdate()
                                        where order_id=@order_id";
                            result = dbManager.ExecuteNonQuery(CommandType.Text, sql);
                            if (group_id != "")//有群組的話
                            {
                                ws.Inset_Percent(order_id, "10", "");
                                if (vmtype == "KVM")
                                {
                                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "order_id is  : " + order_id + " VMtype is : " + vmtype + Environment.NewLine);

                                    dbManager.CreateParameters(4);
                                    dbManager.AddParameters(0, "@order_area", order_area);
                                    dbManager.AddParameters(1, "@company_id", company_id);
                                    dbManager.AddParameters(2, "@vmtype", vmtype);
                                    dbManager.AddParameters(3, "@temp_id", temp_id);
                                    sql = @"SELECT kvm_account,kvm_dsname,kvm_hostname,kvm_pwd,b.hostname as create_on_hostname
                                            FROM config_vm_host a left outer join vm_temp b on a.vmtype=b.vm_type
                                            WHERE area=@order_area and a.company_id=@company_id and vmtype=@vmtype and temp_id=@temp_id";
                                    ds = dbManager.ExecuteDataSet(CommandType.Text, sql);
                                    nCount = ds.Tables[0].Rows.Count;
                                    if (nCount == 1)//如果KVM只有一台HOST
                                    {
                                        #region assign host
                                        kvm_account = ds.Tables[0].Rows[0]["kvm_account"].ToString();
                                        kvm_dsname = ds.Tables[0].Rows[0]["kvm_dsname"].ToString();
                                        kvm_hostname = ds.Tables[0].Rows[0]["kvm_hostname"].ToString();
                                        create_on_hostname = ds.Tables[0].Rows[0]["create_on_hostname"].ToString();
                                        kvm_pwd = CryptoAES.decrypt(ds.Tables[0].Rows[0]["kvm_pwd"].ToString(), "GccA@stanchengGg");

                                        //string kvm_usage = ws.get_KVM_HOST_Usage(kvm_hostname, kvm_account, kvm_pwd);
                                        //string kvm_vm_usage = ws.get_KVM_VM_Usage(kvm_hostname, kvm_dsname, kvm_account, kvm_pwd);
                                        string[] guidf0 = vmapi.init(kvm_hostname, kvm_dsname, kvm_account, kvm_pwd).Split(':');
                                        string[] guidf1 = guidf0[2].Split('\'');

                                        // TODO: add comment
                                        guid = guidf1[1];

                                        string nic_name1 = vmapi.getHostNetworkList(guid, "", kvm_hostname);
                                        //JToken temp = JObject.Parse(nic_name1);

                                        Int32 kvm_count = Regex.Split(nic_name1, "\"message\":")[1].Split(',').Count() / 2;
                                        for (int kc = 1; kc <= kvm_count; kc++)//判斷定單上Vlan_id有沒有與HOST上VLAN_ID相符合
                                        {
                                            if (vlan_id[kc] != (Regex.Split(nic_name1, "\"message\":")[1].Split('[')[1].Split(']')[0].Split('{')[kc].Split('}')[0].Split(',')[1]).Split('\"')[3])
                                            {
                                                exist_group_flag = false;
                                            }
                                            else
                                            {
                                                virNetworkName = Regex.Split(nic_name1, "\"message\":")[1].Split('[')[1].Split(']')[0].Split('{')[kc].Split('}')[0].Split(',')[0].Split('\"')[3];
                                                exist_group_flag = true;
                                                break;
                                            }
                                        }
                                        #endregion
                                    }

                                    #region assign host
                                    //else if (nCount > 1)//KVM大於一台HOST將所有HOST的資源做比較抓出資源最多之HOST
                                    //{
                                    //    int kvm_memory1 = 0;
                                    //    int kvm_memory2 = 0;
                                    //    int kvm_vm_cpu1 = 0;
                                    //    int kvm_vm_cpu2 = 0;
                                    //    for (int i = 0; i < nCount; i++)
                                    //    {
                                    //        string kvm_vm_usage_memory = "";
                                    //        string kvm_useage = "";
                                    //        string kvm_account2 = "";
                                    //        string kvm_dsname2 = "";
                                    //        string kvm_vm_usage_memory2 = "";
                                    //        string kvm_hostname2 = "";
                                    //        string kvm_pwd2 = "";
                                    //        string kvm_useage2 = "";
                                    //        string create_on_hostname2 = "";
                                    //        string virNetworkName2 = "";

                                    //        if (i == 0 && exist_group_flag == false)
                                    //        {

                                    //            kvm_account = ds.Tables[0].Rows[i]["kvm_account"].ToString();
                                    //            kvm_dsname = ds.Tables[0].Rows[i]["kvm_dsname"].ToString();
                                    //            kvm_hostname = ds.Tables[0].Rows[i]["kvm_hostname"].ToString();
                                    //            kvm_pwd = ds.Tables[0].Rows[i]["kvm_pwd"].ToString();
                                    //            create_on_hostname = ds.Tables[0].Rows[0]["create_on_hostname"].ToString();

                                    //            string[] guidf0 = vmapi.init(kvm_hostname, kvm_dsname, kvm_account, kvm_pwd).Split(':');
                                    //            string[] guidf1 = guidf0[2].Split('\'');

                                    //            guid = guidf1[1];

                                    //            string nic_name1 = vmapi.getHostNetworkList(guid, "", kvm_hostname);
                                    //            Int32 kvm_count = Regex.Split(nic_name1, "\"message\":")[1].Split(',').Count() / 2;
                                    //            for (int kc = 1; kc <= kvm_count; kc++)
                                    //            {
                                    //                if (group_id != (Regex.Split(nic_name1, "\"message\":")[1].Split('[')[1].Split(']')[0].Split('{')[kc].Split('}')[0].Split(',')[1]).Split('\"')[3])
                                    //                {
                                    //                    exist_group_flag = false;
                                    //                }
                                    //                else
                                    //                {
                                    //                    virNetworkName = Regex.Split(nic_name1, "\"message\":")[1].Split('[')[1].Split(']')[0].Split('{')[kc].Split('}')[0].Split(',')[0].Split('\"')[3];
                                    //                    exist_group_flag = true;
                                    //                    create_flag = true;
                                    //                    break;
                                    //                }
                                    //            }
                                    //            if (exist_group_flag == true)
                                    //            {
                                    //                kvm_useage = ws.get_KVM_HOST_Usage(kvm_hostname, kvm_account, kvm_pwd);//KVM_HOST有多少MEMORY
                                    //                kvm_vm_usage_memory = ws.get_KVM_VM_Usage(kvm_hostname, kvm_dsname, kvm_account, kvm_pwd);//在此HOST上VM使用多少MEMORY及
                                    //                string kvm_host_memory = kvm_useage.Split(':')[1].Split(']')[0].Trim();
                                    //                string kvm_vm_cpu = kvm_vm_usage_memory.Split(':')[2].Trim();
                                    //                string kvm_vm_memory = kvm_vm_usage_memory.Split(':')[1].Split(' ')[1].ToString().Trim();
                                    //                kvm_memory1 = (Convert.ToInt32(kvm_host_memory)) - (Convert.ToInt32(kvm_vm_memory));
                                    //                kvm_vm_cpu1 = (Convert.ToInt32(kvm_vm_cpu));
                                    //            }

                                    //        }
                                    //        else if (i > 0 && exist_group_flag == false)
                                    //        {

                                    //            kvm_account = ds.Tables[0].Rows[i]["kvm_account"].ToString();
                                    //            kvm_dsname = ds.Tables[0].Rows[i]["kvm_dsname"].ToString();
                                    //            kvm_hostname = ds.Tables[0].Rows[i]["kvm_hostname"].ToString();
                                    //            kvm_pwd = ds.Tables[0].Rows[i]["kvm_pwd"].ToString();
                                    //            create_on_hostname = ds.Tables[0].Rows[0]["create_on_hostname"].ToString();

                                    //            string[] guidf0 = vmapi.init(kvm_hostname, kvm_dsname, kvm_account, kvm_pwd).Split(':');
                                    //            string[] guidf1 = guidf0[2].Split('\'');

                                    //            guid = guidf1[1];

                                    //            string nic_name1 = vmapi.getHostNetworkList(guid, "", kvm_hostname);
                                    //            Int32 kvm_count = Regex.Split(nic_name1, "\"message\":")[1].Split(',').Count() / 2;
                                    //            for (int kc = 1; kc <= kvm_count; kc++)
                                    //            {
                                    //                if (group_id != (Regex.Split(nic_name1, "\"message\":")[1].Split('[')[1].Split(']')[0].Split('{')[kc].Split('}')[0].Split(',')[1]).Split('\"')[3])
                                    //                {
                                    //                    exist_group_flag = false;
                                    //                }
                                    //                else
                                    //                {
                                    //                    virNetworkName = Regex.Split(nic_name1, "\"message\":")[1].Split('[')[1].Split(']')[0].Split('{')[kc].Split('}')[0].Split(',')[0].Split('\"')[3];
                                    //                    exist_group_flag = true;
                                    //                    create_flag = true;
                                    //                    break;
                                    //                }
                                    //            }
                                    //            if (exist_group_flag == true)
                                    //            {
                                    //                kvm_useage = ws.get_KVM_HOST_Usage(kvm_hostname, kvm_account, kvm_pwd);
                                    //                kvm_vm_usage_memory = ws.get_KVM_VM_Usage(kvm_hostname, kvm_dsname, kvm_account, kvm_pwd);
                                    //                string kvm_host_memory = kvm_useage.Split(':')[1].Split(']')[0].Trim();
                                    //                string kvm_vm_cpu = kvm_vm_usage_memory.Split(':')[2].Trim();
                                    //                string kvm_vm_memory = kvm_vm_usage_memory.Split(':')[1].Split(' ')[1].ToString().Trim();
                                    //                kvm_memory1 = (Convert.ToInt32(kvm_host_memory)) - (Convert.ToInt32(kvm_vm_memory));
                                    //                kvm_vm_cpu1 = (Convert.ToInt32(kvm_vm_cpu));
                                    //            }
                                    //        }
                                    //        else if (i > 0 && exist_group_flag == true)
                                    //        {
                                    //            kvm_account2 = ds.Tables[0].Rows[i]["kvm_account"].ToString();
                                    //            kvm_dsname2 = ds.Tables[0].Rows[i]["kvm_dsname"].ToString();
                                    //            kvm_hostname2 = ds.Tables[0].Rows[i]["kvm_hostname"].ToString();
                                    //            kvm_pwd2 = ds.Tables[0].Rows[i]["kvm_pwd"].ToString();
                                    //            create_on_hostname2 = ds.Tables[0].Rows[0]["create_on_hostname"].ToString();

                                    //            string[] guidf0 = vmapi.init(kvm_hostname2, kvm_dsname2, kvm_account2, kvm_pwd2).Split(':');
                                    //            string[] guidf1 = guidf0[2].Split('\'');

                                    //            guid = guidf1[1];

                                    //            string nic_name1 = vmapi.getHostNetworkList(guid, "", kvm_hostname2);
                                    //            Int32 kvm_count = Regex.Split(nic_name1, "\"message\":")[1].Split(',').Count() / 2;
                                    //            for (int kc = 1; kc <= kvm_count; kc++)
                                    //            {
                                    //                for (int nic_count = 0; nic_count <= add_nic_num; nic_count++)
                                    //                {
                                    //                    if (vlan_id[nic_count] != (Regex.Split(nic_name1, "\"message\":")[1].Split('[')[1].Split(']')[0].Split('{')[kc].Split('}')[0].Split(',')[1]).Split('\"')[3])
                                    //                    {
                                    //                        exist_group_flag = false;
                                    //                    }
                                    //                    else
                                    //                    {
                                    //                        virNetworkName = Regex.Split(nic_name1, "\"message\":")[1].Split('[')[1].Split(']')[0].Split('{')[kc].Split('}')[0].Split(',')[0].Split('\"')[3];
                                    //                        exist_group_flag = true;
                                    //                        create_flag = true;
                                    //                        break;
                                    //                    }
                                    //                }
                                    //            }
                                    //            if (exist_group_flag == true)
                                    //            {
                                    //                kvm_useage2 = ws.get_KVM_HOST_Usage(kvm_hostname2, kvm_account2, kvm_pwd2);
                                    //                kvm_vm_usage_memory2 = ws.get_KVM_VM_Usage(kvm_hostname2, kvm_dsname2, kvm_account2, kvm_pwd2);
                                    //                string kvm_host_memory = kvm_useage2.Split(':')[1].Split(']')[0].Trim();
                                    //                string kvm_vm_cpu = kvm_vm_usage_memory2.Split(':')[2].Trim();
                                    //                string kvm_vm_memory = kvm_vm_usage_memory2.Split(':')[1].Split(' ')[1].ToString().Trim();
                                    //                kvm_memory2 = (Convert.ToInt32(kvm_host_memory)) - (Convert.ToInt32(kvm_vm_memory));
                                    //                kvm_vm_cpu2 = (Convert.ToInt32(kvm_vm_cpu));
                                    //                if (kvm_memory1 < kvm_memory2)
                                    //                {
                                    //                    virNetworkName = virNetworkName2;
                                    //                    kvm_account = kvm_account2;
                                    //                    kvm_dsname = kvm_dsname2;
                                    //                    kvm_hostname = kvm_hostname2;
                                    //                    kvm_pwd = kvm_pwd2;
                                    //                    create_on_hostname = create_on_hostname2;
                                    //                    kvm_memory1 = kvm_memory2;
                                    //                }
                                    //                else if (kvm_memory1 == kvm_memory2)
                                    //                {
                                    //                    if (kvm_vm_cpu1 < kvm_vm_cpu2)
                                    //                    {
                                    //                        virNetworkName = virNetworkName2;
                                    //                        kvm_account = kvm_account2;
                                    //                        kvm_dsname = kvm_dsname2;
                                    //                        kvm_hostname = kvm_hostname2;
                                    //                        kvm_pwd = kvm_pwd2;
                                    //                        create_on_hostname = create_on_hostname2;
                                    //                        kvm_memory1 = kvm_memory2;
                                    //                    }
                                    //                }
                                    //            }
                                    //        }
                                    //    }
                                    //}
                                    // 
                                    #endregion
                                }
                                if (vmtype == "VMware")
                                {
                                    dbManager.CreateParameters(4);
                                    dbManager.AddParameters(0, "@order_area", order_area);
                                    dbManager.AddParameters(1, "@company_id", company_id);
                                    dbManager.AddParameters(2, "@vmtype", vmtype);
                                    dbManager.AddParameters(3, "@temp_id", temp_id);
                                    sql = @"SELECT vmware_apiurl,vmware_datacenter_name,vmware_datastore_name,vmware_host_account,vmware_host_name,vmware_host_pwd,b.hostname as create_on_hostname,resource_pool_name,b.datacenter_name as create_on_datacentername,b.temp_id
                                            FROM config_vm_host a left outer join vm_temp b on a.vmtype=b.vm_type 
                                            WHERE area=@order_area and a.company_id=@company_id and vmtype=@vmtype and b.temp_id=@temp_id";
                                    ds = dbManager.ExecuteDataSet(CommandType.Text, sql);
                                    nCount = ds.Tables[0].Rows.Count;
                                    if (nCount == 1)//如果只有一台VMWARE HOST
                                    {
                                        System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "order_id is  : " + order_id + " VMtype is : " + vmtype + Environment.NewLine);

                                        #region assign host
                                        vmware_apiurl = ds.Tables[0].Rows[0]["vmware_apiurl"].ToString();
                                        string vmware_vcenter_ip = vmware_apiurl.Split('/')[2];
                                        vmware_datacenter_name = ds.Tables[0].Rows[0]["vmware_datacenter_name"].ToString();
                                        vmware_datastore_name = ds.Tables[0].Rows[0]["vmware_datastore_name"].ToString();
                                        vmware_host_account = ds.Tables[0].Rows[0]["vmware_host_account"].ToString();
                                        vmware_host_name = ds.Tables[0].Rows[0]["vmware_host_name"].ToString();
                                        vmware_host_pwd = CryptoAES.decrypt(ds.Tables[0].Rows[0]["vmware_host_pwd"].ToString(), "GccA@stanchengGg");
                                        vmware_host_encryp_pwd = ds.Tables[0].Rows[0]["vmware_host_pwd"].ToString();
                                        create_on_hostname = ds.Tables[0].Rows[0]["create_on_hostname"].ToString();
                                        create_on_datacentername = ds.Tables[0].Rows[0]["create_on_datacentername"].ToString();
                                        resource_pool_name = Convert.ToString(ds.Tables[0].Rows[0]["resource_pool_name"]);
                                        //string vmware_usage = vm_useage.get_VMware_HOST_Usage(vmware_vcenter_ip, vmware_host_name, vmware_host_account, vmware_host_pwd);
                                        //string vmware_vm_usage = vm_useage.get_VMware_VM_Usage(vmware_apiurl,vmware_host_account,vmware_host_pwd,vmware_datacenter_name,vmware_host_name);
                                        string[] guidf0 = vmapi.init(vmware_apiurl, vmware_host_account, vmware_host_pwd, vmware_datacenter_name, true).Split(':');
                                        string[] guidf1 = guidf0[2].Split('\'');
                                        guid = guidf1[1];
                                        string nic_name1 = vmapi.getHostNetworkList(guid, vmware_datacenter_name, vmware_host_name);
                                        Int32 vm_count = Regex.Split(nic_name1, "\"message\":")[1].Split(',').Count() / 2;
                                        for (int vc = 1; vc <= vm_count; vc++)
                                        {
                                            if (exist_group_flag != true)
                                            {
                                                for (int nic_count = 0; nic_count <= add_nic_num; nic_count++)
                                                {
                                                    if (vlan_id[nic_count] != (Regex.Split(nic_name1, "\"message\":")[1].Split('[')[1].Split(']')[0].Split('{')[vc].Split('}')[0].Split(',')[1]).Split('\"')[3])
                                                    {
                                                        exist_group_flag = false;
                                                    }
                                                    else
                                                    {
                                                        string vlan_id1 = vlan_id[nic_count];
                                                        virNetworkName = Regex.Split(nic_name1, "\"message\":")[1].Split('[')[1].Split(']')[0].Split('{')[vc].Split('}')[0].Split(',')[0].Split('\"')[3];
                                                        virNetworkName_m = Regex.Split(nic_name1, "\"message\":")[1].Replace("\"", "").Replace("{", "").Replace("[", "").Replace("}", "").Replace("]", "").Replace(")", "").Replace(":", ",").Split(',');
                                                        int count = 0;
                                                        virNetworkName_st[] host_nic_and_vlan = new virNetworkName_st[virNetworkName_m.Count() / 4];
                                                        for (int ii = 3; ii <= virNetworkName_m.Count(); ii = ii + 4)
                                                        {
                                                            host_nic_and_vlan[count].vlan_id = virNetworkName_m[ii];
                                                            host_nic_and_vlan[count].network_name = virNetworkName_m[ii - 2];
                                                            count++;
                                                            //if (ii != virNetworkName_m.Count()-1)
                                                            //{
                                                            //    vlan_id2 += "[" + virNetworkName_m[ii] + "," + virNetworkName_m[ii - 2] + "]" + ",";
                                                            //}
                                                            //else 
                                                            //{
                                                            //    vlan_id2 += "[" + virNetworkName_m[ii] + "," + virNetworkName_m[ii - 2] + "]";
                                                            //}
                                                        }
                                                        for (int iii = 0; iii < vlan_id.Count(); iii++)
                                                        {
                                                            for (int ii = 0; ii < host_nic_and_vlan.Count(); ii++)
                                                            {
                                                                if (vlan_id[iii] == host_nic_and_vlan[ii].vlan_id)
                                                                {
                                                                    vlan_id[iii] = host_nic_and_vlan[ii].vlan_id;
                                                                }
                                                            }
                                                        }
                                                        exist_group_flag = true;
                                                        create_flag = true;
                                                    }
                                                }
                                            }
                                        }
                                        #endregion
                                    }
                                    //
                                    //else if (nCount > 1)//如果大於一台VMWARE HOST將所有HOST的資源做比較抓出資源最多之HOST
                                    //{
                                    #region assign host


                                    //    int vmware_memory1 = 0;
                                    //    int vmware_memory2 = 0;
                                    //    int vmware_vm_cpu1 = 0;
                                    //    int vmware_vm_cpu2 = 0;
                                    //    for (int i = 0; i < nCount; i++)
                                    //    {
                                    //        string vmware_vm_usage_memory = "";
                                    //        string vmware_useage = "";
                                    //        string vmware_vm_usage_memory2 = "";
                                    //        string vmware_host_account2 = "";
                                    //        string vmware_datastore_name2 = "";
                                    //        string vmware_datacenter_name2 = "";
                                    //        string vmware_apiurl2 = "";
                                    //        string vmware_host_name2 = "";
                                    //        string vmware_host_pwd2 = "";
                                    //        string vmware_useage2 = "";
                                    //        string vmware_host_encryp_pwd2 = "";
                                    //        string create_on_hostname2 = "";
                                    //        string create_on_datacentername2 = "";
                                    //        string resource_pool_name2 = "";
                                    //        string virNetworkName2 = "";
                                    //        if (i == 0 && exist_group_flag == false)
                                    //        {
                                    //            vmware_host_account = ds.Tables[0].Rows[i]["vmware_host_account"].ToString();
                                    //            vmware_datastore_name = ds.Tables[0].Rows[i]["vmware_datastore_name"].ToString();
                                    //            vmware_datacenter_name = ds.Tables[0].Rows[i]["vmware_datacenter_name"].ToString();
                                    //            vmware_apiurl = ds.Tables[0].Rows[i]["vmware_apiurl"].ToString();
                                    //            string vmware_vcenter_ip = vmware_apiurl.Split('/')[2];
                                    //            vmware_host_name = ds.Tables[0].Rows[i]["vmware_host_name"].ToString();
                                    //            vmware_host_pwd = CryptoAES.decrypt(ds.Tables[0].Rows[i]["vmware_host_pwd"].ToString(), "GccA@stanchengGg");
                                    //            vmware_host_encryp_pwd = ds.Tables[0].Rows[i]["vmware_host_pwd"].ToString();
                                    //            create_on_hostname = ds.Tables[0].Rows[0]["create_on_hostname"].ToString();
                                    //            create_on_datacentername = ds.Tables[0].Rows[0]["create_on_datacentername"].ToString();
                                    //            resource_pool_name = Convert.ToString(ds.Tables[0].Rows[0]["resource_pool_name"]);
                                    //            string[] guidf0 = vmapi.init(vmware_apiurl, vmware_host_account, vmware_host_pwd, vmware_datacenter_name, true).Split(':');
                                    //            string[] guidf1 = guidf0[2].Split('\'');
                                    //            guid = guidf1[1];
                                    //            string nic_name1 = vmapi.getHostNetworkList(guid, vmware_datacenter_name, vmware_host_name);
                                    //            Int32 vm_count = Regex.Split(nic_name1, "\"message\":")[1].Split(',').Count() / 2;
                                    //            for (int vc = 1; vc <= vm_count; vc++)
                                    //            {
                                    //                for (int nic_count = 0; nic_count <= add_nic_num; nic_count++)
                                    //                {
                                    //                    if (vlan_id[nic_count] != (Regex.Split(nic_name1, "\"message\":")[1].Split('[')[1].Split(']')[0].Split('{')[vc].Split('}')[0].Split(',')[1]).Split('\"')[3])
                                    //                    {
                                    //                        exist_group_flag = false;
                                    //                    }
                                    //                    else
                                    //                    {
                                    //                        virNetworkName = Regex.Split(nic_name1, "\"message\":")[1].Split('[')[1].Split(']')[0].Split('{')[vc].Split('}')[0].Split(',')[0].Split('\"')[3];
                                    //                        exist_group_flag = true;
                                    //                        create_flag = true;
                                    //                        break;
                                    //                    }
                                    //                }
                                    //            }
                                    //            if (exist_group_flag == true)
                                    //            {
                                    //                vmware_useage = ws.get_VMware_HOST_Usage(vmware_vcenter_ip, vmware_host_name, vmware_host_account, vmware_host_pwd);
                                    //                vmware_vm_usage_memory = ws.get_VMware_VM_Usage(vmware_apiurl, vmware_host_account, vmware_host_pwd, vmware_datacenter_name, vmware_host_name);
                                    //                string vmware_host_memory = vmware_useage.Split(':')[1].Split(']')[0].Trim();
                                    //                string vmware_vm_cpu = vmware_vm_usage_memory.Split(':')[2].Trim();
                                    //                string vmware_vm_memory = vmware_vm_usage_memory.Split(':')[1].Split(' ')[1].Split('G')[0].Trim();
                                    //                vmware_memory1 = (Convert.ToInt32(vmware_host_memory)) - (Convert.ToInt32(vmware_vm_memory));
                                    //                vmware_vm_cpu1 = (Convert.ToInt32(vmware_vm_cpu));
                                    //            }
                                    //        }
                                    //        else if (i > 0 && exist_group_flag == false)
                                    //        {
                                    //            vmware_host_account = ds.Tables[0].Rows[i]["vmware_host_account"].ToString();
                                    //            vmware_datastore_name = ds.Tables[0].Rows[i]["vmware_datastore_name"].ToString();
                                    //            vmware_datacenter_name = ds.Tables[0].Rows[i]["vmware_datacenter_name"].ToString();
                                    //            vmware_apiurl = ds.Tables[0].Rows[i]["vmware_apiurl"].ToString();
                                    //            string vmware_vcenter_ip = vmware_apiurl.Split('/')[2];
                                    //            vmware_host_name = ds.Tables[0].Rows[i]["vmware_host_name"].ToString();
                                    //            vmware_host_pwd = CryptoAES.decrypt(ds.Tables[0].Rows[i]["vmware_host_pwd"].ToString(), "GccA@stanchengGg");
                                    //            vmware_host_encryp_pwd = ds.Tables[0].Rows[i]["vmware_host_pwd"].ToString();
                                    //            create_on_hostname = ds.Tables[0].Rows[0]["create_on_hostname"].ToString();
                                    //            create_on_datacentername = ds.Tables[0].Rows[0]["create_on_datacentername"].ToString();
                                    //            resource_pool_name = Convert.ToString(ds.Tables[0].Rows[0]["resource_pool_name"]);
                                    //            string[] guidf0 = vmapi.init(vmware_apiurl, vmware_host_account, vmware_host_pwd, vmware_datacenter_name, true).Split(':');
                                    //            string[] guidf1 = guidf0[2].Split('\'');
                                    //            guid = guidf1[1];
                                    //            string nic_name1 = vmapi.getHostNetworkList(guid, vmware_datacenter_name, vmware_host_name);
                                    //            Int32 vm_count = Regex.Split(nic_name1, "\"message\":")[1].Split(',').Count() / 2;
                                    //            for (int vc = 1; vc <= vm_count; vc++)
                                    //            {
                                    //                for (int nic_count = 0; nic_count <= add_nic_num; nic_count++)
                                    //                {
                                    //                    if (vlan_id[nic_count] != (Regex.Split(nic_name1, "\"message\":")[1].Split('[')[1].Split(']')[0].Split('{')[vc].Split('}')[0].Split(',')[1]).Split('\"')[3])
                                    //                    {
                                    //                        exist_group_flag = false;
                                    //                    }
                                    //                    else
                                    //                    {
                                    //                        virNetworkName = Regex.Split(nic_name1, "\"message\":")[1].Split('[')[1].Split(']')[0].Split('{')[vc].Split('}')[0].Split(',')[0].Split('\"')[3];
                                    //                        exist_group_flag = true;
                                    //                        create_flag = true;
                                    //                        break;
                                    //                    }
                                    //                }
                                    //            }
                                    //            if (exist_group_flag == true)
                                    //            {
                                    //                vmware_useage = ws.get_VMware_HOST_Usage(vmware_vcenter_ip, vmware_host_name, vmware_host_account, vmware_host_pwd);
                                    //                vmware_vm_usage_memory = ws.get_VMware_VM_Usage(vmware_apiurl, vmware_host_account, vmware_host_pwd, vmware_datacenter_name, vmware_host_name);
                                    //                string vmware_host_memory = vmware_useage.Split(':')[1].Split(']')[0].Trim();
                                    //                string vmware_vm_cpu = vmware_vm_usage_memory.Split(':')[2].Trim();
                                    //                string vmware_vm_memory = vmware_vm_usage_memory.Split(':')[1].Split(' ')[1].Split('G')[0].Trim();
                                    //                vmware_memory1 = (Convert.ToInt32(vmware_host_memory)) - (Convert.ToInt32(vmware_vm_memory));
                                    //                vmware_vm_cpu1 = (Convert.ToInt32(vmware_vm_cpu));
                                    //            }

                                    //        }
                                    //        else if (i > 0 && exist_group_flag == true)
                                    //        {
                                    //            vmware_host_account2 = ds.Tables[0].Rows[i]["vmware_host_account"].ToString();
                                    //            vmware_datastore_name2 = ds.Tables[0].Rows[i]["vmware_datastore_name"].ToString();
                                    //            vmware_datacenter_name2 = ds.Tables[0].Rows[i]["vmware_datacenter_name"].ToString();
                                    //            vmware_apiurl2 = ds.Tables[0].Rows[i]["vmware_apiurl"].ToString();
                                    //            string vmware_vcenter_ip = vmware_apiurl.Split('/')[2];
                                    //            vmware_host_name2 = ds.Tables[0].Rows[i]["vmware_host_name"].ToString();
                                    //            vmware_host_pwd2 = CryptoAES.decrypt(ds.Tables[0].Rows[i]["vmware_host_pwd"].ToString(), "GccA@stanchengGg");
                                    //            vmware_host_encryp_pwd2 = ds.Tables[0].Rows[i]["vmware_host_pwd"].ToString();
                                    //            create_on_hostname2 = ds.Tables[0].Rows[0]["create_on_hostname"].ToString();
                                    //            create_on_datacentername2 = ds.Tables[0].Rows[0]["create_on_datacentername"].ToString();
                                    //            resource_pool_name2 = Convert.ToString(ds.Tables[0].Rows[0]["resource_pool_name"]);

                                    //            string[] guidf0 = vmapi.init(vmware_apiurl, vmware_host_account, vmware_host_pwd, vmware_datacenter_name, true).Split(':');
                                    //            string[] guidf1 = guidf0[2].Split('\'');
                                    //            guid = guidf1[1];
                                    //            string nic_name1 = vmapi.getHostNetworkList(guid, vmware_datacenter_name2, vmware_host_name2);
                                    //            Int32 vm_count = Regex.Split(nic_name1, "\"message\":")[1].Split(',').Count() / 2;
                                    //            for (int vc = 1; vc <= vm_count; vc++)
                                    //            {
                                    //                for (int nic_count = 0; nic_count <= add_nic_num; nic_count++)
                                    //                {
                                    //                    if (vlan_id[nic_count] != (Regex.Split(nic_name1, "\"message\":")[1].Split('[')[1].Split(']')[0].Split('{')[vc].Split('}')[0].Split(',')[1]).Split('\"')[3])
                                    //                    {
                                    //                        exist_group_flag = false;
                                    //                    }
                                    //                    else
                                    //                    {
                                    //                        virNetworkName2 = Regex.Split(nic_name1, "\"message\":")[1].Split('[')[1].Split(']')[0].Split('{')[vc].Split('}')[0].Split(',')[0].Split('\"')[3];
                                    //                        exist_group_flag = true;
                                    //                        create_flag = true;
                                    //                        break;
                                    //                    }
                                    //                }
                                    //            }
                                    //            if (exist_group_flag == true)
                                    //            {
                                    //                vmware_useage2 = ws.get_VMware_HOST_Usage(vmware_vcenter_ip, vmware_host_name2, vmware_host_account2, vmware_host_pwd2);
                                    //                vmware_vm_usage_memory2 = ws.get_VMware_VM_Usage(vmware_apiurl2, vmware_host_account2, vmware_host_pwd2, vmware_datacenter_name2, vmware_host_name2);
                                    //                string vmware_host_memory = vmware_useage2.Split(':')[1].Split(']')[0].Trim();
                                    //                string vmware_vm_cpu = vmware_vm_usage_memory2.Split(':')[2].Trim();
                                    //                string vmware_vm_memory = vmware_vm_usage_memory2.Split(':')[1].Split(' ')[1].Split('G')[0].Trim();
                                    //                vmware_memory2 = (Convert.ToInt32(vmware_host_memory)) - (Convert.ToInt32(vmware_vm_memory));
                                    //                vmware_vm_cpu2 = (Convert.ToInt32(vmware_vm_cpu));
                                    //                if (vmware_memory1 < vmware_memory2)
                                    //                {
                                    //                    virNetworkName = virNetworkName2;
                                    //                    virNetworkName = virNetworkName2;
                                    //                    vmware_host_account = vmware_host_account2;
                                    //                    vmware_datastore_name = vmware_datastore_name2;
                                    //                    vmware_datacenter_name = vmware_datacenter_name2;
                                    //                    vmware_apiurl = vmware_apiurl2;
                                    //                    vmware_host_name = vmware_host_name2;
                                    //                    vmware_host_pwd = vmware_host_pwd2;
                                    //                    vmware_host_encryp_pwd = vmware_host_encryp_pwd2;
                                    //                    create_on_hostname = create_on_hostname2;
                                    //                    create_on_datacentername = create_on_datacentername2;
                                    //                    resource_pool_name = resource_pool_name2;
                                    //                    vmware_memory1 = vmware_memory2;
                                    //                }
                                    //                else if (vmware_memory1 == vmware_memory2)
                                    //                {
                                    //                    if (vmware_vm_cpu1 < vmware_vm_cpu2)
                                    //                    {
                                    //                        virNetworkName = virNetworkName2;
                                    //                        vmware_host_account = vmware_host_account2;
                                    //                        vmware_datastore_name = vmware_datastore_name2;
                                    //                        vmware_datacenter_name = vmware_datacenter_name2;
                                    //                        vmware_apiurl = vmware_apiurl2;
                                    //                        vmware_host_name = vmware_host_name2;
                                    //                        vmware_host_pwd = vmware_host_pwd2;
                                    //                        vmware_host_encryp_pwd = vmware_host_encryp_pwd2;
                                    //                        create_on_hostname = create_on_hostname2;
                                    //                        create_on_datacentername = create_on_datacentername2;
                                    //                        resource_pool_name = resource_pool_name2;
                                    //                        vmware_memory1 = vmware_memory2;
                                    //                    }
                                    //                }
                                    //            }
                                    //        }
                                    //    }
                                    #endregion
                                    //}
                                }


                                ws.Inset_Percent(order_id, "20", "");
                                //建立有GROUP之VM
                                vmapi.Timeout = 10000000;
                                try
                                {

                                    if (vmtype == "VMware")
                                    {
                                        if (create_flag == true)
                                        {
                                            string[] vcenter_ip_sp1 = vmware_apiurl.Split('/');
                                            string vcenter_ip = vcenter_ip_sp1[2];
                                            string[] guidf0 = vmapi.init(vmware_apiurl, vmware_host_account, vmware_host_pwd, vmware_datacenter_name, true).Split(':'); //連線至VMWARE HOST
                                            string[] guidf1 = guidf0[2].Split('\'');
                                            guid = guidf1[1];//取得GUID
                                            // string[] VMWare_VFT_sp0 = vmapi.createVMFromTemplate(guid, vpath, create_on_hostname, create_on_datacentername, order_id, vmware_host_name, vmware_datacenter_name, resource_pool_name, vmware_datastore_name).Split(':');//Create VM From Template
                                            // string[] VMWare_VFT_sp1 = VMWare_VFT_sp0[1].Split(',');
                                            string VMWare_VFT_F = "false";//VMWare_VFT_sp1[0];//Create VM From Template 是否成功
                                            System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "create VM is  : " + VMWare_VFT_F + Environment.NewLine);

                                            if (VMWare_VFT_F == "false")  ////還沒creat vm要改成true
                                            {
                                                //vmapi.powerOffVMAsync(guid, vmware_datacenter_name, vmware_host_name, order_id);//poweroff
                                                //ws.Inset_Percent(order_id, "35", "");
                                                //ws.set_create_on_host(order_id, vmware_host_name); // update order information with host name
                                                //vmapi.configVMCPUNum(guid, vmware_datacenter_name, vmware_host_name, order_id, cpu);//修改CPU數量
                                                //vmapi.configVMMemory(guid, vmware_datacenter_name, vmware_host_name, order_id, ram);//修改RAM大小

                                                //if (hdSize != "") //remove HD and add HD
                                                //{
                                                //    vmapi.removeVMDisk(guid, vmware_datacenter_name, vmware_datastore_name, vmware_datacenter_name, order_id, "1");
                                                //    vmapi.addVMDisk(guid, vmware_datacenter_name, vmware_host_name, vmware_datastore_name, order_id, (int.Parse(hdSize) * 1000).ToString());
                                                //}

                                                //string mac1 = vmapi.getVMNICMacList(guid, vmware_datacenter_name, vmware_host_name, order_id);
                                                //VM_set_adapter(vlan_id, virNetworkName_m, guid, vmware_datacenter_name, vmware_host_name, order_id); //暫時 FOR VMWARE   setting adapter. write by jedi
                                                System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "VM hardware set is Complete : " + order_id + Environment.NewLine);
                                                string[] VMWare_Power_on_sp0 = vmapi.powerOnVM(guid, vmware_datacenter_name, vmware_host_name, order_id).Split(':');//開機
                                                string[] VMWare_Power_on_sp1 = VMWare_Power_on_sp0[1].Split(',');
                                                string VMWare_Power_on_F = VMWare_Power_on_sp1[0];//是否開機成功
                                                if (VMWare_Power_on_F == "true")
                                                {
                                                    ws.set_create_on_host(order_id, vmware_host_name);
                                                    RunspaceInvoke invoker = new RunspaceInvoke();
                                                    invoker.Invoke("Set-ExecutionPolicy Unrestricted");
                                                    System.Threading.Thread.Sleep(1000);
                                                    string[] url0 = vmware_apiurl.Split('/');
                                                    vmware_apiurl = url0[2];
                                                    System.Threading.Thread.Sleep(24000);//開機完成後等待240000毫秒
                                                    string vm_network_name = vmapi.getVMNICNetworkList(guid, vmware_datacenter_name, vmware_host_name, order_id).Split(':')[3].Split('"')[1];//取得此台VM之網路名稱

                                                    string vm_ip = vmapi.getVMIpAndMac(guid, vmware_datacenter_name, vmware_host_name, order_id, vm_network_name).Split(':')[2].Split('"')[1];//取得此台VM之IP
                                                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "VM ip is  : " + vm_ip + Environment.NewLine);

                                                    if (os == "0")//For MicroSoft
                                                    {
                                                        System.Threading.Thread.Sleep(1000);
                                                        invoker.Invoke("Set-Item WSMan:\\localhost\\Client\\TrustedHosts " + vm_ip + " -Concatenate -force");
                                                        // TODO: do not write account/password into the script
                                                        string pingvm = @"$username = ""User""
                                                                      $account = """ + vm_account + @"""
                                                                      $password = ConvertTO-SecureString """ + vm_password + @""" -asplaintext -Force
                                                                      $cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $account, $password
                                                                      Invoke-Command -ComputerName " + vm_ip + @" -Authentication default" +
                                                                        @" -credential $cred " +
                                                                        @" -ScriptBlock {" +
                                                                            @"$FileExists = Test-Path ""c:\\AutoProvision"";" +
                                                                            @"If ($FileExists -eq $False){" +
                                                                                @"mkdir ""c:\\AutoProvision"";"
                                                                                + @"echo """ + order_id + " " + group_id + @""" >C:\\AutoProVision\\vmname.txt;"
                                                                                + @"echo ""timeout 30"""
                                                                                + @" `n 'msiexec /forcerestart /i  c:\AutoProVision\VMConfig.msi ALLUSERS=1 DB_Server=""" + db_server + @""" EmailAccount=""" + EmailAccount + "\" EmailPassword=\"" + EmailPassword + "\" smtphost=\"" + smtphost + "\" AutoProvision_WS_url=\"" + AutoProvision_WS_url + "\" /qn'"
                                                                                + @" `n ""timeout 10"" "
                                                                                + @"`n ""restart"" "
                                                                                + @"`n | "
                                                                                + @"Set-Content C:\AutoProVision\vmconfig.bat;"
                                                                                + @"$File = ""c:\AutoProvision\VMConfig.msi"";"
                                                                                + @"$ftp = """ + ftp_folder + "\\VMConfig.msi\";"
                                                                                + @"$File3 = ""c:\AutoProvision\creboot2.exe"";"
                                                                                + @"$ftp3 = """ + ftp_folder + @"\creboot2.exe"";"
                                                                                + @"$File4 = ""c:\AutoProvision\set_ip.exe"";"
                                                                                + @"$ftp4 = """ + ftp_folder + @"\set_ip.exe"";"
                                                                                + @"$webclient = New-Object System.Net.WebClient;"
                                                                                + @"$Username = """ + ConfigurationManager.AppSettings["ftpUsername"] + "\";"
                                                                                + @"$Password = """ + ConfigurationManager.AppSettings["ftpPassword"] + "\";"
                                                                                + @"$webclient = New-Object System.Net.WebClient;"
                                                                                + @"$WebClient.Credentials = New-Object System.Net.Networkcredential($Username, $Password);"
                                                                                + @"$webclient.DownloadFile($ftp, $File);"
                                                                                + @"$webclient.DownloadFile($ftp3, $File3);"
                                                                                + @"$webclient.DownloadFile($ftp4, $File4);"
                                                                                + @"Start-Sleep -s 10 ; "
                                                                            + @"}"
                                                                            + @"c:\AutoProvision\vmconfig.bat;"
                                                                        + "}";
                                                        //使用POWERSHELL 連至VM並對此台VM下達Create一個vmname.txt內容放入訂單編號及群組ID，下載FTP上之AGENT的檔案安裝完AGENT後並重開機。
                                                        execute(pingvm);

                                                    }
                                                    if (os == "1") //For Linux
                                                    {
                                                        string hostSshUserName = "gcca";
                                                        string hostSshPassword = "Gcca@gcca";
                                                        string agentFolderPath = "/opt/AutoConfigAgent";
                                                        // TODO: retrive url from outside
                                                        string wsUrl = AutoProvision_WS_url;
                                                        string hypervisorIpAddress = vmware_host_name;
                                                        string ftpUsername = ftp_user;
                                                        string ftpPassword = ftp_pwd;
                                                        string agentFtpPath = "/AutoConfigAgent";
                                                        string vmProvisionName = order_id;

                                                        executeLinuxScript(vm_ip, hostSshUserName, hostSshPassword, 22,
                                                        agentFolderPath, wsUrl, hypervisorIpAddress, ftp, ftpUsername, ftpPassword, agentFtpPath, vmProvisionName);
                                                    }
                                                    ws.Inset_Percent(order_id, "50", "");
                                                }
                                            }
                                        }
                                        else if (create_flag == false)
                                        {
                                            ws.Inset_Log(order_id, "No any host have your Vlan_id");
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.IO.File.AppendAllText(@"C:\AutoProvision\log.txt", "assign AGent error " + ex.Message + Environment.NewLine);
                                    return;
                                }
                                if (vmtype == "KVM")
                                {
                                    string[] guidf0 = vmapi.init(kvm_hostname, kvm_dsname, kvm_account, kvm_pwd).Split(':');
                                    string[] guidf1 = guidf0[2].Split('\'');
                                    guid = guidf1[1];
                                    //string[] KVM_VFT_sp0 = vmapi.createVMFromTemplate(guid, vpath, create_on_hostname, "", order_id, kvm_hostname, "", "", kvm_dsname).Split(':');
                                    //string[] KVM_VFT_sp1 = KVM_VFT_sp0[1].Split(',');
                                    //string KVM_VFT_F = KVM_VFT_sp1[0];
                                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "create VM is  : " + VMWare_VFT_F + Environment.NewLine);
                                    string KVM_VFT_F = "false";
                                    if (KVM_VFT_F == "false")
                                    {
                                        dbManager.Open();
                                        ws.Inset_Percent(order_id, "35", "");
                                        ws.set_create_on_host(order_id, kvm_hostname);
                                        vmapi.configVMCPUNum(guid, "", kvm_hostname, order_id, cpu);
                                        vmapi.configVMMemory(guid, "", kvm_hostname, order_id, ram);
                                        string nic_name1 = vmapi.getHostNetworkList(guid, vmware_datacenter_name, kvm_hostname);
                                        virNetworkName_m = Regex.Split(nic_name1, "\"message\":")[1].Replace("\"", "").Replace("{", "").Replace("[", "").Replace("}", "").Replace("]", "").Replace(")", "").Replace(":", ",").Split(',');
                                        //
                                        //vmapi.powerOffVMAsync(guid, vmware_datacenter_name, vmware_host_name, order_id);//poweroff
                                        //ws.Inset_Percent(order_id, "35", "");
                                        //ws.set_create_on_host(order_id, vmware_host_name); // update order information with host name
                                        //vmapi.configVMCPUNum(guid, vmware_datacenter_name, vmware_host_name, order_id, cpu);//修改CPU數量
                                        //vmapi.configVMMemory(guid, vmware_datacenter_name, vmware_host_name, order_id, ram);//修改RAM大小

                                        //if (hdSize != "") //remove HD and add HD
                                        //{
                                        //    vmapi.removeVMDisk(guid, vmware_datacenter_name, vmware_datastore_name, vmware_datacenter_name, order_id, "1");
                                        //    vmapi.addVMDisk(guid, vmware_datacenter_name, vmware_host_name, vmware_datastore_name, order_id, (int.Parse(hdSize) * 1000).ToString());
                                        //}


                                        //vmapi.adjustVMNIC(guid, "", kvm_hostname, order_id, "Network adapter 1", virNetworkName);

                                        VM_set_adapter(vlan_id, virNetworkName_m, guid, "", kvm_hostname, order_id);
                                        System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", "VM hardware set is Complete : " + order_id + Environment.NewLine);

                                        string[] KVM_Power_on_sp0 = vmapi.powerOnVM(guid, kvm_hostname, kvm_hostname, order_id).Split(':');
                                        string[] KVM_Power_on_sp1 = KVM_Power_on_sp0[1].Split(',');
                                        string KVM_Power_on_F = KVM_Power_on_sp1[0];

                                        if (KVM_Power_on_F == "true")
                                        {
                                            ws.set_create_on_host(order_id, kvm_hostname);
                                            RunspaceInvoke invoker = new RunspaceInvoke();
                                            invoker.Invoke("Set-ExecutionPolicy Unrestricted");
                                            System.Threading.Thread.Sleep(240000);
                                            string vm_network_name = vmapi.getVMNICNetworkList(guid, "", kvm_hostname, order_id).Split(':')[3].Split('"')[1];
                                            string vm_ip = vmapi.getVMIpAndMac(guid, "", kvm_hostname, order_id, vm_network_name).Split(':')[2].Split('"')[1];

                                            if (os == "0")//For MicroSoft
                                            {
                                                invoker.Invoke("Set-Item WSMan:\\localhost\\Client\\TrustedHosts " + vm_ip + " -Concatenate -force");
                                                string pingvm = "$username = \"User\" \n" +
                                                                "$account = \"User\" \n" +
                                                                "$password = ConvertTO-SecureString \"Passw0rd\" -asplaintext -Force \n" +
                                                                "$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $account, $password \n" +
                                                                "Invoke-Command -ComputerName " + vm_ip + "  -Authentication default -credential $cred -ScriptBlock {$FileExists = Test-Path \"c:\\AutoProvision\";If ($FileExists -eq $False){mkdir \"c:\\AutoProvision\";echo \"" + order_id + " " + group_id + ">C:\\AutoProVision\\vmname.txt;echo \"timeout 30\" `n 'msiexec /forcerestart /i  c:\\AutoProVision\\VMConfig.msi ALLUSERS=1 DB_Server=\"" + db_server + "\" EmailAccount=\"" + EmailAccount + "\" EmailPassword=\"" + EmailPassword + "\" smtphost=\"" + smtphost + "\" AutoProvision_WS_url=\"" + AutoProvision_WS_url + "\" /qn' `n \"timeout 10\" `n \"restart\" `n | Set-Content C:\\AutoProVision\\vmconfig.bat;$File = \"c:\\AutoProvision\\VMConfig.msi\";$ftp = \"" + ftp_folder + "\\VMConfig.msi\";$File3 = \"c:\\AutoProvision\\creboot2.exe\";$ftp3 = \"" + ftp_folder + "\\creboot2.exe\";$File4 = \"c:\\AutoProvision\\set_ip.exe\";$ftp4 = \"" + ftp_folder + "\\set_ip.exe\";$webclient = New-Object System.Net.WebClient;$Username = \"" + ConfigurationManager.AppSettings["ftpUsername"] + "\";$Password = \"" + ConfigurationManager.AppSettings["ftpPassword"] + "\";$webclient = New-Object System.Net.WebClient;$WebClient.Credentials = New-Object System.Net.Networkcredential($Username, $Password);$webclient.DownloadFile($ftp, $File);$webclient.DownloadFile($ftp3, $File3);$webclient.DownloadFile($ftp4, $File4);Start-Sleep -s 10 ; }c:\\AutoProvision\\vmconfig.bat;}";

                                                execute(pingvm);
                                            }
                                            if (os == "1") //For Linux
                                            {
                                                string hostSshUserName = "gcca";
                                                string hostSshPassword = "Gcca@gcca";
                                                string agentFolderPath = "/opt/AutoConfigAgent";
                                                // TODO: retrive url from outside
                                                string wsUrl = AutoProvision_WS_url;
                                                string hypervisorIpAddress = kvm_hostname;
                                                string ftpUsername = ftp_user;
                                                string ftpPassword = ftp_pwd;
                                                string agentFtpPath = "/AutoConfigAgent";
                                                string vmProvisionName = order_id;

                                                executeLinuxScript(vm_ip, hostSshUserName, hostSshPassword, 22,
                                                agentFolderPath, wsUrl, hypervisorIpAddress, ftp, ftpUsername, ftpPassword, agentFtpPath, vmProvisionName);
                                            }
                                            ws.Inset_Percent(order_id, "50", "");
                                        }
                                    }
                                }
                            }
                        }
                        #endregion

                        if (group_id == null || group_id == "" || group_id == "0")
                        {
                            #region nonGroup


                            if (nCount == 1)
                            {
                                group_id = ds.Tables[0].Rows[0]["group_id"].ToString();
                                company_id = ds.Tables[0].Rows[0]["company_id"].ToString();
                                order_id = ds.Tables[0].Rows[0]["order_id"].ToString();
                                os = ds.Tables[0].Rows[0]["os"].ToString();
                                vmtype = ds.Tables[0].Rows[0]["order_vm_type"].ToString();
                                order_area = ds.Tables[0].Rows[0]["order_area"].ToString();
                                temp_id = ds.Tables[0].Rows[0]["temp_id"].ToString();
                                order_vm_type = ds.Tables[0].Rows[0]["order_vm_type"].ToString();
                                vpath = ds.Tables[0].Rows[0]["vpath"].ToString();
                                cpu = ds.Tables[0].Rows[0]["order_cpu"].ToString();
                                hdSize = ds.Tables[0].Rows[0]["order_nhd"].ToString();
                                ram = Convert.ToString(Convert.ToInt16(ds.Tables[0].Rows[0]["order_ram"].ToString()) * 1024);
                                add_nic_num = ds.Tables[0].Rows[0]["vlan_id"].ToString().Split(',').Count() - 1;
                                vlan_id = ds.Tables[0].Rows[0]["vlan_id"].ToString().Split(',');
                                dbManager.CreateParameters(1);
                                dbManager.AddParameters(0, "order_id", order_id);
                                sql = @"UPDATE user_vm_order  
                                         SET order_audit='3',upd_datetime=getdate()  
                                         WHERE order_id=@order_id
                                        update user_vm_order
                                        set create_time=getdate()
                                        where order_id=@order_id";
                                result = dbManager.ExecuteNonQuery(CommandType.Text, sql);

                                ws.Inset_Percent(order_id, "10", "");
                                if (vmtype == "KVM")
                                {
                                    dbManager.CreateParameters(4);
                                    dbManager.AddParameters(0, "@order_area", order_area);
                                    dbManager.AddParameters(1, "@company_id", company_id);
                                    dbManager.AddParameters(2, "@vmtype", vmtype);
                                    dbManager.AddParameters(3, "@temp_id", temp_id);
                                    sql = @"select kvm_account,kvm_dsname,kvm_hostname,kvm_pwd,b.hostname as create_on_hostname
                                from config_vm_host a left outer join vm_temp b on a.vmtype=b.vm_type
                                where area=@order_area and a.company_id=@company_id and vmtype=@vmtype and temp_id=@temp_id";
                                    ds = dbManager.ExecuteDataSet(CommandType.Text, sql);
                                    nCount = ds.Tables[0].Rows.Count;
                                    if (nCount == 1)
                                    {
                                        kvm_account = ds.Tables[0].Rows[0]["kvm_account"].ToString();
                                        kvm_dsname = ds.Tables[0].Rows[0]["kvm_dsname"].ToString();
                                        kvm_hostname = ds.Tables[0].Rows[0]["kvm_hostname"].ToString();
                                        kvm_pwd = CryptoAES.decrypt(ds.Tables[0].Rows[0]["kvm_pwd"].ToString(), "GccA@stanchengGg");
                                        create_on_hostname = ds.Tables[0].Rows[0]["create_on_hostname"].ToString();
                                        //string kvm_usage = vm_useage.get_KVM_HOST_Usage(kvm_hostname, kvm_account, kvm_pwd);
                                        //string kvm_vm_usage = vm_useage.get_KVM_VM_Usage(kvm_hostname, kvm_dsname, kvm_account, kvm_pwd);

                                    }
                                    #region MyRegion
                                    //else if (nCount > 1)
                                    //{
                                    //   


                                    //    int kvm_memory1 = 0;
                                    //    int kvm_memory2 = 0;
                                    //    int kvm_vm_cpu1 = 0;
                                    //    int kvm_vm_cpu2 = 0;
                                    //    for (int i = 0; i < nCount; i++)
                                    //    {
                                    //        string kvm_vm_usage_memory = "";
                                    //        string kvm_useage = "";
                                    //        string kvm_account2 = "";
                                    //        string kvm_dsname2 = "";
                                    //        string kvm_vm_usage_memory2 = "";
                                    //        string kvm_hostname2 = "";
                                    //        string kvm_pwd2 = "";
                                    //        string kvm_useage2 = "";
                                    //        string create_on_hostname2 = "";
                                    //        if (i == 0)
                                    //        {
                                    //            kvm_account = ds.Tables[0].Rows[i]["kvm_account"].ToString();
                                    //            kvm_dsname = ds.Tables[0].Rows[i]["kvm_dsname"].ToString();
                                    //            kvm_hostname = ds.Tables[0].Rows[i]["kvm_hostname"].ToString();
                                    //            kvm_pwd = CryptoAES.decrypt(ds.Tables[0].Rows[0]["kvm_pwd"].ToString(), "GccA@stanchengGg");
                                    //            create_on_hostname = ds.Tables[0].Rows[0]["create_on_hostname"].ToString();
                                    //            kvm_useage = ws.get_KVM_HOST_Usage(kvm_hostname, kvm_account, kvm_pwd);
                                    //            kvm_vm_usage_memory = ws.get_KVM_VM_Usage(kvm_hostname, kvm_dsname, kvm_account, kvm_pwd);
                                    //            string kvm_host_memory = kvm_useage.Split(':')[1].Split(']')[0].Trim();
                                    //            string kvm_vm_cpu = kvm_vm_usage_memory.Split(':')[2].Trim();
                                    //            string kvm_vm_memory = kvm_vm_usage_memory.Split(':')[1].Split(' ')[1].ToString().Trim();
                                    //            kvm_memory1 = (Convert.ToInt32(kvm_host_memory)) - (Convert.ToInt32(kvm_vm_memory));
                                    //            kvm_vm_cpu1 = (Convert.ToInt32(kvm_vm_cpu));
                                    //        }
                                    //        else if (i > 0)
                                    //        {
                                    //            kvm_account2 = ds.Tables[0].Rows[i]["kvm_account"].ToString();
                                    //            kvm_dsname2 = ds.Tables[0].Rows[i]["kvm_dsname"].ToString();
                                    //            kvm_hostname2 = ds.Tables[0].Rows[i]["kvm_hostname"].ToString();
                                    //            kvm_pwd2 = CryptoAES.decrypt(ds.Tables[0].Rows[0]["kvm_pwd"].ToString(), "GccA@stanchengGg");
                                    //            create_on_hostname2 = ds.Tables[0].Rows[0]["create_on_hostname"].ToString();
                                    //            kvm_useage2 = ws.get_KVM_HOST_Usage(kvm_hostname2, kvm_account2, kvm_pwd2);
                                    //            kvm_vm_usage_memory2 = ws.get_KVM_VM_Usage(kvm_hostname2, kvm_dsname2, kvm_account2, kvm_pwd2);
                                    //            string kvm_host_memory = kvm_useage2.Split(':')[1].Split(']')[0].Trim();
                                    //            string kvm_vm_cpu = kvm_vm_usage_memory2.Split(':')[2].Trim();
                                    //            string kvm_vm_memory = kvm_vm_usage_memory2.Split(':')[1].Split(' ')[1].ToString().Trim();
                                    //            kvm_memory2 = (Convert.ToInt32(kvm_host_memory)) - (Convert.ToInt32(kvm_vm_memory));
                                    //            kvm_vm_cpu2 = (Convert.ToInt32(kvm_vm_cpu));
                                    //            if (kvm_memory1 < kvm_memory2)
                                    //            {
                                    //                kvm_account = kvm_account2;
                                    //                kvm_dsname = kvm_dsname2;
                                    //                kvm_hostname = kvm_hostname2;
                                    //                kvm_pwd = kvm_pwd2;
                                    //                create_on_hostname = create_on_hostname2;
                                    //                kvm_memory1 = kvm_memory2;
                                    //            }
                                    //            else if (kvm_memory1 == kvm_memory2)
                                    //            {
                                    //                if (kvm_vm_cpu1 < kvm_vm_cpu2)
                                    //                {
                                    //                    kvm_account = kvm_account2;
                                    //                    kvm_dsname = kvm_dsname2;
                                    //                    kvm_hostname = kvm_hostname2;
                                    //                    kvm_pwd = kvm_pwd2;
                                    //                    create_on_hostname = create_on_hostname2;
                                    //                    kvm_memory1 = kvm_memory2;
                                    //                }
                                    //            }
                                    //        }
                                    //    }
                                    //    
                                    //}
                                    #endregion
                                }
                                if (vmtype == "VMware")
                                {
                                    dbManager.CreateParameters(4);
                                    dbManager.AddParameters(0, "@order_area", order_area);
                                    dbManager.AddParameters(1, "@company_id", company_id);
                                    dbManager.AddParameters(2, "@vmtype", vmtype);
                                    dbManager.AddParameters(3, "@temp_id", temp_id);
                                    sql = @"select vmware_apiurl,vmware_datacenter_name,vmware_datastore_name,vmware_host_account,vmware_host_name,vmware_host_pwd,b.hostname as create_on_hostname,resource_pool_name,b.datacenter_name as create_on_datacentername,b.temp_id
                                from config_vm_host a left outer join vm_temp b on a.vmtype=b.vm_type 
                                where area=@order_area and a.company_id=@company_id and vmtype=@vmtype and b.temp_id=@temp_id";
                                    ds = dbManager.ExecuteDataSet(CommandType.Text, sql);
                                    nCount = ds.Tables[0].Rows.Count;
                                    if (nCount == 1)
                                    {
                                        vmware_apiurl = ds.Tables[0].Rows[0]["vmware_apiurl"].ToString();
                                        string vmware_vcenter_ip = vmware_apiurl.Split('/')[2];
                                        vmware_datacenter_name = ds.Tables[0].Rows[0]["vmware_datacenter_name"].ToString();
                                        vmware_datastore_name = ds.Tables[0].Rows[0]["vmware_datastore_name"].ToString();
                                        vmware_host_account = ds.Tables[0].Rows[0]["vmware_host_account"].ToString();
                                        vmware_host_name = ds.Tables[0].Rows[0]["vmware_host_name"].ToString();
                                        vmware_host_pwd = CryptoAES.decrypt(ds.Tables[0].Rows[0]["vmware_host_pwd"].ToString(), "GccA@stanchengGg");
                                        vmware_host_encryp_pwd = ds.Tables[0].Rows[0]["vmware_host_pwd"].ToString();
                                        create_on_hostname = ds.Tables[0].Rows[0]["create_on_hostname"].ToString();
                                        create_on_datacentername = ds.Tables[0].Rows[0]["create_on_datacentername"].ToString();
                                        resource_pool_name = Convert.ToString(ds.Tables[0].Rows[0]["resource_pool_name"]);
                                        //string vmware_usage = vm_useage.get_VMware_HOST_Usage(vmware_vcenter_ip, vmware_host_name, vmware_host_account, vmware_host_pwd);
                                        //string vmware_vm_usage = vm_useage.get_VMware_VM_Usage(vmware_apiurl,vmware_host_account,vmware_host_pwd,vmware_datacenter_name,vmware_host_name);

                                        /////////06/19  add
                                        string[] guidf0 = vmapi.init(vmware_apiurl, vmware_host_account, vmware_host_pwd, vmware_datacenter_name, true).Split(':');
                                        string[] guidf1 = guidf0[2].Split('\'');
                                        guid = guidf1[1];
                                        string nic_name1 = vmapi.getHostNetworkList(guid, vmware_datacenter_name, vmware_host_name);
                                        Int32 vm_count = Regex.Split(nic_name1, "\"message\":")[1].Split(',').Count() / 2;
                                        for (int vc = 1; vc <= vm_count; vc++)
                                        {
                                            if (exist_group_flag != true)
                                            {
                                                for (int nic_count = 0; nic_count <= add_nic_num; nic_count++)
                                                {
                                                    if (vlan_id[nic_count] != (Regex.Split(nic_name1, "\"message\":")[1].Split('[')[1].Split(']')[0].Split('{')[vc].Split('}')[0].Split(',')[1]).Split('\"')[3])
                                                    {
                                                        exist_group_flag = false;
                                                    }
                                                    else
                                                    {
                                                        string vlan_id1 = vlan_id[nic_count];
                                                        virNetworkName = Regex.Split(nic_name1, "\"message\":")[1].Split('[')[1].Split(']')[0].Split('{')[vc].Split('}')[0].Split(',')[0].Split('\"')[3];
                                                        virNetworkName_m = Regex.Split(nic_name1, "\"message\":")[1].Replace("\"", "").Replace("{", "").Replace("[", "").Replace("}", "").Replace("]", "").Replace(")", "").Replace(":", ",").Split(',');
                                                        int count = 0;
                                                        virNetworkName_st[] host_nic_and_vlan = new virNetworkName_st[virNetworkName_m.Count() / 4];
                                                        for (int ii = 3; ii <= virNetworkName_m.Count(); ii = ii + 4)
                                                        {
                                                            host_nic_and_vlan[count].vlan_id = virNetworkName_m[ii];
                                                            host_nic_and_vlan[count].network_name = virNetworkName_m[ii - 2];

                                                            count++;

                                                            //if (ii != virNetworkName_m.Count()-1)
                                                            //{
                                                            //    vlan_id2 += "[" + virNetworkName_m[ii] + "," + virNetworkName_m[ii - 2] + "]" + ",";
                                                            //}
                                                            //else 
                                                            //{
                                                            //    vlan_id2 += "[" + virNetworkName_m[ii] + "," + virNetworkName_m[ii - 2] + "]";
                                                            //}
                                                        }
                                                        for (int iii = 0; iii < vlan_id.Count(); iii++)
                                                        {
                                                            for (int ii = 0; ii < host_nic_and_vlan.Count(); ii++)
                                                            {
                                                                if (vlan_id[iii] == host_nic_and_vlan[ii].vlan_id)
                                                                {
                                                                    vlan_id[iii] = host_nic_and_vlan[ii].vlan_id;
                                                                }
                                                            }
                                                        }
                                                        exist_group_flag = true;
                                                        create_flag = true;
                                                    }
                                                }
                                            }
                                        }
                                        /////////06/19  add
                                    }
                                    //else if (nCount > 1)
                                    //{
                                    //    int vmware_memory1 = 0;
                                    //    int vmware_memory2 = 0;
                                    //    int vmware_vm_cpu1 = 0;
                                    //    int vmware_vm_cpu2 = 0;
                                    //    for (int i = 0; i < nCount; i++)
                                    //    {
                                    //        string vmware_vm_usage_memory = "";
                                    //        string vmware_useage = "";
                                    //        string vmware_vm_usage_memory2 = "";
                                    //        string vmware_host_account2 = "";
                                    //        string vmware_datastore_name2 = "";
                                    //        string vmware_datacenter_name2 = "";
                                    //        string vmware_apiurl2 = "";
                                    //        string vmware_host_name2 = "";
                                    //        string vmware_host_pwd2 = "";
                                    //        string vmware_useage2 = "";
                                    //        string vmware_host_encryp_pwd2 = "";
                                    //        string create_on_hostname2 = "";
                                    //        string create_on_datacentername2 = "";
                                    //        string resource_pool_name2 = "";
                                    //        if (i == 0)
                                    //        {
                                    //            vmware_host_account = ds.Tables[0].Rows[i]["vmware_host_account"].ToString();
                                    //            vmware_datastore_name = ds.Tables[0].Rows[i]["vmware_datastore_name"].ToString();
                                    //            vmware_datacenter_name = ds.Tables[0].Rows[i]["vmware_datacenter_name"].ToString();
                                    //            vmware_apiurl = ds.Tables[0].Rows[i]["vmware_apiurl"].ToString();
                                    //            string vmware_vcenter_ip = vmware_apiurl.Split('/')[2];
                                    //            vmware_host_name = ds.Tables[0].Rows[i]["vmware_host_name"].ToString();
                                    //            vmware_host_pwd = CryptoAES.decrypt(ds.Tables[0].Rows[i]["vmware_host_pwd"].ToString(), "GccA@stanchengGg");
                                    //            vmware_host_encryp_pwd = ds.Tables[0].Rows[i]["vmware_host_pwd"].ToString();
                                    //            create_on_hostname = ds.Tables[0].Rows[0]["create_on_hostname"].ToString();
                                    //            create_on_datacentername = ds.Tables[0].Rows[0]["create_on_datacentername"].ToString();
                                    //            resource_pool_name = Convert.ToString(ds.Tables[0].Rows[0]["resource_pool_name"]);
                                    //            vmware_useage = ws.get_VMware_HOST_Usage(vmware_vcenter_ip, vmware_host_name, vmware_host_account, vmware_host_pwd);
                                    //            vmware_vm_usage_memory = ws.get_VMware_VM_Usage(vmware_apiurl, vmware_host_account, vmware_host_pwd, vmware_datacenter_name, vmware_host_name);
                                    //            string vmware_host_memory = vmware_useage.Split(':')[1].Split(']')[0].Trim();
                                    //            string vmware_vm_cpu = vmware_vm_usage_memory.Split(':')[2].Trim();
                                    //            string vmware_vm_memory = vmware_vm_usage_memory.Split(':')[1].Split(' ')[1].Split('G')[0].Trim();
                                    //            vmware_memory1 = (Convert.ToInt32(vmware_host_memory)) - (Convert.ToInt32(vmware_vm_memory));
                                    //            vmware_vm_cpu1 = (Convert.ToInt32(vmware_vm_cpu));
                                    //        }
                                    //        //else if (i > 0)
                                    //        //{
                                    //        //    vmware_host_account2 = ds.Tables[0].Rows[i]["vmware_host_account"].ToString();
                                    //        //    vmware_datastore_name2 = ds.Tables[0].Rows[i]["vmware_datastore_name"].ToString();
                                    //        //    vmware_datacenter_name2 = ds.Tables[0].Rows[i]["vmware_datacenter_name"].ToString();
                                    //        //    vmware_apiurl2 = ds.Tables[0].Rows[i]["vmware_apiurl"].ToString();
                                    //        //    string vmware_vcenter_ip = vmware_apiurl.Split('/')[2];
                                    //        //    vmware_host_name2 = ds.Tables[0].Rows[i]["vmware_host_name"].ToString();
                                    //        //    vmware_host_pwd2 = CryptoAES.decrypt(ds.Tables[0].Rows[i]["vmware_host_pwd"].ToString(), "GccA@stanchengGg");
                                    //        //    vmware_host_encryp_pwd2 = ds.Tables[0].Rows[i]["vmware_host_pwd"].ToString();
                                    //        //    create_on_hostname2 = ds.Tables[0].Rows[0]["create_on_hostname"].ToString();
                                    //        //    create_on_datacentername2 = ds.Tables[0].Rows[0]["create_on_datacentername"].ToString();
                                    //        //    resource_pool_name2 = Convert.ToString(ds.Tables[0].Rows[0]["resource_pool_name"]);
                                    //        //    vmware_useage2 = ws.get_VMware_HOST_Usage(vmware_vcenter_ip, vmware_host_name2, vmware_host_account2, vmware_host_pwd2);
                                    //        //    vmware_vm_usage_memory2 = ws.get_VMware_VM_Usage(vmware_apiurl2, vmware_host_account2, vmware_host_pwd2, vmware_datacenter_name2, vmware_host_name2);
                                    //        //    string vmware_host_memory = vmware_useage2.Split(':')[1].Split(']')[0].Trim();
                                    //        //    string vmware_vm_cpu = vmware_vm_usage_memory2.Split(':')[2].Trim();
                                    //        //    string vmware_vm_memory = vmware_vm_usage_memory2.Split(':')[1].Split(' ')[1].Split('G')[0].Trim();
                                    //        //    vmware_memory2 = (Convert.ToInt32(vmware_host_memory)) - (Convert.ToInt32(vmware_vm_memory));
                                    //        //    vmware_vm_cpu2 = (Convert.ToInt32(vmware_vm_cpu));
                                    //        //    if (vmware_memory1 < vmware_memory2)
                                    //        //    {
                                    //        //        vmware_host_account = vmware_host_account2;
                                    //        //        vmware_datastore_name = vmware_datastore_name2;
                                    //        //        vmware_datacenter_name = vmware_datacenter_name2;
                                    //        //        vmware_apiurl = vmware_apiurl2;
                                    //        //        vmware_host_name = vmware_host_name2;
                                    //        //        vmware_host_pwd = vmware_host_pwd2;
                                    //        //        vmware_host_encryp_pwd = vmware_host_encryp_pwd2;
                                    //        //        create_on_hostname = create_on_hostname2;
                                    //        //        create_on_datacentername = create_on_datacentername2;
                                    //        //        resource_pool_name = resource_pool_name2;
                                    //        //        vmware_memory1 = vmware_memory2;
                                    //        //    }
                                    //        //    else if (vmware_memory1 == vmware_memory2)
                                    //        //    {
                                    //        //        if (vmware_vm_cpu1 < vmware_vm_cpu2)
                                    //        //        {
                                    //        //            vmware_host_account = vmware_host_account2;
                                    //        //            vmware_datastore_name = vmware_datastore_name2;
                                    //        //            vmware_datacenter_name = vmware_datacenter_name2;
                                    //        //            vmware_apiurl = vmware_apiurl2;
                                    //        //            vmware_host_name = vmware_host_name2;
                                    //        //            vmware_host_pwd = vmware_host_pwd2;
                                    //        //            vmware_host_encryp_pwd = vmware_host_encryp_pwd2;
                                    //        //            create_on_hostname = create_on_hostname2;
                                    //        //            create_on_datacentername = create_on_datacentername2;
                                    //        //            resource_pool_name = resource_pool_name2;
                                    //        //            vmware_memory1 = vmware_memory2;
                                    //        //        }
                                    //        //    }
                                    //        //}
                                    //    }
                                    //}

                                }


                                ws.Inset_Percent(order_id, "20", "");

                                vmapi.Timeout = 10000000;
                                if (vmtype == "VMware")
                                {
                                    string[] vcenter_ip_sp1 = vmware_apiurl.Split('/');
                                    string vcenter_ip = vcenter_ip_sp1[2];
                                    string[] guidf0 = vmapi.init(vmware_apiurl, vmware_host_account, vmware_host_pwd, vmware_datacenter_name, true).Split(':');
                                    string[] guidf1 = guidf0[2].Split('\'');
                                    guid = guidf1[1];
                                    //string[] VMWare_VFT_sp0 = vmapi.createVMFromTemplate(guid, vpath, create_on_hostname, create_on_datacentername, order_id, vmware_host_name, vmware_datacenter_name, resource_pool_name, vmware_datastore_name).Split(':');
                                    //string[] VMWare_VFT_sp1 = VMWare_VFT_sp0[1].Split(',');
                                    string VMWare_VFT_F = "false";// VMWare_VFT_sp1[0];
                                    if (VMWare_VFT_F == "false")
                                    {

                                        //ws.Inset_Percent(order_id, "35", "");
                                        //ws.set_create_on_host(order_id, vmware_host_name);
                                        //vmapi.configVMCPUNum(guid, vmware_datacenter_name, vmware_host_name, order_id, cpu);
                                        ////vmapi.configVMMemory(guid, vmware_datacenter_name, vmware_host_name, order_id, ram);
                                        //if (hdSize != "") //remove HD and add HD
                                        //{
                                        //    vmapi.removeVMDisk(guid, vmware_datacenter_name, vmware_datastore_name, vmware_datacenter_name, order_id, "1");
                                        //    vmapi.addVMDisk(guid, vmware_datacenter_name, vmware_host_name, vmware_datastore_name, order_id, (int.Parse(hdSize) * 1000).ToString());
                                        //}
                                        //VM_set_adapter(vlan_id, virNetworkName_m, guid, vmware_datacenter_name, vmware_host_name, order_id); //暫時 FOR VMWARE   setting adapter. write by jedi

                                        string[] VMWare_Power_on_sp0 = vmapi.powerOnVM(guid, vmware_datacenter_name, vmware_host_name, order_id).Split(':');
                                        string[] VMWare_Power_on_sp1 = VMWare_Power_on_sp0[1].Split(',');
                                        string VMWare_Power_on_F = VMWare_Power_on_sp1[0];
                                        if (VMWare_Power_on_F == "true")
                                        {
                                            ws.set_create_on_host(order_id, vmware_host_name);
                                            RunspaceInvoke invoker = new RunspaceInvoke();
                                            invoker.Invoke("Set-ExecutionPolicy Unrestricted");
                                            System.Threading.Thread.Sleep(1000);
                                            string[] url0 = vmware_apiurl.Split('/');
                                            vmware_apiurl = url0[2];
                                            //System.Threading.Thread.Sleep(240000);
                                            string vm_network_name = vmapi.getVMNICNetworkList(guid, vmware_datacenter_name, vmware_host_name, order_id).Split(':')[3].Split('"')[1];
                                            string vm_ip = vmapi.getVMIpAndMac(guid, vmware_datacenter_name, vmware_host_name, order_id, vm_network_name).Split(':')[2].Split('"')[1];

                                            if (os == "0")//For MicroSoft
                                            {
                                                System.Threading.Thread.Sleep(1000);
                                                invoker.Invoke("Set-Item WSMan:\\localhost\\Client\\TrustedHosts " + vm_ip + " -Concatenate -force");
                                                // TODO: do not write account/password into the script
                                                string pingvm = @"$username = ""User""
                                                                $account = """ + vm_account + @"""
                                                                $password = ConvertTO-SecureString """ + vm_password + @""" -asplaintext -Force
                                                                $cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $account, $password
                                                                Invoke-Command -ComputerName " + vm_ip + @" -Authentication default" +
                                                                @" -credential $cred " +
                                                                @" -ScriptBlock {" +
                                                                    @"$FileExists = Test-Path ""c:\\AutoProvision"";" +
                                                                    @"If ($FileExists -eq $False){" +
                                                                        @"mkdir ""c:\\AutoProvision"";"
                                                                        + @"echo """ + order_id + " " + group_id + @""" >C:\\AutoProVision\\vmname.txt;"
                                                                        + @"echo ""timeout 30"""
                                                                        + @" `n 'msiexec /forcerestart /i  c:\AutoProVision\VMConfig.msi ALLUSERS=1 DB_Server=""" + db_server + @""" EmailAccount=""" + EmailAccount + "\" EmailPassword=\"" + EmailPassword + "\" smtphost=\"" + smtphost + "\" AutoProvision_WS_url=\"" + AutoProvision_WS_url + "\" /qn'"
                                                                        + @" `n ""timeout 10"" "
                                                                        + @"`n ""restart"" "
                                                                        + @"`n | "
                                                                        + @"Set-Content C:\AutoProVision\vmconfig.bat;"
                                                                        + @"$File = ""c:\AutoProvision\VMConfig.msi"";"
                                                                        + @"$ftp = """ + ftp_folder + "\\VMConfig.msi\";"
                                                                        + @"$File3 = ""c:\AutoProvision\creboot2.exe"";"
                                                                        + @"$ftp3 = """ + ftp_folder + @"\creboot2.exe"";"
                                                                        + @"$File4 = ""c:\AutoProvision\set_ip.exe"";"
                                                                        + @"$ftp4 = """ + ftp_folder + @"\set_ip.exe"";"
                                                                        + @"$webclient = New-Object System.Net.WebClient;"
                                                                        + @"$Username = """ + ConfigurationManager.AppSettings["ftpUsername"] + "\";"
                                                                        + @"$Password = """ + ConfigurationManager.AppSettings["ftpPassword"] + "\";"
                                                                        + @"$webclient = New-Object System.Net.WebClient;"
                                                                        + @"$WebClient.Credentials = New-Object System.Net.Networkcredential($Username, $Password);"
                                                                        + @"$webclient.DownloadFile($ftp, $File);"
                                                                        + @"$webclient.DownloadFile($ftp3, $File3);"
                                                                        + @"$webclient.DownloadFile($ftp4, $File4);"
                                                                        + @"Start-Sleep -s 10 ; "
                                                                    + @"}"
                                                                    + @"c:\AutoProvision\vmconfig.bat;"
                                                                + "}";
                                                //使用POWERSHELL 連至VM並對此台VM下達Create一個vmname.txt內容放入訂單編號及群組ID，下載FTP上之AGENT的檔案安裝完AGENT後並重開機。
                                                execute(pingvm);
                                            }
                                            if (os == "1") //For Linux
                                            {
                                                string hostSshUserName = "gcca";
                                                string hostSshPassword = "Gcca@gcca";
                                                string agentFolderPath = "/opt/AutoConfigAgent";
                                                // TODO: retrive url from outside
                                                string wsUrl = AutoProvision_WS_url;
                                                string hypervisorIpAddress = vmware_host_name;
                                                string ftpUsername = ftp_user;
                                                string ftpPassword = ftp_pwd;
                                                string agentFtpPath = "/AutoConfigAgent";
                                                string vmProvisionName = order_id;

                                                executeLinuxScript(vm_ip, hostSshUserName, hostSshPassword, 22,
                                                agentFolderPath, wsUrl, hypervisorIpAddress, ftp, ftpUsername, ftpPassword, agentFtpPath, vmProvisionName);
                                            }
                                            ws.Inset_Percent(order_id, "50", "");
                                        }
                                    }
                                }

                                if (vmtype == "KVM")
                                {
                                    string[] guidf0 = vmapi.init(kvm_hostname, kvm_dsname, kvm_account, kvm_pwd).Split(':');
                                    string[] guidf1 = guidf0[2].Split('\'');
                                    guid = guidf1[1];
                                    string[] KVM_VFT_sp0 = vmapi.createVMFromTemplate(guid, vpath, create_on_hostname, "", order_id, kvm_hostname, "", "", kvm_dsname).Split(':');
                                    string[] KVM_VFT_sp1 = KVM_VFT_sp0[1].Split(',');
                                    string KVM_VFT_F = KVM_VFT_sp1[0];
                                    if (KVM_VFT_F == "true")
                                    {
                                        dbManager.Open();
                                        ws.Inset_Percent(order_id, "35", "");
                                        ws.set_create_on_host(order_id, vmware_host_name);
                                        vmapi.configVMCPUNum(guid, "", kvm_hostname, order_id, cpu);
                                        vmapi.configVMMemory(guid, "", kvm_hostname, order_id, ram);
                                        vmapi.adjustVMNIC(guid, "", kvm_hostname, order_id, "Network adapter 1", virNetworkName);
                                        string[] KVM_Power_on_sp0 = vmapi.powerOnVM(guid, "", kvm_hostname, order_id).Split(':');
                                        string[] KVM_Power_on_sp1 = KVM_Power_on_sp0[1].Split(',');
                                        string KVM_Power_on_F = KVM_Power_on_sp1[0];

                                        if (KVM_Power_on_F == "true")
                                        {
                                            ws.set_create_on_host(order_id, kvm_hostname);
                                            RunspaceInvoke invoker = new RunspaceInvoke();
                                            invoker.Invoke("Set-ExecutionPolicy Unrestricted");
                                            System.Threading.Thread.Sleep(240000);
                                            string vm_network_name = vmapi.getVMNICNetworkList(guid, "", kvm_hostname, order_id).Split(':')[3].Split('"')[1];
                                            string vm_ip = vmapi.getVMIpAndMac(guid, "", kvm_hostname, order_id, vm_network_name).Split(':')[2].Split('"')[1];
                                            if (os == "0")//For MicroSoft
                                            {
                                                invoker.Invoke("Set-Item WSMan:\\localhost\\Client\\TrustedHosts " + vm_ip + " -Concatenate -force");
                                                string pingvm = "$username = \"User\" \n" +
                                                                "$account = \"User\" \n" +
                                                                "$password = ConvertTO-SecureString \"Passw0rd\" -asplaintext -Force \n" +
                                                                "$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $account, $password \n" +
                                                                "Invoke-Command -ComputerName " + vm_ip + "  -Authentication default -credential $cred -ScriptBlock {$FileExists = Test-Path \"c:\\AutoProvision\";If ($FileExists -eq $False){mkdir \"c:\\AutoProvision\";echo \"" + order_id + "\">C:\\AutoProVision\\vmname.txt;echo \"timeout 30\" `n 'msiexec /forcerestart /i  c:\\AutoProVision\\VMConfig.msi ALLUSERS=1 DB_Server=\"" + db_server + "\" EmailAccount=\"" + EmailAccount + "\" EmailPassword=\"" + EmailPassword + "\" smtphost=\"" + smtphost + "\" AutoProvision_WS_url=\"" + AutoProvision_WS_url + "\" /qn' `n \"timeout 10\" `n \"restart\" `n | Set-Content C:\\AutoProVision\\vmconfig.bat;$File = \"c:\\AutoProvision\\VMConfig.msi\";$ftp = \"" + ftp_folder + "\\VMConfig.msi\";$File3 = \"c:\\AutoProvision\\creboot2.exe\";$ftp3 = \"" + ftp_folder + "\\creboot2.exe\";$File4 = \"c:\\AutoProvision\\set_ip.exe\";$ftp4 = \"" + ftp_folder + "\\set_ip.exe\";$webclient = New-Object System.Net.WebClient;$Username = \"" + ConfigurationManager.AppSettings["ftpUsername"] + "\";$Password = \"" + ConfigurationManager.AppSettings["ftpPassword"] + "\";$webclient = New-Object System.Net.WebClient;$WebClient.Credentials = New-Object System.Net.Networkcredential($Username, $Password);$webclient.DownloadFile($ftp, $File);$webclient.DownloadFile($ftp3, $File3);$webclient.DownloadFile($ftp4, $File4);Start-Sleep -s 10 ; }c:\\AutoProvision\\vmconfig.bat;}";
                                                execute(pingvm);
                                            }
                                            if (os == "1") //For Linux
                                            {
                                                string hostSshUserName = "gcca";
                                                string hostSshPassword = "Gcca@gcca";
                                                string agentFolderPath = "/opt/AutoConfigAgent";
                                                // TODO: retrive url from outside
                                                string wsUrl = AutoProvision_WS_url;
                                                string hypervisorIpAddress = kvm_hostname;
                                                string ftpUsername = ftp_user;
                                                string ftpPassword = ftp_pwd;
                                                string agentFtpPath = "/AutoConfigAgent";
                                                string vmProvisionName = order_id;

                                                executeLinuxScript(vm_ip, hostSshUserName, hostSshPassword, 22,
                                                agentFolderPath, wsUrl, hypervisorIpAddress, ftp, ftpUsername, ftpPassword, agentFtpPath, vmProvisionName);
                                            }
                                            ws.Inset_Percent(order_id, "50", "");
                                        }
                                    }

                                }
                            }
                            #endregion
                        }
                    }
                    catch (Exception)
                    {
                        dbManager.Dispose();
                        ws.Dispose();
                    }
                    finally
                    {
                        dbManager.Dispose();
                        ws.Dispose();
                    }

                }
            }
            catch (Exception)
            {
                dbManager.Dispose();
                ws.Dispose();
            }
            finally
            {
            }
        }


        public static string execute(string script)
        {

            string output3 = "";
            Runspace runspace1 = RunspaceFactory.CreateRunspace();
            runspace1.Open();
            Pipeline pipeline = runspace1.CreatePipeline();
            pipeline.Commands.AddScript(script);
            pipeline.Commands.Add("Out-String");
            try
            {
                var result = pipeline.Invoke();
                output3 = result[0].ToString();
                string[] checkit = Regex.Split(output3, " ");
            }
            catch
            {

            }
            runspace1.Close();
            return output3;
        }

        public struct virNetworkName_st
        {
            public string vlan_id;
            public string network_name;
        }

        public static void executeLinuxScript(string hostName, string hostSshUserName, string hostSshPassword,
            int sshPort, string agentFolderPath, string wsUrl, string hypervisorIpAddress, string ftp,
            string ftpUsername, string ftpPassword, string agentFtpPath, string vmProvisionName)
        {
            GCCA.AutoProv.SSHClient.SSHClient sshClient = new GCCA.AutoProv.SSHClient.SSHClient(hostName, hostSshUserName, hostSshPassword, sshPort);

            // Connect to remote SSH host
            try
            {
                sshClient.Connect();
            }
            catch (Exception)
            {
                return;
            }

            string output = "";
            // Remove firewall policy before login to ftp
            bool result = sshClient.ExecuteShellCommand("bash -c \"sudo iptables -F\"\n", ref output);

            if (!result)
                throw new Exception("Remove firewall policy failed.");

            // Download latest agent from ftp
            result = sshClient.ExecuteShellCommand("bash -c \"sudo bash ~/download_script.sh "
                    + agentFolderPath + " " + ftp + " " + ftpUsername + " "
                    + ftpPassword + " " + agentFtpPath + "\"\n", ref output);

            if (!result)
                throw new Exception("Download latest agent from ftp failed.");

            string startAgentScript = "cd " + agentFolderPath + "\n"
                    + "sudo java -cp autoconfigagent.jar:"
                    + "jetty-client-8.1.9.v20130131.jar:"
                    + "jetty-http-8.1.9.v20130131.jar:"
                    + "jetty-io-8.1.9.v20130131.jar:"
                    + "jetty-util-8.1.9.v20130131.jar:" + "jsch-0.1.49.jar:"
                    + "json-simple-1.1.1.jar gcca.autoprov.AutoConfigAgent "
                    + wsUrl + " " + hypervisorIpAddress + " " + hostName + " "
                    + hostSshUserName + " " + hostSshPassword + " " + sshPort
                    + " " + agentFolderPath + " " + vmProvisionName + " >> /var/log/agent_log.txt " + "\n";

            // Execute java agent
            result = sshClient.ExecuteShellCommand("bash -c \"" + startAgentScript + "\"\n", ref output);

            if (!result)
                throw new Exception("Execute java agent failed.");

            // Disconnect from remote SSH host
            sshClient.Disconnect();
        }
        private static void clean_IP(string vmname)
        {
            string sql;
            DBManager dbManager = new DBManager(DataProvider.SqlServer);
            dbManager.ConnectionString = ConfigurationManager.AppSettings["SSM"].ToString();
            dbManager.Open();
            dbManager.CreateParameters(1);
            dbManager.AddParameters(0, "@order_id", vmname);
            sql = @"update c_ip_list 
                    set ip_address=(select substring(ip_address,0,case charindex(':',ip_address) when ''  then Len(ip_address)+1 else charindex(':',ip_address) end) from c_ip_list a where c_ip_list.row_id=a.row_id),order_id=NULL,used=NULL,upd_userid='Service',used_mac=NULL,upd_datetime=GETDATE()
                    where order_id =@order_id";

            int updateOderId = dbManager.ExecuteNonQuery(CommandType.Text, sql);

            dbManager.Dispose();
        } //clean 該order所有分配的IP

        static void VM_set_adapter(string[] vlan_id, string[] virNetworkName_m, string guid, string vmware_datacenter_name, string vmware_host_name, string order_id) //暫時 FOR VMWARE creat adapter
        {
            try
            {
                Create_VM_Service.AutoProvision_WS.AutoProvision_WS ws = new Create_VM_Service.AutoProvision_WS.AutoProvision_WS();
                Create_VM_Service.VMAPI.GCCA_HypervisorAPI vmapi = new GCCA_HypervisorAPI();

                #region cleaning adapter and database data
                /////cleaning adapter and database data 
                List<JToken> adapter_detail = new List<JToken>();
                string clean = vmapi.getVMNICMacList(guid, vmware_datacenter_name, vmware_host_name, order_id);
                JToken token;
                token = JToken.Parse(clean.Remove(clean.Length - 1).Remove(0, 1).Replace("\\", "\\\\"));
                token.SelectToken("message")[1].SelectToken("Mac").ToString();
                for (int i = 1; i < token.SelectToken("message").Count(); i++) // colculate adapter num
                {
                    string temp = vmapi.removeVMNIC(guid, vmware_datacenter_name, vmware_host_name, order_id, token.SelectToken("message")[i].SelectToken("Nic_Name").ToString(), token.SelectToken("message")[1].SelectToken("Mac").ToString()).ToString();
                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " removeVMNIC is  : " + temp +  + Environment.NewLine);
                }
                DBManager dbManager = new DBManager(DataProvider.SqlServer);
                dbManager.ConnectionString = ConfigurationManager.AppSettings["SSM"].ToString();
                dbManager.Open();
                dbManager.CreateParameters(1);
                dbManager.AddParameters(0, "@order_id", order_id);
                string sql = "";
                sql = @"DELETE FROM order_nic_mac_list
                    WHERE order_id = @order_id";
                int updateOderId = dbManager.ExecuteNonQuery(CommandType.Text, sql);
                System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " change DB data number  : " + updateOderId + +Environment.NewLine);
                clean_IP(order_id);
                //////////////
                #endregion

                List<string> NetworkName = new List<string>();
                int j_test = 0;//單純控制下面迴圈
                int j_result = 1;//控制下面迴圈
                for (int i = 0; i < vlan_id.Length; i++)  //  找出VM上符合訂單上VLAN_ID的網卡名稱
                {
                    while (i < vlan_id.Length)
                    {
                        j_result = string.Compare(vlan_id[i], virNetworkName_m[j_test]);
                        if (j_result == 0)
                        {
                            NetworkName.Add(virNetworkName_m[j_test - 2]);
                            System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " virNetworkName_m is  : " + virNetworkName_m[j_test - 2] + +Environment.NewLine);
                            break;
                        }
                        j_test++;
                    }
                    j_test = 0;
                    j_result = 1;
                }

                for (int i = 0; i < vlan_id.Length; i++)//如果有多張網卡  建立多張網卡
                {
                    if (i == 0)//if 只有一張網卡，更改網卡名稱 
                    { vmapi.adjustVMNIC(guid, vmware_datacenter_name, vmware_host_name, order_id, "Network adapter 1", NetworkName[i]); }//修改網卡內容 只有一張網卡時 修改第一張
                    //string ggg= vlan_id[];
                    else
                    { vmapi.addVMNIC(guid, vmware_datacenter_name, vmware_host_name, order_id, NetworkName[i]); }
                }


                string nic_mac_list = vmapi.getVMNICMacList(guid, vmware_datacenter_name, vmware_host_name, order_id);
                //string[] nic_mac_list2 = nic_mac_list.Split(new string[2]{"\"Mac\":\"","\"}"}, StringSplitOptions.RemoveEmptyEntries); //對API回傳的值做切割，
                string[] nic_mac_list2 = nic_mac_list.Split(new string[4] { "Nic_Name\":\"", "\",", "\"Mac\":\"", "\"}" }, StringSplitOptions.RemoveEmptyEntries); //字串分割  Nic_Name":"    ",    Mac":"   "}
                
                int j_i = 0;
                for (int i = 0; i < vlan_id.Length; i++)//將網卡MAC等資料 塞回DB
                {
                    j_i = 3 * i; //只要陣列中的 i*3+(1or2)的值
                    ws.save_nic_mac(order_id, nic_mac_list2[j_i + 1], nic_mac_list2[j_i + 2], vlan_id[i], vmware_host_name);  //將creat VM的資料(order_id,macID,nic_id,groupID,order_area,vlan_id,host_name) 塞回DB
                    System.IO.File.AppendAllText(@"C:\AutoProvision\logs.txt", " save_nic_mac: " + order_id + " " + nic_mac_list2[j_i + 1] + " " + nic_mac_list2[j_i + 2] + " " + vlan_id[i] + " " + vmware_host_name + " " + +Environment.NewLine);

                }
                dbManager.Dispose();
            }
            catch (Exception)
            {
                return;
            }

        }

    }

}

