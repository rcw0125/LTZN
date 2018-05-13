using System;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Security.Principal;
using System.Threading;
using System.Data.OracleClient;
using System.Collections.Generic;

namespace LTZN
{
    public delegate void MsgEventDelegate(MsgContent msg);
    public delegate void MsgDealEventDelegate(int count);

    public class MsgContent
    {
        public string MsgId = "";
        public string SendUser = "";
        public string ToUser = "";
        public string MyUser = "";
        public string MsgTyp = "";
        public DateTime SendTime;
        public string SendContent = "";
        public string ByMsg = "";
    }

    public class UDPListener
    {
        public static readonly UDPListener instance = new UDPListener();

        protected volatile bool _shouldStop = false;  //停止
        protected Thread workerThread = null;
        private ManualResetEvent allDone = new ManualResetEvent(false);

        private UDPListener()
        {
            workerThread = new Thread(DoWork);
        
        }

        /// <summary>
        /// 消息到达事发生
        /// </summary>
        public event MsgEventDelegate MsgEvent;

        /// <summary>
        /// 消息处理时发生
        /// </summary>
        public event MsgDealEventDelegate MsgDealEvent;

        private const int listenPort = 11000;


        private object lockObj = new object();

        /// <summary>
        /// 未处理理消息集合
        /// </summary>
        public List<MsgContent> UnDealMsg = new List<MsgContent>();

        public  void Start()
        {

            if (workerThread.ThreadState == ThreadState.Running || workerThread.ThreadState == ThreadState.WaitSleepJoin)
            {
                Stop();
            }
            if (workerThread.ThreadState == ThreadState.Suspended || workerThread.ThreadState == ThreadState.SuspendRequested)
            {
                workerThread.Abort();
            }
            if (workerThread.ThreadState == ThreadState.Stopped)
                workerThread = new Thread(DoWork);
            _shouldStop = false;
            workerThread.Start();
        }

        public  void Stop()
        {
            _shouldStop = true;
            allDone.Set();
            if (workerThread.ThreadState == ThreadState.Running)
                workerThread.Join();
        }

        /// <summary>
        /// 开始监听消息
        /// </summary>
        public void DoWork()
        {
            UdpClient listener = new UdpClient(listenPort);
            try
            {
                while (!_shouldStop)
                {
                    IAsyncResult r = listener.BeginReceive(new AsyncCallback(DoReceive), listener);
                    WaitHandle[] handls = new WaitHandle[2];
                    handls[0] = this.allDone;
                    handls[1] = r.AsyncWaitHandle;
                    ManualResetEvent.WaitAny(handls);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            finally
            {
                listener.Close();
            }
        }

        private void DoReceive(IAsyncResult ar)
        {
            UdpClient client = (UdpClient)ar.AsyncState;
            IPEndPoint groupEP = new IPEndPoint(IPAddress.Any, listenPort);
            byte[] bytes = client.EndReceive(ar, ref groupEP);
            string msg = Encoding.Unicode.GetString(bytes, 0, bytes.Length);
            string[] msgs = msg.Split('|');
            if (msgs.Length == 2)
            {
                string msgid = msgs[0];
                string users = msgs[1];
                if (LtznUserManager.instance.CurrentUser != null)
                {
                    List<string> userList = new List<string>();
                    userList.AddRange(users.Split(','));
                    if (userList.Contains(LtznUserManager.instance.CurrentUser.Identity.Name))
                    {
                        Thread thread = new Thread(new ParameterizedThreadStart(readMsg));
                        thread.Start(msgid);
                    }
                }
            }
        }

        /// <summary>
        /// 处理新消息
        /// </summary>
        /// <param name="id"></param>
        private void readMsg(object id)
        {
            if (MsgEvent != null)
            {
                OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
                conn.Open();
                string sql = "select SENDUSER,SENDTIME,TYP,CONTENT from LTMSG where ID=:MSGID";
                OracleCommand cmd = new OracleCommand(sql, conn);
                cmd.Parameters.Add(":MSGID", OracleType.VarChar).Value = (string)id;
                OracleDataReader dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    MsgContent mc = new MsgContent();
                    mc.MsgId = (string)id;
                    mc.SendUser = dr.IsDBNull(0) ? "" : dr.GetString(0);
                    mc.SendTime = dr.IsDBNull(1) ? DateTime.MinValue : dr.GetDateTime(1);
                    mc.MsgTyp = dr.IsDBNull(2) ? "" : dr.GetString(2);
                    mc.SendContent = dr.IsDBNull(3) ? "" : dr.GetString(3);
                    if (LtznUserManager.instance.CurrentUser != null)
                    {
                        mc.MyUser = LtznUserManager.instance.CurrentUser.Identity.Name;
                        addUnDealMsg(mc);
                    }
                    else
                        mc.MyUser = "";
                    MsgEvent(mc);
                }
                dr.Close();
                conn.Close();
            }
        }


        /// <summary>
        /// 添加新消息
        /// </summary>
        /// <param name="msg"></param>
        public void addUnDealMsg(MsgContent msg)
        {
            lock (lockObj)
            {
                UnDealMsg.Add(msg);
            }
            if (MsgDealEvent != null)
            {
                MsgDealEvent(UnDealMsg.Count);
            }
        }


        /// <summary>
        /// 处理消息
        /// </summary>
        /// <param name="msg"></param>
        public void DealMsg(MsgContent msg)
        {
            lock (lockObj)
            {
                OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
                conn.Open();
                string sql = "Update LTUSERMSG Set Chk=1 Where USERNAME=:USERNAME and MSGID=:MSGID";
                OracleCommand cmd = new OracleCommand(sql, conn);
                cmd.Parameters.Add(":USERNAME", OracleType.VarChar).Value = msg.MyUser;
                cmd.Parameters.Add(":MSGID", OracleType.VarChar).Value = msg.MsgId;
                cmd.ExecuteNonQuery();
                conn.Close();
                UnDealMsg.Remove(msg);
            }
            if (MsgDealEvent != null)
            {
                MsgDealEvent(UnDealMsg.Count);
            }
        }


        /// <summary>
        /// 重置未处理消息集合
        /// </summary>
        /// <param name="msgCollection"></param>
        public void resetUnDealMsg()
        {
            List<MsgContent> msgCollection = new List<MsgContent>();
            if (LtznUserManager.instance.CurrentUser != null)
            {
                OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
                conn.Open();
                string sql = "select LTMSG.ID,SENDUSER,SENDTIME,TYP,CONTENT,LTUSERMSG.USERNAME from LTMSG,LTUSERMSG where LTMSG.ID=LTUSERMSG.MSGID and LTUSERMSG.USERNAME=:USERNAME and LTUSERMSG.CHK=0";
                OracleCommand cmd = new OracleCommand(sql, conn);
                cmd.Parameters.Add(":USERNAME", OracleType.VarChar).Value = LtznUserManager.instance.CurrentUser.Identity.Name;
                OracleDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    MsgContent mc = new MsgContent();
                    mc.MsgId = dr.IsDBNull(0) ? "" : dr.GetString(0);
                    mc.SendUser = dr.IsDBNull(1) ? "" : dr.GetString(1);
                    mc.SendTime = dr.IsDBNull(2) ? DateTime.MinValue : dr.GetDateTime(2);
                    mc.MsgTyp = dr.IsDBNull(3) ? "" : dr.GetString(3);
                    mc.SendContent = dr.IsDBNull(4) ? "" : dr.GetString(4);
                    mc.MyUser = dr.IsDBNull(5) ? "" : dr.GetString(5);
                    msgCollection.Add(mc);
                }
                dr.Close();
                conn.Close();
            }
            lock (lockObj)
            {
                UnDealMsg = msgCollection;
            }
            if (MsgDealEvent != null)
            {
                MsgDealEvent(UnDealMsg.Count);
            }
        }

        /// <summary>
        ///  发送消息
        /// </summary>
        /// <param name="msg"></param>
        public void sendMsg(MsgContent msg)
        {
            Guid msgGuid= Guid.NewGuid();
            msg.MsgId = msgGuid.ToString();

            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "Insert into LTMSG(ID,SENDUSER,TOUSER,TYP,SENDTIME,CONTENT,BYMSG) values(:MSGID,:SendUser,:ToUser,:Typ,sysdate,:Content,:ByMsg)";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":MSGID", OracleType.VarChar).Value = msg.MsgId;
            cmd.Parameters.Add(":SendUser", OracleType.VarChar).Value = msg.SendUser;
            cmd.Parameters.Add(":ToUser", OracleType.VarChar).Value =msg.ToUser;
            cmd.Parameters.Add(":Typ", OracleType.VarChar).Value = msg.MsgTyp;
            cmd.Parameters.Add(":Content", OracleType.VarChar).Value = msg.SendContent;
            cmd.Parameters.Add(":ByMsg", OracleType.VarChar).Value = msg.ByMsg;

            cmd.ExecuteNonQuery();

            List<string> userList = new List<string>();
            userList.AddRange(msg.ToUser.Split(','));

            foreach (string user in userList)
            {
                OracleCommand cmdUserMsg = new OracleCommand("Insert into LTUSERMSG(USERNAME,MSGID) values(:USERNAME,:MSGID) ", conn);
                cmdUserMsg.Parameters.Add(":USERNAME", OracleType.VarChar).Value = user;
                cmdUserMsg.Parameters.Add(":MSGID", OracleType.VarChar).Value = msg.MsgId;
                cmdUserMsg.ExecuteNonQuery();
            }
            conn.Close();

            Socket s = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp);
            List<string> broadIps = new List<string>();

            IPAddress localIp = IPAddress.Parse("192.168.38.59");

            string[] ipStr = localIp.ToString().Split('.');
            string broadIpStr = ipStr[0] + "." + ipStr[1] + "." + ipStr[2] + ".255";
            if (!broadIps.Contains(broadIpStr))
            {
                broadIps.Add(broadIpStr);
                IPAddress broadcast = IPAddress.Parse(broadIpStr);
                byte[] sendbuf = Encoding.Unicode.GetBytes(msg.MsgId + '|' + msg.ToUser);
                IPEndPoint ep = new IPEndPoint(broadcast, listenPort);
                s.SendTo(sendbuf, ep);
            }
        }
    }
}
