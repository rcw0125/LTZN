using System;
using System.Collections.Generic;
using System.Text;
using System.Security.Principal;
using System.Data.OracleClient;
using System.Data;
using System.Windows.Forms;

namespace LTZN
{
    public class LtznUserManager
    {
        public static readonly LtznUserManager instance = new LtznUserManager();

        private LtznUserManager()
        {
        }

        private DataTable userRole=new DataTable();
        private bool loadUR = false;

        private LtznUser currentUser = null;

        //用户登陆
        public bool Login(string name,string psw)
        {
            bool authentication=false;

            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "select count(*) from users where yonghu=:name and mima=:psw";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":name", OracleType.VarChar).Value = name;
            cmd.Parameters.Add(":psw", OracleType.VarChar).Value = psw;
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (dr.GetInt32(0) > 0)
                    authentication = true;

            }
            dr.Close();
            conn.Close();
            if (authentication)
            {
                currentUser = new LtznUser(name);
                if (UserChanged != null)
                    UserChanged(this.currentUser);
            }
            return authentication;
        }
  
        //用户注销
        public void Quit()
        {
            currentUser = null;
            if (UserChanged != null)
                UserChanged(this.currentUser);
        }

        //创建用户
        public bool CreateUser(string name, string psw)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "insert into users(yonghu,mima) values(:name,:psw)";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":name", OracleType.VarChar).Value = name;
            cmd.Parameters.Add(":psw", OracleType.VarChar).Value = psw;
            int i = cmd.ExecuteNonQuery();
            conn.Close();
            if (i > 0)
                return true;
            else
                return false;
        }

        //删除用户
        public void DelUser(string name)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "delete from userRole where userName=:name";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":name", OracleType.VarChar).Value = name;
            cmd = null;

            sql = "delete from users where yonghu=:name";
            cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":name", OracleType.VarChar).Value = name;

            cmd.ExecuteNonQuery();
            conn.Close();
        }

        //修改密码
        public bool ChgPsw(string name, string psw)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "update users set mima=:psw where yonghu=:name";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":name", OracleType.VarChar).Value = name;
            cmd.Parameters.Add(":psw", OracleType.VarChar).Value = psw;
            int i=cmd.ExecuteNonQuery();
            conn.Close();
            if (i > 0)
                return true;
            else
                return false;
        }

        //注册用户变更事件
        public void RegisterHandler(Form frm, Action<LtznUser> userChangedHandler)
        {
            frm.Load += delegate(object sender, EventArgs e) { userChangedHandler(this.currentUser); this.UserChanged += userChangedHandler; };
            frm.FormClosed += delegate(object sender, FormClosedEventArgs e) { this.UserChanged -= userChangedHandler; };
        }

        public void addRole(string name, string role)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "insert into UserRole(username,role) values(:name,:role)";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":name", OracleType.VarChar).Value = name;
            cmd.Parameters.Add(":role", OracleType.VarChar).Value = role;
            cmd.ExecuteNonQuery();
            conn.Close();
            loadUR = false;
        }

        public void removeRole(string name, string role)
        {

            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "delete from UserRole where username=:name and  role=:role";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":name", OracleType.VarChar).Value = name;
            cmd.Parameters.Add(":role", OracleType.VarChar).Value = role;
            cmd.ExecuteNonQuery();
            conn.Close();
            loadUR = false;
        }

        public bool checkUserRole(string name, string role)
        {

            bool inRole = false;
            foreach (string rl in getUserRoles(name))
            {
                if (rl == role)
                    inRole = true;
            }
           
            return inRole;
        }
     
        public string[] getUserRoles(string name)
        {  
            
            if (!loadUR)
                reloadUserRole(); 
            DataRow[] drs=userRole.Select("username='" + name + "'");
            string[] roles=new string[drs.Length];
            for (int i = 0; i < drs.Length; i++)
            {
                roles[i] = drs[i][1].ToString();
            }
            return roles;
        }

        private void reloadUserRole()
        {
                OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
                conn.Open();
                string sql = "select username,role from userrole";
                OracleCommand cmd = new OracleCommand(sql, conn);
                OracleDataAdapter da = new OracleDataAdapter(cmd);
                da.Fill(userRole);
                conn.Close();
                loadUR = true;
        }

        public string[] getAllUsers()
        {
            List<string> users = new List<string>();
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "select yonghu from users";
            OracleCommand cmd = new OracleCommand(sql, conn);
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                users.Add(dr.GetString(0));
            }
            dr.Close();
            conn.Close();
            return users.ToArray();
        }

        public string[] getAllRoles()
        {
            List<string> roles = new List<string>();
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "select Name from Roles";
            OracleCommand cmd = new OracleCommand(sql, conn);
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                roles.Add(dr.GetString(0));
            }
            dr.Close();
            conn.Close();
            return roles.ToArray();
        }

        public LtznUser CurrentUser
        {
            get { return currentUser; }

        }
    
        public event Action<LtznUser> UserChanged;

        DateTime now = DateTime.Now;
        
        public DateTime Now
        {
            get
            {
               
                    OracleConnection cn = new OracleConnection();
                    cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
                    cn.Open();
                    OracleCommand selCmd = new OracleCommand();
                    selCmd.Connection = cn;
                    selCmd.CommandText = "Select SYSDATE from DUAL";
                    OracleDataReader dr = selCmd.ExecuteReader();
                    if (dr.Read())
                    {
                        if (!dr.IsDBNull(0))
                            now = dr.GetDateTime(0);
                    }
                    dr.Close();
                    cn.Close();
                    return now;
              
            }
        
        }
    }
   
    public class LtznUser : IPrincipal
    {
        private LtznIdentity myIdentity = new LtznIdentity();

        public LtznUser(string name)
        {
            myIdentity.Name = name;
            myIdentity.IsAuthenticated = true;
        }

        public LtznUser(string name, bool authenticated)
        {
            myIdentity.Name = name;
            myIdentity.IsAuthenticated = authenticated;
        }

        public IIdentity Identity { get { return myIdentity; } }

        public bool IsInRole(string role)
        {
            return LtznUserManager.instance.checkUserRole(myIdentity.Name,role);
        }
    }

    class LtznIdentity : IIdentity
    {
        private string authenticationType = "";
        private bool isAuthenticated = false;
        private string name = "";

        #region IIdentity 成员

        public string AuthenticationType
        {
            get { return authenticationType; }
            set { authenticationType = value; }
        }

        public bool IsAuthenticated
        {
            get { return isAuthenticated; }
            set { isAuthenticated = value; }
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        #endregion

    }

    public class LtznRole
    {
        private string _roleName = "";

        public string RoleName
        {
            get { return _roleName; }
            set { _roleName = value; }
        }

        public bool CanAccess(string funcCode)
        {
            if(allowAccess.Contains(funcCode))
                return true;

            return false;
        }

        public void AllowFuncCode(string funcCode)
        {
            if (!allowAccess.Contains(funcCode))
                allowAccess.Add(funcCode);
        }

        public void RemoveFuncCode(string funcCode)
        {
            if (allowAccess.Contains(funcCode))
                allowAccess.Remove(funcCode);
        }

        private List<string> allowAccess = new List<string>();
   
    }
}
