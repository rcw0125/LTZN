using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;

namespace LTZN
{
    public class User
    {
        public string name;
        public string quan;
    }
    public delegate void  UserChangeHandler(User oldUser,User newUser);
    public static class UserManager
    {
        private static User curUser = null;
        public static event UserChangeHandler OnUserChanged;
        public static void Login(string name, string psw)
        {
            User user = new User();
            user.name = name;

            OnUserChanged(curUser, user);
            curUser = user;
        }
        public static void Quit()
        {
            OnUserChanged(curUser, null);
            curUser = null;
        }
        public static bool Check(string quan)
        {
            bool returnValue = false;
            if (curUser != null)
            {
                if (!String.IsNullOrEmpty(curUser.quan))
                {
                    string[] qs = curUser.quan.Split(',');
                    foreach (string s in qs)
                    {
                        if (s==quan)
                           returnValue = true;
                    }
                }
            }
            return returnValue;
        }
    }
}
