using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CmisSync.Lib.Outlook
{
    public class Config
    {
        private static Config instance;

        public static Config Instance
        {
            get
            {
                if (instance == null)
                {
                    Type type = Type.GetType("CmisSync.Lib.Outlook.ConfigOverride");
                    if (type != null)
                    {
                        instance = (Config)Activator.CreateInstance(type);
                    }
                    else
                    {
                        instance = new Config();
                    }
                }
                return instance;
            }
        }

        protected Config()
        {
        }

        public virtual string ConsumerKey
        {
            get
            {
                return "consumerKey";
            }
        }

        public virtual string ConsumerSecret
        {
            get
            {
                return "consumerSecret";
            }
        }

        public virtual string GrantType
        {
            get
            {
                return "password";
            }
        }


        public virtual string TestUrl
        {
            get
            {
                return "http://localhost:8080";
            }
        }

        public virtual string TestUsername
        {
            get
            {
                return "username";
            }
        }

        public virtual string TestPassword
        {
            get
            {
                return "password";
            }
        }
    }
}
