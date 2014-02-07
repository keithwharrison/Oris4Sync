using System;

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
    }
}
