using log4net.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Proj96333
{
    public class GLog
    {
        private readonly static Type ThisDeclaringType = typeof(GLog);
        private static readonly ILogger s_Instance;

        private static ILogger Instance
        {
            get
            {
                if (null == s_Instance)
                {
                    new GLog();
                }
                return s_Instance;
            }
        }

        static GLog()
        {
            s_Instance = LoggerManager.GetLogger(Assembly.GetCallingAssembly(), "GLog");
        }

        public static void I(String msg)
        {
            GLog.Instance.Log(ThisDeclaringType, Level.Info, msg, null);
        }

        public static void D(String msg)
        {
            GLog.Instance.Log(ThisDeclaringType, Level.Debug, msg, null);
        }

        public static void E(String msg)
        {
            GLog.Instance.Log(ThisDeclaringType, Level.Error, msg, null);
        }

    }
}
