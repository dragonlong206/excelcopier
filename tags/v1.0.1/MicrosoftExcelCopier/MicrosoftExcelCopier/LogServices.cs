using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using log4net;

namespace MicrosoftExcelCopier
{
    public static class LogServices
    {
        private static readonly ILog mainLogger = LogManager.GetLogger("MAINLogger");
        private static readonly ILog debugLogger = LogManager.GetLogger("DEBUGLogger");

        public static void WriteError(object message, Exception ex)
        {
            if (mainLogger != null && mainLogger.IsErrorEnabled)
            {
                mainLogger.Error(message, ex);
            }
        }

        public static void WriteError(object message)
        {
            if (mainLogger != null && mainLogger.IsErrorEnabled)
            {
                mainLogger.Error(message);
            }
        }

        public static void WriteDebug(object message, Exception ex)
        {
            if (debugLogger != null && debugLogger.IsErrorEnabled)
            {
                debugLogger.Debug(message, ex);
            }
        }

        public static void WriteDebug(object message)
        {
            if (debugLogger != null && debugLogger.IsErrorEnabled)
            {
                debugLogger.Debug(message);
            }
        }
    }
}
