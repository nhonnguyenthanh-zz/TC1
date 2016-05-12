using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using log4net.Config;

namespace TC1.utility
{
    class Log
    {
        //Initialize Log4j logs
        private static ILog _Log = LogManager.GetLogger(typeof(Log).Name);

	 
		// This is to print log for the beginning of the test case, as we usually run so many test cases as a test suite
		public static void startTestCase(String sTestCaseName)
        {

            _Log.Info("****************************************************************************************");
            _Log.Info("****************************************************************************************");
            _Log.Info("$$$$$$$$$$$$$$$$$$$$$                 " + sTestCaseName + "       $$$$$$$$$$$$$$$$$$$$$$$$$");
            _Log.Info("****************************************************************************************");
            _Log.Info("****************************************************************************************");

        }

        //This is to print log for the ending of the test case
        public static void endTestCase(String sTestCaseName)
        {
            _Log.Info("XXXXXXXXXXXXXXXXXXXXXXX             " + "-E---N---D-" + "             XXXXXXXXXXXXXXXXXXXXXX");
            _Log.Info("X");
            _Log.Info("X");
            _Log.Info("X");
            _Log.Info("X");

        }

        // Need to create these methods, so that they can be called  
        public static void info(String message)
        {
            _Log.Info(message);
        }

        public static void warn(String message)
        {
            _Log.Warn(message);
        }

        public static void error(String message)
        {
            _Log.Error(message);
        }

        public static void fatal(String message)
        {
            _Log.Fatal(message);
        }

        public static void debug(String message)
        {
            _Log.Debug(message);
        }


    }
}
