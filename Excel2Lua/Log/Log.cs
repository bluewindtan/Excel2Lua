using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Excel2Lua
{
	public sealed class LogMgr
	{
		private static LogMgr instance = null;
		private static readonly object padlock = new object();

		public static LogMgr Instance
		{
			get
			{
				if (null == instance)
				{
					lock (padlock)
					{
						if (null == instance)
						{
							instance = new LogMgr();
						}
					}
				}
				return instance;
			}
		}

		private StreamWriter m_log = null;

		~LogMgr()
		{
			CloseLog();
		}

		public void WriteLog(string strContent)
		{
			if (null != m_log)
			{
				m_log.WriteLine(strContent);
			}
		}

		public void OpenLog()
		{
			if (m_log != null)
			{
				return;
			}
			string logDir = Environment.CurrentDirectory + "\\log";
			if (!Directory.Exists(logDir))
			{
				Directory.CreateDirectory(logDir);
			}
			string logFile = logDir + "\\" + CustomDefine.LOG_NAME + CustomDefine.SUFFIX_LOG;
			if (File.Exists(logFile))
			{
				string logFileBak = logDir + "\\" + CustomDefine.LOG_NAME + "_1" + CustomDefine.SUFFIX_LOG;
				if (File.Exists(logFileBak))
				{
					File.Delete(logFileBak);
				}
				File.Move(logFile, logFileBak);
			}
			m_log = new StreamWriter(logFile, false, CustomDefine.WRITE_ENCODING);
		}

		public void CloseLog()
		{
			if (m_log != null)
			{
				m_log.Close();
				m_log = null;
			}
		}
	}
}
