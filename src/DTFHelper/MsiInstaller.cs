using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Deployment.WindowsInstaller;
using Microsoft.Deployment.WindowsInstaller.Package;
using DTFHelper.Extensions;
using System.Threading;

namespace DTFHelper
{
    public class ActionStartEventArgs : EventArgs
    {
        // TODO: define appropriate members
        public InstallMessage MessageType { get; set; }
        public Record MessageRecord { get; set; }
    }
    public class TerminateEventArgs : EventArgs
    {
        // TODO
    }
    /// <summary>
    /// progress event kind,see the following url.
    /// https://msdn.microsoft.com/en-us/library/aa368786.aspx
    /// </summary>
    public enum ProgressEventKind
    {
        /// <summary>
        /// reset progress bar
        /// </summary>
        /// <list type="bullet">
        /// <item>field 1 = always 0</item>
        /// <item>field 2 = total number of ticks</item>
        /// <item>field 3 = progress direction(0 = forward,1 = back)</item>
        /// <item>field 4 = is script in progress(1 = true, 0 = false)</item>
        /// </list>

        Reset = 0,
        /// <summary>
        /// increment progress
        /// </summary>
        /// <list type="number">
        /// <item>always 1</item>
        /// <item>number of ticks to increment bar</item>
        /// <item>should calculate(if 0,then you should ignore the message)</item>
        /// </list>
        ActionInfo = 1,
        /// <summary>
        /// progress report, should increment progress
        /// <list type="number">
        /// <item>always 2</item>
        /// <item>number of progress</item>
        /// </list>
        /// </summary>
        ProgressReport = 2,
        /// <summary>
        /// unused?
        /// </summary>
        ProgressAddition = 3,
    }
    public class ProgressInfo
    {
        public int TotalProgress { get; set; }
        public bool IsForward { get; set; }
        public int CurrentPosition { get; set; }
        public bool IsScriptInProgress { get; set; }
        public int ProgressTicks { get; set; }
        public bool IsEnableActionData { get; set; }
        public int CurrentTicks { get; set; }
    }
    /// <summary>
    /// https://msdn.microsoft.com/en-us/library/aa368786.aspx
    /// </summary>
    public class ProgressEventArgs :EventArgs
    {
        public int TotalProgress { get; set; }
        public bool IsForward { get; set; }
        public int CurrentTicks { get; set; }
    }
    public class LoggingEventArgs : EventArgs
    {
        public string FormattedMessage { get; set; }
        public InstallLogModes LogMode { get; set; }
    }
    internal static class MsiRecordExtension
    {
        public static ProgressEventArgs ToProgressEventArgs(this ProgressInfo info)
        {
            return new ProgressEventArgs()
            {
                CurrentTicks = info.CurrentTicks
                ,
                IsForward = info.IsForward
                ,
                TotalProgress = info.TotalProgress
            };
        }
        public static ProgressInfo ProcessMessage(this ProgressInfo progressInfo,Record messageRecord)
        {
            if (messageRecord.FieldCount < 1)
            {
                return progressInfo;
            }
            var kind = messageRecord.GetInteger(1);
            switch(kind)
            {
                case (int)ProgressEventKind.Reset:
                    {
                        progressInfo.TotalProgress = messageRecord.GetInteger(2);
                        progressInfo.IsForward = messageRecord.GetInteger(3) == 0;
                        progressInfo.ProgressTicks = progressInfo.IsForward ? 0 : progressInfo.TotalProgress;
                        progressInfo.IsScriptInProgress = messageRecord.GetInteger(4) == 1;
                        progressInfo.CurrentPosition = 0;
                        progressInfo.CurrentTicks = progressInfo.IsScriptInProgress ? progressInfo.TotalProgress : progressInfo.ProgressTicks;
                        return progressInfo;
                    }
                case (int)ProgressEventKind.ActionInfo:
                    {
                        progressInfo.IsEnableActionData = messageRecord.GetInteger(3) != 0;
                        var ticks = progressInfo.IsForward ? messageRecord.GetInteger(2) : (-1 * messageRecord.GetInteger(2));
                        progressInfo.CurrentTicks += ticks;
                        return progressInfo;
                    }
                case (int)ProgressEventKind.ProgressReport:
                    {
                        if (progressInfo.TotalProgress == 0)
                        {
                            return progressInfo;
                        }
                        else
                        {
                            progressInfo.CurrentPosition += messageRecord.GetInteger(2);
                            progressInfo.CurrentTicks += progressInfo.IsForward ? progressInfo.CurrentPosition : (-1 * progressInfo.CurrentPosition);
                            return progressInfo;
                        }
                    }
                    break;
                default:
                    return null;
            }
            return null;
        }
    }
    public class MsiInstaller
    {
        string m_MsiPath;
        public MsiInstaller(string msiPath)
        {
            m_MsiPath = msiPath;
            m_ExternalUIHandler = new ExternalUIRecordHandler(OnExternalUI);
            Properties = new Dictionary<string, string>();
        }
        public IDictionary<string, string> Properties
        {
            get;
            private set;
        }
        bool m_IsForwardProgress = true;
        int m_TotalProgressTicks;
        int m_CurrentPosition = 0;
        int m_CurrentProgressTicks = 0;
        ExternalUIRecordHandler m_ExternalUIHandler;
        ProgressInfo m_ProgressInfo = new ProgressInfo()
        {
        };
        MessageResult OnExternalUI(InstallMessage messageType, Record messageRecord, MessageButtons buttons, MessageIcon icon, MessageDefaultButton defaultButton)
        {
            switch (messageType)
            {
                case InstallMessage.ActionStart:
                    m_ProgressInfo.IsEnableActionData = false;
                    if (OnActionStart != null)
                    {
                        return OnActionStart(this, new ActionStartEventArgs() 
                        {
                            MessageType = messageType,
                            MessageRecord = messageRecord 
                        });
                    }
                    break;
                case InstallMessage.InstallStart:
                    // TODO(before installation start)
                    break;
                case InstallMessage.InstallEnd:
                    // TODO(after installation end)
                    break;
                case InstallMessage.RMFilesInUse:
                    // TODO(List of apps that the user can request Restart Manager to shut down and restart.)
                    break;
                case InstallMessage.ShowDialog:
                    // TODO(show dialog event)
                    break;
                case InstallMessage.CommonData:
                    // TODO(common product info:language id,caption)
                    break;
                case InstallMessage.ActionData:
                    // TODO(formatted data for action)
                    {
                        if(m_ProgressInfo.TotalProgress == 0)
                        {
                            return MessageResult.OK;
                        }
                        if (m_ProgressInfo.IsEnableActionData)
                        {
                            m_ProgressInfo.CurrentTicks = 0;
                        }
                        if(OnProgress != null)
                        {
                            OnProgress(this, m_ProgressInfo.ToProgressEventArgs());
                        }
                    }
                    break;
                case InstallMessage.Error:
                case InstallMessage.Warning:
                case InstallMessage.Info:
                    // TODO(logging)
                    break;
                case InstallMessage.User:
                    // TODO(user request message)
                    break;
                case InstallMessage.Terminate:
                    // TODO(after ui termination,no string data)
                    if(OnTerminate != null)
                    {
                        return OnTerminate(this, new TerminateEventArgs());
                    }
                    break;
                case InstallMessage.Initialize:
                    // TODO(before ui initialize,no string data)
                    break;
                case InstallMessage.Progress:
                    if(OnProgress != null)
                    {
                        m_ProgressInfo.ProcessMessage(messageRecord);
                        return OnProgress(this, m_ProgressInfo.ToProgressEventArgs());
                    }
                    break;
                default:
                    break;
            }
            // nothing to do in this callback. pass another layer
            return MessageResult.None;
        }
        public void ExecuteAdministrativeInstall(string destdir)
        {
            if (!InstallationMutex.WaitOne(1))
            {
                throw new InvalidOperationException("another installation running");
            }
            try
            {
                IsRunningAnotherInstallation = true;
                Installer.InstallProduct(m_MsiPath, string.Format("ACTION=ADMIN TARGETDIR=\"{0}\"", destdir));
            }
            finally
            {
                IsRunningAnotherInstallation = false;
                InstallationMutex.ReleaseMutex();
            }
        }
        public string LogFilePath { get; set; }
        public void ExecuteInstall()
        {
            if (!InstallationMutex.WaitOne(1))
            {
                throw new InvalidOperationException("another installation running");
            }
            try
            {
                Installer.EnableLog(InstallLogModes.Verbose, LogFilePath);
                IsRunningAnotherInstallation = true;
                var oldHandler = Installer.SetExternalUI(new ExternalUIRecordHandler(OnExternalUI)
                    , InstallLogModes.ActionData
                    | InstallLogModes.ActionStart
                    | InstallLogModes.Error
                    | InstallLogModes.Initialize
                    | InstallLogModes.Progress
                    | InstallLogModes.ShowDialog
                    | InstallLogModes.Terminate
                    | InstallLogModes.User
                    | InstallLogModes.Info
                    | InstallLogModes.Warning
                    | InstallLogModes.FilesInUse)
                    ;
                try
                {
                    Installer.InstallProduct(m_MsiPath, string.Join(" "
                        , Properties.Select(kv => string.Format("{0}=\"{1}\"", kv.Key, kv.Value.Replace("\\", "\\\\")))));
                }
                finally
                {
                    Installer.SetExternalUI(oldHandler, InstallLogModes.None);
                }
            }
            finally
            {
                InstallationMutex.ReleaseMutex();
                IsRunningAnotherInstallation = false;
            }
        }
        public delegate MessageResult OnActionStartCallback(object sender, ActionStartEventArgs e);
        public event OnActionStartCallback OnActionStart;
        public delegate MessageResult OnTerminateCallback(object sender, TerminateEventArgs e);
        public event OnTerminateCallback OnTerminate;
        public delegate MessageResult OnProgressCallback(object sender, ProgressEventArgs e);
        public event OnProgressCallback OnProgress;
        public delegate MessageResult OnLoggingMessageCallback(object sender, LoggingEventArgs e);
        public event OnLoggingMessageCallback OnLogging;
        /// <summary>
        /// MSI does not allow the parallel install transaction.
        /// the mutex is for preventing parallel install transaction
        /// </summary>
        static Mutex InstallationMutex = new Mutex();
        static bool m_IsRunningAnoterInstallation = false;
        public bool IsRunningAnotherInstallation
        {
            get
            {
                lock (InstallationLockObject)
                {
                    return m_IsRunningAnoterInstallation;
                }
            }
            private set
            {
                lock(InstallationLockObject)
                {
                    m_IsRunningAnoterInstallation = value;
                }
            }
        }
        static object InstallationLockObject = new object();

        static readonly string[] UpgradeTableColumns = new string[]
        {
            "UpgradeCode",
            "VersionMin",
            "VersionMax",
            "Language",
            "Attributes",
            "Remove",
            "ActionProperty",
        };
        /// <summary>
        /// find related products in system(for more information: https://msdn.microsoft.com/en-us/library/windows/desktop/aa368600%28v=vs.85%29.aspx).
        /// if productid is not found in system and related products are found,msiexec will execute major upgrade.
        /// </summary>
        /// <param name="msiPath">path to msi file</param>
        /// <returns>keyvalue pair of ActionProperty and product id list separeted by ';'</returns>
        public static IDictionary<string, string> GetRelatedProducts(string msiPath)
        {
            using (var db = new InstallPackage(msiPath, DatabaseOpenMode.Transact))
            {
                if (!db.Tables.Any(x => x.Name == "Upgrade"))
                {
                    return new Dictionary<string, string>();
                }
                var upgradeCodeList = db.ExecuteQueryToDictionary(UpgradeTableColumns, "SELECT {0} FROM Upgrade"
                    , string.Join(",", UpgradeTableColumns));
                using (var session = Installer.OpenPackage(db, true))
                {
                    session.DoAction("FindRelatedProducts");
                    Dictionary<string, string> result = new Dictionary<string, string>();
                    foreach (var upgrade in upgradeCodeList)
                    {
                        var propResult = session.GetProductProperty(upgrade["ActionProperty"].ToString());
                        if(!string.IsNullOrEmpty(propResult))
                        {
                            result.Add(upgrade["ActionProperty"].ToString(), propResult);
                        }
                    }
                    return result;
                }
            }
        }
    }
}
