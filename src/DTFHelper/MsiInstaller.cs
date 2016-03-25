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
        public bool IsEnableActionData { get; set; }
        public int ProgressPerAction { get; set; }
        public string CurrentAction { get; set; }
    }
    /// <summary>
    /// https://msdn.microsoft.com/en-us/library/aa368786.aspx
    /// </summary>
    public class ProgressEventArgs : EventArgs
    {
        public int TotalProgress { get; set; }
        public bool IsForward { get; set; }
        public int CurrentPosition { get; set; }
        public string CurrentAction { get; set; }
        public bool IsEnableActionData { get; set; }
    }
    public class LoggingEventArgs : EventArgs
    {
        public string FormattedMessage { get; set; }
        public string PreformatMessage { get; set; }
        public InstallLogModes LogMode { get; set; }
        public IList<string> Parameters { get; set; }
    }
    public class ExceptionEventArgs : EventArgs
    {
        public Exception Exception { get; set; }
    }
    public class FilesInUseEventArgs : EventArgs
    {
        public IList<FilesInUseElement> Files { get; set; }
    }
    public class CancelShowChangedEventArgs : EventArgs
    {
        public bool ShouldEnableCancel { get; set; }
    }
    public class CaptionInformedEventArgs : EventArgs
    {
        public string Title { get; set; }
    }
    public class LanguageInfoEventArgs : EventArgs
    {
        public int LanguageID { get; set; }
        public int CodePage { get; set; }
    }
    internal static class MsiRecordExtension
    {
        public static ProgressEventArgs ToProgressEventArgs(this ProgressInfo info)
        {
            return new ProgressEventArgs()
            {
                CurrentPosition = info.CurrentPosition
                ,
                IsForward = info.IsForward
                ,
                TotalProgress = info.TotalProgress
                ,
                CurrentAction = info.CurrentAction
                ,
                IsEnableActionData = info.IsEnableActionData
            };
        }
        public static ProgressInfo ProcessMessage(this ProgressInfo progressInfo, Record messageRecord)
        {
            if (messageRecord == null)
            {
                return progressInfo;
            }
            if (messageRecord.FieldCount < 1)
            {
                return progressInfo;
            }
            var kind = messageRecord.GetInteger(1);
            switch (kind)
            {
                case (int)ProgressEventKind.Reset:
                    {
                        progressInfo.TotalProgress = messageRecord.GetInteger(2);
                        progressInfo.IsForward = messageRecord.GetInteger(3) == 0;
                        progressInfo.IsScriptInProgress = messageRecord.GetInteger(4) == 1;
                        progressInfo.CurrentPosition = progressInfo.IsForward ? 0 : progressInfo.TotalProgress;
                        return progressInfo;
                    }
                case (int)ProgressEventKind.ActionInfo:
                    {
                        progressInfo.IsEnableActionData = messageRecord.GetInteger(3) != 0;
                        if (messageRecord.GetInteger(3) != 0)
                        {
                            progressInfo.ProgressPerAction = messageRecord.GetInteger(2);
                        }
                        else
                        {
                            progressInfo.ProgressPerAction = 0;
                        }
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
                            return progressInfo;
                        }
                    }
                case (int)ProgressEventKind.ProgressAddition:
                    {
                        progressInfo.TotalProgress += messageRecord.GetInteger(2);
                        return progressInfo;
                    }
                default:
                    return null;
            }
        }
        static List<Tuple<InstallLogModes, InstallMessage>> LogModeMessageList = new List<Tuple<InstallLogModes, InstallMessage>>()
        {
            Tuple.Create(InstallLogModes.Error,InstallMessage.Error),
            Tuple.Create(InstallLogModes.Warning,InstallMessage.Warning),
            Tuple.Create(InstallLogModes.Info,InstallMessage.Info),
            Tuple.Create(InstallLogModes.User,InstallMessage.User),
            Tuple.Create(InstallLogModes.Warning,InstallMessage.Warning),
            Tuple.Create(InstallLogModes.ActionData,InstallMessage.ActionData),
            Tuple.Create(InstallLogModes.ActionStart,InstallMessage.ActionStart),
            Tuple.Create(InstallLogModes.CommonData,InstallMessage.CommonData),
            Tuple.Create(InstallLogModes.FatalExit,InstallMessage.FatalExit),
            Tuple.Create(InstallLogModes.FilesInUse,InstallMessage.FilesInUse),
            Tuple.Create(InstallLogModes.Initialize,InstallMessage.Initialize),
            Tuple.Create(InstallLogModes.OutOfDiskSpace,InstallMessage.OutOfDiskSpace),
            Tuple.Create(InstallLogModes.Progress,InstallMessage.Progress),
            Tuple.Create(InstallLogModes.ResolveSource,InstallMessage.ResolveSource),
            Tuple.Create(InstallLogModes.RMFilesInUse,InstallMessage.RMFilesInUse),
            Tuple.Create(InstallLogModes.ShowDialog,InstallMessage.ShowDialog),
            Tuple.Create(InstallLogModes.Terminate,InstallMessage.Terminate),
        };
        public static InstallLogModes ToInstallLogMode(this InstallMessage msgType)
        {
            var t = LogModeMessageList.FirstOrDefault(x => x.Item2 == msgType);
            if (t != null)
            {
                return t.Item1;
            }
            else
            {
                throw new ArgumentException("unknown argument", "msgType");
            }
        }
    }
    public class FilesInUseElement
    {
        public string ProcessName { get; set; }
        public int? ProcessId { get; set; }
        public string WindowTitle { get; set; }
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
        ExternalUIRecordHandler m_ExternalUIHandler;
        ProgressInfo m_ProgressInfo = new ProgressInfo()
        {
        };
        MessageResult ProcessCommonData(Record messageRecord)
        {
            var type = messageRecord.GetInteger(1);
            switch(type)
            {
                case 0:
                    if(OnLanguageInfo != null)
                    {
                        OnLanguageInfo(this, new LanguageInfoEventArgs()
                        {
                            LanguageID = messageRecord.GetInteger(2)
                            ,
                            CodePage = messageRecord.GetInteger(3)
                        });
                    }
                    break;
                case 1:
                    if(OnCaptionInformed != null)
                    {
                        OnCaptionInformed(this, new CaptionInformedEventArgs() { Title = messageRecord.GetString(2) });
                    }
                    break;
                case 2:
                    if(OnCancelShowChanged != null)
                    {
                        OnCancelShowChanged(this, new CancelShowChangedEventArgs()
                        {
                            ShouldEnableCancel = messageRecord.GetInteger(2) != 0
                        });
                    }
                    break;
                default:
                    break;
            }
            return MessageResult.None;
        }
        MessageResult ProcessMessage(InstallMessage messageType, Record messageRecord, MessageButtons buttons, MessageIcon icon, MessageDefaultButton defaultButton)
        {
            switch (messageType)
            {
                case InstallMessage.ActionStart:
                    m_ProgressInfo.IsEnableActionData = false;
                    m_ProgressInfo.CurrentAction = messageRecord.GetString(1);
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
                    // TODO(before installation start(before initailize))
                    if (OnInitialize != null)
                    {
                        OnInitialize(this, new EventArgs());
                    }
                    break;
                case InstallMessage.InstallEnd:
                    if(OnInstallEnd != null)
                    {
                        OnInstallEnd(this, new EventArgs());
                    }
                    break;
                case InstallMessage.FilesInUse:
                case InstallMessage.RMFilesInUse:
                    // TODO(List of apps that the user can request Restart Manager to shut down and restart.)
                    if (OnFilesInUse != null)
                    {
                        return OnFilesInUse(this, new FilesInUseEventArgs()
                        {
                            Files = Enumerable.Range(1, messageRecord.FieldCount / 2).Select(i =>
                              {
                                  var procName = messageRecord.GetString(i * 2 + 1);
                                  var pidOrTitle = messageRecord.GetString(i * 2 + 2);
                                  int? pid = null;
                                  string title = null;
                                  int tmp = 0;
                                  if(int.TryParse(pidOrTitle,out tmp))
                                  {
                                      pid = tmp;
                                  }
                                  else
                                  {
                                      title = pidOrTitle;
                                  }
                                  return new FilesInUseElement()
                                  {
                                      ProcessName = procName
                                      ,
                                      ProcessId = pid
                                      ,
                                      WindowTitle = title
                                  };
                              }).ToList()
                        });
                    }
                    break;
                case InstallMessage.ShowDialog:
                    // TODO(show dialog event)
                    break;
                case InstallMessage.CommonData:
                    // TODO(common product info:language id,caption)
                    return ProcessCommonData(messageRecord);
                case InstallMessage.ActionData:
                    {
                        if (m_ProgressInfo.TotalProgress == 0)
                        {
                            return MessageResult.OK;
                        }
                        if (m_ProgressInfo.IsEnableActionData)
                        {
                            m_ProgressInfo.CurrentPosition += m_ProgressInfo.ProgressPerAction;
                        }
                        if (OnProgress != null)
                        {
                            OnProgress(this, m_ProgressInfo.ToProgressEventArgs());
                        }
                    }
                    break;
                case InstallMessage.Error:
                case InstallMessage.Warning:
                case InstallMessage.Info:
                case InstallMessage.User:
                    if (OnLogging != null)
                    {
                        OnLogging(this, new LoggingEventArgs()
                        {
                            FormattedMessage = messageRecord.ToString()
                            ,
                            PreformatMessage = messageRecord.FormatString
                            ,
                            LogMode = messageType.ToInstallLogMode()
                            ,
                            Parameters = Enumerable.Range(0, messageRecord.FieldCount + 1).Select(x => messageRecord[x] != null ? messageRecord[x].ToString() : "").ToArray()
                        });
                    }
                    break;
                case InstallMessage.Terminate:
                    if (OnTerminate != null)
                    {
                        return OnTerminate(this, new TerminateEventArgs());
                    }
                    break;
                case InstallMessage.Initialize:
                    if (OnInitialize != null)
                    {
                        return OnInitialize(this, new EventArgs());
                    }
                    break;
                case InstallMessage.Progress:
                    if (OnProgress != null)
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
        MessageResult OnExternalUI(InstallMessage messageType, Record messageRecord, MessageButtons buttons, MessageIcon icon, MessageDefaultButton defaultButton)
        {
            try
            {
                var records = messageRecord != null ? Enumerable.Range(0, messageRecord.FieldCount + 1).Select(i => messageRecord.GetString(i)) : Enumerable.Empty<string>();
                System.Diagnostics.Trace.WriteLine(string.Format("onexternalui({0}):{1},{2}"
                    , messageType
                    , messageRecord == null ? "(null)" : messageRecord.ToString()
                    , messageRecord != null ? string.Join(";", records) : ""
                    ));
                return ProcessMessage(messageType, messageRecord, buttons, icon, defaultButton);
            }
            catch (Exception e)
            {
                if (OnException != null)
                {
                    return OnException(this, new ExceptionEventArgs()
                    {
                        Exception = e
                    });
                }
                return MessageResult.Abort;
            }
        }
        /// <summary>
        /// Execute Administrative installation
        /// </summary>
        /// <param name="destdir">target root dir</param>
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
        string GetProductCode(string msiPath)
        {
            using (var db = new InstallPackage(msiPath, DatabaseOpenMode.ReadOnly))
            {
                return db.Property["ProductCode"];
            }
        }
        public void ExecuteUninstall()
        {
            if (!InstallationMutex.WaitOne(1))
            {
                throw new InvalidOperationException("another installation running");
            }
            try
            {
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
                    | InstallLogModes.FilesInUse
                    | InstallLogModes.CommonData
                    | InstallLogModes.RMFilesInUse
                    )
                    ;
                IsRunningAnotherInstallation = true;
                try
                {
                    var productCode = GetProductCode(m_MsiPath);
                    Installer.ConfigureProduct(productCode, 0, InstallState.Absent
                        , string.Join(" ", Properties.Select(kv => string.Format("{0}=\"{1}\"", kv.Key, kv.Value.Replace("\\", "\\\\")))));
                }
                finally
                {
                    Installer.SetExternalUI(oldHandler, InstallLogModes.None);
                }
            }
            finally
            {
                IsRunningAnotherInstallation = false;
            }
        }
        public void ExecuteInstall()
        {
            if (!InstallationMutex.WaitOne(1))
            {
                throw new InvalidOperationException("another installation running");
            }
            try
            {
                var option = Installer.SetInternalUI(InstallUIOptions.Full);
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
                    | InstallLogModes.FilesInUse
                    | InstallLogModes.CommonData
                    | InstallLogModes.RMFilesInUse
                    )
                    ;
                try
                {
                    Installer.InstallProduct(m_MsiPath, string.Join(" "
                        , Properties.Select(kv => string.Format("{0}=\"{1}\"", kv.Key, kv.Value.Replace("\\", "\\\\")))));
                }
                finally
                {
                    Installer.SetInternalUI(option);
                    Installer.SetExternalUI(oldHandler, InstallLogModes.None);
                }
            }
            finally
            {
                InstallationMutex.ReleaseMutex();
                IsRunningAnotherInstallation = false;
            }
        }
        public event Func<object, ActionStartEventArgs, MessageResult> OnActionStart;
        public event Func<object, TerminateEventArgs, MessageResult> OnTerminate;
        public event Func<object, ProgressEventArgs, MessageResult> OnProgress;
        public event Func<object, LoggingEventArgs, MessageResult> OnLogging;
        public event Func<object, EventArgs, MessageResult> OnInitialize;
        public event Func<object, ExceptionEventArgs, MessageResult> OnException;
        public event Func<object, FilesInUseEventArgs, MessageResult> OnFilesInUse;
        public event Action<object, EventArgs> OnInstallEnd;
        public event Action<object, CancelShowChangedEventArgs> OnCancelShowChanged;
        public event Action<object, CaptionInformedEventArgs> OnCaptionInformed;
        public event Action<object, LanguageInfoEventArgs> OnLanguageInfo;

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
                lock (InstallationLockObject)
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
            var previous = Installer.SetInternalUI(InstallUIOptions.Silent);
            try
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
                            if (!string.IsNullOrEmpty(propResult))
                            {
                                result.Add(upgrade["ActionProperty"].ToString(), propResult);
                            }
                        }
                        return result;
                    }
                }
            }
            finally
            {
                Installer.SetInternalUI(previous);
            }
        }
    }
}
