using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Deployment.WindowsInstaller;
using Microsoft.Deployment.WindowsInstaller.Package;
using DTFHelper.Extensions;

namespace DTFHelper
{
    public class MsiInstaller
    {
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
