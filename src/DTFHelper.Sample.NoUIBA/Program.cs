using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTFHelper.Sample.NoUIBA
{
    using Microsoft.Deployment.WindowsInstaller;
    using Microsoft.Deployment.WindowsInstaller.Package;
    public enum PackageInstallationMode
    {
        InstallNew,
        MinorUpgrade,
        Maintainance,
        MajorUpgrade,
        Uninstall,
    }
    /// <summary>
    /// MSI current installation state
    /// </summary>
    public enum PackageInstallState
    {
        /// <summary>
        /// same package code exists.
        /// you must reinstall or remove.
        /// </summary>
        SamePackageCodeExisted,
        /// <summary>
        /// same product code exists,but package code is not same.
        /// you must do minor upgrade
        /// <remarks>in this state,reinstall and remove will fail</remarks>
        /// </summary>
        SameProductCodeExisted,
        /// <summary>
        /// same upgrade code exists,but product code is not same.
        /// you must do major upgrade
        /// </summary>
        RelatedProductExisted,
        /// <summary>
        /// package is not installed.
        /// </summary>
        NotInstalled,

    }
    class Program
    {
        /// <summary>
        /// get MSI package code.for more information about package code,see https://msdn.microsoft.com/en-us/library/windows/desktop/aa370568%28v=vs.85%29.aspx
        /// </summary>
        /// <param name="msiFilePath">path for MSI file</param>
        /// <returns>MSI package code</returns>
        static string GetPackageCode(string msiFilePath)
        {
            using (var sumInfo = new SummaryInfo(msiFilePath, false))
            {
                return sumInfo.RevisionNumber;
            }
        }
        static PackageInstallationMode GetInstallationMode(string msiFilePath)
        {
            var packageCode = GetPackageCode(msiFilePath);
            using (var db = new InstallPackage(msiFilePath, DatabaseOpenMode.Transact))
            {
                var session = Installer.OpenPackage(db, true);
                session.DoAction("FindRelatedProducts");
                var upgradeCode = db.Property["UpgradeCode"];
                var productCode = db.Property["ProductCode"];
                if (!string.IsNullOrEmpty(upgradeCode))
                {
                    //foreach (var related in ProductInstallation.GetRelatedProducts(upgradeCode))
                    //{
                    //    db.Tables.Select(ti =>
                    //    {
                    //    });
                    //}
                }
                throw new NotImplementedException();
            }
        }
        static void Main(string[] args)
        {
            if (args == null || args.Any())
            {
                Console.Error.WriteLine("you must specify msi file path");
            }
            var msiPath = args[0];
        }
    }
}
