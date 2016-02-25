using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DTFHelper.Test.Properties;
using Xunit;
using Microsoft.Deployment.WindowsInstaller;
using Microsoft.Deployment.WindowsInstaller.Package;

namespace DTFHelper.Test
{
    using System.IO;
    using DTFHelper;
    public class TestMsiInstaller
    {
        [Fact]
        public void TestGetRelatedProduct()
        {
            var tmpPath = Path.GetTempFileName();
            try
            {
                File.WriteAllBytes(tmpPath, Resources.DTFHelperSample);
                var ret = MsiInstaller.GetRelatedProducts(tmpPath);
                Assert.NotNull(ret);
            }
            finally
            {
                try
                {
                    File.Delete(tmpPath);
                }
                catch { }
            }
        }
        [Fact]
        public void TestInstallUninstall()
        {
            var tmpPath = Path.GetTempFileName();
            try
            {
                File.WriteAllBytes(tmpPath, Resources.DTFHelperSample);
                string productCode = "";
                using (var db = new InstallPackage(tmpPath, DatabaseOpenMode.ReadOnly))
                {
                    productCode = db.Property["ProductCode"];
                }
                var installer = new MsiInstaller(tmpPath);
                installer.OnLogging += (sender, ev) =>
                {
                    System.Diagnostics.Trace.WriteLine(string.Format("msg = {0}", new { ev.FormattedMessage, ev.LogMode }));
                    System.Diagnostics.Trace.WriteLine(string.Format("records = {0}", string.Join(" ", ev.Parameters.Select(x => "'" + x + "'"))));
                    return MessageResult.None;
                };
                installer.OnProgress += (sender, ev) =>
                {
                    System.Diagnostics.Trace.WriteLine(string.Format("onprogress = {0}"
                        , new { ev.CurrentTicks, ev.IsForward, ev.TotalProgress }));
                    Console.WriteLine(string.Format("onprogress = {0}"
                        , new { ev.CurrentTicks, ev.IsForward, ev.TotalProgress }));
                    return MessageResult.None;
                };
                installer.ExecuteInstall();
                using (var session = Installer.OpenProduct(productCode))
                {
                    System.Diagnostics.Trace.WriteLine(string.Format("productcode = {0},installed feature = {1}"
                        , productCode
                        , string.Join(",", session.Features.Select(x => x.Name))));
                }
                installer.ExecuteUninstall();
            }
            finally
            {
                try
                {
                    File.Delete(tmpPath);
                }
                catch { }
            }
        }
        [Fact]
        public void TestAdministrativeInstall()
        {
            var tmpPath = Path.GetTempFileName();
            var tmpDir = Path.Combine(Path.GetTempPath(),"TestAdministrativeInstall");
            try
            {
                Directory.CreateDirectory(tmpDir);
                File.WriteAllBytes(tmpPath, Resources.DTFHelperSample);
                var installer = new MsiInstaller(tmpPath);
                installer.ExecuteAdministrativeInstall(tmpDir);
            }
            finally
            {
                try
                {
                    File.Delete(tmpPath);
                }
                catch { }
                try
                {
                    Directory.Delete(tmpDir, true);
                }
                catch { }
            }
        }
    }
}
