using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DTFHelper.Test.Properties;
using Xunit;

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
