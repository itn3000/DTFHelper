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
    }
}
