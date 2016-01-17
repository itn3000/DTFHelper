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
    class Program
    {
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
