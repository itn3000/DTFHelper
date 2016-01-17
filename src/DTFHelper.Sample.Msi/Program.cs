using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTFHelper.Sample.Msi
{
    using WixSharp;
    using WixSharp.CommonTasks;
    class Program
    {
        static void Main(string[] args)
        {
            var project = new Project("DTFHelperSample");
            project.AddDir(new Dir("%ProgramFiles%\\DTFHelperSample"
                , new File("TextFile1.txt")));
            project.GUID = new Guid("{7F9D7171-087C-47C9-91C3-503DBCAED2DD}");
            project.ProductId = new Guid("{0FF8824C-5297-4F18-8566-D7E35407F84E}");
            project.UpgradeCode = new Guid("{E748FC44-0E0A-4783-8DB5-CEAD1BEAB602}");
            Compiler.BuildMsi(project);
        }
    }
}
