using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Deployment.WindowsInstaller;

namespace DTFHelper.Extensions
{
    static class MsiDatabaseExtension
    {
        public static IEnumerable<IDictionary<string, object>> ExecuteQueryToDictionary(this Database db, string[] columns, string sqlquery, params object[] param)
        {
            return db.ExecuteQuery("SELECT {0} FROM Upgrade"
                , string.Join(",", columns)).Cast<object>().Buffer(columns.Count())
                .Select(x => x.Select((field, i) => new { field, i }).ToDictionary(field => columns[field.i], field => field.field))
                ;
        }
    }
}
