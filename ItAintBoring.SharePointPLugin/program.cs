using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ItAintBoring.SharePointPLugin
{
    class Program
    {
        static void Main()
        {
            //treecatsoftware.sharepoint.com
            SharePointClient cli = new SharePointClient("c3d33316-25b9-4779-8ce0-8c9445dff29b", "JrcG47KYf9XJ6eayxWMYgGuFiwuW/390FK4wXg/7Fow=", "ee810a5c-034a-47b1-926d-c745ecaf201b", "treecatsoftware.sharepoint.com");

            Task<string> task = Task.Run<string>(async () => await cli.GetToken());

            var result = task.Result;
        }
    }
}
