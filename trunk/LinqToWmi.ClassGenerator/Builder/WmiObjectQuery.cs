using System;
using System.Collections.Generic;
using System.Text;
using System.Management;
using System.Linq;

namespace LinqToWmi.ProtoGenerator
{
    class WmiMetaInformation
    {
        public ManagementClass GetMetaInformation(string name)
        {
            ManagementScope scope = new ManagementScope(@"\\localhost");

            ManagementPath path = new ManagementPath(name);
            ManagementClass management = new ManagementClass(scope, path, null);
            
            return management;
        }
    }
}
