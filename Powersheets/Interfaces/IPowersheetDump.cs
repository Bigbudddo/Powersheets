using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {

    public interface IPowersheetDump {

        Dictionary<string, string> Columns { get; set; }
    }
}
