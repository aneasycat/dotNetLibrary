using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace AnEasyCat.Office
{
    public class Excel
    {
        string _file="";
        public Excel(string File)
        {
            _file = File;
        }
    }
}
