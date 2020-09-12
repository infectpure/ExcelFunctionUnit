using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFunctionUnit
{
    [Guid("3BC99D1E-8D10-4096-B4F8-530BED98098C")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    class LeapFunction:DefineFunctionBasic 
    {
        public bool IsLeap(int year)
        {
            if (year % 4 == 0 && year % 100 != 0)
            { return true; }
            else if (year % 400 == 0)
            { return true; }   
            else
            { return false; }
        }
    }
}
