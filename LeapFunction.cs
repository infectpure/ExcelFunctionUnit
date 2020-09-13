using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFunctionUnit
{
    [Guid("6234288A-B724-481D-9D71-81279013B46D")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    class LeapFunction:DefineFunctionBasic 
    {
        public LeapFunction():base()
        {

        }
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
