using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LoadDeathsGUI
{
    public class Global
    {
        //Bad idea to declare variables as public (Global) out side of there own class.

        public static string sTheSSN
        {
            get; set;
        }

        public static bool isPrimary
        {
            get; set;
        }

        public static bool isDependent
        {
            get; set;
        }

        public static bool bIsSSNFilled
        {
            get; set;
        }

        public static bool bIsDateSelected
        {
            get; set;
        }

        public static bool bIsPrimaryDependentSelected
        {
            get; set;
        }

        public static string sTheSelectedDate
        {
            get; set;
        }

        public static bool isDeath
        {
            get; set;
        }

        public static bool isDivorce
        {
            get; set;
        }

        public static List<string> HeaderNames
        {
            get; set;
        }

        public static List<string> aCampgainSegments
        {
            get; set;
        }


    }
}
