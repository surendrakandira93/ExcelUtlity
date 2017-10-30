using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelUtility.Models
{
    public class Class : Interface1, Interface

    {
        public int sum(int x, int y)
        {
            return x + y;
        }
    }
}