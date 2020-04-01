using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServiceOutlook.Service
{
    class ServiceTest : IServiceTest
    {
        public string GenerateSqlSelect()
        {
            return "Вроде запустили";
        }
    }
}
