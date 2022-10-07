using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template4435
{
    public partial class Clients
    {
        public int Age
        {
            get
            {
                if (date_birth != null)
                {
                    return DateTime.Now.Year - date_birth.Value.Year;
                }
                else return -1;
            }
        }
    }

    public partial class ClientsJSON
    {
        public int Age
        {
            get
            {
                if (BirthDate != null)
                {
                    return DateTime.Now.Year - BirthDate.Value.Year;
                }
                else return -1;
            }
        }
    }
}
