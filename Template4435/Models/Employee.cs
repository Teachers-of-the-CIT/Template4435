using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Template4435
{
    public partial class Emloyee
    {


        [JsonPropertyName("Position")]
        public string Post { get; set; }
        [JsonPropertyName("FullName")]
        public string FIO { get; set; }
        [JsonPropertyName("Log")]
        public string Login { get; set; }
        [JsonPropertyName("Password")]
        public string Password { get; set; }
        [JsonPropertyName("LastEnter")]
        public string LastEnt { get; set; }
        [JsonPropertyName("TypeEnter")]
        public string Ent { get; set; }

        public string CodeStaff { set
            {
                Id = Convert.ToInt32(value.Remove(0, 3));
            }
        
        }
        [JsonPropertyName("NoId")]
        public int Id
        {
            get;set;
        }
    }
}
