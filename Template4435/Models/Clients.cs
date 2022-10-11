using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Template4435
{
    public partial class Clients
    {
        [JsonPropertyName("Id")]
        public int id { get; set; }
        [JsonPropertyName("FullName")]
        public string FIO { get; set; }
        [JsonPropertyName("BirthDate")]
        public Nullable<System.DateTime> date_birth { get; set; }
        [JsonPropertyName("Index")]
        public string adress_index { get; set; }
        [JsonPropertyName("City")]
        public string adress_gorod { get; set; }
        [JsonPropertyName("Street")]
        public string adress_street { get; set; }
        [JsonPropertyName("Home")]
        public int adress_house { get; set; }
        [JsonPropertyName("Kvartira")]
        public int adress_flat { get; set; }
        [JsonPropertyName("E_mail")]
        public string email { get; set; }
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

        public string Category
        {
            get
            {
                if (Age >= 20 && Age <= 29)
                {
                    return "от 20 до 29";
                }
                if (Age >= 30 && Age <= 39)
                {
                    return "от 30 до 39";
                }
                if (Age >= 40)
                {
                    return "от 40";
                }
                else return "";
            }
        }
    }
}
