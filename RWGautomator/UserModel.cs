using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TRVautomator
{
    public class UserModel
    {
        [JsonProperty(PropertyName = "id")]
        public string id { get; set; }
        public string password { get; set; }
        public bool IsTraining { get; set; } = false;
        public bool IsAdmin { get; set; } = false;
        public override string ToString()
        {
            return JsonConvert.SerializeObject(this);
        }
    }
}
