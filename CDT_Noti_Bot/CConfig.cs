using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace CDT_Noti_Bot
{
    class CConfig
    {
        JObject config = JObject.Parse(System.IO.File.ReadAllText(@"config.json"));

        public string getRealBotToken()
        {
            return config["real_bot_token"].ToString();
        }
    }
}
