using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace Athena
{
    enum TOKEN_TYPE
    {
        TOKEN_TYPE_BOT,
        TOKEN_TYPE_TEST,
        TOKEN_TYPE_SHEET,

        TOKEN_TYPE_MAX
    }

    //enum ADMIN_TYPE
    //{
    //    ADMIN_TYPE_COKE,
    //    ADMIN_TYPE_HYULIN,
    //    ADMIN_TYPE_MANS3UL,
    //    ADMIN_TYPE_LUMINOX,
    //    ADMIN_TYPE_GOGI,
        
    //    ADMIN_TYPE_MAX
    //}

    enum GROUP_TYPE
    {
        GROUP_TYPE_CLAN,
        GROUP_TYPE_ADMIN,
        GROUP_TYPE_GUIDE,

        GROUP_TYPE_MAX
    }

    class CConfig
    {
        public List<string> token_ { get; set; }
        public List<long> group_ { get; set; }
        public List<long> admin_ { get; set; }
        public List<string> admin_ID_ { get; set; }
        public long developer_ { get; set; }

        public void loadConfig()
        {
            string config = System.IO.File.ReadAllText(@"Config/config.json");

            JObject json = JObject.Parse(config);

            if (token_ != null && token_.Count > 0)
                token_.Clear();

            List<string> token = new List<string>();
            token.Add(json["token"]["Bot"].ToString());
            token.Add(json["token"]["Test"].ToString());
            token.Add(json["token"]["Sheet"].ToString());
            token_ = token;

            if (admin_ != null && admin_.Count > 0)
                admin_.Clear();

            if (group_ != null && group_.Count > 0)
                group_.Clear();

            List<long> group = new List<long>();
            group.Add((long)json["group"]["클랜방"]);
            group.Add((long)json["group"]["운영진방"]);
            group.Add((long)json["group"]["사전안내방"]);
            group_ = group;

            List<long> admin = new List<long>();
            foreach (var user in json["admin"])
            {
                admin.Add((long)user.ElementAt(0));
            }
            admin_ = admin;

            List<string> admin_id = new List<string>();
            foreach (var user in json["admin_id"])
            {
                admin_id.Add(user.ElementAt(0).ToString());
            }
            admin_ID_ = admin_id;

            developer_ = (long)json["developer"];
        }

        public string getTokenKey(TOKEN_TYPE type)
        {
            return token_.ElementAt((int)type);
        }

        public bool isAdmin(long userKey)
        {
            return admin_.Contains(userKey);
        }

        public bool isDeveloper(long userKey)
        {
            if (developer_ == userKey)
                return true;

            return false;
        }

        public long getGroupKey(GROUP_TYPE type)
        {
            return group_.ElementAt((int)type);
        }

        public GROUP_TYPE getGroupType(long chatID)
        {
            GROUP_TYPE rtType = GROUP_TYPE.GROUP_TYPE_CLAN;

            foreach (var iter in group_)
            {
                if (iter == chatID)
                {
                    return rtType;
                }

                rtType++;
            }

            return GROUP_TYPE.GROUP_TYPE_MAX;
        }
    }
}
