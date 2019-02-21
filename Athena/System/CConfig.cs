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

    enum ADMIN_TYPE
    {
        ADMIN_TYPE_COKE,
        ADMIN_TYPE_HYULIN,
        ADMIN_TYPE_MANS3UL,
        ADMIN_TYPE_LUMINOX,
        ADMIN_TYPE_GOGI,
        
        ADMIN_TYPE_MAX
    }

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
        public List<long> admin_ { get; set; }
        public List<long> group_ { get; set; }

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

            List<long> admin = new List<long>();
            admin.Add((long)json["admin"]["냉각콜라"]);
            admin.Add((long)json["admin"]["휴린"]);
            admin.Add((long)json["admin"]["만슬"]);
            admin.Add((long)json["admin"]["루미녹스"]);
            admin.Add((long)json["admin"]["고기"]);
            admin_ = admin;

            if (group_ != null && group_.Count > 0)
                group_.Clear();

            List<long> group = new List<long>();
            group.Add((long)json["group"]["클랜방"]);
            group.Add((long)json["group"]["운영진방"]);
            group.Add((long)json["group"]["사전안내방"]);
            group_ = group;
        }

        public string getTokenKey(TOKEN_TYPE type)
        {
            return token_.ElementAt((int)type);
        }

        public long getAdminKey(ADMIN_TYPE type)
        {
            return admin_.ElementAt((int)type);
        }

        public long getGroupKey(GROUP_TYPE type)
        {
            return group_.ElementAt((int)type);
        }

        // Todo : 방 등록 시에 config 및 json 파일 자동 수정
    }
}
