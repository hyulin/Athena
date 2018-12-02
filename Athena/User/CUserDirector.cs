using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Athena
{
    class CUserDirector
    {
        // 유저 딕셔너리
        Dictionary<long, CUser> userInfo = new Dictionary<long, CUser>();
        int userCount = 0;

        // 유저 추가
        public void addUserInfo(long userKey, CUser user)
        {
            if (userKey <= 0)
            {
                // 유저 키가 없음
                return;
            }

            userInfo.Add(userKey, user);
            userCount++;
        }

        // 유저키로 유저 정보 얻어옴
        public CUser getUserInfo(long userKey)
        {
            if (userInfo.ContainsKey(userKey) == true)
            {
                return userInfo[userKey];
            }

            CUser tempUserInfo = new CUser();
            return tempUserInfo;
        }

        // 유저 정보 갱신
        public bool reflechUserInfo(long userKey, CUser user)
        {
            userInfo[userKey].Name = user.Name;
            userInfo[userKey].MainBattleTag = user.MainBattleTag;
            userInfo[userKey].SubBattleTag = user.SubBattleTag;
            userInfo[userKey].Position = user.Position;
            userInfo[userKey].MostPick = user.MostPick;
            userInfo[userKey].OtherPick = user.OtherPick;
            userInfo[userKey].Time = user.Time;
            userInfo[userKey].Info = user.Info;

            return true;
        }

        public int getUserCount()
        {
            return userCount;
        }

        public void addMessage(long userKey, string message, DateTime time)
        {
            CMessage userMessage = new CMessage();
            CUser userInfo = getUserInfo(userKey);

            userMessage.Message = message;
            userMessage.Time = time;

            userInfo.addMessage(userMessage);
        }
    }
}
