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
        List<long> blockUser = new List<long>();

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

        // 전체 유저 정보를 얻어옴
        public Dictionary<long, CUser> getAllUserInfo()
        {
            return userInfo;
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
            if (userInfo.ContainsKey(userKey) == false)
                return false;

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

        public List<CMessage> getMessage(long userKey)
        {
            CUser userInfo = getUserInfo(userKey);

            return userInfo.getMessage();
        }

        public void addPrivateNoti(long userKey, string userID, string notiString, int hour, int min)
        {
            CPrivateNoti privateNoti = new CPrivateNoti();
            CUser userInfo = getUserInfo(userKey);

            privateNoti.Notice = notiString;
            privateNoti.Hour = hour;
            privateNoti.Minute = min;
            privateNoti.UserID = userID;

            userInfo.addPrivateNoti(privateNoti);

            // 파일에 백업
            System.IO.File.AppendAllText(@"Data/" + "Noti_" + userKey + ".txt", hour + "|" + min + "|" + userID + "|" + notiString + "\n", Encoding.UTF8);
        }

        public List<CPrivateNoti> getPrivateNoti(long userKey)
        {
            CUser userInfo = getUserInfo(userKey);

            return userInfo.getPrivateNoti();
        }

        public void RemoveNoti(long userKey, int index)
        {
            var notiQueue = getPrivateNoti(userKey);

            notiQueue.RemoveAt(index);

            string backup = "";
            foreach (var elem in notiQueue)
            {
                backup += elem.Hour + "|" + elem.Minute + "|" + elem.UserID + "|" + elem.Notice + "\n";
            }

            // 파일에 백업
            System.IO.File.WriteAllText(@"Data/" + "Noti_" + userKey + ".txt", backup, Encoding.UTF8);
        }

        public void addMemo(long userKey, string memoString)
        {
            CMemo memo = new CMemo();
            CUser userInfo = getUserInfo(userKey);

            memo.Memo = memoString;

            userInfo.addMemo(memo);

            // 파일에 백업
            System.IO.File.AppendAllText(@"Data/" + "Memo_" + userKey + ".txt", memoString + "\n", Encoding.UTF8);
        }

        public List<CMemo> getMemo(long userKey)
        {
            CUser userInfo = getUserInfo(userKey);

            return userInfo.getMemoList();
        }

        public void RemoveMemo(long userKey, int index)
        {
            var memo = getMemo(userKey);

            memo.RemoveAt(index);

            string backup = "";
            foreach (var elem in memo)
            {
                backup += elem.Memo + "\n";
            }

            // 파일에 백업
            System.IO.File.WriteAllText(@"Data/" + "Memo_" + userKey + ".txt", backup, Encoding.UTF8);
        }

        public void addBlockUser(long userKey)
        {
            blockUser.Add(userKey);
        }

        public void removeBlockUser(long userKey)
        {
            blockUser.Remove(userKey);
        }

        public bool isBlockUser(long userKey)
        {
            if (blockUser.Count > 0 && blockUser.Contains(userKey) == true)
                return true;

            return false;
        }

        public List<CUser> getAdminUser()
        {
            List<CUser> admin = new List<CUser>();

            var allUser = getAllUserInfo();
            foreach (var iter in allUser)
            {
                if (iter.Value.UserType == USER_TYPE.USER_TYPE_ADMIN)
                {
                    admin.Add(iter.Value);
                }
            }
            
            return admin;
        }
    }
}
