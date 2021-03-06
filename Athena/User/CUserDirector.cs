﻿using System;
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

        // 대화명으로 유저 정보 얻어옴
        public CUser getUserInfoByName(string name)
        {
            foreach (var iter in userInfo)
            {
                if (iter.Value.Name.Contains(name) == true)
                {
                    return iter.Value;
                }
            }

            CUser tempUserInfo = new CUser();
            return tempUserInfo;
        }

        // 유저 정보 갱신
        public bool refleshUserInfo(long userKey, CUser user)
        {
            if (userInfo.ContainsKey(userKey) == false)
                return false;

            userInfo[userKey].Name = user.Name;
            userInfo[userKey].MainBattleTag = user.MainBattleTag;
            userInfo[userKey].SubBattleTag = user.SubBattleTag;
            userInfo[userKey].Position = user.Position;
            userInfo[userKey].MostPick = user.MostPick;
            userInfo[userKey].OtherPick = user.OtherPick;
            userInfo[userKey].Team = user.Team;
            userInfo[userKey].Youtube = user.Youtube;
            userInfo[userKey].Twitch = user.Twitch;
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
            System.IO.File.AppendAllText(@"Data/Noti/" + "Noti_" + userKey + ".txt", hour + "|" + min + "|" + userID + "|" + notiString + "\n", Encoding.UTF8);
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
            System.IO.File.WriteAllText(@"Data/Noti/" + "Noti_" + userKey + ".txt", backup, Encoding.UTF8);
        }

        public void addMemo(long userKey, string memoString)
        {
            CMemo memo = new CMemo();
            CUser userInfo = getUserInfo(userKey);

            memo.Memo = memoString;

            userInfo.addMemo(memo);

            // 파일에 백업
            System.IO.File.AppendAllText(@"Data/Memo/" + "Memo_" + userKey + ".txt", memoString + "\n", Encoding.UTF8);
        }

        public List<CMemo> getMemo(long userKey)
        {
            CUser userInfo = getUserInfo(userKey);

            return userInfo.getMemoList();
        }

        public bool RemoveMemo(long userKey, int index)
        {
            var memo = getMemo(userKey);

            if (memo.Count - 1 < index)
                return false;

            memo.RemoveAt(index);

            string backup = "";
            foreach (var elem in memo)
            {
                backup += elem.Memo + "\n";
            }

            // 파일에 백업
            System.IO.File.WriteAllText(@"Data/Memo/" + "Memo_" + userKey + ".txt", backup, Encoding.UTF8);

            return true;
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

        public void increaseChattingCount(long userKey)
        {
            CUser userInfo = getUserInfo(userKey);
            if (userInfo.UserKey > 0)
            {
                userInfo.chattingCount++;

                // 파일에 백업
                System.IO.File.WriteAllText(@"Data/Chatting/" + "ChattingCount_" + userKey.ToString() + ".txt", userInfo.chattingCount.ToString() + "\n", Encoding.UTF8);
            }
        }

        public void setChattingCount(long userKey, ulong count)
        {
            CUser userInfo = getUserInfo(userKey);
            if (userInfo.UserKey > 0)
            {
                userInfo.chattingCount = count;
            }
        }

        public void resetChattingCount(long userKey = 0)
        {
            if (userKey == 0)
            {
                foreach (var iter in userInfo)
                {
                    iter.Value.chattingCount = 0;
                }

                if (System.IO.Directory.Exists(@"Data/Chatting/"))
                {
                    string[] files = System.IO.Directory.GetFiles(@"Data/Chatting/");

                    foreach (string file in files)
                    {
                        string fileName = System.IO.Path.GetFileName(file);
                        string deletefile = @"Data/Chatting/" + fileName;
                        System.IO.File.Delete(deletefile);
                    }
                }

                return;
            }

            CUser user = getUserInfo(userKey);
            if (user.UserKey > 0)
            {
                user.chattingCount = 0;
            }

            string filePath = @"Data/Chatting/" + "ChattingCount_" + userKey.ToString() + ".txt";
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
        }
    }
}
