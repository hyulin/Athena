using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Athena
{
    enum USER_TYPE
    {
        USER_TYPE_NONE,

        USER_TYPE_NORMAL,   // 일반 유저
        USER_TYPE_ADMIN,    // 관리자
        USER_TYPE_BOT,      // 봇

        USER_TYPE_MAX
    }

    enum POSITION
    {
        POSITION_NONE = 0x000,

        POSITION_DPS = 0x001,   // 딜러
        POSITION_TANK = 0x010,  // 탱커
        POSITION_SUPP = 0x100,  // 힐러

        POSITION_FLEX = 0x111   // 플렉스
    }

    class CUser
    {
        public long UserKey { get; set; }
        public string Name { get; set; }
        public USER_TYPE UserType { get; set; }
        public string MainBattleTag { get; set; }
        public string[] SubBattleTag { get; set; }
        public POSITION Position { get; set; }
        public string[] MostPick { get; set; }
        public string OtherPick { get; set; }
        public string Time { get; set; }
        public string Info { get; set; }
        public ulong chattingCount { get; set; }
        
        List<CMessage> MessageList = new List<CMessage>();
        List<CPrivateNoti> PrivateNoti = new List<CPrivateNoti>();
        List<CMemo> MemoList = new List<CMemo>();

        public void addMessage(CMessage message)
        {
            if (MessageList.Count < 100)
            {
                MessageList.Add(message);
            }
            else
            {
                MessageList.RemoveAt(MessageList.Count - 1);
                MessageList.Add(message);
            }
        }

        public void addPrivateNoti(CPrivateNoti privateNoti)
        {
            if (PrivateNoti.Count < 10)
            {
                PrivateNoti.Add(privateNoti);
            }
            else
            {
                PrivateNoti.RemoveAt(PrivateNoti.Count - 1);
                PrivateNoti.Add(privateNoti);
            }
        }

        public void addMemo(CMemo memo)
        {
            if (MemoList.Count < 100)
            {
                MemoList.Add(memo);
            }
            else
            {
                MemoList.RemoveAt(MemoList.Count - 1);
                MemoList.Add(memo);
            }
        }

        public List<CMessage> getMessage()
        {
            return MessageList;
        }

        public List<CPrivateNoti> getPrivateNoti()
        {
            return PrivateNoti;
        }

        public List<CMemo> getMemoList()
        {
            return MemoList;
        }
    }
}
