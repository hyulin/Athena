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
        protected int MessageCount { get; set; }

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

        Queue<CMessage> MessageQueue = new Queue<CMessage>();

        public void addMessage(CMessage message)
        {
            if (MessageCount < 10)
            {
                MessageCount++;
                MessageQueue.Enqueue(message);
            }
            else
            {
                MessageQueue.Dequeue();
                MessageQueue.Enqueue(message);
            }
        }

        public Queue<CMessage> getMessage()
        {
            return MessageQueue;
        }
    }
}
