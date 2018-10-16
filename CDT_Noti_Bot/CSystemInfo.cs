using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CDT_Noti_Bot
{
    class CSystemInfo
    {
        DateTime startTime_ = new DateTime();
        ulong googleSheetReqCount_ = 0;
        ulong messageReqCount_ = 0;

        public string GetNowTime()
        {
            return System.DateTime.Now.ToString();
        }

        public string GetStartTime()
        {
            return startTime_.ToString();
        }

        public string GetRunningTime()
        {
            DateTime nowTime = System.DateTime.Now;
            TimeSpan runningTime = nowTime - startTime_;

            string strValue = runningTime.Days + "일 " + runningTime.Hours + "시간 " + runningTime.Minutes + "분 " + runningTime.Seconds + "초";

            return strValue;
        }

        public ulong GetGoogleSheetReqCount()
        {
            return googleSheetReqCount_;
        }

        public ulong GetMessageReqCount()
        {
            return messageReqCount_;
        }

        public void SetStartTime()
        {
            startTime_ = DateTime.Now;
        }

        public void AppendGoogleSheetCount()
        {
            googleSheetReqCount_++;
        }

        public void AppendMessageReqCount()
        {
            messageReqCount_++;
        }
    }
}
