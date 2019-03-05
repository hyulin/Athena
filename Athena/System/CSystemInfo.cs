using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Athena
{
    class CSystemInfo
    {
        DateTime startTime_ = new DateTime();

        public string GetNowTime()
        {
            return System.DateTime.Now.ToString();
        }

        public string GetStartTime()
        {
            return startTime_.ToString();
        }

        public DateTime GetStartTimeToDate()
        {
            return startTime_;
        }
        
        public DayOfWeek getNowDayOfWeek()
        {
            return System.DateTime.Now.DayOfWeek;
        }

        public string GetRunningTime()
        {
            DateTime nowTime = System.DateTime.Now;
            TimeSpan runningTime = nowTime - startTime_;

            string strValue = runningTime.Days + "일 " + runningTime.Hours + "시간 " + runningTime.Minutes + "분 " + runningTime.Seconds + "초";

            return strValue;
        }

        public void SetStartTime()
        {
            startTime_ = DateTime.Now;
        }
    }
}
