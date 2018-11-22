using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CDT_Noti_Bot
{
    class CCalendarDirector
    {
        CCalendar[] calendar_ = new CCalendar[35];

        public void addCalendar(CCalendar calendar)
        {
            int index = calendar.Day;

            calendar_[index].Day = calendar.Day;
            calendar_[index].WEEK = calendar.WEEK;
            calendar_[index].TODO = calendar.TODO;
        }

        public CCalendar getCalendar(int day)
        {
            if (day > 0 && day < 32)
            {
                return calendar_[day];
            }

            CCalendar empty = new CCalendar();
            return empty;
        }
    }
}
