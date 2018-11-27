using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CDT_Noti_Bot
{
    class CCalendarDirector
    {
        Dictionary<int, CCalendar> calenderDictionary_ = new Dictionary<int, CCalendar>();

        public void addCalendar(CCalendar calendar)
        {
            int index = calendar.Day;

            calenderDictionary_.Add(index, calendar);
        }

        public CCalendar getCalendar(int day)
        {
            if (calenderDictionary_.ContainsKey(day) == true)
            {
                return calenderDictionary_[day];
            }

            CCalendar emptyCalendar = new CCalendar();
            return emptyCalendar;
        }

        public int getCalendarCount()
        {
            return calenderDictionary_.Count;
        }
    }
}
