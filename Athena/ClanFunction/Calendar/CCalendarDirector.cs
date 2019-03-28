using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Athena
{
    class CCalendarDirector
    {
        Dictionary<DateTime, CCalendar> calenderDictionary_ = new Dictionary<DateTime, CCalendar>();

        public void addCalendar(CCalendar calendar)
        {            
            calenderDictionary_.Add(calendar.Time, calendar);
        }

        public CCalendar getCalendar(int year, int month, int day)
        {
            DateTime time = new DateTime(year, month, day);

            if (calenderDictionary_.ContainsKey(time) == true)
            {
                return calenderDictionary_[time];
            }

            CCalendar emptyCalendar = new CCalendar();
            return emptyCalendar;
        }

        public Dictionary<DateTime, CCalendar> getCalendar()
        {
            return calenderDictionary_;
        }

        public int getCalendarCount()
        {
            return calenderDictionary_.Count;
        }
    }
}
