using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CDT_Noti_Bot
{
    enum DAY_WEEK
    {
        DAY_WEEK_SUNDAY,
        DAY_WEEK_MONDAY,
        DAY_WEEK_TUESDAY,
        DAY_WEEK_WEDNESDAY,
        DAY_WEEK_THURSDAY,
        DAY_WEEK_FRIDAY,
        DAY_WEEK_SATURDAY,

        DAY_WEEK_MAX
    }

    class CCalendar
    {
        public int Day { get; set; }
        public DAY_WEEK WEEK { get; set; }
        public string TODO { get; set; }
    }
}
