using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Athena
{
    class CNotice
    {
        string strNotice_ = "";

        public string GetNotice()
        {
            return strNotice_;
        }

        public void SetNotice(string strNotice)
        {
            strNotice_ = strNotice;
        }

        public void AppendNotice(string strNotice)
        {
            strNotice_ += strNotice;
        }
    }
}
