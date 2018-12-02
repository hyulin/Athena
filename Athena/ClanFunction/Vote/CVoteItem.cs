using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Athena
{
    class CVoteItem
    {
        string item_ = "";
        List<string> voter_ = new List<string>();

        // 투표 항목
        public void AddItem(string item)
        {
            item_ = item;
        }
        public string getItem()
        {
            return item_;
        }

        // 투표자
        public void AddVoter(string voter)
        {
            voter_.Add(voter);
        }
        public List<string> getVoter()
        {
            return voter_;
        }
    }
}
