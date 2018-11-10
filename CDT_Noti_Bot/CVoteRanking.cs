using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CDT_Noti_Bot
{
    class CVoteRanking
    {
        int ranking_ = 0;
        string voteItem_ = "";
        int voteCount_ = 0;
        string voteRate_ = "";

        public void setRanking(int ranking, string voteItem, int voteCount, string voteRate)
        {
            ranking_ = ranking;
            voteItem_ = voteItem;
            voteCount_ = voteCount;
            voteRate_ = voteRate;
        }

        public string getVoteItem()
        {
            return voteItem_;
        }

        public int getVoteCount()
        {
            return voteCount_;
        }

        public string getVoteRate()
        {
            return voteRate_;
        }

        public int getRanking()
        {
            return ranking_;
        }
    }
}
