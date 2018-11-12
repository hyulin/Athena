using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CDT_Noti_Bot
{
    class CNaturalLanguage
    {
        //string[] command = {"공지", "조회", "영상", "검색", "모임", "투표", "기록", "안내", "상태"};

        string[] notiCommand = {"공지"};
        string[] refCommand = {"조회", "전적", "점수", "티어"};
        string[] videoCommand = {"영상", "방송"};
        string[] serchCommand = {"검색", "모스트", "포지션", "유저"};
        string[] meetingCommand = {"모임", "정모", "참가", "확정", "불참"};
        string[] voteCommand = {"투표", "선거", "설문"};
        string[] recordCommand = {"기록", "명예", "전당"};
        string[] guideCommand = {"안내", "가이드"};
        string[] statusCommand = {"상태"};

        string[] enterCommand = {"알려", "보여", "?", "궁금", "해줘"};
        string[] ofCommand = {"의", "가", "에", "은", "는"};

        public string DetectionCommand(string text)
        {
            string[] split = text.Split(' ');
            string retCommand = "";

            //--------------------------------------------------------
            // 공지 감지
            //--------------------------------------------------------
            bool isNotiCommand = false;

            foreach (var command in notiCommand)
            {
                foreach (var word in split)
                {
                    if (word.Contains(command))
                    {
                        isNotiCommand = true;
                        break;
                    }
                }

                if (isNotiCommand == true)
                {
                    break;
                }
            }

            if (isNotiCommand == true)
            {
                foreach (var word in split)
                {
                    foreach (var enter in enterCommand)
                    {
                        if (word.Contains(enter))
                        {
                            retCommand = "/공지";
                            return retCommand;
                        }
                    }
                }
            }
            


            //--------------------------------------------------------
            // 조회 감지
            //--------------------------------------------------------
            int battleTagIndex = 0;
            int index = 0;
            bool isRefCommand = false;

            foreach (var word in split)
            {
                foreach (var command in refCommand)
                {
                    if (word.Contains(command))
                    {
                        battleTagIndex = index;
                        isRefCommand = true;
                        break;
                    }
                }

                if (battleTagIndex > 0)
                {
                    break;
                }

                index++;
            }

            if (isRefCommand == true)
            {
                battleTagIndex--;

                string battleTag = split[battleTagIndex];

                foreach (var of in ofCommand)
                {
                    if (battleTag.IndexOf(of) == battleTag.Length - 1)
                    {
                        battleTag = battleTag.Substring(0, battleTag.Length - 1);
                        break;
                    }
                }

                foreach (var word in split)
                {
                    foreach (var enter in enterCommand)
                    {
                        if (word.Contains(enter))
                        {
                            return "/조회 " + battleTag.Trim();
                        }
                    }

                }
            }            

            return retCommand;
        }
    }
}
