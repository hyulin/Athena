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

        string[] notiCommand = { "공지" };
        string[] refCommand = { "조회", "전적", "점수", "티어" };
        string[] videoCommand = { "영상", "방송" };
        string[] serchCommand = { "검색", "모스트", "포지션", "유저" };
        string[] meetingCommand = { "모임", "정모", "참가", "확정", "불참" };
        string[] voteCommand = { "투표", "선거", "설문" };
        string[] recordCommand = { "기록", "명예", "전당" };
        string[] guideCommand = { "안내", "가이드" };
        string[] statusCommand = { "상태" };

        string[] enterCommand = { "알려", "보여", "?", "궁금", "해줘", "뭐야" };
        string[] ofCommand = { "의", "가", "에", "은", "는", "님의", "님" };
        string[] mindCommand = { "어때", "?", "는요", "어떰", "어떨", "어떠", "어떻", "어떤" };

        string[] menu = {
            "백반", "떡볶이", "순대", "김밥", "짜장면", "짬뽕", "볶음밥", "김치찌개", "육개장", "된장찌개", "제육볶음",
            "설렁탕", "회덮밥", "냉면", "돈까스", "함박스테이크", "사골국밥", "순대국밥", "갈비탕", "라면", "라멘", "카레",
            "치킨", "피자", "파스타", "육회비빔밥", "비빔밥", "보쌈", "족발", "막국수", "냉모밀", "소바", "스시", "햄버거",
            "한솥", "삼겹살", "소고기", "곱창", "삼계탕", "양념갈비", "스테이크", "생선구이", "훈제오리", "샐러드", "만두"
        };

        string[] enterMenu = {
            "어떠세요?", "추천합니다.", "땡기네요.", "가보시죠.", "ㄱㄱ", "각이네요", "좋네요.",
            "어떤가요?", "좋을 듯.", "가시죠.", "?", "!", "너로 정했다!", "기대합니다."
        };

        string[] choiceWord = { "이랑", "랑", "," };

        string[] offWork = { "퇴근합니다", "퇴근~", "퇴근!", "퇴근하겠" };

        // 메뉴 추천 감지
        public bool isExistMenu(string message)
        {
            if (message.Contains("뭐") && message.Contains("먹"))
                return true;

            if (message.Contains("뭘") && message.Contains("먹"))
                return true;

            if (message.Contains("어떤") && message.Contains("먹"))
                return true;

            if (message.Contains("밥") && message.Contains("먹"))
                return true;

            if (message.Contains("점심") && message.Contains("먹"))
                return true;

            if (message.Contains("저녁") && message.Contains("먹"))
                return true;

            if (message.Contains("배고픈데") || message.Contains("배고프네") || message.Contains("배고프다") ||
                message.Contains("배고파") || message.Contains("배고픔"))
                return true;
            
            if (message.Contains("메뉴") == true)
            {
                foreach (var word in enterCommand)
                {
                    if (message.Contains(word) == true)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        // 메뉴 추천
        public string getMenu()
        {
            Random menuRandom = new Random();
            Random enterRandom = new Random();

            int menuNum = menuRandom.Next(0, menu.Count());
            int enterNum = enterRandom.Next(0, enterMenu.Count());

            return menu.ElementAt(menuNum) + " " + enterMenu.ElementAt(enterNum);
        }

        // 날씨 감지
        public Tuple<string, string> weatherCall(string message)
        {
            Tuple<string, string> emptyTuple = Tuple.Create("", "");

            if (message.Contains("날씨") == false || message.Contains("내일") == true || message.Contains("어제") == true ||
                message.Contains("모레") == true || message.Contains("엊그제") == true)
            {
                return emptyTuple;
            }

            bool isQuestion = false;
            foreach (var word in mindCommand)
            {
                if (message.Contains(word) == true)
                {
                    isQuestion = true;
                    break;
                }
            }

            if (isQuestion == false)
            {
                foreach (var word in enterCommand)
                {
                    if (message.Contains(word) == true)
                    {
                        isQuestion = true;
                        break;
                    }
                }
            }
            
            if (isQuestion == true)
            {
                Tuple<string, string> tuple = CWeather.getCity(message);
                return tuple;
            }
            
            return emptyTuple;
        }

        // 웃기 감지
        public string laughCall(string message)
        {
            string laughMessage = "";

            if (message.Contains("ㅋ") == true)
            {
                Random random = new Random();
                int randomNum = random.Next(10);
                
                if (randomNum < 3)
                {
                    randomNum = random.Next(3);

                    if (randomNum == 0)
                        laughMessage += "재미있네요ㅎㅎ";
                    else if (randomNum == 1)
                        laughMessage += "재미있어하시니 저도 좋네요ㅎㅎ";
                    else if (randomNum == 2)
                        laughMessage += "빵 터졌어요ㅎㅎ";
                }
                else if(randomNum > 7)
                {
                    randomNum = random.Next(3);

                    for (int i = 1; i < randomNum + 1; i++)
                    {
                        laughMessage += "ㅋㅋ";
                    }
                }
            }

            return laughMessage;
        }

        // 대답하기
        public string replyCall(string message)
        {
            string reply = "";

            Random random = new Random();
            int randomNum = random.Next(10);

            switch (randomNum)
            {
                case 0:
                    reply = "아, 그래요?";
                    break;
                case 1:
                    reply = "정말요?";
                    break;
                case 2:
                    reply = "그렇군요.";
                    break;
                case 3:
                    reply = "아하~";
                    break;
                case 4:
                    reply = ".........";
                    break;
                case 5:
                    reply = "아, 네.";
                    break;
                case 6:
                    reply = "네네~";
                    break;
                case 7:
                    reply = "으잉??";
                    break;
                case 8:
                    reply = "흠...";
                    break;
                case 9:
                    reply = "헐...";
                    break;
            }

            return reply;
        }

        public string offWorkCall(string message)
        {
            string output = "";

            foreach (var word in offWork)
            {
                if (message.Contains(word) == true)
                {
                    Random random = new Random();
                    int randomNum = random.Next(5);

                    switch (randomNum)
                    {
                        case 0:
                            return "수고하셨습니다ㅎㅎ";
                        case 1:
                            return "수고 많으셨어요. :)";
                        case 2:
                            return "고생하셨습니다~";
                        case 3:
                            return "고생 많으셨어요~";
                        case 4:
                            return "";
                    }
                }
            }

            return output;
        }

        // 클랜 기능 감지
        public string FunctionCommand(string text)
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
                            return "/공지";
                        }
                    }
                }
            }



            //--------------------------------------------------------
            // 일정 감지
            //--------------------------------------------------------
            if (text.Contains("일정") == true)
            {
                foreach (var word in enterCommand)
                {
                    if (text.Contains(word) == true)
                        return "/일정";
                }

                foreach (var word in mindCommand)
                {
                    if (text.Contains(word) == true)
                        return "/일정";
                }
            }



            //--------------------------------------------------------
            // 뽑기 감지
            //--------------------------------------------------------
            if (text.Contains("중에") == true)
            {
                string contents = "";
                
                foreach (var word in split)
                {
                    if (word.Contains("아테나") == true)
                        continue;

                    if (word.Contains("중에") == true)
                        break;

                    string[] spliteItem = word.Split(',');

                    foreach (var item in spliteItem)
                    {
                        string replaceItem = item;

                        foreach (var choice in choiceWord)
                        {
                            replaceItem = replaceItem.Replace(choice, "");
                        }

                        contents += " " + replaceItem;
                    }
                }

                return "/뽑기" + contents;
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

                    foreach (var enter in mindCommand)
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
