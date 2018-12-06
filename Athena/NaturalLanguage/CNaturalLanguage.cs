using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// Twitter Korean Processor
using Moda.Korean.TwitterKoreanProcessorCS;

namespace Athena
{
    class CNaturalLanguage
    {
        //string[] command = {"공지", "조회", "영상", "검색", "모임", "투표", "기록", "안내", "상태"};

        string[] notiCommand = { "공지" };
        string[] refCommand = { "조회", "전적", "점수", "티어" };
        string[] videoCommand = { "영상", "방송" };
        string[] videoSubCommand = { "목록", "리스트", "어떤" };
        string[] serchCommand = { "검색", "모스트", "포지션", "유저" };
        string[] meetingCommand = { "모임", "정모", "참가", "확정", "불참" };
        string[] voteCommand = { "투표", "선거", "설문" };
        string[] recordCommand = { "기록", "명예", "전당" };
        string[] guideCommand = { "안내", "가이드" };
        string[] statusCommand = { "상태" };

        string[] enterCommand = { "알려", "보여", "?", "궁금", "해줘", "뭐야", "있나", "있어" };
        string[] ofCommand = { "의", "가", "에", "은", "는", "님의", "님" };
        string[] mindCommand = { "어때", "?", "는요", "어떰", "어떨", "어떠", "어떻", "어떤" };

        string[] morningMenu = { " 토스트", "맥모닝", "빵과 커피", "백반", "삼각김밥", "김밥", "샐러드", "선식" };
        string[] lunchMenu = {
            "백반", "떡볶이", "순대", "김밥", "짜장면", "짬뽕", "볶음밥", "김치찌개", "육개장", "된장찌개", "제육볶음",
            "설렁탕", "회덮밥", "냉면", "돈까스", "함박스테이크", "사골국밥", "순대국밥", "갈비탕", "라면", "라멘", "카레",
            "치킨", "피자", "파스타", "육회비빔밥", "비빔밥", "보쌈", "족발", "막국수", "냉모밀", "소바", "스시", "햄버거",
            "한솥", "삼겹살", "소고기", "곱창", "삼계탕", "양념갈비", "스테이크", "생선구이", "훈제오리", "샐러드", "만두"
        };
        string[] dinnerMenu = {
            "백반", "떡볶이", "순대", "김밥", "짜장면", "짬뽕", "볶음밥", "김치찌개", "육개장", "된장찌개", "제육볶음",
            "설렁탕", "회덮밥", "냉면", "돈까스", "함박스테이크", "사골국밥", "순대국밥", "갈비탕", "라면", "라멘", "카레",
            "치킨", "피자", "파스타", "육회비빔밥", "비빔밥", "보쌈", "족발", "막국수", "냉모밀", "소바", "스시", "햄버거",
            "한솥", "삼겹살", "소고기", "곱창", "삼계탕", "양념갈비", "스테이크", "생선구이", "훈제오리", "샐러드", "만두",
        };
        string[] nightMenu = {
            "라면", "치킨", "피자", "보쌈", "족발", "햄버거", "떡볶이", "만두"
        };

        string[] enterMenu = {
            "어떠세요?", "추천합니다.", "땡기네요.", "가보시죠.", "ㄱㄱ", "각이네요", "좋네요.",
            "어떤가요?", "좋을 듯.", "가시죠.", "?", "!", "기대합니다."
        };

        string[] choiceWord = { "이랑", "랑", "," };

        string[] offWork = { "퇴근합니다", "퇴근 합니다", "퇴근합니당", "퇴근 합니당", "퇴근합니닷", "퇴근 합니닷", "퇴근~", "퇴근!", "퇴근하겠" };

        string[] whatWord = { "뭘", "뭐", "어떤", "무엇" };
        string[] eatWord = { "먹을까", "먹지", "먹어야할", "먹나" };
        string[] hungryWord = { "배고픈데", "배고프네", "배고프다", "배고파", "배고픔" };

        ////////////////////////////////////////////////////////////////
        // 형태소 분석
        ////////////////////////////////////////////////////////////////
        public Tuple<string, string, bool> morphemeProcessor(string message, Queue<CMessage> queue)
        {
            Tuple<string, string, bool> emptyTuple = Tuple.Create("", "", false);

            string normalize = TwitterKoreanProcessorCS.Normalize(message);
            var morpheme = TwitterKoreanProcessorCS.Tokenize(normalize);
            List<KoreanToken> result = new List<KoreanToken>();
            foreach (var token in morpheme)
            {
                result.Add(token);
            }

            string command = "";
            string contents = "";

            //--------------------------------------------------------------------------
            // 클랜 기능 감지
            //--------------------------------------------------------------------------
            string[] natural = FunctionCommand(normalize).Split(' ');
            bool isFirst = true;

            command = natural[0].ToString();

            for (int i = 1; i < natural.Count(); i++)
            {
                if (isFirst == true)
                {
                    contents += natural[i].ToString();
                    isFirst = false;
                }
                else
                {
                    contents += " " + natural[i].ToString();
                }
            }

            if (command != "")
                return Tuple.Create(command, contents, true);

            //--------------------------------------------------------------------------
            // 영상 조회
            //--------------------------------------------------------------------------
            if (normalize.Contains("영상") == true || normalize.Contains("방송") == true)
            {
                bool isVideo = false;

                foreach (var word in enterCommand)
                {
                    if (normalize.Contains(word) == true)
                        isVideo = true;
                }

                if (isVideo == true && normalize.Contains("오늘") == true)
                {
                    int day = 0;
                    if (System.DateTime.Now.Hour < 6)
                        day = System.DateTime.Now.Day - 1;
                    else
                        day = System.DateTime.Now.Day;

                    string date = (System.DateTime.Now.Year - 2000).ToString("D2") + System.DateTime.Now.Month.ToString("D2") + day.ToString("D2");

                    Tuple<string, string, bool> tuple = Tuple.Create("/영상", date, true);
                    return tuple;
                }
                else if (isVideo == true && normalize.Contains("어제") == true)
                {
                    int day = 0;
                    if (System.DateTime.Now.Hour < 6)
                        day = System.DateTime.Now.Day - 2;
                    else
                        day = System.DateTime.Now.Day - 1;

                    string date = (System.DateTime.Now.Year - 2000).ToString("D2") + System.DateTime.Now.Month.ToString("D2") + day.ToString("D2");

                    Tuple<string, string, bool> tuple = Tuple.Create("/영상", date, true);
                    return tuple;
                }
                else if (isVideo == true)
                {
                    Tuple<string, string, bool> tuple = Tuple.Create("/영상", "", true);
                    return tuple;
                }
            }

            //--------------------------------------------------------------------------
            // 메뉴 조회
            //--------------------------------------------------------------------------
            if (isExistMenu(message) == true)
            {
                return Tuple.Create(getMenu(message), "", false);
            }

            //--------------------------------------------------------------------------
            // 퇴근 응답
            //--------------------------------------------------------------------------
            if (message.Contains("퇴근") == true)
            {
                string offWork = offWorkCall(message);
                if (offWork != "")
                    return Tuple.Create(offWork, "", false);
            }

            //--------------------------------------------------------------------------
            // 날씨 감지
            //--------------------------------------------------------------------------
            if (message.Contains("날씨") == true)
            {
                Tuple<string, string> weatherTuple = weatherCall(message);
                if (weatherTuple.Item1 != "" && weatherTuple.Item2 != "")
                    return Tuple.Create("/날씨", weatherTuple.Item2, true);
            }

            //--------------------------------------------------------------------------
            // 그 외
            //--------------------------------------------------------------------------
            foreach (var word in result)
            {
                Random random = new Random();
                int num = random.Next(10);

                // 알 수 없는 단어일 경우
                if (word.Unknown == true)
                {
                    switch (num)
                    {
                        case 0:
                            return Tuple.Create("? 그게 무슨 말이에요?", "", false);
                        case 1:
                            return Tuple.Create("? 처음 듣는 말이네요.", "", false);
                        case 2:
                            return Tuple.Create("? 무슨 말인지 모르겠어요.", "", false);
                    }

                    continue;
                }

                // 감탄사
                if (word.Pos.ToString() == "Exclamation")
                {
                    switch (num)
                    {
                        case 0:
                            return Tuple.Create("와, 정말 놀랍네요.", "", false);
                        case 1:
                            return Tuple.Create("저도 놀라워요.", "", false);
                        case 2:
                            return Tuple.Create("대박이네요.", "", false);
                    }
                }

                foreach (var etc in result)
                {
                    if (etc.Pos.ToString() == "Noun" || etc.Pos.ToString() == "ProperNoun")
                    {
                        Random etcRandom = new Random();
                        int etcNumber = etcRandom.Next(10);
                        if (etcNumber == 1)
                        {
                            // 이전 대화내용 참고 기능
                            foreach (var queMsg in queue)
                            {
                                string time = "";

                                if (queMsg.Time.Minute >= System.DateTime.Now.Minute)
                                    continue;

                                if ((queMsg.Time.Year == System.DateTime.Now.Year) &&
                                    (queMsg.Time.Month == System.DateTime.Now.Month) &&
                                    (queMsg.Time.Day == System.DateTime.Now.Day))
                                    time = "아까";
                                else
                                    time = "저번에";
                                

                                if (queMsg.Message.Contains(etc.Text.ToString()) == true)
                                {
                                    Random queRandom = new Random();
                                    int queNumber = queRandom.Next(5);

                                    switch (etcNumber)
                                    {
                                        case 0:
                                            return Tuple.Create(time + " " + etc.Text.ToString() + "에 대해서 말씀하신 적 있어요. 관심있으신가봐요.", "", false);
                                        case 1:
                                            return Tuple.Create(time + " 말씀하신 " + etc.Text.ToString() + " 어떤가요?", "", false);
                                        case 2:
                                            return Tuple.Create(time + " 비슷한 말씀을 하셨었죠.", "", false);
                                        case 3:
                                            return Tuple.Create("자주 언급을 하시니 저도 " + etc.Text.ToString() + "에 대해서 관심을 가져볼까 해요.", "", false);
                                        case 4:
                                            return Tuple.Create("아, " + etc.Text.ToString() + "에 대해서 " + time + " 말씀하셨었어요. 흥미롭네요.", "", false);
                                    }
                                }
                            }
                            
                            // 그냥 언급
                            etcNumber = etcRandom.Next(5);
                            switch (etcNumber)
                            {
                                case 0:
                                    return Tuple.Create(etc.Text.ToString() + " 좋아하시나봐요.", "", false);
                                case 1:
                                    return Tuple.Create(etc.Text.ToString() + ", 저도 궁금하네요.", "", false);
                                case 2:
                                    return Tuple.Create(etc.Text.ToString() + " 어때요?", "", false);
                                case 3:
                                    return Tuple.Create(etc.Text.ToString() + " 좋나요?", "", false);
                                case 4:
                                    return Tuple.Create(etc.Text.ToString() + " 흥미롭네요.", "", false);
                            }
                        }
                    }
                }
            }

            return emptyTuple;
        }

        // 메뉴 추천 감지
        public bool isExistMenu(string message)
        {
            if (message.Contains("먹었"))
                return false;

            foreach (var word in whatWord)
            {
                if (message.Contains(word) == true)
                {
                    foreach (var eat in eatWord)
                    {
                        if (message.Contains(eat) == true)
                        {
                            return true;
                        }
                    }
                }
            }

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
        public string getMenu(string message)
        {
            Random menuRandom = new Random();
            Random enterRandom = new Random();

            if (message.Contains("아침") == true)
            {
                int menuNum = menuRandom.Next(0, morningMenu.Count());
                int enterNum = enterRandom.Next(0, enterMenu.Count());
                return "아침은 " + morningMenu.ElementAt(menuNum) + " " + enterMenu.ElementAt(enterNum);
            }
            else if (message.Contains("점심") == true)
            {
                int menuNum = menuRandom.Next(0, lunchMenu.Count());
                int enterNum = enterRandom.Next(0, enterMenu.Count());
                return "점심은 " + lunchMenu.ElementAt(menuNum) + " " + enterMenu.ElementAt(enterNum);
            }
            else if (message.Contains("저녁") == true)
            {
                int menuNum = menuRandom.Next(0, dinnerMenu.Count());
                int enterNum = enterRandom.Next(0, enterMenu.Count());
                return "저녁은 " + dinnerMenu.ElementAt(menuNum) + " " + enterMenu.ElementAt(enterNum);
            }
            else if (message.Contains("야식") == true)
            {
                int menuNum = menuRandom.Next(0, nightMenu.Count());
                int enterNum = enterRandom.Next(0, enterMenu.Count());
                return "야식은 " + nightMenu.ElementAt(menuNum) + " " + enterMenu.ElementAt(enterNum);
            }
            else if (DateTime.Now.Hour >= 6 && DateTime.Now.Hour <= 9)
            {
                int menuNum = menuRandom.Next(0, morningMenu.Count());
                int enterNum = enterRandom.Next(0, enterMenu.Count());
                return "아침은 " + morningMenu.ElementAt(menuNum) + " " + enterMenu.ElementAt(enterNum);
            }
            else if (DateTime.Now.Hour >= 10 && DateTime.Now.Hour <= 14)
            {
                int menuNum = menuRandom.Next(0, lunchMenu.Count());
                int enterNum = enterRandom.Next(0, enterMenu.Count());
                return "점심은 " + lunchMenu.ElementAt(menuNum) + " " + enterMenu.ElementAt(enterNum);
            }
            else if (DateTime.Now.Hour >= 15 && DateTime.Now.Hour <= 22)
            {
                int menuNum = menuRandom.Next(0, dinnerMenu.Count());
                int enterNum = enterRandom.Next(0, enterMenu.Count());
                return "저녁은 " + dinnerMenu.ElementAt(menuNum) + " " + enterMenu.ElementAt(enterNum);
            }
            else
            {
                int menuNum = menuRandom.Next(0, nightMenu.Count());
                int enterNum = enterRandom.Next(0, enterMenu.Count());
                return "야식은 " + nightMenu.ElementAt(menuNum) + " " + enterMenu.ElementAt(enterNum);
            }
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

        // 퇴근 감지
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



            //--------------------------------------------------------
            // 영상 감지
            //--------------------------------------------------------
            index = 0;
            bool isVideoCommand = false;
            bool isVideoEnter = false;
            int month = 0;
            int day = 0;

            foreach (var word in videoCommand)
            {
                if (text.Contains(word) == true)
                {
                    isVideoCommand = true;
                    break;
                }
            }

            if (isVideoCommand == true)
            {
                foreach (var enter in enterCommand)
                {
                    if (text.Contains(enter) == true)
                    {
                        isVideoEnter = true;
                        break;
                    }
                }

                if (isVideoEnter == false)
                {
                    foreach (var enter in mindCommand)
                    {
                        if (text.Contains(enter) == true)
                        {
                            isVideoEnter = true;
                            break;
                        }
                    }
                }

                if (isVideoCommand == true && isVideoEnter == true)
                {
                    foreach (var word in split)
                    {
                        if (word.Contains("월") == true)
                        {
                            month = Convert.ToInt32(word.Replace("월", ""));
                            continue;
                        }

                        if (word.Contains("일") == true)
                        {
                            day = Convert.ToInt32(word.Replace("일", ""));
                            continue;
                        }
                    }

                    if (month > 0 && day > 0)
                    {
                        return "/영상 18" + month.ToString("D2") + day.ToString("D2");
                    }
                    else
                    {
                        return "";
                    }
                }
            }

            return retCommand;
        }
    }
}
