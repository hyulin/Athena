using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CDT_Noti_Bot
{
    class CEasterEgg
    {
        string[] strEasterEgg = {
            "방벽 생성기 시험 가동. 준비 완료.",
            "진정하세요.윈스턴. 심박수가 천장을 뚫을 기세입니다.",
            "마지막 유산소 운동 후로, 43일 7시간 29초가 지났습니다.",
            "기억하세요.건강한 몸에 건강한 마음.",
            "이런 뉴스를 볼 때마다 꼭 그러시는군요. 재차 말씀드리지만, 오버워치 요원들을 현역으로 복귀시키는데는 큰 위험이 따릅니다.",
            "2042년 제정된 페트라스 법에 따르면 오버워치 활동은 불법이며, 형사처벌까지 받을 수 있습니다.",
            "보안 프로토콜 오류. 윈스턴! 리퍼가 오버워치 요원 데이터베이스를 추출합니다!",
            "윈스턴... 윈스턴... 윈스턴! 리퍼가 요원들의 위치를 거의 빼냈어요!",
            "바이러스 격리 완료. 핵심 데이터베이스 진단 중. 시스템 복구 시작.",
            "요원 연결 중."
        };

        string[] menu = {
            "백반", "떡볶이", "순대", "김밥", "짜장면", "짬뽕", "볶음밥", "김치찌개", "육개장", "된장찌개", "제육볶음",
            "설렁탕", "회덮밥", "냉면", "돈까스", "함박스테이크", "국밥", "갈비탕", "라면", "라멘", "카레", "치킨",
            "피자", "파스타", "육회비빔밥", "비빔밥", "보쌈", "족발", "막국수", "냉모밀", "소바", "스시", "햄버거", "한솥",
            "삼겹살", "소고기", "곱창", "삼계탕", "양념갈비", "스테이크", "생선구이", "훈제오리", "샐러드", "만두"
        };

        string[] enter = {
            "어떠세요?", "추천합니다.", "땡기네요.", "가보시죠.", "ㄱㄱ", "각", "좋네요.",
            "어떤가요?", "좋을 듯.", "가시죠.", "?", "!", "너로 정했다!", "기대합니다."
        };


        public string getEasterEgg()
        {
            Random random = new Random();

            int iRandomNum = random.Next(0, strEasterEgg.Count());

            return strEasterEgg.ElementAt(iRandomNum);
        }

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

            if (message.Contains("배고픈데") || message.Contains("배고프다") || message.Contains("배고프네요") || message.Contains("배고픔"))
                return true;

            return false;   
        }

        public string getMenu()
        {
            Random menuRandom = new Random();
            Random enterRandom = new Random();

            int menuNum = menuRandom.Next(0, menu.Count());
            int enterNum = enterRandom.Next(0, enter.Count());

            return menu.ElementAt(menuNum) + " " + enter.ElementAt(enterNum);
        }
    }
}
