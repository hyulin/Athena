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

        public string getEasterEgg()
        {
            Random random = new Random();

            int iRandomNum = random.Next(0, strEasterEgg.Count());

            return strEasterEgg.ElementAt(iRandomNum);
        }
    }
}
