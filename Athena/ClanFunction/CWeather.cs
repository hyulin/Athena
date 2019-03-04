using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Athena
{
    class CWeather
    {
        static public Tuple<string, string> getCity(string city)
        {
            string eng = "";
            string kor = "";

            if (city.Contains("서울") == true)
            {
                eng = "seoul";
                kor = "서울";
            }
            else if (city.Contains("경기") == true)
            {
                eng = "gyeonggi-do";
                kor = "경기";
            }
            else if (city.Contains("부산") == true)
            {
                eng = "busan";
                kor = "부산";
            }
            else if (city.Contains("대구") == true)
            {
                eng = "daegu";
                kor = "대구";
            }
            else if (city.Contains("광주") == true)
            {
                eng = "gwangju";
                kor = "광주";
            }
            else if (city.Contains("인천") == true)
            {
                eng = "incheon";
                kor = "인천";
            }
            else if (city.Contains("대전") == true)
            {
                eng = "daejeon";
                kor = "대전";
            }
            else if (city.Contains("울산") == true)
            {
                eng = "ulsan";
                kor = "울산";
            }
            else if (city.Contains("세종") == true)
            {
                eng = "sejong";
                kor = "세종";
            }
            else if (city.Contains("제주") == true)
            {
                eng = "jeju";
                kor = "제주";
            }
            else if (city.Contains("여주") == true)
            {
                eng = "Yeoju";
                kor = "여주";
            }
            else if (city.Contains("양구") == true)
            {
                eng = "Yanggu";
                kor = "양구";
            }
            else if (city.Contains("문경") == true)
            {
                eng = "Mungyeong";
                kor = "문경";
            }
            else if (city.Contains("청송") == true)
            {
                eng = "Cheongsong gun";
                kor = "청송";
            }
            else if (city.Contains("청송") == true)
            {
                eng = "Cheongsong gun";
                kor = "청송";
            }
            else if (city.Contains("충북") == true)
            {
                eng = "Ch’ungch’ŏng-bukto";
                kor = "충북";
            }
            else if (city.Contains("함양") == true)
            {
                eng = "Hamyang";
                kor = "함양";
            }
            else if (city.Contains("연천") == true)
            {
                eng = "yeoncheongun";
                kor = "연천";
            }
            else if (city.Contains("양평") == true)
            {
                eng = "Yangpyong";
                kor = "양평";
            }
            else if (city.Contains("양주") == true)
            {
                eng = "Yangju";
                kor = "양주";
            }
            else if (city.Contains("완주") == true)
            {
                eng = "Wanju";
                kor = "완주";
            }
            else if (city.Contains("부안") == true)
            {
                eng = "Puan";
                kor = "부안";
            }
            else if (city.Contains("아산") == true)
            {
                eng = "Asan";
                kor = "아산";
            }
            else if (city.Contains("문산") == true)
            {
                eng = "Munsan";
                kor = "문산";
            }
            else if (city.Contains("경기광주") == true)
            {
                eng = "Kwangju";
                kor = "경기광주";
            }
            else if (city.Contains("구리") == true)
            {
                eng = "Kuri";
                kor = "구리";
            }
            else if (city.Contains("구미") == true)
            {
                eng = "Kumi";
                kor = "구미";
            }
            else if (city.Contains("김천") == true)
            {
                eng = "Gimcheon";
                kor = "김천";
            }
            else if (city.Contains("김천") == true)
            {
                eng = "Gimcheon";
                kor = "김천";
            }
            else if (city.Contains("가평") == true)
            {
                eng = "Gapyeong";
                kor = "가평";
            }
            else if (city.Contains("화성") == true)
            {
                eng = "Hwaseong";
                kor = "화성";
            }
            else if (city.Contains("화천") == true)
            {
                eng = "Hwacheon";
                kor = "화천";
            }
            else if (city.Contains("홍성") == true)
            {
                eng = "Hongsung";
                kor = "홍성";
            }
            else if (city.Contains("홍천") == true)
            {
                eng = "Hongchon";
                kor = "홍천";
            }
            else if (city.Contains("진안") == true)
            {
                eng = "jin-angun";
                kor = "진안";
            }
            else if (city.Contains("동해") == true)
            {
                eng = "Tonghae";
                kor = "동해";
            }
            else if (city.Contains("성남") == true)
            {
                eng = "Seongnam";
                kor = "성남";
            }
            else if (city.Contains("창녕") == true)
            {
                eng = "Changnyeong";
                kor = "창녕";
            }
            else if (city.Contains("송원") == true)
            {
                eng = "Songwon";
                kor = "송원";
            }
            else if (city.Contains("고성") == true)
            {
                eng = "Kosong";
                kor = "고성";
            }
            else if (city.Contains("기장") == true)
            {
                eng = "Kijang";
                kor = "기장";
            }
            else if (city.Contains("임실") == true)
            {
                eng = "Imsil";
                kor = "임실";
            }
            else if (city.Contains("신안") == true)
            {
                eng = "Sinan";
                kor = "신안";
            }
            else if (city.Contains("경남") == true)
            {
                eng = "Kyŏngsang-namdo";
                kor = "경남";
            }
            else if (city.Contains("경북") == true)
            {
                eng = "Kyŏngsang-bukto";
                kor = "경북";
            }
            else if (city.Contains("충남") == true)
            {
                eng = "Ch’ungch’ŏng-namdo";
                kor = "충남";
            }
            else if (city.Contains("강원") == true)
            {
                eng = "Kangwŏn-do";
                kor = "강원";
            }
            else if (city.Contains("안동") == true)
            {
                eng = "Andong";
                kor = "안동";
            }
            else if (city.Contains("울산") == true)
            {
                eng = "Ulsan";
                kor = "울산";
            }
            else if (city.Contains("천안") == true)
            {
                eng = "Cheonan-si";
                kor = "천안";
            }
            else if (city.Contains("안산") == true)
            {
                eng = "Ansan";
                kor = "안산";
            }
            else if (city.Contains("김해") == true)
            {
                eng = "Kimhae";
                kor = "김해";
            }
            else if (city.Contains("원주") == true)
            {
                eng = "Wonju";
                kor = "원주";
            }
            else if (city.Contains("마산") == true)
            {
                eng = "Masan";
                kor = "마산";
            }
            else if (city.Contains("전북") == true)
            {
                eng = "Chŏlla-bukto";
                kor = "전북";
            }
            else if (city.Contains("익산") == true)
            {
                eng = "Iksan";
                kor = "익산";
            }
            else if (city.Contains("용인") == true)
            {
                eng = "Yongin";
                kor = "용인";
            }
            else if (city.Contains("평택") == true)
            {
                eng = "Pyongtak";
                kor = "평택";
            }
            else if (city.Contains("진주") == true)
            {
                eng = "Chinju";
                kor = "진주";
            }
            else if (city.Contains("부천") == true)
            {
                eng = "Bucheon";
                kor = "부천";
            }
            else if (city.Contains("양산") == true)
            {
                eng = "Yangsan";
                kor = "양산";
            }
            else if (city.Contains("청주") == true)
            {
                eng = "Cheongju";
                kor = "청주";
            }
            else if (city.Contains("춘천") == true)
            {
                eng = "Chuncheon";
                kor = "춘천";
            }
            else if (city.Contains("고양") == true)
            {
                eng = "Goyang";
                kor = "고양";
            }
            else if (city.Contains("전남") == true)
            {
                eng = "Chŏlla-namdo";
                kor = "전남";
            }
            else if (city.Contains("광양") == true)
            {
                eng = "Kwangyang";
                kor = "광양";
            }
            else if (city.Contains("예산") == true)
            {
                eng = "Yesan";
                kor = "예산";
            }
            else if (city.Contains("공주") == true)
            {
                eng = "Kongju";
                kor = "공주";
            }
            else if (city.Contains("진해") == true)
            {
                eng = "Chinhae";
                kor = "진해";
            }
            else if (city.Contains("군산") == true)
            {
                eng = "Kunsan";
                kor = "군산";
            }
            else if (city.Contains("목포") == true)
            {
                eng = "Moppo";
                kor = "목포";
            }
            else if (city.Contains("안양") == true)
            {
                eng = "Anyang";
                kor = "안양";
            }
            else if (city.Contains("전주") == true)
            {
                eng = "Jeonju";
                kor = "전주";
            }
            else if (city.Contains("여수") == true)
            {
                eng = "Reisui";
                kor = "여수";
            }
            else if (city.Contains("속초") == true)
            {
                eng = "Sogcho";
                kor = "속초";
            }
            else if (city.Contains("속초") == true)
            {
                eng = "Sogcho";
                kor = "속초";
            }
            else if (city.Contains("하남") == true)
            {
                eng = "Hanam";
                kor = "하남";
            }
            else if (city.Contains("과천") == true)
            {
                eng = "Kwach’ŏn";
                kor = "과천";
            }
            else if (city.Contains("강릉") == true)
            {
                eng = "Kang-neung";
                kor = "강릉";
            }
            else if (city.Contains("순천") == true)
            {
                eng = "Sunchun";
                kor = "순천";
            }
            else if (city.Contains("밀양") == true)
            {
                eng = "Miryang";
                kor = "밀양";
            }
            else if (city.Contains("순창") == true)
            {
                eng = "Sunch’ang";
                kor = "순창";
            }
            else if (city.Contains("순창") == true)
            {
                eng = "Sunch’ang";
                kor = "순창";
            }
            else if (city.Contains("하동") == true)
            {
                eng = "Hadong";
                kor = "하동";
            }
            else if (city.Contains("상주") == true)
            {
                eng = "Sangju";
                kor = "상주";
            }
            else if (city.Contains("오산") == true)
            {
                eng = "Osan";
                kor = "오산";
            }
            else if (city.Contains("보성") == true)
            {
                eng = "Posung";
                kor = "보성";
            }
            else if (city.Contains("이천") == true)
            {
                eng = "Ichon";
                kor = "이천";
            }
            else if (city.Contains("김제") == true)
            {
                eng = "Kimje";
                kor = "김제";
            }
            else if (city.Contains("논산") == true)
            {
                eng = "Nonsan";
                kor = "논산";
            }
            else if (city.Contains("인제") == true)
            {
                eng = "Inje";
                kor = "인제";
            }
            else if (city.Contains("양양") == true)
            {
                eng = "Yangyang";
                kor = "양양";
            }
            else if (city.Contains("고성") == true)
            {
                eng = "Goseong";
                kor = "고성";
            }
            else if (city.Contains("무주") == true)
            {
                eng = "Muju";
                kor = "무주";
            }
            else if (city.Contains("해남") == true)
            {
                eng = "Haenam";
                kor = "해남";
            }
            else if (city.Contains("무안") == true)
            {
                eng = "Muan";
                kor = "무안";
            }
            else if (city.Contains("나주") == true)
            {
                eng = "Naju";
                kor = "나주";
            }
            else if (city.Contains("경주") == true)
            {
                eng = "Kyonju";
                kor = "경주";
            }
            else if (city.Contains("남해") == true)
            {
                eng = "Namhae";
                kor = "남해";
            }
            else if (city.Contains("광명") == true)
            {
                eng = "Kwangmyong";
                kor = "광명";
            }
            else if (city.Contains("고창") == true)
            {
                eng = "Kochang";
                kor = "고창";
            }
            else if (city.Contains("태백") == true)
            {
                eng = "Taebaek";
                kor = "태백";
            }
            else if (city.Contains("괴산") == true)
            {
                eng = "Koesan";
                kor = "괴산";
            }
            else if (city.Contains("홍천") == true)
            {
                eng = "Hongch’ŏn";
                kor = "홍천";
            }
            else if (city.Contains("남양주") == true)
            {
                eng = "Namyangju-si";
                kor = "남양주";
            }
            else if (city.Contains("강화") == true)
            {
                eng = "Kanghwa";
                kor = "강화";
            }
            else if (city.Contains("옥천") == true)
            {
                eng = "ogcheongun";
                kor = "옥천";
            }
            else if (city.Contains("송정") == true)
            {
                eng = "Songjŏng";
                kor = "송정";
            }
            else if (city.Contains("창원") == true)
            {
                eng = "Changwon";
                kor = "창원";
            }
            else if (city.Contains("안성") == true)
            {
                eng = "Anseong";
                kor = "안성";
            }
            else if (city.Contains("수원") == true)
            {
                eng = "suwon-si";
                kor = "수원";
            }

            Tuple<string, string> tuple = Tuple.Create(eng, kor);

            return tuple;
        }
    }
}
