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

            Tuple<string, string> tuple = Tuple.Create(eng, kor);

            return tuple;
        }
    }
}
