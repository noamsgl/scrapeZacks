# pandas stuff
import pandas as pd
from io import StringIO


import requests

cookies = {
    '_ga': 'GA1.2.1294520602.1610308362',
    '_fbp': 'fb.1.1610308363007.1608728831',
    '_vwo_uuid_v2': 'D8776A850D8025D8390504E31F844A978|aa521abbe2ce29224bae7f9e1b42a99e',
    'com.silverpop.iMAWebCookie': '0ca6b40b-2349-66bd-753c-2dd666abc040',
    '_gid': 'GA1.2.1118877018.1610742967',
    'cto_bundle': 'LHY2Pl9XT053ZU5LMHViQSUyRmZoYnBjVkVZdDRKTTZQN1Z5MyUyRnB4ZEdoYWoxR29tRE82ZzVEZzQxJTJCbjlWN2t5M3FTJTJCYzJuTm4wVW9XQVIlMkZLaXJUSWhTbW1tUE9GOSUyQlJPVVN1VWR6V245VGdKM056WVRiSDA2WiUyQmlVR1k2S3hMMHglMkZzUjd0RE9GOXAyQWZOJTJGZmo1b0NoOWNBbHclM0QlM0Q',
    'cf60519feb1344e434b8444b746a915b': 'cdaeeeba9b4a4c5ebf042c0215a7bb0e',
    'AMCVS_3064401053DB594D0A490D4C%40AdobeOrg': '1',
    'AMCV_3064401053DB594D0A490D4C%40AdobeOrg': '77933605%7CMCIDTS%7C18646%7CMCMID%7C01756172377111412384472593546162443039%7CMCAAMLH-1611573753%7C6%7CMCAAMB-1611573753%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1610976153s%7CNONE%7CMCAID%7CNONE%7CvVersion%7C4.5.1',
    's_gapv': 'Logged%20Out',
    's_vnum': '1612130400126%26vn%3D20',
    's_invisit': 'true',
    'undefined_s': 'First%20Visit',
    's_v17': 'DEF',
    's_cc': 'true',
    'PHPSESSID': 't9n74dom004g0oe2s7i78j8qp0',
    's_p42': 'screening%3A%20stock-screener',
    's_sq': '%5B%5BB%5D%5D',
    '_gat': '1',
    # 'FCCDCF': '[["AKsRol9_jhI0FB-ohN8PGxChw8Tu5-_ER4YW9j0z-fDBEsZvq_2XXC7IveQTAo46GzQ2Lyk3nnSrvcHvttZWKiQMu6MhRb7z4aQiW_bJxo67anR7jOfiH1x46b_w0-6hSVaNpPYwcoU9d2Afdhdhq9vaPDG76w6M5Q=="],null,["[[],[],[],[],null,null,true]",1610970328462]]',
    # '__gads': 'ID=2b4ab2b6302b6d32:T=1610970326:S=ALNI_MYtr3-UxjjDLzC2XbdggXZu1x6YuA',
    # 's_nr': '1610970329993-Repeat',
    # 'CURRENT_POS': 'predefined',
    # 'FROM_PRE_ID': '32663',

    'FCCDCF': '^[^[^\\^AKsRol9WnhEEdHT4JU6DxO4MnYAqxdKkyRu03P0lpbsaqe5pSLK71vW64sObi1RiUfGG-APUJyZGuHC7swOFIVl-A32FFPQObREUb9G4sxWCHZ94EOapIGEUcOswZg2WLYD9X2tnBQmdu9SmHutEH9iygm3uFLM57g==^\\^^],null,^[^\\^^[^[^],^[^],^[^],^[^],null,null,true^]^\\^,1610906856073^]^]',
    '__gads': 'ID=2b4ab2b6302b6d32-223b915ca3a600a5:T=1610971214:RT=1610971214:S=ALNI_MblI8NTJ3hHj0NVuAnMrSyagoamdQ',
    'CURRENT_POS': 'my_screen',
    's_nr': '1610968957767-Repeat',

}

headers = {
    'Connection': 'keep-alive',
    'Upgrade-Insecure-Requests': '1',
    'DNT': '1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.141 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'Sec-Fetch-Site': 'same-origin',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-User': '?1',
    'Sec-Fetch-Dest': 'iframe',
    'Referer': 'https://screener-api.zacks.com/?scr_type=stock&c_id=zacks&c_key=0675466c5b74cfac34f6be7dc37d4fe6a008e212e2ef73bdcd7e9f1f9a9bd377&ref=screening',
    'Accept-Language': 'en-US,en;q=0.9,he;q=0.8',
}

response = requests.get('https://screener-api.zacks.com/export.php', headers=headers, cookies=cookies)

df = pd.read_csv(StringIO(response.text))

print(df)