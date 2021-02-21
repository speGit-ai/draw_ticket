# coding=utf-8

import json
import sys

# 保证兼容python2以及python3

IS_PY3 = sys.version_info.major == 3
if IS_PY3:
    from urllib.request import urlopen
    from urllib.request import Request
    from urllib.error import URLError
    from urllib.parse import urlencode
else:
    from urllib2 import urlopen
    from urllib2 import Request
    from urllib2 import URLError
    from urllib import urlencode

    reload(sys)
    sys.setdefaultencoding('utf8')

# 防止https证书校验不正确
import ssl

ssl._create_default_https_context = ssl._create_unverified_context

# 百度云控制台获取到ak，sk以及
# EasyDL官网获取到URL

# ak
API_KEY = 'NzZD9UylZcXGNUmwbe0ppfmm'

# sk
SECRET_KEY = 'tDg8QPBlkBb4EYt7YvLifhTxgcNvqleI'

# url
EASYDL_TEXT_CLASSIFY_URL = "https://aip.baidubce.com/rpc/2.0/ai_custom/v1/text_cls/spear"

"""  TOKEN start """
TOKEN_URL = 'https://aip.baidubce.com/oauth/2.0/token'

"""
    获取token
"""


def fetch_token():
    params = {'grant_type': 'client_credentials',
              'client_id': API_KEY,
              'client_secret': SECRET_KEY}
    post_data = urlencode(params)
    if (IS_PY3):
        post_data = post_data.encode('utf-8')
    req = Request(TOKEN_URL, post_data)
    try:
        f = urlopen(req, timeout=5)
        result_str = f.read()
    except URLError as err:
        print(err)
    if (IS_PY3):
        result_str = result_str.decode()

    result = json.loads(result_str)

    if ('access_token' in result.keys() and 'scope' in result.keys()):
        if not 'brain_all_scope' in result['scope'].split(' '):
            print('please ensure has check the  ability')
            exit()
        return result['access_token']
    else:
        print('please overwrite the correct API_KEY and SECRET_KEY')
        exit()


"""
    调用远程服务
"""


def request(url, data):
    if IS_PY3:
        req = Request(url, json.dumps(data).encode('utf-8'))
    else:
        req = Request(url, json.dumps(data))

    has_error = False
    try:
        f = urlopen(req)
        result_str = f.read()
        if (IS_PY3):
            result_str = result_str.decode()
        return result_str
    except  URLError as err:
        print(err)


#if __name__ == '__main__':
def class_text(text_draw):
    # 获取access token
    token = fetch_token()

    # 拼接url
    url = EASYDL_TEXT_CLASSIFY_URL + "?access_token=" + token


    # 请求接口
    # 测试好评
    response = request(url,
                       {
                           'text': text_draw,
                           'top_num': 2
                       })

    result_json = json.loads(response)

    result = result_json["results"]

    if result[0]['score'] > result[1]['score']:
        fin = result[0]['name']
    else:
        fin = result[1]['name']
    # 打印好评结果
    # print(text_draw)
    # for obj in result:
    #     print("  评论类别：" + obj['name'] + "    置信度：" + str(obj['score']))
    # print("")
    return fin
