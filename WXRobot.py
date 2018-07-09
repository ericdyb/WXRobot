# 微信聊天机器人

import RequestQuery

import wxpy
import json
import requests


API_KEY = 'a4062760f1c240ef97b788e466064d9b' #申请的图灵机器人API KEY
USER_ID = '21010528' #自行分配的图灵机器人USER ID

def init_bot():
    # newbot = wxpy.Bot()
    newbot = wxpy.Bot(console_qr=False, cache_path=True)
    newbot.file_helper.send('你好，微信机器人已上线！')
    return newbot

# 调用图灵机器人，发送消息并获得机器人的回复
def tuling_txt(text):
    url = 'http://www.tuling123.com/openapi/api'
    payload = {
        'key': API_KEY,
        'info': text,
        'userid': USER_ID
    }
    r = requests.post(url, data=json.dumps(payload))
    result = json.loads(r.content)
    ret_str = '【机器人】' + result['text']
    print('（回复） '+ret_str)
    return ret_str

# 自动回复，重复收到的文字
def auto_repeat(bot):
    @bot.register()
    def print_message(msg):
        print(msg.text)
        reply_str = '【机器人重复回复】'+msg.text
        return reply_str

#自动回复，调用图灵机器人
def auto_reply(bot):

    request_data, sheet_names = RequestQuery.read_excel()

    print('【进入图灵机器人自动应答模式】')

    @bot.register()
    def forward_message(msg):
        print('（收到） '+msg.chat.name+': '+msg.text)
        member = msg.member
        if member != None:
            group_name = member.group.name
        else:
            group_name = '无'
        if (member != None) and (group_name == '再不疯狂我们就老了'):
           if ('需求单号' in msg.text) and (msg.text.find('需求单号') == 0):
                req_id = str(msg.text).split(' ')[-1]
                query_data = RequestQuery.request_query(request_data,sheet_names,'i',req_id)
                if query_data['isfound'] == 1:
                    txt = '【需求查询结果】\n需求单号：{}\n'.format(query_data['id']) + \
                        '需求名称：{}\n'.format(query_data['name']) + \
                        '提出部门：{}\n'.format(query_data['dept']) + \
                        '需求负责人：{}\n'.format(query_data['owner']) + \
                        '所处需求队列类别：{}\n'.format(query_data['queue']) + \
                        '当前所处需求队列优先级：{}\n'.format(query_data['priority']) + \
                        '当前处理人：{}\n'.format(query_data['handler']) + \
                        '当前处理进度：%.0f%%' % query_data['status']
                else:
                    txt = '该需求\'{}\'查无记录,请按\'需求名称 具体需求名称\'或\'需求单号 具体需求单号\'的格式查询，支持模糊匹配。'.format(req_id)
           elif ('需求名称' in msg.text) and (msg.text.find('需求名称') == 0):
                req_name = str(msg.text).split(' ')[-1]
                query_data = RequestQuery.request_query(request_data,sheet_names,'n', req_name)
                if query_data['isfound'] == 1:
                    txt = '【需求查询结果】\n需求单号：{}\n'.format(query_data['id']) + \
                        '需求名称：{}\n'.format(query_data['name']) + \
                        '提出部门：{}\n'.format(query_data['dept']) + \
                        '需求负责人：{}\n'.format(query_data['owner']) + \
                        '所处需求队列类别：{}\n'.format(query_data['queue']) + \
                        '当前所处需求队列优先级：{}\n'.format(query_data['priority']) + \
                        '当前处理人：{}\n'.format(query_data['handler']) + \
                        '当前处理进度：%.0f%%' % query_data['status']
                else:
                    txt = '该需求\'{}\'查无记录，请按\'需求名称 具体需求名称\'或\'需求单号 具体需求单号\'的格式查询，支持模糊匹配。'.format(req_name)
           elif ('机器人' in msg.text) and (msg.text.find('机器人') == 0):
                txt = tuling_txt(str(msg.text).split(' ')[1])
           else:
                print('(不回复来自群 {} 的消息) \n'.format(msg.member.group.name))
                txt = ''
        elif member == None:
            print('(回复来自好友 {} 的消息）\n'.format(msg.sender.name))
            txt = ''
            if msg.sender.name == '逸':
                if ('需求单号' in msg.text) and (msg.text.find('需求单号') == 0):
                    req_id = str(msg.text).split(' ')[-1]
                    query_data = RequestQuery.request_query(request_data, sheet_names, 'i', req_id)
                    if query_data['isfound'] == 1:
                        txt = '【需求查询结果】\n需求单号：{}\n'.format(query_data['id']) + \
                              '需求名称：{}\n'.format(query_data['name']) + \
                              '提出部门：{}\n'.format(query_data['dept']) + \
                              '需求负责人：{}\n'.format(query_data['owner']) + \
                              '所处需求队列类别：{}\n'.format(query_data['queue']) + \
                              '当前所处需求队列优先级：{}\n'.format(query_data['priority']) + \
                              '当前处理人：{}\n'.format(query_data['handler']) + \
                              '当前处理进度：%.0f%%' % query_data['status']
                    else:
                        txt = '该需求\'{}\'查无记录,请按\'需求名称 具体需求名称\'或\'需求单号 具体需求单号\'的格式查询，支持模糊匹配。'.format(req_id)
                elif ('需求名称' in msg.text) and (msg.text.find('需求名称') == 0):
                    req_name = str(msg.text).split(' ')[-1]
                    query_data = RequestQuery.request_query(request_data, sheet_names, 'n', req_name)
                    if query_data['isfound'] == 1:
                        txt = '【需求查询结果】\n需求单号：{}\n'.format(query_data['id']) + \
                              '需求名称：{}\n'.format(query_data['name']) + \
                              '提出部门：{}\n'.format(query_data['dept']) + \
                              '需求负责人：{}\n'.format(query_data['owner']) + \
                              '所处需求队列类别：{}\n'.format(query_data['queue']) + \
                              '当前所处需求队列优先级：{}\n'.format(query_data['priority']) + \
                              '当前处理人：{}\n'.format(query_data['handler']) + \
                              '当前处理进度：%.0f%%' % query_data['status']
                    else:
                        txt = '该需求\'{}\'查无记录，请按\'需求名称 具体需求名称\'或\'需求单号 具体需求单号\'的格式查询，支持模糊匹配。'.format(req_name)
            if txt == '':
                txt = tuling_txt(msg.text)
        else:
            print('(不回复来自群 {} 的消息) \n'.format(msg.member.group.name))
            txt = ''

        return txt


# 好友统计
def friends_stat(bot):
    fristat = bot.friends().stats()
    friend_loc = []
    for province, count in fristat['province'].items():
        if province != "":
            friend_loc.append([province, count])

    # 对人数倒序排序
    friend_loc.sort(key=lambda x: x[1], reverse=True)

    # 打印人数最多的10个地区
    print('【好友地区统计】')
    for item in friend_loc:
        print('地区：{}，人数：{}'.format(item[0],item[1]))

    # 打印性别人数
    print('【好友已知性别统计】')
    for sex, count in fristat['sex'].items():
        # 1代表MALE, 2代表FEMALE
        if sex == 1:
            print('男性：', count)
        elif sex == 2:
            print('女性：', count)
        else:
            print('未知: ', count)

def WXRobot():
    bot = init_bot()

    friends_stat(bot)

    #auto_repeat(bot)

    auto_reply(bot)

    # 进入Python命令行阻塞线程，让程序保持运行
    wxpy.embed()


if __name__ == '__main__':
    WXRobot()