import json
import os
import sys


attack_type_dict = json.loads(open('./config/attack_type_dict.json', 'r', encoding="utf-8").read())
ENV_KEY_LIST = [
    'PROJECT_NAME', 
    'REPORT_ONWER', 
    'WEBDAV_HOSTNAME', 
    'WEBDAV_LOGIN', 
    'WEBDAV_PASSWORD',
    'DATABASE_URL',
    # 'EXCEPT_APP_IDS',
    # 'EXCEPT_IPS'
    ]
env = {}
for key in ENV_KEY_LIST:
    val = os.environ.get(key)
    if not val or val=="":
        print(f"未设置环境变量{key}")
        sys.exit(-1)
    env[key.lower()] = val


def parse_env_list(env_value, default=None):
    """解析环境变量中的逗号分隔列表"""
    if default is None:
        default = []
    
    if not env_value:
        return default
    
    # 去除首尾空格，按逗号分割，过滤空值
    return [
        f"'{item.strip()}'"
        for item in str(env_value).split(',') 
        if item.strip()
    ]

# 使用
except_app_ids = parse_env_list(os.environ.get('EXCEPT_APP_IDS'))
except_ips = parse_env_list(os.environ.get('EXCEPT_IPS'))

config = {
    "project_name": env.get("project_name"),
    "report_onwer": env.get("report_onwer"),
    "database_url": env.get("database_url"),
    'except_app_ids': except_app_ids,
    'except_ips': except_ips,
    'log_level': os.environ.get('LOG_LEVEL', 'INFO').upper(),
    "webdav_options" : {
        "webdav_hostname": env.get("webdav_hostname"),
        "webdav_login":    env.get("webdav_login"),
        "webdav_password": env.get("webdav_password"),
        "disable_check": True
    },
    "attack_type_dict": attack_type_dict
}
