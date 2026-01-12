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
    'DATABASE_URL'
    ]
env = {}
for key in ENV_KEY_LIST:
    val = os.environ.get(key)
    if not val or val=="":
        print(f"未设置环境变量{key}")
        sys.exit(-1)
    env[key.lower()] = val

config = {
    "project_name": env.get("project_name"),
    "report_onwer": env.get("report_onwer"),
    "database_url": env.get("database_url"),
    "webdav_options" : {
        "webdav_hostname": env.get("webdav_hostname"),
        "webdav_login":    env.get("webdav_login"),
        "webdav_password": env.get("webdav_password"),
        "disable_check": True
    },
    "attack_type_dict": attack_type_dict
}
