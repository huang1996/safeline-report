# safeline-report
## safeline-report是一个python脚本容器，和safelinewaf compose部署在一起对接数据库，通过数据库生成waf巡检报告
### 使用方法
修改safelinewaf部署目录的`docker-compose.yaml`文件
``` yaml
  #....safelinewaf配置下新增report服务
  report:
    container_name: safeline-report
    restart: always
    image: huangtao1996/safeline-report
    environment:
      #- DATABASE_URL=postgres://safeline-ce:${POSTGRES_PASSWORD}@safeline-pg/safeline-ce?sslmode=disable
      # 如果密码中有特殊字符请使用下面的连接字符串
      - DATABASE_URL=postgresql://safeline-pg/safeline-ce?user=safeline-ce&password=${POSTGRES_PASSWORD}&sslmode=disable
      - PROJECT_NAME=项目名称
      - REPORT_ONWER=报告审核人
      - WEBDAV_HOSTNAME=webdav主机地址
      - WEBDAV_LOGIN=webdav密码
      - WEBDAV_PASSWORD=webdav密码
    networks:
      safeline-ce:
        ipv4_address: ${SUBNET_PREFIX}.20
```
部署服务
``` bash
docker compose pull report && docker compose up -d report
# 如果想要马上生成报告，可以使用如下命令
docker compose exec -it report python main.py -now
```