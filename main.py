from docx import Document
import psycopg2
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
import matplotlib.font_manager as fm
import datetime
import time
import os
import sys
from config.config import config
from webdav3.client import Client
import schedule
import logging
import logging.handlers
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


today = datetime.date.today()
start_day = today-datetime.timedelta(days=7)

start_time = int(time.mktime(start_day.timetuple()))
end_time = int(time.mktime(today.timetuple())) - 1
end_day = str(datetime.datetime.fromtimestamp(end_time))[:10]



def __render_table_with_data(doc, columns, rows):
    table = doc.add_table(1, len(columns))
    table.style = 'Table Grid'
    # table.style = 'Table Normal'
    hd_cells = table.rows[0].cells
    for i in range(len(columns)):
        hd_cells[i].text = str(columns[i])
    for row in rows:
        cells = table.add_row().cells
        for j in range(len(row)):
            cells[j].text = str(row[j])


def __query_data_from_db(cursor, sql):
    try:
        logger.debug(sql)
        cursor.execute(sql)
        columns = [desc[0] for desc in cursor.description]
        rows = cursor.fetchall()
        cursor.close()
        return columns, rows
    except Exception as e:
        logger.error(e)
        logger.error("查询数据时失败")
        return None
    


def __get_attack_type_name(rows, index):
    row_list = []
    for row in rows:
        row = list(row)
        row[index] = config.get("attack_type_dict").get(f'attack.type.{row[index]}', "未知攻击类型")
        row_list.append(row)
    return row_list


def get_total(doc, conn):
    sql = f"""
        select
            coalesce(sum(case when mss."type" = 'website-req' then mss.value end)::int, 0) as 访问总数,
            coalesce(sum(case when mss."type" = 'website-denied' then mss.value end)::int, 0) as 拦截总数,
            (
            select
                count(*) as 黑名单拦截数
            from
                mgt_rule_detect_log_basic mrdlb
            where
                mrdlb.attack_type =-3
                and mrdlb."timestamp" >= {start_time}
                and mrdlb."timestamp" <= {end_time}
                {f"and mrdlb.site_uuid not in({','.join(config.get('except_app_ids', []))})" if len(config.get('except_app_ids', []))>0 else ''}
                ),
            (
            select
                count(*) as 未拦截数
            from
                mgt_detect_log_basic mdlb
            where
                mdlb."action" = 0
                and mdlb."timestamp" >= {start_time}
                and mdlb."timestamp" <= {end_time}
                {f"and mdlb.site_uuid not in ({','.join(config.get('except_app_ids', []))})" if len(config.get('except_app_ids', []))>0 else ''}
                )
        from
            mgt_system_statistics mss
        where
            mss.created_at >= '{start_day}'
            and
        mss.created_at <= '{end_day}'
        {f"and mss.website not in ({','.join(config.get('except_app_ids', []))})" if len(config.get('except_app_ids', []))>0 else ''}
        """
    
    columns, rows = __query_data_from_db(conn.cursor(), sql)
    return [dict(zip(columns, row)) for row in rows][0]

def get_defens_apps(doc, conn):
    sql = f"""
        select
            mw.id as 应用序号,
            mw."comment" as 应用名称,
            mw.server_names as 域名,
            mw.ports as 开放端口,
            coalesce(SUM(case when mss."type" = 'website-req' then mss.value end)::int,
            0) as 请求次数,
            coalesce(SUM(case when mss."type" = 'website-denied' then mss.value end)::int,
            0) as 拦截次数
        from
            mgt_website mw
        left join mgt_system_statistics mss on
            mw.id = mss.website::bigint
            where
            mss.created_at >= '{start_day}'
            and mss.created_at <= '{end_day}'
            {f"and mw.id not in ({','.join(config.get('except_app_ids', []))})" if len(config.get('except_app_ids', []))>0 else ''}
        group by
            mw.id,
            mw."comment",
            mw.server_names,
            mw.ports
        order by
            mw.id;
            """
    columns, rows = __query_data_from_db(conn.cursor(), sql)
    if len(rows) <= 0:
        doc.add_paragraph(f"暂无防护应用。", style='ReportBodyText')
    else:
        __render_table_with_data(doc, columns, rows)


def get_access_total_by_geos(doc, conn):
    sql = f"""
        select
            country as 国家代号,
            province as 省份,
            city as 城市,
            sum(count) as 访问次数
        from
            statistics_geos sg
        where
            "time" >= {start_time}
            and "time" <= {end_time}
        group by
            country,
            province,
            city
        order by
            访问次数 desc,
            country,
            province,
            city
        """
    columns, rows = __query_data_from_db(conn.cursor(), sql)
    if len(rows) <= 0:
        doc.add_paragraph(f"本周暂无访问数据。", style='ReportBodyText')
    else:
        custom_add_paragraph(doc, f"本周访问数据主要来自:p{rows[0][1]}:sMyEmphasis:p-:p{rows[0][2]}:sMyEmphasis:p，访问次数为:p{rows[0][3]}:sMyEmphasis:p，具体数据可参看下表。")
        __render_table_with_data(doc, columns, rows)


def get_access_total_by_ips(doc, conn):
    sql = f"""
        select 
        si."key" as 访问ip,
        si.attack_type as 访问类型,
        sum(si.count) as 访问次数
        from
        statistics_ips si
        where 
        si."time" >= {start_time}
            and 
        si."time" <= {end_time}
            and 
        si.attack_type = -1
        {f"and si.key not in ({','.join(config.get('except_ips', []))})" if len(config.get('except_ips', [])) > 0 else ''}
        group by si."key",si.attack_type
        order by 访问次数 desc,si.key
        limit 10
        """
    columns, rows = __query_data_from_db(conn.cursor(), sql)
    if len(rows) <= 0:
        doc.add_paragraph(f"本周暂无访问数据。", style='ReportBodyText')
    else:
        rows = __get_attack_type_name(rows, 1)
        custom_add_paragraph(doc, f"本周主要访问IP为:p{rows[0][0]}:sMyEmphasis:p，访问次数为:p{rows[0][2]}:sMyEmphasis:p，具体数据可参看下表。")
        __render_table_with_data(doc, columns, rows)


def get_attack_total_by_ips(doc,conn):
    sql = f"""
        select 
        si."key" as 访问ip,
        si.attack_type as 攻击类型,
        sum(si.count) as 攻击次数
        from
        statistics_ips si
        where 
        si."time" >= {start_time}
            and 
        si."time" <= {end_time}
            and
        si.attack_type > 0
        {f"and si.key not in ({','.join(config.get('except_ips', []))})" if len(config.get('except_ips', [])) > 0 else ''}
        group by si."key",si.attack_type
        order by 攻击次数 desc,si.key
        limit 10
        """
    columns, rows = __query_data_from_db(conn.cursor(), sql)
    if len(rows)<=0:
        doc.add_paragraph("本周暂无攻击数据，您的waf很安全", style='ReportBodyText')
    else:
        rows = __get_attack_type_name(rows, 1)
        custom_add_paragraph(doc, f"本周的攻击主要来自:p{rows[0][0]}:sMyEmphasis:p，攻击类型为:p{rows[0][1]}:sMyEmphasis:p，总计攻击:p{rows[0][2]}:sMyEmphasis:p次，具体数据参看下表。")
        __render_table_with_data(doc, columns, rows)


def get_attack_total_by_type(doc,conn):
    sql = f"""
        select 
        si.attack_type as 攻击类型,
        sum(si.count)::int as 攻击次数
        from
        statistics_ips si
        where 
        si."time" >= {start_time}
            and 
        si."time" <= {end_time}
            and
        si.attack_type > 0 
        {f"and si.key not in ({','.join(config.get('except_ips', []))})" if len(config.get('except_ips', [])) > 0 else ''}
        group by si.attack_type
        order by 攻击次数 desc
        """
    columns, rows = __query_data_from_db(conn.cursor(), sql)
    if len(rows)<=0:
        doc.add_paragraph("本周暂无攻击数据，您的waf很安全", style='ReportBodyText')
    else:
        rows = __get_attack_type_name(rows, 0)

        trans_rows = np.transpose(rows)
        explode = [0.01] * len(trans_rows[0])
        explode[0] = 0.1
        plt.figure(dpi=300)
        plt.pie(trans_rows[1],
            labels=trans_rows[0], # 设置饼图标签
            explode=explode, # 第二部分突出显示，值越大，距离中心越远
            autopct='%.2f%%', # 格式化输出百分比
        )
        plt.title("攻击类型统计图")
        plt.savefig("./攻击类型统计图.png")
        p = custom_add_paragraph(doc, f"本周的主要攻击类型为:p{trans_rows[0][0]}:sMyEmphasis:p，该类型总计攻击:p{trans_rows[1][0]}:sMyEmphasis:p次，具体数据如下图表所示。")
        run = p.add_run()
        run.add_picture("./攻击类型统计图.png")
        __render_table_with_data(doc, columns, rows)



def get_not_defens_log(doc,conn):
    sql = f"""
    select 
    mw."comment" as 被攻击应用,
    mdlb.src_ip as 源IP,
    mdlb.host as 目标主机,
    mdlb.url_path as 请求路径,
    mdlb.dst_port as 目标端口,
    mdlb.country as 国家代码,
    mdlb.province as 省份,
    mdlb.city as 城市,
    mdlb.attack_type as 攻击类型,
    mdlb.updated_at as 攻击时间
    from 
    mgt_detect_log_basic mdlb,
    mgt_website mw
    where
    mdlb.site_uuid::int = mw.id::int
    and
    mdlb."timestamp" >= {start_time}
    and
    mdlb."timestamp" <= {end_time}
    and
    mdlb."action" = 0
    {f"and mdlb.site_uuid not in ({','.join(config.get('except_app_ids', []))})" if len(config.get('except_app_ids', []))>0 else ''}
    """
    columns, rows = __query_data_from_db(conn.cursor(), sql)
    if len(rows) <= 0:
        doc.add_paragraph("本周暂无未拦截攻击，所有攻击都被拒之门外。", style='ReportBodyText')
    else:
        # doc.add_paragraph(f"本周攻击有{len(rows)}条攻击未被拦截，我们将对其进行分析和拦截处理，具体数据参看下表。")
        custom_add_paragraph(doc, f"本周攻击有:p{len(rows)}:sMyEmphasis:p条攻击未被拦截，我们将对其进行分析和拦截处理，具体数据参看下表。")
        rows = __get_attack_type_name(rows, 8)
        __render_table_with_data(doc, columns, rows)

def init_doc():
    doc = Document()
    styles = doc.styles

    H1 = styles['Heading 1']
    H1.font.name = '微软雅黑'
    H1.font.color.rgb = RGBColor(79, 128, 189)

    H2 = styles['Heading 2']
    H2.font.name = '微软雅黑'
    H1.font.color.rgb = RGBColor(79, 128, 189)

    if 'ReportHeading1' not in styles:
        heading_style = styles.add_style('ReportHeading1', 1)
        heading_style.font.name = '黑体'
        heading_style.font.size = Pt(16)
        heading_style.font.bold = True
        heading_style.font.color.rgb = RGBColor(0, 0, 139)
        heading_style.paragraph_format.space_before = Pt(12)
        heading_style.paragraph_format.space_after = Pt(6)

    if 'ReportHeading2' not in styles:
        heading_style = styles.add_style('ReportHeading2', 1)
        heading_style.font.name = '黑体'
        heading_style.font.size = Pt(13)
        heading_style.font.bold = True
        heading_style.font.color.rgb = RGBColor(0, 0, 100)
        heading_style.paragraph_format.space_before = Pt(12)
        heading_style.paragraph_format.space_after = Pt(6)

    if 'ReportBodyText' not in styles:
        body_style = styles.add_style('ReportBodyText', 1)
        body_style.font.name = '宋体'
        body_style.font.size = Pt(10.5)
        body_style.paragraph_format.line_spacing = 1.2
        body_style.paragraph_format.first_line_indent = Pt(21)  # 首行缩进2字符
    
    # 预设强调样式（字符样式）
    if 'MyEmphasis' not in styles:
        emph_style = styles.add_style('MyEmphasis', 2)  # 2=字符样式
        emph_style.font.name = '楷体'
        emph_style.font.italic = True
        emph_style.font.color.rgb = RGBColor(178, 34, 34)
    return doc


def __render_paragraph(paragraph, texts):
    for text in texts:
        run = paragraph.add_run(str(text['value']))
        try:
            run.style = text['style']
        except Exception as e:
            pass
            # logger.error("无style值")

def __render_paragraph_by_template(paragraph, tpl):
    tpl = str(tpl)
    seq = tpl.split(":p")
    for p in seq:
        s = p.split(":s")
        run = paragraph.add_run(str(s[0]))
        # logger.debug(s)
        try:
            run.style=s[1]
        except Exception as e:
            # logger.error(e)
            # logger.debug("无style")
            pass

def custom_add_paragraph(doc, tpl):
    par = doc.add_paragraph("",style='ReportBodyText')
    __render_paragraph_by_template(par, tpl)
    return par


def main():
    logger.info(f"任务触发")
    if not os.path.exists('./report'):
        os.mkdir('./report')
    logger.debug(f"数据库连接信息：{config['database_url']}")
    conn = psycopg2.connect(config['database_url'])
    try:
        # doc = Document()
        doc = init_doc()
        # doc.add_heading(f"{config['project_name']}网络安全巡检报告")
        doc.add_heading("一、巡检必要性概述", level=1)
        doc.add_paragraph("本周对Web应用防火墙（WAF）进行了系统性巡检，旨在及时发现并处置潜在的安全威胁，确保Web应用服务的持续可用性和数据安全性。通过定期巡检，可有效识别异常攻击流量、优化防护策略、验证防护效果，并为安全态势评估提供数据支撑，降低因Web应用漏洞导致的数据泄露和服务中断风险。",style='ReportBodyText')
        doc.add_heading("二、防护应用概览", level=1)
        data = get_total(doc, conn)
        if data['未拦截数'] == 0:
            p = 100
        else:
            p = format(data['拦截总数']/(data['拦截总数']+data['未拦截数'])*100, '.2f')
        # doc.add_paragraph(f"本周waf总体运行平稳，总访问次数为{data['访问总数']}，总拦截次数为{data['拦截总数']}，黑名单拦截次数为{data['黑名单拦截数']}，未拦截攻击次数为{data['未拦截数']}，拦截率为{p}%。")
        template_p = f"本周waf总体运行平稳，总访问次数为:p{data['访问总数']}:sMyEmphasis:p，总拦截次数为:p{data['拦截总数']}:sMyEmphasis:p，黑名单拦截次数为:p{data['黑名单拦截数']}:sMyEmphasis:p，未拦截攻击次数为:p{data['未拦截数']}:sMyEmphasis:p，拦截率为:p{p}:sMyEmphasis:p %。"
        custom_add_paragraph(doc, template_p)
        doc.add_heading("2.1 分应用查看", level=2)
        get_defens_apps(doc, conn)
        doc.add_heading("三、访问数据统计", level=1)
        doc.add_heading("3.1 按地理区域统计访问数据", level=2)
        get_access_total_by_geos(doc, conn)
        doc.add_heading("3.2 按访问IP统计访问数据TOP10", level=2)
        get_access_total_by_ips(doc, conn)
        doc.add_heading("四、攻击数据统计", level=1)
        doc.add_heading("4.1 按攻击方式统计数据", level=2)
        get_attack_total_by_type(doc, conn)
        doc.add_heading("4.2 按攻击IP统计访问数据TOP10", level=2)
        get_attack_total_by_ips(doc, conn)
        doc.add_heading("五、明细数据展示", level=1)
        doc.add_heading("5.1 未拦截攻击明细", level=2)
        get_not_defens_log(doc, conn)

        doc.add_heading("六、报告信息", level=1)
        doc.add_paragraph(f"项目名称：{config['project_name']}")
        doc.add_paragraph(f"报告数据统计周期：{start_day}至{end_day}")
        doc.add_paragraph(f"报告生成时间：{datetime.datetime.now()}")
        doc.add_paragraph(f"报告审核人：{config['report_onwer']}")
        

        conn.close()
        doc_filename = f"{config['project_name']}_{start_day}至{end_day}安全运维周报.docx"
        local_file_path = f'./report/{doc_filename}'
        doc.save(local_file_path)
    except Exception as e:
        logger.error(e)
        logger.error("生成报告失败")


    try:
        if os.path.exists(local_file_path):
            client = Client(config.get("webdav_options"))
            remote_base_path = f'report/{config.get("report_onwer")}'
            client.mkdir(remote_base_path)
            remote_base_path = f'report/{config.get("report_onwer")}/{str(today).replace("-", "")}'
            client.mkdir(remote_base_path)
            remote_file_path = f'{remote_base_path}/{doc_filename}'
            client.upload_sync(remote_path=remote_file_path, local_path=local_file_path)
        else:
            logger.error("本地文件不存在，上传失败")
        logger.info(f"任务结束")
    except Exception as e:
        logger.error(e)
        logger.error("上传文件失败")




def get_logger(name):
    if not os.path.exists('./logs'):
        os.mkdir('./logs')
    logger = logging.getLogger(name)
    logger.setLevel(config.get('log_level'))
    
    # 如果已经配置过，直接返回
    if logger.handlers:
        return logger
    
    # 设置格式
    formatter = logging.Formatter(
        '%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # 输出到文件（自动切割）
    file_handler = logging.handlers.RotatingFileHandler(
        './logs/app.log',
        maxBytes=10*10*1024*1024,  # 100MB
        backupCount=5,
        encoding='utf-8'
    )
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    # 输出到控制台
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    return logger


if __name__ == '__main__':
    logger = get_logger(__name__)
    WQY_FONT_PATH = '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc'
    
    if os.path.exists(WQY_FONT_PATH):
        matplotlib.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei']
        matplotlib.rcParams['axes.unicode_minus'] = False
        logger.info("字体加载OK")
    else:
        matplotlib.rcParams['font.sans-serif'] = ['DejaVu Sans']
        matplotlib.rcParams['axes.unicode_minus'] = False
        logger.debug("使用回退字体: DejaVu Sans")

    if len(sys.argv) == 2 and sys.argv[1] == '-now':
        logger.debug("立即执行生成报告")
        main()
        sys.exit(0)
    
    schedule.every().day.at("12:00").do(main)

    while True:
        logger.info(f"检测定时任务")
        
        today = datetime.date.today()
        start_day = today-datetime.timedelta(days=7)

        logger.debug(f"{start_day}-{end_day}")

        start_time = int(time.mktime(start_day.timetuple()))
        end_time = int(time.mktime(today.timetuple())) - 1
        end_day = str(datetime.datetime.fromtimestamp(end_time))[:10]
        schedule.run_pending()
        time.sleep(3)