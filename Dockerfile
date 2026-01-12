FROM python:3.9.6-slim

# 设置时区为 Asia/Shanghai
ENV TZ=Asia/Shanghai


# 设置清华源并安装时区数据
RUN sed -i 's/deb.debian.org/mirrors.tuna.tsinghua.edu.cn/g' /etc/apt/sources.list && \
    sed -i 's|security.debian.org/debian-security|mirrors.tuna.tsinghua.edu.cn/debian-security|g' /etc/apt/sources.list

# 安装基础依赖、时区配置和中文字体
RUN apt-get update && apt-get install -y \
    gcc \
    vim \
    libpq-dev \
    tzdata \
    # 安装中文字体
    fonts-wqy-microhei \
    fonts-wqy-zenhei \
    ttf-wqy-microhei \
    ttf-wqy-zenhei \
    xfonts-wqy \
    # 安装常用字体
    fontconfig \
    && ln -fs /usr/share/zoneinfo/Asia/Shanghai /etc/localtime \
    && dpkg-reconfigure --frontend noninteractive tzdata \
    && rm -rf /var/lib/apt/lists/*


# 清理字体缓存并重建
RUN fc-cache -fv

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt \
    -i https://pypi.tuna.tsinghua.edu.cn/simple \
    --trusted-host pypi.tuna.tsinghua.edu.cn


COPY . .

CMD ["python", "main.py"]