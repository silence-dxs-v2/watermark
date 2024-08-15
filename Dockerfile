FROM openjdk:8-jdk-slim

RUN apt-get update && apt-get install -y libfreetype6 fontconfig
# 设置工作目录
WORKDIR /app

# 复制构建好的 JAR 文件到镜像中
COPY target/watermark.jar /app/app.jar
# 复制字体到镜像中
COPY target/classes/font/simsun.ttf  /usr/share/fonts/simsun.ttf
# 更新字体缓存
RUN fc-cache -f -v
# 暴露应用的端口
EXPOSE 8100

# 定义启动命令
CMD ["java", "-jar", "app.jar"]