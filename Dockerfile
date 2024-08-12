FROM openjdk:8-jdk

RUN apt-get update && apt-get install -y tini && apt-get  install -y libreoffice

# 复制 jar 文件到容器
COPY target/watermark.jar /app/app.jar
ENV LIBREOFFICE_HOME=/usr/lib/libreoffice
ENV PATH=$JAVA_HOME/bin:$LIBREOFFICE_HOME/program:$PATH

# 设置工作目录
WORKDIR /app

# 设置 ENTRYPOINT 以允许使用 exec
ENTRYPOINT ["tini", "--", "java", "-Djava.security.egd=file:/dev/./urandom", "-jar", "app.jar"]

EXPOSE 8082/tcp