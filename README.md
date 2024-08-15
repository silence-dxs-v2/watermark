# watermark
基于poi给docx、excel、pdf添加水印、word转pdf

## 功能
### word/excel/pdf 加水印

<img width="1193" alt="image" src="https://github.com/user-attachments/assets/eba24684-313f-441d-a8f0-bca0fd1b3759">
### word转pdf


## docker 部署启动
docker pull dockersilencev1/watermark:v1



docker run -d  -p 8100:8082  --name  watermark-web   dockersilencev1/watermark-web:v1

## swagger 
http://localhost:8100//swagger-ui/index.html
