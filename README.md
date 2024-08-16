# watermark
基于poi给docx、excel、pdf添加水印、word转pdf，无需安装第三方软件

## 技术点
- springboot
- poi
- aspose
- swagger

## 功能
### word/excel/pdf 加水印
![image](https://github.com/user-attachments/assets/c57c883d-bb5e-44b8-b0c5-9888936f56cd)



### word转pdf

![image](https://github.com/user-attachments/assets/5d4da050-6d08-462a-acd8-115328e290ed)




## docker 部署启动
docker pull dockersilencev1/watermark:v1



docker run -d  -p 8100:8082  --name  watermark-web   dockersilencev1/watermark:v1

## swagger 
http://localhost:8100//swagger-ui/index.html
