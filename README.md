# watermark
基于poi给docx、excel、pdf添加水印

## 效果

<img width="1193" alt="image" src="https://github.com/user-attachments/assets/eba24684-313f-441d-a8f0-bca0fd1b3759">

## word转pdf
 参考链接
 https://blog.csdn.net/qq_42882229/article/details/140917550


docker run -d -p 8100:8100 --privileged=true -v /apps/my/font:/usr/share/fonts -v /apps/my/config/application.properties:/etc/app/application.properties ghcr.io/jodconverter/jodconverter-examples:rest