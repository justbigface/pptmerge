# PPT Merge Service

合并多个 PPTX 文件为一个。支持 HTTP 表单多文件上传，返回合并后的 PPT 文件。

## 运行方式

### 1. 本地直接运行

```bash
pip install -r requirements.txt
python app/ppt_merge_service.py

接口默认监听 8080 端口。
### 2. Docker部署

docker build -t ppt-merge-service .
docker run -d -p 8080:8080 ppt-merge-service

### 3. API用法

接口：POST /merge_pptx
请求格式：JSON
参数：{"urls": ["<PPT文件URL1>", "<PPT文件URL2>", ...]}（需至少两个URL）
返回：合并后的 merged.pptx

curl调用示例：
curl -X POST -H "Content-Type: application/json" -d '{"urls": ["https://example.com/1.pptx", "https://example.com/2.pptx"]}' http://localhost:8080/merge_pptx --output merged.pptx

## License
MIT

---

## 使用/部署

1. 按照上文目录结构创建代码库。
2. 推送到GitHub：

```bash
git init
git add .
git commit -m "init"
git remote add origin <你的github仓库>
git push origin master
```
3. 构建并运行docker，或本地直接用Python运行。

如需更丰富的合并能力或异常处理、权限控制，请进一步扩展代码！有任何细节需求欢迎补充。

## 注意事项
- 服务需能访问外网，Google Drive文件链接需设置为“所有有链接的人可查看”。
- 长期高并发运行建议使用gunicorn等WSGI服务器部署。
- python-pptx对大PPT有内存消耗，临时文件会自动清理。