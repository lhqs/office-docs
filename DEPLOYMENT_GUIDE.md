# 部署与 SEO 配置指南

> 办公软件精通指南 - 生产环境部署完整指南
> 作者：lhqs (lhqs1314@gmail.com)
> 部署地址：https://docs.office.ninthfeast.com

---

## 📦 部署前准备

### 1. 环境检查

```bash
# 检查 Node.js 版本（需要 >= 20.0）
node --version

# 检查 npm 版本
npm --version

# 安装依赖
npm install
```

### 2. 构建生产版本

```bash
# 清理旧的构建文件
npm run clear

# 构建生产版本
npm run build

# 本地预览构建结果
npm run serve
# 访问 http://localhost:3000
```

### 3. SEO 检查

```bash
# 运行 SEO 检查脚本
npm run seo:check

# 或者一键构建+检查
npm run build:prod
```

---

## 🚀 部署步骤

### 方式一：使用 Vercel（推荐）

1. **连接 Git 仓库**
```bash
# 安装 Vercel CLI
npm i -g vercel

# 登录
vercel login

# 初次部署
vercel

# 生产环境部署
vercel --prod
```

2. **环境配置**
```
Framework Preset: Docusaurus 2
Build Command: npm run build
Output Directory: build
Install Command: npm install
```

3. **域名配置**
```
Domains > Add Domain
输入: docs.office.ninthfeast.com
```

### 方式二：使用 Netlify

1. **部署配置**
创建 `netlify.toml`:
```toml
[build]
  command = "npm run build"
  publish = "build"

[[redirects]]
  from = "/*"
  to = "/index.html"
  status = 200
```

2. **部署**
```bash
# 安装 Netlify CLI
npm i -g netlify-cli

# 部署
netlify deploy --prod
```

### 方式三：传统服务器部署

1. **构建**
```bash
npm run build
```

2. **上传文件**
```bash
# 使用 rsync
rsync -avz --delete build/ user@server:/var/www/docs.office.ninthfeast.com/

# 或使用 FTP/SFTP 工具上传 build 目录
```

3. **Nginx 配置**
```nginx
server {
    listen 80;
    listen [::]:80;
    server_name docs.office.ninthfeast.com;

    # 重定向到 HTTPS
    return 301 https://$server_name$request_uri;
}

server {
    listen 443 ssl http2;
    listen [::]:443 ssl http2;
    server_name docs.office.ninthfeast.com;

    # SSL 证书配置
    ssl_certificate /path/to/fullchain.pem;
    ssl_certificate_key /path/to/privkey.pem;

    # 网站根目录
    root /var/www/docs.office.ninthfeast.com;
    index index.html;

    # Gzip 压缩
    gzip on;
    gzip_types text/plain text/css application/json application/javascript text/xml application/xml application/xml+rss text/javascript;

    # 缓存配置
    location ~* \.(jpg|jpeg|png|gif|ico|css|js|svg|woff|woff2|ttf)$ {
        expires 1y;
        add_header Cache-Control "public, immutable";
    }

    # 主页面配置
    location / {
        try_files $uri $uri/ $uri.html /index.html;
    }

    # 安全头
    add_header X-Frame-Options "SAMEORIGIN" always;
    add_header X-Content-Type-Options "nosniff" always;
    add_header X-XSS-Protection "1; mode=block" always;
}
```

---

## 🔐 SSL 证书配置

### 使用 Let's Encrypt（免费）

```bash
# 安装 Certbot
sudo apt-get update
sudo apt-get install certbot python3-certbot-nginx

# 获取证书
sudo certbot --nginx -d docs.office.ninthfeast.com

# 自动续期
sudo certbot renew --dry-run
```

---

## 🔍 SEO 配置

### 1. Google Search Console

1. **添加资源**
   - 访问: https://search.google.com/search-console
   - 添加资源: `https://docs.office.ninthfeast.com`

2. **验证所有权**

   方式A - HTML 文件验证:
   ```bash
   # 下载验证文件
   # 放到 static/ 目录
   # 重新构建部署
   ```

   方式B - HTML 标签验证（推荐）:
   在 `docusaurus.config.ts` 的 `headTags` 中添加:
   ```ts
   {
     tagName: 'meta',
     attributes: {
       name: 'google-site-verification',
       content: 'YOUR_VERIFICATION_CODE',
     },
   }
   ```

3. **提交 Sitemap**
   ```
   https://docs.office.ninthfeast.com/sitemap.xml
   ```

### 2. 百度搜索资源平台

1. **站点验证**
   - 访问: https://ziyuan.baidu.com
   - 添加网站: `https://docs.office.ninthfeast.com`
   - 验证方式：HTML 标签或文件验证

2. **提交 Sitemap**
   ```
   自动推送: 添加自动推送代码
   手动提交: https://docs.office.ninthfeast.com/sitemap.xml
   ```

3. **百度站长 API 推送**
   ```bash
   curl -H 'Content-Type:text/plain' \
        --data-binary @urls.txt \
        "http://data.zz.baidu.com/urls?site=https://docs.office.ninthfeast.com&token=YOUR_TOKEN"
   ```

### 3. Bing Webmaster Tools

1. **导入 Google Search Console 数据**（推荐）
   - 访问: https://www.bing.com/webmasters
   - 选择"从 Google 导入"

2. **或手动添加**
   - 添加网站
   - 验证所有权
   - 提交 sitemap

---

## 📊 分析工具配置

### Google Analytics 4

1. **创建账号**
   - 访问: https://analytics.google.com
   - 创建新属性: `Office Mastery Guide`

2. **添加跟踪代码**

   在 `docusaurus.config.ts` 中添加:
   ```ts
   presets: [
     [
       'classic',
       {
         // ...
         googleAnalytics: {
           trackingID: 'G-XXXXXXXXXX',
           anonymizeIP: true,
         },
       },
     ],
   ],
   ```

   或使用插件:
   ```ts
   plugins: [
     [
       '@docusaurus/plugin-google-gtag',
       {
         trackingID: 'G-XXXXXXXXXX',
         anonymizeIP: true,
       },
    ],
   ],
   ```

### 百度统计

在 `static/index.html` 或创建自定义插件添加：
```html
<script>
var _hmt = _hmt || [];
(function() {
  var hm = document.createElement("script");
  hm.src = "https://hm.baidu.com/hm.js?YOUR_TRACKING_ID";
  var s = document.getElementsByTagName("script")[0];
  s.parentNode.insertBefore(hm, s);
})();
</script>
```

---

## 🎯 部署后检查清单

### 立即检查

- [ ] 网站可正常访问
- [ ] HTTPS 工作正常
- [ ] robots.txt 可访问: `/robots.txt`
- [ ] sitemap.xml 可访问: `/sitemap.xml`
- [ ] 所有页面正常显示
- [ ] 图片加载正常
- [ ] 移动端适配正常

### 24小时内

- [ ] 提交到 Google Search Console
- [ ] 提交到百度搜索资源平台
- [ ] 提交到 Bing Webmaster
- [ ] 配置 Google Analytics
- [ ] 检查页面加载速度（PageSpeed Insights）

### 一周内

- [ ] 检查搜索引擎抓取情况
- [ ] 查看 Google Analytics 数据
- [ ] 检查是否有抓取错误
- [ ] 监控网站性能

---

## 🔧 常见问题

### Q: 部署后 404 错误
A: 检查：
1. 服务器配置是否正确
2. 文件路径是否正确
3. Nginx/Apache 重写规则

### Q: sitemap 无法访问
A: 确保：
1. `sitemap.xml` 在 build 目录中
2. 服务器允许访问 .xml 文件
3. 路径配置正确

### Q: HTTPS 无法访问
A: 检查：
1. SSL 证书是否正确安装
2. 证书是否过期
3. 防火墙是否开放 443 端口

### Q: 搜索引擎未收录
A: 原因：
1. 刚上线，等待 1-4 周
2. robots.txt 配置错误
3. 未提交 sitemap
4. 内容质量问题

---

## 📈 性能优化

### CDN 配置

使用 Cloudflare CDN:
1. 注册 Cloudflare
2. 添加域名
3. 更新 DNS 记录
4. 开启 CDN 加速

### 图片优化

```bash
# 批量压缩图片
npm install -g imagemin-cli

imagemin static/img/* --out-dir=static/img-optimized
```

### 缓存策略

```nginx
# 浏览器缓存
location ~* \.(jpg|jpeg|png|gif|ico|css|js)$ {
    expires 1y;
    add_header Cache-Control "public, immutable";
}
```

---

## 🔄 更新发布流程

### 1. 本地测试
```bash
npm run build
npm run serve
npm run seo:check
```

### 2. 提交代码
```bash
git add .
git commit -m "feat: 更新内容"
git push origin main
```

### 3. 自动部署
- Vercel/Netlify 会自动部署
- 或手动触发部署

### 4. 验证
- 检查生产环境
- 清除 CDN 缓存（如有）

---

## 📞 技术支持

**作者**: lhqs
**邮箱**: lhqs1314@gmail.com
**网站**: https://docs.office.ninthfeast.com
**GitHub**: https://github.com/lhqs/office-mastery-guide

---

## 📝 更新日志

### 2025-12-22 - v1.0
- ✅ 初始部署
- ✅ SEO 完整优化
- ✅ 结构化数据配置
- ✅ 搜索引擎提交

---

**最后更新**: 2025-12-22
**部署状态**: ✅ 已优化
