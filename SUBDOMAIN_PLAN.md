# 三级域名规划方案

## 核心域名建议

**主域名**：`officemastery.com` 或 `officeplus.com` 或 `officehub.com`

## 三级域名架构

### Phase 1 - 基础架构（当前阶段）

```
officemastery.com              # 主站首页
├── www.officemastery.com      # 主站（可选，重定向到根域名）
└── docs.officemastery.com     # 文档中心（当前 Docusaurus 站点）
```

### Phase 2 - 资源扩展

```
├── templates.officemastery.com # 模板资源库
│   ├── /word                   # Word 模板
│   ├── /excel                  # Excel 模板
│   └── /ppt                    # PPT 模板
│
├── tools.officemastery.com     # 在线工具集
│   ├── /excel-calculator       # Excel 公式计算器
│   ├── /format-converter       # 格式转换工具
│   └── /color-picker           # 配色方案生成器
│
└── shortcuts.officemastery.com # 快捷键速查站
    └── 交互式快捷键查询工具
```

### Phase 3 - 社区与内容

```
├── blog.officemastery.com      # 技术博客
│   └── 办公技巧、案例分析、趋势解读
│
├── community.officemastery.com # 社区论坛
│   ├── 问答板块
│   ├── 经验分享
│   └── 作品展示
│
└── learn.officemastery.com     # 学习平台
    ├── 视频课程
    ├── 交互式练习
    └── 学习路径追踪
```

### Phase 4 - 高级功能

```
├── auto.officemastery.com      # 自动化脚本库
│   ├── VBA 脚本
│   ├── Python 脚本
│   └── Power Automate 模板
│
├── ai.officemastery.com        # AI 辅助工具
│   ├── 智能排版
│   ├── 数据分析助手
│   └── 演示稿生成器
│
└── api.officemastery.com       # API 服务
    └── 开发者文档和接口
```

## 域名命名规范

### 命名原则
1. 简短（5-10个字符）
2. 语义化（见名知意）
3. 小写字母
4. 避免使用数字和连字符
5. SEO 友好

### 推荐词汇库

**功能类**：
- docs, guide, manual, handbook
- learn, course, training, academy
- tools, utils, kit, lab
- templates, resources, assets
- blog, news, insights

**场景类**：
- word, excel, ppt, outlook
- auto, automation, scripts
- collab, team, workspace

**用户类**：
- pro, expert, master
- beginner, starter
- corp, enterprise, edu

## 技术实施

### Nginx 配置示例

```nginx
# 主站
server {
    server_name officemastery.com www.officemastery.com;
    root /var/www/homepage;
}

# 文档站
server {
    server_name docs.officemastery.com;
    root /var/www/docs/build;
}

# 工具站
server {
    server_name tools.officemastery.com;
    root /var/www/tools;
}

# 模板库
server {
    server_name templates.officemastery.com;
    root /var/www/templates;
}
```

### Cloudflare/CDN 配置

```yaml
DNS Records:
  - Type: A
    Name: @
    Content: <主服务器IP>
    Proxied: Yes

  - Type: CNAME
    Name: docs
    Content: officemastery.com
    Proxied: Yes

  - Type: CNAME
    Name: tools
    Content: officemastery.com
    Proxied: Yes

  - Type: CNAME
    Name: templates
    Content: officemastery.com
    Proxied: Yes
```

## SEO 优化建议

### 站点地图结构
```
https://officemastery.com/sitemap.xml
https://docs.officemastery.com/sitemap.xml
https://tools.officemastery.com/sitemap.xml
https://templates.officemastery.com/sitemap.xml
```

### robots.txt 配置
```
User-agent: *
Allow: /

Sitemap: https://officemastery.com/sitemap.xml
Sitemap: https://docs.officemastery.com/sitemap.xml
```

## 品牌一致性

### Logo 与配色
- 所有子站使用统一的主色调：#2c7df0
- 统一的导航栏和 Footer
- 统一的字体系统

### 域名展示
```
页面标题格式：
- 主站：办公软件精通指南 | Office Mastery
- 文档：Word 精通 - 文档中心 | Office Mastery
- 工具：Excel 计算器 | Office Mastery Tools
```

## 监控与分析

### Google Analytics
为每个子域名创建独立的数据视图：
- officemastery.com - 主站
- docs.officemastery.com - 文档
- tools.officemastery.com - 工具
- templates.officemastery.com - 模板

### 性能监控
使用 Uptime Robot 或 StatusCake 监控各子站的可用性

## 未来扩展方向

### 多语言支持
```
en.officemastery.com    # 英文版
zh.officemastery.com    # 简体中文
tw.officemastery.com    # 繁体中文
ja.officemastery.com    # 日文
```

### 地区化部署
```
us.officemastery.com    # 美国节点
eu.officemastery.com    # 欧洲节点
asia.officemastery.com  # 亚洲节点
```

## 成本估算

### 域名成本
- 主域名：$10-15/年
- SSL 证书：Let's Encrypt 免费或 $50-200/年（通配符证书）

### 服务器成本
- 轻量级服务器：$5-20/月
- CDN 流量：按需付费或免费套餐（Cloudflare）

### 总计
初期：约 $100-300/年
扩展期：约 $500-1000/年

---

**更新日期**：2025-12-22
**版本**：v1.0
**负责人**：运营团队
