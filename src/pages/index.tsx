import React from 'react';
import clsx from 'clsx';
import Link from '@docusaurus/Link';
import useDocusaurusContext from '@docusaurus/useDocusaurusContext';
import Layout from '@theme/Layout';
import HomepageFeatures from '@site/src/components/HomepageFeatures';
import StructuredData from '@site/src/components/StructuredData';

import styles from './index.module.css';

type ModuleCard = {
  title: string;
  description: string;
  link: string;
  tag: string;
  accent: 'teal' | 'blue' | 'indigo' | 'orange' | 'pink' | 'green' | 'purple' | 'slate';
};

type QuickLink = {
  title: string;
  description: string;
  link: string;
};

const moduleCards: ModuleCard[] = [
  {
    title: '导览与学习路径',
    description: '学习方法、能力模型与资源索引，快速进入状态。',
    link: '/docs/overview/00-章节目录与学习路径',
    tag: 'Start',
    accent: 'teal',
  },
  {
    title: 'Word 文档处理',
    description: '样式系统、长文档、批量处理与自动化全覆盖。',
    link: '/docs/word/06-Word界面与基础操作精通',
    tag: '排版',
    accent: 'blue',
  },
  {
    title: 'Excel 数据分析',
    description: '函数、透视、Power 系列、VBA 与 Python 集成。',
    link: '/docs/excel/16-Excel界面与数据输入规范',
    tag: '数据',
    accent: 'indigo',
  },
  {
    title: 'PPT 设计与呈现',
    description: '版式、动画、演讲表现与模板复用，提升说服力。',
    link: '/docs/ppt/31-PPT设计基础与原则',
    tag: '表达',
    accent: 'orange',
  },
  {
    title: '沟通与时间管理',
    description: 'Outlook 邮件、日历、任务与自动化处理的全流程。',
    link: '/docs/communication/41-Outlook邮件管理精通',
    tag: '沟通',
    accent: 'pink',
  },
  {
    title: '协作与远程办公',
    description: 'Google Workspace、Microsoft 365 与国内协同生态。',
    link: '/docs/collaboration/46-Google-Workspace全解析',
    tag: '协作',
    accent: 'teal',
  },
  {
    title: '效率与自动化',
    description: '快捷键、插件、RPA、Python 与个人效率系统搭建。',
    link: '/docs/automation/56-快捷键大全与记忆技巧',
    tag: '自动化',
    accent: 'green',
  },
  {
    title: '行业方案与标准化',
    description: '财务、人力、销售、项目等业务场景的模板化落地。',
    link: '/docs/solutions/66-财务会计办公软件应用',
    tag: '行业',
    accent: 'purple',
  },
  {
    title: '治理、安全与未来',
    description: '安全合规、速查索引、社区资源与趋势展望。',
    link: '/docs/ops-trends/76-数据安全与隐私保护',
    tag: '趋势',
    accent: 'slate',
  },
];

const quickLinks: QuickLink[] = [
  {
    title: 'Office 快捷键速查表',
    description: 'Word / Excel / PPT / Outlook 常用快捷键一览。',
    link: '/docs/ops-trends/86-Office快捷键速查表',
  },
  {
    title: '常用函数速查手册',
    description: '财务、数据分析、逻辑等高频函数按场景整理。',
    link: '/docs/ops-trends/87-常用函数速查手册',
  },
  {
    title: '模板资源库',
    description: '高质量文档、报表、PPT 模板和表单清单。',
    link: '/docs/ops-trends/88-模板资源库',
  },
  {
    title: '学习资源与社区',
    description: '视频课程、书单、播客与最佳实践社区导航。',
    link: '/docs/ops-trends/89-学习资源与社区',
  },
];

function ModuleCardItem({title, description, link, tag, accent}: ModuleCard) {
  return (
    <Link to={link} className={clsx(styles.moduleCard, styles[accent])}>
      <div className={styles.pill}>{tag}</div>
      <h3>{title}</h3>
      <p>{description}</p>
      <span className={styles.cardLink}>进入模块 →</span>
    </Link>
  );
}

function QuickLinkCard({title, description, link}: QuickLink) {
  return (
    <Link to={link} className={styles.quickCard}>
      <div>
        <h4>{title}</h4>
        <p>{description}</p>
      </div>
      <span className={styles.cardLink}>查看</span>
    </Link>
  );
}

export default function Home(): React.ReactElement {
  const {siteConfig} = useDocusaurusContext();

  return (
    <Layout
      title="办公软件精通指南 - Word Excel PPT 全栈教程 | Office Mastery Guide"
      description="全面系统的办公软件学习指南，涵盖Word文档处理、Excel数据分析、PPT演示设计、协作工具与办公自动化。90个章节助你从入门到精通，提升职场竞争力。">
      <StructuredData type="WebSite" />
      <StructuredData type="Course" />
      <StructuredData type="Organization" />
      <header className={styles.hero}>
        <div className={styles.heroGlow} />
        <div className="container">
          <div className={styles.heroContent}>
            <span className={styles.heroBadge}>
              <span className={styles.badge}>为每个追求高效的你</span>
            </span>
            <h1 className={styles.heroTitle}>
              让办公软件
              <br />
              成为你的超能力
            </h1>
            <p className={styles.heroSubtitle}>
              从基础到精通，从重复劳动到自动化，从个人效率到团队协作
              <br />
              90 个章节，系统化地帮你构建专业的办公软件技能体系
            </p>
            <div className={styles.actions}>
              <Link
                className="button button--primary button--lg"
                to="/docs/overview/00-章节目录与学习路径">
                开始学习之旅
              </Link>
              <Link
                className="button button--secondary button--lg"
                to="/docs/automation/56-快捷键大全与记忆技巧">
                探索自动化
              </Link>
            </div>
            <div className={styles.highlights}>
              <div className={styles.highlight}>
                <div className={styles.highlightIcon}>📚</div>
                <div>
                  <strong>90 章节</strong>
                  <span>系统化知识体系</span>
                </div>
              </div>
              <div className={styles.highlight}>
                <div className={styles.highlightIcon}>🚀</div>
                <div>
                  <strong>9 大模块</strong>
                  <span>覆盖所有核心场景</span>
                </div>
              </div>
              <div className={styles.highlight}>
                <div className={styles.highlightIcon}>⚡</div>
                <div>
                  <strong>自动化优先</strong>
                  <span>告别重复劳动</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </header>

      <main className={styles.main}>
        <section className={styles.featuredSection}>
          <div className="container">
            <div className={styles.sectionIntro}>
              <h2>核心学习模块</h2>
              <p>
                精心设计的学习路径，从工具掌握到效率提升，
                每个模块都是实战经验的结晶
              </p>
            </div>
            <div className={styles.moduleGrid}>
              {moduleCards.map((item) => (
                <ModuleCardItem key={item.title} {...item} />
              ))}
            </div>
          </div>
        </section>

        <HomepageFeatures />

        <section className={styles.resourceSection}>
          <div className="container">
            <div className={styles.sectionIntro}>
              <h2>快速参考资源</h2>
              <p>需要快速查找？这些资源为你准备好了</p>
            </div>
            <div className={styles.resourceGrid}>
              {quickLinks.map((item) => (
                <QuickLinkCard key={item.title} {...item} />
              ))}
            </div>
          </div>
        </section>

        <section className={styles.ctaSection}>
          <div className="container">
            <div className={styles.ctaCard}>
              <div className={styles.ctaContent}>
                <h2>准备好提升你的办公效率了吗？</h2>
                <p>
                  加入我们，开启系统化学习之旅。无论你是初学者还是进阶用户，
                  这里都有适合你的内容。
                </p>
                <Link
                  className="button button--primary button--lg"
                  to="/docs/overview/00-章节目录与学习路径">
                  立即开始
                </Link>
              </div>
              <div className={styles.ctaStats}>
                <div className={styles.statItem}>
                  <div className={styles.statValue}>Word</div>
                  <div className={styles.statLabel}>专业排版与文档处理</div>
                </div>
                <div className={styles.statItem}>
                  <div className={styles.statValue}>Excel</div>
                  <div className={styles.statLabel}>数据分析与自动化</div>
                </div>
                <div className={styles.statItem}>
                  <div className={styles.statValue}>PPT</div>
                  <div className={styles.statLabel}>演示设计与表达</div>
                </div>
                <div className={styles.statItem}>
                  <div className={styles.statValue}>更多</div>
                  <div className={styles.statLabel}>协作工具与效率系统</div>
                </div>
              </div>
            </div>
          </div>
        </section>
      </main>
    </Layout>
  );
}
