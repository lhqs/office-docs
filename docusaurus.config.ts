import {themes as prismThemes} from 'prism-react-renderer';
import type {Config} from '@docusaurus/types';
import type * as Preset from '@docusaurus/preset-classic';

// This runs in Node.js - Don't use client-side code here (browser APIs, JSX...)

const config: Config = {
  title: '办公软件精通指南 - Word Excel PPT 全栈教程',
  tagline: '从基础到精通，打造专业办公技能体系 | 自动化 · 协作 · 效率提升',
  favicon: 'img/favicon.svg',

  // Future flags, see https://docusaurus.io/docs/api/docusaurus-config#future
  future: {
    v4: true, // Improve compatibility with the upcoming Docusaurus v4
  },

  // Set the production url of your site here
  url: 'https://docs.office.ninthfeast.com',
  // Set the /<baseUrl>/ pathname under which your site is served
  // For GitHub pages deployment, it is often '/<projectName>/'
  baseUrl: '/',

  // GitHub pages deployment config.
  // If you aren't using GitHub pages, you don't need these.
  organizationName: 'lhqs', // Usually your GitHub org/user name.
  projectName: 'office-mastery-guide', // Usually your repo name.

  onBrokenLinks: 'throw',
  onBrokenMarkdownLinks: 'warn',

  // SEO 优化
  headTags: [
    {
      tagName: 'meta',
      attributes: {
        name: 'keywords',
        content: 'Office教程,Word教程,Excel教程,PPT教程,办公软件,办公自动化,VBA,Python办公,数据分析,文档排版,演示设计,协作工具,效率提升,办公技巧',
      },
    },
    {
      tagName: 'meta',
      attributes: {
        name: 'author',
        content: 'lhqs',
      },
    },
    {
      tagName: 'meta',
      attributes: {
        name: 'contact',
        content: 'lhqs1314@gmail.com',
      },
    },
    {
      tagName: 'meta',
      attributes: {
        property: 'og:type',
        content: 'website',
      },
    },
    {
      tagName: 'meta',
      attributes: {
        property: 'og:site_name',
        content: '办公软件精通指南',
      },
    },
    {
      tagName: 'meta',
      attributes: {
        name: 'twitter:card',
        content: 'summary_large_image',
      },
    },
    {
      tagName: 'link',
      attributes: {
        rel: 'canonical',
        href: 'https://docs.office.ninthfeast.com',
      },
    },
  ],

  // Even if you don't use internationalization, you can use this field to set
  // useful metadata like html lang. For example, if your site is Chinese, you
  // may want to replace "en" with "zh-Hans".
  i18n: {
    defaultLocale: 'zh-Hans',
    locales: ['zh-Hans'],
  },

  presets: [
    [
      'classic',
      {
        docs: {
          sidebarPath: './sidebars.ts',
          routeBasePath: 'docs',
          showLastUpdateAuthor: false,
          showLastUpdateTime: false,
          editUrl: 'https://github.com/lhqs/office-mastery-guide/edit/main/',
          // SEO 优化
          breadcrumbs: true,
        },
        blog: false,
        theme: {
          customCss: './src/css/custom.css',
        },
        sitemap: {
          changefreq: 'weekly',
          priority: 0.5,
          ignorePatterns: ['/tags/**'],
          filename: 'sitemap.xml',
        },
      } satisfies Preset.Options,
    ],
  ],

  themeConfig: {
    // Replace with your project's social card
    image: 'img/og-card.svg',

    // SEO 元数据
    metadata: [
      {
        name: 'description',
        content: '全面系统的办公软件学习指南，涵盖Word文档处理、Excel数据分析、PPT演示设计、协作工具与办公自动化。90个章节助你从入门到精通，提升职场竞争力。',
      },
      {
        name: 'keywords',
        content: 'Office,Word,Excel,PPT,办公软件教程,VBA编程,Python办公自动化,数据透视表,邮件合并,演示设计,协作工具,效率工具',
      },
      {
        property: 'og:title',
        content: '办公软件精通指南 - 全栈Office技能体系',
      },
      {
        property: 'og:description',
        content: '90个章节系统化学习Word、Excel、PPT及协作工具，掌握办公自动化技能，提升工作效率。',
      },
      {
        property: 'og:url',
        content: 'https://docs.office.ninthfeast.com',
      },
      {
        property: 'og:image',
        content: 'https://docs.office.ninthfeast.com/img/og-card.svg',
      },
    ],

    colorMode: {
      respectPrefersColorScheme: true,
      defaultMode: 'light',
    },
    navbar: {
      title: '办公软件精通指南',
      logo: {
        alt: 'Office Mastery Guide - 办公软件精通指南',
        src: 'img/logo.svg',
      },
      hideOnScroll: false,
      items: [
        {
          type: 'docSidebar',
          sidebarId: 'guideSidebar',
          position: 'left',
          label: '文档',
        },
        {
          to: '/docs/overview/00-章节目录与学习路径',
          label: '学习路径',
          position: 'left',
        },
        {
          to: '/docs/automation/56-快捷键大全与记忆技巧',
          label: '效率与自动化',
          position: 'left',
        },
      ],
    },
    footer: {
      style: 'dark',
      links: [
        {
          title: '学习导航',
          items: [
            {
              label: '章节总览',
              to: '/docs/overview/00-章节目录与学习路径',
            },
            {
              label: 'Word 精通',
              to: '/docs/word/06-Word界面与基础操作精通',
            },
          ],
        },
        {
          title: '深度专题',
          items: [
            {
              label: 'Excel 数据分析',
              to: '/docs/excel/16-Excel界面与数据输入规范',
            },
            {label: 'PPT 设计力', to: '/docs/ppt/31-PPT设计基础与原则'},
            {label: '行业方案', to: '/docs/solutions/66-财务会计办公软件应用'},
          ],
        },
        {
          title: '资源与联系',
          items: [
            {
              label: '快捷索引',
              to: '/docs/ops-trends/86-Office快捷键速查表',
            },
            {
              label: '学习社区',
              to: '/docs/ops-trends/89-学习资源与社区',
            },
            {
              label: '联系作者',
              href: 'mailto:lhqs1314@gmail.com',
            },
          ],
        },
      ],
      copyright: `© ${new Date().getFullYear()} Office Mastery Guide by <a href="mailto:lhqs1314@gmail.com" target="_blank" rel="noopener noreferrer">lhqs</a> · 构建于 <a href="https://docusaurus.io/" target="_blank" rel="noopener noreferrer">Docusaurus</a>`,
    },
    prism: {
      theme: prismThemes.github,
      darkTheme: prismThemes.dracula,
    },
  } satisfies Preset.ThemeConfig,
};

export default config;
