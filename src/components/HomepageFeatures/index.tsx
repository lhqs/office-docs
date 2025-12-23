import React from 'react';
import styles from './styles.module.css';

type FeatureItem = {
  title: string;
  description: string;
  tag: string;
};

const FeatureList: FeatureItem[] = [
  {
    title: '一眼可见的学习蓝图',
    description:
      '从基础认知到行业方案的 9 大模块，按业务场景和能力成熟度排列，随时跳转到需要的章节。',
    tag: '路径',
  },
  {
    title: '自动化与智能驱动',
    description:
      '涵盖快捷键、宏、VBA、Python、RPA 与 AI 辅助生成，构建可复用的自动化组件库。',
    tag: '自动化',
  },
  {
    title: '设计与表达力升级',
    description:
      '用叙事结构、版式与数据可视化的组合拳，让 Word/PPT/报告在审美与说服力上同步提升。',
    tag: '设计力',
  },
  {
    title: '协作与治理并重',
    description:
      '邮件、日历、项目工具、版本与安全治理全部打通，确保个人高效同时兼顾团队协同。',
    tag: '协作',
  },
];

function Feature({title, description, tag}: FeatureItem) {
  return (
    <div className={styles.card}>
      <span className={styles.badge}>{tag}</span>
      <h3>{title}</h3>
      <p>{description}</p>
    </div>
  );
}

export default function HomepageFeatures(): React.ReactElement {
  return (
    <section className={styles.section}>
      <div className="container">
        <div className={styles.header}>
          <p className={styles.eyebrow}>方法论</p>
          <h2>把高级办公能力拆成可复制的工作流</h2>
          <p>
            每个章节都给出可操作的清单、快捷键或模板链接，尽量减少“只讲概念”的描述，
            让你直接复制到日常工作中。
          </p>
        </div>
        <div className={styles.grid}>
          {FeatureList.map((props) => (
            <Feature key={props.title} {...props} />
          ))}
        </div>
      </div>
    </section>
  );
}
