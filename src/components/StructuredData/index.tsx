import React from 'react';
import Head from '@docusaurus/Head';

interface StructuredDataProps {
  type?: 'WebSite' | 'Course' | 'Article' | 'Organization';
}

export default function StructuredData({type = 'WebSite'}: StructuredDataProps): JSX.Element {
  const baseUrl = 'https://docs.office.ninthfeast.com';

  const getStructuredData = () => {
    switch (type) {
      case 'WebSite':
        return {
          '@context': 'https://schema.org',
          '@type': 'WebSite',
          name: '办公软件精通指南',
          alternateName: 'Office Mastery Guide',
          url: baseUrl,
          description: '全面系统的办公软件学习指南，涵盖Word、Excel、PPT及协作工具，90个章节助你从入门到精通。',
          author: {
            '@type': 'Person',
            name: 'lhqs',
            email: 'lhqs1314@gmail.com',
          },
          publisher: {
            '@type': 'Organization',
            name: 'Office Mastery Guide',
            url: baseUrl,
          },
          inLanguage: 'zh-CN',
          potentialAction: {
            '@type': 'SearchAction',
            target: {
              '@type': 'EntryPoint',
              urlTemplate: `${baseUrl}/search?q={search_term_string}`,
            },
            'query-input': 'required name=search_term_string',
          },
        };

      case 'Course':
        return {
          '@context': 'https://schema.org',
          '@type': 'Course',
          name: '办公软件精通指南',
          description: '系统化学习Word、Excel、PPT等办公软件的全栈教程',
          provider: {
            '@type': 'Organization',
            name: 'Office Mastery Guide',
            sameAs: baseUrl,
          },
          hasCourseInstance: {
            '@type': 'CourseInstance',
            courseMode: 'online',
            courseWorkload: 'P90D',
          },
          educationalLevel: '初级到高级',
          teaches: [
            'Word文档处理',
            'Excel数据分析',
            'PPT演示设计',
            '办公自动化',
            'VBA编程',
            'Python办公应用',
          ],
        };

      case 'Organization':
        return {
          '@context': 'https://schema.org',
          '@type': 'Organization',
          name: 'Office Mastery Guide',
          url: baseUrl,
          logo: `${baseUrl}/img/logo.svg`,
          description: '专注于办公软件技能培训的在线教育平台',
          contactPoint: {
            '@type': 'ContactPoint',
            email: 'lhqs1314@gmail.com',
            contactType: 'customer service',
          },
          sameAs: [
            // 可以添加社交媒体链接
          ],
        };

      default:
        return null;
    }
  };

  const structuredData = getStructuredData();

  if (!structuredData) return null;

  return (
    <Head>
      <script type="application/ld+json">
        {JSON.stringify(structuredData)}
      </script>
    </Head>
  );
}

// 面包屑导航结构化数据
export function BreadcrumbStructuredData({items}: {items: Array<{name: string; url: string}>}): JSX.Element {
  const breadcrumbData = {
    '@context': 'https://schema.org',
    '@type': 'BreadcrumbList',
    itemListElement: items.map((item, index) => ({
      '@type': 'ListItem',
      position: index + 1,
      name: item.name,
      item: item.url,
    })),
  };

  return (
    <Head>
      <script type="application/ld+json">
        {JSON.stringify(breadcrumbData)}
      </script>
    </Head>
  );
}

// 文章结构化数据
export function ArticleStructuredData({
  title,
  description,
  author = 'lhqs',
  datePublished,
  dateModified,
  url,
  image,
}: {
  title: string;
  description: string;
  author?: string;
  datePublished?: string;
  dateModified?: string;
  url: string;
  image?: string;
}): JSX.Element {
  const baseUrl = 'https://docs.office.ninthfeast.com';

  const articleData = {
    '@context': 'https://schema.org',
    '@type': 'Article',
    headline: title,
    description: description,
    author: {
      '@type': 'Person',
      name: author,
      email: 'lhqs1314@gmail.com',
    },
    publisher: {
      '@type': 'Organization',
      name: 'Office Mastery Guide',
      logo: {
        '@type': 'ImageObject',
        url: `${baseUrl}/img/logo.svg`,
      },
    },
    datePublished: datePublished || new Date().toISOString(),
    dateModified: dateModified || new Date().toISOString(),
    mainEntityOfPage: {
      '@type': 'WebPage',
      '@id': url,
    },
    image: image || `${baseUrl}/img/og-card.svg`,
  };

  return (
    <Head>
      <script type="application/ld+json">
        {JSON.stringify(articleData)}
      </script>
    </Head>
  );
}
