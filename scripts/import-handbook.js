/* eslint-disable no-console */
const fs = require('fs');
const path = require('path');

const SOURCE_DIR = path.resolve(
  __dirname,
  '../../../lhqs-handbook/办公软件精通指南',
);
const DOCS_ROOT = path.resolve(__dirname, '../docs');

const buckets = [
  {
    id: 'overview',
    label: '导览与学习路径',
    position: 1,
    range: [0, 5],
    description: '学习路径、能力模型与资源导航，帮助快速上手整套指南。',
  },
  {
    id: 'word',
    label: 'Word 文档处理',
    position: 2,
    range: [6, 15],
    description: '样式、长文档、批量处理与自动化，覆盖 Word 高阶场景。',
  },
  {
    id: 'excel',
    label: 'Excel 数据分析',
    position: 3,
    range: [16, 30],
    description: '数据分析、图表、Power Query/ Pivot、VBA 与 Python 集成。',
  },
  {
    id: 'ppt',
    label: 'PPT 设计与呈现',
    position: 4,
    range: [31, 40],
    description: '版式、动画、演讲与模板设计，让汇报既专业又有张力。',
  },
  {
    id: 'communication',
    label: '沟通与时间管理',
    position: 5,
    range: [41, 45],
    description: 'Outlook 邮件、日历、任务管理与自动化处理的全流程最佳实践。',
  },
  {
    id: 'collaboration',
    label: '协作与远程办公',
    position: 6,
    range: [46, 55],
    description: 'Google Workspace、Microsoft 365 与国内协同生态的高效协作方案。',
  },
  {
    id: 'automation',
    label: '效率与自动化',
    position: 7,
    range: [56, 65],
    description: '快捷键、插件、RPA、Python 与个人效率系统，构建自动化思维。',
  },
  {
    id: 'solutions',
    label: '行业方案与标准化',
    position: 8,
    range: [66, 75],
    description: '财务、人力、销售、项目等业务场景的标准化与模板化落地。',
  },
  {
    id: 'ops-trends',
    label: '治理、安全与未来',
    position: 9,
    range: [76, 90],
    description: '安全合规、速查索引、资源社区与未来趋势的全景观察。',
  },
];

function getBucket(num) {
  const bucket = buckets.find(
    (candidate) => num >= candidate.range[0] && num <= candidate.range[1],
  );
  if (!bucket) {
    throw new Error(`未找到适配分组: ${num}`);
  }
  return bucket;
}

function ensureCategoryMeta(bucket) {
  const dir = path.join(DOCS_ROOT, bucket.id);
  fs.mkdirSync(dir, {recursive: true});
  const metaPath = path.join(dir, '_category_.json');
  const meta = {
    label: bucket.label,
    position: bucket.position,
    link: {
      type: 'generated-index',
      title: bucket.label,
      description: bucket.description,
      slug: `/${bucket.id}`,
    },
  };
  fs.writeFileSync(metaPath, `${JSON.stringify(meta, null, 2)}\n`, 'utf8');
}

function normalizeHeading(content, cleanTitle) {
  const sanitized = content.replace(/^\uFEFF/, '');
  const lines = sanitized.split(/\r?\n/);
  const headingIndex = lines.findIndex((line) => line.trim().startsWith('#'));
  if (headingIndex >= 0) {
    lines[headingIndex] = `# ${cleanTitle}`;
  } else {
    lines.unshift(`# ${cleanTitle}`, '');
  }
  return lines.join('\n');
}

function sanitizeForMdx(content) {
  const lines = content.split(/\r?\n/);
  let inFence = false;
  return lines
    .map((line) => {
      const fenceMatch = line.trimStart().match(/^(```|~~~)/);
      if (fenceMatch) {
        inFence = !inFence;
        return line;
      }
      if (inFence) {
        return line;
      }

      const hasBlockquote = /^\s*>/.test(line);
      const placeholder = '__BLOCKQUOTE__';
      const preserved = hasBlockquote
        ? line.replace(/^\s*>/, (match) => match.replace('>', placeholder))
        : line;

      const escapedAngles = preserved
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
      const escapedBraces = escapedAngles
        .replace(/{/g, '&#123;')
        .replace(/}/g, '&#125;');

      return hasBlockquote
        ? escapedBraces.replace(placeholder, '>')
        : escapedBraces;
    })
    .join('\n');
}

function buildSlug(bucketId, numString, rawTitle) {
  const normalizedTitle = rawTitle.replace(/\s+/g, '-');
  return `/${bucketId}/${numString}-${normalizedTitle}`;
}

function buildFrontmatter(title, position, slug) {
  return [
    '---',
    `title: ${title}`,
    `slug: ${slug}`,
    `sidebar_position: ${position}`,
    '---',
    '',
  ].join('\n');
}

function importDocs() {
  if (!fs.existsSync(SOURCE_DIR)) {
    throw new Error(`未找到源目录: ${SOURCE_DIR}`);
  }

  fs.rmSync(DOCS_ROOT, {recursive: true, force: true});
  fs.mkdirSync(DOCS_ROOT, {recursive: true});

  const files = fs
    .readdirSync(SOURCE_DIR)
    .filter((file) => file.endsWith('.md'))
    .sort((a, b) => {
      const [, aNum] = a.match(/^(\d+)-/) || [];
      const [, bNum] = b.match(/^(\d+)-/) || [];
      return Number(aNum) - Number(bNum);
    });

  const createdCategories = new Set();
  let imported = 0;

  for (const fileName of files) {
    const match = fileName.match(/^(\d+)-(.+)\.md$/);
    if (!match) {
      console.warn(`跳过无法解析的文件名: ${fileName}`);
      continue;
    }
    const [, numString, rawTitle] = match;
    const num = Number(numString);
    const bucket = getBucket(num);

    if (!createdCategories.has(bucket.id)) {
      ensureCategoryMeta(bucket);
      createdCategories.add(bucket.id);
    }

    const destDir = path.join(DOCS_ROOT, bucket.id);
    const destPath = path.join(destDir, fileName);
    const sourcePath = path.join(SOURCE_DIR, fileName);
    const fileContent = fs.readFileSync(sourcePath, 'utf8');

    const normalized = normalizeHeading(fileContent, rawTitle);
    const sanitized = sanitizeForMdx(normalized);
    const slug = buildSlug(bucket.id, numString, rawTitle);
    const finalContent = `${buildFrontmatter(
      rawTitle,
      num,
      slug,
    )}${sanitized.trimStart()}\n`;
    fs.writeFileSync(destPath, finalContent, 'utf8');
    imported += 1;
  }

  console.log(`✅ 已导入 ${imported} 篇文档到 ${DOCS_ROOT}`);
}

importDocs();
