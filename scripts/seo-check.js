#!/usr/bin/env node

/**
 * SEO 检查脚本
 * 用于在构建后检查常见的 SEO 问题
 *
 * 运行: node scripts/seo-check.js
 */

const fs = require('fs');
const path = require('path');
const { glob } = require('glob');

const BUILD_DIR = path.join(__dirname, '../build');
const BASE_URL = 'https://docs.office.ninthfeast.com';

// 颜色输出
const colors = {
  reset: '\x1b[0m',
  red: '\x1b[31m',
  green: '\x1b[32m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
};

function log(message, color = 'reset') {
  console.log(`${colors[color]}${message}${colors.reset}`);
}

// 检查项目
const checks = {
  passed: [],
  warnings: [],
  errors: [],
};

async function checkHTMLFiles() {
  log('\n📄 检查 HTML 文件...', 'blue');

  const htmlFiles = await glob('**/*.html', { cwd: BUILD_DIR });

  log(`找到 ${htmlFiles.length} 个 HTML 文件`, 'green');

  for (const file of htmlFiles) {
    const filePath = path.join(BUILD_DIR, file);
    const content = fs.readFileSync(filePath, 'utf-8');

    // 检查 title
    const titleMatch = content.match(/<title>(.*?)<\/title>/);
    if (!titleMatch) {
      checks.errors.push(`${file}: 缺少 <title> 标签`);
    } else if (titleMatch[1].length < 30) {
      checks.warnings.push(`${file}: title 太短 (${titleMatch[1].length} 字符)`);
    } else if (titleMatch[1].length > 60) {
      checks.warnings.push(`${file}: title 太长 (${titleMatch[1].length} 字符)`);
    }

    // 检查 meta description
    const descMatch = content.match(/<meta name="description" content="(.*?)"/);
    if (!descMatch) {
      checks.errors.push(`${file}: 缺少 meta description`);
    } else if (descMatch[1].length < 120) {
      checks.warnings.push(`${file}: description 太短 (${descMatch[1].length} 字符)`);
    } else if (descMatch[1].length > 160) {
      checks.warnings.push(`${file}: description 太长 (${descMatch[1].length} 字符)`);
    }

    // 检查 h1
    const h1Matches = content.match(/<h1[^>]*>/g);
    if (!h1Matches) {
      checks.warnings.push(`${file}: 缺少 H1 标签`);
    } else if (h1Matches.length > 1) {
      checks.warnings.push(`${file}: 有多个 H1 标签 (${h1Matches.length})`);
    }

    // 检查图片 alt 属性
    const imgRegex = /<img[^>]*>/g;
    const images = content.match(imgRegex) || [];
    for (const img of images) {
      if (!img.includes('alt=')) {
        checks.warnings.push(`${file}: 图片缺少 alt 属性`);
      }
    }
  }

  checks.passed.push('HTML 文件检查完成');
}

function checkSitemap() {
  log('\n🗺️  检查 Sitemap...', 'blue');

  const sitemapPath = path.join(BUILD_DIR, 'sitemap.xml');

  if (!fs.existsSync(sitemapPath)) {
    checks.errors.push('sitemap.xml 不存在');
    return;
  }

  const content = fs.readFileSync(sitemapPath, 'utf-8');
  const urlMatches = content.match(/<url>/g);

  if (urlMatches) {
    log(`Sitemap 包含 ${urlMatches.length} 个 URL`, 'green');
    checks.passed.push('Sitemap 存在且有效');
  } else {
    checks.errors.push('Sitemap 格式错误');
  }

  // 检查是否包含正确的域名
  if (!content.includes(BASE_URL)) {
    checks.warnings.push(`Sitemap 中未找到域名: ${BASE_URL}`);
  }
}

function checkRobotsTxt() {
  log('\n🤖 检查 robots.txt...', 'blue');

  const robotsPath = path.join(BUILD_DIR, 'robots.txt');

  if (!fs.existsSync(robotsPath)) {
    checks.errors.push('robots.txt 不存在');
    return;
  }

  const content = fs.readFileSync(robotsPath, 'utf-8');

  if (content.includes('Sitemap:')) {
    checks.passed.push('robots.txt 包含 Sitemap 引用');
  } else {
    checks.warnings.push('robots.txt 缺少 Sitemap 引用');
  }

  if (content.includes('Disallow: /')) {
    checks.warnings.push('robots.txt 禁止所有爬虫（请检查是否正确）');
  }

  log('robots.txt 存在', 'green');
}

function checkStaticAssets() {
  log('\n🖼️  检查静态资源...', 'blue');

  const requiredFiles = [
    'img/favicon.svg',
    'img/logo.svg',
    'img/og-card.svg',
  ];

  for (const file of requiredFiles) {
    const filePath = path.join(BUILD_DIR, file);
    if (fs.existsSync(filePath)) {
      checks.passed.push(`${file} 存在`);
    } else {
      checks.warnings.push(`${file} 不存在`);
    }
  }
}

function printResults() {
  log('\n' + '='.repeat(60), 'blue');
  log('📊 SEO 检查结果', 'blue');
  log('='.repeat(60), 'blue');

  if (checks.passed.length > 0) {
    log(`\n✅ 通过检查 (${checks.passed.length})`, 'green');
    checks.passed.forEach(item => log(`  • ${item}`, 'green'));
  }

  if (checks.warnings.length > 0) {
    log(`\n⚠️  警告 (${checks.warnings.length})`, 'yellow');
    checks.warnings.forEach(item => log(`  • ${item}`, 'yellow'));
  }

  if (checks.errors.length > 0) {
    log(`\n❌ 错误 (${checks.errors.length})`, 'red');
    checks.errors.forEach(item => log(`  • ${item}`, 'red'));
  }

  log('\n' + '='.repeat(60), 'blue');

  const total = checks.passed.length + checks.warnings.length + checks.errors.length;
  const score = Math.round((checks.passed.length / total) * 100);

  log(`\n总体评分: ${score}/100`, score >= 80 ? 'green' : score >= 60 ? 'yellow' : 'red');
  log('');

  if (checks.errors.length > 0) {
    process.exit(1);
  }
}

// 主函数
async function main() {
  log('\n🔍 开始 SEO 检查...', 'blue');
  log(`构建目录: ${BUILD_DIR}`, 'blue');
  log(`网站域名: ${BASE_URL}`, 'blue');

  if (!fs.existsSync(BUILD_DIR)) {
    log('\n❌ 构建目录不存在，请先运行 npm run build', 'red');
    process.exit(1);
  }

  try {
    await checkHTMLFiles();
    checkSitemap();
    checkRobotsTxt();
    checkStaticAssets();
    printResults();
  } catch (error) {
    log(`\n❌ 检查过程出错: ${error.message}`, 'red');
    process.exit(1);
  }
}

main();
