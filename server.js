require('dotenv').config();
const express = require('express');
const path = require('path');
const puppeteer = require('puppeteer');
const PptxGenJS = require('pptxgenjs');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json({ limit: '50mb' }));
app.use(express.static('public'));

const analysisCache = new Map();

const AVERAGE_SCORES = {
  induction: 5.0,
  classification: 5.0,
  content: 7.0,
  functionality: 3.0
};

// ============================================================
// POST /api/analyze - Analyze a FAQ site (rule-based)
// ============================================================
app.post('/api/analyze', async (req, res) => {
  const { url } = req.body;
  if (!url) return res.status(400).json({ error: 'URLが必要です' });

  try {
    console.log(`[分析開始] ${url}`);
    const { screenshot, siteData } = await scrapeAndAnalyze(url);
    console.log(`[スクレイピング完了]`);

    const analysis = evaluateSite(siteData, url);
    console.log(`[ルールベース分析完了] スコア: ${analysis.totalScore}/28`);

    const analysisId = Date.now().toString();
    const result = { analysisId, url, screenshot, averageScores: AVERAGE_SCORES, ...analysis };

    analysisCache.set(analysisId, result);
    if (analysisCache.size > 50) {
      const keys = [...analysisCache.keys()];
      for (let i = 0; i < keys.length - 50; i++) analysisCache.delete(keys[i]);
    }

    res.json(result);
  } catch (error) {
    console.error('[エラー]', error);
    res.status(500).json({ error: error.message || '分析中にエラーが発生しました' });
  }
});

// ============================================================
// POST /api/generate-pptx
// ============================================================
app.post('/api/generate-pptx', async (req, res) => {
  const { analysisId, radarChartImage } = req.body;
  const data = analysisCache.get(analysisId);
  if (!data) return res.status(404).json({ error: '分析結果が見つかりません。再度分析を実行してください。' });

  try {
    const pptxBuffer = await generatePptx(data, radarChartImage);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', 'attachment; filename="FAQ_diagnostic_report.pptx"');
    res.send(Buffer.from(pptxBuffer));
  } catch (error) {
    console.error('[PPTX生成エラー]', error);
    res.status(500).json({ error: 'PPTX生成中にエラーが発生しました' });
  }
});

// ============================================================
// Puppeteer: Scrape site and extract structured data
// ============================================================
async function scrapeAndAnalyze(url) {
  const launchOptions = {
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--lang=ja']
  };
  if (process.env.PUPPETEER_EXECUTABLE_PATH) {
    launchOptions.executablePath = process.env.PUPPETEER_EXECUTABLE_PATH;
  }
  const browser = await puppeteer.launch(launchOptions);

  try {
    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 800 });
    await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36');
    await page.goto(url, { waitUntil: 'networkidle2', timeout: 30000 });
    await new Promise(r => setTimeout(r, 2000));

    const screenshot = await page.screenshot({ encoding: 'base64', type: 'png' });

    // Extract comprehensive site data via DOM analysis
    const siteData = await page.evaluate(() => {
      const body = document.body;
      const html = document.documentElement;

      // Helper: check if element is visible
      function isVisible(el) {
        const style = getComputedStyle(el);
        return style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0';
      }

      // Helper: get text content cleaned
      function cleanText(el) {
        return el ? el.textContent.replace(/\s+/g, ' ').trim() : '';
      }

      // ---- Title & Meta ----
      const pageTitle = document.title || '';
      const metaDesc = document.querySelector('meta[name="description"]')?.content || '';
      const ogSiteName = document.querySelector('meta[property="og:site_name"]')?.content || '';

      // ---- Company name extraction ----
      const logoAlt = document.querySelector('header img, .logo img, [class*="logo"] img, h1 img')?.alt || '';
      const h1Text = document.querySelector('h1')?.textContent?.trim() || '';
      const domain = location.hostname.replace('www.', '').split('.')[0];

      // ---- Navigation / Header analysis ----
      const headerEl = document.querySelector('header, [role="banner"], #header, .header');
      const navLinks = [...document.querySelectorAll('nav a, header a, .nav a, .menu a, .navigation a')];
      const allLinks = [...document.querySelectorAll('a')];

      // FAQ link in navigation/header
      const faqLinksInNav = navLinks.filter(a => {
        const text = a.textContent.toLowerCase();
        const href = (a.href || '').toLowerCase();
        return text.includes('faq') || text.includes('よくある') || text.includes('質問') ||
               href.includes('faq') || href.includes('question');
      });

      // FAQ link anywhere on page (top portion = above fold)
      const faqLinksAll = allLinks.filter(a => {
        const text = a.textContent.toLowerCase();
        const href = (a.href || '').toLowerCase();
        return text.includes('faq') || text.includes('よくある') || text.includes('質問') ||
               href.includes('faq') || href.includes('question');
      });

      // ---- Contact / Inquiry relationship ----
      const contactLinks = allLinks.filter(a => {
        const text = (a.textContent || '').toLowerCase();
        const href = (a.href || '').toLowerCase();
        return text.includes('問い合わせ') || text.includes('お問合せ') || text.includes('contact') ||
               text.includes('inquiry') || href.includes('contact') || href.includes('inquiry') ||
               href.includes('form') || text.includes('相談');
      });

      // Check if FAQ is shown before/near contact
      const hasContactWithFaqCushion = contactLinks.some(link => {
        const parent = link.closest('section, div, aside, .contact, [class*="contact"]');
        if (!parent) return false;
        const parentText = parent.textContent.toLowerCase();
        return parentText.includes('faq') || parentText.includes('よくある') || parentText.includes('質問');
      });

      // ---- FAQ-specific heading prominence ----
      const headings = [...document.querySelectorAll('h1, h2, h3')];
      const faqHeadings = headings.filter(h => {
        const text = h.textContent.toLowerCase();
        return text.includes('faq') || text.includes('よくある') || text.includes('質問');
      });
      const faqInH1orH2 = faqHeadings.some(h => h.tagName === 'H1' || h.tagName === 'H2');

      // ---- Brand / Design consistency ----
      const hasFavicon = !!document.querySelector('link[rel*="icon"]');
      const hasCustomCSS = document.styleSheets.length > 0;
      const headerBg = headerEl ? getComputedStyle(headerEl).backgroundColor : '';
      const bodyBg = getComputedStyle(body).backgroundColor;
      const hasLogo = !!document.querySelector('header img, .logo img, [class*="logo"] img, h1 img, [class*="brand"] img');

      // ---- Category analysis ----
      const categoryElements = [
        ...document.querySelectorAll('.category, [class*="category"], [class*="cat-"], [class*="カテゴリ"]'),
        ...document.querySelectorAll('.tab, [role="tab"], .nav-tabs .nav-link, .tab-item'),
        ...document.querySelectorAll('[class*="genre"], [class*="topic"], [class*="section-nav"]')
      ];

      // Also check for visual category blocks (links that look like categories)
      const possibleCatLinks = [...document.querySelectorAll('a')].filter(a => {
        const parent = a.parentElement;
        if (!parent) return false;
        const siblings = parent.children.length;
        const text = a.textContent.trim();
        return siblings >= 3 && text.length < 30 && text.length > 1 &&
               (a.className.includes('btn') || a.className.includes('card') || a.className.includes('link') ||
                parent.className.includes('list') || parent.className.includes('grid') || parent.className.includes('menu'));
      });

      const categoryCount = Math.max(categoryElements.length, possibleCatLinks.length > 3 ? possibleCatLinks.length : 0);
      const categoryTexts = categoryElements.length > 0
        ? categoryElements.map(c => cleanText(c).substring(0, 40))
        : possibleCatLinks.slice(0, 10).map(a => a.textContent.trim().substring(0, 40));

      // ---- FAQ Items analysis ----
      const faqSelectors = [
        'details', '.faq-item', '[class*="faq"]', '.accordion-item', '[class*="accordion"]',
        '.question', '[class*="question"]', 'dt', '.collapse-item',
        '[class*="toggle"]', '[itemtype*="FAQPage"] [itemprop="mainEntity"]'
      ];
      let faqItems = [];
      for (const sel of faqSelectors) {
        const items = [...document.querySelectorAll(sel)];
        if (items.length > faqItems.length) faqItems = items;
      }

      // If no FAQ-specific selectors, try to find Q&A patterns
      if (faqItems.length === 0) {
        // Look for Q: A: patterns or numbered questions
        const allPs = [...document.querySelectorAll('p, li, div')];
        const qaItems = allPs.filter(el => {
          const text = el.textContent.trim();
          return /^[QＱ][.．:\s：]/.test(text) || /^質問/.test(text);
        });
        if (qaItems.length > 0) faqItems = qaItems;
      }

      const faqCount = faqItems.length;
      const faqTexts = faqItems.slice(0, 10).map(el => cleanText(el).substring(0, 150));

      // ---- Search functionality ----
      const searchInputs = [
        ...document.querySelectorAll('input[type="search"]'),
        ...document.querySelectorAll('input[placeholder*="検索"]'),
        ...document.querySelectorAll('input[placeholder*="search" i]'),
        ...document.querySelectorAll('input[name*="search" i]'),
        ...document.querySelectorAll('input[name*="query" i]'),
        ...document.querySelectorAll('input[name*="keyword" i]'),
        ...document.querySelectorAll('[class*="search"] input'),
        ...document.querySelectorAll('[role="search"] input'),
        ...document.querySelectorAll('input[aria-label*="検索"]')
      ];
      const hasSearch = searchInputs.length > 0;
      const searchPlaceholder = searchInputs[0]?.placeholder || '';

      // Check for natural language / AI search hints
      const hasNLSearch = searchPlaceholder.includes('文章') || searchPlaceholder.includes('質問を入力') ||
                          !!document.querySelector('[class*="ai-search"], [class*="smart-search"], [class*="nlp"]');

      // ---- Category filtering / browsing ----
      const hasCategoryFilter = categoryCount > 0;
      const hasExpandableCategories = document.querySelectorAll(
        '[class*="accordion"] [class*="category"], [class*="collapse"] [class*="category"], details [class*="category"]'
      ).length > 0 || document.querySelectorAll('[class*="tree"], [class*="nested"]').length > 0;

      // ---- Feedback / Rating ----
      const feedbackElements = [
        ...document.querySelectorAll('[class*="feedback"]'),
        ...document.querySelectorAll('[class*="helpful"]'),
        ...document.querySelectorAll('[class*="rating"]'),
        ...document.querySelectorAll('[class*="vote"]'),
        ...document.querySelectorAll('[class*="thumbs"]'),
        ...document.querySelectorAll('[class*="like"]'),
        ...document.querySelectorAll('button[class*="yes"], button[class*="no"]'),
        ...document.querySelectorAll('[class*="solved"]'),
        ...document.querySelectorAll('[aria-label*="役に立"]'),
        ...document.querySelectorAll('[aria-label*="helpful"]')
      ];
      const hasFeedback = feedbackElements.length > 0;

      // Check if feedback has free text input
      const hasFeedbackText = feedbackElements.some(el => {
        const parent = el.closest('form, div, section');
        return parent && parent.querySelector('textarea, input[type="text"]');
      });

      // ---- Attachments / Media ----
      const contentArea = document.querySelector('main, .main, #main, .content, #content, article') || body;
      const images = contentArea.querySelectorAll('img:not(header img):not(nav img):not(footer img)');
      const pdfLinks = [...contentArea.querySelectorAll('a[href$=".pdf"], a[href*=".pdf?"]')];
      const fileLinks = [...contentArea.querySelectorAll('a[href$=".doc"], a[href$=".docx"], a[href$=".xlsx"], a[href$=".xls"], a[href$=".zip"]')];
      const hasImagePreview = images.length > 0 && [...images].some(img => {
        const w = img.naturalWidth || img.width;
        return w > 100; // Meaningful images, not icons
      });
      const hasFileAttachments = pdfLinks.length > 0 || fileLinks.length > 0;

      // ---- Content quality indicators ----
      const allTextContent = cleanText(contentArea);
      const wordCount = allTextContent.length;

      // Check for bullet lists, tables within answers
      const hasBulletLists = contentArea.querySelectorAll('ul li, ol li').length > 3;
      const hasTables = contentArea.querySelectorAll('table').length > 0;
      const hasSteps = /手順|ステップ|STEP|step\s*\d|①|②|③|1\.|2\.|3\./.test(allTextContent);

      // Check for consistent terminology
      const usesDesuMasu = (allTextContent.match(/です。|ます。|ください。|ません。/g) || []).length;
      const usesDaDearu = (allTextContent.match(/である。|だ。|した。|する。/g) || []).length;
      const isConsistentTone = usesDesuMasu > 0 && usesDaDearu === 0 || usesDaDearu > 0 && usesDesuMasu === 0;
      const hasTechnicalJargon = /API|SDK|HTTP|DNS|SSL|TLS|SMTP|IMAP|LDAP|OAuth/.test(allTextContent);

      // Accordion / expandable answers
      const hasAccordion = document.querySelectorAll(
        'details, [class*="accordion"], [class*="collapse"], [class*="toggle"], [class*="expand"]'
      ).length > 0;

      // ---- Breadcrumbs ----
      const hasBreadcrumbs = !!document.querySelector('[class*="breadcrumb"], [aria-label="breadcrumb"], nav[class*="path"]');

      // ---- Mobile responsive ----
      const hasViewportMeta = !!document.querySelector('meta[name="viewport"]');

      return {
        pageTitle,
        metaDesc,
        ogSiteName,
        logoAlt,
        h1Text,
        domain,
        hasLogo,
        hasFavicon,
        hasCustomCSS,
        faqLinksInNavCount: faqLinksInNav.length,
        faqLinksAllCount: faqLinksAll.length,
        contactLinksCount: contactLinks.length,
        hasContactWithFaqCushion,
        faqInH1orH2,
        faqHeadingsCount: faqHeadings.length,
        categoryCount,
        categoryTexts,
        faqCount,
        faqTexts,
        hasSearch,
        hasNLSearch,
        searchPlaceholder,
        hasCategoryFilter,
        hasExpandableCategories,
        hasFeedback,
        hasFeedbackText,
        hasImagePreview,
        hasFileAttachments,
        pdfLinksCount: pdfLinks.length,
        hasBulletLists,
        hasTables,
        hasSteps,
        isConsistentTone,
        hasTechnicalJargon,
        hasAccordion,
        hasBreadcrumbs,
        hasViewportMeta,
        wordCount,
        usesDesuMasu,
        usesDaDearu,
        contentImagesCount: images.length,
        headingsCount: headings.length
      };
    });

    // Also try to check the top page for FAQ link presence
    let topPageHasFaqLink = false;
    try {
      const parsedUrl = new URL(url);
      const topUrl = parsedUrl.origin;
      if (topUrl !== url && !url.endsWith(parsedUrl.origin + '/')) {
        const topPage = await browser.newPage();
        await topPage.setViewport({ width: 1280, height: 800 });
        await topPage.goto(topUrl, { waitUntil: 'networkidle2', timeout: 15000 });
        topPageHasFaqLink = await topPage.evaluate(() => {
          const links = [...document.querySelectorAll('a')];
          return links.some(a => {
            const text = (a.textContent || '').toLowerCase();
            const href = (a.href || '').toLowerCase();
            return text.includes('faq') || text.includes('よくある') || text.includes('質問') ||
                   href.includes('faq') || href.includes('question');
          });
        });
        await topPage.close();
      }
    } catch (e) {
      // Ignore top page check failure
    }

    siteData.topPageHasFaqLink = topPageHasFaqLink;
    return { screenshot, siteData };
  } finally {
    await browser.close();
  }
}

// ============================================================
// Rule-based evaluation engine
// ============================================================
function evaluateSite(d, url) {
  // ---- A. Induction (誘導) ----
  // 1. FAQサイトへの誘導
  let inductionLink = { score: 0, symbol: '×', comment: '' };
  if (d.faqLinksInNavCount > 0 || d.topPageHasFaqLink) {
    if (d.faqLinksInNavCount > 0 && d.topPageHasFaqLink) {
      inductionLink = { score: 2, symbol: '○', comment: 'ナビゲーションおよびトップページからFAQへの導線が確保されています。' };
    } else {
      inductionLink = { score: 1, symbol: '△', comment: 'FAQへのリンクはありますが、より目立つ位置に配置するとアクセス向上が期待できます。' };
    }
  } else {
    inductionLink = { score: 0, symbol: '×', comment: 'FAQページへの誘導リンクを目立つ位置に設置することで、ユーザーの自己解決率向上が見込めます。' };
  }

  // 2. 問合せフォームとの関係
  let contactRelation = { score: 0, symbol: '×', comment: '' };
  if (d.contactLinksCount > 0 && d.hasContactWithFaqCushion) {
    contactRelation = { score: 2, symbol: '○', comment: 'お問い合わせ前にFAQを案内するクッション導線が設けられています。' };
  } else if (d.contactLinksCount > 0) {
    contactRelation = { score: 1, symbol: '△', comment: 'お問い合わせリンクがありますが、FAQ閲覧を促すクッションページを追加するとさらに効果的です。' };
  } else {
    contactRelation = { score: 0, symbol: '×', comment: 'お問い合わせフォームの手前にFAQ案内を設置することで、問い合わせ数の削減が期待できます。' };
  }

  // 3. 注目のFAQ表出
  let faqProminence = { score: 0, symbol: '×', comment: '' };
  if (d.faqInH1orH2 && d.faqHeadingsCount > 0) {
    faqProminence = { score: 2, symbol: '○', comment: '「よくある質問」が見出しとして目立つ位置に表示されており、ユーザーが見つけやすい構成です。' };
  } else if (d.faqHeadingsCount > 0 || d.faqLinksAllCount > 2) {
    faqProminence = { score: 1, symbol: '△', comment: 'FAQ関連の表示はありますが、より大きな見出しやバナーで表出するとさらに効果的です。' };
  } else {
    faqProminence = { score: 0, symbol: '×', comment: '「よくある質問」を目立つ見出しやバナーで表出することで、ユーザーの注目度が高まります。' };
  }

  // 4. ブランドイメージ
  let brandImage = { score: 0, symbol: '×', comment: '' };
  if (d.hasLogo && d.hasFavicon && d.hasCustomCSS) {
    brandImage = { score: 2, symbol: '○', comment: 'ロゴやファビコン、統一されたデザインでブランドイメージが保たれています。' };
  } else if (d.hasCustomCSS && (d.hasLogo || d.hasFavicon)) {
    brandImage = { score: 1, symbol: '△', comment: '基本的なブランド要素は揃っていますが、ロゴやカラーの統一感をさらに強化できます。' };
  } else {
    brandImage = { score: 0, symbol: '×', comment: 'ロゴやブランドカラーを統一的に配置することで、信頼感のあるFAQサイトになります。' };
  }

  // ---- B. Classification (分類) ----
  // 5. カテゴリ分け
  let categoryOrg = { score: 0, symbol: '×', comment: '' };
  if (d.categoryCount >= 4) {
    categoryOrg = { score: 2, symbol: '○', comment: `${d.categoryCount}件のカテゴリが設けられ、ユーザー視点で整理されています。` };
  } else if (d.categoryCount >= 2) {
    categoryOrg = { score: 1, symbol: '△', comment: 'カテゴリ分けはありますが、より細かく整理するとユーザーが探しやすくなります。' };
  } else {
    categoryOrg = { score: 0, symbol: '×', comment: 'カテゴリ分けを導入することで、ユーザーが目的のFAQにたどり着きやすくなります。' };
  }

  // 6. カテゴリ名称・並び順
  let categoryNaming = { score: 0, symbol: '×', comment: '' };
  if (d.categoryCount >= 4 && d.categoryTexts.length > 0) {
    const hasShortNames = d.categoryTexts.every(t => t.length <= 20);
    if (hasShortNames) {
      categoryNaming = { score: 2, symbol: '○', comment: 'カテゴリ名が簡潔で分かりやすく、直感的に選べる構成です。' };
    } else {
      categoryNaming = { score: 1, symbol: '△', comment: 'カテゴリ名をより簡潔にすると、ユーザーが直感的に選びやすくなります。' };
    }
  } else if (d.categoryCount >= 2) {
    categoryNaming = { score: 1, symbol: '△', comment: 'カテゴリはありますが、名称や並び順を工夫するとさらに使いやすくなります。' };
  } else {
    categoryNaming = { score: 0, symbol: '×', comment: 'ユーザーが直感的に理解できるカテゴリ名と論理的な並び順を検討しましょう。' };
  }

  // ---- C. Content (表記) ----
  // 7. 表記の統一
  let toneConsistency = { score: 0, symbol: '×', comment: '' };
  if (d.isConsistentTone && d.usesDesuMasu > 5) {
    toneConsistency = { score: 2, symbol: '○', comment: '「です・ます」調で統一されており、読みやすい文章です。' };
  } else if (d.usesDesuMasu > 0 || d.usesDaDearu > 0) {
    toneConsistency = { score: 1, symbol: '△', comment: '文体はおおむね統一されていますが、一部混在が見られます。統一するとさらに読みやすくなります。' };
  } else {
    toneConsistency = { score: 1, symbol: '△', comment: '文章量が少ないため十分な判定ができませんが、「です・ます」調での統一をお勧めします。' };
  }

  // 8. 分かりやすい言葉
  let clearLanguage = { score: 0, symbol: '×', comment: '' };
  if (!d.hasTechnicalJargon && d.wordCount > 100) {
    clearLanguage = { score: 2, symbol: '○', comment: '専門用語を避け、一般ユーザーにも分かりやすい表現で記述されています。' };
  } else if (d.hasTechnicalJargon) {
    clearLanguage = { score: 1, symbol: '△', comment: '一部専門用語が使用されています。注釈や補足を加えると、より幅広いユーザーに理解しやすくなります。' };
  } else {
    clearLanguage = { score: 1, symbol: '△', comment: '基本的な表現は分かりやすいですが、用語の補足説明を追加するとさらに親切です。' };
  }

  // 9. 解決できる記述
  let actionableContent = { score: 0, symbol: '×', comment: '' };
  if (d.hasSteps && d.faqCount > 0) {
    actionableContent = { score: 2, symbol: '○', comment: '手順やステップが明示されており、ユーザーが具体的に行動できる記述です。' };
  } else if (d.faqCount > 3 && d.wordCount > 500) {
    actionableContent = { score: 1, symbol: '△', comment: '回答内容はありますが、結論を先に示し具体的なアクションを明記するとさらに効果的です。' };
  } else {
    actionableContent = { score: 0, symbol: '×', comment: '回答に「結論」と「具体的な手順」を明記すると、ユーザーの自己解決率が大きく向上します。' };
  }

  // 10. 見やすいレイアウト
  let readableLayout = { score: 0, symbol: '×', comment: '' };
  if (d.hasBulletLists && (d.hasTables || d.contentImagesCount > 2) && d.hasAccordion) {
    readableLayout = { score: 2, symbol: '○', comment: '箇条書き、表、画像が効果的に活用され、視覚的に見やすいレイアウトです。' };
  } else if (d.hasBulletLists || d.hasAccordion || d.contentImagesCount > 1) {
    readableLayout = { score: 1, symbol: '△', comment: '基本的なレイアウトは整っていますが、箇条書きや図解をさらに活用すると読みやすさが向上します。' };
  } else {
    readableLayout = { score: 0, symbol: '×', comment: '箇条書き、画像、表を活用したレイアウトにすると、情報が格段に見やすくなります。' };
  }

  // ---- D. Functionality (機能) ----
  // 11. ワード検索機能
  let searchFunc = { score: 0, symbol: '×', comment: '' };
  if (d.hasNLSearch) {
    searchFunc = { score: 2, symbol: '○', comment: '自然文対応の高精度な検索機能が搭載されています。' };
  } else if (d.hasSearch) {
    searchFunc = { score: 1, symbol: '△', comment: 'キーワード検索機能があります。自然文検索やAI検索を導入するとさらに利便性が向上します。' };
  } else {
    searchFunc = { score: 0, symbol: '×', comment: '検索ボックスを設置することで、ユーザーが素早く目的のFAQにたどり着けるようになります。' };
  }

  // 12. カテゴリ検索機能
  let categorySearch = { score: 0, symbol: '×', comment: '' };
  if (d.hasCategoryFilter && d.hasExpandableCategories) {
    categorySearch = { score: 2, symbol: '○', comment: 'カテゴリによる絞り込みが見やすく整理されています。' };
  } else if (d.hasCategoryFilter) {
    categorySearch = { score: 1, symbol: '△', comment: 'カテゴリ一覧はありますが、絞り込みや展開表示を追加するとさらに使いやすくなります。' };
  } else {
    categorySearch = { score: 0, symbol: '×', comment: 'カテゴリ別の絞り込み機能を追加すると、FAQ数が増えても快適に検索できます。' };
  }

  // 13. 評価分析機能
  let feedbackFunc = { score: 0, symbol: '×', comment: '' };
  if (d.hasFeedback && d.hasFeedbackText) {
    feedbackFunc = { score: 2, symbol: '○', comment: '自由入力欄を含む評価機能があり、ユーザーの声を収集してFAQ改善に活かせます。' };
  } else if (d.hasFeedback) {
    feedbackFunc = { score: 1, symbol: '△', comment: 'Yes/No形式の評価ボタンがあります。自由入力欄を追加すると改善のヒントが得られます。' };
  } else {
    feedbackFunc = { score: 0, symbol: '×', comment: '「この回答は役に立ちましたか？」などの評価機能を追加すると、FAQ品質の継続的な改善が可能になります。' };
  }

  // 14. 添付ファイル機能
  let attachmentFunc = { score: 0, symbol: '×', comment: '' };
  if (d.hasImagePreview && d.hasFileAttachments) {
    attachmentFunc = { score: 2, symbol: '○', comment: '画像やPDFファイルが活用されており、視覚的にも分かりやすい回答が提供されています。' };
  } else if (d.hasImagePreview || d.hasFileAttachments) {
    attachmentFunc = { score: 1, symbol: '△', comment: '一部ファイルや画像が使用されていますが、図解やマニュアルPDFをさらに充実させると効果的です。' };
  } else {
    attachmentFunc = { score: 0, symbol: '×', comment: '画像や関連PDFを添付すると、文字だけでは伝わりにくい内容も分かりやすく説明できます。' };
  }

  // ---- Build scores ----
  const scores = {
    induction: {
      items: [
        { name: 'FAQサイトへの誘導', ...inductionLink },
        { name: '問合せフォームとの関係', ...contactRelation },
        { name: '注目のFAQ表出', ...faqProminence },
        { name: 'ブランドイメージ', ...brandImage }
      ],
      subtotal: inductionLink.score + contactRelation.score + faqProminence.score + brandImage.score
    },
    classification: {
      items: [
        { name: 'カテゴリ分け', ...categoryOrg },
        { name: 'カテゴリ名称・並び順', ...categoryNaming }
      ],
      subtotal: categoryOrg.score + categoryNaming.score
    },
    content: {
      items: [
        { name: '表記の統一', ...toneConsistency },
        { name: '分かりやすい言葉', ...clearLanguage },
        { name: '解決できる記述', ...actionableContent },
        { name: '見やすいレイアウト', ...readableLayout }
      ],
      subtotal: toneConsistency.score + clearLanguage.score + actionableContent.score + readableLayout.score
    },
    functionality: {
      items: [
        { name: 'ワード検索機能', ...searchFunc },
        { name: 'カテゴリ検索機能', ...categorySearch },
        { name: '評価分析機能', ...feedbackFunc },
        { name: '添付ファイル機能', ...attachmentFunc }
      ],
      subtotal: searchFunc.score + categorySearch.score + feedbackFunc.score + attachmentFunc.score
    }
  };

  const rawTotal = scores.induction.subtotal + scores.classification.subtotal +
                   scores.content.subtotal + scores.functionality.subtotal;
  // Convert to 100-point scale: min 45, max 100 (generous baseline so scores feel positive)
  const totalScore = Math.min(100, Math.round(45 + (rawTotal / 28) * 55));

  // ---- Determine diagnostic type ----
  // Normalize to percentage for comparison
  const catPct = {
    '誘導': scores.induction.subtotal / 8,
    '分類': scores.classification.subtotal / 4,
    '表記': scores.content.subtotal / 8,
    '機能': scores.functionality.subtotal / 8
  };

  let diagnosticType, diagnosticTypeDescription;
  if (totalScore >= 85) {
    diagnosticType = '優良FAQサイト（エキスパート）';
    diagnosticTypeDescription = '高い完成度のFAQサイトです。さらなる磨き上げで業界トップを目指しましょう。';
  } else {
    const lowestCat = Object.entries(catPct).sort((a, b) => a[1] - b[1])[0][0];
    const typeMap = {
      '誘導': { type: '見せる工夫でさらに輝くタイプ', desc: 'FAQへの誘導を強化して、もっと多くのユーザーに届けましょう。' },
      '分類': { type: 'UI整理で使いやすさ倍増タイプ', desc: 'カテゴリ整理で探しやすさを大幅に向上させましょう。' },
      '表記': { type: 'コンテンツを磨いてファン獲得タイプ', desc: 'コンテンツの質を高めてユーザー満足度を向上させましょう。' },
      '機能': { type: '機能を導入して効率化タイプ', desc: 'ASPサービスの機能を活用して運用全体を改革してみませんか？' }
    };
    diagnosticType = typeMap[lowestCat].type;
    diagnosticTypeDescription = typeMap[lowestCat].desc;
  }

  // ---- Company name ----
  const companyName = d.ogSiteName || d.logoAlt || extractCompanyName(d.pageTitle, d.h1Text, d.domain);

  // ---- Summary ----
  const goodPoints = [];
  if (scores.induction.subtotal >= 3) goodPoints.push('FAQへの誘導');
  if (scores.classification.subtotal >= 2) goodPoints.push('カテゴリ整理');
  if (scores.content.subtotal >= 3) goodPoints.push('コンテンツの質');
  if (scores.functionality.subtotal >= 2) goodPoints.push('機能面の充実');
  if (d.faqCount > 0) goodPoints.push(`${d.faqCount}件のFAQ`);

  const goodStr = goodPoints.length > 0
    ? `${goodPoints.join('、')}など、しっかりとした基盤が構築されています。`
    : 'FAQサイトとしての基盤が整っています。';

  const summary = `${goodStr}コンテンツの分類については、${d.categoryCount > 0 ? 'カテゴリによる分類がされており、' : ''}分類内容の充実が期待できます。${diagnosticTypeDescription}`;

  // ---- Proposals ----
  const proposals = generateProposals(scores, d);

  // ---- One-point advice ----
  const onePointAdvice = generateOnePointAdvice(scores, d);

  return {
    companyName,
    diagnosticType,
    diagnosticTypeDescription,
    summary,
    scores,
    totalScore,
    proposals,
    onePointAdvice
  };
}

function extractCompanyName(title, h1, domain) {
  // Try to extract from title (often "FAQ | CompanyName" or "CompanyName - FAQ")
  if (title) {
    const parts = title.split(/[|\-–—/／]/).map(s => s.trim());
    const nonFaq = parts.filter(p => !p.toLowerCase().includes('faq') && !p.includes('よくある') && !p.includes('質問'));
    if (nonFaq.length > 0) return nonFaq[0].substring(0, 30);
  }
  if (h1 && !h1.toLowerCase().includes('faq') && !h1.includes('よくある')) {
    return h1.substring(0, 30);
  }
  return domain;
}

function generateProposals(scores, d) {
  const proposals = [];

  // Find weakest areas and generate proposals
  const allItems = [
    ...scores.induction.items.map(i => ({ ...i, category: 'induction' })),
    ...scores.classification.items.map(i => ({ ...i, category: 'classification' })),
    ...scores.content.items.map(i => ({ ...i, category: 'content' })),
    ...scores.functionality.items.map(i => ({ ...i, category: 'functionality' }))
  ].filter(i => i.score < 2).sort((a, b) => a.score - b.score);

  const proposalTemplates = {
    'ワード検索機能': {
      title: '検索機能の導入で探しやすさを向上',
      current: '現在はFAQ一覧からの目視検索が中心で、ユーザーが「キーワード検索」で素早く回答にたどり着く手段がありません。',
      suggestion: 'FAQページ上部に検索ボックスを設置し、キーワードやタグによる絞り込み検索を導入すると、ユーザーの探索時間が大幅に短縮されます。'
    },
    'カテゴリ検索機能': {
      title: 'カテゴリ絞り込みで目的のFAQへ直行',
      current: 'FAQ一覧は表示されていますが、カテゴリによる絞り込み機能があるとさらに便利です。',
      suggestion: 'カテゴリボタンやタブによる絞り込みUIを追加し、FAQが増えても快適に検索できる構造にしましょう。'
    },
    '評価分析機能': {
      title: 'ユーザー評価機能でFAQ品質を継続改善',
      current: 'FAQ回答に対するユーザーの満足度を把握する仕組みがまだありません。',
      suggestion: '各FAQ回答の下に「この回答は役に立ちましたか？」ボタンを設置し、改善すべきFAQを特定できるようにしましょう。'
    },
    '添付ファイル機能': {
      title: '画像・資料の活用で理解度アップ',
      current: '回答は文章が中心で、画像やPDFなどの補足資料が少ない状態です。',
      suggestion: '手順説明にはスクリーンショットを、詳細な説明にはPDFマニュアルのリンクを添付すると理解が速まります。'
    },
    'FAQサイトへの誘導': {
      title: 'FAQページへの誘導強化でアクセス増加',
      current: 'FAQページへの導線がもう少し目立つ位置にあると、より多くのユーザーに活用されます。',
      suggestion: 'トップページのヘッダーやグローバルナビにFAQリンクを設置し、フッターにも配置すると発見率が向上します。'
    },
    '問合せフォームとの関係': {
      title: '問い合わせ前のFAQ案内で自己解決を促進',
      current: 'お問い合わせフォームとFAQの連携がさらに強化できる余地があります。',
      suggestion: 'お問い合わせフォームの手前に「まずはFAQをご覧ください」というクッションページを設置すると、問い合わせ数の削減につながります。'
    },
    '注目のFAQ表出': {
      title: 'よくある質問の目立つ表出で利用率向上',
      current: 'FAQコンテンツはありますが、「よくある質問」としての目立つ表出を強化できます。',
      suggestion: '大きな見出しやバナーで「よくある質問」を表出し、人気のFAQをトップに表示するとユーザーの利用率が上がります。'
    },
    'カテゴリ分け': {
      title: 'カテゴリ整理で情報の見つけやすさ向上',
      current: 'FAQがフラットに並んでおり、カテゴリ分けを導入すると探しやすさが向上します。',
      suggestion: 'ユーザーの利用シーン別にカテゴリを設定し、各カテゴリにアイコンを付けると直感的に選べる構成になります。'
    },
    '見やすいレイアウト': {
      title: 'ビジュアルレイアウトの強化で読みやすさ改善',
      current: '回答が文章中心で、もう少し視覚的な工夫があると読みやすくなります。',
      suggestion: '箇条書き、番号付きリスト、画像、表を活用し、重要なポイントは太字やハイライトで強調すると情報が伝わりやすくなります。'
    },
    '解決できる記述': {
      title: 'アクション明示で自己解決率を向上',
      current: '回答内容はありますが、結論やアクションをより明確にできる余地があります。',
      suggestion: '各回答の冒頭に結論を記載し、「手順1→手順2→手順3」のようにアクションを明確にすると自己解決率が向上します。'
    },
    '表記の統一': {
      title: '文体・用語の統一で信頼感を向上',
      current: '文体に一部ばらつきが見られます。',
      suggestion: '「です・ます」調に統一し、用語集を作成して一貫した表現にすると、プロフェッショナルな印象が強まります。'
    },
    '分かりやすい言葉': {
      title: '平易な表現で幅広いユーザーに対応',
      current: '一部専門的な表現が使われており、初めてのユーザーには分かりにくい場合があります。',
      suggestion: '専門用語には括弧書きで補足説明を加え、初めてのユーザーでも理解できる表現を心がけましょう。'
    },
    'ブランドイメージ': {
      title: 'デザイン統一でブランド価値を向上',
      current: 'FAQページのデザインがサイト全体のブランドイメージとさらに統一できる余地があります。',
      suggestion: 'サイト全体と同じカラースキーム、ロゴ、フォントをFAQページにも適用し、一貫したブランド体験を提供しましょう。'
    },
    'カテゴリ名称・並び順': {
      title: 'カテゴリ名と並び順の最適化',
      current: 'カテゴリの名称や並び順をさらに工夫する余地があります。',
      suggestion: '利用頻度の高いカテゴリを上位に配置し、ユーザーが直感的に理解できる簡潔なカテゴリ名にしましょう。'
    }
  };

  for (const item of allItems) {
    if (proposals.length >= 3) break;
    const template = proposalTemplates[item.name];
    if (template && !proposals.find(p => p.title === template.title)) {
      proposals.push({
        title: template.title,
        priority: proposals.length === 0 ? '高' : '中',
        current: template.current,
        suggestion: template.suggestion
      });
    }
  }

  // Fill remaining slots if needed
  while (proposals.length < 3) {
    const remaining = Object.values(proposalTemplates).find(t => !proposals.find(p => p.title === t.title));
    if (!remaining) break;
    proposals.push({ ...remaining, priority: '中' });
  }

  return proposals;
}

function generateOnePointAdvice(scores, d) {
  // Pick the most impactful quick win
  if (!d.hasSearch) {
    return {
      focus: '検索ボックスの設置',
      advice: 'FAQページのトップに検索ボックスを1つ設置するだけで、ユーザーが目的のFAQに素早くたどり着けるようになります。プレースホルダーに「例：パスワードの変更方法」と入れると、何を入力すべきか迷わなくなります。'
    };
  }

  if (!d.hasFeedback) {
    return {
      focus: '評価ボタンの追加',
      advice: '各FAQ回答の下に「役に立った / 役に立たなかった」ボタンを設置するだけで、改善すべきFAQが一目で分かるようになります。わずかなHTML追加で実現可能です。'
    };
  }

  if (d.categoryCount > 0 && d.categoryCount < 4) {
    return {
      focus: 'カテゴリの充実',
      advice: `現在${d.categoryCount}カテゴリですが、ユーザーの利用シーン別に5〜8カテゴリに整理すると、FAQが増えても探しやすさを維持できます。カテゴリ名にアイコンを添えるとさらに直感的です。`
    };
  }

  if (d.categoryCount >= 4) {
    return {
      focus: 'カテゴリリンクの活用',
      advice: 'ページのトップにカテゴリが表示されているのはアクセスや探しやすさの向上に有効です。各カテゴリの横にFAQ件数を表示すると、ユーザーがボリューム感を把握しやすくなります。'
    };
  }

  return {
    focus: 'FAQの視認性向上',
    advice: 'ページ上部に「よくある質問TOP5」を目立つデザインで表示するだけで、多くのユーザーが一番知りたい情報にすぐアクセスできるようになります。'
  };
}

// ============================================================
// PPTX Generation
// ============================================================
async function generatePptx(data, radarChartImage) {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: 13.33, height: 7.5 });
  pptx.layout = 'WIDE';

  const slide = pptx.addSlide();
  slide.background = { fill: 'FFFFFF' };

  // Header badge
  slide.addShape(pptx.ShapeType.rect, { x: 0.1, y: 0.1, w: 2.5, h: 0.35, fill: { color: '2980B9' }, rectRadius: 0.03 });
  slide.addText('【FAQサイト診断レポート】', {
    x: 0.1, y: 0.1, w: 2.5, h: 0.35,
    fontSize: 9, color: 'FFFFFF', bold: true, align: 'center', valign: 'middle', fontFace: 'Meiryo'
  });

  // Company name
  slide.addText([
    { text: data.companyName || '', options: { fontSize: 28, bold: true, color: '1A3A5C' } },
    { text: ' 様', options: { fontSize: 20, color: '666666' } }
  ], { x: 0.5, y: 0.35, w: 8, h: 0.7, fontFace: 'Meiryo' });

  // Date
  const today = new Date();
  const dateStr = `${today.getFullYear()}/${String(today.getMonth() + 1).padStart(2, '0')}`;
  slide.addText(`レポート作成日：${dateStr}`, {
    x: 7.5, y: 0.3, w: 2.5, h: 0.25, fontSize: 7, color: '666666', fontFace: 'Meiryo'
  });

  // Score
  slide.addText(`合計: ${data.totalScore} / 100点`, {
    x: 0.1, y: 0.6, w: 2.5, h: 0.35, fontSize: 12, color: '2980B9', bold: true, fontFace: 'Meiryo'
  });

  // Separator lines
  slide.addShape(pptx.ShapeType.rect, { x: 0.1, y: 1.05, w: 4.2, h: 0.07, fill: { color: 'CCCCCC' } });
  slide.addShape(pptx.ShapeType.rect, { x: 4.55, y: 1.05, w: 8.5, h: 0.07, fill: { color: 'CCCCCC' } });

  // Left: Radar chart panel
  slide.addShape(pptx.ShapeType.rect, { x: 0.1, y: 1.12, w: 4.2, h: 3.0, fill: { color: 'F0F4F8' }, line: { color: 'DCE4EC', width: 1 }, rectRadius: 0.05 });
  slide.addText('【全体評価項目】', {
    x: 0.1, y: 1.12, w: 4.2, h: 0.35, fontSize: 10, color: '1A3A5C', bold: true, align: 'center', fontFace: 'Meiryo'
  });
  slide.addText('FAQサイトへの辿り着きやすさ～表記内容～機能面までの\n総合評価点です。', {
    x: 0.1, y: 1.4, w: 4.2, h: 0.4, fontSize: 8, color: '7F8C8D', align: 'center', fontFace: 'Meiryo'
  });

  if (radarChartImage) {
    const imgData = radarChartImage.replace(/^data:image\/\w+;base64,/, '');
    slide.addImage({ data: `image/png;base64,${imgData}`, x: 0.5, y: 1.7, w: 3.5, h: 2.3 });
  }

  // Left: One-point advice with screenshot
  slide.addShape(pptx.ShapeType.rect, { x: 0.1, y: 4.2, w: 4.2, h: 0.3, fill: { color: '2980B9' }, rectRadius: 0.05 });
  slide.addText('ワンポイントアドバイス', { x: 0.2, y: 4.2, w: 2.5, h: 0.3, fontSize: 9, color: 'FFFFFF', bold: true, fontFace: 'Meiryo' });
  slide.addShape(pptx.ShapeType.rect, { x: 0.1, y: 4.5, w: 4.2, h: 2.6, fill: { color: 'F0F7FC' }, line: { color: 'D0E0F0', width: 1 }, rectRadius: 0.05 });

  const adviceText = data.onePointAdvice ? `${data.onePointAdvice.focus}\n\n${data.onePointAdvice.advice}` : '';
  slide.addText(adviceText, {
    x: 0.2, y: 4.55, w: 4.0, h: 1.2, fontSize: 9, color: '2C3E50', fontFace: 'Meiryo', valign: 'top', wrap: true, bold: false
  });

  if (data.screenshot) {
    slide.addImage({
      data: `image/png;base64,${data.screenshot}`, x: 0.2, y: 5.75, w: 3.9, h: 1.25, rounding: true
    });
  }

  // Right: Diagnostic type
  slide.addShape(pptx.ShapeType.rect, { x: 4.55, y: 0.8, w: 8.5, h: 0.35, fill: { color: 'FEF9F0' }, line: { color: 'F0E0C8', width: 1 } });
  slide.addText('【診断タイプ】', { x: 4.6, y: 0.8, w: 1.2, h: 0.35, fontSize: 9, color: '888888', fontFace: 'Meiryo' });
  slide.addText(data.diagnosticType ? `【${data.diagnosticType}】` : '', {
    x: 5.7, y: 0.8, w: 7.0, h: 0.35, fontSize: 14, color: 'C0392B', bold: true, fontFace: 'Meiryo'
  });

  // Right: Main content
  slide.addShape(pptx.ShapeType.rect, { x: 4.55, y: 1.2, w: 8.5, h: 5.9, fill: { color: 'F8FAFC' }, line: { color: 'DCE4EC', width: 1 }, rectRadius: 0.05 });

  let mainText = '';
  if (data.summary) mainText += `今回は「${data.url}」サイトを対象に調査しています。\n${data.summary}\n\n`;
  mainText += '【改善提案】\n';
  if (data.proposals) {
    data.proposals.forEach((p, i) => {
      mainText += `${i + 1}) ${p.title}\n${p.current}\n→ ${p.suggestion}\n\n`;
    });
  }

  slide.addText(mainText, {
    x: 4.7, y: 1.3, w: 8.2, h: 5.6, fontSize: 9.5, color: '2C3E50', fontFace: 'Meiryo',
    valign: 'top', wrap: true, lineSpacingMultiple: 1.4
  });

  // Footer
  slide.addText('※表示されている外部サイトの診断のため、内部の運用体制・ルールまでの診断はしておりません。', {
    x: 0.1, y: 7.15, w: 12.5, h: 0.25, fontSize: 7, color: 'AAAAAA', fontFace: 'Meiryo', align: 'center'
  });

  const arrayBuffer = await pptx.write({ outputType: 'arraybuffer' });
  return arrayBuffer;
}

// ============================================================
// Start server
// ============================================================
app.listen(PORT, () => {
  console.log(`FAQサイト診断レポートツール起動中: http://localhost:${PORT}`);
});
