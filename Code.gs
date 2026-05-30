/**
 * Google Sheets Add-on styled tool for Gantaisaku Zukan (日本がん対策図鑑)
 * File: Code.gs
 * Author: Antigravity Code Assistant
 * Repository: awakuma1113/clinicaltrials
 * 
 * Version 3.0: Added advanced clinical trial results extraction, Results_raw and Results sheet management.
 */

/**
 * Google Sheets が開かれたときに実行されるイベントハンドラ
 * カスタムメニューを追加します。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('がん対策図鑑取り込み')
    .addItem('サイドバーを表示', 'showSidebar')
    .addToUi();
}

/**
 * サイドバーを表示する関数
 */
function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('がん対策図鑑 臨床試験抽出')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 指定されたURLのHTMLを取得し、臨床試験情報を抽出するコントローラー関数
 * 
 * WordPress REST API 経由での取得を第一選択とし、
 * 万が一失敗した場合は通常のHTMLスクレイピングへフォールバックします。
 * 
 * @param {string} url 対象の記事URL
 * @return {Object} 抽出結果オブジェクト
 */
function fetchAndExtract(url) {
  if (!url || !url.startsWith('http')) {
    throw new Error('無効なURLです。httpまたはhttpsから始まるURLを入力してください。');
  }

  try {
    // URLからスラッグ（固有識別キー）を抽出 (例: https://gantaisaku.net/optitrop-lung05/ -> optitrop-lung05)
    let cleanUrl = url.split('?')[0].split('#')[0];
    if (cleanUrl.endsWith('/')) {
      cleanUrl = cleanUrl.slice(0, -1);
    }
    const slugMatch = cleanUrl.match(/\/([^\/]+)$/);
    const slug = slugMatch ? slugMatch[1] : '';

    let html = '';
    let apiTitle = '';
    let apiPublishDate = '';
    let isWpApiSuccess = false;

    // 1. WordPress REST API 経由での取得を試みる (最も安全かつCloudflareを回避しやすい)
    if (slug && url.includes('gantaisaku.net')) {
      try {
        Logger.log('WordPress REST API経由でのデータ取得を試みます。Slug: ' + slug);
        
        // posts (通常の記事) の検索
        const apiPostUrl = 'https://gantaisaku.net/wp-json/wp/v2/posts?slug=' + slug;
        const apiResponse = UrlFetchApp.fetch(apiPostUrl, { 
          muteHttpExceptions: true,
          headers: { 'User-Agent': 'Mozilla/5.0' }
        });
        
        if (apiResponse.getResponseCode() === 200) {
          const posts = JSON.parse(apiResponse.getContentText());
          if (posts && posts.length > 0) {
            const post = posts[0];
            apiTitle = post.title ? post.title.rendered : '';
            html = post.content ? post.content.rendered : '';
            if (post.date) {
              const datePart = post.date.split('T')[0];
              apiPublishDate = datePart.replace(/-/g, '/'); // YYYY-MM-DD -> YYYY/MM/DD
            }
            isWpApiSuccess = true;
            Logger.log('WordPress REST API (posts) での取得に成功しました。ID: ' + post.id);
          }
        }
        
        // postsにない場合は pages (固定ページ) からの取得を試みる
        if (!isWpApiSuccess) {
          const apiPageUrl = 'https://gantaisaku.net/wp-json/wp/v2/pages?slug=' + slug;
          const apiResponsePage = UrlFetchApp.fetch(apiPageUrl, { 
            muteHttpExceptions: true,
            headers: { 'User-Agent': 'Mozilla/5.0' }
          });
          if (apiResponsePage.getResponseCode() === 200) {
            const pages = JSON.parse(apiResponsePage.getContentText());
            if (pages && pages.length > 0) {
              const page = pages[0];
              apiTitle = page.title ? page.title.rendered : '';
              html = page.content ? page.content.rendered : '';
              if (page.date) {
                const datePart = page.date.split('T')[0];
                apiPublishDate = datePart.replace(/-/g, '/');
              }
              isWpApiSuccess = true;
              Logger.log('WordPress REST API (pages) での取得に成功しました。ID: ' + page.id);
            }
          }
        }
      } catch (apiError) {
        Logger.log('WordPress REST APIでの取得中にエラーが発生しました（HTMLフェッチに移行します）: ' + apiError.message);
      }
    }

    // 2. REST APIで取得できなかった場合、通常のHTMLスクレイピングを行う (フォールバック)
    if (!isWpApiSuccess || !html) {
      Logger.log('通常のHTMLフェッチを実行します。');
      const response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
      });

      if (response.getResponseCode() !== 200) {
        throw new Error('ページの取得に失敗しました。ステータスコード: ' + response.getResponseCode());
      }

      html = response.getContentText('UTF-8');
    }

    // 3. データ抽出の実行
    return extractClinicalTrialData(html, url, apiTitle, apiPublishDate);

  } catch (e) {
    Logger.log('エラーが発生しました: ' + e.message);
    throw new Error('解析中にエラーが発生しました: ' + e.message);
  }
}

/**
 * 将来の拡張用：カテゴリURLから記事URLの一覧を抽出するスタブ関数
 */
function extractArticleLinksFromCategory(categoryUrl) {
  try {
    const response = UrlFetchApp.fetch(categoryUrl);
    const html = response.getContentText('UTF-8');
    
    const links = [];
    const linkRegex = /href="([^"]*?gantaisaku\.net\/(?:trial|regimen|news)\/[^"]+?)"/gi;
    let match;
    while ((match = linkRegex.exec(html)) !== null) {
      const url = match[1];
      if (!links.includes(url)) {
        links.push(url);
      }
    }
    return links;
  } catch (e) {
    Logger.log('カテゴリ解析エラー: ' + e.message);
    return [];
  }
}

/**
 * 将来の拡張用：複数のURLを一括処理するコントローラー関数
 */
function bulkProcessArticles(urls) {
  const results = [];
  for (const url of urls) {
    try {
      const data = fetchAndExtract(url);
      results.push(data);
      Utilities.sleep(1000); // サーバー負荷軽減のためのウェイト
    } catch (e) {
      results.push({
        article_url: url,
        error: e.message
      });
    }
  }
  return results;
}

/**
 * HTMLから臨床試験情報を解析・抽出するメイン処理
 * 
 * @param {string} html 取得したHTMLソース（またはWordPress APIのcontent部分）
 * @param {string} url 記事のURL
 * @param {string} [apiTitle] APIから取得したクリーンなタイトル
 * @param {string} [apiPublishDate] APIから取得した公開日
 * @return {Object} 抽出結果オブジェクト
 */
function extractClinicalTrialData(html, url, apiTitle, apiPublishDate) {
  // HTMLタグを剥がしたクリーンな平文も用意（テキストマイニング用）
  const bodyText = cleanHtml(html);
  
  // 各項目の抽出処理
  const title = apiTitle ? cleanHtml(apiTitle) : extractTitle(html);
  const publishDate = apiPublishDate || extractPublishDate(html);
  const cancerType = extractCancerType(title, bodyText);
  const treatmentSetting = extractTreatmentSetting(title, bodyText);
  const endpoint = extractEndpoint(bodyText);
  const endpointFromTitle = extractEndpointFromTitle(title);
  const trialName = extractTrialName(title, bodyText, url);
  const phase = extractPhase(title, bodyText);
  const nctId = extractNctId(bodyText);
  const country = extractCountry(bodyText);
  const publication = extractPublication(bodyText);
  const pubmedId = extractPubmedId(bodyText);
  const originalArticleUrl = extractOriginalArticleUrl(html);
  const qualitativeSummary = extractQualitativeSummary(bodyText);

  return {
    article_url: url,
    article_title: title,
    publish_date: publishDate,
    publication_date: publishDate, // Map both for backwards compatibility
    cancer_type: cancerType,
    treatment_setting: treatmentSetting,
    endpoint: endpoint,
    endpoint_from_title: endpointFromTitle,
    trial_name: trialName,
    phase: phase,
    nct_id: nctId,
    country: country,
    publication: publication,
    pubmed_id: pubmedId,
    original_article_url: originalArticleUrl,
    qualitative_summary: qualitativeSummary,
    summary: qualitativeSummary, // Keep summary mapping for existing structures
    source_checked: '未確認', // 初期値
    notes: publishDate ? '発表日: ' + publishDate : '',
    raw_html: html // 臨床試験結果抽出用に元のHTMLを保持
  };
}

// ==========================================
// 抽出ヘルパー関数群 (Modular Extraction Helpers)
// ==========================================

/**
 * HTMLタグを除去し、テキストコンテンツのみにする
 */
function cleanHtml(html) {
  if (!html) return '';
  
  // スクリプトとスタイルを除去
  let text = html
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<!--[\s\S]*?-->/g, ' ');
  
  // タグを空白に置換
  text = text.replace(/<[^>]*>/g, ' ');
  
  // HTML実体参照・エンティティのデコード
  text = text
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#039;/g, "'")
    .replace(/&#8211;/g, '–')
    .replace(/&#8217;/g, "'")
    .replace(/&#8218;/g, ',')
    .replace(/&#8220;/g, '"')
    .replace(/&#8221;/g, '"')
    .replace(/&#8230;/g, '...')
    .replace(/&#38;/g, '&');
  
  // 連続する空白や改行をクリーンアップ
  return text.replace(/\s+/g, ' ').trim();
}

/**
 * 記事タイトルの抽出 (HTML scraping用フォールバック)
 */
function extractTitle(html) {
  const h1Match = html.match(/<h1[^>]*class="[^"]*entry-title[^"]*"[^>]*>([\s\S]*?)<\/h1>/i) ||
                  html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
  if (h1Match) {
    return cleanHtml(h1Match[1]);
  }
  
  const titleMatch = html.match(/<title>([\s\S]*?)<\/title>/i);
  if (titleMatch) {
    let t = cleanHtml(titleMatch[1]);
    return t.split(/[-|]/)[0].trim();
  }
  return '';
}

/**
 * 記事の発表日（日付）の抽出 (HTML scraping用フォールバック)
 */
function extractPublishDate(html) {
  // 1. timeタグ
  const timeMatch = html.match(/<time[^>]*datetime="([^"]+)"/i) || 
                    html.match(/class="[^"]*(?:date|published|time)[^"]*"[^>]*>([\s\S]*?)<\/time>/i);
  if (timeMatch) {
    const rawDate = timeMatch[1].trim();
    const formatted = rawDate.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (formatted) {
      return formatted[1] + '/' + ('0' + formatted[2]).slice(-2) + '/' + ('0' + formatted[3]).slice(-2);
    }
  }
  
  // 2. 本文内の日付パターン検索
  const dateRegex = /(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日|(\d{4})[-/](\d{1,2})[-/](\d{1,2})/;
  const match = html.match(dateRegex);
  if (match) {
    if (match[1]) {
      return match[1] + '/' + ('0' + match[2]).slice(-2) + '/' + ('0' + match[3]).slice(-2);
    } else if (match[4]) {
      return match[4] + '/' + ('0' + match[5]).slice(-2) + '/' + ('0' + match[6]).slice(-2);
    }
  }
  return '';
}

/**
 * がん種またはカテゴリの抽出
 */
function extractCancerType(title, bodyText) {
  const textToSearch = (title + ' ' + bodyText).toLowerCase();
  
  const cancerDict = [
    { key: '非小細胞肺がん', terms: ['非小細胞肺がん', '非小細胞肺癌', 'nsclc'] },
    { key: '小細胞肺がん', terms: ['小細胞肺がん', '小細胞肺癌', 'sclc'] },
    { key: '肺がん', terms: ['肺がん', '肺癌', '肺腫瘍'] },
    { key: '乳がん', terms: ['乳がん', '乳癌', '乳腺'] },
    { key: '胃がん', terms: ['胃がん', '胃癌', '胃腺腫'] },
    { key: '大腸がん', terms: ['大腸がん', '大腸癌', '結腸', '直腸'] },
    { key: '膵臓がん', terms: ['膵臓がん', '膵がん', '膵癌', '膵管がん'] },
    { key: '肝臓がん', terms: ['肝臓がん', '肝がん', '肝癌', '肝細胞'] },
    { key: '胆道がん', terms: ['胆道がん', '胆道癌', '胆管', '胆のう'] },
    { key: '食道がん', terms: ['食道がん', '食道癌'] },
    { key: '前立腺がん', terms: ['前立腺がん', '前立腺癌'] },
    { key: '腎臓がん', terms: ['腎臓がん', '腎がん', '腎癌', '腎細胞'] },
    { key: '膀胱がん', terms: ['膀胱がん', '膀胱癌', '尿路上皮'] },
    { key: '卵巣がん', terms: ['卵巣がん', '卵巣癌'] },
    { key: '子宮頸がん', terms: ['子宮頸がん', '子宮頸癌'] },
    { key: '子宮体がん', terms: ['子宮体がん', '子宮体癌', '子宮内膜'] },
    { key: '頭頸部がん', terms: ['頭頸部', '頭頸部がん', '頭頸部癌', '咽頭', '喉頭'] },
    { key: '悪性黒色腫', terms: ['悪性黒色腫', 'メラノーマ'] },
    { key: '悪性リンパ腫', terms: ['悪性リンパ腫', 'ホジキン', '非ホジキン', 'リンパ腫'] },
    { key: '白血病', terms: ['白血病', 'leukemia', 'aml', 'all', 'cml', 'cll'] },
    { key: '多発性骨髄腫', terms: ['骨髄腫', 'マイエローマ', 'myeloma'] },
    { key: '脳腫瘍', terms: ['脳腫瘍', 'グリオーマ', '膠芽腫'] },
    { key: '肉腫', terms: ['肉腫', 'サルコーマ', 'sarcoma'] },
    { key: 'がん化学療法', terms: ['固形がん', '固形癌', 'がん全体'] }
  ];

  for (const item of cancerDict) {
    for (const term of item.terms) {
      if (title.toLowerCase().includes(term)) return item.key;
    }
  }

  for (const item of cancerDict) {
    for (const term of item.terms) {
      if (bodyText.toLowerCase().includes(term)) return item.key;
    }
  }
  return '';
}

/**
 * 治療ラインの抽出
 */
function extractTreatmentSetting(title, bodyText) {
  const textToSearch = title + ' ' + bodyText;
  
  const settings = [
    { key: '1次治療', terms: ['1次治療', '一次治療', '初回治療', '未治療', 'first-line', '1st-line'] },
    { key: '2次治療', terms: ['2次治療', '二次治療', '既治療', '前治療歴あり', 'second-line', '2nd-line'] },
    { key: '3次治療以降', terms: ['3次治療', '三次治療', '後方治療', 'third-line', '3rd-line'] },
    { key: '術後補助療法', terms: ['術後補助', '術後化学療法', 'アジュバント', 'adjuvant'] },
    { key: '術前補助療法', terms: ['術前補助', 'ネオアジュバント', 'neoadjuvant'] },
    { key: '維持療法', terms: ['維持療法', 'maintenance'] }
  ];

  for (const item of settings) {
    for (const term of item.terms) {
      if (textToSearch.includes(term)) return item.key;
    }
  }
  return '';
}

/**
 * 評価項目の抽出
 */
function extractEndpoint(bodyText) {
  const textToSearch = bodyText;
  const found = [];
  
  const endpoints = [
    { key: 'PFS (無増悪生存期間)', terms: ['PFS', '無増悪生存期間', 'progression-free survival'] },
    { key: 'OS (全生存期間)', terms: ['OS', '全生存期間', 'overall survival'] },
    { key: 'ORR (客観的奏効率)', terms: ['ORR', '奏効率', '奏効割合', 'objective response rate'] },
    { key: 'DFS (無病生存期間)', terms: ['DFS', '無病生存期間', 'disease-free survival'] },
    { key: 'EFS (無イベント生存期間)', terms: ['EFS', '無イベント生存期間', 'event-free survival'] },
    { key: 'DOR (奏効期間)', terms: ['DOR', '奏効期間', 'duration of response'] },
    { key: 'DCR (病勢コントロール率)', terms: ['DCR', '病勢コントロール率', 'disease control rate'] }
  ];

  for (const item of endpoints) {
    for (const term of item.terms) {
      if (textToSearch.includes(term)) {
        found.push(item.key);
        break;
      }
    }
  }
  
  return found.length > 0 ? found.join(', ') : '';
}

/**
 * 試験名の抽出
 */
function extractTrialName(title, bodyText, url) {
  const excludedNames = ['PFS', 'OS', 'ORR', 'DFS', 'EFS', 'DOR', 'DCR', 'HR', 'QOL', 'AE', 'NEWS', 'POST', 'TRIAL', 'REGIMEN'];

  if (url) {
    let cleanUrl = url.split('?')[0].split('#')[0];
    if (cleanUrl.endsWith('/')) cleanUrl = cleanUrl.slice(0, -1);
    const slugMatch = cleanUrl.match(/\/([^\/]+)$/);
    if (slugMatch) {
      const slug = slugMatch[1].toLowerCase();
      const slugWords = slug.split(/[-_]/).filter(w => w.length >= 3 && w !== 'lung' && w !== 'cancer' && w !== 'trial');
      
      for (const word of slugWords) {
        const escWord = word.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
        const regex = new RegExp('\\b(' + escWord + '[A-Za-z0-9\\-]*|[A-Za-z0-9\\-]*' + escWord + ')\\b', 'i');
        const bodyMatch = bodyText.match(regex);
        if (bodyMatch) {
          const matchedName = bodyMatch[1].trim();
          if (!excludedNames.includes(matchedName.toUpperCase())) {
            return matchedName;
          }
        }
      }
    }
  }

  const bracketRegex = /[（\(]([A-Z0-9\-]{3,20}\s*(?:試験|study|trial)?)[）\)]/ig;
  let match;
  while ((match = bracketRegex.exec(title)) !== null) {
    let name = match[1].replace(/試験|study|trial/gi, '').trim();
    if (name.length >= 3 && isNaN(name) && !excludedNames.includes(name.toUpperCase())) {
      return name;
    }
  }

  const namePattern = /\b([A-Z][A-Za-z0-9\-]{3,15})\b\s*(?:試験|共同試験|臨床試験|study|trial)/;
  const bodyMatch = bodyText.match(namePattern);
  if (bodyMatch && !excludedNames.includes(bodyMatch[1].toUpperCase())) {
    return bodyMatch[1].trim();
  }

  const famousTrials = /\b([A-Z]{3,10}-(?:Breast)?\d{2,3}|[A-Z][a-z]+[A-Z][a-z]+-\d+|[A-Z]{4,12})\b/g;
  let m;
  while ((m = famousTrials.exec(title)) !== null) {
    if (m[1] && isNaN(m[1]) && m[1].length > 3 && !excludedNames.includes(m[1].toUpperCase())) {
      return m[1];
    }
  }

  return '';
}

/**
 * Phaseの抽出
 */
function extractPhase(title, bodyText) {
  const textToSearch = title + ' ' + bodyText;
  
  const phaseRegexes = [
    { key: 'Phase 3', regex: /(?:Phase|第)\s*3\s*(?:相|試験|Phase)?|Phase\s*III|第\s*III\s*相/i },
    { key: 'Phase 2', regex: /(?:Phase|第)\s*2\s*(?:相|試験|Phase)?|Phase\s*II|第\s*II\s*相/i },
    { key: 'Phase 1', regex: /(?:Phase|第)\s*1\s*(?:相|試験|Phase)?|Phase\s*I|第\s*I\s*相/i },
    { key: 'Phase 1/2', regex: /(?:Phase|第)\s*1\/2\s*(?:相|試験|Phase)?|Phase\s*I\/II|第\s*I\/II\s*相/i },
    { key: 'Phase 2/3', regex: /(?:Phase|第)\s*2\/3\s*(?:相|試験|Phase)?|Phase\s*II\/III|第\s*II\/III\s*相/i }
  ];

  for (const item of phaseRegexes) {
    if (item.regex.test(textToSearch)) return item.key;
  }
  return '';
}

/**
 * NCT番号の抽出
 */
function extractNctId(bodyText) {
  const nctRegex = /NCT\s*[-:_]?\s*(\d{8})/i;
  const match = bodyText.match(nctRegex);
  return match ? 'NCT' + match[1] : '';
}

/**
 * 試験実施国の抽出
 */
function extractCountry(bodyText) {
  const textToSearch = bodyText;
  
  if (textToSearch.includes('国際共同') || textToSearch.includes('多施設共同') || (textToSearch.includes('日本') && (textToSearch.includes('米国') || textToSearch.includes('欧州') || textToSearch.includes('海外')))) {
    return '国際共同';
  }
  if (textToSearch.includes('日本国内') || (textToSearch.includes('日本') && !textToSearch.includes('国際共同'))) {
    return '日本';
  }
  if (textToSearch.includes('米国') || textToSearch.includes('アメリカ')) {
    return '米国';
  }
  return '';
}

/**
 * 原著の抽出
 */
function extractPublication(bodyText) {
  const journals = [
    { key: 'New England Journal of Medicine (NEJM)', terms: ['New England Journal of Medicine', 'NEJM'] },
    { key: 'The Lancet', terms: ['The Lancet', 'Lancet'] },
    { key: 'Lancet Oncology', terms: ['Lancet Oncology'] },
    { key: 'Journal of Clinical Oncology (JCO)', terms: ['Journal of Clinical Oncology', 'JCO'] },
    { key: 'JAMA', terms: ['JAMA', 'Journal of the American Medical Association'] },
    { key: 'Annals of Oncology', terms: ['Annals of Oncology'] },
    { key: 'Nature Medicine', terms: ['Nature Medicine'] }
  ];

  for (const j of journals) {
    for (const term of j.terms) {
      const idx = bodyText.indexOf(term);
      if (idx !== -1) {
        const snippet = bodyText.substring(idx, Math.min(idx + 100, bodyText.length));
        const yearMatch = snippet.match(new RegExp(term + '[\\s\\S]*?(\\d{4})', 'i'));
        if (yearMatch) {
          return term + ' (' + yearMatch[1] + ')';
        }
        return j.key;
      }
    }
  }

  const doiRegex = /(?:doi\s*:\s*|https:\/\/doi\.org\/)(10\.\d{4,9}\/[-._;()/:A-Z0-9]+)/i;
  const doiMatch = bodyText.match(doiRegex);
  if (doiMatch) return 'DOI: ' + doiMatch[1];

  return '';
}

/**
 * PubMed IDの抽出
 */
function extractPubmedId(bodyText) {
  const pmidRegex = /(?:PMID\s*[:：\-]?\s*|pubmed\.ncbi\.nlm\.nih\.gov\/|pubmed\s*ID\s*[:：\-]?\s*)(\d{5,10})/i;
  const match = bodyText.match(pmidRegex);
  return match ? match[1] : '';
}

/**
 * 本文要約の作成
 */
function extractSummary(bodyText) {
  if (!bodyText) return '';
  let cleanText = bodyText
    .replace(/\s+/g, ' ')
    .trim();
  
  if (cleanText.length > 250) {
    return cleanText.substring(0, 250) + '...';
  }
  return cleanText;
}

/**
 * 記事タイトルからの評価項目の抽出
 */
function extractEndpointFromTitle(title) {
  if (!title) return '';
  const s = title.toUpperCase();
  const found = [];
  if (s.includes('PFS') || s.includes('無増悪生存期間')) found.push('PFS');
  if (s.includes('OS') || s.includes('全生存期間')) found.push('OS');
  if (s.includes('ORR') || s.includes('奏効率') || s.includes('客観的奏効率')) found.push('ORR');
  if (s.includes('DFS') || s.includes('無病生存期間')) found.push('DFS');
  if (s.includes('EFS') || s.includes('無イベント生存期間')) found.push('EFS');
  return found.length > 0 ? found.join(', ') : '';
}

/**
 * 本文HTMLからの原著URLの抽出
 */
function extractOriginalArticleUrl(html) {
  if (!html) return '';
  // Search for publication links like doi.org, nejm.org, lancet, jco, PubMed, etc.
  const regex = /href="([^"]*?(?:doi\.org|nejm\.org|thelancet\.com|ascopubs\.org|ncbi\.nlm\.nih\.gov\/pubmed\/\d+|pubmed\.ncbi\.nlm\.nih\.gov\/\d+)[^"]*?)"/i;
  const match = html.match(regex);
  if (match) return match[1];
  
  // Fallback to searching for general external links that look like a journal
  const generalRegex = /href="([^"]*?(?:nature\.com|jama\.com|science\.org|annalsofoncology\.org)[^"]*?)"/i;
  const genMatch = html.match(generalRegex);
  return genMatch ? genMatch[1] : '';
}

/**
 * 本文テキストからの定性的要約の作成
 */
function extractQualitativeSummary(bodyText) {
  if (!bodyText) return '';
  const clean = bodyText.replace(/\s+/g, ' ').trim();
  if (clean.length > 300) {
    return clean.substring(0, 300) + '...';
  }
  return clean;
}

/**
 * 将来の拡張用：本文内の結果画像に対するOCRテキスト抽出実行スタブ
 * 
 * @param {string} imageUrl 対象の画像URL
 * @return {string} OCR抽出テキストのプレースホルダー
 */
function performImageOcrStub(imageUrl) {
  Logger.log('OCR Stub called for image: ' + imageUrl);
  // 将来的に Google Cloud Vision API 等を呼び出して画像からテキストを読み取るためのスロット
  return '[OCR: Image analysis not active. Future extension point.]';
}

// ==========================================
// Google Sheets データ書き込み処理 (Database Layer)
// ==========================================

/**
 * 解析された結果を Google Sheet に追加する
 * 
 * @param {Object} data 画面から送られてきた臨床試験情報オブジェクト
 * @return {Object} 処理結果ステータス
 */
function insertRowToSheet(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    const expectedHeaders = [
      'imported_at',
      'article_url',
      'article_title',
      'publication_date',
      'cancer_type',
      'treatment_setting',
      'endpoint',
      'endpoint_from_title',
      'trial_name',
      'phase',
      'nct_id',
      'country',
      'publication',
      'original_article_url',
      'pubmed_id',
      'summary',
      'source_checked',
      'notes'
    ];

    let headerRowIndex = -1;
    let urlColumnIndex = 1;

    // 1. シートが完全に空か、ヘッダーが存在しない場合は作成
    if (values.length === 1 && values[0][0] === '') {
      sheet.appendRow(expectedHeaders);
      headerRowIndex = 0;
    } else {
      for (let i = 0; i < Math.min(values.length, 5); i++) {
        if (values[i].indexOf('article_url') !== -1) {
          headerRowIndex = i;
          urlColumnIndex = values[i].indexOf('article_url');
          break;
        }
      }
      
      if (headerRowIndex === -1) {
        sheet.insertRowBefore(1);
        sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
        headerRowIndex = 0;
        urlColumnIndex = 1;
      }
    }

    const currentValues = sheet.getDataRange().getValues();

    // 2. 同じURLがすでに登録されているかチェック
    const targetUrl = data.article_url ? data.article_url.trim() : '';
    if (targetUrl) {
      for (let r = headerRowIndex + 1; r < currentValues.length; r++) {
        const cellValue = currentValues[r][urlColumnIndex] ? String(currentValues[r][urlColumnIndex]).trim() : '';
        if (cellValue === targetUrl) {
          return {
            success: false,
            message: '重複エラー: この記事URLは既にシートに登録されています。',
            duplicateRow: r + 1
          };
        }
      }
    }

    // 3. 書き込み用データのマッピング
    const headers = currentValues[headerRowIndex];
    const newRow = new Array(headers.length).fill('');
    const now = new Date();
    const formattedTimestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');

    for (let c = 0; c < headers.length; c++) {
      const headerName = headers[c];
      switch (headerName) {
        case 'imported_at':
          newRow[c] = formattedTimestamp;
          break;
        case 'article_url':
          newRow[c] = data.article_url || '';
          break;
        case 'article_title':
          newRow[c] = data.article_title || '';
          break;
        case 'publication_date':
          newRow[c] = data.publication_date || '';
          break;
        case 'cancer_type':
          newRow[c] = data.cancer_type || '';
          break;
        case 'treatment_setting':
          newRow[c] = data.treatment_setting || '';
          break;
        case 'endpoint':
          newRow[c] = data.endpoint || '';
          break;
        case 'endpoint_from_title':
          newRow[c] = data.endpoint_from_title || '';
          break;
        case 'trial_name':
          newRow[c] = data.trial_name || '';
          break;
        case 'phase':
          newRow[c] = data.phase || '';
          break;
        case 'nct_id':
          newRow[c] = data.nct_id || '';
          break;
        case 'country':
          newRow[c] = data.country || '';
          break;
        case 'publication':
          newRow[c] = data.publication || '';
          break;
        case 'original_article_url':
          newRow[c] = data.original_article_url || '';
          break;
        case 'pubmed_id':
          newRow[c] = data.pubmed_id || '';
          break;
        case 'summary':
          newRow[c] = data.summary || '';
          break;
        case 'source_checked':
          newRow[c] = '未確認'; // 初期値固定
          break;
        case 'notes':
          newRow[c] = data.notes || '';
          break;
        default:
          if (data[headerName] !== undefined) {
            newRow[c] = data[headerName];
          }
      }
    }

    // 4. シートへの最終追加
    sheet.appendRow(newRow);
    
    const lastRowIndex = sheet.getLastRow();
    sheet.getRange(lastRowIndex, 1, 1, headers.length).setHorizontalAlignment('left');

    return {
      success: true,
      message: 'シートへ正常に追加されました。'
    };

  } catch (e) {
    Logger.log('追加エラー: ' + e.message);
    return {
      success: false,
      message: 'シート追加中にエラーが発生しました: ' + e.message
    };
  }
}

// =========================================================================
// 新機能：結果抽出・集積シート管理機能 (Results Extraction V3.0)
// =========================================================================

/**
 * 臨床試験結果集積用シート(Results_raw, Results, Trials, Arms, Outcomes)をチェックし、存在しなければ作成する
 */
function initResultsSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // A. Results_raw シートの初期化 (抽出候補用)
  let rawSheet = ss.getSheetByName('Results_raw');
  const expectedRawHeaders = [
    'imported_at',
    'article_url',
    'trial_name',
    'nct_id',
    'pmid',
    'endpoint',
    'source_type',
    'source_text',
    'extracted_numbers',
    'numeric_result_found',
    'result_extraction_status',
    'confidence',
    'checked',
    'notes'
  ];
  if (!rawSheet) {
    rawSheet = ss.insertSheet('Results_raw');
    rawSheet.appendRow(expectedRawHeaders);
    rawSheet.getRange(1, 1, 1, expectedRawHeaders.length).setFontWeight('bold').setBackground('#f1f5f9');
  } else {
    // 既存シートのヘッダーが古い場合は安全に上書きする (スキーマ移行)
    const lastCol = Math.max(rawSheet.getLastColumn(), 1);
    const currentHeaders = rawSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    let isDifferent = currentHeaders.length !== expectedRawHeaders.length;
    if (!isDifferent) {
      for (let j = 0; j < expectedRawHeaders.length; j++) {
        if (currentHeaders[j] !== expectedRawHeaders[j]) {
          isDifferent = true;
          break;
        }
      }
    }
    if (isDifferent) {
      // 既存データがある場合は、安全にヘッダー行のみ更新
      rawSheet.getRange(1, 1, 1, expectedRawHeaders.length).setValues([expectedRawHeaders]);
      rawSheet.getRange(1, 1, 1, expectedRawHeaders.length).setFontWeight('bold').setBackground('#f1f5f9');
    }
  }

  // B. Results シートの初期化 (確認済みデータ用)
  let resultsSheet = ss.getSheetByName('Results');
  const expectedResultsHeaders = [
    'result_id',
    'article_url',
    'trial_name',
    'nct_id',
    'pmid',
    'cancer_type',
    'treatment_setting',
    'population',
    'comparison',
    'endpoint',
    'endpoint_type',
    'arm_a',
    'arm_b',
    'value_a',
    'value_b',
    'median_a',
    'median_b',
    'unit',
    'hr',
    'ci_lower',
    'ci_upper',
    'p_value',
    'orr_a',
    'orr_b',
    'dcr_a',
    'dcr_b',
    'grade3_ae_a',
    'grade3_ae_b',
    'discontinuation_a',
    'discontinuation_b',
    'treatment_related_death_a',
    'treatment_related_death_b',
    'source_type',
    'source_url',
    'source_text',
    'confidence',
    'checked',
    'notes'
  ];
  if (!resultsSheet) {
    resultsSheet = ss.insertSheet('Results');
    resultsSheet.appendRow(expectedResultsHeaders);
    resultsSheet.getRange(1, 1, 1, expectedResultsHeaders.length).setFontWeight('bold').setBackground('#ccfbf1');
  }

  // C. Trials シートの初期化 (ClinicalTrials 基本情報用)
  let trialsSheet = ss.getSheetByName('Trials');
  const expectedTrialsHeaders = [
    'nct_id',
    'brief_title',
    'official_title',
    'phases',
    'study_type',
    'overall_status',
    'conditions',
    'interventions',
    'enrollment_count',
    'start_date',
    'completion_date',
    'primary_completion_date'
  ];
  if (!trialsSheet) {
    trialsSheet = ss.insertSheet('Trials');
    trialsSheet.appendRow(expectedTrialsHeaders);
    trialsSheet.getRange(1, 1, 1, expectedTrialsHeaders.length).setFontWeight('bold').setBackground('#fed7aa');
  }

  // D. Arms シートの初期化 (ClinicalTrials Arm群情報用)
  let armsSheet = ss.getSheetByName('Arms');
  const expectedArmsHeaders = [
    'nct_id',
    'arm_group_label',
    'arm_group_type',
    'arm_group_description',
    'interventions'
  ];
  if (!armsSheet) {
    armsSheet = ss.insertSheet('Arms');
    armsSheet.appendRow(expectedArmsHeaders);
    armsSheet.getRange(1, 1, 1, expectedArmsHeaders.length).setFontWeight('bold').setBackground('#fef08a');
  }

  // E. Outcomes シートの初期化 (ClinicalTrials 評価項目用)
  let outcomesSheet = ss.getSheetByName('Outcomes');
  const expectedOutcomesHeaders = [
    'nct_id',
    'outcome_type',
    'outcome_title',
    'outcome_description',
    'outcome_time_frame'
  ];
  if (!outcomesSheet) {
    outcomesSheet = ss.insertSheet('Outcomes');
    outcomesSheet.appendRow(expectedOutcomesHeaders);
    outcomesSheet.getRange(1, 1, 1, expectedOutcomesHeaders.length).setFontWeight('bold').setBackground('#bfdbfe');
  }
}

/**
 * 記事本文から臨床試験結果の候補文を抽出し、Results_rawシートに保存する
 * 
 * @param {Object} data 記事メタデータと生HTML
 * @return {Object[]} 抽出された結果候補オブジェクトの配列
 */
function extractAndSaveRawResults(data) {
  try {
    // 新しいシートを準備 (存在しなければ自動生成)
    initResultsSheets();

    const url = data.article_url || '';
    const trialName = data.trial_name || '';
    const nctId = data.nct_id || '';
    const pmid = data.pubmed_id || '';
    const rawHtml = data.raw_html || '';
    
    // HTMLタグを剥がして本文テキストを取得
    const bodyText = cleanHtml(rawHtml);
    
    // 候補文の抽出実行 (モジュールエンジンを呼び出し)
    let candidates = TrialResultExtractor.extractCandidates(bodyText, trialName, nctId, pmid, url);
    
    if (candidates.length === 0) {
      // 本文に臨床試験評価項目に関する主要キーワードがあるかチェック (Requirement 5)
      const keywords = ['PFS', 'OS', 'DFS', 'EFS', 'RFS', 'ORR', 'DCR', 'pCR', '無増悪生存期間', '全生存期間', '奏効率', '病勢制御率'];
      let hasKeywords = false;
      for (const kw of keywords) {
        if (bodyText.includes(kw)) {
          hasKeywords = true;
          break;
        }
      }
      
      const status = hasKeywords ? 'image_only' : 'manual_required';
      const notes = hasKeywords ? '本文に結果キーワードが含まれますが、結果数値が検出されませんでした（画像テーブルの可能性があります）。' : '結果数値およびキーワードが本文から見つかりませんでした。手入力またはPubMed補完が必要です。';
      
      const fallbackCandidate = {
        article_url: url,
        trial_name: trialName,
        nct_id: nctId,
        pmid: pmid,
        detected_endpoint: data.endpoint_from_title || data.endpoint || 'Other',
        source_type: 'がん対策図鑑',
        source_text: notes,
        detected_numbers: {},
        numeric_result_found: false,
        result_extraction_status: status,
        confidence: 'low',
        checked: '未確認',
        notes: notes
      };
      candidates = [fallbackCandidate];
    }

    saveRawCandidatesToSheet(candidates, url);
    return candidates;

  } catch (e) {
    Logger.log('結果抽出エラー: ' + e.message);
    throw new Error('結果候補の抽出中にエラーが発生しました: ' + e.message);
  }
}

/**
 * ユーザーが確認・編集した構造化臨床試験結果を Results シートに保存する
 * 
 * @param {Object} verifiedData 確認済みの構造化データ
 * @return {Object} 処理結果ステータス
 */
function insertVerifiedResult(verifiedData) {
  try {
    initResultsSheets();
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const resultsSheet = ss.getSheetByName('Results');
    const headers = resultsSheet.getRange(1, 1, 1, resultsSheet.getLastColumn()).getValues()[0];
    
    // ユニークなResult IDを生成 (RES-XXXXX)
    const randId = 'RES-' + Math.floor(10000 + Math.random() * 90000);
    verifiedData.result_id = randId;
    verifiedData.checked = '確認済'; // 確認マーク (Requirement 9)

    // ヘッダーのカラム順にデータを配列にマッピング
    const newRow = new Array(headers.length).fill('');
    for (let c = 0; c < headers.length; c++) {
      const colName = headers[c];
      if (verifiedData[colName] !== undefined) {
        newRow[c] = verifiedData[colName];
      }
    }

    resultsSheet.appendRow(newRow);
    
    // 保存に成功したため、該当する Results_raw の checked 列を「確認済」に更新する
    try {
      const rawSheet = ss.getSheetByName('Results_raw');
      const lastRow = rawSheet.getLastRow();
      if (lastRow > 1) {
        const urls = rawSheet.getRange(2, 2, lastRow - 1, 1).getValues(); // Column 2: article_url
        const texts = rawSheet.getRange(2, 8, lastRow - 1, 1).getValues(); // Column 8: source_text
        for (let r = 0; r < urls.length; r++) {
          if (urls[r][0] === verifiedData.article_url && texts[r][0] === verifiedData.source_text) {
            rawSheet.getRange(r + 2, 13).setValue('確認済'); // Column 13: checked
            break;
          }
        }
      }
    } catch (rawUpdateErr) {
      Logger.log('Results_raw更新エラー (無視して進行): ' + rawUpdateErr.message);
    }

    return {
      success: true,
      message: '臨床試験結果が「Results」シートに正常に登録されました！ ID: ' + randId
    };

  } catch (e) {
    Logger.log('結果登録エラー: ' + e.message);
    return {
      success: false,
      message: '結果の登録中にエラーが発生しました: ' + e.message
    };
  }
}

// ==========================================
// インテリジェントパースエンジン (TrialResultExtractor Namespace)
// ==========================================

const TrialResultExtractor = {
  
  /**
   * 本文テキストから結果候補文をすべて抽出して構造化モデルを作成する
   */
  extractCandidates: function(bodyText, trialName, nctId, pmid, url) {
    if (!bodyText) return [];
    
    // 1. 文割りの実行
    const sentences = this.splitIntoSentences(bodyText);
    const candidates = [];
    
    // キーワード一覧 (Requirement 4)
    const keywords = [
      'PFS', 'OS', 'DFS', 'EFS', 'RFS', 'ORR', 'DCR', 'pCR',
      '無増悪生存期間', '全生存期間', '奏効率', '病勢制御率',
      '中央値', 'HR', 'ハザード比', '95%CI', '95％信頼区間', 'p値',
      'Grade 3', 'グレード3', '有害事象', '治療関連死'
    ];

    for (const sentence of sentences) {
      // キーワードが1つでも含まれるかチェック
      let matchedKeyword = null;
      for (const kw of keywords) {
        if (sentence.includes(kw)) {
          matchedKeyword = kw;
          break; // 最初に見つかったキーワード
        }
      }
      
      if (matchedKeyword) {
        // 2. エンドポイントの自動推定
        const endpoint = this.estimateEndpoint(sentence);
        
        // 3. 数値情報の正規表現抽出 (Requirement 6 / 11: 存在しない数値は補完しない)
        const numbers = this.extractNumbers(sentence);
        
        // 4. 確信度 (Confidence) の判定 (Requirement 12)
        const confidence = this.calculateConfidence(endpoint, numbers, sentence);
        
        // 5. プレビュー編集フォームで扱いやすいように ARM や主要数値を割り当て (初期提案値)
        let arm_a = '';
        let arm_b = '';
        let value_a = '';
        let value_b = '';
        let median_a = '';
        let median_b = '';
        
        // 抽出された割合 (%) の自動マッピング
        if (numbers.percentages && numbers.percentages.length >= 2) {
          value_a = numbers.percentages[0];
          value_b = numbers.percentages[1];
        } else if (numbers.percentages && numbers.percentages.length === 1) {
          value_a = numbers.percentages[0];
        }
        
        // 抽出された「か月」(median) の自動マッピング
        if (numbers.medians && numbers.medians.length >= 2) {
          median_a = numbers.medians[0];
          median_b = numbers.medians[1];
        } else if (numbers.medians && numbers.medians.length === 1) {
          median_a = numbers.medians[0];
        }

        // Endpoint別値マッピング補完
        let orr_a = '', orr_b = '';
        let dcr_a = '', dcr_b = '';
        let grade3_ae_a = '', grade3_ae_b = '';
        
        if (endpoint === 'ORR') {
          orr_a = value_a; orr_b = value_b;
        } else if (endpoint === 'DCR') {
          dcr_a = value_a; dcr_b = value_b;
        } else if (endpoint === 'Safety') {
          grade3_ae_a = value_a; grade3_ae_b = value_b;
        }

        candidates.push({
          article_url: url,
          trial_name: trialName,
          nct_id: nctId,
          pmid: pmid,
          detected_endpoint: endpoint,
          detected_result_text: sentence,
          detected_numbers: numbers,
          source_text: sentence,
          confidence: confidence,
          
          // 編集用構造化データの初期枠 (ユーザーがサイドバーで調整可能)
          comparison: '',
          endpoint: endpoint,
          endpoint_type: '主要評価項目',
          arm_a: arm_a,
          arm_b: arm_b,
          value_a: value_a,
          value_b: value_b,
          median_a: median_a,
          median_b: median_b,
          unit: numbers.medians && numbers.medians.length > 0 ? 'か月' : (numbers.percentages && numbers.percentages.length > 0 ? '%' : ''),
          hr: numbers.hr || '',
          ci_lower: numbers.ci_lower || '',
          ci_upper: numbers.ci_upper || '',
          p_value: numbers.p_value || '',
          orr_a: orr_a,
          orr_b: orr_b,
          dcr_a: dcr_a,
          dcr_b: dcr_b,
          grade3_ae_a: grade3_ae_a,
          grade3_ae_b: grade3_ae_b,
          discontinuation_a: '',
          discontinuation_b: '',
          treatment_related_death_a: '',
          treatment_related_death_b: '',
          notes: ''
        });
      }
    }
    
    return candidates;
  },

  /**
   * 文単位での分割
   */
  splitIntoSentences: function(text) {
    if (!text) return [];
    
    // 全角の「。」「！」「？」、改行、および直後に大文字や日本語が続く「. + 半角スペース」で文分割
    const rawSentences = text.split(/(?:。|！|？|\r?\n|\.\s+(?=[A-Z\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF]))/);
    
    return rawSentences
      .map(s => s.trim())
      .filter(s => s.length > 6 && !s.startsWith('<') && !s.includes('http')); // 無効なパーツの除去
  },

  /**
   * キーワードに基づくエンドポイントの推定 (Requirement 5)
   */
  estimateEndpoint: function(sentence) {
    const s = sentence;
    if (s.includes('PFS') || s.includes('無増悪生存期間')) return 'PFS';
    if (s.includes('OS') || s.includes('全生存期間')) return 'OS';
    if (s.includes('ORR') || s.includes('奏効率')) return 'ORR';
    if (s.includes('DCR') || s.includes('病勢制御率')) return 'DCR';
    if (s.includes('DFS') || s.includes('無病生存期間')) return 'DFS';
    if (s.includes('EFS') || s.includes('無イベント生存期間')) return 'EFS';
    if (s.includes('RFS') || s.includes('無再発生存期間')) return 'RFS';
    if (s.includes('pCR') || s.includes('病理学的完全奏効')) return 'pCR';
    if (s.includes('Grade 3') || s.includes('グレード3') || s.includes('有害事象') || s.includes('治療関連死')) return 'Safety';
    
    return 'Other';
  },

  /**
   * 文脈からの数値情報の正規表現抽出 (Requirement 6 / 11)
   */
  extractNumbers: function(sentence) {
    const s = sentence;
    
    // A. ハザード比 (HR) の抽出
    let hrVal = '';
    const hrMatch = s.match(/HR\s*[:：＝=]?\s*(\d+\.?\d*)|ハザード比\s*[:：＝=]?\s*(\d+\.?\d*)/i);
    if (hrMatch) {
      hrVal = hrMatch[1] || hrMatch[2];
    }

    // B. 95%信頼区間 (95%CI) の抽出 (上限・下限)
    let ciLower = '';
    let ciUpper = '';
    const ciMatch = s.match(/(?:95%|９５％)\s*(?:CI|信頼区間)\s*[:：＝=]?\s*[（(]?\s*(\d+\.?\d*)\s*[-–,，]\s*(\d+\.?\d*)\s*[）)]?/i);
    if (ciMatch) {
      ciLower = ciMatch[1];
      ciUpper = ciMatch[2];
    }

    // C. p値 (p-value) の抽出
    let pVal = '';
    const pMatch = s.match(/p\s*(?:value|値)?\s*[:：＝=]?\s*([<＜>＞=＝]?\s*\d+\.?\d+)/i);
    if (pMatch) {
      pVal = pMatch[1].replace(/\s+/g, ''); // 空白除去
    }

    // D. 月数/中央値 (median) の抽出 (例: 35.5か月, 12 months)
    const medians = [];
    const medianRegex = /(\d+\.?\d*)\s*(?:か月|ヶ月|カ月|ヶ月|月|months)/gi;
    let mMatch;
    while ((mMatch = medianRegex.exec(s)) !== null) {
      medians.push(mMatch[1]);
    }
    
    // E. 割合 (%) の抽出 (例: 80%, 25.4％)
    const percentages = [];
    const pctRegex = /(\d+\.?\d*)\s*[%％]/g;
    let pPct;
    while ((pPct = pctRegex.exec(s)) !== null) {
      percentages.push(pPct[1]);
    }

    return {
      hr: hrVal,
      ci_lower: ciLower,
      ci_upper: ciUpper,
      p_value: pVal,
      medians: medians,
      percentages: percentages
    };
  },

  /**
   * 確信度 (Confidence) の判定ロジック (Requirement 12)
   */
  calculateConfidence: function(endpoint, numbers, sentence) {
    const hasEndpoint = endpoint !== 'Other';
    const hasKeyNumber = numbers.hr || (numbers.medians && numbers.medians.length > 0) || (numbers.percentages && numbers.percentages.length > 0);
    const hasSourceText = sentence && sentence.length > 15;
    
    if (hasEndpoint && hasKeyNumber && hasSourceText) {
      // エンドポイント、数値、ソースがすべて揃っている場合
      return 'high';
    } else if (hasEndpoint && hasSourceText) {
      // エンドポイントはあるが、数値が不明瞭な場合
      return 'medium';
    } else {
      // 単にキーワードだけが含まれる場合
      return 'low';
    }
  }
};

/**
 * PMIDからPubMedのAbstract（抄録）テキストをXML形式で取得して平文に整形する
 * 
 * @param {string} pmid PubMed ID
 * @return {string} 抄録テキスト
 */
function fetchPubMedAbstractText(pmid) {
  if (!pmid) return '';
  const cleanPmid = String(pmid).trim();
  try {
    const url = 'https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pubmed&id=' + cleanPmid + '&retmode=xml';
    const response = UrlFetchApp.fetch(url, { 
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    
    if (response.getResponseCode() !== 200) {
      Logger.log('PubMed API fetch failed: ' + response.getResponseCode());
      return '';
    }
    
    const xmlText = response.getContentText();
    
    // AbstractTextエレメントを正規表現で走査し、Label属性（BACKGROUND, METHODS, RESULTS等）も拾う
    const abstractMatch = xmlText.match(/<AbstractText[^>]*>([\s\S]*?)<\/AbstractText>/g);
    if (abstractMatch) {
      return abstractMatch.map(m => {
        let text = m.replace(/<[^>]*>/g, '').trim(); // HTMLタグ除去
        // Label属性をパース
        const labelAttr = m.match(/Label="([^"]+)"/i);
        if (labelAttr && labelAttr[1]) {
          text = labelAttr[1] + ': ' + text;
        }
        return text;
      }).join(' ');
    }
    
    // AbstractTextがない場合はArticleTitle等で代替
    const titleMatch = xmlText.match(/<ArticleTitle>([\s\S]*?)<\/ArticleTitle>/i);
    if (titleMatch) {
      return titleMatch[1].replace(/<[^>]*>/g, '').trim();
    }
    
    return '';
  } catch (e) {
    Logger.log('PubMed API エラー: ' + e.message);
    return '';
  }
}

/**
 * NCT IDからClinicalTrials.gov API v2経由で試験詳細情報を取得する
 * 
 * @param {string} nctId NCT ID (例: NCT03450278)
 * @return {Object|null} 試験テキストと生データのオブジェクト
 */
function fetchClinicalTrialsInfo(nctId) {
  if (!nctId) return null;
  let cleanNct = String(nctId).trim().toUpperCase();
  if (!cleanNct.startsWith('NCT')) {
    cleanNct = 'NCT' + cleanNct;
  }
  
  try {
    const url = 'https://clinicaltrials.gov/api/v2/studies/' + cleanNct;
    const response = UrlFetchApp.fetch(url, { 
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    
    if (response.getResponseCode() !== 200) {
      Logger.log('ClinicalTrials API v2 fetch failed: ' + response.getResponseCode());
      return null;
    }
    
    const data = JSON.parse(response.getContentText());
    const protocol = data.protocolSection || {};
    
    const title = protocol.identificationModule ? (protocol.identificationModule.officialTitle || protocol.identificationModule.briefTitle || '') : '';
    const briefSummary = protocol.descriptionModule ? (protocol.descriptionModule.briefSummary || '') : '';
    const detailedDesc = protocol.descriptionModule ? (protocol.descriptionModule.detailedDescription || '') : '';
    
    // 評価項目の情報
    let outcomesText = '';
    if (protocol.outcomesModule && protocol.outcomesModule.primaryOutcomes) {
      outcomesText = 'Primary Outcomes:\n' + protocol.outcomesModule.primaryOutcomes.map((o, idx) => {
        return `${idx + 1}. ${o.measure || ''} (${o.timeFrame || ''}) - ${o.description || ''}`;
      }).join('\n');
    }
    
    // 結果データがAPI上にあれば含める
    let resultsText = '';
    if (data.resultsSection) {
      resultsText = 'Trial Results Section Available on ClinicalTrials.gov.';
    }
    
    const fullText = [
      'Title: ' + title,
      'Brief Summary:\n' + briefSummary,
      'Detailed Description:\n' + detailedDesc,
      outcomesText,
      resultsText
    ].filter(Boolean).join('\n\n');
    
    return {
      title: title,
      fullText: fullText,
      raw: data
    };
  } catch (e) {
    Logger.log('ClinicalTrials API エラー: ' + e.message);
    return null;
  }
}

/**
 * PubMed (NCBI E-utilities) から結果候補を取得し、Results_rawに一時保存する
 * 
 * @param {Object} data 記事データ
 * @return {Object[]} 抽出された候補文オブジェクト配列
 */
function supplementFromPubMed(data) {
  try {
    const pmid = data.pubmed_id;
    if (!pmid) throw new Error('PubMed IDが見つかりません。');
    
    const abstract = fetchPubMedAbstractText(pmid);
    if (!abstract) throw new Error('PubMedから抄録（Abstract）を取得できませんでした。');
    
    const trialName = data.trial_name || '';
    const nctId = data.nct_id || '';
    const sourceUrl = 'https://pubmed.ncbi.nlm.nih.gov/' + pmid;
    
    // E-utilitiesテキストに対して抽出を実行
    const candidates = TrialResultExtractor.extractCandidates(abstract, trialName, nctId, pmid, data.article_url);
    
    if (candidates.length === 0) {
      const dummyCandidate = {
        article_url: data.article_url,
        trial_name: trialName,
        nct_id: nctId,
        pmid: pmid,
        detected_endpoint: data.endpoint_from_title || data.endpoint || 'Other',
        source_type: 'PubMed',
        source_text: 'PubMedから抄録を取得しましたが、結果数値は検出されませんでした。',
        detected_numbers: {},
        numeric_result_found: false,
        result_extraction_status: 'abstract_no_numeric_result',
        confidence: 'low',
        checked: '未確認',
        notes: 'PubMedから補完 (結果数値なし)'
      };
      saveRawCandidatesToSheet([dummyCandidate], sourceUrl);
      return [dummyCandidate];
    }
    
    // source_url, source_textの明示的な割り当て
    for (const c of candidates) {
      c.source_url = sourceUrl;
      c.source_type = 'PubMed';
      c.source_text = c.source_text || c.detected_result_text;
      
      const hasNumbers = !!(c.detected_numbers.hr || (c.detected_numbers.medians && c.detected_numbers.medians.length > 0) || (c.detected_numbers.percentages && c.detected_numbers.percentages.length > 0));
      c.numeric_result_found = hasNumbers;
      c.result_extraction_status = hasNumbers ? 'abstract_found' : 'abstract_no_numeric_result';
      c.notes = 'PubMedから補完';
    }
    
    saveRawCandidatesToSheet(candidates, sourceUrl);
    return candidates;
  } catch (e) {
    Logger.log('PubMed補完エラー: ' + e.message);
    throw new Error('PubMedからの補完中にエラーが発生しました: ' + e.message);
  }
}

/**
 * ClinicalTrials.gov から結果候補を取得し、Results_rawに一時保存する
 * 
 * @param {Object} data 記事データ
 * @return {Object[]} 抽出された候補文オブジェクト配列
 */
function supplementFromClinicalTrials(data) {
  try {
    const nctId = data.nct_id;
    if (!nctId) throw new Error('NCT IDが見つかりません。');
    
    const trialInfo = fetchClinicalTrialsInfo(nctId);
    if (!trialInfo || !trialInfo.fullText) throw new Error('ClinicalTrials.govから試験詳細テキストを取得できませんでした。');
    
    const trialName = data.trial_name || '';
    const pmid = data.pubmed_id || '';
    const sourceUrl = 'https://clinicaltrials.gov/study/' + nctId;
    
    // ClinicalTrials詳細テキストに対して抽出を実行
    const candidates = TrialResultExtractor.extractCandidates(trialInfo.fullText, trialName, nctId, pmid, data.article_url);
    
    if (candidates.length === 0) {
      const dummyCandidate = {
        article_url: data.article_url,
        trial_name: trialName,
        nct_id: nctId,
        pmid: pmid,
        detected_endpoint: data.endpoint_from_title || data.endpoint || 'Other',
        source_type: 'ClinicalTrials.gov',
        source_text: 'ClinicalTrials.govからデータを取得しましたが、結果数値は掲載されていませんでした。',
        detected_numbers: {},
        numeric_result_found: false,
        result_extraction_status: 'clinicaltrials_no_results',
        confidence: 'low',
        checked: '未確認',
        notes: 'ClinicalTrials.govから補完 (結果数値なし)'
      };
      saveRawCandidatesToSheet([dummyCandidate], sourceUrl);
      return [dummyCandidate];
    }
    
    // source_url, source_textの明示的な割り当て
    for (const c of candidates) {
      c.source_url = sourceUrl;
      c.source_type = 'ClinicalTrials.gov';
      c.source_text = c.source_text || c.detected_result_text;
      
      const hasNumbers = !!(c.detected_numbers.hr || (c.detected_numbers.medians && c.detected_numbers.medians.length > 0) || (c.detected_numbers.percentages && c.detected_numbers.percentages.length > 0));
      c.numeric_result_found = hasNumbers;
      c.result_extraction_status = hasNumbers ? 'clinicaltrials_found' : 'clinicaltrials_no_results';
      c.notes = 'ClinicalTrials.govから補完';
    }
    
    saveRawCandidatesToSheet(candidates, sourceUrl);
    return candidates;
  } catch (e) {
    Logger.log('ClinicalTrials補完エラー: ' + e.message);
    throw new Error('ClinicalTrials.govからの補完中にエラーが発生しました: ' + e.message);
  }
}

/**
 * ユーザーがサイドバーで調整・編集した結果候補（未確認）を Results_raw シートに上書きまたは保存する
 * 
 * @param {Object} data ユーザー編集済みの候補データ
 * @return {Object} 処理結果ステータス
 */
function saveSingleRawCandidate(data) {
  try {
    initResultsSheets();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawSheet = ss.getSheetByName('Results_raw');
    
    // 重複チェック (同じURLかつ同じsource_textが既に存在するか確認して更新、なければ追加)
    let existingRow = -1;
    const lastRow = rawSheet.getLastRow();
    if (lastRow > 1) {
      const urlVals = rawSheet.getRange(2, 2, lastRow - 1, 1).getValues(); // Column 2: article_url
      const textVals = rawSheet.getRange(2, 8, lastRow - 1, 1).getValues(); // Column 8: source_text
      for (let i = 0; i < urlVals.length; i++) {
        if (urlVals[i][0] === data.article_url && textVals[i][0] === data.source_text) {
          existingRow = i + 2;
          break;
        }
      }
    }
    
    // Reconstruct numbers JSON structure
    const numbers = {
      hr: data.hr || '',
      ci_lower: data.ci_lower || '',
      ci_upper: data.ci_upper || '',
      p_value: data.p_value || '',
      medians: [data.median_a || '', data.median_b || ''].filter(Boolean),
      percentages: [data.value_a || '', data.value_b || ''].filter(Boolean)
    };
    
    const hasNumbers = !!(numbers.hr || (numbers.medians && numbers.medians.length > 0) || (numbers.percentages && numbers.percentages.length > 0));
    const numericResultFound = data.numeric_result_found !== undefined ? data.numeric_result_found : hasNumbers;
    let status = data.result_extraction_status;
    if (!status) {
      status = numericResultFound ? 'text_found' : 'image_only';
    }
    
    const now = new Date();
    const formattedTimestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    
    const rowValues = [
      data.imported_at || formattedTimestamp,
      data.article_url || '',
      data.trial_name || '',
      data.nct_id || '',
      data.pmid || '',
      data.endpoint || '',
      data.source_type || 'がん対策図鑑',
      data.source_text || '',
      JSON.stringify(numbers),
      numericResultFound,
      status,
      data.confidence || 'low',
      data.checked || '未確認',
      data.notes || ''
    ];
    
    if (existingRow !== -1) {
      rawSheet.getRange(existingRow, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      rawSheet.appendRow(rowValues);
    }
    
    return {
      success: true,
      message: '結果候補が Results_raw シートに正常に保存されました！'
    };
  } catch (e) {
    Logger.log('単一結果候補保存エラー: ' + e.message);
    return {
      success: false,
      message: '結果候補の保存中にエラーが発生しました: ' + e.message
    };
  }
}

/**
 * 抽出された候補配列を Results_raw シートに一括ログ保存するヘルパー関数
 * 
 * @param {Object[]} candidates 候補配列
 * @param {string} sourceUrl ソース元URL (PubMed / ClinicalTrials)
 */
function saveRawCandidatesToSheet(candidates, sourceUrl) {
  initResultsSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName('Results_raw');
  
  for (const c of candidates) {
    let isDuplicate = false;
    const lastRow = rawSheet.getLastRow();
    if (lastRow > 1) {
      const urlVals = rawSheet.getRange(2, 2, lastRow - 1, 1).getValues(); // article_url is Column 2
      const textVals = rawSheet.getRange(2, 8, lastRow - 1, 1).getValues(); // source_text is Column 8
      for (let i = 0; i < urlVals.length; i++) {
        if (urlVals[i][0] === c.article_url && textVals[i][0] === c.source_text) {
          isDuplicate = true;
          break;
        }
      }
    }
    
    if (!isDuplicate) {
      // Determine numeric_result_found
      const numbers = c.detected_numbers || c.extracted_numbers || {};
      const hasNumbers = !!(numbers.hr || (numbers.medians && numbers.medians.length > 0) || (numbers.percentages && numbers.percentages.length > 0));
      
      const numericResultFound = c.numeric_result_found !== undefined ? c.numeric_result_found : hasNumbers;
      let status = c.result_extraction_status;
      if (!status) {
        status = numericResultFound ? 'text_found' : 'image_only';
      }
      
      const now = new Date();
      const formattedTimestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
      
      const rowValues = [
        c.imported_at || formattedTimestamp,
        c.article_url || '',
        c.trial_name || '',
        c.nct_id || '',
        c.pmid || '',
        c.detected_endpoint || c.endpoint || '',
        c.source_type || (sourceUrl && sourceUrl.includes('pubmed') ? 'PubMed' : (sourceUrl && sourceUrl.includes('clinicaltrials') ? 'ClinicalTrials.gov' : 'がん対策図鑑')),
        c.source_text || '',
        JSON.stringify(numbers),
        numericResultFound,
        status,
        c.confidence || 'low',
        c.checked || '未確認',
        c.notes || ''
      ];
      
      rawSheet.appendRow(rowValues);
    }
  }
}

// ==========================================
// GitHubリポジトリ連携機能 (GitHub Sync V1.0)
// ==========================================

/**
 * GitHubの連携設定をUserPropertiesに安全に保存する
 * 
 * @param {string} token GitHub Personal Access Token
 * @param {string} repo リポジトリパス (例: owner/repo)
 * @param {string} branch ブランチ名
 * @return {Object} 処理結果ステータス
 */
function saveGitHubSettings(token, repo, branch) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    if (token !== undefined && token !== null) {
      const cleanToken = token.trim();
      // マスキング入力の場合は保存をスキップ
      if (cleanToken !== '' && !cleanToken.startsWith('***')) {
        userProperties.setProperty('GITHUB_TOKEN', cleanToken);
      }
    }
    userProperties.setProperty('GITHUB_REPO', (repo || 'awakuma1113/clinicaltrials').trim());
    userProperties.setProperty('GITHUB_BRANCH', (branch || 'main').trim());
    
    return {
      success: true,
      message: 'GitHubの設定を正常に保存しました。'
    };
  } catch (e) {
    Logger.log('GitHub設定保存エラー: ' + e.message);
    return {
      success: false,
      message: '設定の保存中にエラーが発生しました: ' + e.message
    };
  }
}

/**
 * 保存されているGitHub設定を取得する（トークンはセキュリティのため部分マスク）
 * 
 * @return {Object} 設定情報
 */
function getGitHubSettings() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const token = userProperties.getProperty('GITHUB_TOKEN') || '';
    const repo = userProperties.getProperty('GITHUB_REPO') || 'awakuma1113/clinicaltrials';
    const branch = userProperties.getProperty('GITHUB_BRANCH') || 'main';
    
    // トークンが存在する場合はマスクして返す
    let maskedToken = '';
    if (token) {
      maskedToken = token.substring(0, 4) + '****************' + token.substring(token.length - 4);
    }
    
    return {
      hasToken: !!token,
      maskedToken: maskedToken,
      repo: repo,
      branch: branch
    };
  } catch (e) {
    Logger.log('GitHub設定取得エラー: ' + e.message);
    return {
      hasToken: false,
      maskedToken: '',
      repo: 'awakuma1113/clinicaltrials',
      branch: 'main'
    };
  }
}

/**
 * 指定されたシートのデータをCSVフォーマットのテキストにシリアライズする
 * 
 * @param {string} sheetName 対象のシート名
 * @return {string} CSV文字列
 */
function convertSheetToCsvText(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return '';
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return '';
  
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  let csvText = '';
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const csvRow = [];
    for (let j = 0; j < row.length; j++) {
      let cellValue = '';
      if (row[j] instanceof Date) {
        cellValue = Utilities.formatDate(row[j], ss.getSpreadsheetTimeZone(), 'yyyy/MM/dd HH:mm:ss');
      } else if (row[j] !== null && row[j] !== undefined) {
        cellValue = String(row[j]);
      }
      
      // ダブルクォートの二重化と、カンマ/クォート/改行を含む場合のエスケープ
      if (cellValue.includes('"') || cellValue.includes(',') || cellValue.includes('\n') || cellValue.includes('\r')) {
        cellValue = '"' + cellValue.replace(/"/g, '""') + '"';
      }
      csvRow.push(cellValue);
    }
    csvText += csvRow.join(',') + '\r\n';
  }
  return csvText;
}

/**
 * GitHub APIの認証情報およびリポジトリ・ブランチの接続テストを行う
 * 
 * @param {string} token APIトークン
 * @param {string} repo リポジトリパス (owner/repo)
 * @param {string} branch ブランチ名
 * @return {Object} 接続結果ステータス
 */
function testGitHubConnection(token, repo, branch) {
  try {
    const cleanRepo = (repo || 'awakuma1113/clinicaltrials').trim();
    const cleanBranch = (branch || 'main').trim();
    let cleanToken = (token || '').trim();
    
    // マスク文字列または空の場合は保存された設定からロード
    if (!cleanToken || cleanToken.startsWith('***') || cleanToken === '********') {
      cleanToken = PropertiesService.getUserProperties().getProperty('GITHUB_TOKEN') || '';
      if (!cleanToken) {
        return { success: false, message: '設定エラー: GitHub Personal Access Token (PAT) が入力されていません。' };
      }
    }
    
    // 1. リポジトリへのアクセス可否をチェック
    const url = 'https://api.github.com/repos/' + cleanRepo;
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + cleanToken,
        'Accept': 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28',
        'User-Agent': 'Mozilla/5.0 (Google Apps Script)'
      },
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    if (responseCode === 200) {
      // 2. ブランチの存在チェック
      const branchUrl = 'https://api.github.com/repos/' + cleanRepo + '/branches/' + cleanBranch;
      const branchResponse = UrlFetchApp.fetch(branchUrl, {
        method: 'GET',
        headers: {
          'Authorization': 'Bearer ' + cleanToken,
          'Accept': 'application/vnd.github+json',
          'X-GitHub-Api-Version': '2022-11-28',
          'User-Agent': 'Mozilla/5.0 (Google Apps Script)'
        },
        muteHttpExceptions: true
      });
      
      if (branchResponse.getResponseCode() === 200) {
        return {
          success: true,
          message: `リポジトリおよびブランチ「${cleanBranch}」への接続に成功しました！データを同期可能です。`
        };
      } else {
        return {
          success: false,
          message: `リポジトリにはアクセスできましたが、ブランチ「${cleanBranch}」が見つかりませんでした。(エラーコード: ${branchResponse.getResponseCode()})`
        };
      }
    } else if (responseCode === 401 || responseCode === 403) {
      return {
        success: false,
        message: '認証エラー：GitHub PATが無効か、十分なリポジトリ操作権限がありません。'
      };
    } else if (responseCode === 404) {
      return {
        success: false,
        message: `エラー：リポジトリ「${cleanRepo}」が見つかりません。プライベートリポジトリの場合は、PATのアクセス権限を確認してください。`
      };
    } else {
      return {
        success: false,
        message: `エラーが発生しました。GitHub APIステータスコード: ${responseCode}`
      };
    }
  } catch (e) {
    return {
      success: false,
      message: '接続テスト中に例外エラーが発生しました: ' + e.message
    };
  }
}

/**
 * GitHub API Helper: 指定ファイルのSHAハッシュを取得する (更新時のConflict回避用)
 */
function getGitHubFileSha(token, repo, branch, path) {
  const url = 'https://api.github.com/repos/' + repo + '/contents/' + path + '?ref=' + branch;
  const response = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Accept': 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'User-Agent': 'Mozilla/5.0 (Google Apps Script)'
    },
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() === 200) {
    const data = JSON.parse(response.getContentText());
    return data.sha;
  }
  return null;
}

/**
 * GitHub API Helper: 指定ファイルをコミット・プッシュする
 */
function commitFileToGitHub(token, repo, branch, path, contentText, commitMessage, existingSha) {
  const url = 'https://api.github.com/repos/' + repo + '/contents/' + path;
  
  // UTF-8文字コードを担保した上でBase64エンコード
  const base64Content = Utilities.base64Encode(contentText, Utilities.Charset.UTF_8);
  
  const payload = {
    message: commitMessage,
    content: base64Content,
    branch: branch
  };
  
  if (existingSha) {
    payload.sha = existingSha;
  }
  
  const response = UrlFetchApp.fetch(url, {
    method: 'PUT',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Accept': 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'Content-Type': 'application/json',
      'User-Agent': 'Mozilla/5.0 (Google Apps Script)'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  const responseCode = response.getResponseCode();
  if (responseCode === 200 || responseCode === 201) {
    const data = JSON.parse(response.getContentText());
    return {
      success: true,
      sha: data.content.sha,
      htmlUrl: data.commit.html_url
    };
  } else {
    const errText = response.getContentText();
    Logger.log(`GitHub Commit [${path}] Error: ${errText}`);
    return {
      success: false,
      message: `APIエラー ${responseCode} - ${errText}`
    };
  }
}

/**
 * GitHub API Helper: リポジトリが空（未初期化）かどうかをチェックする
 */
function checkIfRepositoryIsEmpty(token, repo) {
  const url = 'https://api.github.com/repos/' + repo + '/contents';
  const response = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Accept': 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'User-Agent': 'Mozilla/5.0 (Google Apps Script)'
    },
    muteHttpExceptions: true
  });
  
  const responseCode = response.getResponseCode();
  if (responseCode === 404) return true; // 空または未初期化リポジトリは通常404を返します
  if (responseCode === 200) {
    const list = JSON.parse(response.getContentText());
    return list.length === 0;
  }
  return true;
}

/**
 * 初回初期化用の美しく設計されたREADME.mdファイルコンテンツを生成する
 */
function getReadmeContent(repo, activeSheetName) {
  return `# 日本がん対策図鑑 臨床試験データベース (Gantaisaku Zukan Clinical Trial Database)

Google スプレッドシートから自動抽出・収集されたがん治療関連の臨床試験結果データベースです。
スプレッドシートの add-on ツール「がん対策図鑑取り込み」から直接 GitHub に同期されています。

## リポジトリ構成

本リポジトリには、スプレッドシートから同期された以下の CSV データファイルが格納されています。

| ファイル名 | 説明 | 主要なカラム |
| :--- | :--- | :--- |
| **\`clinical_trials.csv\`** | 主要な試験記事データ | タイトル、がん種、治療ライン、試験名、Phase、NCT番号、実施国、ジャーナル情報、要約等 |
| **\`results.csv\`** | 確認済みの構造化臨床試験データ | 試験群/対照群、評価項目（PFS/OS/ORR 等）、中央値、ハザード比（HR）、95%信頼区間、p値、有害事象率等 |
| **\`results_raw.csv\`** | 未確認の一次ソース抽出候補データ | 自動抽出された候補文、ハザード比（HR）、95%信頼区間、p値、確信度（High/Medium/Low）、確認ステータス等 |

## データ更新について

データの更新は、Google スプレッドシートのサイドバーから自動同期を実行することで反映されます。同期時にはスプレッドシートの最新の内容が CSV 形式に変換され、本リポジトリにコミットされます。

---
*Created by [日本がん対策図鑑 臨床試験抽出ツール](https://github.com/${repo})*
`;
}

/**
 * 臨床試験スプレッドシートの主要データをCSV形式でGitHubへ一括プッシュする
 * 
 * @param {string} token APIトークン
 * @param {string} repo リポジトリパス (owner/repo)
 * @param {string} branch ブランチ名
 * @param {string} [commitMessage] 任意のコミットメッセージ
 * @return {Object} 処理ステータスとログ詳細
 */
function pushCsvFilesToGitHub(token, repo, branch, commitMessage) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    let cleanToken = (token || '').trim();
    
    // マスク文字列または未入力の場合は保存されたPATをロード
    if (!cleanToken || cleanToken.startsWith('***') || cleanToken === '********') {
      cleanToken = userProperties.getProperty('GITHUB_TOKEN') || '';
    }
    
    if (!cleanToken) {
      return { success: false, message: 'エラー：GitHub Personal Access Token (PAT) が設定されていません。' };
    }
    
    const cleanRepo = (repo || userProperties.getProperty('GITHUB_REPO') || 'awakuma1113/clinicaltrials').trim();
    const cleanBranch = (branch || userProperties.getProperty('GITHUB_BRANCH') || 'main').trim();
    
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    const cleanMsg = (commitMessage || 'Spreadsheet data sync: ' + nowStr).trim();
    
    // 1. 各種シートをシリアライズ
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();
    const activeSheetName = activeSheet.getName();
    
    const activeCsvText = convertSheetToCsvText(activeSheetName);
    const resultsCsvText = convertSheetToCsvText('Results');
    const resultsRawCsvText = convertSheetToCsvText('Results_raw');
    const trialsCsvText = convertSheetToCsvText('Trials');
    const armsCsvText = convertSheetToCsvText('Arms');
    const outcomesCsvText = convertSheetToCsvText('Outcomes');
    
    const filesToPush = [];
    
    if (activeCsvText) {
      filesToPush.push({
        path: 'clinical_trials.csv',
        content: activeCsvText,
        description: `スプレッドシート主要データ (\`${activeSheetName}\` シート)`
      });
    }
    
    if (resultsCsvText) {
      filesToPush.push({
        path: 'results.csv',
        content: resultsCsvText,
        description: '確認済み構造化結果 (\`Results\` シート)'
      });
    }
    
    if (resultsRawCsvText) {
      filesToPush.push({
        path: 'results_raw.csv',
        content: resultsRawCsvText,
        description: '抽出候補生データ (\`Results_raw\` シート)'
      });
    }

    if (trialsCsvText) {
      filesToPush.push({
        path: 'trials.csv',
        content: trialsCsvText,
        description: '試験基本情報 (\`Trials\` シート)'
      });
    }

    if (armsCsvText) {
      filesToPush.push({
        path: 'arms.csv',
        content: armsCsvText,
        description: '試験アーム群 (\`Arms\` シート)'
      });
    }

    if (outcomesCsvText) {
      filesToPush.push({
        path: 'outcomes.csv',
        content: outcomesCsvText,
        description: '試験評価項目 (\`Outcomes\` シート)'
      });
    }
    
    if (filesToPush.length === 0) {
      return { success: false, message: 'エラー：同期可能なデータシートが存在しませんでした。' };
    }
    
    // 2. リポジトリが空の場合にREADME.mdを自動差し込み
    const isRepoEmpty = checkIfRepositoryIsEmpty(cleanToken, cleanRepo);
    if (isRepoEmpty) {
      const readmeContent = getReadmeContent(cleanRepo, activeSheetName);
      filesToPush.unshift({
        path: 'README.md',
        content: readmeContent,
        description: 'リポジトリREADME.md初期化'
      });
    }
    
    // 3. 各ファイルの順次コミット実行
    const logDetails = [];
    let allSuccessful = true;
    let finalCommitUrl = '';
    
    for (const file of filesToPush) {
      const existingSha = getGitHubFileSha(cleanToken, cleanRepo, cleanBranch, file.path);
      const commitRes = commitFileToGitHub(cleanToken, cleanRepo, cleanBranch, file.path, file.content, `Update ${file.path}: ${cleanMsg}`, existingSha);
      
      if (commitRes.success) {
        logDetails.push(`✅ ${file.description} のコミットに成功しました。`);
        finalCommitUrl = commitRes.htmlUrl;
      } else {
        allSuccessful = false;
        logDetails.push(`❌ ${file.description} のコミットに失敗しました: ${commitRes.message}`);
      }
    }
    
    // 正常同期できたら、入力された情報を設定として永続保存
    if (allSuccessful) {
      saveGitHubSettings(cleanToken, cleanRepo, cleanBranch);
    }
    
    return {
      success: allSuccessful,
      message: allSuccessful ? 'GitHubへのデータ同期がすべて完了しました！' : '一部のファイルの同期に失敗しました。詳細ログを確認してください。',
      details: logDetails,
      htmlUrl: finalCommitUrl || `https://github.com/${cleanRepo}`
    };
    
  } catch (e) {
    Logger.log('GitHubデータ同期エラー: ' + e.message);
    return {
      success: false,
      message: '同期処理中に例外エラーが発生しました: ' + e.message
    };
  }
}

/**
 * 現在入力されているNCT番号、または選択中の行のnct_id列からNCT IDを取得する
 */
function getNctIdFromActiveSelectionOrForm(sidebarNctId) {
  if (sidebarNctId && sidebarNctId.trim().toUpperCase().startsWith('NCT')) {
    return sidebarNctId.trim().toUpperCase();
  }
  
  // Otherwise check active row
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeCell = sheet.getActiveCell();
  const activeRowIndex = activeCell.getRow();
  
  if (activeRowIndex <= 1) {
    throw new Error('ヘッダー行以外の有効なデータ行を選択するか、サイドバーにNCT番号を入力してください。');
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const nctColIndex = headers.indexOf('nct_id');
  
  if (nctColIndex === -1) {
    throw new Error('現在のシートに「nct_id」列が存在しません。データ行を選択するか、サイドバーにNCT番号を入力してください。');
  }
  
  const nctIdVal = sheet.getRange(activeRowIndex, nctColIndex + 1).getValue();
  if (!nctIdVal || !String(nctIdVal).trim().toUpperCase().startsWith('NCT')) {
    throw new Error('選択した行の「nct_id」列に有効なNCT番号が見つかりません。');
  }
  
  return String(nctIdVal).trim().toUpperCase();
}

/**
 * シート選択行からNCT IDを読み取る (サイドバー連携インジケータ用)
 */
function getNctIdFromSheetSelectionOnly() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeCell = sheet.getActiveCell();
    const activeRowIndex = activeCell.getRow();
    if (activeRowIndex <= 1) return null;
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nctColIndex = headers.indexOf('nct_id');
    if (nctColIndex === -1) return null;
    
    const nctIdVal = sheet.getRange(activeRowIndex, nctColIndex + 1).getValue();
    if (nctIdVal && String(nctIdVal).trim().toUpperCase().startsWith('NCT')) {
      return { nctId: String(nctIdVal).trim().toUpperCase() };
    }
    return null;
  } catch (e) {
    return null;
  }
}

/**
 * ClinicalTrials.gov API v2経由で試験詳細情報を新規シート(Trials, Arms, Outcomes)およびResults_rawに取得・保存する
 * 
 * @param {string} sidebarNctId サイドバーに入力されているNCT ID
 * @return {Object} 処理ステータスと抽出された候補リスト
 */
function supplementFromClinicalTrialsApi(sidebarNctId) {
  try {
    // 1. NCT IDを取得 (入力値 or 選択行)
    const nctId = getNctIdFromActiveSelectionOrForm(sidebarNctId);
    
    // 2. API v2 からデータを取得
    const ctData = fetchClinicalTrialsInfo(nctId);
    if (!ctData) {
      throw new Error('ClinicalTrials.govからNCT ID「' + nctId + '」の情報が見つからないか、APIリクエストに失敗しました。');
    }
    
    const protocol = ctData.raw.protocolSection || {};
    const results = ctData.raw.resultsSection || null;
    
    // 3. Trials シートに保存
    saveToTrialsSheet(nctId, protocol);
    
    // 4. Arms シートに保存
    saveToArmsSheet(nctId, protocol);
    
    // 5. Outcomes シートに保存
    saveToOutcomesSheet(nctId, protocol);
    
    // 6. Results_raw シートに保存
    let candidates = [];
    if (results) {
      candidates = parseAndSaveResultsSection(nctId, results);
    }
    
    if (candidates.length === 0) {
      // 結果が無い場合は clinicaltrials_no_results として Results_raw に記録
      const now = new Date();
      const formattedTimestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
      
      const dummyCandidate = {
        imported_at: formattedTimestamp,
        article_url: '',
        trial_name: protocol.identificationModule ? (protocol.identificationModule.officialTitle || protocol.identificationModule.briefTitle || '') : '',
        nct_id: nctId,
        pmid: extractPmidFromProtocol(protocol),
        endpoint: 'All Outcomes',
        source_type: 'ClinicalTrials.gov',
        source_text: results ? 'ClinicalTrials.govに結果セクションはありますが、具体的な統計解析値はパースできませんでした。' : 'ClinicalTrials.govに結果情報が掲載されていません。',
        extracted_numbers: {},
        numeric_result_found: false,
        result_extraction_status: 'clinicaltrials_no_results',
        confidence: 'low',
        checked: '未確認',
        notes: results ? '結果掲載あり（詳細数値なし）' : '結果未掲載'
      };
      
      saveRawCandidatesToSheet([dummyCandidate], 'https://clinicaltrials.gov/study/' + nctId);
      candidates.push(dummyCandidate);
    }
    
    return {
      nct_id: nctId,
      trial_name: protocol.identificationModule ? (protocol.identificationModule.officialTitle || protocol.identificationModule.briefTitle || '') : '',
      candidates: candidates
    };
    
  } catch (e) {
    Logger.log('ClinicalTrials API v2補完エラー: ' + e.message);
    throw new Error('ClinicalTrials.govからの補完中にエラーが発生しました: ' + e.message);
  }
}

/**
 * Trials シートに試験基本情報を保存する (重複時は更新)
 */
function saveToTrialsSheet(nctId, protocol) {
  initResultsSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trialsSheet = ss.getSheetByName('Trials');
  
  const lastRow = trialsSheet.getLastRow();
  let existingRow = -1;
  if (lastRow > 1) {
    const ids = trialsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === nctId) {
        existingRow = i + 2;
        break;
      }
    }
  }
  
  const idModule = protocol.identificationModule || {};
  const statusModule = protocol.statusModule || {};
  const designModule = protocol.designModule || {};
  const condsModule = protocol.conditionsModule || {};
  const armsModule = protocol.armsInterventionsModule || {};
  
  const briefTitle = idModule.briefTitle || '';
  const officialTitle = idModule.officialTitle || '';
  const phases = designModule.phases ? designModule.phases.join(', ') : '';
  const studyType = designModule.studyType || '';
  const overallStatus = statusModule.overallStatus || '';
  const conditions = condsModule.conditions ? condsModule.conditions.join(', ') : '';
  
  const interventions = armsModule.interventions ? 
    armsModule.interventions.map(i => (i.type || '') + ': ' + (i.name || '')).join(', ') : '';
  
  const enrollmentCount = designModule.enrollmentInfo ? (designModule.enrollmentInfo.count || '') : '';
  const startDate = statusModule.startDateStruct ? (statusModule.startDateStruct.date || '') : '';
  const completionDate = statusModule.completionDateStruct ? (statusModule.completionDateStruct.date || '') : '';
  const primaryCompletionDate = statusModule.primaryCompletionDateStruct ? (statusModule.primaryCompletionDateStruct.date || '') : '';
  
  const rowValues = [
    nctId,
    briefTitle,
    officialTitle,
    phases,
    studyType,
    overallStatus,
    conditions,
    interventions,
    enrollmentCount,
    startDate,
    completionDate,
    primaryCompletionDate
  ];
  
  if (existingRow !== -1) {
    trialsSheet.getRange(existingRow, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    trialsSheet.appendRow(rowValues);
  }
}

/**
 * Arms シートにアーム群情報を保存する (既存定義は削除後追加)
 */
function saveToArmsSheet(nctId, protocol) {
  initResultsSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const armsSheet = ss.getSheetByName('Arms');
  
  const armsModule = protocol.armsInterventionsModule || {};
  const armGroups = armsModule.armGroups || [];
  
  if (armGroups.length === 0) return;
  
  // 既存アームを削除
  const lastRow = armsSheet.getLastRow();
  if (lastRow > 1) {
    const ids = armsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) {
      if (ids[i][0] === nctId) {
        armsSheet.deleteRow(i + 2);
      }
    }
  }
  
  for (const arm of armGroups) {
    const label = arm.label || '';
    const type = arm.type || '';
    const description = arm.description || '';
    const interventions = arm.interventionNames ? arm.interventionNames.join(', ') : '';
    
    armsSheet.appendRow([
      nctId,
      label,
      type,
      description,
      interventions
    ]);
  }
}

/**
 * Outcomes シートに主要/副次評価項目定義を保存する (既存定義は削除後追加)
 */
function saveToOutcomesSheet(nctId, protocol) {
  initResultsSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outcomesSheet = ss.getSheetByName('Outcomes');
  
  const outcomesModule = protocol.outcomesModule || {};
  const primary = outcomesModule.primaryOutcomes || [];
  const secondary = outcomesModule.secondaryOutcomes || [];
  const other = outcomesModule.otherOutcomes || [];
  
  if (primary.length === 0 && secondary.length === 0 && other.length === 0) return;
  
  // 既存アウトカムを削除
  const lastRow = outcomesSheet.getLastRow();
  if (lastRow > 1) {
    const ids = outcomesSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) {
      if (ids[i][0] === nctId) {
        outcomesSheet.deleteRow(i + 2);
      }
    }
  }
  
  const outcomesToAppend = [];
  
  primary.forEach(o => {
    outcomesToAppend.push([nctId, 'Primary', o.measure || '', o.description || '', o.timeFrame || '']);
  });
  secondary.forEach(o => {
    outcomesToAppend.push([nctId, 'Secondary', o.measure || '', o.description || '', o.timeFrame || '']);
  });
  other.forEach(o => {
    outcomesToAppend.push([nctId, 'Other', o.measure || '', o.description || '', o.timeFrame || '']);
  });
  
  for (const row of outcomesToAppend) {
    outcomesSheet.appendRow(row);
  }
}

/**
 * ClinicalTrials.gov API v2 結果セクションから統計結果をパースして Results_raw に保存
 */
function parseAndSaveResultsSection(nctId, results) {
  const outcomes = results.outcomeMeasuresModule ? (results.outcomeMeasuresModule.outcomeMeasures || []) : [];
  const candidates = [];
  
  const now = new Date();
  const formattedTimestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  const sourceUrl = 'https://clinicaltrials.gov/study/' + nctId;
  
  for (const measure of outcomes) {
    const title = measure.title || '';
    const description = measure.description || '';
    const endpoint = TrialResultExtractor.estimateEndpoint(title);
    
    const analyses = measure.analyses || [];
    for (const analysis of analyses) {
      const pValue = analysis.pValue || '';
      const paramType = analysis.paramType || '';
      const paramValue = analysis.paramValue || '';
      const ciLower = analysis.ciLowerLimit || '';
      const ciUpper = analysis.ciUpperLimit || '';
      
      if (!paramValue && !pValue) continue;
      
      const sourceText = `Outcome Measure: ${title}. Statistical Analysis: ${paramType} = ${paramValue}` + 
                         (ciLower || ciUpper ? ` (95% CI: ${ciLower}-${ciUpper})` : '') + 
                         (pValue ? `, p = ${pValue}` : '') + 
                         (analysis.statisticalMethod ? ` [Method: ${analysis.statisticalMethod}]` : '');
      
      let hr = '';
      if (paramType.toLowerCase().includes('hazard ratio') || paramType.toLowerCase().includes('hr')) {
        hr = paramValue;
      }
      
      const numbers = {
        hr: hr,
        ci_lower: ciLower,
        ci_upper: ciUpper,
        p_value: pValue,
        medians: [],
        percentages: []
      };
      
      const classes = measure.classes || [];
      const medians = [];
      const percentages = [];
      classes.forEach(cls => {
        const categories = cls.categories || [];
        categories.forEach(cat => {
          const measurements = cat.measurements || [];
          measurements.forEach(m => {
            if (m.value !== undefined) {
              if (measure.unitOfMeasure && measure.unitOfMeasure.toLowerCase().includes('month')) {
                medians.push(m.value);
              } else if (measure.unitOfMeasure && measure.unitOfMeasure.includes('%')) {
                percentages.push(m.value);
              }
            }
          });
        });
      });
      
      if (medians.length > 0) numbers.medians = medians;
      if (percentages.length > 0) numbers.percentages = percentages;
      
      candidates.push({
        imported_at: formattedTimestamp,
        article_url: '',
        trial_name: '',
        nct_id: nctId,
        pmid: '',
        detected_endpoint: endpoint,
        detected_result_text: sourceText,
        source_type: 'ClinicalTrials.gov',
        source_text: sourceText,
        detected_numbers: numbers,
        numeric_result_found: true,
        result_extraction_status: 'clinicaltrials_found',
        confidence: 'high',
        checked: '未確認',
        notes: 'ClinicalTrials.gov API 結果セクションより自動抽出'
      });
    }
  }
  
  if (candidates.length > 0) {
    saveRawCandidatesToSheet(candidates, sourceUrl);
  }
  
  return candidates;
}

/**
 * ClinicalTrials.gov プロトコルからPMIDを抽出
 */
function extractPmidFromProtocol(protocol) {
  const references = protocol.referencesModule ? (protocol.referencesModule.references || []) : [];
  for (const ref of references) {
    if (ref.type === 'RESULT' || ref.type === 'BACKGROUND') {
      if (ref.pmid) return ref.pmid;
    }
  }
  return '';
}
