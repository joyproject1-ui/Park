/**
 * FDA Warning Letter 주간 알림 — Google Apps Script 판
 *
 * 식약처 행정처분 메일을 보내는 것과 같은 구글 계정에서 돌리기 위한 것입니다.
 * MailApp 이 계정 권한으로 직접 보내므로 앱 비밀번호나 SMTP 설정이 필요 없습니다.
 *
 * 설치:
 *   1) 식약처 스크립트가 있는 Apps Script 프로젝트에 이 파일을 추가
 *   2) installWeeklyTrigger() 를 한 번 실행 (권한 승인 요청이 뜨면 허용)
 *   3) previewFdaWarningLetters() 로 메일 없이 결과만 먼저 확인
 *
 * 설정은 프로젝트 설정 > 스크립트 속성에서 바꿉니다 (없으면 아래 기본값):
 *   MAIL_TO        수신자, 쉼표 구분
 *   SINCE_DAYS     최근 며칠 (기본 7)
 *   STERILE_ONLY   'true' 면 무균·주사제·점안제 관련 건만
 */

var DEFAULTS = {
  MAIL_TO: 'jhp5408@hanlim.com',
  SINCE_DAYS: '7',
  STERILE_ONLY: 'false'
};

var LANDING_PAGE =
  'https://www.fda.gov/inspections-compliance-enforcement-and-criminal-investigations' +
  '/compliance-actions-and-activities/warning-letters';
var RSS_URL =
  'https://www.fda.gov/about-fda/contact-fda/stay-informed/rss-feeds/warning-letters/rss.xml';
var DATATABLES_URL =
  'https://www.fda.gov/datatables/views/data.json' +
  '?total_count_needed=true&view_display_id=warning_letter_solr_block' +
  '&view_name=warning_letter_solr_index&length=100&start=0';

var STATE_KEY = 'FDA_WL_NOTIFIED';
var STATE_LIMIT = 300;          // 스크립트 속성 용량 한도(9KB) 때문에 오래된 것부터 버립니다.
var MAX_DETAIL_FETCHES = 30;    // 본문을 열어볼 최대 건수
var TIME_BUDGET_MS = 4 * 60 * 1000;  // 실행 제한 6분 전에 스스로 멈추기 위한 여유

// 파이썬 판(gmpai/warningletters.py)과 같은 키워드를 씁니다. 한쪽만 고치지 마세요.
var STERILE_KEYWORDS = {
  'sterile': '무균', 'sterility': '무균성', 'aseptic': '무균조작',
  'ophthalmic': '점안제', 'eye drop': '점안제', 'injectable': '주사제',
  'injection': '주사제', 'parenteral': '주사제', 'media fill': '미디어필',
  'environmental monitoring': '환경모니터링', 'endotoxin': '엔도톡신',
  'pyrogen': '발열성물질', 'particulate': '이물', 'isolator': '아이솔레이터',
  'rabs': 'RABS', 'smoke stud': '기류시험', '503b': '503B 조제시설',
  'outsourcing facility': '503B 조제시설'
};

var CGMP_KEYWORDS = {
  'data integrity': '데이터 완전성', 'out-of-specification': 'OOS',
  'out of specification': 'OOS', 'cgmp': 'CGMP', 'quality unit': '품질부서',
  'stability': '안정성시험', 'cleaning validation': '세척밸리데이션',
  'process validation': '공정밸리데이션', 'contamination': '오염',
  'recall': '회수', 'audit trail': '감사추적'
};

// --------------------------------------------------------------- 진입점

/** 주 1회 트리거가 부르는 함수. */
function sendFdaWarningLetters() {
  return run_(false);
}

/** 메일을 보내지 않고 실행 기록에만 남깁니다. 첫 확인용. */
function previewFdaWarningLetters() {
  return run_(true);
}

function run_(dryRun) {
  var cfg = config_();
  var collected = collect_();
  if (!collected.letters.length) {
    Logger.log('수집 실패 — ' + collected.problems.join(' / '));
    return 0;
  }
  Logger.log('[' + collected.source + '] 목록 ' + collected.letters.length + '건');

  var letters = withinDays_(collected.letters, cfg.sinceDays);
  Logger.log('최근 ' + cfg.sinceDays + '일 해당: ' + letters.length + '건');
  if (!letters.length) return 0;

  letters = enrich_(letters);
  letters = letters.filter(function (l) {
    return cfg.sterileOnly ? l.sterileHits.length : (l.sterileHits.length || l.cgmpHits.length);
  });
  Logger.log('관련 건 선별 후: ' + letters.length + '건');

  var state = loadState_();
  var fresh = sortLetters_(letters.filter(function (l) { return !state[keyOf_(l)]; }));
  Logger.log('신규: ' + fresh.length + '건 (중복 제외 ' + (letters.length - fresh.length) + '건)');
  if (!fresh.length) return 0;

  fresh.forEach(function (l) {
    Logger.log((l.sterileHits.length ? '[무균] ' : '[CGMP] ') + l.company +
               ' | ' + (l.letterDate || '-') + ' | ' + l.url);
  });

  if (dryRun) {
    Logger.log('dry-run — 메일을 보내지 않았습니다.');
    return fresh.length;
  }

  MailApp.sendEmail({
    to: cfg.recipients.join(','),
    subject: '[FDA] 신규 Warning Letter ' + fresh.length + '건',
    body: renderText_(fresh, sinceDate_(cfg.sinceDays)),
    htmlBody: renderHtml_(fresh, sinceDate_(cfg.sinceDays), collected.source)
  });
  saveState_(state, fresh);
  Logger.log('발송 완료 → ' + cfg.recipients.join(', '));
  return fresh.length;
}

function config_() {
  var props = PropertiesService.getScriptProperties();
  function get(name) { return props.getProperty(name) || DEFAULTS[name]; }
  return {
    recipients: get('MAIL_TO').split(',').map(function (s) { return s.trim(); })
                              .filter(function (s) { return s; }),
    sinceDays: parseInt(get('SINCE_DAYS'), 10) || 7,
    sterileOnly: String(get('STERILE_ONLY')).toLowerCase() === 'true'
  };
}

// --------------------------------------------------------------- 수집

function fetchText_(url) {
  var response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { 'User-Agent': 'Mozilla/5.0 gmp-ai-regs/1.0', 'Accept-Language': 'en' }
  });
  var code = response.getResponseCode();
  if (code < 200 || code >= 300) throw new Error('HTTP ' + code);
  return response.getContentText();
}

/** RSS → 목록 JSON 순으로 시도해 먼저 성공한 결과를 씁니다. */
function collect_() {
  var sources = [
    { name: 'rss', url: RSS_URL, parse: parseRss_ },
    { name: 'datatables', url: DATATABLES_URL, parse: parseDatatables_ }
  ];
  var problems = [];
  for (var i = 0; i < sources.length; i++) {
    try {
      var letters = sources[i].parse(fetchText_(sources[i].url));
      if (letters.length) return { letters: letters, source: sources[i].name, problems: problems };
      problems.push(sources[i].name + ': 항목 0건');
    } catch (err) {
      problems.push(sources[i].name + ': ' + err.message);
    }
  }
  return { letters: [], source: '', problems: problems };
}

function parseRss_(xml) {
  var channel = XmlService.parse(xml).getRootElement().getChild('channel');
  if (!channel) throw new Error('channel 요소가 없습니다');
  return channel.getChildren('item').map(function (item) {
    var title = (item.getChildText('title') || '').trim();
    var parts = title.split(' - ').map(function (p) { return p.trim(); })
                     .filter(function (p) { return p; });
    var letterDate = '';
    for (var i = parts.length - 1; i >= 1 && !letterDate; i--) letterDate = parseDate_(parts[i]);
    return newLetter_({
      company: parts.length ? parts[0] : title,
      url: (item.getChildText('link') || '').trim(),
      letterDate: letterDate,
      postedDate: parseDate_(item.getChildText('pubDate') || ''),
      subject: (item.getChildText('description') || '').trim().slice(0, 300)
    });
  }).filter(function (l) { return l.url; });
}

function parseDatatables_(text) {
  var payload = JSON.parse(text);
  var rows = payload && payload.data ? payload.data : payload;
  if (!(rows instanceof Array)) throw new Error('data 배열이 없습니다');
  return rows.map(function (row) {
    var rawCompany = findRaw_(row, ['company', 'legal_name', 'title']);
    return newLetter_({
      company: stripTags_(rawCompany),
      url: extractHref_(rawCompany) || extractHref_(findRaw_(row, ['link', 'url'])),
      letterDate: parseDate_(pick_(row, ['issue_date', 'letter_issue', 'issue'])),
      postedDate: parseDate_(pick_(row, ['change_date', 'posted', 'update'])),
      office: pick_(row, ['issuing_office', 'office']),
      subject: pick_(row, ['subject'])
    });
  }).filter(function (l) { return l.company; });
}

function findRaw_(row, needles) {
  for (var n = 0; n < needles.length; n++) {
    for (var key in row) {
      if (key.toLowerCase().indexOf(needles[n]) >= 0 &&
          typeof row[key] === 'string' && row[key].trim()) return row[key];
    }
  }
  return '';
}

function pick_(row, needles) { return stripTags_(findRaw_(row, needles)); }

function stripTags_(value) {
  return String(value || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

function extractHref_(value) {
  var match = /href="([^"]+)"/.exec(String(value || ''));
  if (!match) return '';
  return match[1].indexOf('http') === 0 ? match[1] : 'https://www.fda.gov' + match[1];
}

function newLetter_(fields) {
  return {
    company: fields.company || '', url: fields.url || '',
    letterDate: fields.letterDate || '', postedDate: fields.postedDate || '',
    office: fields.office || '', subject: fields.subject || '',
    sterileHits: [], cgmpHits: [], cfr: []
  };
}

function keyOf_(letter) { return letter.url || (letter.company + '|' + letter.letterDate); }
function bestDate_(letter) { return letter.postedDate || letter.letterDate; }

// --------------------------------------------------------------- 날짜

var MONTHS = {
  january: 1, february: 2, march: 3, april: 4, may: 5, june: 6, july: 7,
  august: 8, september: 9, october: 10, november: 11, december: 12,
  jan: 1, feb: 2, mar: 3, apr: 4, jun: 6, jul: 7, aug: 8, sep: 9, sept: 9,
  oct: 10, nov: 11, dec: 12
};

function pad_(n) { return (n < 10 ? '0' : '') + n; }

/** 여러 표기의 날짜를 YYYY-MM-DD 로. 읽지 못하면 빈 문자열. */
function parseDate_(value) {
  var text = String(value || '').trim();
  if (!text) return '';

  var iso = /\d{4}-\d{2}-\d{2}/.exec(text);
  if (iso) return iso[0];

  var slash = /(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(text);
  if (slash) return slash[3] + '-' + pad_(+slash[1]) + '-' + pad_(+slash[2]);

  var named = /([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s*(\d{4})/.exec(text);
  if (named) {
    var month = MONTHS[named[1].toLowerCase()];
    if (month) return named[3] + '-' + pad_(month) + '-' + pad_(+named[2]);
  }

  var parsed = new Date(text);   // RFC822 pubDate 등
  if (!isNaN(parsed.getTime())) {
    return parsed.getUTCFullYear() + '-' + pad_(parsed.getUTCMonth() + 1) + '-' +
           pad_(parsed.getUTCDate());
  }
  return '';
}

function sinceDate_(days) {
  var d = new Date(Date.now() - days * 86400000);
  return d.getUTCFullYear() + '-' + pad_(d.getUTCMonth() + 1) + '-' + pad_(d.getUTCDate());
}

// --------------------------------------------------------------- 선별

/** 지난 N일 이내. 날짜를 못 읽은 건은 판단할 수 없으므로 남깁니다. */
function withinDays_(letters, days) {
  var cutoff = sinceDate_(days);
  return letters.filter(function (l) {
    var date = bestDate_(l);
    return !date || date >= cutoff;
  });
}

/** 본문을 열어 키워드와 21 CFR 인용을 채웁니다. 시간·건수 예산 안에서만. */
function enrich_(letters) {
  var deadline = Date.now() + TIME_BUDGET_MS;
  var fetched = 0;
  letters.forEach(function (letter) {
    var text = letter.company + ' ' + letter.subject;
    if (letter.url && fetched < MAX_DETAIL_FETCHES && Date.now() < deadline) {
      try {
        var body = fetchText_(letter.url);
        text = stripTags_(body);
        fetched++;
        if (!letter.office) {
          var office = /(?:Division of|Office of)[^.<\n]{3,80}/.exec(body);
          if (office) letter.office = stripTags_(office[0]);
        }
      } catch (err) {
        Logger.log('본문 실패 ' + letter.company + ': ' + err.message);
      }
    }
    classify_(letter, text);
  });
  return letters;
}

function classify_(letter, text) {
  var lowered = String(text).toLowerCase();
  letter.sterileHits = matchKeywords_(lowered, STERILE_KEYWORDS);
  letter.cgmpHits = matchKeywords_(lowered, CGMP_KEYWORDS);
  letter.cfr = uniqueSorted_((String(text).match(/21\s*CFR\s*\d{3}\.\d+(?:\([a-z0-9]+\))*/gi) || [])
    .map(function (m) { return m.replace(/^21\s*CFR\s*/i, ''); }));
  return letter;
}

function matchKeywords_(lowered, table) {
  var hits = [];
  for (var english in table) {
    if (lowered.indexOf(english) >= 0 && hits.indexOf(table[english]) < 0) hits.push(table[english]);
  }
  return hits.sort();
}

function uniqueSorted_(values) {
  return values.filter(function (v, i) { return values.indexOf(v) === i; }).sort();
}

/** 무균 관련 건을 위로, 그 다음 최신순. */
function sortLetters_(letters) {
  return letters.slice().sort(function (a, b) {
    var sterile = (b.sterileHits.length ? 1 : 0) - (a.sterileHits.length ? 1 : 0);
    if (sterile) return sterile;
    var dateA = bestDate_(a) || '0', dateB = bestDate_(b) || '0';
    if (dateA !== dateB) return dateA < dateB ? 1 : -1;
    return a.company < b.company ? -1 : (a.company > b.company ? 1 : 0);
  });
}

// --------------------------------------------------------------- 발송 이력

function loadState_() {
  var raw = PropertiesService.getScriptProperties().getProperty(STATE_KEY);
  if (!raw) return {};
  try { return JSON.parse(raw) || {}; } catch (err) { return {}; }
}

function saveState_(state, letters) {
  var stamp = new Date().toISOString();
  letters.forEach(function (l) { state[keyOf_(l)] = { date: bestDate_(l), at: stamp }; });

  // 속성 하나당 9KB 한도가 있으므로 오래된 기록부터 버립니다.
  var keys = Object.keys(state);
  if (keys.length > STATE_LIMIT) {
    keys.sort(function (a, b) {
      return String(state[a].at || '') < String(state[b].at || '') ? -1 : 1;
    }).slice(0, keys.length - STATE_LIMIT).forEach(function (k) { delete state[k]; });
  }
  PropertiesService.getScriptProperties().setProperty(STATE_KEY, JSON.stringify(state));
}

/** 발송 이력을 지웁니다. 지난 건을 다시 받아보고 싶을 때만 쓰세요. */
function resetNotifiedState() {
  PropertiesService.getScriptProperties().deleteProperty(STATE_KEY);
  Logger.log('발송 이력을 지웠습니다.');
}

// --------------------------------------------------------------- 본문

function esc_(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

var TH_STYLE = 'border:1px solid #d0d7de;background:#f2f5f8;padding:7px 10px;' +
               'text-align:left;width:140px;vertical-align:top;white-space:nowrap;';
var TD_STYLE = 'border:1px solid #d0d7de;padding:7px 10px;vertical-align:top;line-height:1.55;';
var TABLE_STYLE = 'border-collapse:collapse;width:100%;margin:4px 0 18px 0;font-size:13px;';
var SECTION_STYLE = 'color:#1f5f8b;font-size:13px;font-weight:bold;margin:12px 0 2px 0;';

function row_(label, value, raw) {
  var shown = raw ? value : esc_(value);
  return '<tr><th style="' + TH_STYLE + '">' + esc_(label) + '</th>' +
         '<td style="' + TD_STYLE + '">' + (shown || '-') + '</td></tr>';
}

function renderHtml_(letters, since, source) {
  var sterile = letters.filter(function (l) { return l.sterileHits.length; }).length;
  var html =
    '<html><body style="font-family:\'Malgun Gothic\',AppleGothic,sans-serif;color:#222;">' +
    '<h2 style="font-size:17px;margin:0 0 4px 0;">FDA Warning Letter 신규 ' + letters.length +
    '건 (' + since.replace(/-/g, '') + ' 이후)</h2>' +
    '<p style="font-size:13px;color:#444;margin:0 0 14px 0;">무균·주사제·점안제 관련 <b>' +
    sterile + '건</b> 포함. 분류는 본문 키워드 자동 판정이므로 원문 확인이 필요합니다.</p>';

  letters.forEach(function (letter, index) {
    var tag = letter.sterileHits.length ? '🔴 무균 관련' : '· CGMP';
    var link = letter.url
      ? '<a href="' + esc_(letter.url) + '">' + esc_(letter.url) + '</a>' : '-';
    html +=
      '<h3 style="font-size:15px;margin:26px 0 6px 0;color:#111;">' + (index + 1) + '. ' +
      esc_(letter.company) +
      ' <span style="font-weight:normal;color:#666;font-size:12px;">(' + tag + ')</span></h3>' +
      '<div style="' + SECTION_STYLE + '">■ 기본정보</div>' +
      '<table style="' + TABLE_STYLE + '">' +
        row_('업체명', letter.company) +
        row_('발행 사무소', letter.office) +
        row_('서한일자', letter.letterDate) +
        row_('게시일자', letter.postedDate) +
        row_('원문 링크', link, true) +
      '</table>' +
      '<div style="' + SECTION_STYLE + '">■ 지적 관련성</div>' +
      '<table style="' + TABLE_STYLE + '">' +
        row_('해당 키워드', letter.sterileHits.concat(letter.cgmpHits).join(', ')) +
        row_('인용 조항', letter.cfr.map(function (c) { return '21 CFR ' + c; }).join(', ')) +
      '</table>';
  });

  return html +
    '<hr style="border:none;border-top:1px solid #e1e4e8;margin:26px 0 10px 0;">' +
    '<p style="font-size:11px;color:#888;line-height:1.6;">출처: <a href="' + LANDING_PAGE +
    '">FDA Warning Letters</a>' + (source ? ' (수집 경로: ' + esc_(source) + ')' : '') +
    '<br>생성: ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm') +
    ' KST · Apps Script</p></body></html>';
}

function renderText_(letters, since) {
  var lines = ['FDA Warning Letter 신규 ' + letters.length + '건 (' + since + ' 이후)', ''];
  letters.forEach(function (letter, index) {
    lines.push((index + 1) + '. ' + (letter.sterileHits.length ? '[무균]' : '[CGMP]') + ' ' +
               letter.company);
    lines.push('   서한일자: ' + (letter.letterDate || '-') +
               ' / 게시일자: ' + (letter.postedDate || '-'));
    lines.push('   사무소: ' + (letter.office || '-'));
    lines.push('   키워드: ' + (letter.sterileHits.concat(letter.cgmpHits).join(', ') || '-'));
    lines.push('   조항: ' +
      (letter.cfr.map(function (c) { return '21 CFR ' + c; }).join(', ') || '-'));
    lines.push('   원문: ' + letter.url);
    lines.push('');
  });
  return lines.join('\n');
}

// --------------------------------------------------------------- 트리거

/**
 * 매주 수요일 오전 9시에 실행되도록 겁니다.
 * FDA 가 경고장 목록을 통상 화요일(미 동부시간)에 갱신하기 때문입니다.
 * 시각은 프로젝트 설정의 시간대를 따르므로 Asia/Seoul 인지 확인하세요.
 */
function installWeeklyTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('sendFdaWarningLetters')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
    .atHour(9)
    .create();
  Logger.log('매주 수요일 09시 트리거를 걸었습니다.');
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'sendFdaWarningLetters') ScriptApp.deleteTrigger(trigger);
  });
}
