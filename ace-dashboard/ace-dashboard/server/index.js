/**
 * ACE 學習者圖表報告系統 - OOP 架構後端
 * @version 2.0.0
 */

const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');
const compression = require('compression');
const NodeCache = require('node-cache');
const path = require('path');

// ==================== 工具類別 ====================

class CacheManager {
  constructor(ttl = 3600) {
    this.cache = new NodeCache({ stdTTL: ttl, checkperiod: 600 });
  }
  get(key) { return this.cache.get(key); }
  set(key, value) { return this.cache.set(key, value); }
  flush() { return this.cache.flushAll(); }
}

class DateUtils {
  static parseMonthKey(dateValue) {
    if (!dateValue) return null;
    let d;
    if (dateValue instanceof Date) {
      d = dateValue;
    } else if (typeof dateValue === 'number') {
      d = new Date(new Date(1899, 11, 30).getTime() + dateValue * 86400000);
    } else if (typeof dateValue === 'string') {
      const match = dateValue.match(/Date\((\d+),\s*(\d+),\s*(\d+)\)/);
      d = match ? new Date(+match[1], +match[2], +match[3]) : new Date(dateValue);
    } else return null;
    if (isNaN(d.getTime())) return null;
    return `${d.getFullYear()}-${(d.getMonth() + 1).toString().padStart(2, '0')}`;
  }

  static getMinMonth() {
    const now = new Date();
    const min = new Date(now.getFullYear() - 7, now.getMonth(), 1);
    return `${min.getFullYear()}-${(min.getMonth() + 1).toString().padStart(2, '0')}`;
  }

  // 月份轉學期 (民國年-1 或 民國年-2)
  static monthToSemester(monthKey) {
    const [year, month] = monthKey.split('-').map(Number);
    if (month >= 8) return `${year - 1911}-1`;
    if (month === 1) return `${year - 1912}-1`;
    return `${year - 1912}-2`;
  }

  // 學期轉月份陣列
  static semesterToMonths(semesterKey) {
    const [rocYear, sem] = semesterKey.split('-').map(Number);
    const y = rocYear + 1911;
    if (sem === 1) {
      return [`${y}-08`, `${y}-09`, `${y}-10`, `${y}-11`, `${y}-12`, `${y + 1}-01`];
    }
    return [`${y + 1}-02`, `${y + 1}-03`, `${y + 1}-04`, `${y + 1}-05`, `${y + 1}-06`, `${y + 1}-07`];
  }

  // 月份轉學年度
  static monthToYear(monthKey) {
    const [year, month] = monthKey.split('-').map(Number);
    return (month >= 8 ? year - 1911 : year - 1912).toString();
  }

  // 學年度轉月份陣列
  static yearToMonths(rocYear) {
    const y = parseInt(rocYear) + 1911;
    return [
      `${y}-08`, `${y}-09`, `${y}-10`, `${y}-11`, `${y}-12`,
      `${y + 1}-01`, `${y + 1}-02`, `${y + 1}-03`, `${y + 1}-04`, `${y + 1}-05`, `${y + 1}-06`, `${y + 1}-07`
    ];
  }
}

class Validator {
  static sanitize(input) {
    if (typeof input !== 'string') return input;
    return input.replace(/[<>]/g, '').replace(/javascript:/gi, '').trim().slice(0, 1000);
  }
  static isValidChart(type) {
    return /^chart(0|1|2|3|4|5|6|7|8|9|10|11)$/.test(type);
  }
  static isValidDataType(type) {
    return ['new', 'cumulative'].includes(type);
  }
}

// ==================== 分類器類別 ====================

class CollegeClassifier {
  static categories = [
    '醫學院', '生農學院', '工學院', '文學院', '理學院', '公衛學院',
    '共教學院', '社科學院', '電資學院', '創新學院', '生科學院',
    '法學院', '管理學院', '國際學院', '重點科技學院', '進修推廣學院', '其他'
  ];

  static classify(str) {
    if (!str) return '其他';
    const s = String(str);
    const map = {
      '醫學': '醫學院', '生物資源': '生農學院', '農學': '生農學院', '生農': '生農學院',
      '工學院': '工學院', '文學院': '文學院', '理學院': '理學院',
      '公共衛生': '公衛學院', '公衛': '公衛學院', '共同教育': '共教學院', '共同': '共教學院',
      '社會科學': '社科學院', '社科': '社科學院', '電機資訊': '電資學院', '電資': '電資學院',
      '創新設計': '創新學院', '創新': '創新學院', '生命科學': '生科學院', '生科': '生科學院',
      '法學院': '法學院', '管理學院': '管理學院', '國際學院': '國際學院', '國際': '國際學院',
      '重點科技': '重點科技學院', '進修推廣': '進修推廣學院', '推廣': '進修推廣學院'
    };
    for (const [k, v] of Object.entries(map)) {
      if (s.includes(k)) return v;
    }
    return '其他';
  }
}

class TeacherClassifier {
  static getType(job) {
    if (!job) return null;
    const s = String(job);
    if (s.includes('臨床') || s.includes('Clinical')) return '臨床教師';
    if (s.includes('專案') || s.includes('Project')) return '專案教師';
    if (s.includes('兼任') || s.includes('Adjunct')) return '兼任教師';
    if (['教授 Professor', '副教授 Associate Professor', '助理教授 Assistant Professor', '講師 Lecturer'].includes(s.trim())) {
      return '專任教師';
    }
    return null;
  }

  static getRank(job) {
    if (!job) return null;
    const s = String(job);
    if (s.includes('助理教授') || s.includes('Assistant Professor')) return '助理教授';
    if (s.includes('副教授') || s.includes('Associate Professor')) return '副教授';
    if (s.includes('教授') || s.includes('Professor')) return '教授';
    if (s.includes('講師') || s.includes('Lecturer')) return '講師';
    return null;
  }

  static fullTimeMap = {
    '教授 Professor': '教授', '副教授 Associate Professor': '副教授',
    '助理教授 Assistant Professor': '助理教授', '講師 Lecturer': '講師'
  };
  static adjunctMap = {
    '兼任教授 Adjunct Professor': '教授', '兼任副教授 Adjunct Associate Professor': '副教授',
    '兼任助理教授 Adjunct Assistant Professor': '助理教授', '兼任講師 Adjunct Lecturer': '講師'
  };
  static clinicalMap = {
    '臨床教授 Clinical Professor': '教授', '臨床副教授 Clinical Associate Professor': '副教授',
    '臨床助理教授 Clinical Assistant Professor': '助理教授', '臨床講師 Clinical Lecturer': '講師'
  };
}

class StudentClassifier {
  static levels = ['大學生', '研究生', '博士生'];

  static getLevel(identity, level) {
    const s = String(identity || '') + String(level || '');
    if (s.includes('博士') || s.includes('PhD') || s.includes('Doctoral')) return '博士生';
    if (s.includes('碩士') || s.includes('研究') || s.includes('Master') || s.includes('Graduate')) return '研究生';
    if (s.includes('大學') || s.includes('學士') || s.includes('Undergraduate') || s.includes('Bachelor') || s.includes('學生')) return '大學生';
    return null;
  }
}

// ==================== Google Sheets 服務 ====================

class GoogleSheetsService {
  constructor(spreadsheetId, cache) {
    this.spreadsheetId = spreadsheetId;
    this.cache = cache;
  }

  async fetch(sheetName) {
    const key = `sheet_${sheetName}`;
    const cached = this.cache.get(key);
    if (cached) return cached;

    const url = `https://docs.google.com/spreadsheets/d/${this.spreadsheetId}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(sheetName)}`;
    const res = await fetch(url);
    const text = await res.text();
    const match = text.match(/google\.visualization\.Query\.setResponse\(([\s\S]*)\);?$/);
    if (!match) throw new Error('解析失敗');

    const data = JSON.parse(match[1]);
    if (data.status === 'error') throw new Error(data.errors?.[0]?.message || '獲取失敗');

    const headers = data.table.cols.map(c => c.label || c.id);
    const rows = data.table.rows.map(r => {
      const row = {};
      r.c.forEach((cell, i) => { row[headers[i]] = cell ? (cell.v ?? cell.f ?? '') : ''; });
      return row;
    });

    const result = { headers, rows };
    this.cache.set(key, result);
    return result;
  }

  async getAll() {
    const [identity, staff, student, teacher] = await Promise.all([
      this.fetch('身分數據'),
      this.fetch('臺大教職員工數據庫'),
      this.fetch('臺大學生數據'),
      this.fetch('臺大教師數據')
    ]);
    return { identity, staff, student, teacher };
  }
}

// ==================== 圖表生成器類別 ====================

class ChartGenerator {
  constructor(months, dataType) {
    this.months = months;
    this.monthSet = new Set(months);
    this.dataType = dataType;
  }

  initCounts(cats) {
    const counts = {};
    this.months.forEach(m => { counts[m] = {}; cats.forEach(c => counts[m][c] = 0); });
    return counts;
  }

  buildDatasets(counts, cats) {
    const datasets = [];
    if (this.dataType === 'cumulative') {
      cats.forEach(c => {
        let cum = 0;
        const data = this.months.map(m => { cum += counts[m][c] || 0; return cum; });
        if (data.some(v => v > 0)) datasets.push({ label: c, data });
      });
    } else {
      cats.forEach(c => {
        const data = this.months.map(m => counts[m][c] || 0);
        if (data.some(v => v > 0)) datasets.push({ label: c, data });
      });
    }
    return datasets;
  }

  chartType() { return this.months.length === 1 ? 'pie' : 'bar'; }
}

class Chart0Generator extends ChartGenerator {
  generate(data) {
    const counts = this.initCounts(['總計']);
    data.rows.forEach(r => {
      const date = Object.values(r).find(v => DateUtils.parseMonthKey(v));
      const m = DateUtils.parseMonthKey(date);
      if (m && this.monthSet.has(m)) counts[m]['總計']++;
    });
    return {
      title: this.dataType === 'cumulative' ? '總報名人數累計' : '總報名人數',
      chartType: 'line', labels: this.months, datasets: this.buildDatasets(counts, ['總計'])
    };
  }
}

class Chart1Generator extends ChartGenerator {
  generate(data) {
    const cats = ['教師', '學生', '研究員', '其他'];
    const counts = this.initCounts(cats);
    data.rows.forEach(r => {
      const id = Object.values(r)[0], date = Object.values(r)[24];
      const m = DateUtils.parseMonthKey(date);
      if (m && this.monthSet.has(m) && id) {
        const s = String(id);
        let cat = '其他';
        if (s.includes('教師')) cat = '教師';
        else if (s.includes('學生')) cat = '學生';
        else if (s.includes('研究員')) cat = '研究員';
        counts[m][cat]++;
      }
    });
    return {
      title: this.dataType === 'cumulative' ? '校內外報名者累計分布' : '校內外報名者分布',
      chartType: this.chartType(), labels: this.months, datasets: this.buildDatasets(counts, cats)
    };
  }
}

class Chart2Generator extends ChartGenerator {
  generate(staff, student) {
    const cats = ['臺大教師', '臺大學生', '研究員'];
    const counts = this.initCounts(cats);
    staff.rows.forEach(r => {
      const id = Object.values(r)[4], date = Object.values(r)[17];
      const m = DateUtils.parseMonthKey(date);
      if (m && this.monthSet.has(m) && id) {
        const s = String(id);
        if (s.includes('教師')) counts[m]['臺大教師']++;
        else if (s.includes('研究員')) counts[m]['研究員']++;
      }
    });
    student.rows.forEach(r => {
      const id = Object.values(r)[1], date = Object.values(r)[12];
      const m = DateUtils.parseMonthKey(date);
      if (m && this.monthSet.has(m) && id && String(id).includes('學生')) counts[m]['臺大學生']++;
    });
    return {
      title: this.dataType === 'cumulative' ? '台大報名者累計分布' : '台大報名者分布',
      chartType: this.chartType(), labels: this.months, datasets: this.buildDatasets(counts, cats)
    };
  }
}

class Chart3Generator extends ChartGenerator {
  generate(data) {
    const cats = ['專任教師', '兼任教師', '專案教師', '臨床教師'];
    const counts = this.initCounts(cats);
    data.rows.forEach(r => {
      const job = Object.values(r)[6], date = Object.values(r)[12];
      const m = DateUtils.parseMonthKey(date);
      if (m && this.monthSet.has(m)) {
        const type = TeacherClassifier.getType(job);
        if (type) counts[m][type]++;
      }
    });
    return {
      title: this.dataType === 'cumulative' ? '教師所有職級累計分布' : '教師所有職級分布',
      chartType: this.chartType(), labels: this.months, datasets: this.buildDatasets(counts, cats)
    };
  }
}

class TeacherRankGenerator extends ChartGenerator {
  constructor(months, dataType, type, titleMap, keywords = []) {
    super(months, dataType);
    this.type = type;
    this.titleMap = titleMap;
    this.keywords = keywords;
  }

  generate(data) {
    const cats = ['教授', '副教授', '助理教授', '講師'];
    const counts = this.initCounts(cats);
    data.rows.forEach(r => {
      const job = Object.values(r)[6], date = Object.values(r)[12];
      const m = DateUtils.parseMonthKey(date);
      if (m && this.monthSet.has(m) && job) {
        const s = String(job).trim();
        if (this.titleMap) {
          const rank = this.titleMap[s];
          if (rank) counts[m][rank]++;
        } else if (this.keywords.length) {
          if (this.keywords.some(k => s.includes(k))) {
            const rank = TeacherClassifier.getRank(s);
            if (rank) counts[m][rank]++;
          }
        }
      }
    });
    return {
      title: this.dataType === 'cumulative' ? `${this.type}職級累計分布` : `${this.type}職級分布`,
      chartType: this.chartType(), labels: this.months, datasets: this.buildDatasets(counts, cats)
    };
  }
}

class Chart8Generator extends ChartGenerator {
  generate(data) {
    const cats = StudentClassifier.levels;
    const counts = this.initCounts(cats);
    data.rows.forEach(r => {
      const id = Object.values(r)[1], lv = Object.values(r)[2], date = Object.values(r)[12];
      const m = DateUtils.parseMonthKey(date);
      if (m && this.monthSet.has(m)) {
        const level = StudentClassifier.getLevel(id, lv);
        if (level) counts[m][level]++;
      }
    });
    return {
      title: this.dataType === 'cumulative' ? '學生所有職級累計分布' : '學生所有職級分布',
      chartType: this.chartType(), labels: this.months, datasets: this.buildDatasets(counts, cats)
    };
  }
}

class CollegeGenerator extends ChartGenerator {
  constructor(months, dataType, type) {
    super(months, dataType);
    this.type = type;
  }

  generate(teacher, student) {
    const cats = CollegeClassifier.categories;
    const counts = this.initCounts(cats);
    if (this.type === 'teacher' || this.type === 'combined') {
      teacher.rows.forEach(r => {
        const col = Object.values(r)[4], date = Object.values(r)[12];
        const m = DateUtils.parseMonthKey(date);
        if (m && this.monthSet.has(m) && col) counts[m][CollegeClassifier.classify(col)]++;
      });
    }
    if (this.type === 'student' || this.type === 'combined') {
      student.rows.forEach(r => {
        const col = Object.values(r)[7], date = Object.values(r)[12];
        const m = DateUtils.parseMonthKey(date);
        if (m && this.monthSet.has(m) && col) counts[m][CollegeClassifier.classify(col)]++;
      });
    }
    const names = { teacher: '教師學院', student: '學生學院', combined: '教師與學生學院' };
    return {
      title: this.dataType === 'cumulative' ? `${names[this.type]}累計分布` : `${names[this.type]}分布`,
      chartType: this.chartType(), labels: this.months, datasets: this.buildDatasets(counts, cats)
    };
  }
}

// ==================== 圖表服務 ====================

class ChartService {
  constructor(sheets) { this.sheets = sheets; }

  async generate(chartType, months, dataType) {
    const data = await this.sheets.getAll();
    const gens = {
      chart0: () => new Chart0Generator(months, dataType).generate(data.identity),
      chart1: () => new Chart1Generator(months, dataType).generate(data.identity),
      chart2: () => new Chart2Generator(months, dataType).generate(data.staff, data.student),
      chart3: () => new Chart3Generator(months, dataType).generate(data.teacher),
      chart4: () => new TeacherRankGenerator(months, dataType, '專任教師', TeacherClassifier.fullTimeMap).generate(data.teacher),
      chart5: () => new TeacherRankGenerator(months, dataType, '兼任教師', TeacherClassifier.adjunctMap).generate(data.teacher),
      chart6: () => new TeacherRankGenerator(months, dataType, '專案教師', null, ['專案', 'Project']).generate(data.teacher),
      chart7: () => new TeacherRankGenerator(months, dataType, '臨床教師', TeacherClassifier.clinicalMap).generate(data.teacher),
      chart8: () => new Chart8Generator(months, dataType).generate(data.student),
      chart9: () => new CollegeGenerator(months, dataType, 'teacher').generate(data.teacher, data.student),
      chart10: () => new CollegeGenerator(months, dataType, 'student').generate(data.teacher, data.student),
      chart11: () => new CollegeGenerator(months, dataType, 'combined').generate(data.teacher, data.student)
    };
    if (!gens[chartType]) throw new Error('無效的圖表類型');
    return gens[chartType]();
  }

  async getTimeOptions() {
    const data = await this.sheets.fetch('身分數據');
    const monthSet = new Set();
    data.rows.forEach(r => {
      const m = DateUtils.parseMonthKey(Object.values(r)[24]);
      if (m) monthSet.add(m);
    });
    const min = DateUtils.getMinMonth();
    const months = [...monthSet].filter(m => m >= min).sort();
    const semesters = [...new Set(months.map(m => DateUtils.monthToSemester(m)))].sort();
    const years = [...new Set(months.map(m => DateUtils.monthToYear(m)))].sort();
    return { months, semesters, years };
  }
}

// ==================== Express App ====================

class App {
  constructor() {
    this.app = express();
    this.port = process.env.PORT || 3000;
    this.cache = new CacheManager();
    this.sheets = new GoogleSheetsService('1OMjAbOwTssGqKHBC0oM-C-ds17MgMQbPsWCfH2bemjY', this.cache);
    this.charts = new ChartService(this.sheets);
    this.setup();
  }

  setup() {
    this.app.use(helmet({ contentSecurityPolicy: false, crossOriginEmbedderPolicy: false }));
    this.app.use(cors({ origin: true }));
    this.app.use('/api/', rateLimit({ windowMs: 15 * 60 * 1000, max: 100 }));
    this.app.use(compression());
    this.app.use(express.json({ limit: '10kb' }));
    this.app.use(express.static(path.join(__dirname, '../public')));

    this.app.get('/api/health', (_, res) => res.json({ status: 'ok' }));

    this.app.get('/api/time-options', async (_, res) => {
      try {
        const data = await this.charts.getTimeOptions();
        res.json({ success: true, data });
      } catch (e) {
        res.status(500).json({ success: false, error: e.message });
      }
    });

    this.app.post('/api/chart-data', async (req, res) => {
      try {
        const { chartType, timeSelections, timeMode, dataType } = req.body;
        if (!Validator.isValidChart(chartType)) return res.status(400).json({ error: '無效圖表' });
        if (!Validator.isValidDataType(dataType)) return res.status(400).json({ error: '無效類型' });

        const opts = await this.charts.getTimeOptions();
        const availSet = new Set(opts.months);
        let months = [];

        if (timeMode === 'month') {
          months = timeSelections.filter(m => availSet.has(m));
        } else if (timeMode === 'semester') {
          timeSelections.forEach(s => DateUtils.semesterToMonths(s).forEach(m => { if (availSet.has(m)) months.push(m); }));
        } else if (timeMode === 'year') {
          timeSelections.forEach(y => DateUtils.yearToMonths(y).forEach(m => { if (availSet.has(m)) months.push(m); }));
        }

        months = [...new Set(months)].sort();
        if (!months.length) return res.status(400).json({ error: '無有效時間' });

        const result = await this.charts.generate(chartType, months, dataType);
        res.json({ success: true, data: result });
      } catch (e) {
        res.status(500).json({ success: false, error: e.message });
      }
    });

    this.app.post('/api/clear-cache', (_, res) => {
      this.cache.flush();
      res.json({ success: true });
    });

    this.app.use((req, res) => {
      res.sendFile(path.join(__dirname, '../public/index.html'));
    });
  }

  start() {
    this.app.listen(this.port, () => {
      console.log(`\n  ACE 圖表系統 v2.0 | http://localhost:${this.port}\n`);
    });
  }
}

new App().start();
