(function (root) {
  const monthMap = {
    ene: '01',
    enero: '01',
    feb: '02',
    febrero: '02',
    mar: '03',
    marzo: '03',
    abr: '04',
    abril: '04',
    may: '05',
    mayo: '05',
    jun: '06',
    junio: '06',
    jul: '07',
    julio: '07',
    ago: '08',
    agosto: '08',
    sep: '09',
    set: '09',
    septiembre: '09',
    setiembre: '09',
    oct: '10',
    octubre: '10',
    nov: '11',
    noviembre: '11',
    dic: '12',
    diciembre: '12'
  };

  function normalizeText(value) {
    return String(value || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[|]/g, 'I')
      .toLowerCase();
  }

  function compactOcrText(value) {
    return normalizeText(value)
      .replace(/[0]/g, 'o')
      .replace(/[1]/g, 'l')
      .replace(/[^a-z0-9]+/g, ' ')
      .trim();
  }

  function normalizeYear(value) {
    const year = Number(value);
    if (value.length === 2) {
      return String(year >= 70 ? 1900 + year : 2000 + year);
    }

    return String(year).padStart(4, '0');
  }

  function monthFromText(value) {
    const normalized = normalizeText(value).replace(/[^a-z]/g, '');
    if (monthMap[normalized]) {
      return monthMap[normalized];
    }

    const match = Object.keys(monthMap).find((month) => normalized.startsWith(month.slice(0, 3)));
    return match ? monthMap[match] : '';
  }

  function parseDate(text) {
    const normalized = normalizeText(text);
    const monthDate = normalized.match(/\b(\d{1,2})\s*[\/\-. ]\s*([a-z]{3,12})\s*[\/\-. ]\s*(\d{2,4})\b/i);

    if (monthDate) {
      const day = monthDate[1].padStart(2, '0');
      const month = monthFromText(monthDate[2]);
      const year = normalizeYear(monthDate[3]);

      if (month) {
        return {
          date: `${year}-${month}-${day}`,
          rawDate: monthDate[0]
        };
      }
    }

    const numericDate = normalized.match(/\b(\d{1,2})\s*[\/\-.]\s*(\d{1,2})\s*[\/\-.]\s*(\d{2,4})\b/);

    if (numericDate) {
      const day = numericDate[1].padStart(2, '0');
      const month = numericDate[2].padStart(2, '0');
      const year = normalizeYear(numericDate[3]);

      if (Number(month) >= 1 && Number(month) <= 12 && Number(day) >= 1 && Number(day) <= 31) {
        return {
          date: `${year}-${month}-${day}`,
          rawDate: numericDate[0]
        };
      }
    }

    return {
      date: '',
      rawDate: ''
    };
  }

  function parseLocalizedAmount(value) {
    const cleaned = String(value || '').replace(/[^\d.,]/g, '');
    const lastDot = cleaned.lastIndexOf('.');
    const lastComma = cleaned.lastIndexOf(',');
    const decimalIndex = Math.max(lastDot, lastComma);

    if (decimalIndex === -1) {
      return Number(cleaned.replace(/[^\d]/g, ''));
    }

    const whole = cleaned.slice(0, decimalIndex).replace(/[^\d]/g, '');
    const decimals = cleaned.slice(decimalIndex + 1).replace(/[^\d]/g, '').slice(0, 2);
    return Number(`${whole}.${decimals.padEnd(2, '0')}`);
  }

  function extractAmounts(value) {
    const amountPattern = /(?:CRC|¢|₡|\$)?\s*(\d{1,3}(?:[.,]\d{3})+(?:[.,]\d{2})|\d{1,7}[.,]\d{2})/gi;
    const amounts = [];
    let match;

    while ((match = amountPattern.exec(value)) !== null) {
      const raw = match[1];
      const amount = parseLocalizedAmount(raw);

      if (Number.isFinite(amount)) {
        amounts.push({
          raw,
          amount
        });
      }
    }

    return amounts;
  }

  function lineScore(line, index, totalLines) {
    const normalized = compactOcrText(line);
    let score = 0;

    if (/\btotal\b/.test(normalized)) score += 30;
    if (/\bnet[oa0]?\b|\bnete\b|\bneto\b/.test(normalized)) score += 60;
    if (/\bbruto\b|\bsubtotal\b|\biva\b|\bi v a\b|\bdescuento\b|\bimpuesto\b/.test(normalized)) score -= 45;

    score += Math.min(index / Math.max(totalLines, 1), 1) * 8;
    return score;
  }

  function parseTotalNeto(text) {
    const lines = String(text || '')
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);
    const candidates = [];

    lines.forEach((line, index) => {
      const nearby = [line, lines[index + 1] || ''].join(' ');
      const amounts = extractAmounts(nearby);
      const score = lineScore(line, index, lines.length);

      amounts.forEach((amount) => {
        candidates.push({
          ...amount,
          line,
          score
        });
      });
    });

    if (!candidates.length) {
      return {
        amount: null,
        rawAmount: '',
        amountLine: ''
      };
    }

    const strong = candidates
      .filter((candidate) => candidate.score >= 80)
      .sort((a, b) => b.score - a.score || b.amount - a.amount)[0];

    if (strong) {
      return {
        amount: strong.amount,
        rawAmount: strong.raw,
        amountLine: strong.line
      };
    }

    const fallback = candidates
      .filter((candidate) => candidate.amount > 0)
      .sort((a, b) => b.score - a.score || b.amount - a.amount)[0];

    return {
      amount: fallback?.amount ?? null,
      rawAmount: fallback?.raw || '',
      amountLine: fallback?.line || ''
    };
  }

  function parseReceiptText(text) {
    const total = parseTotalNeto(text);
    const date = parseDate(text);

    return {
      amount: total.amount,
      rawAmount: total.rawAmount,
      amountLine: total.amountLine,
      date: date.date,
      rawDate: date.rawDate
    };
  }

  const api = {
    extractAmounts,
    parseDate,
    parseLocalizedAmount,
    parseReceiptText,
    parseTotalNeto
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  }

  root.ReceiptParser = api;
}(typeof window !== 'undefined' ? window : globalThis));
