/*! xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
function make_xlsx_lib(XLSX) {
  XLSX.version = "0.16.6";
  let current_codepage = 1200;
  let current_ansi = 1252;
  /* global cptable:true, window */
  if (typeof module !== "undefined" && typeof require !== "undefined") {
    if (typeof cptable === "undefined") {
      if (typeof global !== "undefined")
        global.cptable = require("./cpexcel.js");
      else if (typeof window !== "undefined")
        window.cptable = require("./cpexcel.js");
    }
  }

  const VALID_ANSI = [874, 932, 936, 949, 950];
  for (let i = 0; i <= 8; ++i) VALID_ANSI.push(1250 + i);
  /* ECMA-376 Part I 18.4.1 charset to codepage mapping */
  const CS2CP = {
    0: 1252 /* ANSI */,
    1: 65001 /* DEFAULT */,
    2: 65001 /* SYMBOL */,
    77: 10000 /* MAC */,
    128: 932 /* SHIFTJIS */,
    129: 949 /* HANGUL */,
    130: 1361 /* JOHAB */,
    134: 936 /* GB2312 */,
    136: 950 /* CHINESEBIG5 */,
    161: 1253 /* GREEK */,
    162: 1254 /* TURKISH */,
    163: 1258 /* VIETNAMESE */,
    177: 1255 /* HEBREW */,
    178: 1256 /* ARABIC */,
    186: 1257 /* BALTIC */,
    204: 1251 /* RUSSIAN */,
    222: 874 /* THAI */,
    238: 1250 /* EASTEUROPE */,
    255: 1252 /* OEM */,
    69: 6969 /* MISC */
  };

  const set_ansi = function(cp) {
    if (VALID_ANSI.indexOf(cp) == -1) return;
    current_ansi = CS2CP[0] = cp;
  };
  function reset_ansi() {
    set_ansi(1252);
  }

  let set_cp = function(cp) {
    current_codepage = cp;
    set_ansi(cp);
  };
  function reset_cp() {
    set_cp(1200);
    reset_ansi();
  }

  function char_codes(data) {
    const o = [];
    for (let i = 0, len = data.length; i < len; ++i) o[i] = data.charCodeAt(i);
    return o;
  }

  function utf16leread(data) {
    const o = [];
    for (let i = 0; i < data.length >> 1; ++i)
      o[i] = String.fromCharCode(
        data.charCodeAt(2 * i) + (data.charCodeAt(2 * i + 1) << 8)
      );
    return o.join("");
  }
  function utf16beread(data) {
    const o = [];
    for (let i = 0; i < data.length >> 1; ++i)
      o[i] = String.fromCharCode(
        data.charCodeAt(2 * i + 1) + (data.charCodeAt(2 * i) << 8)
      );
    return o.join("");
  }

  let debom = function(data) {
    const c1 = data.charCodeAt(0);
    const c2 = data.charCodeAt(1);
    if (c1 == 0xff && c2 == 0xfe) return utf16leread(data.slice(2));
    if (c1 == 0xfe && c2 == 0xff) return utf16beread(data.slice(2));
    if (c1 == 0xfeff) return data.slice(1);
    return data;
  };

  let _getchar = function _gc1(x) {
    return String.fromCharCode(x);
  };
  let _getansi = function _ga1(x) {
    return String.fromCharCode(x);
  };
  if (typeof cptable !== "undefined") {
    set_cp = function(cp) {
      current_codepage = cp;
      set_ansi(cp);
    };
    debom = function(data) {
      if (data.charCodeAt(0) === 0xff && data.charCodeAt(1) === 0xfe) {
        return cptable.utils.decode(1200, char_codes(data.slice(2)));
      }
      return data;
    };
    _getchar = function _gc2(x) {
      if (current_codepage === 1200) return String.fromCharCode(x);
      return cptable.utils.decode(current_codepage, [x & 255, x >> 8])[0];
    };
    _getansi = function _ga2(x) {
      return cptable.utils.decode(current_ansi, [x])[0];
    };
  }
  const DENSE = null;
  const DIF_XL = true;
  const Base64 = (function make_b64() {
    const map =
      "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
    return {
      encode(input) {
        let o = "";
        let c1 = 0;
        let c2 = 0;
        let c3 = 0;
        let e1 = 0;
        let e2 = 0;
        let e3 = 0;
        let e4 = 0;
        for (let i = 0; i < input.length; ) {
          c1 = input.charCodeAt(i++);
          e1 = c1 >> 2;

          c2 = input.charCodeAt(i++);
          e2 = ((c1 & 3) << 4) | (c2 >> 4);

          c3 = input.charCodeAt(i++);
          e3 = ((c2 & 15) << 2) | (c3 >> 6);
          e4 = c3 & 63;
          if (isNaN(c2)) {
            e3 = e4 = 64;
          } else if (isNaN(c3)) {
            e4 = 64;
          }
          o +=
            map.charAt(e1) + map.charAt(e2) + map.charAt(e3) + map.charAt(e4);
        }
        return o;
      },
      decode: function b64_decode(input) {
        let o = "";
        let c1 = 0;
        let c2 = 0;
        let c3 = 0;
        let e1 = 0;
        let e2 = 0;
        let e3 = 0;
        let e4 = 0;
        input = input.replace(/[^\w\+\/\=]/g, "");
        for (let i = 0; i < input.length; ) {
          e1 = map.indexOf(input.charAt(i++));
          e2 = map.indexOf(input.charAt(i++));
          c1 = (e1 << 2) | (e2 >> 4);
          o += String.fromCharCode(c1);

          e3 = map.indexOf(input.charAt(i++));
          c2 = ((e2 & 15) << 4) | (e3 >> 2);
          if (e3 !== 64) {
            o += String.fromCharCode(c2);
          }

          e4 = map.indexOf(input.charAt(i++));
          c3 = ((e3 & 3) << 6) | e4;
          if (e4 !== 64) {
            o += String.fromCharCode(c3);
          }
        }
        return o;
      }
    };
  })();
  const has_buf =
    typeof Buffer !== "undefined" &&
    typeof process !== "undefined" &&
    typeof process.versions !== "undefined" &&
    !!process.versions.node;

  let Buffer_from = function() {};

  if (typeof Buffer !== "undefined") {
    let nbfs = !Buffer.from;
    if (!nbfs)
      try {
        Buffer.from("foo", "utf8");
      } catch (e) {
        nbfs = true;
      }
    Buffer_from = nbfs
      ? function(buf, enc) {
          return enc ? new Buffer(buf, enc) : new Buffer(buf);
        }
      : Buffer.from.bind(Buffer);
    // $FlowIgnore
    if (!Buffer.alloc)
      Buffer.alloc = function(n) {
        return new Buffer(n);
      };
    // $FlowIgnore
    if (!Buffer.allocUnsafe)
      Buffer.allocUnsafe = function(n) {
        return new Buffer(n);
      };
  }

  function new_raw_buf(len) {
    /* jshint -W056 */
    return has_buf ? Buffer.alloc(len) : new Array(len);
    /* jshint +W056 */
  }

  function new_unsafe_buf(len) {
    /* jshint -W056 */
    return has_buf ? Buffer.allocUnsafe(len) : new Array(len);
    /* jshint +W056 */
  }

  const s2a = function s2a(s) {
    if (has_buf) return Buffer_from(s, "binary");
    return s.split("").map(function(x) {
      return x.charCodeAt(0) & 0xff;
    });
  };

  function a2s(data) {
    if (Array.isArray(data))
      return data
        .map(function(c) {
          return String.fromCharCode(c);
        })
        .join("");
    const o = [];
    for (let i = 0; i < data.length; ++i) o[i] = String.fromCharCode(data[i]);
    return o.join("");
  }

  function ab2a(data) {
    if (typeof ArrayBuffer === "undefined") throw new Error("Unsupported");
    if (data instanceof ArrayBuffer) return ab2a(new Uint8Array(data));
    const o = new Array(data.length);
    for (let i = 0; i < data.length; ++i) o[i] = data[i];
    return o;
  }

  let bconcat = function(bufs) {
    return [].concat.apply([], bufs);
  };

  const chr0 = /\u0000/g;
  const chr1 = /[\u0001-\u0006]/g;
  /* ssf.js (C) 2013-present SheetJS -- http://sheetjs.com */
  /* jshint -W041 */
  const SSF = {};
  const make_ssf = function make_ssf(SSF) {
    SSF.version = "0.11.2";
    function _strrev(x) {
      let o = "";
      let i = x.length - 1;
      while (i >= 0) o += x.charAt(i--);
      return o;
    }
    function fill(c, l) {
      let o = "";
      while (o.length < l) o += c;
      return o;
    }
    function pad0(v, d) {
      const t = `${v}`;
      return t.length >= d ? t : fill("0", d - t.length) + t;
    }
    function pad_(v, d) {
      const t = `${v}`;
      return t.length >= d ? t : fill(" ", d - t.length) + t;
    }
    function rpad_(v, d) {
      const t = `${v}`;
      return t.length >= d ? t : t + fill(" ", d - t.length);
    }
    function pad0r1(v, d) {
      const t = `${Math.round(v)}`;
      return t.length >= d ? t : fill("0", d - t.length) + t;
    }
    function pad0r2(v, d) {
      const t = `${v}`;
      return t.length >= d ? t : fill("0", d - t.length) + t;
    }
    const p2_32 = Math.pow(2, 32);
    function pad0r(v, d) {
      if (v > p2_32 || v < -p2_32) return pad0r1(v, d);
      const i = Math.round(v);
      return pad0r2(i, d);
    }
    function isgeneral(s, i) {
      i = i || 0;
      return (
        s.length >= 7 + i &&
        (s.charCodeAt(i) | 32) === 103 &&
        (s.charCodeAt(i + 1) | 32) === 101 &&
        (s.charCodeAt(i + 2) | 32) === 110 &&
        (s.charCodeAt(i + 3) | 32) === 101 &&
        (s.charCodeAt(i + 4) | 32) === 114 &&
        (s.charCodeAt(i + 5) | 32) === 97 &&
        (s.charCodeAt(i + 6) | 32) === 108
      );
    }
    const days = [
      ["Sun", "Sunday"],
      ["Mon", "Monday"],
      ["Tue", "Tuesday"],
      ["Wed", "Wednesday"],
      ["Thu", "Thursday"],
      ["Fri", "Friday"],
      ["Sat", "Saturday"]
    ];
    const months = [
      ["J", "Jan", "January"],
      ["F", "Feb", "February"],
      ["M", "Mar", "March"],
      ["A", "Apr", "April"],
      ["M", "May", "May"],
      ["J", "Jun", "June"],
      ["J", "Jul", "July"],
      ["A", "Aug", "August"],
      ["S", "Sep", "September"],
      ["O", "Oct", "October"],
      ["N", "Nov", "November"],
      ["D", "Dec", "December"]
    ];
    function init_table(t) {
      t[0] = "General";
      t[1] = "0";
      t[2] = "0.00";
      t[3] = "#,##0";
      t[4] = "#,##0.00";
      t[9] = "0%";
      t[10] = "0.00%";
      t[11] = "0.00E+00";
      t[12] = "# ?/?";
      t[13] = "# ??/??";
      t[14] = "m/d/yy";
      t[15] = "d-mmm-yy";
      t[16] = "d-mmm";
      t[17] = "mmm-yy";
      t[18] = "h:mm AM/PM";
      t[19] = "h:mm:ss AM/PM";
      t[20] = "h:mm";
      t[21] = "h:mm:ss";
      t[22] = "m/d/yy h:mm";
      t[37] = "#,##0 ;(#,##0)";
      t[38] = "#,##0 ;[Red](#,##0)";
      t[39] = "#,##0.00;(#,##0.00)";
      t[40] = "#,##0.00;[Red](#,##0.00)";
      t[45] = "mm:ss";
      t[46] = "[h]:mm:ss";
      t[47] = "mmss.0";
      t[48] = "##0.0E+0";
      t[49] = "@";
      t[56] = '"上午/下午 "hh"時"mm"分"ss"秒 "';
    }

    const table_fmt = {};
    init_table(table_fmt);
    /* Defaults determined by systematically testing in Excel 2019 */

    /* These formats appear to default to other formats in the table */
    const default_map = [];
    let defi = 0;

    //  5 -> 37 ...  8 -> 40
    for (defi = 5; defi <= 8; ++defi) default_map[defi] = 32 + defi;

    // 23 ->  0 ... 26 ->  0
    for (defi = 23; defi <= 26; ++defi) default_map[defi] = 0;

    // 27 -> 14 ... 31 -> 14
    for (defi = 27; defi <= 31; ++defi) default_map[defi] = 14;
    // 50 -> 14 ... 58 -> 14
    for (defi = 50; defi <= 58; ++defi) default_map[defi] = 14;

    // 59 ->  1 ... 62 ->  4
    for (defi = 59; defi <= 62; ++defi) default_map[defi] = defi - 58;
    // 67 ->  9 ... 68 -> 10
    for (defi = 67; defi <= 68; ++defi) default_map[defi] = defi - 58;
    // 72 -> 14 ... 75 -> 17
    for (defi = 72; defi <= 75; ++defi) default_map[defi] = defi - 58;

    // 69 -> 12 ... 71 -> 14
    for (defi = 67; defi <= 68; ++defi) default_map[defi] = defi - 57;

    // 76 -> 20 ... 78 -> 22
    for (defi = 76; defi <= 78; ++defi) default_map[defi] = defi - 56;

    // 79 -> 45 ... 81 -> 47
    for (defi = 79; defi <= 81; ++defi) default_map[defi] = defi - 34;

    // 82 ->  0 ... 65536 -> 0 (omitted)

    /* These formats technically refer to Accounting formats with no equivalent */
    const default_str = [];

    //  5 -- Currency,   0 decimal, black negative
    default_str[5] = default_str[63] = '"$"#,##0_);\\("$"#,##0\\)';
    //  6 -- Currency,   0 decimal, red   negative
    default_str[6] = default_str[64] = '"$"#,##0_);[Red]\\("$"#,##0\\)';
    //  7 -- Currency,   2 decimal, black negative
    default_str[7] = default_str[65] = '"$"#,##0.00_);\\("$"#,##0.00\\)';
    //  8 -- Currency,   2 decimal, red   negative
    default_str[8] = default_str[66] = '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)';

    // 41 -- Accounting, 0 decimal, No Symbol
    default_str[41] = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)';
    // 42 -- Accounting, 0 decimal, $  Symbol
    default_str[42] = '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)';
    // 43 -- Accounting, 2 decimal, No Symbol
    default_str[43] = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
    // 44 -- Accounting, 2 decimal, $  Symbol
    default_str[44] =
      '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';
    function frac(x, D, mixed) {
      const sgn = x < 0 ? -1 : 1;
      let B = x * sgn;
      let P_2 = 0;
      let P_1 = 1;
      let P = 0;
      let Q_2 = 1;
      let Q_1 = 0;
      let Q = 0;
      let A = Math.floor(B);
      while (Q_1 < D) {
        A = Math.floor(B);
        P = A * P_1 + P_2;
        Q = A * Q_1 + Q_2;
        if (B - A < 0.00000005) break;
        B = 1 / (B - A);
        P_2 = P_1;
        P_1 = P;
        Q_2 = Q_1;
        Q_1 = Q;
      }
      if (Q > D) {
        if (Q_1 > D) {
          Q = Q_2;
          P = P_2;
        } else {
          Q = Q_1;
          P = P_1;
        }
      }
      if (!mixed) return [0, sgn * P, Q];
      const q = Math.floor((sgn * P) / Q);
      return [q, sgn * P - q * Q, Q];
    }
    function parse_date_code(v, opts, b2) {
      if (v > 2958465 || v < 0) return null;
      let date = v | 0;
      let time = Math.floor(86400 * (v - date));
      let dow = 0;
      let dout = [];
      const out = {
        D: date,
        T: time,
        u: 86400 * (v - date) - time,
        y: 0,
        m: 0,
        d: 0,
        H: 0,
        M: 0,
        S: 0,
        q: 0
      };
      if (Math.abs(out.u) < 1e-6) out.u = 0;
      if (opts && opts.date1904) date += 1462;
      if (out.u > 0.9999) {
        out.u = 0;
        if (++time == 86400) {
          out.T = time = 0;
          ++date;
          ++out.D;
        }
      }
      if (date === 60) {
        dout = b2 ? [1317, 10, 29] : [1900, 2, 29];
        dow = 3;
      } else if (date === 0) {
        dout = b2 ? [1317, 8, 29] : [1900, 1, 0];
        dow = 6;
      } else {
        if (date > 60) --date;
        /* 1 = Jan 1 1900 in Gregorian */
        const d = new Date(1900, 0, 1);
        d.setDate(d.getDate() + date - 1);
        dout = [d.getFullYear(), d.getMonth() + 1, d.getDate()];
        dow = d.getDay();
        if (date < 60) dow = (dow + 6) % 7;
        if (b2) dow = fix_hijri(d, dout);
      }
      out.y = dout[0];
      out.m = dout[1];
      out.d = dout[2];
      out.S = time % 60;
      time = Math.floor(time / 60);
      out.M = time % 60;
      time = Math.floor(time / 60);
      out.H = time;
      out.q = dow;
      return out;
    }
    SSF.parse_date_code = parse_date_code;
    const basedate = new Date(1899, 11, 31, 0, 0, 0);
    const dnthresh = basedate.getTime();
    const base1904 = new Date(1900, 2, 1, 0, 0, 0);
    function datenum_local(v, date1904) {
      let epoch = v.getTime();
      if (date1904) epoch -= 1461 * 24 * 60 * 60 * 1000;
      else if (v >= base1904) epoch += 24 * 60 * 60 * 1000;
      return (
        (epoch -
          (dnthresh +
            (v.getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000)) /
        (24 * 60 * 60 * 1000)
      );
    }
    /* The longest 32-bit integer text is "-4294967296", exactly 11 chars */
    function general_fmt_int(v) {
      return v.toString(10);
    }
    SSF._general_int = general_fmt_int;

    /* ECMA-376 18.8.30 numFmt */
    /* Note: `toPrecision` uses standard form when prec > E and E >= -6 */
    const general_fmt_num = (function make_general_fmt_num() {
      const trailing_zeroes_and_decimal = /(?:\.0*|(\.\d*[1-9])0+)$/;
      function strip_decimal(o) {
        return o.indexOf(".") == -1
          ? o
          : o.replace(trailing_zeroes_and_decimal, "$1");
      }

      /* General Exponential always shows 2 digits exp and trims the mantissa */
      const mantissa_zeroes_and_decimal = /(?:\.0*|(\.\d*[1-9])0+)[Ee]/;
      const exp_with_single_digit = /(E[+-])(\d)$/;
      function normalize_exp(o) {
        if (o.indexOf("E") == -1) return o;
        return o
          .replace(mantissa_zeroes_and_decimal, "$1E")
          .replace(exp_with_single_digit, "$10$2");
      }

      /* exponent >= -9 and <= 9 */
      function small_exp(v) {
        const w = v < 0 ? 12 : 11;
        let o = strip_decimal(v.toFixed(12));
        if (o.length <= w) return o;
        o = v.toPrecision(10);
        if (o.length <= w) return o;
        return v.toExponential(5);
      }

      /* exponent >= 11 or <= -10 likely exponential */
      function large_exp(v) {
        const o = strip_decimal(v.toFixed(11));
        return o.length > (v < 0 ? 12 : 11) || o === "0" || o === "-0"
          ? v.toPrecision(6)
          : o;
      }

      function general_fmt_num_base(v) {
        const V = Math.floor(Math.log(Math.abs(v)) * Math.LOG10E);
        let o;

        if (V >= -4 && V <= -1) o = v.toPrecision(10 + V);
        else if (Math.abs(V) <= 9) o = small_exp(v);
        else if (V === 10) o = v.toFixed(10).substr(0, 12);
        else o = large_exp(v);

        return strip_decimal(normalize_exp(o.toUpperCase()));
      }

      return general_fmt_num_base;
    })();
    SSF._general_num = general_fmt_num;

    /*
	"General" rules:
	- text is passed through ("@")
	- booleans are rendered as TRUE/FALSE
	- "up to 11 characters" displayed for numbers
	- Default date format (code 14) used for Dates

	TODO: technically the display depends on the width of the cell
*/
    function general_fmt(v, opts) {
      switch (typeof v) {
        case "string":
          return v;
        case "boolean":
          return v ? "TRUE" : "FALSE";
        case "number":
          return (v | 0) === v ? v.toString(10) : general_fmt_num(v);
        case "undefined":
          return "";
        case "object":
          if (v == null) return "";
          if (v instanceof Date)
            return format(14, datenum_local(v, opts && opts.date1904), opts);
      }
      throw new Error(`unsupported value in General format: ${v}`);
    }
    SSF._general = general_fmt;
    function fix_hijri(date, o) {
      /* TODO: properly adjust y/m/d and  */
      o[0] -= 581;
      let dow = date.getDay();
      if (date < 60) dow = (dow + 6) % 7;
      return dow;
    }
    // var THAI_DIGITS = "\u0E50\u0E51\u0E52\u0E53\u0E54\u0E55\u0E56\u0E57\u0E58\u0E59".split("");
    /* jshint -W086 */
    function write_date(type, fmt, val, ss0) {
      let o = "";
      let ss = 0;
      let tt = 0;
      let { y } = val;
      let out;
      let outl = 0;
      switch (type) {
        case 98 /* 'b' buddhist year */:
          y = val.y + 543;
        /* falls through */
        case 121 /* 'y' year */:
          switch (fmt.length) {
            case 1:
            case 2:
              out = y % 100;
              outl = 2;
              break;
            default:
              out = y % 10000;
              outl = 4;
              break;
          }
          break;
        case 109 /* 'm' month */:
          switch (fmt.length) {
            case 1:
            case 2:
              out = val.m;
              outl = fmt.length;
              break;
            case 3:
              return months[val.m - 1][1];
            case 5:
              return months[val.m - 1][0];
            default:
              return months[val.m - 1][2];
          }
          break;
        case 100 /* 'd' day */:
          switch (fmt.length) {
            case 1:
            case 2:
              out = val.d;
              outl = fmt.length;
              break;
            case 3:
              return days[val.q][0];
            default:
              return days[val.q][1];
          }
          break;
        case 104 /* 'h' 12-hour */:
          switch (fmt.length) {
            case 1:
            case 2:
              out = 1 + ((val.H + 11) % 12);
              outl = fmt.length;
              break;
            default:
              throw `bad hour format: ${fmt}`;
          }
          break;
        case 72 /* 'H' 24-hour */:
          switch (fmt.length) {
            case 1:
            case 2:
              out = val.H;
              outl = fmt.length;
              break;
            default:
              throw `bad hour format: ${fmt}`;
          }
          break;
        case 77 /* 'M' minutes */:
          switch (fmt.length) {
            case 1:
            case 2:
              out = val.M;
              outl = fmt.length;
              break;
            default:
              throw `bad minute format: ${fmt}`;
          }
          break;
        case 115 /* 's' seconds */:
          if (
            fmt != "s" &&
            fmt != "ss" &&
            fmt != ".0" &&
            fmt != ".00" &&
            fmt != ".000"
          )
            throw `bad second format: ${fmt}`;
          if (val.u === 0 && (fmt == "s" || fmt == "ss"))
            return pad0(val.S, fmt.length);
          if (ss0 >= 2) tt = ss0 === 3 ? 1000 : 100;
          else tt = ss0 === 1 ? 10 : 1;
          ss = Math.round(tt * (val.S + val.u));
          if (ss >= 60 * tt) ss = 0;
          if (fmt === "s") return ss === 0 ? "0" : `${ss / tt}`;
          o = pad0(ss, 2 + ss0);
          if (fmt === "ss") return o.substr(0, 2);
          return `.${o.substr(2, fmt.length - 1)}`;
        case 90 /* 'Z' absolute time */:
          switch (fmt) {
            case "[h]":
            case "[hh]":
              out = val.D * 24 + val.H;
              break;
            case "[m]":
            case "[mm]":
              out = (val.D * 24 + val.H) * 60 + val.M;
              break;
            case "[s]":
            case "[ss]":
              out =
                ((val.D * 24 + val.H) * 60 + val.M) * 60 +
                Math.round(val.S + val.u);
              break;
            default:
              throw `bad abstime format: ${fmt}`;
          }
          outl = fmt.length === 3 ? 1 : 2;
          break;
        case 101 /* 'e' era */:
          out = y;
          outl = 1;
          break;
      }
      const outstr = outl > 0 ? pad0(out, outl) : "";
      return outstr;
    }
    /* jshint +W086 */
    function commaify(s) {
      const w = 3;
      if (s.length <= w) return s;
      let j = s.length % w;
      let o = s.substr(0, j);
      for (; j != s.length; j += w)
        o += (o.length > 0 ? "," : "") + s.substr(j, w);
      return o;
    }
    var write_num = (function make_write_num() {
      const pct1 = /%/g;
      function write_num_pct(type, fmt, val) {
        const sfmt = fmt.replace(pct1, "");
        const mul = fmt.length - sfmt.length;
        return (
          write_num(type, sfmt, val * Math.pow(10, 2 * mul)) + fill("%", mul)
        );
      }
      function write_num_cm(type, fmt, val) {
        let idx = fmt.length - 1;
        while (fmt.charCodeAt(idx - 1) === 44) --idx;
        return write_num(
          type,
          fmt.substr(0, idx),
          val / Math.pow(10, 3 * (fmt.length - idx))
        );
      }
      function write_num_exp(fmt, val) {
        let o;
        const idx = fmt.indexOf("E") - fmt.indexOf(".") - 1;
        if (fmt.match(/^#+0.0E\+0$/)) {
          if (val == 0) return "0.0E+0";
          else if (val < 0) return `-${write_num_exp(fmt, -val)}`;
          let period = fmt.indexOf(".");
          if (period === -1) period = fmt.indexOf("E");
          let ee = Math.floor(Math.log(val) * Math.LOG10E) % period;
          if (ee < 0) ee += period;
          o = (val / Math.pow(10, ee)).toPrecision(
            idx + 1 + ((period + ee) % period)
          );
          if (o.indexOf("e") === -1) {
            const fakee = Math.floor(Math.log(val) * Math.LOG10E);
            if (o.indexOf(".") === -1)
              o = `${o.charAt(0)}.${o.substr(1)}E+${fakee - o.length + ee}`;
            else o += `E+${fakee - ee}`;
            while (o.substr(0, 2) === "0.") {
              o = `${o.charAt(0) + o.substr(2, period)}.${o.substr(
                2 + period
              )}`;
              o = o.replace(/^0+([1-9])/, "$1").replace(/^0+\./, "0.");
            }
            o = o.replace(/\+-/, "-");
          }
          o = o.replace(/^([+-]?)(\d*)\.(\d*)[Ee]/, function($$, $1, $2, $3) {
            return `${$1 +
              $2 +
              $3.substr(0, (period + ee) % period)}.${$3.substr(ee)}E`;
          });
        } else o = val.toExponential(idx);
        if (fmt.match(/E\+00$/) && o.match(/e[+-]\d$/))
          o = `${o.substr(0, o.length - 1)}0${o.charAt(o.length - 1)}`;
        if (fmt.match(/E\-/) && o.match(/e\+/)) o = o.replace(/e\+/, "e");
        return o.replace("e", "E");
      }
      const frac1 = /# (\?+)( ?)\/( ?)(\d+)/;
      function write_num_f1(r, aval, sign) {
        const den = parseInt(r[4], 10);
        const rr = Math.round(aval * den);
        const base = Math.floor(rr / den);
        const myn = rr - base * den;
        const myd = den;
        return `${sign + (base === 0 ? "" : `${base}`)} ${
          myn === 0
            ? fill(" ", r[1].length + 1 + r[4].length)
            : `${pad_(myn, r[1].length) + r[2]}/${r[3]}${pad0(
                myd,
                r[4].length
              )}`
        }`;
      }
      function write_num_f2(r, aval, sign) {
        return (
          sign +
          (aval === 0 ? "" : `${aval}`) +
          fill(" ", r[1].length + 2 + r[4].length)
        );
      }
      const dec1 = /^#*0*\.([0#]+)/;
      const closeparen = /\).*[0#]/;
      const phone = /\(###\) ###\\?-####/;
      function hashq(str) {
        let o = "";
        let cc;
        for (let i = 0; i != str.length; ++i)
          switch ((cc = str.charCodeAt(i))) {
            case 35:
              break;
            case 63:
              o += " ";
              break;
            case 48:
              o += "0";
              break;
            default:
              o += String.fromCharCode(cc);
          }
        return o;
      }
      function rnd(val, d) {
        const dd = Math.pow(10, d);
        return `${Math.round(val * dd) / dd}`;
      }
      function dec(val, d) {
        const _frac = val - Math.floor(val);
        const dd = Math.pow(10, d);
        if (d < `${Math.round(_frac * dd)}`.length) return 0;
        return Math.round(_frac * dd);
      }
      function carry(val, d) {
        if (
          d < `${Math.round((val - Math.floor(val)) * Math.pow(10, d))}`.length
        ) {
          return 1;
        }
        return 0;
      }
      function flr(val) {
        if (val < 2147483647 && val > -2147483648)
          return `${val >= 0 ? val | 0 : (val - 1) | 0}`;
        return `${Math.floor(val)}`;
      }
      function write_num_flt(type, fmt, val) {
        if (type.charCodeAt(0) === 40 && !fmt.match(closeparen)) {
          const ffmt = fmt
            .replace(/\( */, "")
            .replace(/ \)/, "")
            .replace(/\)/, "");
          if (val >= 0) return write_num_flt("n", ffmt, val);
          return `(${write_num_flt("n", ffmt, -val)})`;
        }
        if (fmt.charCodeAt(fmt.length - 1) === 44)
          return write_num_cm(type, fmt, val);
        if (fmt.indexOf("%") !== -1) return write_num_pct(type, fmt, val);
        if (fmt.indexOf("E") !== -1) return write_num_exp(fmt, val);
        if (fmt.charCodeAt(0) === 36)
          return `$${write_num_flt(
            type,
            fmt.substr(fmt.charAt(1) == " " ? 2 : 1),
            val
          )}`;
        let o;
        let r;
        let ri;
        let ff;
        const aval = Math.abs(val);
        const sign = val < 0 ? "-" : "";
        if (fmt.match(/^00+$/)) return sign + pad0r(aval, fmt.length);
        if (fmt.match(/^[#?]+$/)) {
          o = pad0r(val, 0);
          if (o === "0") o = "";
          return o.length > fmt.length
            ? o
            : hashq(fmt.substr(0, fmt.length - o.length)) + o;
        }
        if ((r = fmt.match(frac1))) return write_num_f1(r, aval, sign);
        if (fmt.match(/^#+0+$/))
          return sign + pad0r(aval, fmt.length - fmt.indexOf("0"));
        if ((r = fmt.match(dec1))) {
          o = rnd(val, r[1].length)
            .replace(/^([^\.]+)$/, `$1.${hashq(r[1])}`)
            .replace(/\.$/, `.${hashq(r[1])}`)
            .replace(/\.(\d*)$/, function($$, $1) {
              return `.${$1}${fill("0", hashq(r[1]).length - $1.length)}`;
            });
          return fmt.indexOf("0.") !== -1 ? o : o.replace(/^0\./, ".");
        }
        fmt = fmt.replace(/^#+([0.])/, "$1");
        if ((r = fmt.match(/^(0*)\.(#*)$/))) {
          return (
            sign +
            rnd(aval, r[2].length)
              .replace(/\.(\d*[1-9])0*$/, ".$1")
              .replace(/^(-?\d*)$/, "$1.")
              .replace(/^0\./, r[1].length ? "0." : ".")
          );
        }
        if ((r = fmt.match(/^#{1,3},##0(\.?)$/)))
          return sign + commaify(pad0r(aval, 0));
        if ((r = fmt.match(/^#,##0\.([#0]*0)$/))) {
          return val < 0
            ? `-${write_num_flt(type, fmt, -val)}`
            : `${commaify(
                `${Math.floor(val) + carry(val, r[1].length)}`
              )}.${pad0(dec(val, r[1].length), r[1].length)}`;
        }
        if ((r = fmt.match(/^#,#*,#0/)))
          return write_num_flt(type, fmt.replace(/^#,#*,/, ""), val);
        if ((r = fmt.match(/^([0#]+)(\\?-([0#]+))+$/))) {
          o = _strrev(write_num_flt(type, fmt.replace(/[\\-]/g, ""), val));
          ri = 0;
          return _strrev(
            _strrev(fmt.replace(/\\/g, "")).replace(/[0#]/g, function(x) {
              return ri < o.length ? o.charAt(ri++) : x === "0" ? "0" : "";
            })
          );
        }
        if (fmt.match(phone)) {
          o = write_num_flt(type, "##########", val);
          return `(${o.substr(0, 3)}) ${o.substr(3, 3)}-${o.substr(6)}`;
        }
        let oa = "";
        if ((r = fmt.match(/^([#0?]+)( ?)\/( ?)([#0?]+)/))) {
          ri = Math.min(r[4].length, 7);
          ff = frac(aval, Math.pow(10, ri) - 1, false);
          o = `${sign}`;
          oa = write_num("n", r[1], ff[1]);
          if (oa.charAt(oa.length - 1) == " ")
            oa = `${oa.substr(0, oa.length - 1)}0`;
          o += `${oa + r[2]}/${r[3]}`;
          oa = rpad_(ff[2], ri);
          if (oa.length < r[4].length)
            oa = hashq(r[4].substr(r[4].length - oa.length)) + oa;
          o += oa;
          return o;
        }
        if ((r = fmt.match(/^# ([#0?]+)( ?)\/( ?)([#0?]+)/))) {
          ri = Math.min(Math.max(r[1].length, r[4].length), 7);
          ff = frac(aval, Math.pow(10, ri) - 1, true);
          return `${sign + (ff[0] || (ff[1] ? "" : "0"))} ${
            ff[1]
              ? `${pad_(ff[1], ri) + r[2]}/${r[3]}${rpad_(ff[2], ri)}`
              : fill(" ", 2 * ri + 1 + r[2].length + r[3].length)
          }`;
        }
        if ((r = fmt.match(/^[#0?]+$/))) {
          o = pad0r(val, 0);
          if (fmt.length <= o.length) return o;
          return hashq(fmt.substr(0, fmt.length - o.length)) + o;
        }
        if ((r = fmt.match(/^([#0?]+)\.([#0]+)$/))) {
          o = `${val
            .toFixed(Math.min(r[2].length, 10))
            .replace(/([^0])0+$/, "$1")}`;
          ri = o.indexOf(".");
          const lres = fmt.indexOf(".") - ri;
          const rres = fmt.length - o.length - lres;
          return hashq(fmt.substr(0, lres) + o + fmt.substr(fmt.length - rres));
        }
        if ((r = fmt.match(/^00,000\.([#0]*0)$/))) {
          ri = dec(val, r[1].length);
          return val < 0
            ? `-${write_num_flt(type, fmt, -val)}`
            : `${commaify(flr(val))
                .replace(/^\d,\d{3}$/, "0$&")
                .replace(/^\d*$/, function($$) {
                  return `00,${
                    $$.length < 3 ? pad0(0, 3 - $$.length) : ""
                  }${$$}`;
                })}.${pad0(ri, r[1].length)}`;
        }
        switch (fmt) {
          case "###,##0.00":
            return write_num_flt(type, "#,##0.00", val);
          case "###,###":
          case "##,###":
          case "#,###":
            var x = commaify(pad0r(aval, 0));
            return x !== "0" ? sign + x : "";
          case "###,###.00":
            return write_num_flt(type, "###,##0.00", val).replace(/^0\./, ".");
          case "#,###.00":
            return write_num_flt(type, "#,##0.00", val).replace(/^0\./, ".");
          default:
        }
        throw new Error(`unsupported format |${fmt}|`);
      }
      function write_num_cm2(type, fmt, val) {
        let idx = fmt.length - 1;
        while (fmt.charCodeAt(idx - 1) === 44) --idx;
        return write_num(
          type,
          fmt.substr(0, idx),
          val / Math.pow(10, 3 * (fmt.length - idx))
        );
      }
      function write_num_pct2(type, fmt, val) {
        const sfmt = fmt.replace(pct1, "");
        const mul = fmt.length - sfmt.length;
        return (
          write_num(type, sfmt, val * Math.pow(10, 2 * mul)) + fill("%", mul)
        );
      }
      function write_num_exp2(fmt, val) {
        let o;
        const idx = fmt.indexOf("E") - fmt.indexOf(".") - 1;
        if (fmt.match(/^#+0.0E\+0$/)) {
          if (val == 0) return "0.0E+0";
          else if (val < 0) return `-${write_num_exp2(fmt, -val)}`;
          let period = fmt.indexOf(".");
          if (period === -1) period = fmt.indexOf("E");
          let ee = Math.floor(Math.log(val) * Math.LOG10E) % period;
          if (ee < 0) ee += period;
          o = (val / Math.pow(10, ee)).toPrecision(
            idx + 1 + ((period + ee) % period)
          );
          if (!o.match(/[Ee]/)) {
            const fakee = Math.floor(Math.log(val) * Math.LOG10E);
            if (o.indexOf(".") === -1)
              o = `${o.charAt(0)}.${o.substr(1)}E+${fakee - o.length + ee}`;
            else o += `E+${fakee - ee}`;
            o = o.replace(/\+-/, "-");
          }
          o = o.replace(/^([+-]?)(\d*)\.(\d*)[Ee]/, function($$, $1, $2, $3) {
            return `${$1 +
              $2 +
              $3.substr(0, (period + ee) % period)}.${$3.substr(ee)}E`;
          });
        } else o = val.toExponential(idx);
        if (fmt.match(/E\+00$/) && o.match(/e[+-]\d$/))
          o = `${o.substr(0, o.length - 1)}0${o.charAt(o.length - 1)}`;
        if (fmt.match(/E\-/) && o.match(/e\+/)) o = o.replace(/e\+/, "e");
        return o.replace("e", "E");
      }
      function write_num_int(type, fmt, val) {
        if (type.charCodeAt(0) === 40 && !fmt.match(closeparen)) {
          const ffmt = fmt
            .replace(/\( */, "")
            .replace(/ \)/, "")
            .replace(/\)/, "");
          if (val >= 0) return write_num_int("n", ffmt, val);
          return `(${write_num_int("n", ffmt, -val)})`;
        }
        if (fmt.charCodeAt(fmt.length - 1) === 44)
          return write_num_cm2(type, fmt, val);
        if (fmt.indexOf("%") !== -1) return write_num_pct2(type, fmt, val);
        if (fmt.indexOf("E") !== -1) return write_num_exp2(fmt, val);
        if (fmt.charCodeAt(0) === 36)
          return `$${write_num_int(
            type,
            fmt.substr(fmt.charAt(1) == " " ? 2 : 1),
            val
          )}`;
        let o;
        let r;
        let ri;
        let ff;
        const aval = Math.abs(val);
        const sign = val < 0 ? "-" : "";
        if (fmt.match(/^00+$/)) return sign + pad0(aval, fmt.length);
        if (fmt.match(/^[#?]+$/)) {
          o = `${val}`;
          if (val === 0) o = "";
          return o.length > fmt.length
            ? o
            : hashq(fmt.substr(0, fmt.length - o.length)) + o;
        }
        if ((r = fmt.match(frac1))) return write_num_f2(r, aval, sign);
        if (fmt.match(/^#+0+$/))
          return sign + pad0(aval, fmt.length - fmt.indexOf("0"));
        if ((r = fmt.match(dec1))) {
          o = `${val}`
            .replace(/^([^\.]+)$/, `$1.${hashq(r[1])}`)
            .replace(/\.$/, `.${hashq(r[1])}`);
          o = o.replace(/\.(\d*)$/, function($$, $1) {
            return `.${$1}${fill("0", hashq(r[1]).length - $1.length)}`;
          });
          return fmt.indexOf("0.") !== -1 ? o : o.replace(/^0\./, ".");
        }
        fmt = fmt.replace(/^#+([0.])/, "$1");
        if ((r = fmt.match(/^(0*)\.(#*)$/))) {
          return (
            sign +
            `${aval}`
              .replace(/\.(\d*[1-9])0*$/, ".$1")
              .replace(/^(-?\d*)$/, "$1.")
              .replace(/^0\./, r[1].length ? "0." : ".")
          );
        }
        if ((r = fmt.match(/^#{1,3},##0(\.?)$/)))
          return sign + commaify(`${aval}`);
        if ((r = fmt.match(/^#,##0\.([#0]*0)$/))) {
          return val < 0
            ? `-${write_num_int(type, fmt, -val)}`
            : `${commaify(`${val}`)}.${fill("0", r[1].length)}`;
        }
        if ((r = fmt.match(/^#,#*,#0/)))
          return write_num_int(type, fmt.replace(/^#,#*,/, ""), val);
        if ((r = fmt.match(/^([0#]+)(\\?-([0#]+))+$/))) {
          o = _strrev(write_num_int(type, fmt.replace(/[\\-]/g, ""), val));
          ri = 0;
          return _strrev(
            _strrev(fmt.replace(/\\/g, "")).replace(/[0#]/g, function(x) {
              return ri < o.length ? o.charAt(ri++) : x === "0" ? "0" : "";
            })
          );
        }
        if (fmt.match(phone)) {
          o = write_num_int(type, "##########", val);
          return `(${o.substr(0, 3)}) ${o.substr(3, 3)}-${o.substr(6)}`;
        }
        let oa = "";
        if ((r = fmt.match(/^([#0?]+)( ?)\/( ?)([#0?]+)/))) {
          ri = Math.min(r[4].length, 7);
          ff = frac(aval, Math.pow(10, ri) - 1, false);
          o = `${sign}`;
          oa = write_num("n", r[1], ff[1]);
          if (oa.charAt(oa.length - 1) == " ")
            oa = `${oa.substr(0, oa.length - 1)}0`;
          o += `${oa + r[2]}/${r[3]}`;
          oa = rpad_(ff[2], ri);
          if (oa.length < r[4].length)
            oa = hashq(r[4].substr(r[4].length - oa.length)) + oa;
          o += oa;
          return o;
        }
        if ((r = fmt.match(/^# ([#0?]+)( ?)\/( ?)([#0?]+)/))) {
          ri = Math.min(Math.max(r[1].length, r[4].length), 7);
          ff = frac(aval, Math.pow(10, ri) - 1, true);
          return `${sign + (ff[0] || (ff[1] ? "" : "0"))} ${
            ff[1]
              ? `${pad_(ff[1], ri) + r[2]}/${r[3]}${rpad_(ff[2], ri)}`
              : fill(" ", 2 * ri + 1 + r[2].length + r[3].length)
          }`;
        }
        if ((r = fmt.match(/^[#0?]+$/))) {
          o = `${val}`;
          if (fmt.length <= o.length) return o;
          return hashq(fmt.substr(0, fmt.length - o.length)) + o;
        }
        if ((r = fmt.match(/^([#0]+)\.([#0]+)$/))) {
          o = `${val
            .toFixed(Math.min(r[2].length, 10))
            .replace(/([^0])0+$/, "$1")}`;
          ri = o.indexOf(".");
          const lres = fmt.indexOf(".") - ri;
          const rres = fmt.length - o.length - lres;
          return hashq(fmt.substr(0, lres) + o + fmt.substr(fmt.length - rres));
        }
        if ((r = fmt.match(/^00,000\.([#0]*0)$/))) {
          return val < 0
            ? `-${write_num_int(type, fmt, -val)}`
            : `${commaify(`${val}`)
                .replace(/^\d,\d{3}$/, "0$&")
                .replace(/^\d*$/, function($$) {
                  return `00,${
                    $$.length < 3 ? pad0(0, 3 - $$.length) : ""
                  }${$$}`;
                })}.${pad0(0, r[1].length)}`;
        }
        switch (fmt) {
          case "###,###":
          case "##,###":
          case "#,###":
            var x = commaify(`${aval}`);
            return x !== "0" ? sign + x : "";
          default:
            if (fmt.match(/\.[0#?]*$/))
              return (
                write_num_int(type, fmt.slice(0, fmt.lastIndexOf(".")), val) +
                hashq(fmt.slice(fmt.lastIndexOf(".")))
              );
        }
        throw new Error(`unsupported format |${fmt}|`);
      }
      return function write_num(type, fmt, val) {
        return (val | 0) === val
          ? write_num_int(type, fmt, val)
          : write_num_flt(type, fmt, val);
      };
    })();
    function split_fmt(fmt) {
      const out = [];
      let in_str = false; /* , cc */
      for (var i = 0, j = 0; i < fmt.length; ++i)
        switch (/* cc= */ fmt.charCodeAt(i)) {
          case 34 /* '"' */:
            in_str = !in_str;
            break;
          case 95:
          case 42:
          case 92 /* '_' '*' '\\' */:
            ++i;
            break;
          case 59 /* ';' */:
            out[out.length] = fmt.substr(j, i - j);
            j = i + 1;
        }
      out[out.length] = fmt.substr(j);
      if (in_str === true)
        throw new Error(`Format |${fmt}| unterminated string `);
      return out;
    }
    SSF._split = split_fmt;
    const abstime = /\[[HhMmSs\u0E0A\u0E19\u0E17]*\]/;
    function fmt_is_date(fmt) {
      let i = 0;
      /* cc = 0, */ let c = "";
      let o = "";
      while (i < fmt.length) {
        switch ((c = fmt.charAt(i))) {
          case "G":
            if (isgeneral(fmt, i)) i += 6;
            i++;
            break;
          case '"':
            for (; /* cc= */ fmt.charCodeAt(++i) !== 34 && i < fmt.length; ) {
              /* empty */
            }
            ++i;
            break;
          case "\\":
            i += 2;
            break;
          case "_":
            i += 2;
            break;
          case "@":
            ++i;
            break;
          case "B":
          case "b":
            if (fmt.charAt(i + 1) === "1" || fmt.charAt(i + 1) === "2")
              return true;
          /* falls through */
          case "M":
          case "D":
          case "Y":
          case "H":
          case "S":
          case "E":
          /* falls through */
          case "m":
          case "d":
          case "y":
          case "h":
          case "s":
          case "e":
          case "g":
            return true;
          case "A":
          case "a":
          case "上":
            if (fmt.substr(i, 3).toUpperCase() === "A/P") return true;
            if (fmt.substr(i, 5).toUpperCase() === "AM/PM") return true;
            if (fmt.substr(i, 5).toUpperCase() === "上午/下午") return true;
            ++i;
            break;
          case "[":
            o = c;
            while (fmt.charAt(i++) !== "]" && i < fmt.length)
              o += fmt.charAt(i);
            if (o.match(abstime)) return true;
            break;
          case ".":
          /* falls through */
          case "0":
          case "#":
            while (
              i < fmt.length &&
              ("0#?.,E+-%".indexOf((c = fmt.charAt(++i))) > -1 ||
                (c == "\\" &&
                  fmt.charAt(i + 1) == "-" &&
                  "0#".indexOf(fmt.charAt(i + 2)) > -1))
            ) {
              /* empty */
            }
            break;
          case "?":
            while (fmt.charAt(++i) === c) {
              /* empty */
            }
            break;
          case "*":
            ++i;
            if (fmt.charAt(i) == " " || fmt.charAt(i) == "*") ++i;
            break;
          case "(":
          case ")":
            ++i;
            break;
          case "1":
          case "2":
          case "3":
          case "4":
          case "5":
          case "6":
          case "7":
          case "8":
          case "9":
            while (
              i < fmt.length &&
              "0123456789".indexOf(fmt.charAt(++i)) > -1
            ) {
              /* empty */
            }
            break;
          case " ":
            ++i;
            break;
          default:
            ++i;
            break;
        }
      }
      return false;
    }
    SSF.is_date = fmt_is_date;
    function eval_fmt(fmt, v, opts, flen) {
      const out = [];
      let o = "";
      let i = 0;
      let c = "";
      let lst = "t";
      let dt;
      let j;
      let cc;
      let hr = "H";
      /* Tokenize */
      while (i < fmt.length) {
        switch ((c = fmt.charAt(i))) {
          case "G" /* General */:
            if (!isgeneral(fmt, i))
              throw new Error(`unrecognized character ${c} in ${fmt}`);
            out[out.length] = { t: "G", v: "General" };
            i += 7;
            break;
          case '"' /* Literal text */:
            for (o = ""; (cc = fmt.charCodeAt(++i)) !== 34 && i < fmt.length; )
              o += String.fromCharCode(cc);
            out[out.length] = { t: "t", v: o };
            ++i;
            break;
          case "\\":
            var w = fmt.charAt(++i);
            var t = w === "(" || w === ")" ? w : "t";
            out[out.length] = { t, v: w };
            ++i;
            break;
          case "_":
            out[out.length] = { t: "t", v: " " };
            i += 2;
            break;
          case "@" /* Text Placeholder */:
            out[out.length] = { t: "T", v };
            ++i;
            break;
          case "B":
          case "b":
            if (fmt.charAt(i + 1) === "1" || fmt.charAt(i + 1) === "2") {
              if (dt == null) {
                dt = parse_date_code(v, opts, fmt.charAt(i + 1) === "2");
                if (dt == null) return "";
              }
              out[out.length] = { t: "X", v: fmt.substr(i, 2) };
              lst = c;
              i += 2;
              break;
            }
          /* falls through */
          case "M":
          case "D":
          case "Y":
          case "H":
          case "S":
          case "E":
            c = c.toLowerCase();
          /* falls through */
          case "m":
          case "d":
          case "y":
          case "h":
          case "s":
          case "e":
          case "g":
            if (v < 0) return "";
            if (dt == null) {
              dt = parse_date_code(v, opts);
              if (dt == null) return "";
            }
            o = c;
            while (++i < fmt.length && fmt.charAt(i).toLowerCase() === c)
              o += c;
            if (c === "m" && lst.toLowerCase() === "h") c = "M";
            if (c === "h") c = hr;
            out[out.length] = { t: c, v: o };
            lst = c;
            break;
          case "A":
          case "a":
          case "上":
            var q = { t: c, v: c };
            if (dt == null) dt = parse_date_code(v, opts);
            if (fmt.substr(i, 3).toUpperCase() === "A/P") {
              if (dt != null) q.v = dt.H >= 12 ? "P" : "A";
              q.t = "T";
              hr = "h";
              i += 3;
            } else if (fmt.substr(i, 5).toUpperCase() === "AM/PM") {
              if (dt != null) q.v = dt.H >= 12 ? "PM" : "AM";
              q.t = "T";
              i += 5;
              hr = "h";
            } else if (fmt.substr(i, 5).toUpperCase() === "上午/下午") {
              if (dt != null) q.v = dt.H >= 12 ? "下午" : "上午";
              q.t = "T";
              i += 5;
              hr = "h";
            } else {
              q.t = "t";
              ++i;
            }
            if (dt == null && q.t === "T") return "";
            out[out.length] = q;
            lst = c;
            break;
          case "[":
            o = c;
            while (fmt.charAt(i++) !== "]" && i < fmt.length)
              o += fmt.charAt(i);
            if (o.slice(-1) !== "]") throw `unterminated "[" block: |${o}|`;
            if (o.match(abstime)) {
              if (dt == null) {
                dt = parse_date_code(v, opts);
                if (dt == null) return "";
              }
              out[out.length] = { t: "Z", v: o.toLowerCase() };
              lst = o.charAt(1);
            } else if (o.indexOf("$") > -1) {
              o = (o.match(/\$([^-\[\]]*)/) || [])[1] || "$";
              if (!fmt_is_date(fmt)) out[out.length] = { t: "t", v: o };
            }
            break;
          /* Numbers */
          case ".":
            if (dt != null) {
              o = c;
              while (++i < fmt.length && (c = fmt.charAt(i)) === "0") o += c;
              out[out.length] = { t: "s", v: o };
              break;
            }
          /* falls through */
          case "0":
          case "#":
            o = c;
            while (
              ++i < fmt.length &&
              "0#?.,E+-%".indexOf((c = fmt.charAt(i))) > -1
            )
              o += c;
            out[out.length] = { t: "n", v: o };
            break;
          case "?":
            o = c;
            while (fmt.charAt(++i) === c) o += c;
            out[out.length] = { t: c, v: o };
            lst = c;
            break;
          case "*":
            ++i;
            if (fmt.charAt(i) == " " || fmt.charAt(i) == "*") ++i;
            break; // **
          case "(":
          case ")":
            out[out.length] = { t: flen === 1 ? "t" : c, v: c };
            ++i;
            break;
          case "1":
          case "2":
          case "3":
          case "4":
          case "5":
          case "6":
          case "7":
          case "8":
          case "9":
            o = c;
            while (i < fmt.length && "0123456789".indexOf(fmt.charAt(++i)) > -1)
              o += fmt.charAt(i);
            out[out.length] = { t: "D", v: o };
            break;
          case " ":
            out[out.length] = { t: c, v: c };
            ++i;
            break;
          case "$":
            out[out.length] = { t: "t", v: "$" };
            ++i;
            break;
          default:
            if (",$-+/():!^&'~{}<>=€acfijklopqrtuvwxzP".indexOf(c) === -1)
              throw new Error(`unrecognized character ${c} in ${fmt}`);
            out[out.length] = { t: "t", v: c };
            ++i;
            break;
        }
      }

      /* Scan for date/time parts */
      let bt = 0;
      let ss0 = 0;
      let ssm;
      for (i = out.length - 1, lst = "t"; i >= 0; --i) {
        switch (out[i].t) {
          case "h":
          case "H":
            out[i].t = hr;
            lst = "h";
            if (bt < 1) bt = 1;
            break;
          case "s":
            if ((ssm = out[i].v.match(/\.0+$/)))
              ss0 = Math.max(ss0, ssm[0].length - 1);
            if (bt < 3) bt = 3;
          /* falls through */
          case "d":
          case "y":
          case "M":
          case "e":
            lst = out[i].t;
            break;
          case "m":
            if (lst === "s") {
              out[i].t = "M";
              if (bt < 2) bt = 2;
            }
            break;
          case "X" /* if(out[i].v === "B2"); */:
            break;
          case "Z":
            if (bt < 1 && out[i].v.match(/[Hh]/)) bt = 1;
            if (bt < 2 && out[i].v.match(/[Mm]/)) bt = 2;
            if (bt < 3 && out[i].v.match(/[Ss]/)) bt = 3;
        }
      }
      /* time rounding depends on presence of minute / second / usec fields */
      switch (bt) {
        case 0:
          break;
        case 1:
          if (dt.u >= 0.5) {
            dt.u = 0;
            ++dt.S;
          }
          if (dt.S >= 60) {
            dt.S = 0;
            ++dt.M;
          }
          if (dt.M >= 60) {
            dt.M = 0;
            ++dt.H;
          }
          break;
        case 2:
          if (dt.u >= 0.5) {
            dt.u = 0;
            ++dt.S;
          }
          if (dt.S >= 60) {
            dt.S = 0;
            ++dt.M;
          }
          break;
      }

      /* replace fields */
      let nstr = "";
      let jj;
      for (i = 0; i < out.length; ++i) {
        switch (out[i].t) {
          case "t":
          case "T":
          case " ":
          case "D":
            break;
          case "X":
            out[i].v = "";
            out[i].t = ";";
            break;
          case "d":
          case "m":
          case "y":
          case "h":
          case "H":
          case "M":
          case "s":
          case "e":
          case "b":
          case "Z":
            out[i].v = write_date(out[i].t.charCodeAt(0), out[i].v, dt, ss0);
            out[i].t = "t";
            break;
          case "n":
          case "?":
            jj = i + 1;
            while (
              out[jj] != null &&
              ((c = out[jj].t) === "?" ||
                c === "D" ||
                ((c === " " || c === "t") &&
                  out[jj + 1] != null &&
                  (out[jj + 1].t === "?" ||
                    (out[jj + 1].t === "t" && out[jj + 1].v === "/"))) ||
                (out[i].t === "(" && (c === " " || c === "n" || c === ")")) ||
                (c === "t" &&
                  (out[jj].v === "/" ||
                    (out[jj].v === " " &&
                      out[jj + 1] != null &&
                      out[jj + 1].t == "?"))))
            ) {
              out[i].v += out[jj].v;
              out[jj] = { v: "", t: ";" };
              ++jj;
            }
            nstr += out[i].v;
            i = jj - 1;
            break;
          case "G":
            out[i].t = "t";
            out[i].v = general_fmt(v, opts);
            break;
        }
      }
      let vv = "";
      let myv;
      let ostr;
      if (nstr.length > 0) {
        if (nstr.charCodeAt(0) == 40) {
          /* '(' */ myv = v < 0 && nstr.charCodeAt(0) === 45 ? -v : v;
          ostr = write_num("n", nstr, myv);
        } else {
          myv = v < 0 && flen > 1 ? -v : v;
          ostr = write_num("n", nstr, myv);
          if (myv < 0 && out[0] && out[0].t == "t") {
            ostr = ostr.substr(1);
            out[0].v = `-${out[0].v}`;
          }
        }
        jj = ostr.length - 1;
        let decpt = out.length;
        for (i = 0; i < out.length; ++i)
          if (out[i] != null && out[i].t != "t" && out[i].v.indexOf(".") > -1) {
            decpt = i;
            break;
          }
        let lasti = out.length;
        if (decpt === out.length && ostr.indexOf("E") === -1) {
          for (i = out.length - 1; i >= 0; --i) {
            if (out[i] == null || "n?".indexOf(out[i].t) === -1) continue;
            if (jj >= out[i].v.length - 1) {
              jj -= out[i].v.length;
              out[i].v = ostr.substr(jj + 1, out[i].v.length);
            } else if (jj < 0) out[i].v = "";
            else {
              out[i].v = ostr.substr(0, jj + 1);
              jj = -1;
            }
            out[i].t = "t";
            lasti = i;
          }
          if (jj >= 0 && lasti < out.length)
            out[lasti].v = ostr.substr(0, jj + 1) + out[lasti].v;
        } else if (decpt !== out.length && ostr.indexOf("E") === -1) {
          jj = ostr.indexOf(".") - 1;
          for (i = decpt; i >= 0; --i) {
            if (out[i] == null || "n?".indexOf(out[i].t) === -1) continue;
            j =
              out[i].v.indexOf(".") > -1 && i === decpt
                ? out[i].v.indexOf(".") - 1
                : out[i].v.length - 1;
            vv = out[i].v.substr(j + 1);
            for (; j >= 0; --j) {
              if (
                jj >= 0 &&
                (out[i].v.charAt(j) === "0" || out[i].v.charAt(j) === "#")
              )
                vv = ostr.charAt(jj--) + vv;
            }
            out[i].v = vv;
            out[i].t = "t";
            lasti = i;
          }
          if (jj >= 0 && lasti < out.length)
            out[lasti].v = ostr.substr(0, jj + 1) + out[lasti].v;
          jj = ostr.indexOf(".") + 1;
          for (i = decpt; i < out.length; ++i) {
            if (
              out[i] == null ||
              ("n?(".indexOf(out[i].t) === -1 && i !== decpt)
            )
              continue;
            j =
              out[i].v.indexOf(".") > -1 && i === decpt
                ? out[i].v.indexOf(".") + 1
                : 0;
            vv = out[i].v.substr(0, j);
            for (; j < out[i].v.length; ++j) {
              if (jj < ostr.length) vv += ostr.charAt(jj++);
            }
            out[i].v = vv;
            out[i].t = "t";
            lasti = i;
          }
        }
      }
      for (i = 0; i < out.length; ++i)
        if (out[i] != null && "n?".indexOf(out[i].t) > -1) {
          myv = flen > 1 && v < 0 && i > 0 && out[i - 1].v === "-" ? -v : v;
          out[i].v = write_num(out[i].t, out[i].v, myv);
          out[i].t = "t";
        }
      let retval = "";
      for (i = 0; i !== out.length; ++i) if (out[i] != null) retval += out[i].v;
      return retval;
    }
    SSF._eval = eval_fmt;
    const cfregex = /\[[=<>]/;
    const cfregex2 = /\[(=|>[=]?|<[>=]?)(-?\d+(?:\.\d*)?)\]/;
    function chkcond(v, rr) {
      if (rr == null) return false;
      const thresh = parseFloat(rr[2]);
      switch (rr[1]) {
        case "=":
          if (v == thresh) return true;
          break;
        case ">":
          if (v > thresh) return true;
          break;
        case "<":
          if (v < thresh) return true;
          break;
        case "<>":
          if (v != thresh) return true;
          break;
        case ">=":
          if (v >= thresh) return true;
          break;
        case "<=":
          if (v <= thresh) return true;
          break;
      }
      return false;
    }
    function choose_fmt(f, v) {
      let fmt = split_fmt(f);
      let l = fmt.length;
      const lat = fmt[l - 1].indexOf("@");
      if (l < 4 && lat > -1) --l;
      if (fmt.length > 4)
        throw new Error(`cannot find right format for |${fmt.join("|")}|`);
      if (typeof v !== "number")
        return [4, fmt.length === 4 || lat > -1 ? fmt[fmt.length - 1] : "@"];
      switch (fmt.length) {
        case 1:
          fmt =
            lat > -1
              ? ["General", "General", "General", fmt[0]]
              : [fmt[0], fmt[0], fmt[0], "@"];
          break;
        case 2:
          fmt =
            lat > -1
              ? [fmt[0], fmt[0], fmt[0], fmt[1]]
              : [fmt[0], fmt[1], fmt[0], "@"];
          break;
        case 3:
          fmt =
            lat > -1
              ? [fmt[0], fmt[1], fmt[0], fmt[2]]
              : [fmt[0], fmt[1], fmt[2], "@"];
          break;
        case 4:
          break;
      }
      const ff = v > 0 ? fmt[0] : v < 0 ? fmt[1] : fmt[2];
      if (fmt[0].indexOf("[") === -1 && fmt[1].indexOf("[") === -1)
        return [l, ff];
      if (fmt[0].match(cfregex) != null || fmt[1].match(cfregex) != null) {
        const m1 = fmt[0].match(cfregex2);
        const m2 = fmt[1].match(cfregex2);
        return chkcond(v, m1)
          ? [l, fmt[0]]
          : chkcond(v, m2)
          ? [l, fmt[1]]
          : [l, fmt[m1 != null && m2 != null ? 2 : 1]];
      }
      return [l, ff];
    }
    function format(fmt, v, o) {
      if (o == null) o = {};
      let sfmt = "";
      switch (typeof fmt) {
        case "string":
          if (fmt == "m/d/yy" && o.dateNF) sfmt = o.dateNF;
          else sfmt = fmt;
          break;
        case "number":
          if (fmt == 14 && o.dateNF) sfmt = o.dateNF;
          else sfmt = (o.table != null ? o.table : table_fmt)[fmt];
          if (sfmt == null)
            sfmt =
              (o.table && o.table[default_map[fmt]]) ||
              table_fmt[default_map[fmt]];
          if (sfmt == null) sfmt = default_str[fmt] || "General";
          break;
      }
      if (isgeneral(sfmt, 0)) return general_fmt(v, o);
      if (v instanceof Date) v = datenum_local(v, o.date1904);
      const f = choose_fmt(sfmt, v);
      if (isgeneral(f[1])) return general_fmt(v, o);
      if (v === true) v = "TRUE";
      else if (v === false) v = "FALSE";
      else if (v === "" || v == null) return "";
      return eval_fmt(f[1], v, o, f[0]);
    }
    function load_entry(fmt, idx) {
      if (typeof idx !== "number") {
        idx = +idx || -1;
        for (let i = 0; i < 0x0188; ++i) {
          if (table_fmt[i] == undefined) {
            if (idx < 0) idx = i;
            continue;
          }
          if (table_fmt[i] == fmt) {
            idx = i;
            break;
          }
        }
        if (idx < 0) idx = 0x187;
      }
      table_fmt[idx] = fmt;
      return idx;
    }
    SSF.load = load_entry;
    SSF._table = table_fmt;
    SSF.get_table = function get_table() {
      return table_fmt;
    };
    SSF.load_table = function load_table(tbl) {
      for (let i = 0; i != 0x0188; ++i)
        if (tbl[i] !== undefined) load_entry(tbl[i], i);
    };
    SSF.init_table = init_table;
    SSF.format = format;
  };
  make_ssf(SSF);
  /* map from xlml named formats to SSF TODO: localize */
  const XLMLFormatMap /* {[string]:string} */ = {
    "General Number": "General",
    "General Date": SSF._table[22],
    "Long Date": "dddd, mmmm dd, yyyy",
    "Medium Date": SSF._table[15],
    "Short Date": SSF._table[14],
    "Long Time": SSF._table[19],
    "Medium Time": SSF._table[18],
    "Short Time": SSF._table[20],
    Currency: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
    Fixed: SSF._table[2],
    Standard: SSF._table[4],
    Percent: SSF._table[10],
    Scientific: SSF._table[11],
    "Yes/No": '"Yes";"Yes";"No";@',
    "True/False": '"True";"True";"False";@',
    "On/Off": '"Yes";"Yes";"No";@'
  };

  const SSFImplicit /* {[number]:string} */ = {
    "5": '"$"#,##0_);\\("$"#,##0\\)',
    "6": '"$"#,##0_);[Red]\\("$"#,##0\\)',
    "7": '"$"#,##0.00_);\\("$"#,##0.00\\)',
    "8": '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
    "23": "General",
    "24": "General",
    "25": "General",
    "26": "General",
    "27": "m/d/yy",
    "28": "m/d/yy",
    "29": "m/d/yy",
    "30": "m/d/yy",
    "31": "m/d/yy",
    "32": "h:mm:ss",
    "33": "h:mm:ss",
    "34": "h:mm:ss",
    "35": "h:mm:ss",
    "36": "m/d/yy",
    "41": '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
    "42": '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
    "43": '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
    "44": '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
    "50": "m/d/yy",
    "51": "m/d/yy",
    "52": "m/d/yy",
    "53": "m/d/yy",
    "54": "m/d/yy",
    "55": "m/d/yy",
    "56": "m/d/yy",
    "57": "m/d/yy",
    "58": "m/d/yy",
    "59": "0",
    "60": "0.00",
    "61": "#,##0",
    "62": "#,##0.00",
    "63": '"$"#,##0_);\\("$"#,##0\\)',
    "64": '"$"#,##0_);[Red]\\("$"#,##0\\)',
    "65": '"$"#,##0.00_);\\("$"#,##0.00\\)',
    "66": '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
    "67": "0%",
    "68": "0.00%",
    "69": "# ?/?",
    "70": "# ??/??",
    "71": "m/d/yy",
    "72": "m/d/yy",
    "73": "d-mmm-yy",
    "74": "d-mmm",
    "75": "mmm-yy",
    "76": "h:mm",
    "77": "h:mm:ss",
    "78": "m/d/yy h:mm",
    "79": "mm:ss",
    "80": "[h]:mm:ss",
    "81": "mmss.0"
  };

  /* dateNF parse TODO: move to SSF */
  const dateNFregex = /[dD]+|[mM]+|[yYeE]+|[Hh]+|[Ss]+/g;
  function dateNF_regex(dateNF) {
    let fmt = typeof dateNF === "number" ? SSF._table[dateNF] : dateNF;
    fmt = fmt.replace(dateNFregex, "(\\d+)");
    return new RegExp(`^${fmt}$`);
  }
  function dateNF_fix(str, dateNF, match) {
    let Y = -1;
    let m = -1;
    let d = -1;
    let H = -1;
    let M = -1;
    let S = -1;
    (dateNF.match(dateNFregex) || []).forEach(function(n, i) {
      const v = parseInt(match[i + 1], 10);
      switch (n.toLowerCase().charAt(0)) {
        case "y":
          Y = v;
          break;
        case "d":
          d = v;
          break;
        case "h":
          H = v;
          break;
        case "s":
          S = v;
          break;
        case "m":
          if (H >= 0) M = v;
          else m = v;
          break;
      }
    });
    if (S >= 0 && M == -1 && m >= 0) {
      M = m;
      m = -1;
    }
    let datestr = `${`${Y >= 0 ? Y : new Date().getFullYear()}`.slice(
      -4
    )}-${`00${m >= 1 ? m : 1}`.slice(-2)}-${`00${d >= 1 ? d : 1}`.slice(-2)}`;
    if (datestr.length == 7) datestr = `0${datestr}`;
    if (datestr.length == 8) datestr = `20${datestr}`;
    const timestr = `${`00${H >= 0 ? H : 0}`.slice(-2)}:${`00${
      M >= 0 ? M : 0
    }`.slice(-2)}:${`00${S >= 0 ? S : 0}`.slice(-2)}`;
    if (H == -1 && M == -1 && S == -1) return datestr;
    if (Y == -1 && m == -1 && d == -1) return timestr;
    return `${datestr}T${timestr}`;
  }

  const DO_NOT_EXPORT_CFB = true;
  /* cfb.js (C) 2013-present SheetJS -- http://sheetjs.com */
  /* vim: set ts=2: */
  /* jshint eqnull:true */
  /* exported CFB */
  /* global Uint8Array:false, Uint16Array:false */

  /* crc32.js (C) 2014-present SheetJS -- http://sheetjs.com */
  /* vim: set ts=2: */
  /* exported CRC32 */
  let CRC32;
  (function(factory) {
    /* jshint ignore:start */
    /*eslint-disable */
    factory((CRC32 = {}));
    /* eslint-enable */
    /* jshint ignore:end */
  })(function(CRC32) {
    CRC32.version = "1.2.0";
    /* see perf/crc32table.js */
    /* global Int32Array */
    function signed_crc_table() {
      let c = 0;
      const table = new Array(256);

      for (let n = 0; n != 256; ++n) {
        c = n;
        c = c & 1 ? -306674912 ^ (c >>> 1) : c >>> 1;
        c = c & 1 ? -306674912 ^ (c >>> 1) : c >>> 1;
        c = c & 1 ? -306674912 ^ (c >>> 1) : c >>> 1;
        c = c & 1 ? -306674912 ^ (c >>> 1) : c >>> 1;
        c = c & 1 ? -306674912 ^ (c >>> 1) : c >>> 1;
        c = c & 1 ? -306674912 ^ (c >>> 1) : c >>> 1;
        c = c & 1 ? -306674912 ^ (c >>> 1) : c >>> 1;
        c = c & 1 ? -306674912 ^ (c >>> 1) : c >>> 1;
        table[n] = c;
      }

      return typeof Int32Array !== "undefined" ? new Int32Array(table) : table;
    }

    const T = signed_crc_table();
    function crc32_bstr(bstr, seed) {
      let C = seed ^ -1;
      const L = bstr.length - 1;
      for (var i = 0; i < L; ) {
        C = (C >>> 8) ^ T[(C ^ bstr.charCodeAt(i++)) & 0xff];
        C = (C >>> 8) ^ T[(C ^ bstr.charCodeAt(i++)) & 0xff];
      }
      if (i === L) C = (C >>> 8) ^ T[(C ^ bstr.charCodeAt(i)) & 0xff];
      return C ^ -1;
    }

    function crc32_buf(buf, seed) {
      if (buf.length > 10000) return crc32_buf_8(buf, seed);
      let C = seed ^ -1;
      const L = buf.length - 3;
      for (var i = 0; i < L; ) {
        C = (C >>> 8) ^ T[(C ^ buf[i++]) & 0xff];
        C = (C >>> 8) ^ T[(C ^ buf[i++]) & 0xff];
        C = (C >>> 8) ^ T[(C ^ buf[i++]) & 0xff];
        C = (C >>> 8) ^ T[(C ^ buf[i++]) & 0xff];
      }
      while (i < L + 3) C = (C >>> 8) ^ T[(C ^ buf[i++]) & 0xff];
      return C ^ -1;
    }

    function crc32_buf_8(buf, seed) {
      let C = seed ^ -1;
      const L = buf.length - 7;
      for (var i = 0; i < L; ) {
        C = (C >>> 8) ^ T[(C ^ buf[i++]) & 0xff];
        C = (C >>> 8) ^ T[(C ^ buf[i++]) & 0xff];
        C = (C >>> 8) ^ T[(C ^ buf[i++]) & 0xff];
        C = (C >>> 8) ^ T[(C ^ buf[i++]) & 0xff];
        C = (C >>> 8) ^ T[(C ^ buf[i++]) & 0xff];
        C = (C >>> 8) ^ T[(C ^ buf[i++]) & 0xff];
        C = (C >>> 8) ^ T[(C ^ buf[i++]) & 0xff];
        C = (C >>> 8) ^ T[(C ^ buf[i++]) & 0xff];
      }
      while (i < L + 7) C = (C >>> 8) ^ T[(C ^ buf[i++]) & 0xff];
      return C ^ -1;
    }

    function crc32_str(str, seed) {
      let C = seed ^ -1;
      for (var i = 0, L = str.length, c, d; i < L; ) {
        c = str.charCodeAt(i++);
        if (c < 0x80) {
          C = (C >>> 8) ^ T[(C ^ c) & 0xff];
        } else if (c < 0x800) {
          C = (C >>> 8) ^ T[(C ^ (192 | ((c >> 6) & 31))) & 0xff];
          C = (C >>> 8) ^ T[(C ^ (128 | (c & 63))) & 0xff];
        } else if (c >= 0xd800 && c < 0xe000) {
          c = (c & 1023) + 64;
          d = str.charCodeAt(i++) & 1023;
          C = (C >>> 8) ^ T[(C ^ (240 | ((c >> 8) & 7))) & 0xff];
          C = (C >>> 8) ^ T[(C ^ (128 | ((c >> 2) & 63))) & 0xff];
          C =
            (C >>> 8) ^
            T[(C ^ (128 | ((d >> 6) & 15) | ((c & 3) << 4))) & 0xff];
          C = (C >>> 8) ^ T[(C ^ (128 | (d & 63))) & 0xff];
        } else {
          C = (C >>> 8) ^ T[(C ^ (224 | ((c >> 12) & 15))) & 0xff];
          C = (C >>> 8) ^ T[(C ^ (128 | ((c >> 6) & 63))) & 0xff];
          C = (C >>> 8) ^ T[(C ^ (128 | (c & 63))) & 0xff];
        }
      }
      return C ^ -1;
    }
    CRC32.table = T;
    CRC32.bstr = crc32_bstr;
    CRC32.buf = crc32_buf;
    CRC32.str = crc32_str;
  });
  /* [MS-CFB] v20171201 */
  var CFB = (function _CFB() {
    const exports = {};
    exports.version = "1.1.4";
    /* [MS-CFB] 2.6.4 */
    function namecmp(l, r) {
      const L = l.split("/");
      const R = r.split("/");
      for (let i = 0, c = 0, Z = Math.min(L.length, R.length); i < Z; ++i) {
        if ((c = L[i].length - R[i].length)) return c;
        if (L[i] != R[i]) return L[i] < R[i] ? -1 : 1;
      }
      return L.length - R.length;
    }
    function dirname(p) {
      if (p.charAt(p.length - 1) == "/")
        return p.slice(0, -1).indexOf("/") === -1 ? p : dirname(p.slice(0, -1));
      const c = p.lastIndexOf("/");
      return c === -1 ? p : p.slice(0, c + 1);
    }

    function filename(p) {
      if (p.charAt(p.length - 1) == "/") return filename(p.slice(0, -1));
      const c = p.lastIndexOf("/");
      return c === -1 ? p : p.slice(c + 1);
    }
    /* -------------------------------------------------------------------------- */
    /* DOS Date format:
   high|YYYYYYYm.mmmddddd.HHHHHMMM.MMMSSSSS|low
   add 1980 to stored year
   stored second should be doubled
*/

    /* write JS date to buf as a DOS date */
    function write_dos_date(buf, date) {
      if (typeof date === "string") date = new Date(date);
      let hms = date.getHours();
      hms = (hms << 6) | date.getMinutes();
      hms = (hms << 5) | (date.getSeconds() >>> 1);
      buf.write_shift(2, hms);
      let ymd = date.getFullYear() - 1980;
      ymd = (ymd << 4) | (date.getMonth() + 1);
      ymd = (ymd << 5) | date.getDate();
      buf.write_shift(2, ymd);
    }

    /* read four bytes from buf and interpret as a DOS date */
    function parse_dos_date(buf) {
      let hms = buf.read_shift(2) & 0xffff;
      let ymd = buf.read_shift(2) & 0xffff;
      const val = new Date();
      const d = ymd & 0x1f;
      ymd >>>= 5;
      const m = ymd & 0x0f;
      ymd >>>= 4;
      val.setMilliseconds(0);
      val.setFullYear(ymd + 1980);
      val.setMonth(m - 1);
      val.setDate(d);
      const S = hms & 0x1f;
      hms >>>= 5;
      const M = hms & 0x3f;
      hms >>>= 6;
      val.setHours(hms);
      val.setMinutes(M);
      val.setSeconds(S << 1);
      return val;
    }
    function parse_extra_field(blob) {
      prep_blob(blob, 0);
      const o = {};
      let flags = 0;
      while (blob.l <= blob.length - 4) {
        const type = blob.read_shift(2);
        const sz = blob.read_shift(2);
        const tgt = blob.l + sz;
        const p = {};
        switch (type) {
          /* UNIX-style Timestamps */
          case 0x5455:
            {
              flags = blob.read_shift(1);
              if (flags & 1) p.mtime = blob.read_shift(4);
              /* for some reason, CD flag corresponds to LFH */
              if (sz > 5) {
                if (flags & 2) p.atime = blob.read_shift(4);
                if (flags & 4) p.ctime = blob.read_shift(4);
              }
              if (p.mtime) p.mt = new Date(p.mtime * 1000);
            }
            break;
        }
        blob.l = tgt;
        o[type] = p;
      }
      return o;
    }
    let fs;
    function get_fs() {
      return fs || (fs = require("fs"));
    }
    function parse(file, options) {
      if (file[0] == 0x50 && file[1] == 0x4b) return parse_zip(file, options);
      if (file.length < 512)
        throw new Error(`CFB file size ${file.length} < 512`);
      let mver = 3;
      let ssz = 512;
      let nmfs = 0; // number of mini FAT sectors
      let difat_sec_cnt = 0;
      let dir_start = 0;
      let minifat_start = 0;
      let difat_start = 0;

      const fat_addrs = []; // locations of FAT sectors

      /* [MS-CFB] 2.2 Compound File Header */
      let blob = file.slice(0, 512);
      prep_blob(blob, 0);

      /* major version */
      const mv = check_get_mver(blob);
      mver = mv[0];
      switch (mver) {
        case 3:
          ssz = 512;
          break;
        case 4:
          ssz = 4096;
          break;
        case 0:
          if (mv[1] == 0) return parse_zip(file, options);
        /* falls through */
        default:
          throw new Error(`Major Version: Expected 3 or 4 saw ${mver}`);
      }

      /* reprocess header */
      if (ssz !== 512) {
        blob = file.slice(0, ssz);
        prep_blob(blob, 28 /* blob.l */);
      }
      /* Save header for final object */
      const header = file.slice(0, ssz);

      check_shifts(blob, mver);

      // Number of Directory Sectors
      const dir_cnt = blob.read_shift(4, "i");
      if (mver === 3 && dir_cnt !== 0)
        throw new Error(`# Directory Sectors: Expected 0 saw ${dir_cnt}`);

      // Number of FAT Sectors
      blob.l += 4;

      // First Directory Sector Location
      dir_start = blob.read_shift(4, "i");

      // Transaction Signature
      blob.l += 4;

      // Mini Stream Cutoff Size
      blob.chk("00100000", "Mini Stream Cutoff Size: ");

      // First Mini FAT Sector Location
      minifat_start = blob.read_shift(4, "i");

      // Number of Mini FAT Sectors
      nmfs = blob.read_shift(4, "i");

      // First DIFAT sector location
      difat_start = blob.read_shift(4, "i");

      // Number of DIFAT Sectors
      difat_sec_cnt = blob.read_shift(4, "i");

      // Grab FAT Sector Locations
      for (let q = -1, j = 0; j < 109; ++j) {
        /* 109 = (512 - blob.l)>>>2; */
        q = blob.read_shift(4, "i");
        if (q < 0) break;
        fat_addrs[j] = q;
      }

      /** Break the file up into sectors */
      const sectors = sectorify(file, ssz);

      sleuth_fat(difat_start, difat_sec_cnt, sectors, ssz, fat_addrs);

      /** Chains */
      const sector_list = make_sector_list(sectors, dir_start, fat_addrs, ssz);

      sector_list[dir_start].name = "!Directory";
      if (nmfs > 0 && minifat_start !== ENDOFCHAIN)
        sector_list[minifat_start].name = "!MiniFAT";
      sector_list[fat_addrs[0]].name = "!FAT";
      sector_list.fat_addrs = fat_addrs;
      sector_list.ssz = ssz;

      /* [MS-CFB] 2.6.1 Compound File Directory Entry */
      const files = {};
      const Paths = [];
      const FileIndex = [];
      const FullPaths = [];
      read_directory(
        dir_start,
        sector_list,
        sectors,
        Paths,
        nmfs,
        files,
        FileIndex,
        minifat_start
      );

      build_full_paths(FileIndex, FullPaths, Paths);
      Paths.shift();

      const o = {
        FileIndex,
        FullPaths
      };

      // $FlowIgnore
      if (options && options.raw) o.raw = { header, sectors };
      return o;
    } // parse

    /* [MS-CFB] 2.2 Compound File Header -- read up to major version */
    function check_get_mver(blob) {
      if (blob[blob.l] == 0x50 && blob[blob.l + 1] == 0x4b) return [0, 0];
      // header signature 8
      blob.chk(HEADER_SIGNATURE, "Header Signature: ");

      // clsid 16
      // blob.chk(HEADER_CLSID, 'CLSID: ');
      blob.l += 16;

      // minor version 2
      const mver = blob.read_shift(2, "u");

      return [blob.read_shift(2, "u"), mver];
    }
    function check_shifts(blob, mver) {
      let shift = 0x09;

      // Byte Order
      // blob.chk('feff', 'Byte Order: '); // note: some writers put 0xffff
      blob.l += 2;

      // Sector Shift
      switch ((shift = blob.read_shift(2))) {
        case 0x09:
          if (mver != 3)
            throw new Error(`Sector Shift: Expected 9 saw ${shift}`);
          break;
        case 0x0c:
          if (mver != 4)
            throw new Error(`Sector Shift: Expected 12 saw ${shift}`);
          break;
        default:
          throw new Error(`Sector Shift: Expected 9 or 12 saw ${shift}`);
      }

      // Mini Sector Shift
      blob.chk("0600", "Mini Sector Shift: ");

      // Reserved
      blob.chk("000000000000", "Reserved: ");
    }

    /** Break the file up into sectors */
    function sectorify(file, ssz) {
      const nsectors = Math.ceil(file.length / ssz) - 1;
      const sectors = [];
      for (let i = 1; i < nsectors; ++i)
        sectors[i - 1] = file.slice(i * ssz, (i + 1) * ssz);
      sectors[nsectors - 1] = file.slice(nsectors * ssz);
      return sectors;
    }

    /* [MS-CFB] 2.6.4 Red-Black Tree */
    function build_full_paths(FI, FP, Paths) {
      let i = 0;
      let L = 0;
      let R = 0;
      let C = 0;
      let j = 0;
      const pl = Paths.length;
      const dad = [];
      const q = [];

      for (; i < pl; ++i) {
        dad[i] = q[i] = i;
        FP[i] = Paths[i];
      }

      for (; j < q.length; ++j) {
        i = q[j];
        L = FI[i].L;
        R = FI[i].R;
        C = FI[i].C;
        if (dad[i] === i) {
          if (L !== -1 /* NOSTREAM */ && dad[L] !== L) dad[i] = dad[L];
          if (R !== -1 && dad[R] !== R) dad[i] = dad[R];
        }
        if (C !== -1 /* NOSTREAM */) dad[C] = i;
        if (L !== -1 && i != dad[i]) {
          dad[L] = dad[i];
          if (q.lastIndexOf(L) < j) q.push(L);
        }
        if (R !== -1 && i != dad[i]) {
          dad[R] = dad[i];
          if (q.lastIndexOf(R) < j) q.push(R);
        }
      }
      for (i = 1; i < pl; ++i)
        if (dad[i] === i) {
          if (R !== -1 /* NOSTREAM */ && dad[R] !== R) dad[i] = dad[R];
          else if (L !== -1 && dad[L] !== L) dad[i] = dad[L];
        }

      for (i = 1; i < pl; ++i) {
        if (FI[i].type === 0 /* unknown */) continue;
        j = i;
        if (j != dad[j])
          do {
            j = dad[j];
            FP[i] = `${FP[j]}/${FP[i]}`;
          } while (j !== 0 && dad[j] !== -1 && j != dad[j]);
        dad[i] = -1;
      }

      FP[0] += "/";
      for (i = 1; i < pl; ++i) {
        if (FI[i].type !== 2 /* stream */) FP[i] += "/";
      }
    }

    function get_mfat_entry(entry, payload, mini) {
      const { start } = entry;
      let { size } = entry;
      // return (payload.slice(start*MSSZ, start*MSSZ + size));
      const o = [];
      let idx = start;
      while (mini && size > 0 && idx >= 0) {
        o.push(payload.slice(idx * MSSZ, idx * MSSZ + MSSZ));
        size -= MSSZ;
        idx = __readInt32LE(mini, idx * 4);
      }
      if (o.length === 0) return new_buf(0);
      return bconcat(o).slice(0, entry.size);
    }

    /** Chase down the rest of the DIFAT chain to build a comprehensive list
     DIFAT chains by storing the next sector number as the last 32 bits */
    function sleuth_fat(idx, cnt, sectors, ssz, fat_addrs) {
      let q = ENDOFCHAIN;
      if (idx === ENDOFCHAIN) {
        if (cnt !== 0) throw new Error("DIFAT chain shorter than expected");
      } else if (idx !== -1 /* FREESECT */) {
        const sector = sectors[idx];
        const m = (ssz >>> 2) - 1;
        if (!sector) return;
        for (let i = 0; i < m; ++i) {
          if ((q = __readInt32LE(sector, i * 4)) === ENDOFCHAIN) break;
          fat_addrs.push(q);
        }
        sleuth_fat(
          __readInt32LE(sector, ssz - 4),
          cnt - 1,
          sectors,
          ssz,
          fat_addrs
        );
      }
    }

    /** Follow the linked list of sectors for a given starting point */
    function get_sector_list(sectors, start, fat_addrs, ssz, chkd) {
      const buf = [];
      const buf_chain = [];
      if (!chkd) chkd = [];
      const modulus = ssz - 1;
      let j = 0;
      let jj = 0;
      for (j = start; j >= 0; ) {
        chkd[j] = true;
        buf[buf.length] = j;
        buf_chain.push(sectors[j]);
        const addr = fat_addrs[Math.floor((j * 4) / ssz)];
        jj = (j * 4) & modulus;
        if (ssz < 4 + jj)
          throw new Error(`FAT boundary crossed: ${j} 4 ${ssz}`);
        if (!sectors[addr]) break;
        j = __readInt32LE(sectors[addr], jj);
      }
      return { nodes: buf, data: __toBuffer([buf_chain]) };
    }

    /** Chase down the sector linked lists */
    function make_sector_list(sectors, dir_start, fat_addrs, ssz) {
      const sl = sectors.length;
      const sector_list = [];
      const chkd = [];
      let buf = [];
      let buf_chain = [];
      const modulus = ssz - 1;
      let i = 0;
      let j = 0;
      let k = 0;
      let jj = 0;
      for (i = 0; i < sl; ++i) {
        buf = [];
        k = i + dir_start;
        if (k >= sl) k -= sl;
        if (chkd[k]) continue;
        buf_chain = [];
        const seen = [];
        for (j = k; j >= 0; ) {
          seen[j] = true;
          chkd[j] = true;
          buf[buf.length] = j;
          buf_chain.push(sectors[j]);
          const addr = fat_addrs[Math.floor((j * 4) / ssz)];
          jj = (j * 4) & modulus;
          if (ssz < 4 + jj)
            throw new Error(`FAT boundary crossed: ${j} 4 ${ssz}`);
          if (!sectors[addr]) break;
          j = __readInt32LE(sectors[addr], jj);
          if (seen[j]) break;
        }
        sector_list[k] = { nodes: buf, data: __toBuffer([buf_chain]) };
      }
      return sector_list;
    }

    /* [MS-CFB] 2.6.1 Compound File Directory Entry */
    function read_directory(
      dir_start,
      sector_list,
      sectors,
      Paths,
      nmfs,
      files,
      FileIndex,
      mini
    ) {
      let minifat_store = 0;
      const pl = Paths.length ? 2 : 0;
      const sector = sector_list[dir_start].data;
      let i = 0;
      let namelen = 0;
      let name;
      for (; i < sector.length; i += 128) {
        const blob = sector.slice(i, i + 128);
        prep_blob(blob, 64);
        namelen = blob.read_shift(2);
        name = __utf16le(blob, 0, namelen - pl);
        Paths.push(name);
        const o = {
          name,
          type: blob.read_shift(1),
          color: blob.read_shift(1),
          L: blob.read_shift(4, "i"),
          R: blob.read_shift(4, "i"),
          C: blob.read_shift(4, "i"),
          clsid: blob.read_shift(16),
          state: blob.read_shift(4, "i"),
          start: 0,
          size: 0
        };
        const ctime =
          blob.read_shift(2) +
          blob.read_shift(2) +
          blob.read_shift(2) +
          blob.read_shift(2);
        if (ctime !== 0) o.ct = read_date(blob, blob.l - 8);
        const mtime =
          blob.read_shift(2) +
          blob.read_shift(2) +
          blob.read_shift(2) +
          blob.read_shift(2);
        if (mtime !== 0) o.mt = read_date(blob, blob.l - 8);
        o.start = blob.read_shift(4, "i");
        o.size = blob.read_shift(4, "i");
        if (o.size < 0 && o.start < 0) {
          o.size = o.type = 0;
          o.start = ENDOFCHAIN;
          o.name = "";
        }
        if (o.type === 5) {
          /* root */
          minifat_store = o.start;
          if (nmfs > 0 && minifat_store !== ENDOFCHAIN)
            sector_list[minifat_store].name = "!StreamData";
          /* minifat_size = o.size; */
        } else if (o.size >= 4096 /* MSCSZ */) {
          o.storage = "fat";
          if (sector_list[o.start] === undefined)
            sector_list[o.start] = get_sector_list(
              sectors,
              o.start,
              sector_list.fat_addrs,
              sector_list.ssz
            );
          sector_list[o.start].name = o.name;
          o.content = sector_list[o.start].data.slice(0, o.size);
        } else {
          o.storage = "minifat";
          if (o.size < 0) o.size = 0;
          else if (
            minifat_store !== ENDOFCHAIN &&
            o.start !== ENDOFCHAIN &&
            sector_list[minifat_store]
          ) {
            o.content = get_mfat_entry(
              o,
              sector_list[minifat_store].data,
              (sector_list[mini] || {}).data
            );
          }
        }
        if (o.content) prep_blob(o.content, 0);
        files[name] = o;
        FileIndex.push(o);
      }
    }

    function read_date(blob, offset) {
      return new Date(
        ((__readUInt32LE(blob, offset + 4) / 1e7) * Math.pow(2, 32) +
          __readUInt32LE(blob, offset) / 1e7 -
          11644473600) *
          1000
      );
    }

    function read_file(filename, options) {
      get_fs();
      return parse(fs.readFileSync(filename), options);
    }

    function read(blob, options) {
      switch ((options && options.type) || "base64") {
        case "file":
          return read_file(blob, options);
        case "base64":
          return parse(s2a(Base64.decode(blob)), options);
        case "binary":
          return parse(s2a(blob), options);
      }
      return parse(blob, options);
    }

    function init_cfb(cfb, opts) {
      const o = opts || {};
      const root = o.root || "Root Entry";
      if (!cfb.FullPaths) cfb.FullPaths = [];
      if (!cfb.FileIndex) cfb.FileIndex = [];
      if (cfb.FullPaths.length !== cfb.FileIndex.length)
        throw new Error("inconsistent CFB structure");
      if (cfb.FullPaths.length === 0) {
        cfb.FullPaths[0] = `${root}/`;
        cfb.FileIndex[0] = { name: root, type: 5 };
      }
      if (o.CLSID) cfb.FileIndex[0].clsid = o.CLSID;
      seed_cfb(cfb);
    }
    function seed_cfb(cfb) {
      const nm = "\u0001Sh33tJ5";
      if (CFB.find(cfb, `/${nm}`)) return;
      const p = new_buf(4);
      p[0] = 55;
      p[1] = p[3] = 50;
      p[2] = 54;
      cfb.FileIndex.push({
        name: nm,
        type: 2,
        content: p,
        size: 4,
        L: 69,
        R: 69,
        C: 69
      });
      cfb.FullPaths.push(cfb.FullPaths[0] + nm);
      rebuild_cfb(cfb);
    }
    function rebuild_cfb(cfb, f) {
      init_cfb(cfb);
      let gc = false;
      let s = false;
      for (var i = cfb.FullPaths.length - 1; i >= 0; --i) {
        const _file = cfb.FileIndex[i];
        switch (_file.type) {
          case 0:
            if (s) gc = true;
            else {
              cfb.FileIndex.pop();
              cfb.FullPaths.pop();
            }
            break;
          case 1:
          case 2:
          case 5:
            s = true;
            if (isNaN(_file.R * _file.L * _file.C)) gc = true;
            if (_file.R > -1 && _file.L > -1 && _file.R == _file.L) gc = true;
            break;
          default:
            gc = true;
            break;
        }
      }
      if (!gc && !f) return;

      const now = new Date(1987, 1, 19);
      let j = 0;
      const data = [];
      for (i = 0; i < cfb.FullPaths.length; ++i) {
        if (cfb.FileIndex[i].type === 0) continue;
        data.push([cfb.FullPaths[i], cfb.FileIndex[i]]);
      }
      for (i = 0; i < data.length; ++i) {
        const dad = dirname(data[i][0]);
        s = false;
        for (j = 0; j < data.length; ++j) if (data[j][0] === dad) s = true;
        if (!s)
          data.push([
            dad,
            {
              name: filename(dad).replace("/", ""),
              type: 1,
              clsid: HEADER_CLSID,
              ct: now,
              mt: now,
              content: null
            }
          ]);
      }

      data.sort(function(x, y) {
        return namecmp(x[0], y[0]);
      });
      cfb.FullPaths = [];
      cfb.FileIndex = [];
      for (i = 0; i < data.length; ++i) {
        cfb.FullPaths[i] = data[i][0];
        cfb.FileIndex[i] = data[i][1];
      }
      for (i = 0; i < data.length; ++i) {
        const elt = cfb.FileIndex[i];
        const nm = cfb.FullPaths[i];

        elt.name = filename(nm).replace("/", "");
        elt.L = elt.R = elt.C = -(elt.color = 1);
        elt.size = elt.content ? elt.content.length : 0;
        elt.start = 0;
        elt.clsid = elt.clsid || HEADER_CLSID;
        if (i === 0) {
          elt.C = data.length > 1 ? 1 : -1;
          elt.size = 0;
          elt.type = 5;
        } else if (nm.slice(-1) == "/") {
          for (j = i + 1; j < data.length; ++j)
            if (dirname(cfb.FullPaths[j]) == nm) break;
          elt.C = j >= data.length ? -1 : j;
          for (j = i + 1; j < data.length; ++j)
            if (dirname(cfb.FullPaths[j]) == dirname(nm)) break;
          elt.R = j >= data.length ? -1 : j;
          elt.type = 1;
        } else {
          if (dirname(cfb.FullPaths[i + 1] || "") == dirname(nm)) elt.R = i + 1;
          elt.type = 2;
        }
      }
    }

    function _write(cfb, options) {
      const _opts = options || {};
      rebuild_cfb(cfb);
      if (_opts.fileType == "zip") return write_zip(cfb, _opts);
      const L = (function(cfb) {
        let mini_size = 0;
        let fat_size = 0;
        for (let i = 0; i < cfb.FileIndex.length; ++i) {
          const file = cfb.FileIndex[i];
          if (!file.content) continue;
          const flen = file.content.length;
          if (flen > 0) {
            if (flen < 0x1000) mini_size += (flen + 0x3f) >> 6;
            else fat_size += (flen + 0x01ff) >> 9;
          }
        }
        const dir_cnt = (cfb.FullPaths.length + 3) >> 2;
        const mini_cnt = (mini_size + 7) >> 3;
        const mfat_cnt = (mini_size + 0x7f) >> 7;
        const fat_base = mini_cnt + fat_size + dir_cnt + mfat_cnt;
        let fat_cnt = (fat_base + 0x7f) >> 7;
        let difat_cnt = fat_cnt <= 109 ? 0 : Math.ceil((fat_cnt - 109) / 0x7f);
        while ((fat_base + fat_cnt + difat_cnt + 0x7f) >> 7 > fat_cnt)
          difat_cnt = ++fat_cnt <= 109 ? 0 : Math.ceil((fat_cnt - 109) / 0x7f);
        const L = [
          1,
          difat_cnt,
          fat_cnt,
          mfat_cnt,
          dir_cnt,
          fat_size,
          mini_size,
          0
        ];
        cfb.FileIndex[0].size = mini_size << 6;
        L[7] =
          (cfb.FileIndex[0].start = L[0] + L[1] + L[2] + L[3] + L[4] + L[5]) +
          ((L[6] + 7) >> 3);
        return L;
      })(cfb);
      const o = new_buf(L[7] << 9);
      let i = 0;
      let T = 0;
      {
        for (i = 0; i < 8; ++i) o.write_shift(1, HEADER_SIG[i]);
        for (i = 0; i < 8; ++i) o.write_shift(2, 0);
        o.write_shift(2, 0x003e);
        o.write_shift(2, 0x0003);
        o.write_shift(2, 0xfffe);
        o.write_shift(2, 0x0009);
        o.write_shift(2, 0x0006);
        for (i = 0; i < 3; ++i) o.write_shift(2, 0);
        o.write_shift(4, 0);
        o.write_shift(4, L[2]);
        o.write_shift(4, L[0] + L[1] + L[2] + L[3] - 1);
        o.write_shift(4, 0);
        o.write_shift(4, 1 << 12);
        o.write_shift(4, L[3] ? L[0] + L[1] + L[2] - 1 : ENDOFCHAIN);
        o.write_shift(4, L[3]);
        o.write_shift(-4, L[1] ? L[0] - 1 : ENDOFCHAIN);
        o.write_shift(4, L[1]);
        for (i = 0; i < 109; ++i) o.write_shift(-4, i < L[2] ? L[1] + i : -1);
      }
      if (L[1]) {
        for (T = 0; T < L[1]; ++T) {
          for (; i < 236 + T * 127; ++i)
            o.write_shift(-4, i < L[2] ? L[1] + i : -1);
          o.write_shift(-4, T === L[1] - 1 ? ENDOFCHAIN : T + 1);
        }
      }
      const chainit = function(w) {
        for (T += w; i < T - 1; ++i) o.write_shift(-4, i + 1);
        if (w) {
          ++i;
          o.write_shift(-4, ENDOFCHAIN);
        }
      };
      T = i = 0;
      for (T += L[1]; i < T; ++i) o.write_shift(-4, consts.DIFSECT);
      for (T += L[2]; i < T; ++i) o.write_shift(-4, consts.FATSECT);
      chainit(L[3]);
      chainit(L[4]);
      let j = 0;
      let flen = 0;
      let file = cfb.FileIndex[0];
      for (; j < cfb.FileIndex.length; ++j) {
        file = cfb.FileIndex[j];
        if (!file.content) continue;
        flen = file.content.length;
        if (flen < 0x1000) continue;
        file.start = T;
        chainit((flen + 0x01ff) >> 9);
      }
      chainit((L[6] + 7) >> 3);
      while (o.l & 0x1ff) o.write_shift(-4, consts.ENDOFCHAIN);
      T = i = 0;
      for (j = 0; j < cfb.FileIndex.length; ++j) {
        file = cfb.FileIndex[j];
        if (!file.content) continue;
        flen = file.content.length;
        if (!flen || flen >= 0x1000) continue;
        file.start = T;
        chainit((flen + 0x3f) >> 6);
      }
      while (o.l & 0x1ff) o.write_shift(-4, consts.ENDOFCHAIN);
      for (i = 0; i < L[4] << 2; ++i) {
        const nm = cfb.FullPaths[i];
        if (!nm || nm.length === 0) {
          for (j = 0; j < 17; ++j) o.write_shift(4, 0);
          for (j = 0; j < 3; ++j) o.write_shift(4, -1);
          for (j = 0; j < 12; ++j) o.write_shift(4, 0);
          continue;
        }
        file = cfb.FileIndex[i];
        if (i === 0) file.start = file.size ? file.start - 1 : ENDOFCHAIN;
        const _nm = (i === 0 && _opts.root) || file.name;
        flen = 2 * (_nm.length + 1);
        o.write_shift(64, _nm, "utf16le");
        o.write_shift(2, flen);
        o.write_shift(1, file.type);
        o.write_shift(1, file.color);
        o.write_shift(-4, file.L);
        o.write_shift(-4, file.R);
        o.write_shift(-4, file.C);
        if (!file.clsid) for (j = 0; j < 4; ++j) o.write_shift(4, 0);
        else o.write_shift(16, file.clsid, "hex");
        o.write_shift(4, file.state || 0);
        o.write_shift(4, 0);
        o.write_shift(4, 0);
        o.write_shift(4, 0);
        o.write_shift(4, 0);
        o.write_shift(4, file.start);
        o.write_shift(4, file.size);
        o.write_shift(4, 0);
      }
      for (i = 1; i < cfb.FileIndex.length; ++i) {
        file = cfb.FileIndex[i];
        if (file.size >= 0x1000) {
          o.l = (file.start + 1) << 9;
          for (j = 0; j < file.size; ++j) o.write_shift(1, file.content[j]);
          for (; j & 0x1ff; ++j) o.write_shift(1, 0);
        }
      }
      for (i = 1; i < cfb.FileIndex.length; ++i) {
        file = cfb.FileIndex[i];
        if (file.size > 0 && file.size < 0x1000) {
          for (j = 0; j < file.size; ++j) o.write_shift(1, file.content[j]);
          for (; j & 0x3f; ++j) o.write_shift(1, 0);
        }
      }
      while (o.l < o.length) o.write_shift(1, 0);
      return o;
    }
    /* [MS-CFB] 2.6.4 (Unicode 3.0.1 case conversion) */
    function find(cfb, path) {
      const UCFullPaths = cfb.FullPaths.map(function(x) {
        return x.toUpperCase();
      });
      const UCPaths = UCFullPaths.map(function(x) {
        const y = x.split("/");
        return y[y.length - (x.slice(-1) == "/" ? 2 : 1)];
      });
      let k = false;
      if (path.charCodeAt(0) === 47 /* "/" */) {
        k = true;
        path = UCFullPaths[0].slice(0, -1) + path;
      } else k = path.indexOf("/") !== -1;
      let UCPath = path.toUpperCase();
      let w =
        k === true ? UCFullPaths.indexOf(UCPath) : UCPaths.indexOf(UCPath);
      if (w !== -1) return cfb.FileIndex[w];

      const m = !UCPath.match(chr1);
      UCPath = UCPath.replace(chr0, "");
      if (m) UCPath = UCPath.replace(chr1, "!");
      for (w = 0; w < UCFullPaths.length; ++w) {
        if (
          (m ? UCFullPaths[w].replace(chr1, "!") : UCFullPaths[w]).replace(
            chr0,
            ""
          ) == UCPath
        )
          return cfb.FileIndex[w];
        if (
          (m ? UCPaths[w].replace(chr1, "!") : UCPaths[w]).replace(chr0, "") ==
          UCPath
        )
          return cfb.FileIndex[w];
      }
      return null;
    }
    /** CFB Constants */
    var MSSZ = 64; /* Mini Sector Size = 1<<6 */
    // var MSCSZ = 4096; /* Mini Stream Cutoff Size */
    /* 2.1 Compound File Sector Numbers and Types */
    var ENDOFCHAIN = -2;
    /* 2.2 Compound File Header */
    var HEADER_SIGNATURE = "d0cf11e0a1b11ae1";
    var HEADER_SIG = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
    var HEADER_CLSID = "00000000000000000000000000000000";
    var consts = {
      /* 2.1 Compund File Sector Numbers and Types */
      MAXREGSECT: -6,
      DIFSECT: -4,
      FATSECT: -3,
      ENDOFCHAIN,
      FREESECT: -1,
      /* 2.2 Compound File Header */
      HEADER_SIGNATURE,
      HEADER_MINOR_VERSION: "3e00",
      MAXREGSID: -6,
      NOSTREAM: -1,
      HEADER_CLSID,
      /* 2.6.1 Compound File Directory Entry */
      EntryTypes: [
        "unknown",
        "storage",
        "stream",
        "lockbytes",
        "property",
        "root"
      ]
    };

    function write_file(cfb, filename, options) {
      get_fs();
      const o = _write(cfb, options);
      fs.writeFileSync(filename, o);
    }

    function a2s(o) {
      const out = new Array(o.length);
      for (let i = 0; i < o.length; ++i) out[i] = String.fromCharCode(o[i]);
      return out.join("");
    }

    function write(cfb, options) {
      const o = _write(cfb, options);
      switch (options && options.type) {
        case "file":
          get_fs();
          fs.writeFileSync(options.filename, o);
          return o;
        case "binary":
          return a2s(o);
        case "base64":
          return Base64.encode(a2s(o));
      }
      return o;
    }
    /* node < 8.1 zlib does not expose bytesRead, so default to pure JS */
    let _zlib;
    function use_zlib(zlib) {
      try {
        const { InflateRaw } = zlib;
        const InflRaw = new InflateRaw();
        InflRaw._processChunk(new Uint8Array([3, 0]), InflRaw._finishFlushFlag);
        if (InflRaw.bytesRead) _zlib = zlib;
        else throw new Error("zlib does not expose bytesRead");
      } catch (e) {
        console.error(`cannot use native zlib: ${e.message || e}`);
      }
    }

    function _inflateRawSync(payload, usz) {
      if (!_zlib) return _inflate(payload, usz);
      const { InflateRaw } = _zlib;
      const InflRaw = new InflateRaw();
      const out = InflRaw._processChunk(
        payload.slice(payload.l),
        InflRaw._finishFlushFlag
      );
      payload.l += InflRaw.bytesRead;
      return out;
    }

    function _deflateRawSync(payload) {
      return _zlib ? _zlib.deflateRawSync(payload) : _deflate(payload);
    }
    const CLEN_ORDER = [
      16,
      17,
      18,
      0,
      8,
      7,
      9,
      6,
      10,
      5,
      11,
      4,
      12,
      3,
      13,
      2,
      14,
      1,
      15
    ];

    /*  LEN_ID = [ 257, 258, 259, 260, 261, 262, 263, 264, 265, 266, 267, 268, 269, 270, 271, 272, 273, 274, 275, 276, 277, 278, 279, 280, 281, 282, 283, 284, 285 ]; */
    const LEN_LN = [
      3,
      4,
      5,
      6,
      7,
      8,
      9,
      10,
      11,
      13,
      15,
      17,
      19,
      23,
      27,
      31,
      35,
      43,
      51,
      59,
      67,
      83,
      99,
      115,
      131,
      163,
      195,
      227,
      258
    ];

    /*  DST_ID = [  0,  1,  2,  3,  4,  5,  6,  7,  8,  9, 10, 11, 12, 13,  14,  15,  16,  17,  18,  19,   20,   21,   22,   23,   24,   25,   26,    27,    28,    29 ]; */
    const DST_LN = [
      1,
      2,
      3,
      4,
      5,
      7,
      9,
      13,
      17,
      25,
      33,
      49,
      65,
      97,
      129,
      193,
      257,
      385,
      513,
      769,
      1025,
      1537,
      2049,
      3073,
      4097,
      6145,
      8193,
      12289,
      16385,
      24577
    ];

    function bit_swap_8(n) {
      const t =
        (((n << 1) | (n << 11)) & 0x22110) | (((n << 5) | (n << 15)) & 0x88440);
      return ((t >> 16) | (t >> 8) | t) & 0xff;
    }

    const use_typed_arrays = typeof Uint8Array !== "undefined";

    const bitswap8 = use_typed_arrays ? new Uint8Array(1 << 8) : [];
    for (let q = 0; q < 1 << 8; ++q) bitswap8[q] = bit_swap_8(q);

    function bit_swap_n(n, b) {
      let rev = bitswap8[n & 0xff];
      if (b <= 8) return rev >>> (8 - b);
      rev = (rev << 8) | bitswap8[(n >> 8) & 0xff];
      if (b <= 16) return rev >>> (16 - b);
      rev = (rev << 8) | bitswap8[(n >> 16) & 0xff];
      return rev >>> (24 - b);
    }

    /* helpers for unaligned bit reads */
    function read_bits_2(buf, bl) {
      const w = bl & 7;
      const h = bl >>> 3;
      return ((buf[h] | (w <= 6 ? 0 : buf[h + 1] << 8)) >>> w) & 0x03;
    }
    function read_bits_3(buf, bl) {
      const w = bl & 7;
      const h = bl >>> 3;
      return ((buf[h] | (w <= 5 ? 0 : buf[h + 1] << 8)) >>> w) & 0x07;
    }
    function read_bits_4(buf, bl) {
      const w = bl & 7;
      const h = bl >>> 3;
      return ((buf[h] | (w <= 4 ? 0 : buf[h + 1] << 8)) >>> w) & 0x0f;
    }
    function read_bits_5(buf, bl) {
      const w = bl & 7;
      const h = bl >>> 3;
      return ((buf[h] | (w <= 3 ? 0 : buf[h + 1] << 8)) >>> w) & 0x1f;
    }
    function read_bits_7(buf, bl) {
      const w = bl & 7;
      const h = bl >>> 3;
      return ((buf[h] | (w <= 1 ? 0 : buf[h + 1] << 8)) >>> w) & 0x7f;
    }

    /* works up to n = 3 * 8 + 1 = 25 */
    function read_bits_n(buf, bl, n) {
      const w = bl & 7;
      const h = bl >>> 3;
      const f = (1 << n) - 1;
      let v = buf[h] >>> w;
      if (n < 8 - w) return v & f;
      v |= buf[h + 1] << (8 - w);
      if (n < 16 - w) return v & f;
      v |= buf[h + 2] << (16 - w);
      if (n < 24 - w) return v & f;
      v |= buf[h + 3] << (24 - w);
      return v & f;
    }

    /* until ArrayBuffer#realloc is a thing, fake a realloc */
    function realloc(b, sz) {
      const L = b.length;
      const M = 2 * L > sz ? 2 * L : sz + 5;
      let i = 0;
      if (L >= sz) return b;
      if (has_buf) {
        const o = new_unsafe_buf(M);
        // $FlowIgnore
        if (b.copy) b.copy(o);
        else for (; i < b.length; ++i) o[i] = b[i];
        return o;
      } else if (use_typed_arrays) {
        const a = new Uint8Array(M);
        if (a.set) a.set(b);
        else for (; i < b.length; ++i) a[i] = b[i];
        return a;
      }
      b.length = M;
      return b;
    }

    /* zero-filled arrays for older browsers */
    function zero_fill_array(n) {
      const o = new Array(n);
      for (let i = 0; i < n; ++i) o[i] = 0;
      return o;
    }
    var _deflate = (function() {
      const _deflateRaw = (function() {
        return function deflateRaw(data, out) {
          let boff = 0;
          while (boff < data.length) {
            let L = Math.min(0xffff, data.length - boff);
            const h = boff + L == data.length;
            /* TODO: this is only type 0 stored */
            out.write_shift(1, +h);
            out.write_shift(2, L);
            out.write_shift(2, ~L & 0xffff);
            while (L-- > 0) out[out.l++] = data[boff++];
          }
          return out.l;
        };
      })();

      return function(data) {
        const buf = new_buf(50 + Math.floor(data.length * 1.1));
        const off = _deflateRaw(data, buf);
        return buf.slice(0, off);
      };
    })();
    /* modified inflate function also moves original read head */

    /* build tree (used for literals and lengths) */
    function build_tree(clens, cmap, MAX) {
      let maxlen = 1;
      let w = 0;
      let i = 0;
      let j = 0;
      let ccode = 0;
      let L = clens.length;

      const bl_count = use_typed_arrays
        ? new Uint16Array(32)
        : zero_fill_array(32);
      for (i = 0; i < 32; ++i) bl_count[i] = 0;

      for (i = L; i < MAX; ++i) clens[i] = 0;
      L = clens.length;

      const ctree = use_typed_arrays ? new Uint16Array(L) : zero_fill_array(L); // []

      /* build code tree */
      for (i = 0; i < L; ++i) {
        bl_count[(w = clens[i])]++;
        if (maxlen < w) maxlen = w;
        ctree[i] = 0;
      }
      bl_count[0] = 0;
      for (i = 1; i <= maxlen; ++i)
        bl_count[i + 16] = ccode = (ccode + bl_count[i - 1]) << 1;
      for (i = 0; i < L; ++i) {
        ccode = clens[i];
        if (ccode != 0) ctree[i] = bl_count[ccode + 16]++;
      }

      /* cmap[maxlen + 4 bits] = (off&15) + (lit<<4) reverse mapping */
      let cleni = 0;
      for (i = 0; i < L; ++i) {
        cleni = clens[i];
        if (cleni != 0) {
          ccode = bit_swap_n(ctree[i], maxlen) >> (maxlen - cleni);
          for (j = (1 << (maxlen + 4 - cleni)) - 1; j >= 0; --j)
            cmap[ccode | (j << cleni)] = (cleni & 15) | (i << 4);
        }
      }
      return maxlen;
    }

    const fix_lmap = use_typed_arrays
      ? new Uint16Array(512)
      : zero_fill_array(512);
    const fix_dmap = use_typed_arrays
      ? new Uint16Array(32)
      : zero_fill_array(32);
    if (!use_typed_arrays) {
      for (var i = 0; i < 512; ++i) fix_lmap[i] = 0;
      for (i = 0; i < 32; ++i) fix_dmap[i] = 0;
    }
    (function() {
      const dlens = [];
      let i = 0;
      for (; i < 32; i++) dlens.push(5);
      build_tree(dlens, fix_dmap, 32);

      const clens = [];
      i = 0;
      for (; i <= 143; i++) clens.push(8);
      for (; i <= 255; i++) clens.push(9);
      for (; i <= 279; i++) clens.push(7);
      for (; i <= 287; i++) clens.push(8);
      build_tree(clens, fix_lmap, 288);
    })();

    const dyn_lmap = use_typed_arrays
      ? new Uint16Array(32768)
      : zero_fill_array(32768);
    const dyn_dmap = use_typed_arrays
      ? new Uint16Array(32768)
      : zero_fill_array(32768);
    const dyn_cmap = use_typed_arrays
      ? new Uint16Array(128)
      : zero_fill_array(128);
    let dyn_len_1 = 1;
    let dyn_len_2 = 1;

    /* 5.5.3 Expanding Huffman Codes */
    function dyn(data, boff) {
      /* nomenclature from RFC1951 refers to bit values; these are offset by the implicit constant */
      const _HLIT = read_bits_5(data, boff) + 257;
      boff += 5;
      const _HDIST = read_bits_5(data, boff) + 1;
      boff += 5;
      const _HCLEN = read_bits_4(data, boff) + 4;
      boff += 4;
      let w = 0;

      /* grab and store code lengths */
      const clens = use_typed_arrays ? new Uint8Array(19) : zero_fill_array(19);
      const ctree = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
      let maxlen = 1;
      const bl_count = use_typed_arrays
        ? new Uint8Array(8)
        : zero_fill_array(8);
      const next_code = use_typed_arrays
        ? new Uint8Array(8)
        : zero_fill_array(8);
      const L = clens.length; /* 19 */
      for (var i = 0; i < _HCLEN; ++i) {
        clens[CLEN_ORDER[i]] = w = read_bits_3(data, boff);
        if (maxlen < w) maxlen = w;
        bl_count[w]++;
        boff += 3;
      }

      /* build code tree */
      let ccode = 0;
      bl_count[0] = 0;
      for (i = 1; i <= maxlen; ++i)
        next_code[i] = ccode = (ccode + bl_count[i - 1]) << 1;
      for (i = 0; i < L; ++i)
        if ((ccode = clens[i]) != 0) ctree[i] = next_code[ccode]++;
      /* cmap[7 bits from stream] = (off&7) + (lit<<3) */
      let cleni = 0;
      for (i = 0; i < L; ++i) {
        cleni = clens[i];
        if (cleni != 0) {
          ccode = bitswap8[ctree[i]] >> (8 - cleni);
          for (let j = (1 << (7 - cleni)) - 1; j >= 0; --j)
            dyn_cmap[ccode | (j << cleni)] = (cleni & 7) | (i << 3);
        }
      }

      /* read literal and dist codes at once */
      const hcodes = [];
      maxlen = 1;
      for (; hcodes.length < _HLIT + _HDIST; ) {
        ccode = dyn_cmap[read_bits_7(data, boff)];
        boff += ccode & 7;
        switch ((ccode >>>= 3)) {
          case 16:
            w = 3 + read_bits_2(data, boff);
            boff += 2;
            ccode = hcodes[hcodes.length - 1];
            while (w-- > 0) hcodes.push(ccode);
            break;
          case 17:
            w = 3 + read_bits_3(data, boff);
            boff += 3;
            while (w-- > 0) hcodes.push(0);
            break;
          case 18:
            w = 11 + read_bits_7(data, boff);
            boff += 7;
            while (w-- > 0) hcodes.push(0);
            break;
          default:
            hcodes.push(ccode);
            if (maxlen < ccode) maxlen = ccode;
            break;
        }
      }

      /* build literal / length trees */
      const h1 = hcodes.slice(0, _HLIT);
      const h2 = hcodes.slice(_HLIT);
      for (i = _HLIT; i < 286; ++i) h1[i] = 0;
      for (i = _HDIST; i < 30; ++i) h2[i] = 0;
      dyn_len_1 = build_tree(h1, dyn_lmap, 286);
      dyn_len_2 = build_tree(h2, dyn_dmap, 30);
      return boff;
    }

    /* return [ data, bytesRead ] */
    function inflate(data, usz) {
      /* shortcircuit for empty buffer [0x03, 0x00] */
      if (data[0] == 3 && !(data[1] & 0x3)) {
        return [new_raw_buf(usz), 2];
      }

      /* bit offset */
      let boff = 0;

      /* header includes final bit and type bits */
      let header = 0;

      let outbuf = new_unsafe_buf(usz || 1 << 18);
      let woff = 0;
      let OL = outbuf.length >>> 0;
      let max_len_1 = 0;
      let max_len_2 = 0;

      while ((header & 1) == 0) {
        header = read_bits_3(data, boff);
        boff += 3;
        if (header >>> 1 == 0) {
          /* Stored block */
          if (boff & 7) boff += 8 - (boff & 7);
          /* 2 bytes sz, 2 bytes bit inverse */
          let sz = data[boff >>> 3] | (data[(boff >>> 3) + 1] << 8);
          boff += 32;
          /* push sz bytes */
          if (!usz && OL < woff + sz) {
            outbuf = realloc(outbuf, woff + sz);
            OL = outbuf.length;
          }
          if (typeof data.copy === "function") {
            // $FlowIgnore
            data.copy(outbuf, woff, boff >>> 3, (boff >>> 3) + sz);
            woff += sz;
            boff += 8 * sz;
          } else
            while (sz-- > 0) {
              outbuf[woff++] = data[boff >>> 3];
              boff += 8;
            }
          continue;
        } else if (header >>> 1 == 1) {
          /* Fixed Huffman */
          max_len_1 = 9;
          max_len_2 = 5;
        } else {
          /* Dynamic Huffman */
          boff = dyn(data, boff);
          max_len_1 = dyn_len_1;
          max_len_2 = dyn_len_2;
        }
        if (!usz && OL < woff + 32767) {
          outbuf = realloc(outbuf, woff + 32767);
          OL = outbuf.length;
        }
        for (;;) {
          // while(true) is apparently out of vogue in modern JS circles
          /* ingest code and move read head */
          let bits = read_bits_n(data, boff, max_len_1);
          let code = header >>> 1 == 1 ? fix_lmap[bits] : dyn_lmap[bits];
          boff += code & 15;
          code >>>= 4;
          /* 0-255 are literals, 256 is end of block token, 257+ are copy tokens */
          if (((code >>> 8) & 0xff) === 0) outbuf[woff++] = code;
          else if (code == 256) break;
          else {
            code -= 257;
            let len_eb = code < 8 ? 0 : (code - 4) >> 2;
            if (len_eb > 5) len_eb = 0;
            let tgt = woff + LEN_LN[code];
            /* length extra bits */
            if (len_eb > 0) {
              tgt += read_bits_n(data, boff, len_eb);
              boff += len_eb;
            }

            /* dist code */
            bits = read_bits_n(data, boff, max_len_2);
            code = header >>> 1 == 1 ? fix_dmap[bits] : dyn_dmap[bits];
            boff += code & 15;
            code >>>= 4;
            const dst_eb = code < 4 ? 0 : (code - 2) >> 1;
            let dst = DST_LN[code];
            /* dist extra bits */
            if (dst_eb > 0) {
              dst += read_bits_n(data, boff, dst_eb);
              boff += dst_eb;
            }

            /* in the common case, manual byte copy is faster than TA set / Buffer copy */
            if (!usz && OL < tgt) {
              outbuf = realloc(outbuf, tgt);
              OL = outbuf.length;
            }
            while (woff < tgt) {
              outbuf[woff] = outbuf[woff - dst];
              ++woff;
            }
          }
        }
      }
      return [usz ? outbuf : outbuf.slice(0, woff), (boff + 7) >>> 3];
    }

    function _inflate(payload, usz) {
      const data = payload.slice(payload.l || 0);
      const out = inflate(data, usz);
      payload.l += out[1];
      return out[0];
    }

    function warn_or_throw(wrn, msg) {
      if (wrn) {
        if (typeof console !== "undefined") console.error(msg);
      } else throw new Error(msg);
    }

    function parse_zip(file, options) {
      const blob = file;
      prep_blob(blob, 0);

      const FileIndex = [];
      const FullPaths = [];
      const o = {
        FileIndex,
        FullPaths
      };
      init_cfb(o, { root: options.root });

      /* find end of central directory, start just after signature */
      let i = blob.length - 4;
      while (
        (blob[i] != 0x50 ||
          blob[i + 1] != 0x4b ||
          blob[i + 2] != 0x05 ||
          blob[i + 3] != 0x06) &&
        i >= 0
      )
        --i;
      blob.l = i + 4;

      /* parse end of central directory */
      blob.l += 4;
      const fcnt = blob.read_shift(2);
      blob.l += 6;
      const start_cd = blob.read_shift(4);

      /* parse central directory */
      blob.l = start_cd;

      for (i = 0; i < fcnt; ++i) {
        /* trust local file header instead of CD entry */
        blob.l += 20;
        const csz = blob.read_shift(4);
        const usz = blob.read_shift(4);
        const namelen = blob.read_shift(2);
        const efsz = blob.read_shift(2);
        const fcsz = blob.read_shift(2);
        blob.l += 8;
        const offset = blob.read_shift(4);
        const EF = parse_extra_field(
          blob.slice(blob.l + namelen, blob.l + namelen + efsz)
        );
        blob.l += namelen + efsz + fcsz;

        const L = blob.l;
        blob.l = offset + 4;
        parse_local_file(blob, csz, usz, o, EF);
        blob.l = L;
      }

      return o;
    }

    /* head starts just after local file header signature */
    function parse_local_file(blob, csz, usz, o, EF) {
      /* [local file header] */
      blob.l += 2;
      const flags = blob.read_shift(2);
      const meth = blob.read_shift(2);
      let date = parse_dos_date(blob);

      if (flags & 0x2041) throw new Error("Unsupported ZIP encryption");
      let crc32 = blob.read_shift(4);
      let _csz = blob.read_shift(4);
      let _usz = blob.read_shift(4);

      const namelen = blob.read_shift(2);
      const efsz = blob.read_shift(2);

      // TODO: flags & (1<<11) // UTF8
      let name = "";
      for (let i = 0; i < namelen; ++i)
        name += String.fromCharCode(blob[blob.l++]);
      if (efsz) {
        const ef = parse_extra_field(blob.slice(blob.l, blob.l + efsz));
        if ((ef[0x5455] || {}).mt) date = ef[0x5455].mt;
        if (((EF || {})[0x5455] || {}).mt) date = EF[0x5455].mt;
      }
      blob.l += efsz;

      /* [encryption header] */

      /* [file data] */
      let data = blob.slice(blob.l, blob.l + _csz);
      switch (meth) {
        case 8:
          data = _inflateRawSync(blob, _usz);
          break;
        case 0:
          break;
        default:
          throw new Error(`Unsupported ZIP Compression method ${meth}`);
      }

      /* [data descriptor] */
      let wrn = false;
      if (flags & 8) {
        crc32 = blob.read_shift(4);
        if (crc32 == 0x08074b50) {
          crc32 = blob.read_shift(4);
          wrn = true;
        }
        _csz = blob.read_shift(4);
        _usz = blob.read_shift(4);
      }

      if (_csz != csz)
        warn_or_throw(wrn, `Bad compressed size: ${csz} != ${_csz}`);
      if (_usz != usz)
        warn_or_throw(wrn, `Bad uncompressed size: ${usz} != ${_usz}`);
      const _crc32 = CRC32.buf(data, 0);
      if (crc32 >> 0 != _crc32 >> 0)
        warn_or_throw(wrn, `Bad CRC32 checksum: ${crc32} != ${_crc32}`);
      cfb_add(o, name, data, { unsafe: true, mt: date });
    }
    function write_zip(cfb, options) {
      const _opts = options || {};
      const out = [];
      const cdirs = [];
      let o = new_buf(1);
      const method = _opts.compression ? 8 : 0;
      let flags = 0;
      const desc = false;
      if (desc) flags |= 8;
      let i = 0;
      let j = 0;

      let start_cd = 0;
      let fcnt = 0;
      const root = cfb.FullPaths[0];
      let fp = root;
      let fi = cfb.FileIndex[0];
      const crcs = [];
      let sz_cd = 0;

      for (i = 1; i < cfb.FullPaths.length; ++i) {
        fp = cfb.FullPaths[i].slice(root.length);
        fi = cfb.FileIndex[i];
        if (!fi.size || !fi.content || fp == "\u0001Sh33tJ5") continue;
        const start = start_cd;

        /* TODO: CP437 filename */
        let namebuf = new_buf(fp.length);
        for (j = 0; j < fp.length; ++j)
          namebuf.write_shift(1, fp.charCodeAt(j) & 0x7f);
        namebuf = namebuf.slice(0, namebuf.l);
        crcs[fcnt] = CRC32.buf(fi.content, 0);

        let outbuf = fi.content;
        if (method == 8) outbuf = _deflateRawSync(outbuf);

        /* local file header */
        o = new_buf(30);
        o.write_shift(4, 0x04034b50);
        o.write_shift(2, 20);
        o.write_shift(2, flags);
        o.write_shift(2, method);
        /* TODO: last mod file time/date */
        if (fi.mt) write_dos_date(o, fi.mt);
        else o.write_shift(4, 0);
        o.write_shift(-4, flags & 8 ? 0 : crcs[fcnt]);
        o.write_shift(4, flags & 8 ? 0 : outbuf.length);
        o.write_shift(4, flags & 8 ? 0 : fi.content.length);
        o.write_shift(2, namebuf.length);
        o.write_shift(2, 0);

        start_cd += o.length;
        out.push(o);
        start_cd += namebuf.length;
        out.push(namebuf);

        /* TODO: encryption header ? */
        start_cd += outbuf.length;
        out.push(outbuf);

        /* data descriptor */
        if (flags & 8) {
          o = new_buf(12);
          o.write_shift(-4, crcs[fcnt]);
          o.write_shift(4, outbuf.length);
          o.write_shift(4, fi.content.length);
          start_cd += o.l;
          out.push(o);
        }

        /* central directory */
        o = new_buf(46);
        o.write_shift(4, 0x02014b50);
        o.write_shift(2, 0);
        o.write_shift(2, 20);
        o.write_shift(2, flags);
        o.write_shift(2, method);
        o.write_shift(4, 0); /* TODO: last mod file time/date */
        o.write_shift(-4, crcs[fcnt]);

        o.write_shift(4, outbuf.length);
        o.write_shift(4, fi.content.length);
        o.write_shift(2, namebuf.length);
        o.write_shift(2, 0);
        o.write_shift(2, 0);
        o.write_shift(2, 0);
        o.write_shift(2, 0);
        o.write_shift(4, 0);
        o.write_shift(4, start);

        sz_cd += o.l;
        cdirs.push(o);
        sz_cd += namebuf.length;
        cdirs.push(namebuf);
        ++fcnt;
      }

      /* end of central directory */
      o = new_buf(22);
      o.write_shift(4, 0x06054b50);
      o.write_shift(2, 0);
      o.write_shift(2, 0);
      o.write_shift(2, fcnt);
      o.write_shift(2, fcnt);
      o.write_shift(4, sz_cd);
      o.write_shift(4, start_cd);
      o.write_shift(2, 0);

      return bconcat([bconcat(out), bconcat(cdirs), o]);
    }

    function cfb_add(cfb, name, content, opts) {
      const unsafe = opts && opts.unsafe;
      if (!unsafe) init_cfb(cfb);
      let file = !unsafe && CFB.find(cfb, name);
      if (!file) {
        let fpath = cfb.FullPaths[0];
        if (name.slice(0, fpath.length) == fpath) fpath = name;
        else {
          if (fpath.slice(-1) != "/") fpath += "/";
          fpath = (fpath + name).replace("//", "/");
        }
        file = { name: filename(name), type: 2 };
        cfb.FileIndex.push(file);
        cfb.FullPaths.push(fpath);
        if (!unsafe) CFB.utils.cfb_gc(cfb);
      }
      file.content = content;
      file.size = content ? content.length : 0;
      if (opts) {
        if (opts.CLSID) file.clsid = opts.CLSID;
        if (opts.mt) file.mt = opts.mt;
        if (opts.ct) file.ct = opts.ct;
      }
      return file;
    }

    return exports;
  })();

  if (
    typeof require !== "undefined" &&
    typeof module !== "undefined" &&
    typeof DO_NOT_EXPORT_CFB === "undefined"
  ) {
    module.exports = CFB;
  }
  let _fs;
  if (typeof require !== "undefined")
    try {
      _fs = require("fs");
    } catch (e) {}

  /* read binary data from file */
  function read_binary(path) {
    if (typeof _fs !== "undefined") return _fs.readFileSync(path);
    // $FlowIgnore
    if (
      typeof $ !== "undefined" &&
      typeof File !== "undefined" &&
      typeof Folder !== "undefined"
    )
      try {
        // extendscript
        // $FlowIgnore
        const infile = File(path);
        infile.open("r");
        infile.encoding = "binary";
        const data = infile.read();
        infile.close();
        return data;
      } catch (e) {
        if (!e.message || !e.message.match(/onstruct/)) throw e;
      }
    throw new Error(`Cannot access file ${path}`);
  }
  function keys(o) {
    const ks = Object.keys(o);
    const o2 = [];
    for (let i = 0; i < ks.length; ++i)
      if (Object.prototype.hasOwnProperty.call(o, ks[i])) o2.push(ks[i]);
    return o2;
  }

  function evert_key(obj, key) {
    const o = [];
    const K = keys(obj);
    for (let i = 0; i !== K.length; ++i)
      if (o[obj[K[i]][key]] == null) o[obj[K[i]][key]] = K[i];
    return o;
  }

  function evert(obj) {
    const o = [];
    const K = keys(obj);
    for (let i = 0; i !== K.length; ++i) o[obj[K[i]]] = K[i];
    return o;
  }

  function evert_num(obj) {
    const o = [];
    const K = keys(obj);
    for (let i = 0; i !== K.length; ++i) o[obj[K[i]]] = parseInt(K[i], 10);
    return o;
  }

  const basedate = new Date(1899, 11, 30, 0, 0, 0); // 2209161600000
  function datenum(v, date1904) {
    let epoch = v.getTime();
    if (date1904) epoch -= 1462 * 24 * 60 * 60 * 1000;
    const dnthresh =
      basedate.getTime() +
      (v.getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000;
    return (epoch - dnthresh) / (24 * 60 * 60 * 1000);
  }
  const refdate = new Date();
  const dnthresh =
    basedate.getTime() +
    (refdate.getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000;
  const refoffset = refdate.getTimezoneOffset();
  function numdate(v) {
    const out = new Date();
    out.setTime(v * 24 * 60 * 60 * 1000 + dnthresh);
    if (out.getTimezoneOffset() !== refoffset) {
      out.setTime(
        out.getTime() + (out.getTimezoneOffset() - refoffset) * 60000
      );
    }
    return out;
  }

  /* ISO 8601 Duration */
  function parse_isodur(s) {
    let sec = 0;
    let mt = 0;
    let time = false;
    const m = s.match(
      /P([0-9\.]+Y)?([0-9\.]+M)?([0-9\.]+D)?T([0-9\.]+H)?([0-9\.]+M)?([0-9\.]+S)?/
    );
    if (!m) throw new Error(`|${s}| is not an ISO8601 Duration`);
    for (let i = 1; i != m.length; ++i) {
      if (!m[i]) continue;
      mt = 1;
      if (i > 3) time = true;
      switch (m[i].slice(m[i].length - 1)) {
        case "Y":
          throw new Error(
            `Unsupported ISO Duration Field: ${m[i].slice(m[i].length - 1)}`
          );
        case "D":
          mt *= 24;
        /* falls through */
        case "H":
          mt *= 60;
        /* falls through */
        case "M":
          if (!time) throw new Error("Unsupported ISO Duration Field: M");
          else mt *= 60;
        /* falls through */
        case "S":
          break;
      }
      sec += mt * parseInt(m[i], 10);
    }
    return sec;
  }

  let good_pd_date = new Date("2017-02-19T19:06:09.000Z");
  if (isNaN(good_pd_date.getFullYear())) good_pd_date = new Date("2/19/17");
  const good_pd = good_pd_date.getFullYear() == 2017;
  /* parses a date as a local date */
  function parseDate(str, fixdate) {
    const d = new Date(str);
    if (good_pd) {
      if (fixdate > 0)
        d.setTime(d.getTime() + d.getTimezoneOffset() * 60 * 1000);
      else if (fixdate < 0)
        d.setTime(d.getTime() - d.getTimezoneOffset() * 60 * 1000);
      return d;
    }
    if (str instanceof Date) return str;
    if (good_pd_date.getFullYear() == 1917 && !isNaN(d.getFullYear())) {
      const s = d.getFullYear();
      if (str.indexOf(`${s}`) > -1) return d;
      d.setFullYear(d.getFullYear() + 100);
      return d;
    }
    const n = str.match(/\d+/g) || ["2017", "2", "19", "0", "0", "0"];
    let out = new Date(
      +n[0],
      +n[1] - 1,
      +n[2],
      +n[3] || 0,
      +n[4] || 0,
      +n[5] || 0
    );
    if (str.indexOf("Z") > -1)
      out = new Date(out.getTime() - out.getTimezoneOffset() * 60 * 1000);
    return out;
  }

  function cc2str(arr) {
    let o = "";
    for (let i = 0; i != arr.length; ++i) o += String.fromCharCode(arr[i]);
    return o;
  }

  function dup(o) {
    if (typeof JSON !== "undefined" && !Array.isArray(o))
      return JSON.parse(JSON.stringify(o));
    if (typeof o !== "object" || o == null) return o;
    if (o instanceof Date) return new Date(o.getTime());
    const out = {};
    for (const k in o)
      if (Object.prototype.hasOwnProperty.call(o, k)) out[k] = dup(o[k]);
    return out;
  }

  function fill(c, l) {
    let o = "";
    while (o.length < l) o += c;
    return o;
  }

  /* TODO: stress test */
  function fuzzynum(s) {
    let v = Number(s);
    if (!isNaN(v)) return v;
    let wt = 1;
    let ss = s
      .replace(/([\d]),([\d])/g, "$1$2")
      .replace(/[$]/g, "")
      .replace(/[%]/g, function() {
        wt *= 100;
        return "";
      });
    if (!isNaN((v = Number(ss)))) return v / wt;
    ss = ss.replace(/[(](.*)[)]/, function($$, $1) {
      wt = -wt;
      return $1;
    });
    if (!isNaN((v = Number(ss)))) return v / wt;
    return v;
  }
  function fuzzydate(s) {
    const o = new Date(s);
    const n = new Date(NaN);
    const y = o.getYear();
    const m = o.getMonth();
    const d = o.getDate();
    if (isNaN(d)) return n;
    if (y < 0 || y > 8099) return n;
    if ((m > 0 || d > 1) && y != 101) return o;
    if (
      s.toLowerCase().match(/jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/)
    )
      return o;
    if (s.match(/[^-0-9:,\/\\]/)) return n;
    return o;
  }

  const safe_split_regex = "abacaba".split(/(:?b)/i).length == 5;
  function split_regex(str, re, def) {
    if (safe_split_regex || typeof re === "string") return str.split(re);
    const p = str.split(re);
    const o = [p[0]];
    for (let i = 1; i < p.length; ++i) {
      o.push(def);
      o.push(p[i]);
    }
    return o;
  }
  function getdatastr(data) {
    if (!data) return null;
    if (data.data) return debom(data.data);
    if (data.asNodeBuffer && has_buf)
      return debom(data.asNodeBuffer().toString("binary"));
    if (data.asBinary) return debom(data.asBinary());
    if (data._data && data._data.getContent)
      return debom(
        cc2str(Array.prototype.slice.call(data._data.getContent(), 0))
      );
    if (data.content && data.type) return debom(cc2str(data.content));
    return null;
  }

  function getdatabin(data) {
    if (!data) return null;
    if (data.data) return char_codes(data.data);
    if (data.asNodeBuffer && has_buf) return data.asNodeBuffer();
    if (data._data && data._data.getContent) {
      const o = data._data.getContent();
      if (typeof o === "string") return char_codes(o);
      return Array.prototype.slice.call(o);
    }
    if (data.content && data.type) return data.content;
    return null;
  }

  function getdata(data) {
    return data && data.name.slice(-4) === ".bin"
      ? getdatabin(data)
      : getdatastr(data);
  }

  /* Part 2 Section 10.1.2 "Mapping Content Types" Names are case-insensitive */
  /* OASIS does not comment on filename case sensitivity */
  function safegetzipfile(zip, file) {
    const k = zip.FullPaths || keys(zip.files);
    const f = file.toLowerCase();
    const g = f.replace(/\//g, "\\");
    for (let i = 0; i < k.length; ++i) {
      const n = k[i].toLowerCase();
      if (f == n || g == n) return zip.files[k[i]];
    }
    return null;
  }

  function getzipfile(zip, file) {
    const o = safegetzipfile(zip, file);
    if (o == null) throw new Error(`Cannot find file ${file} in zip`);
    return o;
  }

  function getzipdata(zip, file, safe) {
    if (!safe) return getdata(getzipfile(zip, file));
    if (!file) return null;
    try {
      return getzipdata(zip, file);
    } catch (e) {
      return null;
    }
  }

  function getzipstr(zip, file, safe) {
    if (!safe) return getdatastr(getzipfile(zip, file));
    if (!file) return null;
    try {
      return getzipstr(zip, file);
    } catch (e) {
      return null;
    }
  }

  function zipentries(zip) {
    const k = zip.FullPaths || keys(zip.files);
    const o = [];
    for (let i = 0; i < k.length; ++i) if (k[i].slice(-1) != "/") o.push(k[i]);
    return o.sort();
  }

  let jszip;
  /* global JSZipSync:true */
  if (typeof JSZipSync !== "undefined") jszip = JSZipSync;
  if (typeof exports !== "undefined") {
    if (typeof module !== "undefined" && module.exports) {
      if (typeof jszip === "undefined") jszip = require("./jszip.js");
    }
  }

  function zip_read(d, o) {
    let zip;
    if (jszip)
      switch (o.type) {
        case "base64":
          zip = new jszip(d, { base64: true });
          break;
        case "binary":
        case "array":
          zip = new jszip(d, { base64: false });
          break;
        case "buffer":
          zip = new jszip(d);
          break;
        default:
          throw new Error(`Unrecognized type ${o.type}`);
      }
    else
      switch (o.type) {
        case "base64":
          zip = CFB.read(d, { type: "base64" });
          break;
        case "binary":
          zip = CFB.read(d, { type: "binary" });
          break;
        case "buffer":
        case "array":
          zip = CFB.read(d, { type: "buffer" });
          break;
        default:
          throw new Error(`Unrecognized type ${o.type}`);
      }
    return zip;
  }

  function resolve_path(path, base) {
    if (path.charAt(0) == "/") return path.slice(1);
    const result = base.split("/");
    if (base.slice(-1) != "/") result.pop(); // folder path
    const target = path.split("/");
    while (target.length !== 0) {
      const step = target.shift();
      if (step === "..") result.pop();
      else if (step !== ".") result.push(step);
    }
    return result.join("/");
  }
  const XML_HEADER =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
  const attregexg = /([^"\s?>\/]+)\s*=\s*((?:")([^"]*)(?:")|(?:')([^']*)(?:')|([^'">\s]+))/g;
  let tagregex = /<[\/\?]?[a-zA-Z0-9:_-]+(?:\s+[^"\s?>\/]+\s*=\s*(?:"[^"]*"|'[^']*'|[^'">\s=]+))*\s?[\/\?]?>/gm;

  if (!XML_HEADER.match(tagregex)) tagregex = /<[^>]*>/g;
  const nsregex = /<\w*:/;
  const nsregex2 = /<(\/?)\w+:/;
  function parsexmltag(tag, skip_root, skip_LC) {
    const z = {};
    let eq = 0;
    let c = 0;
    for (; eq !== tag.length; ++eq)
      if ((c = tag.charCodeAt(eq)) === 32 || c === 10 || c === 13) break;
    if (!skip_root) z[0] = tag.slice(0, eq);
    if (eq === tag.length) return z;
    const m = tag.match(attregexg);
    let j = 0;
    let v = "";
    let i = 0;
    let q = "";
    let cc = "";
    let quot = 1;
    if (m)
      for (i = 0; i != m.length; ++i) {
        cc = m[i];
        for (c = 0; c != cc.length; ++c) if (cc.charCodeAt(c) === 61) break;
        q = cc.slice(0, c).trim();
        while (cc.charCodeAt(c + 1) == 32) ++c;
        quot = (eq = cc.charCodeAt(c + 1)) == 34 || eq == 39 ? 1 : 0;
        v = cc.slice(c + 1 + quot, cc.length - quot);
        for (j = 0; j != q.length; ++j) if (q.charCodeAt(j) === 58) break;
        if (j === q.length) {
          if (q.indexOf("_") > 0) q = q.slice(0, q.indexOf("_")); // from ods
          z[q] = v;
          if (!skip_LC) z[q.toLowerCase()] = v;
        } else {
          const k =
            (j === 5 && q.slice(0, 5) === "xmlns" ? "xmlns" : "") +
            q.slice(j + 1);
          if (z[k] && q.slice(j - 3, j) == "ext") continue; // from ods
          z[k] = v;
          if (!skip_LC) z[k.toLowerCase()] = v;
        }
      }
    return z;
  }
  function strip_ns(x) {
    return x.replace(nsregex2, "<$1");
  }

  const encodings = {
    "&quot;": '"',
    "&apos;": "'",
    "&gt;": ">",
    "&lt;": "<",
    "&amp;": "&"
  };
  const rencoding = evert(encodings);

  // TODO: CP remap (need to read file version to determine OS)
  const unescapexml = (function() {
    /* 22.4.2.4 bstr (Basic String) */
    const encregex = /&(?:quot|apos|gt|lt|amp|#x?([\da-fA-F]+));/gi;
    const coderegex = /_x([\da-fA-F]{4})_/gi;
    return function unescapexml(text) {
      const s = `${text}`;
      const i = s.indexOf("<![CDATA[");
      if (i == -1)
        return s
          .replace(encregex, function($$, $1) {
            return (
              encodings[$$] ||
              String.fromCharCode(
                parseInt($1, $$.indexOf("x") > -1 ? 16 : 10)
              ) ||
              $$
            );
          })
          .replace(coderegex, function(m, c) {
            return String.fromCharCode(parseInt(c, 16));
          });
      const j = s.indexOf("]]>");
      return (
        unescapexml(s.slice(0, i)) +
        s.slice(i + 9, j) +
        unescapexml(s.slice(j + 3))
      );
    };
  })();

  const decregex = /[&<>'"]/g;

  const htmlcharegex = /[\u0000-\u001f]/g;
  function escapehtml(text) {
    const s = `${text}`;
    return s
      .replace(decregex, function(y) {
        return rencoding[y];
      })
      .replace(/\n/g, "<br/>")
      .replace(htmlcharegex, function(s) {
        return `&#x${`000${s.charCodeAt(0).toString(16)}`.slice(-4)};`;
      });
  }

  /* TODO: handle codepages */
  const xlml_fixstr = (function() {
    const entregex = /&#(\d+);/g;
    function entrepl($$, $1) {
      return String.fromCharCode(parseInt($1, 10));
    }
    return function xlml_fixstr(str) {
      return str.replace(entregex, entrepl);
    };
  })();

  function parsexmlbool(value) {
    switch (value) {
      case 1:
      case true:
      case "1":
      case "true":
      case "TRUE":
        return true;
      /* case '0': case 'false': case 'FALSE': */
      default:
        return false;
    }
  }

  let utf8read = function utf8reada(orig) {
    let out = "";
    let i = 0;
    let c = 0;
    let d = 0;
    let e = 0;
    let f = 0;
    let w = 0;
    while (i < orig.length) {
      c = orig.charCodeAt(i++);
      if (c < 128) {
        out += String.fromCharCode(c);
        continue;
      }
      d = orig.charCodeAt(i++);
      if (c > 191 && c < 224) {
        f = (c & 31) << 6;
        f |= d & 63;
        out += String.fromCharCode(f);
        continue;
      }
      e = orig.charCodeAt(i++);
      if (c < 240) {
        out += String.fromCharCode(
          ((c & 15) << 12) | ((d & 63) << 6) | (e & 63)
        );
        continue;
      }
      f = orig.charCodeAt(i++);
      w =
        (((c & 7) << 18) | ((d & 63) << 12) | ((e & 63) << 6) | (f & 63)) -
        65536;
      out += String.fromCharCode(0xd800 + ((w >>> 10) & 1023));
      out += String.fromCharCode(0xdc00 + (w & 1023));
    }
    return out;
  };

  var utf8write = function(orig) {
    const out = [];
    let i = 0;
    let c = 0;
    let d = 0;
    while (i < orig.length) {
      c = orig.charCodeAt(i++);
      switch (true) {
        case c < 128:
          out.push(String.fromCharCode(c));
          break;
        case c < 2048:
          out.push(String.fromCharCode(192 + (c >> 6)));
          out.push(String.fromCharCode(128 + (c & 63)));
          break;
        case c >= 55296 && c < 57344:
          c -= 55296;
          d = orig.charCodeAt(i++) - 56320 + (c << 10);
          out.push(String.fromCharCode(240 + ((d >> 18) & 7)));
          out.push(String.fromCharCode(144 + ((d >> 12) & 63)));
          out.push(String.fromCharCode(128 + ((d >> 6) & 63)));
          out.push(String.fromCharCode(128 + (d & 63)));
          break;
        default:
          out.push(String.fromCharCode(224 + (c >> 12)));
          out.push(String.fromCharCode(128 + ((c >> 6) & 63)));
          out.push(String.fromCharCode(128 + (c & 63)));
      }
    }
    return out.join("");
  };

  if (has_buf) {
    const utf8readb = function utf8readb(data) {
      const out = Buffer.alloc(2 * data.length);
      let w;
      let i;
      let j = 1;
      let k = 0;
      let ww = 0;
      let c;
      for (i = 0; i < data.length; i += j) {
        j = 1;
        if ((c = data.charCodeAt(i)) < 128) w = c;
        else if (c < 224) {
          w = (c & 31) * 64 + (data.charCodeAt(i + 1) & 63);
          j = 2;
        } else if (c < 240) {
          w =
            (c & 15) * 4096 +
            (data.charCodeAt(i + 1) & 63) * 64 +
            (data.charCodeAt(i + 2) & 63);
          j = 3;
        } else {
          j = 4;
          w =
            (c & 7) * 262144 +
            (data.charCodeAt(i + 1) & 63) * 4096 +
            (data.charCodeAt(i + 2) & 63) * 64 +
            (data.charCodeAt(i + 3) & 63);
          w -= 65536;
          ww = 0xd800 + ((w >>> 10) & 1023);
          w = 0xdc00 + (w & 1023);
        }
        if (ww !== 0) {
          out[k++] = ww & 255;
          out[k++] = ww >>> 8;
          ww = 0;
        }
        out[k++] = w % 256;
        out[k++] = w >>> 8;
      }
      return out.slice(0, k).toString("ucs2");
    };
    const corpus = "foo bar baz\u00e2\u0098\u0083\u00f0\u009f\u008d\u00a3";
    if (utf8read(corpus) == utf8readb(corpus)) utf8read = utf8readb;
    const utf8readc = function utf8readc(data) {
      return Buffer_from(data, "binary").toString("utf8");
    };
    if (utf8read(corpus) == utf8readc(corpus)) utf8read = utf8readc;

    utf8write = function(data) {
      return Buffer_from(data, "utf8").toString("binary");
    };
  }

  // matches <foo>...</foo> extracts content
  const matchtag = (function() {
    const mtcache = {};
    return function matchtag(f, g) {
      const t = `${f}|${g || ""}`;
      if (mtcache[t]) return mtcache[t];
      return (mtcache[t] = new RegExp(
        `<(?:\\w+:)?${f}(?: xml:space="preserve")?(?:[^>]*)>([\\s\\S]*?)</(?:\\w+:)?${f}>`,
        g || ""
      ));
    };
  })();

  const htmldecode = (function() {
    const entities = [
      ["nbsp", " "],
      ["middot", "·"],
      ["quot", '"'],
      ["apos", "'"],
      ["gt", ">"],
      ["lt", "<"],
      ["amp", "&"]
    ].map(function(x) {
      return [new RegExp(`&${x[0]};`, "ig"), x[1]];
    });
    return function htmldecode(str) {
      let o = str
        // Remove new lines and spaces from start of content
        .replace(/^[\t\n\r ]+/, "")
        // Remove new lines and spaces from end of content
        .replace(/[\t\n\r ]+$/, "")
        // Added line which removes any white space characters after and before html tags
        .replace(/>\s+/g, ">")
        .replace(/\s+</g, "<")
        // Replace remaining new lines and spaces with space
        .replace(/[\t\n\r ]+/g, " ")
        // Replace <br> tags with new lines
        .replace(/<\s*[bB][rR]\s*\/?>/g, "\n")
        // Strip HTML elements
        .replace(/<[^>]*>/g, "");
      for (let i = 0; i < entities.length; ++i)
        o = o.replace(entities[i][0], entities[i][1]);
      return o;
    };
  })();

  const vtregex = (function() {
    const vt_cache = {};
    return function vt_regex(bt) {
      if (vt_cache[bt] !== undefined) return vt_cache[bt];
      return (vt_cache[bt] = new RegExp(
        `<(?:vt:)?${bt}>([\\s\\S]*?)</(?:vt:)?${bt}>`,
        "g"
      ));
    };
  })();
  const vtvregex = /<\/?(?:vt:)?variant>/g;
  const vtmregex = /<(?:vt:)([^>]*)>([\s\S]*)</;
  function parseVector(data, opts) {
    const h = parsexmltag(data);

    const matches = data.match(vtregex(h.baseType)) || [];
    const res = [];
    if (matches.length != h.size) {
      if (opts.WTF)
        throw new Error(
          `unexpected vector length ${matches.length} != ${h.size}`
        );
      return res;
    }
    matches.forEach(function(x) {
      const v = x.replace(vtvregex, "").match(vtmregex);
      if (v) res.push({ v: utf8read(v[2]), t: v[1] });
    });
    return res;
  }

  const wtregex = /(^\s|\s$|\n)/;

  function wxt_helper(h) {
    return keys(h)
      .map(function(k) {
        return ` ${k}="${h[k]}"`;
      })
      .join("");
  }
  function writextag(f, g, h) {
    return `<${f}${h != null ? wxt_helper(h) : ""}${
      g != null
        ? `${g.match(wtregex) ? ' xml:space="preserve"' : ""}>${g}</${f}`
        : "/"
    }>`;
  }

  const XMLNS = {
    dc: "http://purl.org/dc/elements/1.1/",
    dcterms: "http://purl.org/dc/terms/",
    dcmitype: "http://purl.org/dc/dcmitype/",
    mx: "http://schemas.microsoft.com/office/mac/excel/2008/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    sjs:
      "http://schemas.openxmlformats.org/package/2006/sheetjs/core-properties",
    vt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    xsi: "http://www.w3.org/2001/XMLSchema-instance",
    xsd: "http://www.w3.org/2001/XMLSchema"
  };

  XMLNS.main = [
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://purl.oclc.org/ooxml/spreadsheetml/main",
    "http://schemas.microsoft.com/office/excel/2006/main",
    "http://schemas.microsoft.com/office/excel/2006/2"
  ];

  function read_double_le(b, idx) {
    const s = 1 - 2 * (b[idx + 7] >>> 7);
    let e = ((b[idx + 7] & 0x7f) << 4) + ((b[idx + 6] >>> 4) & 0x0f);
    let m = b[idx + 6] & 0x0f;
    for (let i = 5; i >= 0; --i) m = m * 256 + b[idx + i];
    if (e == 0x7ff) return m == 0 ? s * Infinity : NaN;
    if (e == 0) e = -1022;
    else {
      e -= 1023;
      m += Math.pow(2, 52);
    }
    return s * Math.pow(2, e - 52) * m;
  }

  function write_double_le(b, v, idx) {
    const bs = (v < 0 || 1 / v == -Infinity ? 1 : 0) << 7;
    let e = 0;
    let m = 0;
    const av = bs ? -v : v;
    if (!isFinite(av)) {
      e = 0x7ff;
      m = isNaN(v) ? 0x6969 : 0;
    } else if (av == 0) e = m = 0;
    else {
      e = Math.floor(Math.log(av) / Math.LN2);
      m = av * Math.pow(2, 52 - e);
      if (e <= -1023 && (!isFinite(m) || m < Math.pow(2, 52))) {
        e = -1022;
      } else {
        m -= Math.pow(2, 52);
        e += 1023;
      }
    }
    for (let i = 0; i <= 5; ++i, m /= 256) b[idx + i] = m & 0xff;
    b[idx + 6] = ((e & 0x0f) << 4) | (m & 0xf);
    b[idx + 7] = (e >> 4) | bs;
  }

  var __toBuffer = function(bufs) {
    const x = [];
    const w = 10240;
    for (let i = 0; i < bufs[0].length; ++i)
      if (bufs[0][i])
        for (let j = 0, L = bufs[0][i].length; j < L; j += w)
          x.push.apply(x, bufs[0][i].slice(j, j + w));
    return x;
  };
  const ___toBuffer = __toBuffer;
  var __utf16le = function(b, s, e) {
    const ss = [];
    for (let i = s; i < e; i += 2)
      ss.push(String.fromCharCode(__readUInt16LE(b, i)));
    return ss.join("").replace(chr0, "");
  };
  const ___utf16le = __utf16le;
  let __hexlify = function(b, s, l) {
    const ss = [];
    for (let i = s; i < s + l; ++i) ss.push(`0${b[i].toString(16)}`.slice(-2));
    return ss.join("");
  };
  const ___hexlify = __hexlify;
  let __utf8 = function(b, s, e) {
    const ss = [];
    for (let i = s; i < e; i++) ss.push(String.fromCharCode(__readUInt8(b, i)));
    return ss.join("");
  };
  const ___utf8 = __utf8;
  let __lpstr = function(b, i) {
    const len = __readUInt32LE(b, i);
    return len > 0 ? __utf8(b, i + 4, i + 4 + len - 1) : "";
  };
  const ___lpstr = __lpstr;
  let __cpstr = function(b, i) {
    const len = __readUInt32LE(b, i);
    return len > 0 ? __utf8(b, i + 4, i + 4 + len - 1) : "";
  };
  const ___cpstr = __cpstr;
  let __lpwstr = function(b, i) {
    const len = 2 * __readUInt32LE(b, i);
    return len > 0 ? __utf8(b, i + 4, i + 4 + len - 1) : "";
  };
  const ___lpwstr = __lpwstr;
  let __lpp4;
  let ___lpp4;
  __lpp4 = ___lpp4 = function lpp4_(b, i) {
    const len = __readUInt32LE(b, i);
    return len > 0 ? __utf16le(b, i + 4, i + 4 + len) : "";
  };
  let __8lpp4 = function(b, i) {
    const len = __readUInt32LE(b, i);
    return len > 0 ? __utf8(b, i + 4, i + 4 + len) : "";
  };
  const ___8lpp4 = __8lpp4;
  let __double;
  let ___double;
  __double = ___double = function(b, idx) {
    return read_double_le(b, idx);
  };
  let is_buf = function is_buf_a(a) {
    return Array.isArray(a);
  };

  if (has_buf) {
    __utf16le = function(b, s, e) {
      if (!Buffer.isBuffer(b)) return ___utf16le(b, s, e);
      return b
        .toString("utf16le", s, e)
        .replace(chr0, "") /* .replace(chr1,'!') */;
    };
    __hexlify = function(b, s, l) {
      return Buffer.isBuffer(b)
        ? b.toString("hex", s, s + l)
        : ___hexlify(b, s, l);
    };
    __lpstr = function lpstr_b(b, i) {
      if (!Buffer.isBuffer(b)) return ___lpstr(b, i);
      const len = b.readUInt32LE(i);
      return len > 0 ? b.toString("utf8", i + 4, i + 4 + len - 1) : "";
    };
    __cpstr = function cpstr_b(b, i) {
      if (!Buffer.isBuffer(b)) return ___cpstr(b, i);
      const len = b.readUInt32LE(i);
      return len > 0 ? b.toString("utf8", i + 4, i + 4 + len - 1) : "";
    };
    __lpwstr = function lpwstr_b(b, i) {
      if (!Buffer.isBuffer(b)) return ___lpwstr(b, i);
      const len = 2 * b.readUInt32LE(i);
      return b.toString("utf16le", i + 4, i + 4 + len - 1);
    };
    __lpp4 = function lpp4_b(b, i) {
      if (!Buffer.isBuffer(b)) return ___lpp4(b, i);
      const len = b.readUInt32LE(i);
      return b.toString("utf16le", i + 4, i + 4 + len);
    };
    __8lpp4 = function lpp4_8b(b, i) {
      if (!Buffer.isBuffer(b)) return ___8lpp4(b, i);
      const len = b.readUInt32LE(i);
      return b.toString("utf8", i + 4, i + 4 + len);
    };
    __utf8 = function utf8_b(b, s, e) {
      return Buffer.isBuffer(b) ? b.toString("utf8", s, e) : ___utf8(b, s, e);
    };
    __toBuffer = function(bufs) {
      return bufs[0].length > 0 && Buffer.isBuffer(bufs[0][0])
        ? Buffer.concat(bufs[0])
        : ___toBuffer(bufs);
    };
    bconcat = function(bufs) {
      return Buffer.isBuffer(bufs[0])
        ? Buffer.concat(bufs)
        : [].concat.apply([], bufs);
    };
    __double = function double_(b, i) {
      if (Buffer.isBuffer(b)) return b.readDoubleLE(i);
      return ___double(b, i);
    };
    is_buf = function is_buf_b(a) {
      return Buffer.isBuffer(a) || Array.isArray(a);
    };
  }

  /* from js-xls */
  if (typeof cptable !== "undefined") {
    __utf16le = function(b, s, e) {
      return cptable.utils.decode(1200, b.slice(s, e)).replace(chr0, "");
    };
    __utf8 = function(b, s, e) {
      return cptable.utils.decode(65001, b.slice(s, e));
    };
    __lpstr = function(b, i) {
      const len = __readUInt32LE(b, i);
      return len > 0
        ? cptable.utils.decode(current_ansi, b.slice(i + 4, i + 4 + len - 1))
        : "";
    };
    __cpstr = function(b, i) {
      const len = __readUInt32LE(b, i);
      return len > 0
        ? cptable.utils.decode(
            current_codepage,
            b.slice(i + 4, i + 4 + len - 1)
          )
        : "";
    };
    __lpwstr = function(b, i) {
      const len = 2 * __readUInt32LE(b, i);
      return len > 0
        ? cptable.utils.decode(1200, b.slice(i + 4, i + 4 + len - 1))
        : "";
    };
    __lpp4 = function(b, i) {
      const len = __readUInt32LE(b, i);
      return len > 0
        ? cptable.utils.decode(1200, b.slice(i + 4, i + 4 + len))
        : "";
    };
    __8lpp4 = function(b, i) {
      const len = __readUInt32LE(b, i);
      return len > 0
        ? cptable.utils.decode(65001, b.slice(i + 4, i + 4 + len))
        : "";
    };
  }

  var __readUInt8 = function(b, idx) {
    return b[idx];
  };
  var __readUInt16LE = function(b, idx) {
    return b[idx + 1] * (1 << 8) + b[idx];
  };
  const __readInt16LE = function(b, idx) {
    const u = b[idx + 1] * (1 << 8) + b[idx];
    return u < 0x8000 ? u : (0xffff - u + 1) * -1;
  };
  var __readUInt32LE = function(b, idx) {
    return (
      b[idx + 3] * (1 << 24) + (b[idx + 2] << 16) + (b[idx + 1] << 8) + b[idx]
    );
  };
  var __readInt32LE = function(b, idx) {
    return (b[idx + 3] << 24) | (b[idx + 2] << 16) | (b[idx + 1] << 8) | b[idx];
  };
  const __readInt32BE = function(b, idx) {
    return (b[idx] << 24) | (b[idx + 1] << 16) | (b[idx + 2] << 8) | b[idx + 3];
  };

  function ReadShift(size, t) {
    let o = "";
    let oI;
    let oR;
    const oo = [];
    let w;
    let vv;
    let i;
    let loc;
    switch (t) {
      case "dbcs":
        loc = this.l;
        if (has_buf && Buffer.isBuffer(this))
          o = this.slice(this.l, this.l + 2 * size).toString("utf16le");
        else
          for (i = 0; i < size; ++i) {
            o += String.fromCharCode(__readUInt16LE(this, loc));
            loc += 2;
          }
        size *= 2;
        break;

      case "utf8":
        o = __utf8(this, this.l, this.l + size);
        break;
      case "utf16le":
        size *= 2;
        o = __utf16le(this, this.l, this.l + size);
        break;

      case "wstr":
        if (typeof cptable !== "undefined")
          o = cptable.utils.decode(
            current_codepage,
            this.slice(this.l, this.l + 2 * size)
          );
        else return ReadShift.call(this, size, "dbcs");
        size = 2 * size;
        break;

      /* [MS-OLEDS] 2.1.4 LengthPrefixedAnsiString */
      case "lpstr-ansi":
        o = __lpstr(this, this.l);
        size = 4 + __readUInt32LE(this, this.l);
        break;
      case "lpstr-cp":
        o = __cpstr(this, this.l);
        size = 4 + __readUInt32LE(this, this.l);
        break;
      /* [MS-OLEDS] 2.1.5 LengthPrefixedUnicodeString */
      case "lpwstr":
        o = __lpwstr(this, this.l);
        size = 4 + 2 * __readUInt32LE(this, this.l);
        break;
      /* [MS-OFFCRYPTO] 2.1.2 Length-Prefixed Padded Unicode String (UNICODE-LP-P4) */
      case "lpp4":
        size = 4 + __readUInt32LE(this, this.l);
        o = __lpp4(this, this.l);
        if (size & 0x02) size += 2;
        break;
      /* [MS-OFFCRYPTO] 2.1.3 Length-Prefixed UTF-8 String (UTF-8-LP-P4) */
      case "8lpp4":
        size = 4 + __readUInt32LE(this, this.l);
        o = __8lpp4(this, this.l);
        if (size & 0x03) size += 4 - (size & 0x03);
        break;

      case "cstr":
        size = 0;
        o = "";
        while ((w = __readUInt8(this, this.l + size++)) !== 0)
          oo.push(_getchar(w));
        o = oo.join("");
        break;
      case "_wstr":
        size = 0;
        o = "";
        while ((w = __readUInt16LE(this, this.l + size)) !== 0) {
          oo.push(_getchar(w));
          size += 2;
        }
        size += 2;
        o = oo.join("");
        break;

      /* sbcs and dbcs support continue records in the SST way TODO codepages */
      case "dbcs-cont":
        o = "";
        loc = this.l;
        for (i = 0; i < size; ++i) {
          if (this.lens && this.lens.indexOf(loc) !== -1) {
            w = __readUInt8(this, loc);
            this.l = loc + 1;
            vv = ReadShift.call(this, size - i, w ? "dbcs-cont" : "sbcs-cont");
            return oo.join("") + vv;
          }
          oo.push(_getchar(__readUInt16LE(this, loc)));
          loc += 2;
        }
        o = oo.join("");
        size *= 2;
        break;

      case "cpstr":
        if (typeof cptable !== "undefined") {
          o = cptable.utils.decode(
            current_codepage,
            this.slice(this.l, this.l + size)
          );
          break;
        }
      /* falls through */
      case "sbcs-cont":
        o = "";
        loc = this.l;
        for (i = 0; i != size; ++i) {
          if (this.lens && this.lens.indexOf(loc) !== -1) {
            w = __readUInt8(this, loc);
            this.l = loc + 1;
            vv = ReadShift.call(this, size - i, w ? "dbcs-cont" : "sbcs-cont");
            return oo.join("") + vv;
          }
          oo.push(_getchar(__readUInt8(this, loc)));
          loc += 1;
        }
        o = oo.join("");
        break;

      default:
        switch (size) {
          case 1:
            oI = __readUInt8(this, this.l);
            this.l++;
            return oI;
          case 2:
            oI = (t === "i" ? __readInt16LE : __readUInt16LE)(this, this.l);
            this.l += 2;
            return oI;
          case 4:
          case -4:
            if (t === "i" || (this[this.l + 3] & 0x80) === 0) {
              oI = (size > 0 ? __readInt32LE : __readInt32BE)(this, this.l);
              this.l += 4;
              return oI;
            } else {
              oR = __readUInt32LE(this, this.l);
              this.l += 4;
            }
            return oR;
          case 8:
          case -8:
            if (t === "f") {
              if (size == 8) oR = __double(this, this.l);
              else
                oR = __double(
                  [
                    this[this.l + 7],
                    this[this.l + 6],
                    this[this.l + 5],
                    this[this.l + 4],
                    this[this.l + 3],
                    this[this.l + 2],
                    this[this.l + 1],
                    this[this.l + 0]
                  ],
                  0
                );
              this.l += 8;
              return oR;
            } else size = 8;
          /* falls through */
          case 16:
            o = __hexlify(this, this.l, size);
            break;
        }
    }
    this.l += size;
    return o;
  }

  const __writeUInt32LE = function(b, val, idx) {
    b[idx] = val & 0xff;
    b[idx + 1] = (val >>> 8) & 0xff;
    b[idx + 2] = (val >>> 16) & 0xff;
    b[idx + 3] = (val >>> 24) & 0xff;
  };
  const __writeInt32LE = function(b, val, idx) {
    b[idx] = val & 0xff;
    b[idx + 1] = (val >> 8) & 0xff;
    b[idx + 2] = (val >> 16) & 0xff;
    b[idx + 3] = (val >> 24) & 0xff;
  };
  const __writeUInt16LE = function(b, val, idx) {
    b[idx] = val & 0xff;
    b[idx + 1] = (val >>> 8) & 0xff;
  };

  function WriteShift(t, val, f) {
    let size = 0;
    let i = 0;
    if (f === "dbcs") {
      for (i = 0; i != val.length; ++i)
        __writeUInt16LE(this, val.charCodeAt(i), this.l + 2 * i);
      size = 2 * val.length;
    } else if (f === "sbcs") {
      if (typeof cptable !== "undefined" && current_ansi == 874) {
        /* TODO: use tables directly, don't encode */
        for (i = 0; i != val.length; ++i) {
          const cppayload = cptable.utils.encode(current_ansi, val.charAt(i));
          this[this.l + i] = cppayload[0];
        }
      } else {
        val = val.replace(/[^\x00-\x7F]/g, "_");
        for (i = 0; i != val.length; ++i)
          this[this.l + i] = val.charCodeAt(i) & 0xff;
      }
      size = val.length;
    } else if (f === "hex") {
      for (; i < t; ++i) {
        this[this.l++] = parseInt(val.slice(2 * i, 2 * i + 2), 16) || 0;
      }
      return this;
    } else if (f === "utf16le") {
      const end = Math.min(this.l + t, this.length);
      for (i = 0; i < Math.min(val.length, t); ++i) {
        const cc = val.charCodeAt(i);
        this[this.l++] = cc & 0xff;
        this[this.l++] = cc >> 8;
      }
      while (this.l < end) this[this.l++] = 0;
      return this;
    } else
      switch (t) {
        case 1:
          size = 1;
          this[this.l] = val & 0xff;
          break;
        case 2:
          size = 2;
          this[this.l] = val & 0xff;
          val >>>= 8;
          this[this.l + 1] = val & 0xff;
          break;
        case 3:
          size = 3;
          this[this.l] = val & 0xff;
          val >>>= 8;
          this[this.l + 1] = val & 0xff;
          val >>>= 8;
          this[this.l + 2] = val & 0xff;
          break;
        case 4:
          size = 4;
          __writeUInt32LE(this, val, this.l);
          break;
        case 8:
          size = 8;
          if (f === "f") {
            write_double_le(this, val, this.l);
            break;
          }
        /* falls through */
        case 16:
          break;
        case -4:
          size = 4;
          __writeInt32LE(this, val, this.l);
          break;
      }
    this.l += size;
    return this;
  }

  function CheckField(hexstr, fld) {
    const m = __hexlify(this, this.l, hexstr.length >> 1);
    if (m !== hexstr) throw new Error(`${fld}Expected ${hexstr} saw ${m}`);
    this.l += hexstr.length >> 1;
  }

  function prep_blob(blob, pos) {
    blob.l = pos;
    blob.read_shift = ReadShift;
    blob.chk = CheckField;
    blob.write_shift = WriteShift;
  }

  function parsenoop(blob, length) {
    blob.l += length;
  }

  function new_buf(sz) {
    const o = new_raw_buf(sz);
    prep_blob(o, 0);
    return o;
  }

  /* [MS-XLSB] 2.1.4 Record */
  function recordhopper(data, cb, opts) {
    if (!data) return;
    let tmpbyte;
    let cntbyte;
    let length;
    prep_blob(data, data.l || 0);
    const L = data.length;
    let RT = 0;
    let tgt = 0;
    while (data.l < L) {
      RT = data.read_shift(1);
      if (RT & 0x80) RT = (RT & 0x7f) + ((data.read_shift(1) & 0x7f) << 7);
      const R = XLSBRecordEnum[RT] || XLSBRecordEnum[0xffff];
      tmpbyte = data.read_shift(1);
      length = tmpbyte & 0x7f;
      for (cntbyte = 1; cntbyte < 4 && tmpbyte & 0x80; ++cntbyte)
        length += ((tmpbyte = data.read_shift(1)) & 0x7f) << (7 * cntbyte);
      tgt = data.l + length;
      const d = (R.f || parsenoop)(data, length, opts);
      data.l = tgt;
      if (cb(d, R.n, RT)) return;
    }
  }

  /* control buffer usage for fixed-length buffers */
  function buf_array() {
    const bufs = [];
    const blksz = has_buf ? 256 : 2048;
    const newblk = function ba_newblk(sz) {
      const o = new_buf(sz);
      prep_blob(o, 0);
      return o;
    };

    let curbuf = newblk(blksz);

    const endbuf = function ba_endbuf() {
      if (!curbuf) return;
      if (curbuf.length > curbuf.l) {
        curbuf = curbuf.slice(0, curbuf.l);
        curbuf.l = curbuf.length;
      }
      if (curbuf.length > 0) bufs.push(curbuf);
      curbuf = null;
    };

    const next = function ba_next(sz) {
      if (curbuf && sz < curbuf.length - curbuf.l) return curbuf;
      endbuf();
      return (curbuf = newblk(Math.max(sz + 1, blksz)));
    };

    const end = function ba_end() {
      endbuf();
      return __toBuffer([bufs]);
    };

    const push = function ba_push(buf) {
      endbuf();
      curbuf = buf;
      if (curbuf.l == null) curbuf.l = curbuf.length;
      next(blksz);
    };

    return { next, push, end, _bufs: bufs };
  }

  function write_record(ba, type, payload, length) {
    const t = +XLSBRE[type];
    let l;
    if (isNaN(t)) return; // TODO: throw something here?
    if (!length) length = XLSBRecordEnum[t].p || (payload || []).length || 0;
    l = 1 + (t >= 0x80 ? 1 : 0) + 1 /* + length */;
    if (length >= 0x80) ++l;
    if (length >= 0x4000) ++l;
    if (length >= 0x200000) ++l;
    const o = ba.next(l);
    if (t <= 0x7f) o.write_shift(1, t);
    else {
      o.write_shift(1, (t & 0x7f) + 0x80);
      o.write_shift(1, t >> 7);
    }
    for (let i = 0; i != 4; ++i) {
      if (length >= 0x80) {
        o.write_shift(1, (length & 0x7f) + 0x80);
        length >>= 7;
      } else {
        o.write_shift(1, length);
        break;
      }
    }
    if (length > 0 && is_buf(payload)) ba.push(payload);
  }
  /* XLS ranges enforced */
  function shift_cell_xls(cell, tgt, opts) {
    const out = dup(cell);
    if (tgt.s) {
      if (out.cRel) out.c += tgt.s.c;
      if (out.rRel) out.r += tgt.s.r;
    } else {
      if (out.cRel) out.c += tgt.c;
      if (out.rRel) out.r += tgt.r;
    }
    if (!opts || opts.biff < 12) {
      while (out.c >= 0x100) out.c -= 0x100;
      while (out.r >= 0x10000) out.r -= 0x10000;
    }
    return out;
  }

  function shift_range_xls(cell, range, opts) {
    const out = dup(cell);
    out.s = shift_cell_xls(out.s, range.s, opts);
    out.e = shift_cell_xls(out.e, range.s, opts);
    return out;
  }

  function encode_cell_xls(c, biff) {
    if (c.cRel && c.c < 0) {
      c = dup(c);
      while (c.c < 0) c.c += biff > 8 ? 0x4000 : 0x100;
    }
    if (c.rRel && c.r < 0) {
      c = dup(c);
      while (c.r < 0) c.r += biff > 8 ? 0x100000 : biff > 5 ? 0x10000 : 0x4000;
    }
    let s = encode_cell(c);
    if (!c.cRel && c.cRel != null) s = fix_col(s);
    if (!c.rRel && c.rRel != null) s = fix_row(s);
    return s;
  }

  function encode_range_xls(r, opts) {
    if (r.s.r == 0 && !r.s.rRel) {
      if (
        r.e.r ==
          (opts.biff >= 12 ? 0xfffff : opts.biff >= 8 ? 0x10000 : 0x4000) &&
        !r.e.rRel
      ) {
        return `${(r.s.cRel ? "" : "$") + encode_col(r.s.c)}:${
          r.e.cRel ? "" : "$"
        }${encode_col(r.e.c)}`;
      }
    }
    if (r.s.c == 0 && !r.s.cRel) {
      if (r.e.c == (opts.biff >= 12 ? 0x3fff : 0xff) && !r.e.cRel) {
        return `${(r.s.rRel ? "" : "$") + encode_row(r.s.r)}:${
          r.e.rRel ? "" : "$"
        }${encode_row(r.e.r)}`;
      }
    }
    return `${encode_cell_xls(r.s, opts.biff)}:${encode_cell_xls(
      r.e,
      opts.biff
    )}`;
  }
  const OFFCRYPTO = {};

  const make_offcrypto = function(O, _crypto) {
    let crypto;
    if (typeof _crypto !== "undefined") crypto = _crypto;
    else if (typeof require !== "undefined") {
      try {
        crypto = require("crypto");
      } catch (e) {
        crypto = null;
      }
    }

    O.rc4 = function(key, data) {
      const S = new Array(256);
      let c = 0;
      let i = 0;
      let j = 0;
      let t = 0;
      for (i = 0; i != 256; ++i) S[i] = i;
      for (i = 0; i != 256; ++i) {
        j = (j + S[i] + key[i % key.length].charCodeAt(0)) & 255;
        t = S[i];
        S[i] = S[j];
        S[j] = t;
      }
      // $FlowIgnore
      i = j = 0;
      const out = new_raw_buf(data.length);
      for (c = 0; c != data.length; ++c) {
        i = (i + 1) & 255;
        j = (j + S[i]) % 256;
        t = S[i];
        S[i] = S[j];
        S[j] = t;
        out[c] = data[c] ^ S[(S[i] + S[j]) & 255];
      }
      return out;
    };

    O.md5 = function(hex) {
      if (!crypto) throw new Error("Unsupported crypto");
      return crypto
        .createHash("md5")
        .update(hex)
        .digest("hex");
    };
  };
  /* global crypto:true */
  make_offcrypto(OFFCRYPTO, typeof crypto !== "undefined" ? crypto : undefined);

  function decode_row(rowstr) {
    return parseInt(unfix_row(rowstr), 10) - 1;
  }
  function encode_row(row) {
    return `${row + 1}`;
  }
  function fix_row(cstr) {
    return cstr.replace(/([A-Z]|^)(\d+)$/, "$1$$$2");
  }
  function unfix_row(cstr) {
    return cstr.replace(/\$(\d+)$/, "$1");
  }

  function decode_col(colstr) {
    const c = unfix_col(colstr);
    let d = 0;
    let i = 0;
    for (; i !== c.length; ++i) d = 26 * d + c.charCodeAt(i) - 64;
    return d - 1;
  }
  function encode_col(col) {
    if (col < 0) throw new Error(`invalid column ${col}`);
    let s = "";
    for (++col; col; col = Math.floor((col - 1) / 26))
      s = String.fromCharCode(((col - 1) % 26) + 65) + s;
    return s;
  }
  function fix_col(cstr) {
    return cstr.replace(/^([A-Z])/, "$$$1");
  }
  function unfix_col(cstr) {
    return cstr.replace(/^\$([A-Z])/, "$1");
  }

  function decode_cell(cstr) {
    let R = 0;
    let C = 0;
    for (let i = 0; i < cstr.length; ++i) {
      const cc = cstr.charCodeAt(i);
      if (cc >= 48 && cc <= 57) R = 10 * R + (cc - 48);
      else if (cc >= 65 && cc <= 90) C = 26 * C + (cc - 64);
    }
    return { c: C - 1, r: R - 1 };
  }

  function encode_cell(cell) {
    let col = cell.c + 1;
    let s = "";
    for (; col; col = ((col - 1) / 26) | 0)
      s = String.fromCharCode(((col - 1) % 26) + 65) + s;
    return s + (cell.r + 1);
  }
  function decode_range(range) {
    const idx = range.indexOf(":");
    if (idx == -1) return { s: decode_cell(range), e: decode_cell(range) };
    return {
      s: decode_cell(range.slice(0, idx)),
      e: decode_cell(range.slice(idx + 1))
    };
  }
  function encode_range(cs, ce) {
    if (typeof ce === "undefined" || typeof ce === "number") {
      return encode_range(cs.s, cs.e);
    }
    if (typeof cs !== "string") cs = encode_cell(cs);
    if (typeof ce !== "string") ce = encode_cell(ce);
    return cs == ce ? cs : `${cs}:${ce}`;
  }

  function safe_decode_range(range) {
    const o = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };
    let idx = 0;
    let i = 0;
    let cc = 0;
    const len = range.length;
    for (idx = 0; i < len; ++i) {
      if ((cc = range.charCodeAt(i) - 64) < 1 || cc > 26) break;
      idx = 26 * idx + cc;
    }
    o.s.c = --idx;

    for (idx = 0; i < len; ++i) {
      if ((cc = range.charCodeAt(i) - 48) < 0 || cc > 9) break;
      idx = 10 * idx + cc;
    }
    o.s.r = --idx;

    if (i === len || range.charCodeAt(++i) === 58) {
      o.e.c = o.s.c;
      o.e.r = o.s.r;
      return o;
    }

    for (idx = 0; i != len; ++i) {
      if ((cc = range.charCodeAt(i) - 64) < 1 || cc > 26) break;
      idx = 26 * idx + cc;
    }
    o.e.c = --idx;

    for (idx = 0; i != len; ++i) {
      if ((cc = range.charCodeAt(i) - 48) < 0 || cc > 9) break;
      idx = 10 * idx + cc;
    }
    o.e.r = --idx;
    return o;
  }

  function safe_format_cell(cell, v) {
    const q = cell.t == "d" && v instanceof Date;
    if (cell.z != null)
      try {
        return (cell.w = SSF.format(cell.z, q ? datenum(v) : v));
      } catch (e) {}
    try {
      return (cell.w = SSF.format(
        (cell.XF || {}).numFmtId || (q ? 14 : 0),
        q ? datenum(v) : v
      ));
    } catch (e) {
      return `${v}`;
    }
  }

  function format_cell(cell, v, o) {
    if (cell == null || cell.t == null || cell.t == "z") return "";
    if (cell.w !== undefined) return cell.w;
    if (cell.t == "d" && !cell.z && o && o.dateNF) cell.z = o.dateNF;
    if (v == undefined) return safe_format_cell(cell, cell.v);
    return safe_format_cell(cell, v);
  }

  function sheet_to_workbook(sheet, opts) {
    const n = opts && opts.sheet ? opts.sheet : "Sheet1";
    const sheets = {};
    sheets[n] = sheet;
    return { SheetNames: [n], Sheets: sheets };
  }

  function sheet_add_aoa(_ws, data, opts) {
    const o = opts || {};
    let dense = _ws ? Array.isArray(_ws) : o.dense;
    if (DENSE != null && dense == null) dense = DENSE;
    const ws = _ws || (dense ? [] : {});
    let _R = 0;
    let _C = 0;
    if (ws && o.origin != null) {
      if (typeof o.origin === "number") _R = o.origin;
      else {
        const _origin =
          typeof o.origin === "string" ? decode_cell(o.origin) : o.origin;
        _R = _origin.r;
        _C = _origin.c;
      }
      if (!ws["!ref"]) ws["!ref"] = "A1:A1";
    }
    const range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
    if (ws["!ref"]) {
      const _range = safe_decode_range(ws["!ref"]);
      range.s.c = _range.s.c;
      range.s.r = _range.s.r;
      range.e.c = Math.max(range.e.c, _range.e.c);
      range.e.r = Math.max(range.e.r, _range.e.r);
      if (_R == -1) range.e.r = _R = _range.e.r + 1;
    }
    for (let R = 0; R != data.length; ++R) {
      if (!data[R]) continue;
      if (!Array.isArray(data[R]))
        throw new Error("aoa_to_sheet expects an array of arrays");
      for (let C = 0; C != data[R].length; ++C) {
        if (typeof data[R][C] === "undefined") continue;
        let cell = { v: data[R][C] };
        const __R = _R + R;
        const __C = _C + C;
        if (range.s.r > __R) range.s.r = __R;
        if (range.s.c > __C) range.s.c = __C;
        if (range.e.r < __R) range.e.r = __R;
        if (range.e.c < __C) range.e.c = __C;
        if (
          data[R][C] &&
          typeof data[R][C] === "object" &&
          !Array.isArray(data[R][C]) &&
          !(data[R][C] instanceof Date)
        )
          cell = data[R][C];
        else {
          if (Array.isArray(cell.v)) {
            cell.f = data[R][C][1];
            cell.v = cell.v[0];
          }
          if (cell.v === null) {
            if (cell.f) cell.t = "n";
            else if (!o.sheetStubs) continue;
            else cell.t = "z";
          } else if (typeof cell.v === "number") cell.t = "n";
          else if (typeof cell.v === "boolean") cell.t = "b";
          else if (cell.v instanceof Date) {
            cell.z = o.dateNF || SSF._table[14];
            if (o.cellDates) {
              cell.t = "d";
              cell.w = SSF.format(cell.z, datenum(cell.v));
            } else {
              cell.t = "n";
              cell.v = datenum(cell.v);
              cell.w = SSF.format(cell.z, cell.v);
            }
          } else cell.t = "s";
        }
        if (dense) {
          if (!ws[__R]) ws[__R] = [];
          if (ws[__R][__C] && ws[__R][__C].z) cell.z = ws[__R][__C].z;
          ws[__R][__C] = cell;
        } else {
          const cell_ref = encode_cell({ c: __C, r: __R });
          if (ws[cell_ref] && ws[cell_ref].z) cell.z = ws[cell_ref].z;
          ws[cell_ref] = cell;
        }
      }
    }
    if (range.s.c < 10000000) ws["!ref"] = encode_range(range);
    return ws;
  }
  function aoa_to_sheet(data, opts) {
    return sheet_add_aoa(null, data, opts);
  }

  /* [MS-XLSB] 2.5.168 */
  function parse_XLWideString(data) {
    const cchCharacters = data.read_shift(4);
    return cchCharacters === 0 ? "" : data.read_shift(cchCharacters, "dbcs");
  }
  function write_XLWideString(data, o) {
    let _null = false;
    if (o == null) {
      _null = true;
      o = new_buf(4 + 2 * data.length);
    }
    o.write_shift(4, data.length);
    if (data.length > 0) o.write_shift(0, data, "dbcs");
    return _null ? o.slice(0, o.l) : o;
  }

  /* [MS-XLSB] 2.5.91 */
  // function parse_LPWideString(data) {
  //	var cchCharacters = data.read_shift(2);
  //	return cchCharacters === 0 ? "" : data.read_shift(cchCharacters, "utf16le");
  // }

  /* [MS-XLSB] 2.5.143 */
  function parse_StrRun(data) {
    return { ich: data.read_shift(2), ifnt: data.read_shift(2) };
  }

  /* [MS-XLSB] 2.5.121 */
  function parse_RichStr(data, length) {
    const start = data.l;
    const flags = data.read_shift(1);
    const str = parse_XLWideString(data);
    const rgsStrRun = [];
    const z = { t: str, h: str };
    if ((flags & 1) !== 0) {
      /* fRichStr */
      /* TODO: formatted string */
      const dwSizeStrRun = data.read_shift(4);
      for (let i = 0; i != dwSizeStrRun; ++i)
        rgsStrRun.push(parse_StrRun(data));
      z.r = rgsStrRun;
    } else z.r = [{ ich: 0, ifnt: 0 }];
    // if((flags & 2) !== 0) { /* fExtStr */
    //	/* TODO: phonetic string */
    // }
    data.l = start + length;
    return z;
  }
  /* [MS-XLSB] 2.4.328 BrtCommentText (RichStr w/1 run) */
  const parse_BrtCommentText = parse_RichStr;

  /* [MS-XLSB] 2.5.9 */
  function parse_XLSBCell(data) {
    const col = data.read_shift(4);
    let iStyleRef = data.read_shift(2);
    iStyleRef += data.read_shift(1) << 16;
    data.l++; // var fPhShow = data.read_shift(1);
    return { c: col, iStyleRef };
  }
  function write_XLSBCell(cell, o) {
    if (o == null) o = new_buf(8);
    o.write_shift(-4, cell.c);
    o.write_shift(3, cell.iStyleRef || cell.s);
    o.write_shift(1, 0); /* fPhShow */
    return o;
  }

  /* [MS-XLSB] 2.5.21 */
  const parse_XLSBCodeName = parse_XLWideString;

  /* [MS-XLSB] 2.5.166 */
  function parse_XLNullableWideString(data) {
    const cchCharacters = data.read_shift(4);
    return cchCharacters === 0 || cchCharacters === 0xffffffff
      ? ""
      : data.read_shift(cchCharacters, "dbcs");
  }
  function write_XLNullableWideString(data, o) {
    let _null = false;
    if (o == null) {
      _null = true;
      o = new_buf(127);
    }
    o.write_shift(4, data.length > 0 ? data.length : 0xffffffff);
    if (data.length > 0) o.write_shift(0, data, "dbcs");
    return _null ? o.slice(0, o.l) : o;
  }

  /* [MS-XLSB] 2.5.165 */
  const parse_XLNameWideString = parse_XLWideString;
  // var write_XLNameWideString = write_XLWideString;

  /* [MS-XLSB] 2.5.114 */
  const parse_RelID = parse_XLNullableWideString;
  const write_RelID = write_XLNullableWideString;

  /* [MS-XLS] 2.5.217 ; [MS-XLSB] 2.5.122 */
  function parse_RkNumber(data) {
    const b = data.slice(data.l, data.l + 4);
    const fX100 = b[0] & 1;
    const fInt = b[0] & 2;
    data.l += 4;
    b[0] &= 0xfc; // b[0] &= ~3;
    const RK =
      fInt === 0
        ? __double([0, 0, 0, 0, b[0], b[1], b[2], b[3]], 0)
        : __readInt32LE(b, 0) >> 2;
    return fX100 ? RK / 100 : RK;
  }
  function write_RkNumber(data, o) {
    if (o == null) o = new_buf(4);
    let fX100 = 0;
    let fInt = 0;
    const d100 = data * 100;
    if (data == (data | 0) && data >= -(1 << 29) && data < 1 << 29) {
      fInt = 1;
    } else if (d100 == (d100 | 0) && d100 >= -(1 << 29) && d100 < 1 << 29) {
      fInt = 1;
      fX100 = 1;
    }
    if (fInt) o.write_shift(-4, ((fX100 ? d100 : data) << 2) + (fX100 + 2));
    else throw new Error(`unsupported RkNumber ${data}`); // TODO
  }

  /* [MS-XLSB] 2.5.117 RfX */
  function parse_RfX(data) {
    const cell = { s: {}, e: {} };
    cell.s.r = data.read_shift(4);
    cell.e.r = data.read_shift(4);
    cell.s.c = data.read_shift(4);
    cell.e.c = data.read_shift(4);
    return cell;
  }
  function write_RfX(r, o) {
    if (!o) o = new_buf(16);
    o.write_shift(4, r.s.r);
    o.write_shift(4, r.e.r);
    o.write_shift(4, r.s.c);
    o.write_shift(4, r.e.c);
    return o;
  }

  /* [MS-XLSB] 2.5.153 UncheckedRfX */
  const parse_UncheckedRfX = parse_RfX;
  const write_UncheckedRfX = write_RfX;

  /* [MS-XLSB] 2.5.155 UncheckedSqRfX */
  // function parse_UncheckedSqRfX(data) {
  //	var cnt = data.read_shift(4);
  //	var out = [];
  //	for(var i = 0; i < cnt; ++i) {
  //		var rng = parse_UncheckedRfX(data);
  //		out.push(encode_range(rng));
  //	}
  //	return out.join(",");
  // }
  // function write_UncheckedSqRfX(sqrfx) {
  //	var parts = sqrfx.split(/\s*,\s*/);
  //	var o = new_buf(4); o.write_shift(4, parts.length);
  //	var out = [o];
  //	parts.forEach(function(rng) {
  //		out.push(write_UncheckedRfX(safe_decode_range(rng)));
  //	});
  //	return bconcat(out);
  // }

  /* [MS-XLS] 2.5.342 ; [MS-XLSB] 2.5.171 */
  /* TODO: error checking, NaN and Infinity values are not valid Xnum */
  function parse_Xnum(data) {
    return data.read_shift(8, "f");
  }
  function write_Xnum(data, o) {
    return (o || new_buf(8)).write_shift(8, data, "f");
  }

  /* [MS-XLSB] 2.4.324 BrtColor */
  function parse_BrtColor(data) {
    const out = {};
    const d = data.read_shift(1);

    // var fValidRGB = d & 1;
    const xColorType = d >>> 1;

    const index = data.read_shift(1);
    const nTS = data.read_shift(2, "i");
    const bR = data.read_shift(1);
    const bG = data.read_shift(1);
    const bB = data.read_shift(1);
    data.l++; // var bAlpha = data.read_shift(1);

    switch (xColorType) {
      case 0:
        out.auto = 1;
        break;
      case 1:
        out.index = index;
        var icv = XLSIcv[index];
        /* automatic pseudo index 81 */
        if (icv) out.rgb = rgb2Hex(icv);
        break;
      case 2:
        /* if(!fValidRGB) throw new Error("invalid"); */
        out.rgb = rgb2Hex([bR, bG, bB]);
        break;
      case 3:
        out.theme = index;
        break;
    }
    if (nTS != 0) out.tint = nTS > 0 ? nTS / 32767 : nTS / 32768;

    return out;
  }

  /* [MS-XLSB] 2.5.52 */
  function parse_FontFlags(data) {
    const d = data.read_shift(1);
    data.l++;
    const out = {
      fBold: d & 0x01,
      fItalic: d & 0x02,
      fUnderline: d & 0x04,
      fStrikeout: d & 0x08,
      fOutline: d & 0x10,
      fShadow: d & 0x20,
      fCondense: d & 0x40,
      fExtend: d & 0x80
    };
    return out;
  }

  /* [MS-OLEDS] 2.3.1 and 2.3.2 */
  function parse_ClipboardFormatOrString(o, w) {
    // $FlowIgnore
    const ClipFmt = {
      2: "BITMAP",
      3: "METAFILEPICT",
      8: "DIB",
      14: "ENHMETAFILE"
    };
    const m = o.read_shift(4);
    switch (m) {
      case 0x00000000:
        return "";
      case 0xffffffff:
      case 0xfffffffe:
        return ClipFmt[o.read_shift(4)] || "";
    }
    if (m > 0x190) throw new Error(`Unsupported Clipboard: ${m.toString(16)}`);
    o.l -= 4;
    return o.read_shift(0, w == 1 ? "lpstr" : "lpwstr");
  }
  function parse_ClipboardFormatOrAnsiString(o) {
    return parse_ClipboardFormatOrString(o, 1);
  }
  function parse_ClipboardFormatOrUnicodeString(o) {
    return parse_ClipboardFormatOrString(o, 2);
  }

  /* [MS-OLEPS] 2.2 PropertyType */
  // var VT_EMPTY    = 0x0000;
  // var VT_NULL     = 0x0001;
  const VT_I2 = 0x0002;
  const VT_I4 = 0x0003;
  // var VT_R4       = 0x0004;
  // var VT_R8       = 0x0005;
  // var VT_CY       = 0x0006;
  // var VT_DATE     = 0x0007;
  // var VT_BSTR     = 0x0008;
  // var VT_ERROR    = 0x000A;
  const VT_BOOL = 0x000b;
  const VT_VARIANT = 0x000c;
  // var VT_DECIMAL  = 0x000E;
  // var VT_I1       = 0x0010;
  // var VT_UI1      = 0x0011;
  // var VT_UI2      = 0x0012;
  const VT_UI4 = 0x0013;
  // var VT_I8       = 0x0014;
  // var VT_UI8      = 0x0015;
  // var VT_INT      = 0x0016;
  // var VT_UINT     = 0x0017;
  const VT_LPSTR = 0x001e;
  // var VT_LPWSTR   = 0x001F;
  const VT_FILETIME = 0x0040;
  const VT_BLOB = 0x0041;
  // var VT_STREAM   = 0x0042;
  // var VT_STORAGE  = 0x0043;
  // var VT_STREAMED_Object  = 0x0044;
  // var VT_STORED_Object    = 0x0045;
  // var VT_BLOB_Object      = 0x0046;
  const VT_CF = 0x0047;
  // var VT_CLSID    = 0x0048;
  // var VT_VERSIONED_STREAM = 0x0049;
  const VT_VECTOR = 0x1000;
  // var VT_ARRAY    = 0x2000;

  const VT_STRING = 0x0050; // 2.3.3.1.11 VtString
  const VT_USTR = 0x0051; // 2.3.3.1.12 VtUnalignedString
  const VT_CUSTOM = [VT_STRING, VT_USTR];

  /* [MS-OSHARED] 2.3.3.2.2.1 Document Summary Information PIDDSI */
  const DocSummaryPIDDSI = {
    0x01: { n: "CodePage", t: VT_I2 },
    0x02: { n: "Category", t: VT_STRING },
    0x03: { n: "PresentationFormat", t: VT_STRING },
    0x04: { n: "ByteCount", t: VT_I4 },
    0x05: { n: "LineCount", t: VT_I4 },
    0x06: { n: "ParagraphCount", t: VT_I4 },
    0x07: { n: "SlideCount", t: VT_I4 },
    0x08: { n: "NoteCount", t: VT_I4 },
    0x09: { n: "HiddenCount", t: VT_I4 },
    0x0a: { n: "MultimediaClipCount", t: VT_I4 },
    0x0b: { n: "ScaleCrop", t: VT_BOOL },
    0x0c: { n: "HeadingPairs", t: VT_VECTOR | VT_VARIANT },
    0x0d: { n: "TitlesOfParts", t: VT_VECTOR | VT_LPSTR },
    0x0e: { n: "Manager", t: VT_STRING },
    0x0f: { n: "Company", t: VT_STRING },
    0x10: { n: "LinksUpToDate", t: VT_BOOL },
    0x11: { n: "CharacterCount", t: VT_I4 },
    0x13: { n: "SharedDoc", t: VT_BOOL },
    0x16: { n: "HyperlinksChanged", t: VT_BOOL },
    0x17: { n: "AppVersion", t: VT_I4, p: "version" },
    0x18: { n: "DigSig", t: VT_BLOB },
    0x1a: { n: "ContentType", t: VT_STRING },
    0x1b: { n: "ContentStatus", t: VT_STRING },
    0x1c: { n: "Language", t: VT_STRING },
    0x1d: { n: "Version", t: VT_STRING },
    0xff: {}
  };

  /* [MS-OSHARED] 2.3.3.2.1.1 Summary Information Property Set PIDSI */
  const SummaryPIDSI = {
    0x01: { n: "CodePage", t: VT_I2 },
    0x02: { n: "Title", t: VT_STRING },
    0x03: { n: "Subject", t: VT_STRING },
    0x04: { n: "Author", t: VT_STRING },
    0x05: { n: "Keywords", t: VT_STRING },
    0x06: { n: "Comments", t: VT_STRING },
    0x07: { n: "Template", t: VT_STRING },
    0x08: { n: "LastAuthor", t: VT_STRING },
    0x09: { n: "RevNumber", t: VT_STRING },
    0x0a: { n: "EditTime", t: VT_FILETIME },
    0x0b: { n: "LastPrinted", t: VT_FILETIME },
    0x0c: { n: "CreatedDate", t: VT_FILETIME },
    0x0d: { n: "ModifiedDate", t: VT_FILETIME },
    0x0e: { n: "PageCount", t: VT_I4 },
    0x0f: { n: "WordCount", t: VT_I4 },
    0x10: { n: "CharCount", t: VT_I4 },
    0x11: { n: "Thumbnail", t: VT_CF },
    0x12: { n: "Application", t: VT_STRING },
    0x13: { n: "DocSecurity", t: VT_I4 },
    0xff: {}
  };

  /* [MS-OLEPS] 2.18 */
  const SpecialProperties = {
    0x80000000: { n: "Locale", t: VT_UI4 },
    0x80000003: { n: "Behavior", t: VT_UI4 },
    0x72627262: {}
  };

  (function() {
    for (const y in SpecialProperties)
      if (Object.prototype.hasOwnProperty.call(SpecialProperties, y))
        DocSummaryPIDDSI[y] = SummaryPIDSI[y] = SpecialProperties[y];
  })();

  /* [MS-XLS] 2.4.63 Country/Region codes */
  const CountryEnum = {
    0x0001: "US", // United States
    0x0002: "CA", // Canada
    0x0003: "", // Latin America (except Brazil)
    0x0007: "RU", // Russia
    0x0014: "EG", // Egypt
    0x001e: "GR", // Greece
    0x001f: "NL", // Netherlands
    0x0020: "BE", // Belgium
    0x0021: "FR", // France
    0x0022: "ES", // Spain
    0x0024: "HU", // Hungary
    0x0027: "IT", // Italy
    0x0029: "CH", // Switzerland
    0x002b: "AT", // Austria
    0x002c: "GB", // United Kingdom
    0x002d: "DK", // Denmark
    0x002e: "SE", // Sweden
    0x002f: "NO", // Norway
    0x0030: "PL", // Poland
    0x0031: "DE", // Germany
    0x0034: "MX", // Mexico
    0x0037: "BR", // Brazil
    0x003d: "AU", // Australia
    0x0040: "NZ", // New Zealand
    0x0042: "TH", // Thailand
    0x0051: "JP", // Japan
    0x0052: "KR", // Korea
    0x0054: "VN", // Viet Nam
    0x0056: "CN", // China
    0x005a: "TR", // Turkey
    0x0069: "JS", // Ramastan
    0x00d5: "DZ", // Algeria
    0x00d8: "MA", // Morocco
    0x00da: "LY", // Libya
    0x015f: "PT", // Portugal
    0x0162: "IS", // Iceland
    0x0166: "FI", // Finland
    0x01a4: "CZ", // Czech Republic
    0x0376: "TW", // Taiwan
    0x03c1: "LB", // Lebanon
    0x03c2: "JO", // Jordan
    0x03c3: "SY", // Syria
    0x03c4: "IQ", // Iraq
    0x03c5: "KW", // Kuwait
    0x03c6: "SA", // Saudi Arabia
    0x03cb: "AE", // United Arab Emirates
    0x03cc: "IL", // Israel
    0x03ce: "QA", // Qatar
    0x03d5: "IR", // Iran
    0xffff: "US" // United States
  };

  /* [MS-XLS] 2.5.127 */
  const XLSFillPattern = [
    null,
    "solid",
    "mediumGray",
    "darkGray",
    "lightGray",
    "darkHorizontal",
    "darkVertical",
    "darkDown",
    "darkUp",
    "darkGrid",
    "darkTrellis",
    "lightHorizontal",
    "lightVertical",
    "lightDown",
    "lightUp",
    "lightGrid",
    "lightTrellis",
    "gray125",
    "gray0625"
  ];

  function rgbify(arr) {
    return arr.map(function(x) {
      return [(x >> 16) & 255, (x >> 8) & 255, x & 255];
    });
  }

  /* [MS-XLS] 2.5.161 */
  /* [MS-XLSB] 2.5.75 Icv */
  const _XLSIcv = rgbify([
    /* Color Constants */
    0x000000,
    0xffffff,
    0xff0000,
    0x00ff00,
    0x0000ff,
    0xffff00,
    0xff00ff,
    0x00ffff,

    /* Overridable Defaults */
    0x000000,
    0xffffff,
    0xff0000,
    0x00ff00,
    0x0000ff,
    0xffff00,
    0xff00ff,
    0x00ffff,

    0x800000,
    0x008000,
    0x000080,
    0x808000,
    0x800080,
    0x008080,
    0xc0c0c0,
    0x808080,
    0x9999ff,
    0x993366,
    0xffffcc,
    0xccffff,
    0x660066,
    0xff8080,
    0x0066cc,
    0xccccff,

    0x000080,
    0xff00ff,
    0xffff00,
    0x00ffff,
    0x800080,
    0x800000,
    0x008080,
    0x0000ff,
    0x00ccff,
    0xccffff,
    0xccffcc,
    0xffff99,
    0x99ccff,
    0xff99cc,
    0xcc99ff,
    0xffcc99,

    0x3366ff,
    0x33cccc,
    0x99cc00,
    0xffcc00,
    0xff9900,
    0xff6600,
    0x666699,
    0x969696,
    0x003366,
    0x339966,
    0x003300,
    0x333300,
    0x993300,
    0x993366,
    0x333399,
    0x333333,

    /* Other entries to appease BIFF8/12 */
    0xffffff /* 0x40 icvForeground ?? */,
    0x000000 /* 0x41 icvBackground ?? */,
    0x000000 /* 0x42 icvFrame ?? */,
    0x000000 /* 0x43 icv3D ?? */,
    0x000000 /* 0x44 icv3DText ?? */,
    0x000000 /* 0x45 icv3DHilite ?? */,
    0x000000 /* 0x46 icv3DShadow ?? */,
    0x000000 /* 0x47 icvHilite ?? */,
    0x000000 /* 0x48 icvCtlText ?? */,
    0x000000 /* 0x49 icvCtlScrl ?? */,
    0x000000 /* 0x4A icvCtlInv ?? */,
    0x000000 /* 0x4B icvCtlBody ?? */,
    0x000000 /* 0x4C icvCtlFrame ?? */,
    0x000000 /* 0x4D icvCtlFore ?? */,
    0x000000 /* 0x4E icvCtlBack ?? */,
    0x000000 /* 0x4F icvCtlNeutral */,
    0x000000 /* 0x50 icvInfoBk ?? */,
    0x000000 /* 0x51 icvInfoText ?? */
  ]);
  var XLSIcv = dup(_XLSIcv);

  /* [MS-XLSB] 2.5.97.2 */
  const BErr = {
    0x00: "#NULL!",
    0x07: "#DIV/0!",
    0x0f: "#VALUE!",
    0x17: "#REF!",
    0x1d: "#NAME?",
    0x24: "#NUM!",
    0x2a: "#N/A",
    0x2b: "#GETTING_DATA",
    0xff: "#WTF?"
  };
  const RBErr = evert_num(
    BErr
  ); /* Parts enumerated in OPC spec, MS-XLSB and MS-XLSX */
  /* 12.3 Part Summary <SpreadsheetML> */
  /* 14.2 Part Summary <DrawingML> */
  /* [MS-XLSX] 2.1 Part Enumerations ; [MS-XLSB] 2.1.7 Part Enumeration */
  const ct2type /* {[string]:string} */ = {
    /* Workbook */
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml":
      "workbooks",

    /* Worksheet */
    "application/vnd.ms-excel.binIndexWs": "TODO" /* Binary Index */,

    /* Macrosheet */
    "application/vnd.ms-excel.intlmacrosheet": "TODO",
    "application/vnd.ms-excel.binIndexMs": "TODO" /* Binary Index */,

    /* File Properties */
    "application/vnd.openxmlformats-package.core-properties+xml": "coreprops",
    "application/vnd.openxmlformats-officedocument.custom-properties+xml":
      "custprops",
    "application/vnd.openxmlformats-officedocument.extended-properties+xml":
      "extprops",

    /* Custom Data Properties */
    "application/vnd.openxmlformats-officedocument.customXmlProperties+xml":
      "TODO",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.customProperty":
      "TODO",

    /* PivotTable */
    "application/vnd.ms-excel.pivotTable": "TODO",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml":
      "TODO",

    /* Chart Objects */
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml": "TODO",

    /* Chart Colors */
    "application/vnd.ms-office.chartcolorstyle+xml": "TODO",

    /* Chart Style */
    "application/vnd.ms-office.chartstyle+xml": "TODO",

    /* Chart Advanced */
    "application/vnd.ms-office.chartex+xml": "TODO",

    /* Calculation Chain */
    "application/vnd.ms-excel.calcChain": "calcchains",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml":
      "calcchains",

    /* Printer Settings */
    "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings":
      "TODO",

    /* ActiveX */
    "application/vnd.ms-office.activeX": "TODO",
    "application/vnd.ms-office.activeX+xml": "TODO",

    /* Custom Toolbars */
    "application/vnd.ms-excel.attachedToolbars": "TODO",

    /* External Data Connections */
    "application/vnd.ms-excel.connections": "TODO",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml":
      "TODO",

    /* External Links */
    "application/vnd.ms-excel.externalLink": "links",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml":
      "links",

    /* Metadata */
    "application/vnd.ms-excel.sheetMetadata": "TODO",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml":
      "TODO",

    /* PivotCache */
    "application/vnd.ms-excel.pivotCacheDefinition": "TODO",
    "application/vnd.ms-excel.pivotCacheRecords": "TODO",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml":
      "TODO",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml":
      "TODO",

    /* Query Table */
    "application/vnd.ms-excel.queryTable": "TODO",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml":
      "TODO",

    /* Shared Workbook */
    "application/vnd.ms-excel.userNames": "TODO",
    "application/vnd.ms-excel.revisionHeaders": "TODO",
    "application/vnd.ms-excel.revisionLog": "TODO",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml":
      "TODO",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml":
      "TODO",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.userNames+xml":
      "TODO",

    /* Single Cell Table */
    "application/vnd.ms-excel.tableSingleCells": "TODO",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.tableSingleCells+xml":
      "TODO",

    /* Slicer */
    "application/vnd.ms-excel.slicer": "TODO",
    "application/vnd.ms-excel.slicerCache": "TODO",
    "application/vnd.ms-excel.slicer+xml": "TODO",
    "application/vnd.ms-excel.slicerCache+xml": "TODO",

    /* Sort Map */
    "application/vnd.ms-excel.wsSortMap": "TODO",

    /* Table */
    "application/vnd.ms-excel.table": "TODO",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml":
      "TODO",

    /* Themes */
    "application/vnd.openxmlformats-officedocument.theme+xml": "themes",

    /* Theme Override */
    "application/vnd.openxmlformats-officedocument.themeOverride+xml": "TODO",

    /* Timeline */
    "application/vnd.ms-excel.Timeline+xml": "TODO" /* verify */,
    "application/vnd.ms-excel.TimelineCache+xml": "TODO" /* verify */,

    /* VBA */
    "application/vnd.ms-office.vbaProject": "vba",
    "application/vnd.ms-office.vbaProjectSignature": "vba",

    /* Volatile Dependencies */
    "application/vnd.ms-office.volatileDependencies": "TODO",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.volatileDependencies+xml":
      "TODO",

    /* Control Properties */
    "application/vnd.ms-excel.controlproperties+xml": "TODO",

    /* Data Model */
    "application/vnd.openxmlformats-officedocument.model+data": "TODO",

    /* Survey */
    "application/vnd.ms-excel.Survey+xml": "TODO",

    /* Drawing */
    "application/vnd.openxmlformats-officedocument.drawing+xml": "drawings",
    "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml":
      "TODO",
    "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml":
      "TODO",
    "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml":
      "TODO",
    "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml":
      "TODO",
    "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml":
      "TODO",

    /* VML */
    "application/vnd.openxmlformats-officedocument.vmlDrawing": "TODO",

    "application/vnd.openxmlformats-package.relationships+xml": "rels",
    "application/vnd.openxmlformats-officedocument.oleObject": "TODO",

    /* Image */
    "image/png": "TODO",

    sheet: "js"
  };

  XMLNS.CT = "http://schemas.openxmlformats.org/package/2006/content-types";

  function new_ct() {
    return {
      workbooks: [],
      sheets: [],
      charts: [],
      dialogs: [],
      macros: [],
      rels: [],
      strs: [],
      comments: [],
      links: [],
      coreprops: [],
      extprops: [],
      custprops: [],
      themes: [],
      styles: [],
      calcchains: [],
      vba: [],
      drawings: [],
      TODO: [],
      xmlns: ""
    };
  }

  function parse_ct(data) {
    const ct = new_ct();
    if (!data || !data.match) return ct;
    const ctext = {};
    (data.match(tagregex) || []).forEach(function(x) {
      const y = parsexmltag(x);
      switch (y[0].replace(nsregex, "<")) {
        case "<?xml":
          break;
        case "<Types":
          ct.xmlns = y[`xmlns${(y[0].match(/<(\w+):/) || ["", ""])[1]}`];
          break;
        case "<Default":
          ctext[y.Extension] = y.ContentType;
          break;
        case "<Override":
          if (ct[ct2type[y.ContentType]] !== undefined)
            ct[ct2type[y.ContentType]].push(y.PartName);
          break;
      }
    });
    if (ct.xmlns !== XMLNS.CT)
      throw new Error(`Unknown Namespace: ${ct.xmlns}`);
    ct.calcchain = ct.calcchains.length > 0 ? ct.calcchains[0] : "";
    ct.sst = ct.strs.length > 0 ? ct.strs[0] : "";
    ct.style = ct.styles.length > 0 ? ct.styles[0] : "";
    ct.defaults = ctext;
    delete ct.calcchains;
    return ct;
  }

  /* 9.3 Relationships */
  const RELS = {
    WB:
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    SHEET:
      "http://sheetjs.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    HLINK:
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    VML:
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
    XPATH:
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
    XMISS:
      "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing",
    XLINK:
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
    VBA: "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
  };

  /* 9.3.3 Representing Relationships */
  function get_rels_path(file) {
    const n = file.lastIndexOf("/");
    return `${file.slice(0, n + 1)}_rels/${file.slice(n + 1)}.rels`;
  }

  function parse_rels(data, currentFilePath) {
    const rels = { "!id": {} };
    if (!data) return rels;
    if (currentFilePath.charAt(0) !== "/") {
      currentFilePath = `/${currentFilePath}`;
    }
    const hash = {};

    (data.match(tagregex) || []).forEach(function(x) {
      const y = parsexmltag(x);
      /* 9.3.2.2 OPC_Relationships */
      if (y[0] === "<Relationship") {
        const rel = {};
        rel.Type = y.Type;
        rel.Target = y.Target;
        rel.Id = y.Id;
        rel.TargetMode = y.TargetMode;
        const canonictarget =
          y.TargetMode === "External"
            ? y.Target
            : resolve_path(y.Target, currentFilePath);
        rels[canonictarget] = rel;
        hash[y.Id] = rel;
      }
    });
    rels["!id"] = hash;
    return rels;
  }

  XMLNS.RELS = "http://schemas.openxmlformats.org/package/2006/relationships";

  /* Open Document Format for Office Applications (OpenDocument) Version 1.2 */
  /* Part 3 Section 4 Manifest File */
  const CT_ODS = "application/vnd.oasis.opendocument.spreadsheet";
  function parse_manifest(d, opts) {
    const str = xlml_normalize(d);
    let Rn;
    let FEtag;
    while ((Rn = xlmlregex.exec(str)))
      switch (Rn[3]) {
        case "manifest":
          break; // 4.2 <manifest:manifest>
        case "file-entry": // 4.3 <manifest:file-entry>
          FEtag = parsexmltag(Rn[0], false);
          if (FEtag.path == "/" && FEtag.type !== CT_ODS)
            throw new Error("This OpenDocument is not a spreadsheet");
          break;
        case "encryption-data": // 4.4 <manifest:encryption-data>
        case "algorithm": // 4.5 <manifest:algorithm>
        case "start-key-generation": // 4.6 <manifest:start-key-generation>
        case "key-derivation": // 4.7 <manifest:key-derivation>
          throw new Error("Unsupported ODS Encryption");
        default:
          if (opts && opts.WTF) throw Rn;
      }
  }

  /* ECMA-376 Part II 11.1 Core Properties Part */
  /* [MS-OSHARED] 2.3.3.2.[1-2].1 (PIDSI/PIDDSI) */
  const CORE_PROPS = [
    ["cp:category", "Category"],
    ["cp:contentStatus", "ContentStatus"],
    ["cp:keywords", "Keywords"],
    ["cp:lastModifiedBy", "LastAuthor"],
    ["cp:lastPrinted", "LastPrinted"],
    ["cp:revision", "RevNumber"],
    ["cp:version", "Version"],
    ["dc:creator", "Author"],
    ["dc:description", "Comments"],
    ["dc:identifier", "Identifier"],
    ["dc:language", "Language"],
    ["dc:subject", "Subject"],
    ["dc:title", "Title"],
    ["dcterms:created", "CreatedDate", "date"],
    ["dcterms:modified", "ModifiedDate", "date"]
  ];

  XMLNS.CORE_PROPS =
    "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
  RELS.CORE_PROPS =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

  const CORE_PROPS_REGEX = (function() {
    const r = new Array(CORE_PROPS.length);
    for (let i = 0; i < CORE_PROPS.length; ++i) {
      const f = CORE_PROPS[i];
      const g = `(?:${f[0].slice(0, f[0].indexOf(":"))}:)${f[0].slice(
        f[0].indexOf(":") + 1
      )}`;
      r[i] = new RegExp(`<${g}[^>]*>([\\s\\S]*?)<\/${g}>`);
    }
    return r;
  })();

  function parse_core_props(data) {
    const p = {};
    data = utf8read(data);

    for (let i = 0; i < CORE_PROPS.length; ++i) {
      const f = CORE_PROPS[i];
      const cur = data.match(CORE_PROPS_REGEX[i]);
      if (cur != null && cur.length > 0) p[f[1]] = unescapexml(cur[1]);
      if (f[2] === "date" && p[f[1]]) p[f[1]] = parseDate(p[f[1]]);
    }

    return p;
  }

  /* 15.2.12.3 Extended File Properties Part */
  /* [MS-OSHARED] 2.3.3.2.[1-2].1 (PIDSI/PIDDSI) */
  const EXT_PROPS = [
    ["Application", "Application", "string"],
    ["AppVersion", "AppVersion", "string"],
    ["Company", "Company", "string"],
    ["DocSecurity", "DocSecurity", "string"],
    ["Manager", "Manager", "string"],
    ["HyperlinksChanged", "HyperlinksChanged", "bool"],
    ["SharedDoc", "SharedDoc", "bool"],
    ["LinksUpToDate", "LinksUpToDate", "bool"],
    ["ScaleCrop", "ScaleCrop", "bool"],
    ["HeadingPairs", "HeadingPairs", "raw"],
    ["TitlesOfParts", "TitlesOfParts", "raw"]
  ];

  XMLNS.EXT_PROPS =
    "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
  RELS.EXT_PROPS =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";

  const PseudoPropsPairs = [
    "Worksheets",
    "SheetNames",
    "NamedRanges",
    "DefinedNames",
    "Chartsheets",
    "ChartNames"
  ];
  function load_props_pairs(HP, TOP, props, opts) {
    let v = [];
    if (typeof HP === "string") v = parseVector(HP, opts);
    else
      for (let j = 0; j < HP.length; ++j)
        v = v.concat(
          HP[j].map(function(hp) {
            return { v: hp };
          })
        );
    const parts =
      typeof TOP === "string"
        ? parseVector(TOP, opts).map(function(x) {
            return x.v;
          })
        : TOP;
    let idx = 0;
    let len = 0;
    if (parts.length > 0)
      for (let i = 0; i !== v.length; i += 2) {
        len = +v[i + 1].v;
        switch (v[i].v) {
          case "Worksheets":
          case "工作表":
          case "Листы":
          case "أوراق العمل":
          case "ワークシート":
          case "גליונות עבודה":
          case "Arbeitsblätter":
          case "Çalışma Sayfaları":
          case "Feuilles de calcul":
          case "Fogli di lavoro":
          case "Folhas de cálculo":
          case "Planilhas":
          case "Regneark":
          case "Hojas de cálculo":
          case "Werkbladen":
            props.Worksheets = len;
            props.SheetNames = parts.slice(idx, idx + len);
            break;

          case "Named Ranges":
          case "Rangos con nombre":
          case "名前付き一覧":
          case "Benannte Bereiche":
          case "Navngivne områder":
            props.NamedRanges = len;
            props.DefinedNames = parts.slice(idx, idx + len);
            break;

          case "Charts":
          case "Diagramme":
            props.Chartsheets = len;
            props.ChartNames = parts.slice(idx, idx + len);
            break;
        }
        idx += len;
      }
  }

  function parse_ext_props(data, p, opts) {
    const q = {};
    if (!p) p = {};
    data = utf8read(data);

    EXT_PROPS.forEach(function(f) {
      const xml = (data.match(matchtag(f[0])) || [])[1];
      switch (f[2]) {
        case "string":
          if (xml) p[f[1]] = unescapexml(xml);
          break;
        case "bool":
          p[f[1]] = xml === "true";
          break;
        case "raw":
          var cur = data.match(
            new RegExp(`<${f[0]}[^>]*>([\\s\\S]*?)<\/${f[0]}>`)
          );
          if (cur && cur.length > 0) q[f[1]] = cur[1];
          break;
      }
    });

    if (q.HeadingPairs && q.TitlesOfParts)
      load_props_pairs(q.HeadingPairs, q.TitlesOfParts, p, opts);

    return p;
  }

  /* 15.2.12.2 Custom File Properties Part */
  XMLNS.CUST_PROPS =
    "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
  RELS.CUST_PROPS =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";

  const custregex = /<[^>]+>[^<]*/g;
  function parse_cust_props(data, opts) {
    const p = {};
    let name = "";
    const m = data.match(custregex);
    if (m)
      for (let i = 0; i != m.length; ++i) {
        const x = m[i];
        const y = parsexmltag(x);
        switch (y[0]) {
          case "<?xml":
            break;
          case "<Properties":
            break;
          case "<property":
            name = unescapexml(y.name);
            break;
          case "</property>":
            name = null;
            break;
          default:
            if (x.indexOf("<vt:") === 0) {
              const toks = x.split(">");
              const type = toks[0].slice(4);
              const text = toks[1];
              /* 22.4.2.32 (CT_Variant). Omit the binary types from 22.4 (Variant Types) */
              switch (type) {
                case "lpstr":
                case "bstr":
                case "lpwstr":
                  p[name] = unescapexml(text);
                  break;
                case "bool":
                  p[name] = parsexmlbool(text);
                  break;
                case "i1":
                case "i2":
                case "i4":
                case "i8":
                case "int":
                case "uint":
                  p[name] = parseInt(text, 10);
                  break;
                case "r4":
                case "r8":
                case "decimal":
                  p[name] = parseFloat(text);
                  break;
                case "filetime":
                case "date":
                  p[name] = parseDate(text);
                  break;
                case "cy":
                case "error":
                  p[name] = unescapexml(text);
                  break;
                default:
                  if (type.slice(-1) == "/") break;
                  if (opts.WTF && typeof console !== "undefined")
                    console.warn("Unexpected", x, type, toks);
              }
            } else if (x.slice(0, 2) === "</") {
              /* empty */
            } else if (opts.WTF) throw new Error(x);
        }
      }
    return p;
  }

  /* Common Name -> XLML Name */
  const XLMLDocPropsMap = {
    Title: "Title",
    Subject: "Subject",
    Author: "Author",
    Keywords: "Keywords",
    Comments: "Description",
    LastAuthor: "LastAuthor",
    RevNumber: "Revision",
    Application: "AppName",
    /* TotalTime: 'TotalTime', */
    LastPrinted: "LastPrinted",
    CreatedDate: "Created",
    ModifiedDate: "LastSaved",
    /* Pages */
    /* Words */
    /* Characters */
    Category: "Category",
    /* PresentationFormat */
    Manager: "Manager",
    Company: "Company",
    /* Guid */
    /* HyperlinkBase */
    /* Bytes */
    /* Lines */
    /* Paragraphs */
    /* CharactersWithSpaces */
    AppVersion: "Version",

    ContentStatus: "ContentStatus" /* NOTE: missing from schema */,
    Identifier: "Identifier" /* NOTE: missing from schema */,
    Language: "Language" /* NOTE: missing from schema */
  };
  const evert_XLMLDPM = evert(XLMLDocPropsMap);

  function xlml_set_prop(Props, tag, val) {
    tag = evert_XLMLDPM[tag] || tag;
    Props[tag] = val;
  }

  /* [MS-DTYP] 2.3.3 FILETIME */
  /* [MS-OLEDS] 2.1.3 FILETIME (Packet Version) */
  /* [MS-OLEPS] 2.8 FILETIME (Packet Version) */
  function parse_FILETIME(blob) {
    const dwLowDateTime = blob.read_shift(4);
    const dwHighDateTime = blob.read_shift(4);
    return new Date(
      ((dwHighDateTime / 1e7) * Math.pow(2, 32) +
        dwLowDateTime / 1e7 -
        11644473600) *
        1000
    )
      .toISOString()
      .replace(/\.000/, "");
  }
  function write_FILETIME(time) {
    const date = typeof time === "string" ? new Date(Date.parse(time)) : time;
    const t = date.getTime() / 1000 + 11644473600;
    let l = t % Math.pow(2, 32);
    let h = (t - l) / Math.pow(2, 32);
    l *= 1e7;
    h *= 1e7;
    const w = (l / Math.pow(2, 32)) | 0;
    if (w > 0) {
      l %= Math.pow(2, 32);
      h += w;
    }
    const o = new_buf(8);
    o.write_shift(4, l);
    o.write_shift(4, h);
    return o;
  }

  /* [MS-OSHARED] 2.3.3.1.4 Lpstr */
  function parse_lpstr(blob, type, pad) {
    const start = blob.l;
    const str = blob.read_shift(0, "lpstr-cp");
    if (pad) while ((blob.l - start) & 3) ++blob.l;
    return str;
  }

  /* [MS-OSHARED] 2.3.3.1.6 Lpwstr */
  function parse_lpwstr(blob, type, pad) {
    const str = blob.read_shift(0, "lpwstr");
    if (pad) blob.l += (4 - ((str.length + 1) & 3)) & 3;
    return str;
  }

  /* [MS-OSHARED] 2.3.3.1.11 VtString */
  /* [MS-OSHARED] 2.3.3.1.12 VtUnalignedString */
  function parse_VtStringBase(blob, stringType, pad) {
    if (stringType === 0x1f /* VT_LPWSTR */) return parse_lpwstr(blob);
    return parse_lpstr(blob, stringType, pad);
  }

  function parse_VtString(blob, t, pad) {
    return parse_VtStringBase(blob, t, pad === false ? 0 : 4);
  }
  function parse_VtUnalignedString(blob, t) {
    if (!t) throw new Error("VtUnalignedString must have positive length");
    return parse_VtStringBase(blob, t, 0);
  }

  /* [MS-OSHARED] 2.3.3.1.9 VtVecUnalignedLpstrValue */
  function parse_VtVecUnalignedLpstrValue(blob) {
    const length = blob.read_shift(4);
    const ret = [];
    for (let i = 0; i != length; ++i)
      ret[i] = blob.read_shift(0, "lpstr-cp").replace(chr0, "");
    return ret;
  }

  /* [MS-OSHARED] 2.3.3.1.10 VtVecUnalignedLpstr */
  function parse_VtVecUnalignedLpstr(blob) {
    return parse_VtVecUnalignedLpstrValue(blob);
  }

  /* [MS-OSHARED] 2.3.3.1.13 VtHeadingPair */
  function parse_VtHeadingPair(blob) {
    const headingString = parse_TypedPropertyValue(blob, VT_USTR);
    const headerParts = parse_TypedPropertyValue(blob, VT_I4);
    return [headingString, headerParts];
  }

  /* [MS-OSHARED] 2.3.3.1.14 VtVecHeadingPairValue */
  function parse_VtVecHeadingPairValue(blob) {
    const cElements = blob.read_shift(4);
    const out = [];
    for (let i = 0; i != cElements / 2; ++i)
      out.push(parse_VtHeadingPair(blob));
    return out;
  }

  /* [MS-OSHARED] 2.3.3.1.15 VtVecHeadingPair */
  function parse_VtVecHeadingPair(blob) {
    // NOTE: When invoked, wType & padding were already consumed
    return parse_VtVecHeadingPairValue(blob);
  }

  /* [MS-OLEPS] 2.18.1 Dictionary (uses 2.17, 2.16) */
  function parse_dictionary(blob, CodePage) {
    const cnt = blob.read_shift(4);
    const dict = {};
    for (let j = 0; j != cnt; ++j) {
      const pid = blob.read_shift(4);
      const len = blob.read_shift(4);
      dict[pid] = blob
        .read_shift(len, CodePage === 0x4b0 ? "utf16le" : "utf8")
        .replace(chr0, "")
        .replace(chr1, "!");
      if (CodePage === 0x4b0 && len % 2) blob.l += 2;
    }
    if (blob.l & 3) blob.l = (blob.l >> (2 + 1)) << 2;
    return dict;
  }

  /* [MS-OLEPS] 2.9 BLOB */
  function parse_BLOB(blob) {
    const size = blob.read_shift(4);
    const bytes = blob.slice(blob.l, blob.l + size);
    blob.l += size;
    if ((size & 3) > 0) blob.l += (4 - (size & 3)) & 3;
    return bytes;
  }

  /* [MS-OLEPS] 2.11 ClipboardData */
  function parse_ClipboardData(blob) {
    // TODO
    const o = {};
    o.Size = blob.read_shift(4);
    // o.Format = blob.read_shift(4);
    blob.l += o.Size + 3 - ((o.Size - 1) % 4);
    return o;
  }

  /* [MS-OLEPS] 2.15 TypedPropertyValue */
  function parse_TypedPropertyValue(blob, type, _opts) {
    const t = blob.read_shift(2);
    let ret;
    const opts = _opts || {};
    blob.l += 2;
    if (type !== VT_VARIANT)
      if (t !== type && VT_CUSTOM.indexOf(type) === -1)
        throw new Error(`Expected type ${type} saw ${t}`);
    switch (type === VT_VARIANT ? t : type) {
      case 0x02 /* VT_I2 */:
        ret = blob.read_shift(2, "i");
        if (!opts.raw) blob.l += 2;
        return ret;
      case 0x03 /* VT_I4 */:
        ret = blob.read_shift(4, "i");
        return ret;
      case 0x0b /* VT_BOOL */:
        return blob.read_shift(4) !== 0x0;
      case 0x13 /* VT_UI4 */:
        ret = blob.read_shift(4);
        return ret;
      case 0x1e /* VT_LPSTR */:
        return parse_lpstr(blob, t, 4).replace(chr0, "");
      case 0x1f /* VT_LPWSTR */:
        return parse_lpwstr(blob);
      case 0x40 /* VT_FILETIME */:
        return parse_FILETIME(blob);
      case 0x41 /* VT_BLOB */:
        return parse_BLOB(blob);
      case 0x47 /* VT_CF */:
        return parse_ClipboardData(blob);
      case 0x50 /* VT_STRING */:
        return parse_VtString(blob, t, !opts.raw).replace(chr0, "");
      case 0x51 /* VT_USTR */:
        return parse_VtUnalignedString(blob, t /* , 4 */).replace(chr0, "");
      case 0x100c /* VT_VECTOR|VT_VARIANT */:
        return parse_VtVecHeadingPair(blob);
      case 0x101e /* VT_LPSTR */:
        return parse_VtVecUnalignedLpstr(blob);
      default:
        throw new Error(`TypedPropertyValue unrecognized type ${type} ${t}`);
    }
  }
  function write_TypedPropertyValue(type, value) {
    const o = new_buf(4);
    let p = new_buf(4);
    o.write_shift(4, type == 0x50 ? 0x1f : type);
    switch (type) {
      case 0x03 /* VT_I4 */:
        p.write_shift(-4, value);
        break;
      case 0x05 /* VT_I4 */:
        p = new_buf(8);
        p.write_shift(8, value, "f");
        break;
      case 0x0b /* VT_BOOL */:
        p.write_shift(4, value ? 0x01 : 0x00);
        break;
      case 0x40 /* VT_FILETIME */:
        p = write_FILETIME(value);
        break;
      case 0x1f /* VT_LPWSTR */:
      case 0x50 /* VT_STRING */:
        p = new_buf(4 + 2 * (value.length + 1) + (value.length % 2 ? 0 : 2));
        p.write_shift(4, value.length + 1);
        p.write_shift(0, value, "dbcs");
        while (p.l != p.length) p.write_shift(1, 0);
        break;
      default:
        throw new Error(
          `TypedPropertyValue unrecognized type ${type} ${value}`
        );
    }
    return bconcat([o, p]);
  }

  /* [MS-OLEPS] 2.20 PropertySet */
  function parse_PropertySet(blob, PIDSI) {
    const start_addr = blob.l;
    const size = blob.read_shift(4);
    const NumProps = blob.read_shift(4);
    const Props = [];
    let i = 0;
    let CodePage = 0;
    let Dictionary = -1;
    let DictObj = {};
    for (i = 0; i != NumProps; ++i) {
      const PropID = blob.read_shift(4);
      const Offset = blob.read_shift(4);
      Props[i] = [PropID, Offset + start_addr];
    }
    Props.sort(function(x, y) {
      return x[1] - y[1];
    });
    const PropH = {};
    for (i = 0; i != NumProps; ++i) {
      if (blob.l !== Props[i][1]) {
        let fail = true;
        if (i > 0 && PIDSI)
          switch (PIDSI[Props[i - 1][0]].t) {
            case 0x02 /* VT_I2 */:
              if (blob.l + 2 === Props[i][1]) {
                blob.l += 2;
                fail = false;
              }
              break;
            case 0x50 /* VT_STRING */:
              if (blob.l <= Props[i][1]) {
                blob.l = Props[i][1];
                fail = false;
              }
              break;
            case 0x100c /* VT_VECTOR|VT_VARIANT */:
              if (blob.l <= Props[i][1]) {
                blob.l = Props[i][1];
                fail = false;
              }
              break;
          }
        if ((!PIDSI || i == 0) && blob.l <= Props[i][1]) {
          fail = false;
          blob.l = Props[i][1];
        }
        if (fail)
          throw new Error(
            `Read Error: Expected address ${Props[i][1]} at ${blob.l} :${i}`
          );
      }
      if (PIDSI) {
        const piddsi = PIDSI[Props[i][0]];
        PropH[piddsi.n] = parse_TypedPropertyValue(blob, piddsi.t, {
          raw: true
        });
        if (piddsi.p === "version")
          PropH[piddsi.n] = `${String(PropH[piddsi.n] >> 16)}.${`0000${String(
            PropH[piddsi.n] & 0xffff
          )}`.slice(-4)}`;
        if (piddsi.n == "CodePage")
          switch (PropH[piddsi.n]) {
            case 0:
              PropH[piddsi.n] = 1252;
            /* falls through */
            case 874:
            case 932:
            case 936:
            case 949:
            case 950:
            case 1250:
            case 1251:
            case 1253:
            case 1254:
            case 1255:
            case 1256:
            case 1257:
            case 1258:
            case 10000:
            case 1200:
            case 1201:
            case 1252:
            case 65000:
            case -536:
            case 65001:
            case -535:
              set_cp((CodePage = (PropH[piddsi.n] >>> 0) & 0xffff));
              break;
            default:
              throw new Error(`Unsupported CodePage: ${PropH[piddsi.n]}`);
          }
      } else if (Props[i][0] === 0x1) {
        CodePage = PropH.CodePage = parse_TypedPropertyValue(blob, VT_I2);
        set_cp(CodePage);
        if (Dictionary !== -1) {
          const oldpos = blob.l;
          blob.l = Props[Dictionary][1];
          DictObj = parse_dictionary(blob, CodePage);
          blob.l = oldpos;
        }
      } else if (Props[i][0] === 0) {
        if (CodePage === 0) {
          Dictionary = i;
          blob.l = Props[i + 1][1];
          continue;
        }
        DictObj = parse_dictionary(blob, CodePage);
      } else {
        const name = DictObj[Props[i][0]];
        var val;
        /* [MS-OSHARED] 2.3.3.2.3.1.2 + PROPVARIANT */
        switch (blob[blob.l]) {
          case 0x41 /* VT_BLOB */:
            blob.l += 4;
            val = parse_BLOB(blob);
            break;
          case 0x1e /* VT_LPSTR */:
            blob.l += 4;
            val = parse_VtString(blob, blob[blob.l - 4]).replace(
              /\u0000+$/,
              ""
            );
            break;
          case 0x1f /* VT_LPWSTR */:
            blob.l += 4;
            val = parse_VtString(blob, blob[blob.l - 4]).replace(
              /\u0000+$/,
              ""
            );
            break;
          case 0x03 /* VT_I4 */:
            blob.l += 4;
            val = blob.read_shift(4, "i");
            break;
          case 0x13 /* VT_UI4 */:
            blob.l += 4;
            val = blob.read_shift(4);
            break;
          case 0x05 /* VT_R8 */:
            blob.l += 4;
            val = blob.read_shift(8, "f");
            break;
          case 0x0b /* VT_BOOL */:
            blob.l += 4;
            val = parsebool(blob, 4);
            break;
          case 0x40 /* VT_FILETIME */:
            blob.l += 4;
            val = parseDate(parse_FILETIME(blob));
            break;
          default:
            throw new Error(`unparsed value: ${blob[blob.l]}`);
        }
        PropH[name] = val;
      }
    }
    blob.l = start_addr + size; /* step ahead to skip padding */
    return PropH;
  }
  const XLSPSSkip = [
    "CodePage",
    "Thumbnail",
    "_PID_LINKBASE",
    "_PID_HLINKS",
    "SystemIdentifier",
    "FMTID"
  ].concat(PseudoPropsPairs);
  function guess_property_type(val) {
    switch (typeof val) {
      case "boolean":
        return 0x0b;
      case "number":
        return (val | 0) == val ? 0x03 : 0x05;
      case "string":
        return 0x1f;
      case "object":
        if (val instanceof Date) return 0x40;
        break;
    }
    return -1;
  }
  function write_PropertySet(entries, RE, PIDSI) {
    const hdr = new_buf(8);
    const piao = [];
    const prop = [];
    let sz = 8;
    let i = 0;

    let pr = new_buf(8);
    let pio = new_buf(8);
    pr.write_shift(4, 0x0002);
    pr.write_shift(4, 0x04b0);
    pio.write_shift(4, 0x0001);
    prop.push(pr);
    piao.push(pio);
    sz += 8 + pr.length;

    if (!RE) {
      pio = new_buf(8);
      pio.write_shift(4, 0);
      piao.unshift(pio);

      const bufs = [new_buf(4)];
      bufs[0].write_shift(4, entries.length);
      for (i = 0; i < entries.length; ++i) {
        const value = entries[i][0];
        pr = new_buf(
          4 + 4 + 2 * (value.length + 1) + (value.length % 2 ? 0 : 2)
        );
        pr.write_shift(4, i + 2);
        pr.write_shift(4, value.length + 1);
        pr.write_shift(0, value, "dbcs");
        while (pr.l != pr.length) pr.write_shift(1, 0);
        bufs.push(pr);
      }
      pr = bconcat(bufs);
      prop.unshift(pr);
      sz += 8 + pr.length;
    }

    for (i = 0; i < entries.length; ++i) {
      if (RE && !RE[entries[i][0]]) continue;
      if (XLSPSSkip.indexOf(entries[i][0]) > -1) continue;
      if (entries[i][1] == null) continue;

      let val = entries[i][1];
      let idx = 0;
      if (RE) {
        idx = +RE[entries[i][0]];
        const pinfo = PIDSI[idx];
        if (pinfo.p == "version" && typeof val === "string") {
          const arr = val.split(".");
          val = (+arr[0] << 16) + (+arr[1] || 0);
        }
        pr = write_TypedPropertyValue(pinfo.t, val);
      } else {
        let T = guess_property_type(val);
        if (T == -1) {
          T = 0x1f;
          val = String(val);
        }
        pr = write_TypedPropertyValue(T, val);
      }
      prop.push(pr);

      pio = new_buf(8);
      pio.write_shift(4, !RE ? 2 + i : idx);
      piao.push(pio);

      sz += 8 + pr.length;
    }

    let w = 8 * (prop.length + 1);
    for (i = 0; i < prop.length; ++i) {
      piao[i].write_shift(4, w);
      w += prop[i].length;
    }
    hdr.write_shift(4, sz);
    hdr.write_shift(4, prop.length);
    return bconcat([hdr].concat(piao).concat(prop));
  }

  /* [MS-OLEPS] 2.21 PropertySetStream */
  function parse_PropertySetStream(file, PIDSI, clsid) {
    const blob = file.content;
    if (!blob) return {};
    prep_blob(blob, 0);

    let NumSets;
    let FMTID0;
    let FMTID1;
    let Offset0;
    let Offset1 = 0;
    blob.chk("feff", "Byte Order: ");

    /* var vers = */ blob.read_shift(2); // TODO: check version
    const SystemIdentifier = blob.read_shift(4);
    const CLSID = blob.read_shift(16);
    if (CLSID !== CFB.utils.consts.HEADER_CLSID && CLSID !== clsid)
      throw new Error(`Bad PropertySet CLSID ${CLSID}`);
    NumSets = blob.read_shift(4);
    if (NumSets !== 1 && NumSets !== 2)
      throw new Error(`Unrecognized #Sets: ${NumSets}`);
    FMTID0 = blob.read_shift(16);
    Offset0 = blob.read_shift(4);

    if (NumSets === 1 && Offset0 !== blob.l)
      throw new Error(`Length mismatch: ${Offset0} !== ${blob.l}`);
    else if (NumSets === 2) {
      FMTID1 = blob.read_shift(16);
      Offset1 = blob.read_shift(4);
    }
    const PSet0 = parse_PropertySet(blob, PIDSI);

    const rval = { SystemIdentifier };
    for (var y in PSet0) rval[y] = PSet0[y];
    // rval.blob = blob;
    rval.FMTID = FMTID0;
    // rval.PSet0 = PSet0;
    if (NumSets === 1) return rval;
    if (Offset1 - blob.l == 2) blob.l += 2;
    if (blob.l !== Offset1)
      throw new Error(`Length mismatch 2: ${blob.l} !== ${Offset1}`);
    let PSet1;
    try {
      PSet1 = parse_PropertySet(blob, null);
    } catch (e) {
      /* empty */
    }
    for (y in PSet1) rval[y] = PSet1[y];
    rval.FMTID = [FMTID0, FMTID1]; // TODO: verify FMTID0/1
    return rval;
  }

  function parsenoop2(blob, length) {
    blob.read_shift(length);
    return null;
  }
  function writezeroes(n, o) {
    if (!o) o = new_buf(n);
    for (let j = 0; j < n; ++j) o.write_shift(1, 0);
    return o;
  }

  function parslurp(blob, length, cb) {
    const arr = [];
    const target = blob.l + length;
    while (blob.l < target) arr.push(cb(blob, target - blob.l));
    if (target !== blob.l) throw new Error("Slurp error");
    return arr;
  }

  function parsebool(blob, length) {
    return blob.read_shift(length) === 0x1;
  }
  function writebool(v, o) {
    if (!o) o = new_buf(2);
    o.write_shift(2, +!!v);
    return o;
  }

  function parseuint16(blob) {
    return blob.read_shift(2, "u");
  }
  function writeuint16(v, o) {
    if (!o) o = new_buf(2);
    o.write_shift(2, v);
    return o;
  }
  function parseuint16a(blob, length) {
    return parslurp(blob, length, parseuint16);
  }

  /* --- 2.5 Structures --- */

  /* [MS-XLS] 2.5.10 Bes (boolean or error) */
  function parse_Bes(blob) {
    const v = blob.read_shift(1);
    const t = blob.read_shift(1);
    return t === 0x01 ? v : v === 0x01;
  }
  function write_Bes(v, t, o) {
    if (!o) o = new_buf(2);
    o.write_shift(1, +v);
    o.write_shift(1, t == "e" ? 1 : 0);
    return o;
  }

  /* [MS-XLS] 2.5.240 ShortXLUnicodeString */
  function parse_ShortXLUnicodeString(blob, length, opts) {
    const cch = blob.read_shift(opts && opts.biff >= 12 ? 2 : 1);
    let encoding = "sbcs-cont";
    const cp = current_codepage;
    if (opts && opts.biff >= 8) current_codepage = 1200;
    if (!opts || opts.biff == 8) {
      const fHighByte = blob.read_shift(1);
      if (fHighByte) {
        encoding = "dbcs-cont";
      }
    } else if (opts.biff == 12) {
      encoding = "wstr";
    }
    if (opts.biff >= 2 && opts.biff <= 5) encoding = "cpstr";
    const o = cch ? blob.read_shift(cch, encoding) : "";
    current_codepage = cp;
    return o;
  }

  /* 2.5.293 XLUnicodeRichExtendedString */
  function parse_XLUnicodeRichExtendedString(blob) {
    const cp = current_codepage;
    current_codepage = 1200;
    const cch = blob.read_shift(2);
    const flags = blob.read_shift(1);
    const /* fHighByte = flags & 0x1, */ fExtSt = flags & 0x4;
    const fRichSt = flags & 0x8;
    const width = 1 + (flags & 0x1); // 0x0 -> utf8, 0x1 -> dbcs
    let cRun = 0;
    let cbExtRst;
    const z = {};
    if (fRichSt) cRun = blob.read_shift(2);
    if (fExtSt) cbExtRst = blob.read_shift(4);
    const encoding = width == 2 ? "dbcs-cont" : "sbcs-cont";
    const msg = cch === 0 ? "" : blob.read_shift(cch, encoding);
    if (fRichSt) blob.l += 4 * cRun; // TODO: parse this
    if (fExtSt) blob.l += cbExtRst; // TODO: parse this
    z.t = msg;
    if (!fRichSt) {
      z.raw = `<t>${z.t}</t>`;
      z.r = z.t;
    }
    current_codepage = cp;
    return z;
  }
  function write_XLUnicodeRichExtendedString(xlstr) {
    const str = xlstr.t || "";
    const nfmts = 1;

    const hdr = new_buf(3 + (nfmts > 1 ? 2 : 0));
    hdr.write_shift(2, str.length);
    hdr.write_shift(1, (nfmts > 1 ? 0x08 : 0x00) | 0x01);
    if (nfmts > 1) hdr.write_shift(2, nfmts);

    const otext = new_buf(2 * str.length);
    otext.write_shift(2 * str.length, str, "utf16le");

    const out = [hdr, otext];

    return bconcat(out);
  }

  /* 2.5.296 XLUnicodeStringNoCch */
  function parse_XLUnicodeStringNoCch(blob, cch, opts) {
    let retval;
    if (opts) {
      if (opts.biff >= 2 && opts.biff <= 5)
        return blob.read_shift(cch, "cpstr");
      if (opts.biff >= 12) return blob.read_shift(cch, "dbcs-cont");
    }
    const fHighByte = blob.read_shift(1);
    if (fHighByte === 0) {
      retval = blob.read_shift(cch, "sbcs-cont");
    } else {
      retval = blob.read_shift(cch, "dbcs-cont");
    }
    return retval;
  }

  /* 2.5.294 XLUnicodeString */
  function parse_XLUnicodeString(blob, length, opts) {
    const cch = blob.read_shift(opts && opts.biff == 2 ? 1 : 2);
    if (cch === 0) {
      blob.l++;
      return "";
    }
    return parse_XLUnicodeStringNoCch(blob, cch, opts);
  }
  /* BIFF5 override */
  function parse_XLUnicodeString2(blob, length, opts) {
    if (opts.biff > 5) return parse_XLUnicodeString(blob, length, opts);
    const cch = blob.read_shift(1);
    if (cch === 0) {
      blob.l++;
      return "";
    }
    return blob.read_shift(
      cch,
      opts.biff <= 4 || !blob.lens ? "cpstr" : "sbcs-cont"
    );
  }
  /* TODO: BIFF5 and lower, codepage awareness */
  function write_XLUnicodeString(str, opts, o) {
    if (!o) o = new_buf(3 + 2 * str.length);
    o.write_shift(2, str.length);
    o.write_shift(1, 1);
    o.write_shift(31, str, "utf16le");
    return o;
  }

  /* [MS-XLS] 2.5.61 ControlInfo */
  function parse_ControlInfo(blob) {
    const flags = blob.read_shift(1);
    blob.l++;
    const accel = blob.read_shift(2);
    blob.l += 2;
    return [flags, accel];
  }

  /* [MS-OSHARED] 2.3.7.6 URLMoniker TODO: flags */
  function parse_URLMoniker(blob) {
    const len = blob.read_shift(4);
    const start = blob.l;
    let extra = false;
    if (len > 24) {
      /* look ahead */
      blob.l += len - 24;
      if (blob.read_shift(16) === "795881f43b1d7f48af2c825dc4852763")
        extra = true;
      blob.l = start;
    }
    const url = blob
      .read_shift((extra ? len - 24 : len) >> 1, "utf16le")
      .replace(chr0, "");
    if (extra) blob.l += 24;
    return url;
  }

  /* [MS-OSHARED] 2.3.7.8 FileMoniker TODO: all fields */
  function parse_FileMoniker(blob) {
    blob.l += 2; // var cAnti = blob.read_shift(2);
    const ansiPath = blob.read_shift(0, "lpstr-ansi");
    blob.l += 2; // var endServer = blob.read_shift(2);
    if (blob.read_shift(2) != 0xdead) throw new Error("Bad FileMoniker");
    const sz = blob.read_shift(4);
    if (sz === 0) return ansiPath.replace(/\\/g, "/");
    const bytes = blob.read_shift(4);
    if (blob.read_shift(2) != 3) throw new Error("Bad FileMoniker");
    const unicodePath = blob
      .read_shift(bytes >> 1, "utf16le")
      .replace(chr0, "");
    return unicodePath;
  }

  /* [MS-OSHARED] 2.3.7.2 HyperlinkMoniker TODO: all the monikers */
  function parse_HyperlinkMoniker(blob, length) {
    const clsid = blob.read_shift(16);
    length -= 16;
    switch (clsid) {
      case "e0c9ea79f9bace118c8200aa004ba90b":
        return parse_URLMoniker(blob, length);
      case "0303000000000000c000000000000046":
        return parse_FileMoniker(blob, length);
      default:
        throw new Error(`Unsupported Moniker ${clsid}`);
    }
  }

  /* [MS-OSHARED] 2.3.7.9 HyperlinkString */
  function parse_HyperlinkString(blob) {
    const len = blob.read_shift(4);
    const o = len > 0 ? blob.read_shift(len, "utf16le").replace(chr0, "") : "";
    return o;
  }

  /* [MS-OSHARED] 2.3.7.1 Hyperlink Object */
  function parse_Hyperlink(blob, length) {
    const end = blob.l + length;
    const sVer = blob.read_shift(4);
    if (sVer !== 2) throw new Error(`Unrecognized streamVersion: ${sVer}`);
    const flags = blob.read_shift(2);
    blob.l += 2;
    let displayName;
    let targetFrameName;
    let moniker;
    let oleMoniker;
    let Loc = "";
    let guid;
    let fileTime;
    if (flags & 0x0010) displayName = parse_HyperlinkString(blob, end - blob.l);
    if (flags & 0x0080)
      targetFrameName = parse_HyperlinkString(blob, end - blob.l);
    if ((flags & 0x0101) === 0x0101)
      moniker = parse_HyperlinkString(blob, end - blob.l);
    if ((flags & 0x0101) === 0x0001)
      oleMoniker = parse_HyperlinkMoniker(blob, end - blob.l);
    if (flags & 0x0008) Loc = parse_HyperlinkString(blob, end - blob.l);
    if (flags & 0x0020) guid = blob.read_shift(16);
    if (flags & 0x0040) fileTime = parse_FILETIME(blob /* , 8 */);
    blob.l = end;
    let target = targetFrameName || moniker || oleMoniker || "";
    if (target && Loc) target += `#${Loc}`;
    if (!target) target = `#${Loc}`;
    const out = { Target: target };
    if (guid) out.guid = guid;
    if (fileTime) out.time = fileTime;
    if (displayName) out.Tooltip = displayName;
    return out;
  }
  function write_Hyperlink(hl) {
    const out = new_buf(512);
    let i = 0;
    let { Target } = hl;
    let F = Target.indexOf("#") > -1 ? 0x1f : 0x17;
    switch (Target.charAt(0)) {
      case "#":
        F = 0x1c;
        break;
      case ".":
        F &= ~2;
        break;
    }
    out.write_shift(4, 2);
    out.write_shift(4, F);
    let data = [8, 6815827, 6619237, 4849780, 83];
    for (i = 0; i < data.length; ++i) out.write_shift(4, data[i]);
    if (F == 0x1c) {
      Target = Target.slice(1);
      out.write_shift(4, Target.length + 1);
      for (i = 0; i < Target.length; ++i)
        out.write_shift(2, Target.charCodeAt(i));
      out.write_shift(2, 0);
    } else if (F & 0x02) {
      data = "e0 c9 ea 79 f9 ba ce 11 8c 82 00 aa 00 4b a9 0b".split(" ");
      for (i = 0; i < data.length; ++i)
        out.write_shift(1, parseInt(data[i], 16));
      out.write_shift(4, 2 * (Target.length + 1));
      for (i = 0; i < Target.length; ++i)
        out.write_shift(2, Target.charCodeAt(i));
      out.write_shift(2, 0);
    } else {
      data = "03 03 00 00 00 00 00 00 c0 00 00 00 00 00 00 46".split(" ");
      for (i = 0; i < data.length; ++i)
        out.write_shift(1, parseInt(data[i], 16));
      let P = 0;
      while (
        Target.slice(P * 3, P * 3 + 3) == "../" ||
        Target.slice(P * 3, P * 3 + 3) == "..\\"
      )
        ++P;
      out.write_shift(2, P);
      out.write_shift(4, Target.length + 1);
      for (i = 0; i < Target.length; ++i)
        out.write_shift(1, Target.charCodeAt(i) & 0xff);
      out.write_shift(1, 0);
      out.write_shift(2, 0xffff);
      out.write_shift(2, 0xdead);
      for (i = 0; i < 6; ++i) out.write_shift(4, 0);
    }
    return out.slice(0, out.l);
  }

  /* 2.5.178 LongRGBA */
  function parse_LongRGBA(blob) {
    const r = blob.read_shift(1);
    const g = blob.read_shift(1);
    const b = blob.read_shift(1);
    const a = blob.read_shift(1);
    return [r, g, b, a];
  }

  /* 2.5.177 LongRGB */
  function parse_LongRGB(blob, length) {
    const x = parse_LongRGBA(blob, length);
    x[3] = 0;
    return x;
  }

  /* [MS-XLS] 2.5.19 */
  function parse_XLSCell(blob) {
    const rw = blob.read_shift(2); // 0-indexed
    const col = blob.read_shift(2);
    const ixfe = blob.read_shift(2);
    return { r: rw, c: col, ixfe };
  }
  function write_XLSCell(R, C, ixfe, o) {
    if (!o) o = new_buf(6);
    o.write_shift(2, R);
    o.write_shift(2, C);
    o.write_shift(2, ixfe || 0);
    return o;
  }

  /* [MS-XLS] 2.5.134 */
  function parse_frtHeader(blob) {
    const rt = blob.read_shift(2);
    const flags = blob.read_shift(2); // TODO: parse these flags
    blob.l += 8;
    return { type: rt, flags };
  }

  function parse_OptXLUnicodeString(blob, length, opts) {
    return length === 0 ? "" : parse_XLUnicodeString2(blob, length, opts);
  }

  /* [MS-XLS] 2.5.344 */
  function parse_XTI(blob, length, opts) {
    const w = opts.biff > 8 ? 4 : 2;
    const iSupBook = blob.read_shift(w);
    const itabFirst = blob.read_shift(w, "i");
    const itabLast = blob.read_shift(w, "i");
    return [iSupBook, itabFirst, itabLast];
  }

  /* [MS-XLS] 2.5.218 */
  function parse_RkRec(blob) {
    const ixfe = blob.read_shift(2);
    const RK = parse_RkNumber(blob);
    return [ixfe, RK];
  }

  /* [MS-XLS] 2.5.1 */
  function parse_AddinUdf(blob, length, opts) {
    blob.l += 4;
    length -= 4;
    let l = blob.l + length;
    const udfName = parse_ShortXLUnicodeString(blob, length, opts);
    const cb = blob.read_shift(2);
    l -= blob.l;
    if (cb !== l)
      throw new Error(`Malformed AddinUdf: padding = ${l} != ${cb}`);
    blob.l += cb;
    return udfName;
  }

  /* [MS-XLS] 2.5.209 TODO: Check sizes */
  function parse_Ref8U(blob) {
    const rwFirst = blob.read_shift(2);
    const rwLast = blob.read_shift(2);
    const colFirst = blob.read_shift(2);
    const colLast = blob.read_shift(2);
    return { s: { c: colFirst, r: rwFirst }, e: { c: colLast, r: rwLast } };
  }
  function write_Ref8U(r, o) {
    if (!o) o = new_buf(8);
    o.write_shift(2, r.s.r);
    o.write_shift(2, r.e.r);
    o.write_shift(2, r.s.c);
    o.write_shift(2, r.e.c);
    return o;
  }

  /* [MS-XLS] 2.5.211 */
  function parse_RefU(blob) {
    const rwFirst = blob.read_shift(2);
    const rwLast = blob.read_shift(2);
    const colFirst = blob.read_shift(1);
    const colLast = blob.read_shift(1);
    return { s: { c: colFirst, r: rwFirst }, e: { c: colLast, r: rwLast } };
  }

  /* [MS-XLS] 2.5.207 */
  const parse_Ref = parse_RefU;

  /* [MS-XLS] 2.5.143 */
  function parse_FtCmo(blob) {
    blob.l += 4;
    const ot = blob.read_shift(2);
    const id = blob.read_shift(2);
    const flags = blob.read_shift(2);
    blob.l += 12;
    return [id, ot, flags];
  }

  /* [MS-XLS] 2.5.149 */
  function parse_FtNts(blob) {
    const out = {};
    blob.l += 4;
    blob.l += 16; // GUID TODO
    out.fSharedNote = blob.read_shift(2);
    blob.l += 4;
    return out;
  }

  /* [MS-XLS] 2.5.142 */
  function parse_FtCf(blob) {
    const out = {};
    blob.l += 4;
    blob.cf = blob.read_shift(2);
    return out;
  }

  /* [MS-XLS] 2.5.140 - 2.5.154 and friends */
  function parse_FtSkip(blob) {
    blob.l += 2;
    blob.l += blob.read_shift(2);
  }
  const FtTab = {
    0x00: parse_FtSkip /* FtEnd */,
    0x04: parse_FtSkip /* FtMacro */,
    0x05: parse_FtSkip /* FtButton */,
    0x06: parse_FtSkip /* FtGmo */,
    0x07: parse_FtCf /* FtCf */,
    0x08: parse_FtSkip /* FtPioGrbit */,
    0x09: parse_FtSkip /* FtPictFmla */,
    0x0a: parse_FtSkip /* FtCbls */,
    0x0b: parse_FtSkip /* FtRbo */,
    0x0c: parse_FtSkip /* FtSbs */,
    0x0d: parse_FtNts /* FtNts */,
    0x0e: parse_FtSkip /* FtSbsFmla */,
    0x0f: parse_FtSkip /* FtGboData */,
    0x10: parse_FtSkip /* FtEdoData */,
    0x11: parse_FtSkip /* FtRboData */,
    0x12: parse_FtSkip /* FtCblsData */,
    0x13: parse_FtSkip /* FtLbsData */,
    0x14: parse_FtSkip /* FtCblsFmla */,
    0x15: parse_FtCmo
  };
  function parse_FtArray(blob, length) {
    const tgt = blob.l + length;
    const fts = [];
    while (blob.l < tgt) {
      const ft = blob.read_shift(2);
      blob.l -= 2;
      try {
        fts.push(FtTab[ft](blob, tgt - blob.l));
      } catch (e) {
        blob.l = tgt;
        return fts;
      }
    }
    if (blob.l != tgt) blob.l = tgt; // throw new Error("bad Object Ft-sequence");
    return fts;
  }

  /* --- 2.4 Records --- */

  /* [MS-XLS] 2.4.21 */
  function parse_BOF(blob, length) {
    const o = { BIFFVer: 0, dt: 0 };
    o.BIFFVer = blob.read_shift(2);
    length -= 2;
    if (length >= 2) {
      o.dt = blob.read_shift(2);
      blob.l -= 2;
    }
    switch (o.BIFFVer) {
      case 0x0600: /* BIFF8 */
      case 0x0500: /* BIFF5 */
      case 0x0400: /* BIFF4 */
      case 0x0300: /* BIFF3 */
      case 0x0200: /* BIFF2 */
      case 0x0002:
      case 0x0007 /* BIFF2 */:
        break;
      default:
        if (length > 6) throw new Error(`Unexpected BIFF Ver ${o.BIFFVer}`);
    }

    blob.read_shift(length);
    return o;
  }
  function write_BOF(wb, t, o) {
    let h = 0x0600;
    let w = 16;
    switch (o.bookType) {
      case "biff8":
        break;
      case "biff5":
        h = 0x0500;
        w = 8;
        break;
      case "biff4":
        h = 0x0004;
        w = 6;
        break;
      case "biff3":
        h = 0x0003;
        w = 6;
        break;
      case "biff2":
        h = 0x0002;
        w = 4;
        break;
      case "xla":
        break;
      default:
        throw new Error("unsupported BIFF version");
    }
    const out = new_buf(w);
    out.write_shift(2, h);
    out.write_shift(2, t);
    if (w > 4) out.write_shift(2, 0x7262);
    if (w > 6) out.write_shift(2, 0x07cd);
    if (w > 8) {
      out.write_shift(2, 0xc009);
      out.write_shift(2, 0x0001);
      out.write_shift(2, 0x0706);
      out.write_shift(2, 0x0000);
    }
    return out;
  }

  /* [MS-XLS] 2.4.146 */
  function parse_InterfaceHdr(blob, length) {
    if (length === 0) return 0x04b0;
    if (blob.read_shift(2) !== 0x04b0) {
      /* empty */
    }
    return 0x04b0;
  }

  /* [MS-XLS] 2.4.349 */
  function parse_WriteAccess(blob, length, opts) {
    if (opts.enc) {
      blob.l += length;
      return "";
    }
    const { l } = blob;
    // TODO: make sure XLUnicodeString doesnt overrun
    const UserName = parse_XLUnicodeString2(blob, 0, opts);
    blob.read_shift(length + l - blob.l);
    return UserName;
  }
  function write_WriteAccess(s, opts) {
    const b8 = !opts || opts.biff == 8;
    const o = new_buf(b8 ? 112 : 54);
    o.write_shift(opts.biff == 8 ? 2 : 1, 7);
    if (b8) o.write_shift(1, 0);
    o.write_shift(4, 0x33336853);
    o.write_shift(4, 0x00534a74 | (b8 ? 0 : 0x20000000));
    while (o.l < o.length) o.write_shift(1, b8 ? 0 : 32);
    return o;
  }

  /* [MS-XLS] 2.4.351 */
  function parse_WsBool(blob, length, opts) {
    const flags =
      (opts && opts.biff == 8) || length == 2
        ? blob.read_shift(2)
        : ((blob.l += length), 0);
    return { fDialog: flags & 0x10 };
  }

  /* [MS-XLS] 2.4.28 */
  function parse_BoundSheet8(blob, length, opts) {
    const pos = blob.read_shift(4);
    const hidden = blob.read_shift(1) & 0x03;
    let dt = blob.read_shift(1);
    switch (dt) {
      case 0:
        dt = "Worksheet";
        break;
      case 1:
        dt = "Macrosheet";
        break;
      case 2:
        dt = "Chartsheet";
        break;
      case 6:
        dt = "VBAModule";
        break;
    }
    let name = parse_ShortXLUnicodeString(blob, 0, opts);
    if (name.length === 0) name = "Sheet1";
    return { pos, hs: hidden, dt, name };
  }
  function write_BoundSheet8(data, opts) {
    const w = !opts || opts.biff >= 8 ? 2 : 1;
    const o = new_buf(8 + w * data.name.length);
    o.write_shift(4, data.pos);
    o.write_shift(1, data.hs || 0);
    o.write_shift(1, data.dt);
    o.write_shift(1, data.name.length);
    if (opts.biff >= 8) o.write_shift(1, 1);
    o.write_shift(
      w * data.name.length,
      data.name,
      opts.biff < 8 ? "sbcs" : "utf16le"
    );
    const out = o.slice(0, o.l);
    out.l = o.l;
    return out;
  }

  /* [MS-XLS] 2.4.265 TODO */
  function parse_SST(blob, length) {
    const end = blob.l + length;
    const cnt = blob.read_shift(4);
    const ucnt = blob.read_shift(4);
    const strs = [];
    for (let i = 0; i != ucnt && blob.l < end; ++i) {
      strs.push(parse_XLUnicodeRichExtendedString(blob));
    }
    strs.Count = cnt;
    strs.Unique = ucnt;
    return strs;
  }
  function write_SST(sst, opts) {
    const header = new_buf(8);
    header.write_shift(4, sst.Count);
    header.write_shift(4, sst.Unique);
    const strs = [];
    for (let j = 0; j < sst.length; ++j)
      strs[j] = write_XLUnicodeRichExtendedString(sst[j], opts);
    const o = bconcat([header].concat(strs));
    o.parts = [header.length].concat(
      strs.map(function(str) {
        return str.length;
      })
    );
    return o;
  }

  /* [MS-XLS] 2.4.107 */
  function parse_ExtSST(blob, length) {
    const extsst = {};
    extsst.dsst = blob.read_shift(2);
    blob.l += length - 2;
    return extsst;
  }

  /* [MS-XLS] 2.4.221 TODO: check BIFF2-4 */
  function parse_Row(blob) {
    const z = {};
    z.r = blob.read_shift(2);
    z.c = blob.read_shift(2);
    z.cnt = blob.read_shift(2) - z.c;
    const miyRw = blob.read_shift(2);
    blob.l += 4; // reserved(2), unused(2)
    const flags = blob.read_shift(1); // various flags
    blob.l += 3; // reserved(8), ixfe(12), flags(4)
    if (flags & 0x07) z.level = flags & 0x07;
    // collapsed: flags & 0x10
    if (flags & 0x20) z.hidden = true;
    if (flags & 0x40) z.hpt = miyRw / 20;
    return z;
  }

  /* [MS-XLS] 2.4.125 */
  function parse_ForceFullCalculation(blob) {
    const header = parse_frtHeader(blob);
    if (header.type != 0x08a3)
      throw new Error(`Invalid Future Record ${header.type}`);
    const fullcalc = blob.read_shift(4);
    return fullcalc !== 0x0;
  }

  /* [MS-XLS] 2.4.215 rt */
  function parse_RecalcId(blob) {
    blob.read_shift(2);
    return blob.read_shift(4);
  }

  /* [MS-XLS] 2.4.87 */
  function parse_DefaultRowHeight(blob, length, opts) {
    let f = 0;
    if (!(opts && opts.biff == 2)) {
      f = blob.read_shift(2);
    }
    let miyRw = blob.read_shift(2);
    if (opts && opts.biff == 2) {
      f = 1 - (miyRw >> 15);
      miyRw &= 0x7fff;
    }
    const fl = {
      Unsynced: f & 1,
      DyZero: (f & 2) >> 1,
      ExAsc: (f & 4) >> 2,
      ExDsc: (f & 8) >> 3
    };
    return [fl, miyRw];
  }

  /* [MS-XLS] 2.4.345 TODO */
  function parse_Window1(blob) {
    const xWn = blob.read_shift(2);
    const yWn = blob.read_shift(2);
    const dxWn = blob.read_shift(2);
    const dyWn = blob.read_shift(2);
    const flags = blob.read_shift(2);
    const iTabCur = blob.read_shift(2);
    const iTabFirst = blob.read_shift(2);
    const ctabSel = blob.read_shift(2);
    const wTabRatio = blob.read_shift(2);
    return {
      Pos: [xWn, yWn],
      Dim: [dxWn, dyWn],
      Flags: flags,
      CurTab: iTabCur,
      FirstTab: iTabFirst,
      Selected: ctabSel,
      TabRatio: wTabRatio
    };
  }
  function write_Window1() {
    const o = new_buf(18);
    o.write_shift(2, 0);
    o.write_shift(2, 0);
    o.write_shift(2, 0x7260);
    o.write_shift(2, 0x44c0);
    o.write_shift(2, 0x38);
    o.write_shift(2, 0);
    o.write_shift(2, 0);
    o.write_shift(2, 1);
    o.write_shift(2, 0x01f4);
    return o;
  }
  /* [MS-XLS] 2.4.346 TODO */
  function parse_Window2(blob, length, opts) {
    if (opts && opts.biff >= 2 && opts.biff < 5) return {};
    const f = blob.read_shift(2);
    return { RTL: f & 0x40 };
  }
  function write_Window2(view) {
    const o = new_buf(18);
    let f = 0x6b6;
    if (view && view.RTL) f |= 0x40;
    o.write_shift(2, f);
    o.write_shift(4, 0);
    o.write_shift(4, 64);
    o.write_shift(4, 0);
    o.write_shift(4, 0);
    return o;
  }

  /* [MS-XLS] 2.4.189 TODO */
  function parse_Pane(/* blob, length, opts */) {}

  /* [MS-XLS] 2.4.122 TODO */
  function parse_Font(blob, length, opts) {
    const o = {
      dyHeight: blob.read_shift(2),
      fl: blob.read_shift(2)
    };
    switch ((opts && opts.biff) || 8) {
      case 2:
        break;
      case 3:
      case 4:
        blob.l += 2;
        break;
      default:
        blob.l += 10;
        break;
    }
    o.name = parse_ShortXLUnicodeString(blob, 0, opts);
    return o;
  }
  function write_Font(data, opts) {
    const name = data.name || "Arial";
    const b5 = opts && opts.biff == 5;
    const w = b5 ? 15 + name.length : 16 + 2 * name.length;
    const o = new_buf(w);
    o.write_shift(2, (data.sz || 12) * 20);
    o.write_shift(4, 0);
    o.write_shift(2, 400);
    o.write_shift(4, 0);
    o.write_shift(2, 0);
    o.write_shift(1, name.length);
    if (!b5) o.write_shift(1, 1);
    o.write_shift((b5 ? 1 : 2) * name.length, name, b5 ? "sbcs" : "utf16le");
    return o;
  }

  /* [MS-XLS] 2.4.149 */
  function parse_LabelSst(blob) {
    const cell = parse_XLSCell(blob);
    cell.isst = blob.read_shift(4);
    return cell;
  }
  function write_LabelSst(R, C, v, os) {
    const o = new_buf(10);
    write_XLSCell(R, C, os, o);
    o.write_shift(4, v);
    return o;
  }

  /* [MS-XLS] 2.4.148 */
  function parse_Label(blob, length, opts) {
    const target = blob.l + length;
    const cell = parse_XLSCell(blob, 6);
    if (opts.biff == 2) blob.l++;
    const str = parse_XLUnicodeString(blob, target - blob.l, opts);
    cell.val = str;
    return cell;
  }
  function write_Label(R, C, v, os, opts) {
    const b8 = !opts || opts.biff == 8;
    const o = new_buf(6 + 2 + +b8 + (1 + b8) * v.length);
    write_XLSCell(R, C, os, o);
    o.write_shift(2, v.length);
    if (b8) o.write_shift(1, 1);
    o.write_shift((1 + b8) * v.length, v, b8 ? "utf16le" : "sbcs");
    return o;
  }

  /* [MS-XLS] 2.4.126 Number Formats */
  function parse_Format(blob, length, opts) {
    const numFmtId = blob.read_shift(2);
    const fmtstr = parse_XLUnicodeString2(blob, 0, opts);
    return [numFmtId, fmtstr];
  }
  function write_Format(i, f, opts, o) {
    const b5 = opts && opts.biff == 5;
    if (!o) o = new_buf(b5 ? 3 + f.length : 5 + 2 * f.length);
    o.write_shift(2, i);
    o.write_shift(b5 ? 1 : 2, f.length);
    if (!b5) o.write_shift(1, 1);
    o.write_shift((b5 ? 1 : 2) * f.length, f, b5 ? "sbcs" : "utf16le");
    const out = o.length > o.l ? o.slice(0, o.l) : o;
    if (out.l == null) out.l = out.length;
    return out;
  }
  const parse_BIFF2Format = parse_XLUnicodeString2;

  /* [MS-XLS] 2.4.90 */
  function parse_Dimensions(blob, length, opts) {
    const end = blob.l + length;
    const w = opts.biff == 8 || !opts.biff ? 4 : 2;
    const r = blob.read_shift(w);
    const R = blob.read_shift(w);
    const c = blob.read_shift(2);
    const C = blob.read_shift(2);
    blob.l = end;
    return { s: { r, c }, e: { r: R, c: C } };
  }
  function write_Dimensions(range, opts) {
    const w = opts.biff == 8 || !opts.biff ? 4 : 2;
    const o = new_buf(2 * w + 6);
    o.write_shift(w, range.s.r);
    o.write_shift(w, range.e.r + 1);
    o.write_shift(2, range.s.c);
    o.write_shift(2, range.e.c + 1);
    o.write_shift(2, 0);
    return o;
  }

  /* [MS-XLS] 2.4.220 */
  function parse_RK(blob) {
    const rw = blob.read_shift(2);
    const col = blob.read_shift(2);
    const rkrec = parse_RkRec(blob);
    return { r: rw, c: col, ixfe: rkrec[0], rknum: rkrec[1] };
  }

  /* [MS-XLS] 2.4.175 */
  function parse_MulRk(blob, length) {
    const target = blob.l + length - 2;
    const rw = blob.read_shift(2);
    const col = blob.read_shift(2);
    const rkrecs = [];
    while (blob.l < target) rkrecs.push(parse_RkRec(blob));
    if (blob.l !== target) throw new Error("MulRK read error");
    const lastcol = blob.read_shift(2);
    if (rkrecs.length != lastcol - col + 1)
      throw new Error("MulRK length mismatch");
    return { r: rw, c: col, C: lastcol, rkrec: rkrecs };
  }
  /* [MS-XLS] 2.4.174 */
  function parse_MulBlank(blob, length) {
    const target = blob.l + length - 2;
    const rw = blob.read_shift(2);
    const col = blob.read_shift(2);
    const ixfes = [];
    while (blob.l < target) ixfes.push(blob.read_shift(2));
    if (blob.l !== target) throw new Error("MulBlank read error");
    const lastcol = blob.read_shift(2);
    if (ixfes.length != lastcol - col + 1)
      throw new Error("MulBlank length mismatch");
    return { r: rw, c: col, C: lastcol, ixfe: ixfes };
  }

  /* [MS-XLS] 2.5.20 2.5.249 TODO: interpret values here */
  function parse_CellStyleXF(blob, length, style, opts) {
    const o = {};
    const a = blob.read_shift(4);
    const b = blob.read_shift(4);
    const c = blob.read_shift(4);
    const d = blob.read_shift(2);
    o.patternType = XLSFillPattern[c >> 26];

    if (!opts.cellStyles) return o;
    o.alc = a & 0x07;
    o.fWrap = (a >> 3) & 0x01;
    o.alcV = (a >> 4) & 0x07;
    o.fJustLast = (a >> 7) & 0x01;
    o.trot = (a >> 8) & 0xff;
    o.cIndent = (a >> 16) & 0x0f;
    o.fShrinkToFit = (a >> 20) & 0x01;
    o.iReadOrder = (a >> 22) & 0x02;
    o.fAtrNum = (a >> 26) & 0x01;
    o.fAtrFnt = (a >> 27) & 0x01;
    o.fAtrAlc = (a >> 28) & 0x01;
    o.fAtrBdr = (a >> 29) & 0x01;
    o.fAtrPat = (a >> 30) & 0x01;
    o.fAtrProt = (a >> 31) & 0x01;

    o.dgLeft = b & 0x0f;
    o.dgRight = (b >> 4) & 0x0f;
    o.dgTop = (b >> 8) & 0x0f;
    o.dgBottom = (b >> 12) & 0x0f;
    o.icvLeft = (b >> 16) & 0x7f;
    o.icvRight = (b >> 23) & 0x7f;
    o.grbitDiag = (b >> 30) & 0x03;

    o.icvTop = c & 0x7f;
    o.icvBottom = (c >> 7) & 0x7f;
    o.icvDiag = (c >> 14) & 0x7f;
    o.dgDiag = (c >> 21) & 0x0f;

    o.icvFore = d & 0x7f;
    o.icvBack = (d >> 7) & 0x7f;
    o.fsxButton = (d >> 14) & 0x01;
    return o;
  }
  // function parse_CellXF(blob, length, opts) {return parse_CellStyleXF(blob,length,0, opts);}
  // function parse_StyleXF(blob, length, opts) {return parse_CellStyleXF(blob,length,1, opts);}

  /* [MS-XLS] 2.4.353 TODO: actually do this right */
  function parse_XF(blob, length, opts) {
    const o = {};
    o.ifnt = blob.read_shift(2);
    o.numFmtId = blob.read_shift(2);
    o.flags = blob.read_shift(2);
    o.fStyle = (o.flags >> 2) & 0x01;
    length -= 6;
    o.data = parse_CellStyleXF(blob, length, o.fStyle, opts);
    return o;
  }
  function write_XF(data, ixfeP, opts, o) {
    const b5 = opts && opts.biff == 5;
    if (!o) o = new_buf(b5 ? 16 : 20);
    o.write_shift(2, 0);
    if (data.style) {
      o.write_shift(2, data.numFmtId || 0);
      o.write_shift(2, 0xfff4);
    } else {
      o.write_shift(2, data.numFmtId || 0);
      o.write_shift(2, ixfeP << 4);
    }
    o.write_shift(4, 0);
    o.write_shift(4, 0);
    if (!b5) o.write_shift(4, 0);
    o.write_shift(2, 0);
    return o;
  }

  /* [MS-XLS] 2.4.134 */
  function parse_Guts(blob) {
    blob.l += 4;
    const out = [blob.read_shift(2), blob.read_shift(2)];
    if (out[0] !== 0) out[0]--;
    if (out[1] !== 0) out[1]--;
    if (out[0] > 7 || out[1] > 7)
      throw new Error(`Bad Gutters: ${out.join("|")}`);
    return out;
  }
  function write_Guts(guts) {
    const o = new_buf(8);
    o.write_shift(4, 0);
    o.write_shift(2, guts[0] ? guts[0] + 1 : 0);
    o.write_shift(2, guts[1] ? guts[1] + 1 : 0);
    return o;
  }

  /* [MS-XLS] 2.4.24 */
  function parse_BoolErr(blob, length, opts) {
    const cell = parse_XLSCell(blob, 6);
    if (opts.biff == 2) ++blob.l;
    const val = parse_Bes(blob, 2);
    cell.val = val;
    cell.t = val === true || val === false ? "b" : "e";
    return cell;
  }
  function write_BoolErr(R, C, v, os, opts, t) {
    const o = new_buf(8);
    write_XLSCell(R, C, os, o);
    write_Bes(v, t, o);
    return o;
  }

  /* [MS-XLS] 2.4.180 Number */
  function parse_Number(blob) {
    const cell = parse_XLSCell(blob, 6);
    const xnum = parse_Xnum(blob, 8);
    cell.val = xnum;
    return cell;
  }
  function write_Number(R, C, v, os) {
    const o = new_buf(14);
    write_XLSCell(R, C, os, o);
    write_Xnum(v, o);
    return o;
  }

  const parse_XLHeaderFooter = parse_OptXLUnicodeString; // TODO: parse 2.4.136

  /* [MS-XLS] 2.4.271 */
  function parse_SupBook(blob, length, opts) {
    const end = blob.l + length;
    const ctab = blob.read_shift(2);
    const cch = blob.read_shift(2);
    opts.sbcch = cch;
    if (cch == 0x0401 || cch == 0x3a01) return [cch, ctab];
    if (cch < 0x01 || cch > 0xff)
      throw new Error(`Unexpected SupBook type: ${cch}`);
    const virtPath = parse_XLUnicodeStringNoCch(blob, cch);
    /* TODO: 2.5.277 Virtual Path */
    const rgst = [];
    while (end > blob.l) rgst.push(parse_XLUnicodeString(blob));
    return [cch, ctab, virtPath, rgst];
  }

  /* [MS-XLS] 2.4.105 TODO */
  function parse_ExternName(blob, length, opts) {
    const flags = blob.read_shift(2);
    let body;
    const o = {
      fBuiltIn: flags & 0x01,
      fWantAdvise: (flags >>> 1) & 0x01,
      fWantPict: (flags >>> 2) & 0x01,
      fOle: (flags >>> 3) & 0x01,
      fOleLink: (flags >>> 4) & 0x01,
      cf: (flags >>> 5) & 0x3ff,
      fIcon: (flags >>> 15) & 0x01
    };
    if (opts.sbcch === 0x3a01) body = parse_AddinUdf(blob, length - 2, opts);
    // else throw new Error("unsupported SupBook cch: " + opts.sbcch);
    o.body = body || blob.read_shift(length - 2);
    if (typeof body === "string") o.Name = body;
    return o;
  }

  /* [MS-XLS] 2.4.150 TODO */
  const XLSLblBuiltIn = [
    "_xlnm.Consolidate_Area",
    "_xlnm.Auto_Open",
    "_xlnm.Auto_Close",
    "_xlnm.Extract",
    "_xlnm.Database",
    "_xlnm.Criteria",
    "_xlnm.Print_Area",
    "_xlnm.Print_Titles",
    "_xlnm.Recorder",
    "_xlnm.Data_Form",
    "_xlnm.Auto_Activate",
    "_xlnm.Auto_Deactivate",
    "_xlnm.Sheet_Title",
    "_xlnm._FilterDatabase"
  ];
  function parse_Lbl(blob, length, opts) {
    const target = blob.l + length;
    const flags = blob.read_shift(2);
    const chKey = blob.read_shift(1);
    const cch = blob.read_shift(1);
    const cce = blob.read_shift(opts && opts.biff == 2 ? 1 : 2);
    let itab = 0;
    if (!opts || opts.biff >= 5) {
      if (opts.biff != 5) blob.l += 2;
      itab = blob.read_shift(2);
      if (opts.biff == 5) blob.l += 2;
      blob.l += 4;
    }
    let name = parse_XLUnicodeStringNoCch(blob, cch, opts);
    if (flags & 0x20) name = XLSLblBuiltIn[name.charCodeAt(0)];
    let npflen = target - blob.l;
    if (opts && opts.biff == 2) --npflen;
    const rgce =
      target == blob.l || cce === 0
        ? []
        : parse_NameParsedFormula(blob, npflen, opts, cce);
    return {
      chKey,
      Name: name,
      itab,
      rgce
    };
  }

  /* [MS-XLS] 2.4.106 TODO: verify filename encoding */
  function parse_ExternSheet(blob, length, opts) {
    if (opts.biff < 8) return parse_BIFF5ExternSheet(blob, length, opts);
    const o = [];
    const target = blob.l + length;
    let len = blob.read_shift(opts.biff > 8 ? 4 : 2);
    while (len-- !== 0) o.push(parse_XTI(blob, opts.biff > 8 ? 12 : 6, opts));
    // [iSupBook, itabFirst, itabLast];
    if (blob.l != target)
      throw new Error(`Bad ExternSheet: ${blob.l} != ${target}`);
    return o;
  }
  function parse_BIFF5ExternSheet(blob, length, opts) {
    if (blob[blob.l + 1] == 0x03) blob[blob.l]++;
    const o = parse_ShortXLUnicodeString(blob, length, opts);
    return o.charCodeAt(0) == 0x03 ? o.slice(1) : o;
  }

  /* [MS-XLS] 2.4.176 TODO: check older biff */
  function parse_NameCmt(blob, length, opts) {
    if (opts.biff < 8) {
      blob.l += length;
      return;
    }
    const cchName = blob.read_shift(2);
    const cchComment = blob.read_shift(2);
    const name = parse_XLUnicodeStringNoCch(blob, cchName, opts);
    const comment = parse_XLUnicodeStringNoCch(blob, cchComment, opts);
    return [name, comment];
  }

  /* [MS-XLS] 2.4.260 */
  function parse_ShrFmla(blob, length, opts) {
    const ref = parse_RefU(blob, 6);
    blob.l++;
    const cUse = blob.read_shift(1);
    length -= 8;
    return [parse_SharedParsedFormula(blob, length, opts), cUse, ref];
  }

  /* [MS-XLS] 2.4.4 TODO */
  function parse_Array(blob, length, opts) {
    const ref = parse_Ref(blob, 6);
    /* TODO: fAlwaysCalc */
    switch (opts.biff) {
      case 2:
        blob.l++;
        length -= 7;
        break;
      case 3:
      case 4:
        blob.l += 2;
        length -= 8;
        break;
      default:
        blob.l += 6;
        length -= 12;
    }
    return [ref, parse_ArrayParsedFormula(blob, length, opts, ref)];
  }

  /* [MS-XLS] 2.4.173 */
  function parse_MTRSettings(blob) {
    const fMTREnabled = blob.read_shift(4) !== 0x00;
    const fUserSetThreadCount = blob.read_shift(4) !== 0x00;
    const cUserThreadCount = blob.read_shift(4);
    return [fMTREnabled, fUserSetThreadCount, cUserThreadCount];
  }

  /* [MS-XLS] 2.5.186 TODO: BIFF5 */
  function parse_NoteSh(blob, length, opts) {
    if (opts.biff < 8) return;
    const row = blob.read_shift(2);
    const col = blob.read_shift(2);
    const flags = blob.read_shift(2);
    const idObj = blob.read_shift(2);
    const stAuthor = parse_XLUnicodeString2(blob, 0, opts);
    if (opts.biff < 8) blob.read_shift(1);
    return [{ r: row, c: col }, stAuthor, idObj, flags];
  }

  /* [MS-XLS] 2.4.179 */
  function parse_Note(blob, length, opts) {
    /* TODO: Support revisions */
    return parse_NoteSh(blob, length, opts);
  }

  /* [MS-XLS] 2.4.168 */
  function parse_MergeCells(blob, length) {
    const merges = [];
    let cmcs = blob.read_shift(2);
    while (cmcs--) merges.push(parse_Ref8U(blob, length));
    return merges;
  }
  function write_MergeCells(merges) {
    const o = new_buf(2 + merges.length * 8);
    o.write_shift(2, merges.length);
    for (let i = 0; i < merges.length; ++i) write_Ref8U(merges[i], o);
    return o;
  }

  /* [MS-XLS] 2.4.181 TODO: parse all the things! */
  function parse_Obj(blob, length, opts) {
    if (opts && opts.biff < 8) return parse_BIFF5Obj(blob, length, opts);
    const cmo = parse_FtCmo(blob, 22); // id, ot, flags
    const fts = parse_FtArray(blob, length - 22, cmo[1]);
    return { cmo, ft: fts };
  }
  /* from older spec */
  const parse_BIFF5OT = [];
  parse_BIFF5OT[0x08] = function(blob, length) {
    const tgt = blob.l + length;
    blob.l += 10; // todo
    const cf = blob.read_shift(2);
    blob.l += 4;
    blob.l += 2; // var cbPictFmla = blob.read_shift(2);
    blob.l += 2;
    blob.l += 2; // var grbit = blob.read_shift(2);
    blob.l += 4;
    const cchName = blob.read_shift(1);
    blob.l += cchName; // TODO: stName
    blob.l = tgt; // TODO: fmla
    return { fmt: cf };
  };

  function parse_BIFF5Obj(blob, length, opts) {
    blob.l += 4; // var cnt = blob.read_shift(4);
    const ot = blob.read_shift(2);
    const id = blob.read_shift(2);
    const grbit = blob.read_shift(2);
    blob.l += 2; // var colL = blob.read_shift(2);
    blob.l += 2; // var dxL = blob.read_shift(2);
    blob.l += 2; // var rwT = blob.read_shift(2);
    blob.l += 2; // var dyT = blob.read_shift(2);
    blob.l += 2; // var colR = blob.read_shift(2);
    blob.l += 2; // var dxR = blob.read_shift(2);
    blob.l += 2; // var rwB = blob.read_shift(2);
    blob.l += 2; // var dyB = blob.read_shift(2);
    blob.l += 2; // var cbMacro = blob.read_shift(2);
    blob.l += 6;
    length -= 36;
    const fts = [];
    fts.push((parse_BIFF5OT[ot] || parsenoop)(blob, length, opts));
    return { cmo: [id, ot, grbit], ft: fts };
  }

  /* [MS-XLS] 2.4.329 TODO: parse properly */
  function parse_TxO(blob, length, opts) {
    const s = blob.l;
    let texts = "";
    try {
      blob.l += 4;
      const ot = (opts.lastobj || { cmo: [0, 0] }).cmo[1];
      let controlInfo; // eslint-disable-line no-unused-vars
      if ([0, 5, 7, 11, 12, 14].indexOf(ot) == -1) blob.l += 6;
      else controlInfo = parse_ControlInfo(blob, 6, opts);
      const cchText = blob.read_shift(2);
      /* var cbRuns = */ blob.read_shift(2);
      /* var ifntEmpty = */ parseuint16(blob, 2);
      const len = blob.read_shift(2);
      blob.l += len;
      // var fmla = parse_ObjFmla(blob, s + length - blob.l);

      for (let i = 1; i < blob.lens.length - 1; ++i) {
        if (blob.l - s != blob.lens[i])
          throw new Error("TxO: bad continue record");
        const hdr = blob[blob.l];
        const t = parse_XLUnicodeStringNoCch(
          blob,
          blob.lens[i + 1] - blob.lens[i] - 1
        );
        texts += t;
        if (texts.length >= (hdr ? cchText : 2 * cchText)) break;
      }
      if (texts.length !== cchText && texts.length !== cchText * 2) {
        throw new Error(`cchText: ${cchText} != ${texts.length}`);
      }

      blob.l = s + length;
      /* [MS-XLS] 2.5.272 TxORuns */
      //	var rgTxoRuns = [];
      //	for(var j = 0; j != cbRuns/8-1; ++j) blob.l += 8;
      //	var cchText2 = blob.read_shift(2);
      //	if(cchText2 !== cchText) throw new Error("TxOLastRun mismatch: " + cchText2 + " " + cchText);
      //	blob.l += 6;
      //	if(s + length != blob.l) throw new Error("TxO " + (s + length) + ", at " + blob.l);
      return { t: texts };
    } catch (e) {
      blob.l = s + length;
      return { t: texts };
    }
  }

  /* [MS-XLS] 2.4.140 */
  function parse_HLink(blob, length) {
    const ref = parse_Ref8U(blob, 8);
    blob.l += 16; /* CLSID */
    const hlink = parse_Hyperlink(blob, length - 24);
    return [ref, hlink];
  }
  function write_HLink(hl) {
    const O = new_buf(24);
    const ref = decode_cell(hl[0]);
    O.write_shift(2, ref.r);
    O.write_shift(2, ref.r);
    O.write_shift(2, ref.c);
    O.write_shift(2, ref.c);
    const clsid = "d0 c9 ea 79 f9 ba ce 11 8c 82 00 aa 00 4b a9 0b".split(" ");
    for (let i = 0; i < 16; ++i) O.write_shift(1, parseInt(clsid[i], 16));
    return bconcat([O, write_Hyperlink(hl[1])]);
  }

  /* [MS-XLS] 2.4.141 */
  function parse_HLinkTooltip(blob, length) {
    blob.read_shift(2);
    const ref = parse_Ref8U(blob, 8);
    let wzTooltip = blob.read_shift((length - 10) / 2, "dbcs-cont");
    wzTooltip = wzTooltip.replace(chr0, "");
    return [ref, wzTooltip];
  }
  function write_HLinkTooltip(hl) {
    const TT = hl[1].Tooltip;
    const O = new_buf(10 + 2 * (TT.length + 1));
    O.write_shift(2, 0x0800);
    const ref = decode_cell(hl[0]);
    O.write_shift(2, ref.r);
    O.write_shift(2, ref.r);
    O.write_shift(2, ref.c);
    O.write_shift(2, ref.c);
    for (let i = 0; i < TT.length; ++i) O.write_shift(2, TT.charCodeAt(i));
    O.write_shift(2, 0);
    return O;
  }

  /* [MS-XLS] 2.4.63 */
  function parse_Country(blob) {
    const o = [0, 0];
    let d;
    d = blob.read_shift(2);
    o[0] = CountryEnum[d] || d;
    d = blob.read_shift(2);
    o[1] = CountryEnum[d] || d;
    return o;
  }
  function write_Country(o) {
    if (!o) o = new_buf(4);
    o.write_shift(2, 0x01);
    o.write_shift(2, 0x01);
    return o;
  }

  /* [MS-XLS] 2.4.50 ClrtClient */
  function parse_ClrtClient(blob) {
    let ccv = blob.read_shift(2);
    const o = [];
    while (ccv-- > 0) o.push(parse_LongRGB(blob, 8));
    return o;
  }

  /* [MS-XLS] 2.4.188 */
  function parse_Palette(blob) {
    let ccv = blob.read_shift(2);
    const o = [];
    while (ccv-- > 0) o.push(parse_LongRGB(blob, 8));
    return o;
  }

  /* [MS-XLS] 2.4.354 */
  function parse_XFCRC(blob) {
    blob.l += 2;
    const o = { cxfs: 0, crc: 0 };
    o.cxfs = blob.read_shift(2);
    o.crc = blob.read_shift(4);
    return o;
  }

  /* [MS-XLS] 2.4.53 TODO: parse flags */
  /* [MS-XLSB] 2.4.323 TODO: parse flags */
  function parse_ColInfo(blob, length, opts) {
    if (!opts.cellStyles) return parsenoop(blob, length);
    const w = opts && opts.biff >= 12 ? 4 : 2;
    const colFirst = blob.read_shift(w);
    const colLast = blob.read_shift(w);
    const coldx = blob.read_shift(w);
    const ixfe = blob.read_shift(w);
    const flags = blob.read_shift(2);
    if (w == 2) blob.l += 2;
    const o = { s: colFirst, e: colLast, w: coldx, ixfe, flags };
    if (opts.biff >= 5 || !opts.biff) o.level = (flags >> 8) & 0x7;
    return o;
  }

  /* [MS-XLS] 2.4.257 */
  function parse_Setup(blob, length) {
    const o = {};
    if (length < 32) return o;
    blob.l += 16;
    o.header = parse_Xnum(blob, 8);
    o.footer = parse_Xnum(blob, 8);
    blob.l += 2;
    return o;
  }

  /* [MS-XLS] 2.4.261 */
  function parse_ShtProps(blob, length, opts) {
    const def = { area: false };
    if (opts.biff != 5) {
      blob.l += length;
      return def;
    }
    const d = blob.read_shift(1);
    blob.l += 3;
    if (d & 0x10) def.area = true;
    return def;
  }

  /* [MS-XLS] 2.4.241 */
  function write_RRTabId(n) {
    const out = new_buf(2 * n);
    for (let i = 0; i < n; ++i) out.write_shift(2, i + 1);
    return out;
  }

  const parse_Blank = parse_XLSCell; /* [MS-XLS] 2.4.20 Just the cell */
  const parse_Scl = parseuint16a; /* [MS-XLS] 2.4.247 num, den */
  const parse_String = parse_XLUnicodeString; /* [MS-XLS] 2.4.268 */

  /* --- Specific to versions before BIFF8 --- */
  function parse_ImData(blob) {
    const cf = blob.read_shift(2);
    const env = blob.read_shift(2);
    const lcb = blob.read_shift(4);
    const o = {
      fmt: cf,
      env,
      len: lcb,
      data: blob.slice(blob.l, blob.l + lcb)
    };
    blob.l += lcb;
    return o;
  }

  /* BIFF2_??? where ??? is the name from [XLS] */
  function parse_BIFF2STR(blob, length, opts) {
    const cell = parse_XLSCell(blob, 6);
    ++blob.l;
    const str = parse_XLUnicodeString2(blob, length - 7, opts);
    cell.t = "str";
    cell.val = str;
    return cell;
  }

  function parse_BIFF2NUM(blob) {
    const cell = parse_XLSCell(blob, 6);
    ++blob.l;
    const num = parse_Xnum(blob, 8);
    cell.t = "n";
    cell.val = num;
    return cell;
  }
  function write_BIFF2NUM(r, c, val) {
    const out = new_buf(15);
    write_BIFF2Cell(out, r, c);
    out.write_shift(8, val, "f");
    return out;
  }

  function parse_BIFF2INT(blob) {
    const cell = parse_XLSCell(blob, 6);
    ++blob.l;
    const num = blob.read_shift(2);
    cell.t = "n";
    cell.val = num;
    return cell;
  }
  function write_BIFF2INT(r, c, val) {
    const out = new_buf(9);
    write_BIFF2Cell(out, r, c);
    out.write_shift(2, val);
    return out;
  }

  function parse_BIFF2STRING(blob) {
    const cch = blob.read_shift(1);
    if (cch === 0) {
      blob.l++;
      return "";
    }
    return blob.read_shift(cch, "sbcs-cont");
  }

  /* TODO: convert to BIFF8 font struct */
  function parse_BIFF2FONTXTRA(blob, length) {
    blob.l += 6; // unknown
    blob.l += 2; // font weight "bls"
    blob.l += 1; // charset
    blob.l += 3; // unknown
    blob.l += 1; // font family
    blob.l += length - 13;
  }

  /* TODO: parse rich text runs */
  function parse_RString(blob, length, opts) {
    const end = blob.l + length;
    const cell = parse_XLSCell(blob, 6);
    const cch = blob.read_shift(2);
    const str = parse_XLUnicodeStringNoCch(blob, cch, opts);
    blob.l = end;
    cell.t = "str";
    cell.val = str;
    return cell;
  }
  /* from js-harb (C) 2014-present  SheetJS */
  const DBF = (function() {
    const dbf_codepage_map = {
      /* Code Pages Supported by Visual FoxPro */
      0x01: 437,
      0x02: 850,
      0x03: 1252,
      0x04: 10000,
      0x64: 852,
      0x65: 866,
      0x66: 865,
      0x67: 861,
      0x68: 895,
      0x69: 620,
      0x6a: 737,
      0x6b: 857,
      0x78: 950,
      0x79: 949,
      0x7a: 936,
      0x7b: 932,
      0x7c: 874,
      0x7d: 1255,
      0x7e: 1256,
      0x96: 10007,
      0x97: 10029,
      0x98: 10006,
      0xc8: 1250,
      0xc9: 1251,
      0xca: 1254,
      0xcb: 1253,

      /* shapefile DBF extension */
      0x00: 20127,
      0x08: 865,
      0x09: 437,
      0x0a: 850,
      0x0b: 437,
      0x0d: 437,
      0x0e: 850,
      0x0f: 437,
      0x10: 850,
      0x11: 437,
      0x12: 850,
      0x13: 932,
      0x14: 850,
      0x15: 437,
      0x16: 850,
      0x17: 865,
      0x18: 437,
      0x19: 437,
      0x1a: 850,
      0x1b: 437,
      0x1c: 863,
      0x1d: 850,
      0x1f: 852,
      0x22: 852,
      0x23: 852,
      0x24: 860,
      0x25: 850,
      0x26: 866,
      0x37: 850,
      0x40: 852,
      0x4d: 936,
      0x4e: 949,
      0x4f: 950,
      0x50: 874,
      0x57: 1252,
      0x58: 1252,
      0x59: 1252,

      0xff: 16969
    };
    const dbf_reverse_map = evert({
      0x01: 437,
      0x02: 850,
      0x03: 1252,
      0x04: 10000,
      0x64: 852,
      0x65: 866,
      0x66: 865,
      0x67: 861,
      0x68: 895,
      0x69: 620,
      0x6a: 737,
      0x6b: 857,
      0x78: 950,
      0x79: 949,
      0x7a: 936,
      0x7b: 932,
      0x7c: 874,
      0x7d: 1255,
      0x7e: 1256,
      0x96: 10007,
      0x97: 10029,
      0x98: 10006,
      0xc8: 1250,
      0xc9: 1251,
      0xca: 1254,
      0xcb: 1253,
      0x00: 20127
    });
    const DBF_SUPPORTED_VERSIONS = [
      0x02,
      0x03,
      0x30,
      0x31,
      0x83,
      0x8b,
      0x8c,
      0xf5
    ];
    /* TODO: find an actual specification */
    function dbf_to_aoa(buf, opts) {
      let out = [];
      /* TODO: browser based */
      let d = new_raw_buf(1);
      switch (opts.type) {
        case "base64":
          d = s2a(Base64.decode(buf));
          break;
        case "binary":
          d = s2a(buf);
          break;
        case "buffer":
        case "array":
          d = buf;
          break;
      }
      prep_blob(d, 0);
      /* header */
      const ft = d.read_shift(1);
      let memo = false;
      let vfp = false;
      let l7 = false;
      switch (ft) {
        case 0x02:
        case 0x03:
          break;
        case 0x30:
          vfp = true;
          memo = true;
          break;
        case 0x31:
          vfp = true;
          break;
        case 0x83:
          memo = true;
          break;
        case 0x8b:
          memo = true;
          break;
        case 0x8c:
          memo = true;
          l7 = true;
          break;
        case 0xf5:
          memo = true;
          break;
        default:
          throw new Error(`DBF Unsupported Version: ${ft.toString(16)}`);
      }
      let /* filedate = new Date(), */ nrow = 0;
      let fpos = 0;
      if (ft == 0x02) nrow = d.read_shift(2);
      /* filedate = new Date(d.read_shift(1) + 1900, d.read_shift(1) - 1, d.read_shift(1)); */ d.l += 3;
      if (ft != 0x02) nrow = d.read_shift(4);
      if (ft != 0x02) fpos = d.read_shift(2);
      const rlen = d.read_shift(2);

      let /* flags = 0, */ current_cp = 1252;
      if (ft != 0x02) {
        d.l += 16;
        /* flags = */ d.read_shift(1);
        // if(memo && ((flags & 0x02) === 0)) throw new Error("DBF Flags " + flags.toString(16) + " ft " + ft.toString(16));

        /* codepage present in FoxPro */
        if (d[d.l] !== 0) current_cp = dbf_codepage_map[d[d.l]];
        d.l += 1;

        d.l += 2;
      }
      if (l7) d.l += 36;
      const fields = [];
      let field = {};
      const hend = fpos - 10 - (vfp ? 264 : 0);
      const ww = l7 ? 32 : 11;
      while (ft == 0x02 ? d.l < d.length && d[d.l] != 0x0d : d.l < hend) {
        field = {};
        field.name = cptable.utils
          .decode(current_cp, d.slice(d.l, d.l + ww))
          .replace(/[\u0000\r\n].*$/g, "");
        d.l += ww;
        field.type = String.fromCharCode(d.read_shift(1));
        if (ft != 0x02 && !l7) field.offset = d.read_shift(4);
        field.len = d.read_shift(1);
        if (ft == 0x02) field.offset = d.read_shift(2);
        field.dec = d.read_shift(1);
        if (field.name.length) fields.push(field);
        if (ft != 0x02) d.l += l7 ? 13 : 14;
        switch (field.type) {
          case "B": // VFP Double
            if ((!vfp || field.len != 8) && opts.WTF)
              console.log(`Skipping ${field.name}:${field.type}`);
            break;
          case "G": // General
          case "P": // Picture
            if (opts.WTF) console.log(`Skipping ${field.name}:${field.type}`);
            break;
          case "C": // character
          case "D": // date
          case "F": // floating point
          case "I": // long
          case "L": // boolean
          case "M": // memo
          case "N": // number
          case "O": // double
          case "T": // datetime
          case "Y": // currency
          case "0": // VFP _NullFlags
          case "@": // timestamp
          case "+": // autoincrement
            break;
          default:
            throw new Error(`Unknown Field Type: ${field.type}`);
        }
      }
      if (d[d.l] !== 0x0d) d.l = fpos - 1;
      else if (ft == 0x02) d.l = 0x209;
      if (ft != 0x02) {
        if (d.read_shift(1) !== 0x0d)
          throw new Error(`DBF Terminator not found ${d.l} ${d[d.l]}`);
        d.l = fpos;
      }
      /* data */
      let R = 0;
      let C = 0;
      out[0] = [];
      for (C = 0; C != fields.length; ++C) out[0][C] = fields[C].name;
      while (nrow-- > 0) {
        if (d[d.l] === 0x2a) {
          d.l += rlen;
          continue;
        }
        ++d.l;
        out[++R] = [];
        C = 0;
        for (C = 0; C != fields.length; ++C) {
          const dd = d.slice(d.l, d.l + fields[C].len);
          d.l += fields[C].len;
          prep_blob(dd, 0);
          const s = cptable.utils.decode(current_cp, dd);
          switch (fields[C].type) {
            case "C":
              out[R][C] = cptable.utils.decode(current_cp, dd);
              out[R][C] = out[R][C].trim();
              break;
            case "D":
              if (s.length === 8)
                out[R][C] = new Date(
                  +s.slice(0, 4),
                  +s.slice(4, 6) - 1,
                  +s.slice(6, 8)
                );
              else out[R][C] = s;
              break;
            case "F":
              out[R][C] = parseFloat(s.trim());
              break;
            case "+":
            case "I":
              out[R][C] = l7
                ? dd.read_shift(-4, "i") ^ 0x80000000
                : dd.read_shift(4, "i");
              break;
            case "L":
              switch (s.toUpperCase()) {
                case "Y":
                case "T":
                  out[R][C] = true;
                  break;
                case "N":
                case "F":
                  out[R][C] = false;
                  break;
                case " ":
                case "?":
                  out[R][C] = false;
                  break; /* NOTE: technically uninitialized */
                default:
                  throw new Error(`DBF Unrecognized L:|${s}|`);
              }
              break;
            case "M" /* TODO: handle memo files */:
              if (!memo)
                throw new Error(
                  `DBF Unexpected MEMO for type ${ft.toString(16)}`
                );
              out[R][C] = `##MEMO##${
                l7 ? parseInt(s.trim(), 10) : dd.read_shift(4)
              }`;
              break;
            case "N":
              out[R][C] = +s.replace(/\u0000/g, "").trim();
              break;
            case "@":
              out[R][C] = new Date(dd.read_shift(-8, "f") - 0x388317533400);
              break;
            case "T":
              out[R][C] = new Date(
                (dd.read_shift(4) - 0x253d8c) * 0x5265c00 + dd.read_shift(4)
              );
              break;
            case "Y":
              out[R][C] = dd.read_shift(4, "i") / 1e4;
              break;
            case "O":
              out[R][C] = -dd.read_shift(-8, "f");
              break;
            case "B":
              if (vfp && fields[C].len == 8) {
                out[R][C] = dd.read_shift(8, "f");
                break;
              }
            /* falls through */
            case "G":
            case "P":
              dd.l += fields[C].len;
              break;
            case "0":
              if (fields[C].name === "_NullFlags") break;
            /* falls through */
            default:
              throw new Error(`DBF Unsupported data type ${fields[C].type}`);
          }
        }
      }
      if (ft != 0x02)
        if (d.l < d.length && d[d.l++] != 0x1a)
          throw new Error(
            `DBF EOF Marker missing ${d.l - 1} of ${d.length} ${d[
              d.l - 1
            ].toString(16)}`
          );
      if (opts && opts.sheetRows) out = out.slice(0, opts.sheetRows);
      return out;
    }

    function dbf_to_sheet(buf, opts) {
      const o = opts || {};
      if (!o.dateNF) o.dateNF = "yyyymmdd";
      return aoa_to_sheet(dbf_to_aoa(buf, o), o);
    }

    function dbf_to_workbook(buf, opts) {
      try {
        return sheet_to_workbook(dbf_to_sheet(buf, opts), opts);
      } catch (e) {
        if (opts && opts.WTF) throw e;
      }
      return { SheetNames: [], Sheets: {} };
    }

    const _RLEN = { B: 8, C: 250, L: 1, D: 8, "?": 0, "": 0 };
    function sheet_to_dbf(ws, opts) {
      const o = opts || {};
      if (+o.codepage >= 0) set_cp(+o.codepage);
      if (o.type == "string") throw new Error("Cannot write DBF to JS string");
      const ba = buf_array();
      const aoa = sheet_to_json(ws, { header: 1, raw: true, cellDates: true });
      const headers = aoa[0];
      const data = aoa.slice(1);
      let i = 0;
      let j = 0;
      let hcnt = 0;
      let rlen = 1;
      for (i = 0; i < headers.length; ++i) {
        if (i == null) continue;
        ++hcnt;
        if (typeof headers[i] === "number")
          headers[i] = headers[i].toString(10);
        if (typeof headers[i] !== "string")
          throw new Error(
            `DBF Invalid column name ${headers[i]} |${typeof headers[i]}|`
          );
        if (headers.indexOf(headers[i]) !== i)
          for (j = 0; j < 1024; ++j)
            if (headers.indexOf(`${headers[i]}_${j}`) == -1) {
              headers[i] += `_${j}`;
              break;
            }
      }
      const range = safe_decode_range(ws["!ref"]);
      const coltypes = [];
      for (i = 0; i <= range.e.c - range.s.c; ++i) {
        const col = [];
        for (j = 0; j < data.length; ++j) {
          if (data[j][i] != null) col.push(data[j][i]);
        }
        if (col.length == 0 || headers[i] == null) {
          coltypes[i] = "?";
          continue;
        }
        let guess = "";
        let _guess = "";
        for (j = 0; j < col.length; ++j) {
          switch (typeof col[j]) {
            /* TODO: check if L2 compat is desired */
            case "number":
              _guess = "B";
              break;
            case "string":
              _guess = "C";
              break;
            case "boolean":
              _guess = "L";
              break;
            case "object":
              _guess = col[j] instanceof Date ? "D" : "C";
              break;
            default:
              _guess = "C";
          }
          guess = guess && guess != _guess ? "C" : _guess;
          if (guess == "C") break;
        }
        rlen += _RLEN[guess] || 0;
        coltypes[i] = guess;
      }

      const h = ba.next(32);
      h.write_shift(4, 0x13021130);
      h.write_shift(4, data.length);
      h.write_shift(2, 296 + 32 * hcnt);
      h.write_shift(2, rlen);
      for (i = 0; i < 4; ++i) h.write_shift(4, 0);
      h.write_shift(
        4,
        0x00000000 | ((+dbf_reverse_map[current_ansi] || 0x03) << 8)
      );

      for (i = 0, j = 0; i < headers.length; ++i) {
        if (headers[i] == null) continue;
        const hf = ba.next(32);
        const _f = `${headers[i].slice(
          -10
        )}\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00`.slice(0, 11);
        hf.write_shift(1, _f, "sbcs");
        hf.write_shift(1, coltypes[i] == "?" ? "C" : coltypes[i], "sbcs");
        hf.write_shift(4, j);
        hf.write_shift(1, _RLEN[coltypes[i]] || 0);
        hf.write_shift(1, 0);
        hf.write_shift(1, 0x02);
        hf.write_shift(4, 0);
        hf.write_shift(1, 0);
        hf.write_shift(4, 0);
        hf.write_shift(4, 0);
        j += _RLEN[coltypes[i]] || 0;
      }

      const hb = ba.next(264);
      hb.write_shift(4, 0x0000000d);
      for (i = 0; i < 65; ++i) hb.write_shift(4, 0x00000000);
      for (i = 0; i < data.length; ++i) {
        const rout = ba.next(rlen);
        rout.write_shift(1, 0);
        for (j = 0; j < headers.length; ++j) {
          if (headers[j] == null) continue;
          switch (coltypes[j]) {
            case "L":
              rout.write_shift(
                1,
                data[i][j] == null ? 0x3f : data[i][j] ? 0x54 : 0x46
              );
              break;
            case "B":
              rout.write_shift(8, data[i][j] || 0, "f");
              break;
            case "D":
              if (!data[i][j]) rout.write_shift(8, "00000000", "sbcs");
              else {
                rout.write_shift(
                  4,
                  `0000${data[i][j].getFullYear()}`.slice(-4),
                  "sbcs"
                );
                rout.write_shift(
                  2,
                  `00${data[i][j].getMonth() + 1}`.slice(-2),
                  "sbcs"
                );
                rout.write_shift(
                  2,
                  `00${data[i][j].getDate()}`.slice(-2),
                  "sbcs"
                );
              }
              break;
            case "C":
              var _s = String(data[i][j] || "");
              rout.write_shift(1, _s, "sbcs");
              for (hcnt = 0; hcnt < 250 - _s.length; ++hcnt)
                rout.write_shift(1, 0x20);
              break;
          }
        }
        // data
      }
      ba.next(1).write_shift(1, 0x1a);
      return ba.end();
    }
    return {
      versions: DBF_SUPPORTED_VERSIONS,
      to_workbook: dbf_to_workbook,
      to_sheet: dbf_to_sheet,
      from_sheet: sheet_to_dbf
    };
  })();

  const SYLK = (function() {
    /* TODO: stress test sequences */
    const sylk_escapes = {
      AA: "À",
      BA: "Á",
      CA: "Â",
      DA: 195,
      HA: "Ä",
      JA: 197,
      AE: "È",
      BE: "É",
      CE: "Ê",
      HE: "Ë",
      AI: "Ì",
      BI: "Í",
      CI: "Î",
      HI: "Ï",
      AO: "Ò",
      BO: "Ó",
      CO: "Ô",
      DO: 213,
      HO: "Ö",
      AU: "Ù",
      BU: "Ú",
      CU: "Û",
      HU: "Ü",
      Aa: "à",
      Ba: "á",
      Ca: "â",
      Da: 227,
      Ha: "ä",
      Ja: 229,
      Ae: "è",
      Be: "é",
      Ce: "ê",
      He: "ë",
      Ai: "ì",
      Bi: "í",
      Ci: "î",
      Hi: "ï",
      Ao: "ò",
      Bo: "ó",
      Co: "ô",
      Do: 245,
      Ho: "ö",
      Au: "ù",
      Bu: "ú",
      Cu: "û",
      Hu: "ü",
      KC: "Ç",
      Kc: "ç",
      q: "æ",
      z: "œ",
      a: "Æ",
      j: "Œ",
      DN: 209,
      Dn: 241,
      Hy: 255,
      S: 169,
      c: 170,
      R: 174,
      B: 180,
      0: 176,
      1: 177,
      2: 178,
      3: 179,
      5: 181,
      6: 182,
      7: 183,
      Q: 185,
      k: 186,
      b: 208,
      i: 216,
      l: 222,
      s: 240,
      y: 248,
      "!": 161,
      '"': 162,
      "#": 163,
      "(": 164,
      "%": 165,
      "'": 167,
      "H ": 168,
      "+": 171,
      ";": 187,
      "<": 188,
      "=": 189,
      ">": 190,
      "?": 191,
      "{": 223
    };
    const sylk_char_regex = new RegExp(
      `\u001BN(${keys(sylk_escapes)
        .join("|")
        .replace(/\|\|\|/, "|\\||")
        .replace(/([?()+])/g, "\\$1")}|\\|)`,
      "gm"
    );
    const sylk_char_fn = function(_, $1) {
      const o = sylk_escapes[$1];
      return typeof o === "number" ? _getansi(o) : o;
    };
    const decode_sylk_char = function($$, $1, $2) {
      const newcc =
        (($1.charCodeAt(0) - 0x20) << 4) | ($2.charCodeAt(0) - 0x30);
      return newcc == 59 ? $$ : _getansi(newcc);
    };
    sylk_escapes["|"] = 254;
    /* TODO: find an actual specification */
    function sylk_to_aoa(d, opts) {
      switch (opts.type) {
        case "base64":
          return sylk_to_aoa_str(Base64.decode(d), opts);
        case "binary":
          return sylk_to_aoa_str(d, opts);
        case "buffer":
          return sylk_to_aoa_str(d.toString("binary"), opts);
        case "array":
          return sylk_to_aoa_str(cc2str(d), opts);
      }
      throw new Error(`Unrecognized type ${opts.type}`);
    }
    function sylk_to_aoa_str(str, opts) {
      const records = str.split(/[\n\r]+/);
      let R = -1;
      let C = -1;
      let ri = 0;
      let rj = 0;
      let arr = [];
      const formats = [];
      let next_cell_format = null;
      const sht = {};
      const rowinfo = [];
      const colinfo = [];
      let cw = [];
      let Mval = 0;
      let j;
      if (+opts.codepage >= 0) set_cp(+opts.codepage);
      for (; ri !== records.length; ++ri) {
        Mval = 0;
        const rstr = records[ri]
          .trim()
          .replace(/\x1B([\x20-\x2F])([\x30-\x3F])/g, decode_sylk_char)
          .replace(sylk_char_regex, sylk_char_fn);
        const record = rstr
          .replace(/;;/g, "\u0000")
          .split(";")
          .map(function(x) {
            return x.replace(/\u0000/g, ";");
          });
        const RT = record[0];
        var val;
        if (rstr.length > 0)
          switch (RT) {
            case "ID":
              break; /* header */
            case "E":
              break; /* EOF */
            case "B":
              break; /* dimensions */
            case "O":
              break; /* options? */
            case "P":
              if (record[1].charAt(0) == "P")
                formats.push(rstr.slice(3).replace(/;;/g, ";"));
              break;
            case "C":
              var C_seen_K = false;
              var C_seen_X = false;
              for (rj = 1; rj < record.length; ++rj)
                switch (record[rj].charAt(0)) {
                  case "X":
                    C = parseInt(record[rj].slice(1)) - 1;
                    C_seen_X = true;
                    break;
                  case "Y":
                    R = parseInt(record[rj].slice(1)) - 1;
                    if (!C_seen_X) C = 0;
                    for (j = arr.length; j <= R; ++j) arr[j] = [];
                    break;
                  case "K":
                    val = record[rj].slice(1);
                    if (val.charAt(0) === '"')
                      val = val.slice(1, val.length - 1);
                    else if (val === "TRUE") val = true;
                    else if (val === "FALSE") val = false;
                    else if (!isNaN(fuzzynum(val))) {
                      val = fuzzynum(val);
                      if (
                        next_cell_format !== null &&
                        SSF.is_date(next_cell_format)
                      )
                        val = numdate(val);
                    } else if (!isNaN(fuzzydate(val).getDate())) {
                      val = parseDate(val);
                    }
                    if (
                      typeof cptable !== "undefined" &&
                      typeof val === "string" &&
                      (opts || {}).type != "string" &&
                      (opts || {}).codepage
                    )
                      val = cptable.utils.decode(opts.codepage, val);
                    C_seen_K = true;
                    break;
                  case "E":
                    var formula = rc_to_a1(record[rj].slice(1), { r: R, c: C });
                    arr[R][C] = [arr[R][C], formula];
                    break;
                  default:
                    if (opts && opts.WTF)
                      throw new Error(`SYLK bad record ${rstr}`);
                }
              if (C_seen_K) {
                arr[R][C] = val;
                next_cell_format = null;
              }
              break;
            case "F":
              var F_seen = 0;
              for (rj = 1; rj < record.length; ++rj)
                switch (record[rj].charAt(0)) {
                  case "X":
                    C = parseInt(record[rj].slice(1)) - 1;
                    ++F_seen;
                    break;
                  case "Y":
                    R = parseInt(record[rj].slice(1)) - 1; /* C = 0; */
                    for (j = arr.length; j <= R; ++j) arr[j] = [];
                    break;
                  case "M":
                    Mval = parseInt(record[rj].slice(1)) / 20;
                    break;
                  case "F":
                    break; /* ??? */
                  case "G":
                    break; /* hide grid */
                  case "P":
                    next_cell_format = formats[parseInt(record[rj].slice(1))];
                    break;
                  case "S":
                    break; /* cell style */
                  case "D":
                    break; /* column */
                  case "N":
                    break; /* font */
                  case "W":
                    cw = record[rj].slice(1).split(" ");
                    for (
                      j = parseInt(cw[0], 10);
                      j <= parseInt(cw[1], 10);
                      ++j
                    ) {
                      Mval = parseInt(cw[2], 10);
                      colinfo[j - 1] =
                        Mval === 0 ? { hidden: true } : { wch: Mval };
                      process_col(colinfo[j - 1]);
                    }
                    break;
                  case "C" /* default column format */:
                    C = parseInt(record[rj].slice(1)) - 1;
                    if (!colinfo[C]) colinfo[C] = {};
                    break;
                  case "R" /* row properties */:
                    R = parseInt(record[rj].slice(1)) - 1;
                    if (!rowinfo[R]) rowinfo[R] = {};
                    if (Mval > 0) {
                      rowinfo[R].hpt = Mval;
                      rowinfo[R].hpx = pt2px(Mval);
                    } else if (Mval === 0) rowinfo[R].hidden = true;
                    break;
                  default:
                    if (opts && opts.WTF)
                      throw new Error(`SYLK bad record ${rstr}`);
                }
              if (F_seen < 1) next_cell_format = null;
              break;
            default:
              if (opts && opts.WTF) throw new Error(`SYLK bad record ${rstr}`);
          }
      }
      if (rowinfo.length > 0) sht["!rows"] = rowinfo;
      if (colinfo.length > 0) sht["!cols"] = colinfo;
      if (opts && opts.sheetRows) arr = arr.slice(0, opts.sheetRows);
      return [arr, sht];
    }

    function sylk_to_sheet(d, opts) {
      const aoasht = sylk_to_aoa(d, opts);
      const aoa = aoasht[0];
      const ws = aoasht[1];
      const o = aoa_to_sheet(aoa, opts);
      keys(ws).forEach(function(k) {
        o[k] = ws[k];
      });
      return o;
    }

    function sylk_to_workbook(d, opts) {
      return sheet_to_workbook(sylk_to_sheet(d, opts), opts);
    }

    function write_ws_cell_sylk(cell, ws, R, C) {
      let o = `C;Y${R + 1};X${C + 1};K`;
      switch (cell.t) {
        case "n":
          o += cell.v || 0;
          if (cell.f && !cell.F) o += `;E${a1_to_rc(cell.f, { r: R, c: C })}`;
          break;
        case "b":
          o += cell.v ? "TRUE" : "FALSE";
          break;
        case "e":
          o += cell.w || cell.v;
          break;
        case "d":
          o += `"${cell.w || cell.v}"`;
          break;
        case "s":
          o += `"${cell.v.replace(/"/g, "")}"`;
          break;
      }
      return o;
    }

    function write_ws_cols_sylk(out, cols) {
      cols.forEach(function(col, i) {
        let rec = `F;W${i + 1} ${i + 1} `;
        if (col.hidden) rec += "0";
        else {
          if (typeof col.width === "number") col.wpx = width2px(col.width);
          if (typeof col.wpx === "number") col.wch = px2char(col.wpx);
          if (typeof col.wch === "number") rec += Math.round(col.wch);
        }
        if (rec.charAt(rec.length - 1) != " ") out.push(rec);
      });
    }

    function write_ws_rows_sylk(out, rows) {
      rows.forEach(function(row, i) {
        let rec = "F;";
        if (row.hidden) rec += "M0;";
        else if (row.hpt) rec += `M${20 * row.hpt};`;
        else if (row.hpx) rec += `M${20 * px2pt(row.hpx)};`;
        if (rec.length > 2) out.push(`${rec}R${i + 1}`);
      });
    }

    function sheet_to_sylk(ws, opts) {
      const preamble = ["ID;PWXL;N;E"];
      const o = [];
      const r = safe_decode_range(ws["!ref"]);
      let cell;
      const dense = Array.isArray(ws);
      const RS = "\r\n";

      preamble.push("P;PGeneral");
      preamble.push("F;P0;DG0G8;M255");
      if (ws["!cols"]) write_ws_cols_sylk(preamble, ws["!cols"]);
      if (ws["!rows"]) write_ws_rows_sylk(preamble, ws["!rows"]);

      preamble.push(
        `B;Y${r.e.r - r.s.r + 1};X${r.e.c - r.s.c + 1};D${[
          r.s.c,
          r.s.r,
          r.e.c,
          r.e.r
        ].join(" ")}`
      );
      for (let R = r.s.r; R <= r.e.r; ++R) {
        for (let C = r.s.c; C <= r.e.c; ++C) {
          const coord = encode_cell({ r: R, c: C });
          cell = dense ? (ws[R] || [])[C] : ws[coord];
          if (!cell || (cell.v == null && (!cell.f || cell.F))) continue;
          o.push(write_ws_cell_sylk(cell, ws, R, C, opts));
        }
      }
      return `${preamble.join(RS) + RS + o.join(RS) + RS}E${RS}`;
    }

    return {
      to_workbook: sylk_to_workbook,
      to_sheet: sylk_to_sheet,
      from_sheet: sheet_to_sylk
    };
  })();

  const DIF = (function() {
    function dif_to_aoa(d, opts) {
      switch (opts.type) {
        case "base64":
          return dif_to_aoa_str(Base64.decode(d), opts);
        case "binary":
          return dif_to_aoa_str(d, opts);
        case "buffer":
          return dif_to_aoa_str(d.toString("binary"), opts);
        case "array":
          return dif_to_aoa_str(cc2str(d), opts);
      }
      throw new Error(`Unrecognized type ${opts.type}`);
    }
    function dif_to_aoa_str(str, opts) {
      const records = str.split("\n");
      let R = -1;
      let C = -1;
      let ri = 0;
      let arr = [];
      for (; ri !== records.length; ++ri) {
        if (records[ri].trim() === "BOT") {
          arr[++R] = [];
          C = 0;
          continue;
        }
        if (R < 0) continue;
        const metadata = records[ri].trim().split(",");
        const type = metadata[0];
        const value = metadata[1];
        ++ri;
        let data = records[ri].trim();
        switch (+type) {
          case -1:
            if (data === "BOT") {
              arr[++R] = [];
              C = 0;
              continue;
            } else if (data !== "EOD")
              throw new Error(`Unrecognized DIF special command ${data}`);
            break;
          case 0:
            if (data === "TRUE") arr[R][C] = true;
            else if (data === "FALSE") arr[R][C] = false;
            else if (!isNaN(fuzzynum(value))) arr[R][C] = fuzzynum(value);
            else if (!isNaN(fuzzydate(value).getDate()))
              arr[R][C] = parseDate(value);
            else arr[R][C] = value;
            ++C;
            break;
          case 1:
            data = data.slice(1, data.length - 1);
            arr[R][C++] = data !== "" ? data : null;
            break;
        }
        if (data === "EOD") break;
      }
      if (opts && opts.sheetRows) arr = arr.slice(0, opts.sheetRows);
      return arr;
    }

    function dif_to_sheet(str, opts) {
      return aoa_to_sheet(dif_to_aoa(str, opts), opts);
    }
    function dif_to_workbook(str, opts) {
      return sheet_to_workbook(dif_to_sheet(str, opts), opts);
    }

    const sheet_to_dif = (function() {
      const push_field = function pf(o, topic, v, n, s) {
        o.push(topic);
        o.push(`${v},${n}`);
        o.push(`"${s.replace(/"/g, '""')}"`);
      };
      const push_value = function po(o, type, v, s) {
        o.push(`${type},${v}`);
        o.push(type == 1 ? `"${s.replace(/"/g, '""')}"` : s);
      };
      return function sheet_to_dif(ws) {
        const o = [];
        const r = safe_decode_range(ws["!ref"]);
        let cell;
        const dense = Array.isArray(ws);
        push_field(o, "TABLE", 0, 1, "sheetjs");
        push_field(o, "VECTORS", 0, r.e.r - r.s.r + 1, "");
        push_field(o, "TUPLES", 0, r.e.c - r.s.c + 1, "");
        push_field(o, "DATA", 0, 0, "");
        for (let R = r.s.r; R <= r.e.r; ++R) {
          push_value(o, -1, 0, "BOT");
          for (let C = r.s.c; C <= r.e.c; ++C) {
            const coord = encode_cell({ r: R, c: C });
            cell = dense ? (ws[R] || [])[C] : ws[coord];
            if (!cell) {
              push_value(o, 1, 0, "");
              continue;
            }
            switch (cell.t) {
              case "n":
                var val = DIF_XL ? cell.w : cell.v;
                if (!val && cell.v != null) val = cell.v;
                if (val == null) {
                  if (DIF_XL && cell.f && !cell.F)
                    push_value(o, 1, 0, `=${cell.f}`);
                  else push_value(o, 1, 0, "");
                } else push_value(o, 0, val, "V");
                break;
              case "b":
                push_value(o, 0, cell.v ? 1 : 0, cell.v ? "TRUE" : "FALSE");
                break;
              case "s":
                push_value(
                  o,
                  1,
                  0,
                  !DIF_XL || isNaN(cell.v) ? cell.v : `="${cell.v}"`
                );
                break;
              case "d":
                if (!cell.w)
                  cell.w = SSF.format(
                    cell.z || SSF._table[14],
                    datenum(parseDate(cell.v))
                  );
                if (DIF_XL) push_value(o, 0, cell.w, "V");
                else push_value(o, 1, 0, cell.w);
                break;
              default:
                push_value(o, 1, 0, "");
            }
          }
        }
        push_value(o, -1, 0, "EOD");
        const RS = "\r\n";
        const oo = o.join(RS);
        // while((oo.length & 0x7F) != 0) oo += "\0";
        return oo;
      };
    })();
    return {
      to_workbook: dif_to_workbook,
      to_sheet: dif_to_sheet,
      from_sheet: sheet_to_dif
    };
  })();

  const ETH = (function() {
    function decode(s) {
      return s
        .replace(/\\b/g, "\\")
        .replace(/\\c/g, ":")
        .replace(/\\n/g, "\n");
    }
    function encode(s) {
      return s
        .replace(/\\/g, "\\b")
        .replace(/:/g, "\\c")
        .replace(/\n/g, "\\n");
    }

    function eth_to_aoa(str, opts) {
      const records = str.split("\n");
      let R = -1;
      let C = -1;
      let ri = 0;
      let arr = [];
      for (; ri !== records.length; ++ri) {
        const record = records[ri].trim().split(":");
        if (record[0] !== "cell") continue;
        const addr = decode_cell(record[1]);
        if (arr.length <= addr.r)
          for (R = arr.length; R <= addr.r; ++R) if (!arr[R]) arr[R] = [];
        R = addr.r;
        C = addr.c;
        switch (record[2]) {
          case "t":
            arr[R][C] = decode(record[3]);
            break;
          case "v":
            arr[R][C] = +record[3];
            break;
          case "vtf":
            var _f = record[record.length - 1];
          /* falls through */
          case "vtc":
            switch (record[3]) {
              case "nl":
                arr[R][C] = !!+record[4];
                break;
              default:
                arr[R][C] = +record[4];
                break;
            }
            if (record[2] == "vtf") arr[R][C] = [arr[R][C], _f];
        }
      }
      if (opts && opts.sheetRows) arr = arr.slice(0, opts.sheetRows);
      return arr;
    }

    function eth_to_sheet(d, opts) {
      return aoa_to_sheet(eth_to_aoa(d, opts), opts);
    }
    function eth_to_workbook(d, opts) {
      return sheet_to_workbook(eth_to_sheet(d, opts), opts);
    }

    const header = [
      "socialcalc:version:1.5",
      "MIME-Version: 1.0",
      "Content-Type: multipart/mixed; boundary=SocialCalcSpreadsheetControlSave"
    ].join("\n");

    const sep = `${[
      "--SocialCalcSpreadsheetControlSave",
      "Content-type: text/plain; charset=UTF-8"
    ].join("\n")}\n`;

    /* TODO: the other parts */
    const meta = ["# SocialCalc Spreadsheet Control Save", "part:sheet"].join(
      "\n"
    );

    const end = "--SocialCalcSpreadsheetControlSave--";

    function sheet_to_eth_data(ws) {
      if (!ws || !ws["!ref"]) return "";
      const o = [];
      let oo = [];
      let cell;
      let coord = "";
      const r = decode_range(ws["!ref"]);
      const dense = Array.isArray(ws);
      for (let R = r.s.r; R <= r.e.r; ++R) {
        for (let C = r.s.c; C <= r.e.c; ++C) {
          coord = encode_cell({ r: R, c: C });
          cell = dense ? (ws[R] || [])[C] : ws[coord];
          if (!cell || cell.v == null || cell.t === "z") continue;
          oo = ["cell", coord, "t"];
          switch (cell.t) {
            case "s":
            case "str":
              oo.push(encode(cell.v));
              break;
            case "n":
              if (!cell.f) {
                oo[2] = "v";
                oo[3] = cell.v;
              } else {
                oo[2] = "vtf";
                oo[3] = "n";
                oo[4] = cell.v;
                oo[5] = encode(cell.f);
              }
              break;
            case "b":
              oo[2] = `vt${cell.f ? "f" : "c"}`;
              oo[3] = "nl";
              oo[4] = cell.v ? "1" : "0";
              oo[5] = encode(cell.f || (cell.v ? "TRUE" : "FALSE"));
              break;
            case "d":
              var t = datenum(parseDate(cell.v));
              oo[2] = "vtc";
              oo[3] = "nd";
              oo[4] = `${t}`;
              oo[5] = cell.w || SSF.format(cell.z || SSF._table[14], t);
              break;
            case "e":
              continue;
          }
          o.push(oo.join(":"));
        }
      }
      o.push(`sheet:c:${r.e.c - r.s.c + 1}:r:${r.e.r - r.s.r + 1}:tvf:1`);
      o.push("valueformat:1:text-wiki");
      // o.push("copiedfrom:" + ws['!ref']); // clipboard only
      return o.join("\n");
    }

    function sheet_to_eth(ws) {
      return [header, sep, meta, sep, sheet_to_eth_data(ws), end].join("\n");
      // return ["version:1.5", sheet_to_eth_data(ws)].join("\n"); // clipboard form
    }

    return {
      to_workbook: eth_to_workbook,
      to_sheet: eth_to_sheet,
      from_sheet: sheet_to_eth
    };
  })();

  const PRN = (function() {
    function set_text_arr(data, arr, R, C, o) {
      if (o.raw) arr[R][C] = data;
      else if (data === "TRUE") arr[R][C] = true;
      else if (data === "FALSE") arr[R][C] = false;
      else if (data === "") {
        /* empty */
      } else if (!isNaN(fuzzynum(data))) arr[R][C] = fuzzynum(data);
      else if (!isNaN(fuzzydate(data).getDate())) arr[R][C] = parseDate(data);
      else arr[R][C] = data;
    }

    function prn_to_aoa_str(f, opts) {
      const o = opts || {};
      let arr = [];
      if (!f || f.length === 0) return arr;
      const lines = f.split(/[\r\n]/);
      let L = lines.length - 1;
      while (L >= 0 && lines[L].length === 0) --L;
      let start = 10;
      let idx = 0;
      let R = 0;
      for (; R <= L; ++R) {
        idx = lines[R].indexOf(" ");
        if (idx == -1) idx = lines[R].length;
        else idx++;
        start = Math.max(start, idx);
      }
      for (R = 0; R <= L; ++R) {
        arr[R] = [];
        /* TODO: confirm that widths are always 10 */
        let C = 0;
        set_text_arr(lines[R].slice(0, start).trim(), arr, R, C, o);
        for (C = 1; C <= (lines[R].length - start) / 10 + 1; ++C)
          set_text_arr(
            lines[R].slice(start + (C - 1) * 10, start + C * 10).trim(),
            arr,
            R,
            C,
            o
          );
      }
      if (o.sheetRows) arr = arr.slice(0, o.sheetRows);
      return arr;
    }

    // List of accepted CSV separators
    const guess_seps = {
      0x2c: ",",
      0x09: "\t",
      0x3b: ";"
    };

    // CSV separator weights to be used in case of equal numbers
    const guess_sep_weights = {
      0x2c: 3,
      0x09: 2,
      0x3b: 1
    };

    function guess_sep(str) {
      let cnt = {};
      let instr = false;
      let end = 0;
      let cc = 0;
      for (; end < str.length; ++end) {
        if ((cc = str.charCodeAt(end)) == 0x22) instr = !instr;
        else if (!instr && cc in guess_seps) cnt[cc] = (cnt[cc] || 0) + 1;
      }

      cc = [];
      for (end in cnt)
        if (Object.prototype.hasOwnProperty.call(cnt, end)) {
          cc.push([cnt[end], end]);
        }

      if (!cc.length) {
        cnt = guess_sep_weights;
        for (end in cnt)
          if (Object.prototype.hasOwnProperty.call(cnt, end)) {
            cc.push([cnt[end], end]);
          }
      }

      cc.sort(function(a, b) {
        return a[0] - b[0] || guess_sep_weights[a[1]] - guess_sep_weights[b[1]];
      });

      return guess_seps[cc.pop()[1]];
    }

    function dsv_to_sheet_str(str, opts) {
      const o = opts || {};
      let sep = "";
      if (DENSE != null && o.dense == null) o.dense = DENSE;
      const ws = o.dense ? [] : {};
      const range = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };

      if (str.slice(0, 4) == "sep=") {
        // If the line ends in \r\n
        if (str.charCodeAt(5) == 13 && str.charCodeAt(6) == 10) {
          sep = str.charAt(4);
          str = str.slice(7);
        }
        // If line ends in \r OR \n
        else if (str.charCodeAt(5) == 13 || str.charCodeAt(5) == 10) {
          //
          sep = str.charAt(4);
          str = str.slice(6);
        }
      } else sep = guess_sep(str.slice(0, 1024));
      let R = 0;
      let C = 0;
      let v = 0;
      let start = 0;
      let end = 0;
      const sepcc = sep.charCodeAt(0);
      let instr = false;
      let cc = 0;
      str = str.replace(/\r\n/gm, "\n");
      const _re = o.dateNF != null ? dateNF_regex(o.dateNF) : null;
      function finish_cell() {
        let s = str.slice(start, end);
        const cell = {};
        if (s.charAt(0) == '"' && s.charAt(s.length - 1) == '"')
          s = s.slice(1, -1).replace(/""/g, '"');
        if (s.length === 0) cell.t = "z";
        else if (o.raw) {
          cell.t = "s";
          cell.v = s;
        } else if (s.trim().length === 0) {
          cell.t = "s";
          cell.v = s;
        } else if (s.charCodeAt(0) == 0x3d) {
          if (s.charCodeAt(1) == 0x22 && s.charCodeAt(s.length - 1) == 0x22) {
            cell.t = "s";
            cell.v = s.slice(2, -1).replace(/""/g, '"');
          } else if (fuzzyfmla(s)) {
            cell.t = "n";
            cell.f = s.slice(1);
          } else {
            cell.t = "s";
            cell.v = s;
          }
        } else if (s == "TRUE") {
          cell.t = "b";
          cell.v = true;
        } else if (s == "FALSE") {
          cell.t = "b";
          cell.v = false;
        } else if (!isNaN((v = fuzzynum(s)))) {
          cell.t = "n";
          if (o.cellText !== false) cell.w = s;
          cell.v = v;
        } else if (!isNaN(fuzzydate(s).getDate()) || (_re && s.match(_re))) {
          cell.z = o.dateNF || SSF._table[14];
          let k = 0;
          if (_re && s.match(_re)) {
            s = dateNF_fix(s, o.dateNF, s.match(_re) || []);
            k = 1;
          }
          if (o.cellDates) {
            cell.t = "d";
            cell.v = parseDate(s, k);
          } else {
            cell.t = "n";
            cell.v = datenum(parseDate(s, k));
          }
          if (o.cellText !== false)
            cell.w = SSF.format(
              cell.z,
              cell.v instanceof Date ? datenum(cell.v) : cell.v
            );
          if (!o.cellNF) delete cell.z;
        } else {
          cell.t = "s";
          cell.v = s;
        }
        if (cell.t == "z") {
        } else if (o.dense) {
          if (!ws[R]) ws[R] = [];
          ws[R][C] = cell;
        } else ws[encode_cell({ c: C, r: R })] = cell;
        start = end + 1;
        if (range.e.c < C) range.e.c = C;
        if (range.e.r < R) range.e.r = R;
        if (cc == sepcc) ++C;
        else {
          C = 0;
          ++R;
          if (o.sheetRows && o.sheetRows <= R) return true;
        }
      }
      outer: for (; end < str.length; ++end)
        switch ((cc = str.charCodeAt(end))) {
          case 0x22:
            instr = !instr;
            break;
          case sepcc:
          case 0x0a:
          case 0x0d:
            if (!instr && finish_cell()) break outer;
            break;
          default:
            break;
        }
      if (end - start > 0) finish_cell();

      ws["!ref"] = encode_range(range);
      return ws;
    }

    function prn_to_sheet_str(str, opts) {
      if (!(opts && opts.PRN)) return dsv_to_sheet_str(str, opts);
      if (str.slice(0, 4) == "sep=") return dsv_to_sheet_str(str, opts);
      if (
        str.indexOf("\t") >= 0 ||
        str.indexOf(",") >= 0 ||
        str.indexOf(";") >= 0
      )
        return dsv_to_sheet_str(str, opts);
      return aoa_to_sheet(prn_to_aoa_str(str, opts), opts);
    }

    function prn_to_sheet(d, opts) {
      let str = "";
      const bytes = opts.type == "string" ? [0, 0, 0, 0] : firstbyte(d, opts);
      switch (opts.type) {
        case "base64":
          str = Base64.decode(d);
          break;
        case "binary":
          str = d;
          break;
        case "buffer":
          if (opts.codepage == 65001) str = d.toString("utf8");
          else if (opts.codepage && typeof cptable !== "undefined")
            str = cptable.utils.decode(opts.codepage, d);
          else str = d.toString("binary");
          break;
        case "array":
          str = cc2str(d);
          break;
        case "string":
          str = d;
          break;
        default:
          throw new Error(`Unrecognized type ${opts.type}`);
      }
      if (bytes[0] == 0xef && bytes[1] == 0xbb && bytes[2] == 0xbf)
        str = utf8read(str.slice(3));
      else if (
        opts.type == "binary" &&
        typeof cptable !== "undefined" &&
        opts.codepage
      )
        str = cptable.utils.decode(
          opts.codepage,
          cptable.utils.encode(1252, str)
        );
      if (str.slice(0, 19) == "socialcalc:version:")
        return ETH.to_sheet(opts.type == "string" ? str : utf8read(str), opts);
      return prn_to_sheet_str(str, opts);
    }

    function prn_to_workbook(d, opts) {
      return sheet_to_workbook(prn_to_sheet(d, opts), opts);
    }

    function sheet_to_prn(ws) {
      const o = [];
      const r = safe_decode_range(ws["!ref"]);
      let cell;
      const dense = Array.isArray(ws);
      for (let R = r.s.r; R <= r.e.r; ++R) {
        const oo = [];
        for (let C = r.s.c; C <= r.e.c; ++C) {
          const coord = encode_cell({ r: R, c: C });
          cell = dense ? (ws[R] || [])[C] : ws[coord];
          if (!cell || cell.v == null) {
            oo.push("          ");
            continue;
          }
          let w = (cell.w || (format_cell(cell), cell.w) || "").slice(0, 10);
          while (w.length < 10) w += " ";
          oo.push(w + (C === 0 ? " " : ""));
        }
        o.push(oo.join(""));
      }
      return o.join("\n");
    }

    return {
      to_workbook: prn_to_workbook,
      to_sheet: prn_to_sheet,
      from_sheet: sheet_to_prn
    };
  })();

  /* Excel defaults to SYLK but warns if data is not valid */
  function read_wb_ID(d, opts) {
    const o = opts || {};
    const OLD_WTF = !!o.WTF;
    o.WTF = true;
    try {
      const out = SYLK.to_workbook(d, o);
      o.WTF = OLD_WTF;
      return out;
    } catch (e) {
      o.WTF = OLD_WTF;
      if (!e.message.match(/SYLK bad record ID/) && OLD_WTF) throw e;
      return PRN.to_workbook(d, opts);
    }
  }

  const WK_ = (function() {
    function lotushopper(data, cb, opts) {
      if (!data) return;
      prep_blob(data, data.l || 0);
      const Enum = opts.Enum || WK1Enum;
      while (data.l < data.length) {
        const RT = data.read_shift(2);
        const R = Enum[RT] || Enum[0xff];
        const length = data.read_shift(2);
        const tgt = data.l + length;
        const d = (R.f || parsenoop)(data, length, opts);
        data.l = tgt;
        if (cb(d, R.n, RT)) return;
      }
    }

    function lotus_to_workbook(d, opts) {
      switch (opts.type) {
        case "base64":
          return lotus_to_workbook_buf(s2a(Base64.decode(d)), opts);
        case "binary":
          return lotus_to_workbook_buf(s2a(d), opts);
        case "buffer":
        case "array":
          return lotus_to_workbook_buf(d, opts);
      }
      throw `Unsupported type ${opts.type}`;
    }

    function lotus_to_workbook_buf(d, opts) {
      if (!d) return d;
      const o = opts || {};
      if (DENSE != null && o.dense == null) o.dense = DENSE;
      let s = o.dense ? [] : {};
      let n = "Sheet1";
      let sidx = 0;
      const sheets = {};
      const snames = [n];

      let refguess = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
      const sheetRows = o.sheetRows || 0;

      if (d[2] == 0x02) o.Enum = WK1Enum;
      else if (d[2] == 0x1a) o.Enum = WK3Enum;
      else if (d[2] == 0x0e) {
        o.Enum = WK3Enum;
        o.qpro = true;
        d.l = 0;
      } else throw new Error(`Unrecognized LOTUS BOF ${d[2]}`);
      lotushopper(
        d,
        function(val, Rn, RT) {
          if (d[2] == 0x02)
            switch (RT) {
              case 0x00:
                o.vers = val;
                if (val >= 0x1000) o.qpro = true;
                break;
              case 0x06:
                refguess = val;
                break; /* RANGE */
              case 0x0f /* LABEL */:
                if (!o.qpro) val[1].v = val[1].v.slice(1);
              /* falls through */
              case 0x0d: /* INTEGER */
              case 0x0e: /* NUMBER */
              case 0x10: /* FORMULA */
              case 0x33 /* STRING */:
                /* TODO: actual translation of the format code */
                if (
                  RT == 0x0e &&
                  (val[2] & 0x70) == 0x70 &&
                  (val[2] & 0x0f) > 1 &&
                  (val[2] & 0x0f) < 15
                ) {
                  val[1].z = o.dateNF || SSF._table[14];
                  if (o.cellDates) {
                    val[1].t = "d";
                    val[1].v = numdate(val[1].v);
                  }
                }
                if (o.dense) {
                  if (!s[val[0].r]) s[val[0].r] = [];
                  s[val[0].r][val[0].c] = val[1];
                } else s[encode_cell(val[0])] = val[1];
                break;
            }
          else
            switch (RT) {
              case 0x16 /* LABEL16 */:
                val[1].v = val[1].v.slice(1);
              /* falls through */
              case 0x17: /* NUMBER17 */
              case 0x18: /* NUMBER18 */
              case 0x19: /* FORMULA19 */
              case 0x25: /* NUMBER25 */
              case 0x27: /* NUMBER27 */
              case 0x28 /* FORMULA28 */:
                if (val[3] > sidx) {
                  s["!ref"] = encode_range(refguess);
                  sheets[n] = s;
                  s = o.dense ? [] : {};
                  refguess = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
                  sidx = val[3];
                  n = `Sheet${sidx + 1}`;
                  snames.push(n);
                }
                if (sheetRows > 0 && val[0].r >= sheetRows) break;
                if (o.dense) {
                  if (!s[val[0].r]) s[val[0].r] = [];
                  s[val[0].r][val[0].c] = val[1];
                } else s[encode_cell(val[0])] = val[1];
                if (refguess.e.c < val[0].c) refguess.e.c = val[0].c;
                if (refguess.e.r < val[0].r) refguess.e.r = val[0].r;
                break;
              default:
                break;
            }
        },
        o
      );

      s["!ref"] = encode_range(refguess);
      sheets[n] = s;
      return { SheetNames: snames, Sheets: sheets };
    }

    function parse_RANGE(blob) {
      const o = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };
      o.s.c = blob.read_shift(2);
      o.s.r = blob.read_shift(2);
      o.e.c = blob.read_shift(2);
      o.e.r = blob.read_shift(2);
      if (o.s.c == 0xffff) o.s.c = o.e.c = o.s.r = o.e.r = 0;
      return o;
    }

    function parse_cell(blob, length, opts) {
      const o = [{ c: 0, r: 0 }, { t: "n", v: 0 }, 0];
      if (opts.qpro && opts.vers != 0x5120) {
        o[0].c = blob.read_shift(1);
        blob.l++;
        o[0].r = blob.read_shift(2);
        blob.l += 2;
      } else {
        o[2] = blob.read_shift(1);
        o[0].c = blob.read_shift(2);
        o[0].r = blob.read_shift(2);
      }
      return o;
    }

    function parse_LABEL(blob, length, opts) {
      const tgt = blob.l + length;
      const o = parse_cell(blob, length, opts);
      o[1].t = "s";
      if (opts.vers == 0x5120) {
        blob.l++;
        const len = blob.read_shift(1);
        o[1].v = blob.read_shift(len, "utf8");
        return o;
      }
      if (opts.qpro) blob.l++;
      o[1].v = blob.read_shift(tgt - blob.l, "cstr");
      return o;
    }

    function parse_INTEGER(blob, length, opts) {
      const o = parse_cell(blob, length, opts);
      o[1].v = blob.read_shift(2, "i");
      return o;
    }

    function parse_NUMBER(blob, length, opts) {
      const o = parse_cell(blob, length, opts);
      o[1].v = blob.read_shift(8, "f");
      return o;
    }

    function parse_FORMULA(blob, length, opts) {
      const tgt = blob.l + length;
      const o = parse_cell(blob, length, opts);
      /* TODO: formula */
      o[1].v = blob.read_shift(8, "f");
      if (opts.qpro) blob.l = tgt;
      else {
        const flen = blob.read_shift(2);
        blob.l += flen;
      }
      return o;
    }

    function parse_cell_3(blob) {
      const o = [{ c: 0, r: 0 }, { t: "n", v: 0 }, 0];
      o[0].r = blob.read_shift(2);
      o[3] = blob[blob.l++];
      o[0].c = blob[blob.l++];
      return o;
    }

    function parse_LABEL_16(blob, length) {
      const o = parse_cell_3(blob, length);
      o[1].t = "s";
      o[1].v = blob.read_shift(length - 4, "cstr");
      return o;
    }

    function parse_NUMBER_18(blob, length) {
      const o = parse_cell_3(blob, length);
      o[1].v = blob.read_shift(2);
      let v = o[1].v >> 1;
      /* TODO: figure out all of the corner cases */
      if (o[1].v & 0x1) {
        switch (v & 0x07) {
          case 1:
            v = (v >> 3) * 500;
            break;
          case 2:
            v = (v >> 3) / 20;
            break;
          case 4:
            v = (v >> 3) / 2000;
            break;
          case 6:
            v = (v >> 3) / 16;
            break;
          case 7:
            v = (v >> 3) / 64;
            break;
          default:
            throw `unknown NUMBER_18 encoding ${v & 0x07}`;
        }
      }
      o[1].v = v;
      return o;
    }

    function parse_NUMBER_17(blob, length) {
      const o = parse_cell_3(blob, length);
      const v1 = blob.read_shift(4);
      const v2 = blob.read_shift(4);
      let e = blob.read_shift(2);
      if (e == 0xffff) {
        o[1].v = 0;
        return o;
      }
      const s = e & 0x8000;
      e = (e & 0x7fff) - 16446;
      o[1].v =
        (s * 2 - 1) *
        ((e > 0 ? v2 << e : v2 >>> -e) +
          (e > -32 ? v1 << (e + 32) : v1 >>> -(e + 32)));
      return o;
    }

    function parse_FORMULA_19(blob, length) {
      const o = parse_NUMBER_17(blob, 14);
      blob.l += length - 14; /* TODO: formula */
      return o;
    }

    function parse_NUMBER_25(blob, length) {
      const o = parse_cell_3(blob, length);
      const v1 = blob.read_shift(4);
      o[1].v = v1 >> 6;
      return o;
    }

    function parse_NUMBER_27(blob, length) {
      const o = parse_cell_3(blob, length);
      const v1 = blob.read_shift(8, "f");
      o[1].v = v1;
      return o;
    }

    function parse_FORMULA_28(blob, length) {
      const o = parse_NUMBER_27(blob, 14);
      blob.l += length - 10; /* TODO: formula */
      return o;
    }

    var WK1Enum = {
      0x0000: { n: "BOF", f: parseuint16 },
      0x0001: { n: "EOF" },
      0x0002: { n: "CALCMODE" },
      0x0003: { n: "CALCORDER" },
      0x0004: { n: "SPLIT" },
      0x0005: { n: "SYNC" },
      0x0006: { n: "RANGE", f: parse_RANGE },
      0x0007: { n: "WINDOW1" },
      0x0008: { n: "COLW1" },
      0x0009: { n: "WINTWO" },
      0x000a: { n: "COLW2" },
      0x000b: { n: "NAME" },
      0x000c: { n: "BLANK" },
      0x000d: { n: "INTEGER", f: parse_INTEGER },
      0x000e: { n: "NUMBER", f: parse_NUMBER },
      0x000f: { n: "LABEL", f: parse_LABEL },
      0x0010: { n: "FORMULA", f: parse_FORMULA },
      0x0018: { n: "TABLE" },
      0x0019: { n: "ORANGE" },
      0x001a: { n: "PRANGE" },
      0x001b: { n: "SRANGE" },
      0x001c: { n: "FRANGE" },
      0x001d: { n: "KRANGE1" },
      0x0020: { n: "HRANGE" },
      0x0023: { n: "KRANGE2" },
      0x0024: { n: "PROTEC" },
      0x0025: { n: "FOOTER" },
      0x0026: { n: "HEADER" },
      0x0027: { n: "SETUP" },
      0x0028: { n: "MARGINS" },
      0x0029: { n: "LABELFMT" },
      0x002a: { n: "TITLES" },
      0x002b: { n: "SHEETJS" },
      0x002d: { n: "GRAPH" },
      0x002e: { n: "NGRAPH" },
      0x002f: { n: "CALCCOUNT" },
      0x0030: { n: "UNFORMATTED" },
      0x0031: { n: "CURSORW12" },
      0x0032: { n: "WINDOW" },
      0x0033: { n: "STRING", f: parse_LABEL },
      0x0037: { n: "PASSWORD" },
      0x0038: { n: "LOCKED" },
      0x003c: { n: "QUERY" },
      0x003d: { n: "QUERYNAME" },
      0x003e: { n: "PRINT" },
      0x003f: { n: "PRINTNAME" },
      0x0040: { n: "GRAPH2" },
      0x0041: { n: "GRAPHNAME" },
      0x0042: { n: "ZOOM" },
      0x0043: { n: "SYMSPLIT" },
      0x0044: { n: "NSROWS" },
      0x0045: { n: "NSCOLS" },
      0x0046: { n: "RULER" },
      0x0047: { n: "NNAME" },
      0x0048: { n: "ACOMM" },
      0x0049: { n: "AMACRO" },
      0x004a: { n: "PARSE" },
      0x00ff: { n: "", f: parsenoop }
    };

    var WK3Enum = {
      0x0000: { n: "BOF" },
      0x0001: { n: "EOF" },
      0x0003: { n: "??" },
      0x0004: { n: "??" },
      0x0005: { n: "??" },
      0x0006: { n: "??" },
      0x0007: { n: "??" },
      0x0009: { n: "??" },
      0x000a: { n: "??" },
      0x000b: { n: "??" },
      0x000c: { n: "??" },
      0x000e: { n: "??" },
      0x000f: { n: "??" },
      0x0010: { n: "??" },
      0x0011: { n: "??" },
      0x0012: { n: "??" },
      0x0013: { n: "??" },
      0x0015: { n: "??" },
      0x0016: { n: "LABEL16", f: parse_LABEL_16 },
      0x0017: { n: "NUMBER17", f: parse_NUMBER_17 },
      0x0018: { n: "NUMBER18", f: parse_NUMBER_18 },
      0x0019: { n: "FORMULA19", f: parse_FORMULA_19 },
      0x001a: { n: "??" },
      0x001b: { n: "??" },
      0x001c: { n: "??" },
      0x001d: { n: "??" },
      0x001e: { n: "??" },
      0x001f: { n: "??" },
      0x0021: { n: "??" },
      0x0025: { n: "NUMBER25", f: parse_NUMBER_25 },
      0x0027: { n: "NUMBER27", f: parse_NUMBER_27 },
      0x0028: { n: "FORMULA28", f: parse_FORMULA_28 },
      0x00ff: { n: "", f: parsenoop }
    };
    return {
      to_workbook: lotus_to_workbook
    };
  })();
  /* 18.4.7 rPr CT_RPrElt */
  function parse_rpr(rpr) {
    const font = {};
    const m = rpr.match(tagregex);
    let i = 0;
    let pass = false;
    if (m)
      for (; i != m.length; ++i) {
        const y = parsexmltag(m[i]);
        switch (y[0].replace(/\w*:/g, "")) {
          /* 18.8.12 condense CT_BooleanProperty */
          /* ** not required . */
          case "<condense":
            break;
          /* 18.8.17 extend CT_BooleanProperty */
          /* ** not required . */
          case "<extend":
            break;
          /* 18.8.36 shadow CT_BooleanProperty */
          /* ** not required . */
          case "<shadow":
            if (!y.val) break;
          /* falls through */
          case "<shadow>":
          case "<shadow/>":
            font.shadow = 1;
            break;
          case "</shadow>":
            break;

          /* 18.4.1 charset CT_IntProperty TODO */
          case "<charset":
            if (y.val == "1") break;
            font.cp = CS2CP[parseInt(y.val, 10)];
            break;

          /* 18.4.2 outline CT_BooleanProperty TODO */
          case "<outline":
            if (!y.val) break;
          /* falls through */
          case "<outline>":
          case "<outline/>":
            font.outline = 1;
            break;
          case "</outline>":
            break;

          /* 18.4.5 rFont CT_FontName */
          case "<rFont":
            font.name = y.val;
            break;

          /* 18.4.11 sz CT_FontSize */
          case "<sz":
            font.sz = y.val;
            break;

          /* 18.4.10 strike CT_BooleanProperty */
          case "<strike":
            if (!y.val) break;
          /* falls through */
          case "<strike>":
          case "<strike/>":
            font.strike = 1;
            break;
          case "</strike>":
            break;

          /* 18.4.13 u CT_UnderlineProperty */
          case "<u":
            if (!y.val) break;
            switch (y.val) {
              case "double":
                font.uval = "double";
                break;
              case "singleAccounting":
                font.uval = "single-accounting";
                break;
              case "doubleAccounting":
                font.uval = "double-accounting";
                break;
            }
          /* falls through */
          case "<u>":
          case "<u/>":
            font.u = 1;
            break;
          case "</u>":
            break;

          /* 18.8.2 b */
          case "<b":
            if (y.val == "0") break;
          /* falls through */
          case "<b>":
          case "<b/>":
            font.b = 1;
            break;
          case "</b>":
            break;

          /* 18.8.26 i */
          case "<i":
            if (y.val == "0") break;
          /* falls through */
          case "<i>":
          case "<i/>":
            font.i = 1;
            break;
          case "</i>":
            break;

          /* 18.3.1.15 color CT_Color TODO: tint, theme, auto, indexed */
          case "<color":
            if (y.rgb) font.color = y.rgb.slice(2, 8);
            break;

          /* 18.8.18 family ST_FontFamily */
          case "<family":
            font.family = y.val;
            break;

          /* 18.4.14 vertAlign CT_VerticalAlignFontProperty TODO */
          case "<vertAlign":
            font.valign = y.val;
            break;

          /* 18.8.35 scheme CT_FontScheme TODO */
          case "<scheme":
            break;

          /* 18.2.10 extLst CT_ExtensionList ? */
          case "<extLst":
          case "<extLst>":
          case "</extLst>":
            break;
          case "<ext":
            pass = true;
            break;
          case "</ext>":
            pass = false;
            break;
          default:
            if (y[0].charCodeAt(1) !== 47 && !pass)
              throw new Error(`Unrecognized rich format ${y[0]}`);
        }
      }
    return font;
  }

  const parse_rs = (function() {
    const tregex = matchtag("t");
    const rpregex = matchtag("rPr");
    /* 18.4.4 r CT_RElt */
    function parse_r(r) {
      /* 18.4.12 t ST_Xstring */
      const t = r.match(tregex); /* , cp = 65001 */
      if (!t) return { t: "s", v: "" };

      const o = { t: "s", v: unescapexml(t[1]) };
      const rpr = r.match(rpregex);
      if (rpr) o.s = parse_rpr(rpr[1]);
      return o;
    }
    const rregex = /<(?:\w+:)?r>/g;
    const rend = /<\/(?:\w+:)?r>/;
    return function parse_rs(rs) {
      return rs
        .replace(rregex, "")
        .split(rend)
        .map(parse_r)
        .filter(function(r) {
          return r.v;
        });
    };
  })();

  /* Parse a list of <r> tags */
  const rs_to_html = (function parse_rs_factory() {
    const nlregex = /(\r\n|\n)/g;
    function parse_rpr2(font, intro, outro) {
      const style = [];

      if (font.u) style.push("text-decoration: underline;");
      if (font.uval) style.push(`text-underline-style:${font.uval};`);
      if (font.sz) style.push(`font-size:${font.sz}pt;`);
      if (font.outline) style.push("text-effect: outline;");
      if (font.shadow) style.push("text-shadow: auto;");
      intro.push(`<span style="${style.join("")}">`);

      if (font.b) {
        intro.push("<b>");
        outro.push("</b>");
      }
      if (font.i) {
        intro.push("<i>");
        outro.push("</i>");
      }
      if (font.strike) {
        intro.push("<s>");
        outro.push("</s>");
      }

      let align = font.valign || "";
      if (align == "superscript" || align == "super") align = "sup";
      else if (align == "subscript") align = "sub";
      if (align != "") {
        intro.push(`<${align}>`);
        outro.push(`</${align}>`);
      }

      outro.push("</span>");
      return font;
    }

    /* 18.4.4 r CT_RElt */
    function r_to_html(r) {
      const terms = [[], r.v, []];
      if (!r.v) return "";

      if (r.s) parse_rpr2(r.s, terms[0], terms[2]);

      return (
        terms[0].join("") +
        terms[1].replace(nlregex, "<br/>") +
        terms[2].join("")
      );
    }

    return function parse_rs(rs) {
      return rs.map(r_to_html).join("");
    };
  })();

  /* 18.4.8 si CT_Rst */
  const sitregex = /<(?:\w+:)?t[^>]*>([^<]*)<\/(?:\w+:)?t>/g;
  const sirregex = /<(?:\w+:)?r>/;
  const sirphregex = /<(?:\w+:)?rPh.*?>([\s\S]*?)<\/(?:\w+:)?rPh>/g;
  function parse_si(x, opts) {
    const html = opts ? opts.cellHTML : true;
    const z = {};
    if (!x) return { t: "" };
    // var y;
    /* 18.4.12 t ST_Xstring (Plaintext String) */
    // TODO: is whitespace actually valid here?
    if (x.match(/^\s*<(?:\w+:)?t[^>]*>/)) {
      z.t = unescapexml(
        utf8read(x.slice(x.indexOf(">") + 1).split(/<\/(?:\w+:)?t>/)[0] || "")
      );
      z.r = utf8read(x);
      if (html) z.h = escapehtml(z.t);
    } else if (/* y = */ x.match(sirregex)) {
      /* 18.4.4 r CT_RElt (Rich Text Run) */
      z.r = utf8read(x);
      z.t = unescapexml(
        utf8read(
          (x.replace(sirphregex, "").match(sitregex) || [])
            .join("")
            .replace(tagregex, "")
        )
      );
      if (html) z.h = rs_to_html(parse_rs(z.r));
    }
    /* 18.4.3 phoneticPr CT_PhoneticPr (TODO: needed for Asian support) */
    /* 18.4.6 rPh CT_PhoneticRun (TODO: needed for Asian support) */
    return z;
  }

  /* 18.4 Shared String Table */
  const sstr0 = /<(?:\w+:)?sst([^>]*)>([\s\S]*)<\/(?:\w+:)?sst>/;
  const sstr1 = /<(?:\w+:)?(?:si|sstItem)>/g;
  const sstr2 = /<\/(?:\w+:)?(?:si|sstItem)>/;
  function parse_sst_xml(data, opts) {
    const s = [];
    let ss = "";
    if (!data) return s;
    /* 18.4.9 sst CT_Sst */
    let sst = data.match(sstr0);
    if (sst) {
      ss = sst[2].replace(sstr1, "").split(sstr2);
      for (let i = 0; i != ss.length; ++i) {
        const o = parse_si(ss[i].trim(), opts);
        if (o != null) s[s.length] = o;
      }
      sst = parsexmltag(sst[1]);
      s.Count = sst.count;
      s.Unique = sst.uniqueCount;
    }
    return s;
  }

  RELS.SST =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
  /* [MS-XLSB] 2.4.221 BrtBeginSst */
  function parse_BrtBeginSst(data) {
    return [data.read_shift(4), data.read_shift(4)];
  }

  /* [MS-XLSB] 2.1.7.45 Shared Strings */
  function parse_sst_bin(data, opts) {
    const s = [];
    let pass = false;
    recordhopper(data, function hopper_sst(val, R_n, RT) {
      switch (RT) {
        case 0x009f /* 'BrtBeginSst' */:
          s.Count = val[0];
          s.Unique = val[1];
          break;
        case 0x0013 /* 'BrtSSTItem' */:
          s.push(val);
          break;
        case 0x00a0 /* 'BrtEndSst' */:
          return true;

        case 0x0023 /* 'BrtFRTBegin' */:
          pass = true;
          break;
        case 0x0024 /* 'BrtFRTEnd' */:
          pass = false;
          break;

        default:
          if (R_n.indexOf("Begin") > 0) {
            /* empty */
          } else if (R_n.indexOf("End") > 0) {
            /* empty */
          }
          if (!pass || opts.WTF)
            throw new Error(`Unexpected record ${RT} ${R_n}`);
      }
    });
    return s;
  }

  function _JS2ANSI(str) {
    if (typeof cptable !== "undefined")
      return cptable.utils.encode(current_ansi, str);
    const o = [];
    const oo = str.split("");
    for (let i = 0; i < oo.length; ++i) o[i] = oo[i].charCodeAt(0);
    return o;
  }

  /* [MS-OFFCRYPTO] 2.1.4 Version */
  function parse_CRYPTOVersion(blob, length) {
    const o = {};
    o.Major = blob.read_shift(2);
    o.Minor = blob.read_shift(2);
    if (length >= 4) blob.l += length - 4;
    return o;
  }

  /* [MS-OFFCRYPTO] 2.1.5 DataSpaceVersionInfo */
  function parse_DataSpaceVersionInfo(blob) {
    const o = {};
    o.id = blob.read_shift(0, "lpp4");
    o.R = parse_CRYPTOVersion(blob, 4);
    o.U = parse_CRYPTOVersion(blob, 4);
    o.W = parse_CRYPTOVersion(blob, 4);
    return o;
  }

  /* [MS-OFFCRYPTO] 2.1.6.1 DataSpaceMapEntry Structure */
  function parse_DataSpaceMapEntry(blob) {
    const len = blob.read_shift(4);
    const end = blob.l + len - 4;
    const o = {};
    let cnt = blob.read_shift(4);
    const comps = [];
    /* [MS-OFFCRYPTO] 2.1.6.2 DataSpaceReferenceComponent Structure */
    while (cnt-- > 0)
      comps.push({ t: blob.read_shift(4), v: blob.read_shift(0, "lpp4") });
    o.name = blob.read_shift(0, "lpp4");
    o.comps = comps;
    if (blob.l != end)
      throw new Error(`Bad DataSpaceMapEntry: ${blob.l} != ${end}`);
    return o;
  }

  /* [MS-OFFCRYPTO] 2.1.6 DataSpaceMap */
  function parse_DataSpaceMap(blob) {
    const o = [];
    blob.l += 4; // must be 0x8
    let cnt = blob.read_shift(4);
    while (cnt-- > 0) o.push(parse_DataSpaceMapEntry(blob));
    return o;
  }

  /* [MS-OFFCRYPTO] 2.1.7 DataSpaceDefinition */
  function parse_DataSpaceDefinition(blob) {
    const o = [];
    blob.l += 4; // must be 0x8
    let cnt = blob.read_shift(4);
    while (cnt-- > 0) o.push(blob.read_shift(0, "lpp4"));
    return o;
  }

  /* [MS-OFFCRYPTO] 2.1.8 DataSpaceDefinition */
  function parse_TransformInfoHeader(blob) {
    const o = {};
    /* var len = */ blob.read_shift(4);
    blob.l += 4; // must be 0x1
    o.id = blob.read_shift(0, "lpp4");
    o.name = blob.read_shift(0, "lpp4");
    o.R = parse_CRYPTOVersion(blob, 4);
    o.U = parse_CRYPTOVersion(blob, 4);
    o.W = parse_CRYPTOVersion(blob, 4);
    return o;
  }

  function parse_Primary(blob) {
    /* [MS-OFFCRYPTO] 2.2.6 IRMDSTransformInfo */
    const hdr = parse_TransformInfoHeader(blob);
    /* [MS-OFFCRYPTO] 2.1.9 EncryptionTransformInfo */
    hdr.ename = blob.read_shift(0, "8lpp4");
    hdr.blksz = blob.read_shift(4);
    hdr.cmode = blob.read_shift(4);
    if (blob.read_shift(4) != 0x04) throw new Error("Bad !Primary record");
    return hdr;
  }

  /* [MS-OFFCRYPTO] 2.3.2 Encryption Header */
  function parse_EncryptionHeader(blob, length) {
    const tgt = blob.l + length;
    const o = {};
    o.Flags = blob.read_shift(4) & 0x3f;
    blob.l += 4;
    o.AlgID = blob.read_shift(4);
    let valid = false;
    switch (o.AlgID) {
      case 0x660e:
      case 0x660f:
      case 0x6610:
        valid = o.Flags == 0x24;
        break;
      case 0x6801:
        valid = o.Flags == 0x04;
        break;
      case 0:
        valid = o.Flags == 0x10 || o.Flags == 0x04 || o.Flags == 0x24;
        break;
      default:
        throw `Unrecognized encryption algorithm: ${o.AlgID}`;
    }
    if (!valid) throw new Error("Encryption Flags/AlgID mismatch");
    o.AlgIDHash = blob.read_shift(4);
    o.KeySize = blob.read_shift(4);
    o.ProviderType = blob.read_shift(4);
    blob.l += 8;
    o.CSPName = blob.read_shift((tgt - blob.l) >> 1, "utf16le");
    blob.l = tgt;
    return o;
  }

  /* [MS-OFFCRYPTO] 2.3.3 Encryption Verifier */
  function parse_EncryptionVerifier(blob, length) {
    const o = {};
    const tgt = blob.l + length;
    blob.l += 4; // SaltSize must be 0x10
    o.Salt = blob.slice(blob.l, blob.l + 16);
    blob.l += 16;
    o.Verifier = blob.slice(blob.l, blob.l + 16);
    blob.l += 16;
    /* var sz = */ blob.read_shift(4);
    o.VerifierHash = blob.slice(blob.l, tgt);
    blob.l = tgt;
    return o;
  }

  /* [MS-OFFCRYPTO] 2.3.4.* EncryptionInfo Stream */
  function parse_EncryptionInfo(blob) {
    const vers = parse_CRYPTOVersion(blob);
    switch (vers.Minor) {
      case 0x02:
        return [vers.Minor, parse_EncInfoStd(blob, vers)];
      case 0x03:
        return [vers.Minor, parse_EncInfoExt(blob, vers)];
      case 0x04:
        return [vers.Minor, parse_EncInfoAgl(blob, vers)];
    }
    throw new Error(
      `ECMA-376 Encrypted file unrecognized Version: ${vers.Minor}`
    );
  }

  /* [MS-OFFCRYPTO] 2.3.4.5  EncryptionInfo Stream (Standard Encryption) */
  function parse_EncInfoStd(blob) {
    const flags = blob.read_shift(4);
    if ((flags & 0x3f) != 0x24) throw new Error("EncryptionInfo mismatch");
    const sz = blob.read_shift(4);
    // var tgt = blob.l + sz;
    const hdr = parse_EncryptionHeader(blob, sz);
    const verifier = parse_EncryptionVerifier(blob, blob.length - blob.l);
    return { t: "Std", h: hdr, v: verifier };
  }
  /* [MS-OFFCRYPTO] 2.3.4.6  EncryptionInfo Stream (Extensible Encryption) */
  function parse_EncInfoExt() {
    throw new Error("File is password-protected: ECMA-376 Extensible");
  }
  /* [MS-OFFCRYPTO] 2.3.4.10 EncryptionInfo Stream (Agile Encryption) */
  function parse_EncInfoAgl(blob) {
    const KeyData = [
      "saltSize",
      "blockSize",
      "keyBits",
      "hashSize",
      "cipherAlgorithm",
      "cipherChaining",
      "hashAlgorithm",
      "saltValue"
    ];
    blob.l += 4;
    const xml = blob.read_shift(blob.length - blob.l, "utf8");
    const o = {};
    xml.replace(tagregex, function xml_agile(x) {
      const y = parsexmltag(x);
      switch (strip_ns(y[0])) {
        case "<?xml":
          break;
        case "<encryption":
        case "</encryption>":
          break;
        case "<keyData":
          KeyData.forEach(function(k) {
            o[k] = y[k];
          });
          break;
        case "<dataIntegrity":
          o.encryptedHmacKey = y.encryptedHmacKey;
          o.encryptedHmacValue = y.encryptedHmacValue;
          break;
        case "<keyEncryptors>":
        case "<keyEncryptors":
          o.encs = [];
          break;
        case "</keyEncryptors>":
          break;

        case "<keyEncryptor":
          o.uri = y.uri;
          break;
        case "</keyEncryptor>":
          break;
        case "<encryptedKey":
          o.encs.push(y);
          break;
        default:
          throw y[0];
      }
    });
    return o;
  }

  /* [MS-OFFCRYPTO] 2.3.5.1 RC4 CryptoAPI Encryption Header */
  function parse_RC4CryptoHeader(blob, length) {
    const o = {};
    const vers = (o.EncryptionVersionInfo = parse_CRYPTOVersion(blob, 4));
    length -= 4;
    if (vers.Minor != 2)
      throw new Error(`unrecognized minor version code: ${vers.Minor}`);
    if (vers.Major > 4 || vers.Major < 2)
      throw new Error(`unrecognized major version code: ${vers.Major}`);
    o.Flags = blob.read_shift(4);
    length -= 4;
    const sz = blob.read_shift(4);
    length -= 4;
    o.EncryptionHeader = parse_EncryptionHeader(blob, sz);
    length -= sz;
    o.EncryptionVerifier = parse_EncryptionVerifier(blob, length);
    return o;
  }
  /* [MS-OFFCRYPTO] 2.3.6.1 RC4 Encryption Header */
  function parse_RC4Header(blob) {
    const o = {};
    const vers = (o.EncryptionVersionInfo = parse_CRYPTOVersion(blob, 4));
    if (vers.Major != 1 || vers.Minor != 1)
      throw `unrecognized version code ${vers.Major} : ${vers.Minor}`;
    o.Salt = blob.read_shift(16);
    o.EncryptedVerifier = blob.read_shift(16);
    o.EncryptedVerifierHash = blob.read_shift(16);
    return o;
  }

  /* [MS-OFFCRYPTO] 2.3.7.1 Binary Document Password Verifier Derivation */
  function crypto_CreatePasswordVerifier_Method1(Password) {
    let Verifier = 0x0000;
    let PasswordArray;
    const PasswordDecoded = _JS2ANSI(Password);
    const len = PasswordDecoded.length + 1;
    let i;
    let PasswordByte;
    let Intermediate1;
    let Intermediate2;
    let Intermediate3;
    PasswordArray = new_raw_buf(len);
    PasswordArray[0] = PasswordDecoded.length;
    for (i = 1; i != len; ++i) PasswordArray[i] = PasswordDecoded[i - 1];
    for (i = len - 1; i >= 0; --i) {
      PasswordByte = PasswordArray[i];
      Intermediate1 = (Verifier & 0x4000) === 0x0000 ? 0 : 1;
      Intermediate2 = (Verifier << 1) & 0x7fff;
      Intermediate3 = Intermediate1 | Intermediate2;
      Verifier = Intermediate3 ^ PasswordByte;
    }
    return Verifier ^ 0xce4b;
  }

  /* [MS-OFFCRYPTO] 2.3.7.2 Binary Document XOR Array Initialization */
  const crypto_CreateXorArray_Method1 = (function() {
    const PadArray = [
      0xbb,
      0xff,
      0xff,
      0xba,
      0xff,
      0xff,
      0xb9,
      0x80,
      0x00,
      0xbe,
      0x0f,
      0x00,
      0xbf,
      0x0f,
      0x00
    ];
    const InitialCode = [
      0xe1f0,
      0x1d0f,
      0xcc9c,
      0x84c0,
      0x110c,
      0x0e10,
      0xf1ce,
      0x313e,
      0x1872,
      0xe139,
      0xd40f,
      0x84f9,
      0x280c,
      0xa96a,
      0x4ec3
    ];
    const XorMatrix = [
      0xaefc,
      0x4dd9,
      0x9bb2,
      0x2745,
      0x4e8a,
      0x9d14,
      0x2a09,
      0x7b61,
      0xf6c2,
      0xfda5,
      0xeb6b,
      0xc6f7,
      0x9dcf,
      0x2bbf,
      0x4563,
      0x8ac6,
      0x05ad,
      0x0b5a,
      0x16b4,
      0x2d68,
      0x5ad0,
      0x0375,
      0x06ea,
      0x0dd4,
      0x1ba8,
      0x3750,
      0x6ea0,
      0xdd40,
      0xd849,
      0xa0b3,
      0x5147,
      0xa28e,
      0x553d,
      0xaa7a,
      0x44d5,
      0x6f45,
      0xde8a,
      0xad35,
      0x4a4b,
      0x9496,
      0x390d,
      0x721a,
      0xeb23,
      0xc667,
      0x9cef,
      0x29ff,
      0x53fe,
      0xa7fc,
      0x5fd9,
      0x47d3,
      0x8fa6,
      0x0f6d,
      0x1eda,
      0x3db4,
      0x7b68,
      0xf6d0,
      0xb861,
      0x60e3,
      0xc1c6,
      0x93ad,
      0x377b,
      0x6ef6,
      0xddec,
      0x45a0,
      0x8b40,
      0x06a1,
      0x0d42,
      0x1a84,
      0x3508,
      0x6a10,
      0xaa51,
      0x4483,
      0x8906,
      0x022d,
      0x045a,
      0x08b4,
      0x1168,
      0x76b4,
      0xed68,
      0xcaf1,
      0x85c3,
      0x1ba7,
      0x374e,
      0x6e9c,
      0x3730,
      0x6e60,
      0xdcc0,
      0xa9a1,
      0x4363,
      0x86c6,
      0x1dad,
      0x3331,
      0x6662,
      0xccc4,
      0x89a9,
      0x0373,
      0x06e6,
      0x0dcc,
      0x1021,
      0x2042,
      0x4084,
      0x8108,
      0x1231,
      0x2462,
      0x48c4
    ];
    const Ror = function(Byte) {
      return ((Byte / 2) | (Byte * 128)) & 0xff;
    };
    const XorRor = function(byte1, byte2) {
      return Ror(byte1 ^ byte2);
    };
    const CreateXorKey_Method1 = function(Password) {
      let XorKey = InitialCode[Password.length - 1];
      let CurrentElement = 0x68;
      for (let i = Password.length - 1; i >= 0; --i) {
        let Char = Password[i];
        for (let j = 0; j != 7; ++j) {
          if (Char & 0x40) XorKey ^= XorMatrix[CurrentElement];
          Char *= 2;
          --CurrentElement;
        }
      }
      return XorKey;
    };
    return function(password) {
      const Password = _JS2ANSI(password);
      const XorKey = CreateXorKey_Method1(Password);
      let Index = Password.length;
      const ObfuscationArray = new_raw_buf(16);
      for (let i = 0; i != 16; ++i) ObfuscationArray[i] = 0x00;
      let Temp;
      let PasswordLastChar;
      let PadIndex;
      if ((Index & 1) === 1) {
        Temp = XorKey >> 8;
        ObfuscationArray[Index] = XorRor(PadArray[0], Temp);
        --Index;
        Temp = XorKey & 0xff;
        PasswordLastChar = Password[Password.length - 1];
        ObfuscationArray[Index] = XorRor(PasswordLastChar, Temp);
      }
      while (Index > 0) {
        --Index;
        Temp = XorKey >> 8;
        ObfuscationArray[Index] = XorRor(Password[Index], Temp);
        --Index;
        Temp = XorKey & 0xff;
        ObfuscationArray[Index] = XorRor(Password[Index], Temp);
      }
      Index = 15;
      PadIndex = 15 - Password.length;
      while (PadIndex > 0) {
        Temp = XorKey >> 8;
        ObfuscationArray[Index] = XorRor(PadArray[PadIndex], Temp);
        --Index;
        --PadIndex;
        Temp = XorKey & 0xff;
        ObfuscationArray[Index] = XorRor(Password[Index], Temp);
        --Index;
        --PadIndex;
      }
      return ObfuscationArray;
    };
  })();

  /* [MS-OFFCRYPTO] 2.3.7.3 Binary Document XOR Data Transformation Method 1 */
  const crypto_DecryptData_Method1 = function(
    password,
    Data,
    XorArrayIndex,
    XorArray,
    O
  ) {
    /* If XorArray is set, use it; if O is not set, make changes in-place */
    if (!O) O = Data;
    if (!XorArray) XorArray = crypto_CreateXorArray_Method1(password);
    let Index;
    let Value;
    for (Index = 0; Index != Data.length; ++Index) {
      Value = Data[Index];
      Value ^= XorArray[XorArrayIndex];
      Value = ((Value >> 5) | (Value << 3)) & 0xff;
      O[Index] = Value;
      ++XorArrayIndex;
    }
    return [O, XorArrayIndex, XorArray];
  };

  const crypto_MakeXorDecryptor = function(password) {
    let XorArrayIndex = 0;
    const XorArray = crypto_CreateXorArray_Method1(password);
    return function(Data) {
      const O = crypto_DecryptData_Method1("", Data, XorArrayIndex, XorArray);
      XorArrayIndex = O[1];
      return O[0];
    };
  };

  /* 2.5.343 */
  function parse_XORObfuscation(blob, length, opts, out) {
    const o = { key: parseuint16(blob), verificationBytes: parseuint16(blob) };
    if (opts.password)
      o.verifier = crypto_CreatePasswordVerifier_Method1(opts.password);
    out.valid = o.verificationBytes === o.verifier;
    if (out.valid) out.insitu = crypto_MakeXorDecryptor(opts.password);
    return o;
  }

  /* 2.4.117 */
  function parse_FilePassHeader(blob, length, oo) {
    const o = oo || {};
    o.Info = blob.read_shift(2);
    blob.l -= 2;
    if (o.Info === 1) o.Data = parse_RC4Header(blob, length);
    else o.Data = parse_RC4CryptoHeader(blob, length);
    return o;
  }
  function parse_FilePass(blob, length, opts) {
    const o = {
      Type: opts.biff >= 8 ? blob.read_shift(2) : 0
    }; /* wEncryptionType */
    if (o.Type) parse_FilePassHeader(blob, length - 2, o);
    else
      parse_XORObfuscation(blob, opts.biff >= 8 ? length : length - 2, opts, o);
    return o;
  }

  const RTF = (function() {
    function rtf_to_sheet(d, opts) {
      switch (opts.type) {
        case "base64":
          return rtf_to_sheet_str(Base64.decode(d), opts);
        case "binary":
          return rtf_to_sheet_str(d, opts);
        case "buffer":
          return rtf_to_sheet_str(d.toString("binary"), opts);
        case "array":
          return rtf_to_sheet_str(cc2str(d), opts);
      }
      throw new Error(`Unrecognized type ${opts.type}`);
    }

    function rtf_to_sheet_str(str, opts) {
      const o = opts || {};
      const ws = o.dense ? [] : {};
      const range = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };

      // TODO: parse
      if (!str.match(/\\trowd/)) throw new Error("RTF missing table");

      ws["!ref"] = encode_range(range);
      return ws;
    }

    function rtf_to_workbook(d, opts) {
      return sheet_to_workbook(rtf_to_sheet(d, opts), opts);
    }

    /* TODO: this is a stub */
    function sheet_to_rtf(ws) {
      const o = ["{\\rtf1\\ansi"];
      const r = safe_decode_range(ws["!ref"]);
      let cell;
      const dense = Array.isArray(ws);
      for (let R = r.s.r; R <= r.e.r; ++R) {
        o.push("\\trowd\\trautofit1");
        for (var C = r.s.c; C <= r.e.c; ++C) o.push(`\\cellx${C + 1}`);
        o.push("\\pard\\intbl");
        for (C = r.s.c; C <= r.e.c; ++C) {
          const coord = encode_cell({ r: R, c: C });
          cell = dense ? (ws[R] || [])[C] : ws[coord];
          if (!cell || (cell.v == null && (!cell.f || cell.F))) continue;
          o.push(` ${cell.w || (format_cell(cell), cell.w)}`);
          o.push("\\cell");
        }
        o.push("\\pard\\intbl\\row");
      }
      return `${o.join("")}}`;
    }

    return {
      to_workbook: rtf_to_workbook,
      to_sheet: rtf_to_sheet,
      from_sheet: sheet_to_rtf
    };
  })();
  function hex2RGB(h) {
    const o = h.slice(h[0] === "#" ? 1 : 0).slice(0, 6);
    return [
      parseInt(o.slice(0, 2), 16),
      parseInt(o.slice(2, 4), 16),
      parseInt(o.slice(4, 6), 16)
    ];
  }
  function rgb2Hex(rgb) {
    for (var i = 0, o = 1; i != 3; ++i)
      o = o * 256 + (rgb[i] > 255 ? 255 : rgb[i] < 0 ? 0 : rgb[i]);
    return o
      .toString(16)
      .toUpperCase()
      .slice(1);
  }

  function rgb2HSL(rgb) {
    const R = rgb[0] / 255;
    const G = rgb[1] / 255;
    const B = rgb[2] / 255;
    const M = Math.max(R, G, B);
    const m = Math.min(R, G, B);
    const C = M - m;
    if (C === 0) return [0, 0, R];

    let H6 = 0;
    let S = 0;
    const L2 = M + m;
    S = C / (L2 > 1 ? 2 - L2 : L2);
    switch (M) {
      case R:
        H6 = ((G - B) / C + 6) % 6;
        break;
      case G:
        H6 = (B - R) / C + 2;
        break;
      case B:
        H6 = (R - G) / C + 4;
        break;
    }
    return [H6 / 6, S, L2 / 2];
  }

  function hsl2RGB(hsl) {
    const H = hsl[0];
    const S = hsl[1];
    const L = hsl[2];
    const C = S * 2 * (L < 0.5 ? L : 1 - L);
    const m = L - C / 2;
    const rgb = [m, m, m];
    const h6 = 6 * H;

    let X;
    if (S !== 0)
      switch (h6 | 0) {
        case 0:
        case 6:
          X = C * h6;
          rgb[0] += C;
          rgb[1] += X;
          break;
        case 1:
          X = C * (2 - h6);
          rgb[0] += X;
          rgb[1] += C;
          break;
        case 2:
          X = C * (h6 - 2);
          rgb[1] += C;
          rgb[2] += X;
          break;
        case 3:
          X = C * (4 - h6);
          rgb[1] += X;
          rgb[2] += C;
          break;
        case 4:
          X = C * (h6 - 4);
          rgb[2] += C;
          rgb[0] += X;
          break;
        case 5:
          X = C * (6 - h6);
          rgb[2] += X;
          rgb[0] += C;
          break;
      }
    for (let i = 0; i != 3; ++i) rgb[i] = Math.round(rgb[i] * 255);
    return rgb;
  }

  /* 18.8.3 bgColor tint algorithm */
  function rgb_tint(hex, tint) {
    if (tint === 0) return hex;
    const hsl = rgb2HSL(hex2RGB(hex));
    if (tint < 0) hsl[2] *= 1 + tint;
    else hsl[2] = 1 - (1 - hsl[2]) * (1 - tint);
    return rgb2Hex(hsl2RGB(hsl));
  }

  /* 18.3.1.13 width calculations */
  /* [MS-OI29500] 2.1.595 Column Width & Formatting */
  const DEF_MDW = 6;
  const MAX_MDW = 15;
  const MIN_MDW = 1;
  let MDW = DEF_MDW;
  function width2px(width) {
    return Math.floor((width + Math.round(128 / MDW) / 256) * MDW);
  }
  function px2char(px) {
    return Math.floor(((px - 5) / MDW) * 100 + 0.5) / 100;
  }
  function char2width(chr) {
    return Math.round(((chr * MDW + 5) / MDW) * 256) / 256;
  }
  // function px2char_(px) { return (((px - 5)/MDW * 100 + 0.5))/100; }
  // function char2width_(chr) { return (((chr * MDW + 5)/MDW*256))/256; }
  function cycle_width(collw) {
    return char2width(px2char(width2px(collw)));
  }
  /* XLSX/XLSB/XLS specify width in units of MDW */
  function find_mdw_colw(collw) {
    let delta = Math.abs(collw - cycle_width(collw));
    let _MDW = MDW;
    if (delta > 0.005)
      for (MDW = MIN_MDW; MDW < MAX_MDW; ++MDW)
        if (Math.abs(collw - cycle_width(collw)) <= delta) {
          delta = Math.abs(collw - cycle_width(collw));
          _MDW = MDW;
        }
    MDW = _MDW;
  }
  /* XLML specifies width in terms of pixels */
  /* function find_mdw_wpx(wpx) {
	var delta = Infinity, guess = 0, _MDW = MIN_MDW;
	for(MDW=MIN_MDW; MDW<MAX_MDW; ++MDW) {
		guess = char2width_(px2char_(wpx))*256;
		guess = (guess) % 1;
		if(guess > 0.5) guess--;
		if(Math.abs(guess) < delta) { delta = Math.abs(guess); _MDW = MDW; }
	}
	MDW = _MDW;
} */

  function process_col(coll) {
    if (coll.width) {
      coll.wpx = width2px(coll.width);
      coll.wch = px2char(coll.wpx);
      coll.MDW = MDW;
    } else if (coll.wpx) {
      coll.wch = px2char(coll.wpx);
      coll.width = char2width(coll.wch);
      coll.MDW = MDW;
    } else if (typeof coll.wch === "number") {
      coll.width = char2width(coll.wch);
      coll.wpx = width2px(coll.width);
      coll.MDW = MDW;
    }
    if (coll.customWidth) delete coll.customWidth;
  }

  const DEF_PPI = 96;
  const PPI = DEF_PPI;
  function px2pt(px) {
    return (px * 96) / PPI;
  }
  function pt2px(pt) {
    return (pt * PPI) / 96;
  }

  /* [MS-EXSPXML3] 2.4.54 ST_enmPattern */
  const XLMLPatternTypeMap = {
    None: "none",
    Solid: "solid",
    Gray50: "mediumGray",
    Gray75: "darkGray",
    Gray25: "lightGray",
    HorzStripe: "darkHorizontal",
    VertStripe: "darkVertical",
    ReverseDiagStripe: "darkDown",
    DiagStripe: "darkUp",
    DiagCross: "darkGrid",
    ThickDiagCross: "darkTrellis",
    ThinHorzStripe: "lightHorizontal",
    ThinVertStripe: "lightVertical",
    ThinReverseDiagStripe: "lightDown",
    ThinHorzCross: "lightGrid"
  };

  /* 18.8.5 borders CT_Borders */
  function parse_borders(t, styles, themes, opts) {
    styles.Borders = [];
    let border = {};
    let pass = false;
    (t[0].match(tagregex) || []).forEach(function(x) {
      const y = parsexmltag(x);
      switch (strip_ns(y[0])) {
        case "<borders":
        case "<borders>":
        case "</borders>":
          break;

        /* 18.8.4 border CT_Border */
        case "<border":
        case "<border>":
        case "<border/>":
          border = {};
          if (y.diagonalUp) border.diagonalUp = parsexmlbool(y.diagonalUp);
          if (y.diagonalDown)
            border.diagonalDown = parsexmlbool(y.diagonalDown);
          styles.Borders.push(border);
          break;
        case "</border>":
          break;

        /* note: not in spec, appears to be CT_BorderPr */
        case "<left/>":
          break;
        case "<left":
        case "<left>":
          break;
        case "</left>":
          break;

        /* note: not in spec, appears to be CT_BorderPr */
        case "<right/>":
          break;
        case "<right":
        case "<right>":
          break;
        case "</right>":
          break;

        /* 18.8.43 top CT_BorderPr */
        case "<top/>":
          break;
        case "<top":
        case "<top>":
          break;
        case "</top>":
          break;

        /* 18.8.6 bottom CT_BorderPr */
        case "<bottom/>":
          break;
        case "<bottom":
        case "<bottom>":
          break;
        case "</bottom>":
          break;

        /* 18.8.13 diagonal CT_BorderPr */
        case "<diagonal":
        case "<diagonal>":
        case "<diagonal/>":
          break;
        case "</diagonal>":
          break;

        /* 18.8.25 horizontal CT_BorderPr */
        case "<horizontal":
        case "<horizontal>":
        case "<horizontal/>":
          break;
        case "</horizontal>":
          break;

        /* 18.8.44 vertical CT_BorderPr */
        case "<vertical":
        case "<vertical>":
        case "<vertical/>":
          break;
        case "</vertical>":
          break;

        /* 18.8.37 start CT_BorderPr */
        case "<start":
        case "<start>":
        case "<start/>":
          break;
        case "</start>":
          break;

        /* 18.8.16 end CT_BorderPr */
        case "<end":
        case "<end>":
        case "<end/>":
          break;
        case "</end>":
          break;

        /* 18.8.? color CT_Color */
        case "<color":
        case "<color>":
          break;
        case "<color/>":
        case "</color>":
          break;

        /* 18.2.10 extLst CT_ExtensionList ? */
        case "<extLst":
        case "<extLst>":
        case "</extLst>":
          break;
        case "<ext":
          pass = true;
          break;
        case "</ext>":
          pass = false;
          break;
        default:
          if (opts && opts.WTF) {
            if (!pass) throw new Error(`unrecognized ${y[0]} in borders`);
          }
      }
    });
  }

  /* 18.8.21 fills CT_Fills */
  function parse_fills(t, styles, themes, opts) {
    styles.Fills = [];
    let fill = {};
    let pass = false;
    (t[0].match(tagregex) || []).forEach(function(x) {
      const y = parsexmltag(x);
      switch (strip_ns(y[0])) {
        case "<fills":
        case "<fills>":
        case "</fills>":
          break;

        /* 18.8.20 fill CT_Fill */
        case "<fill>":
        case "<fill":
        case "<fill/>":
          fill = {};
          styles.Fills.push(fill);
          break;
        case "</fill>":
          break;

        /* 18.8.24 gradientFill CT_GradientFill */
        case "<gradientFill>":
          break;
        case "<gradientFill":
        case "</gradientFill>":
          styles.Fills.push(fill);
          fill = {};
          break;

        /* 18.8.32 patternFill CT_PatternFill */
        case "<patternFill":
        case "<patternFill>":
          if (y.patternType) fill.patternType = y.patternType;
          break;
        case "<patternFill/>":
        case "</patternFill>":
          break;

        /* 18.8.3 bgColor CT_Color */
        case "<bgColor":
          if (!fill.bgColor) fill.bgColor = {};
          if (y.indexed) fill.bgColor.indexed = parseInt(y.indexed, 10);
          if (y.theme) fill.bgColor.theme = parseInt(y.theme, 10);
          if (y.tint) fill.bgColor.tint = parseFloat(y.tint);
          /* Excel uses ARGB strings */
          if (y.rgb) fill.bgColor.rgb = y.rgb.slice(-6);
          break;
        case "<bgColor/>":
        case "</bgColor>":
          break;

        /* 18.8.19 fgColor CT_Color */
        case "<fgColor":
          if (!fill.fgColor) fill.fgColor = {};
          if (y.theme) fill.fgColor.theme = parseInt(y.theme, 10);
          if (y.tint) fill.fgColor.tint = parseFloat(y.tint);
          /* Excel uses ARGB strings */
          if (y.rgb != null) fill.fgColor.rgb = y.rgb.slice(-6);
          break;
        case "<fgColor/>":
        case "</fgColor>":
          break;

        /* 18.8.38 stop CT_GradientStop */
        case "<stop":
        case "<stop/>":
          break;
        case "</stop>":
          break;

        /* 18.8.? color CT_Color */
        case "<color":
        case "<color/>":
          break;
        case "</color>":
          break;

        /* 18.2.10 extLst CT_ExtensionList ? */
        case "<extLst":
        case "<extLst>":
        case "</extLst>":
          break;
        case "<ext":
          pass = true;
          break;
        case "</ext>":
          pass = false;
          break;
        default:
          if (opts && opts.WTF) {
            if (!pass) throw new Error(`unrecognized ${y[0]} in fills`);
          }
      }
    });
  }

  /* 18.8.23 fonts CT_Fonts */
  function parse_fonts(t, styles, themes, opts) {
    styles.Fonts = [];
    let font = {};
    let pass = false;
    (t[0].match(tagregex) || []).forEach(function(x) {
      const y = parsexmltag(x);
      switch (strip_ns(y[0])) {
        case "<fonts":
        case "<fonts>":
        case "</fonts>":
          break;

        /* 18.8.22 font CT_Font */
        case "<font":
        case "<font>":
          break;
        case "</font>":
        case "<font/>":
          styles.Fonts.push(font);
          font = {};
          break;

        /* 18.8.29 name CT_FontName */
        case "<name":
          if (y.val) font.name = utf8read(y.val);
          break;
        case "<name/>":
        case "</name>":
          break;

        /* 18.8.2  b CT_BooleanProperty */
        case "<b":
          font.bold = y.val ? parsexmlbool(y.val) : 1;
          break;
        case "<b/>":
          font.bold = 1;
          break;

        /* 18.8.26 i CT_BooleanProperty */
        case "<i":
          font.italic = y.val ? parsexmlbool(y.val) : 1;
          break;
        case "<i/>":
          font.italic = 1;
          break;

        /* 18.4.13 u CT_UnderlineProperty */
        case "<u":
          switch (y.val) {
            case "none":
              font.underline = 0x00;
              break;
            case "single":
              font.underline = 0x01;
              break;
            case "double":
              font.underline = 0x02;
              break;
            case "singleAccounting":
              font.underline = 0x21;
              break;
            case "doubleAccounting":
              font.underline = 0x22;
              break;
          }
          break;
        case "<u/>":
          font.underline = 1;
          break;

        /* 18.4.10 strike CT_BooleanProperty */
        case "<strike":
          font.strike = y.val ? parsexmlbool(y.val) : 1;
          break;
        case "<strike/>":
          font.strike = 1;
          break;

        /* 18.4.2  outline CT_BooleanProperty */
        case "<outline":
          font.outline = y.val ? parsexmlbool(y.val) : 1;
          break;
        case "<outline/>":
          font.outline = 1;
          break;

        /* 18.8.36 shadow CT_BooleanProperty */
        case "<shadow":
          font.shadow = y.val ? parsexmlbool(y.val) : 1;
          break;
        case "<shadow/>":
          font.shadow = 1;
          break;

        /* 18.8.12 condense CT_BooleanProperty */
        case "<condense":
          font.condense = y.val ? parsexmlbool(y.val) : 1;
          break;
        case "<condense/>":
          font.condense = 1;
          break;

        /* 18.8.17 extend CT_BooleanProperty */
        case "<extend":
          font.extend = y.val ? parsexmlbool(y.val) : 1;
          break;
        case "<extend/>":
          font.extend = 1;
          break;

        /* 18.4.11 sz CT_FontSize */
        case "<sz":
          if (y.val) font.sz = +y.val;
          break;
        case "<sz/>":
        case "</sz>":
          break;

        /* 18.4.14 vertAlign CT_VerticalAlignFontProperty */
        case "<vertAlign":
          if (y.val) font.vertAlign = y.val;
          break;
        case "<vertAlign/>":
        case "</vertAlign>":
          break;

        /* 18.8.18 family CT_FontFamily */
        case "<family":
          if (y.val) font.family = parseInt(y.val, 10);
          break;
        case "<family/>":
        case "</family>":
          break;

        /* 18.8.35 scheme CT_FontScheme */
        case "<scheme":
          if (y.val) font.scheme = y.val;
          break;
        case "<scheme/>":
        case "</scheme>":
          break;

        /* 18.4.1 charset CT_IntProperty */
        case "<charset":
          if (y.val == "1") break;
          y.codepage = CS2CP[parseInt(y.val, 10)];
          break;

        /* 18.?.? color CT_Color */
        case "<color":
          if (!font.color) font.color = {};
          if (y.auto) font.color.auto = parsexmlbool(y.auto);

          if (y.rgb) font.color.rgb = y.rgb.slice(-6);
          else if (y.indexed) {
            font.color.index = parseInt(y.indexed, 10);
            let icv = XLSIcv[font.color.index];
            if (font.color.index == 81) icv = XLSIcv[1];
            if (!icv) throw new Error(x);
            font.color.rgb =
              icv[0].toString(16) + icv[1].toString(16) + icv[2].toString(16);
          } else if (y.theme) {
            font.color.theme = parseInt(y.theme, 10);
            if (y.tint) font.color.tint = parseFloat(y.tint);
            if (
              y.theme &&
              themes.themeElements &&
              themes.themeElements.clrScheme
            ) {
              font.color.rgb = rgb_tint(
                themes.themeElements.clrScheme[font.color.theme].rgb,
                font.color.tint || 0
              );
            }
          }

          break;
        case "<color/>":
        case "</color>":
          break;

        /* note: sometimes mc:AlternateContent appears bare */
        case "<AlternateContent":
          pass = true;
          break;
        case "</AlternateContent>":
          pass = false;
          break;

        /* 18.2.10 extLst CT_ExtensionList ? */
        case "<extLst":
        case "<extLst>":
        case "</extLst>":
          break;
        case "<ext":
          pass = true;
          break;
        case "</ext>":
          pass = false;
          break;
        default:
          if (opts && opts.WTF) {
            if (!pass) throw new Error(`unrecognized ${y[0]} in fonts`);
          }
      }
    });
  }

  /* 18.8.31 numFmts CT_NumFmts */
  function parse_numFmts(t, styles, opts) {
    styles.NumberFmt = [];
    const k /* Array<number> */ = keys(SSF._table);
    for (var i = 0; i < k.length; ++i)
      styles.NumberFmt[k[i]] = SSF._table[k[i]];
    const m = t[0].match(tagregex);
    if (!m) return;
    for (i = 0; i < m.length; ++i) {
      const y = parsexmltag(m[i]);
      switch (strip_ns(y[0])) {
        case "<numFmts":
        case "</numFmts>":
        case "<numFmts/>":
        case "<numFmts>":
          break;
        case "<numFmt":
          {
            const f = unescapexml(utf8read(y.formatCode));
            let j = parseInt(y.numFmtId, 10);
            styles.NumberFmt[j] = f;
            if (j > 0) {
              if (j > 0x188) {
                for (j = 0x188; j > 0x3c; --j)
                  if (styles.NumberFmt[j] == null) break;
                styles.NumberFmt[j] = f;
              }
              SSF.load(f, j);
            }
          }
          break;
        case "</numFmt>":
          break;
        default:
          if (opts.WTF) throw new Error(`unrecognized ${y[0]} in numFmts`);
      }
    }
  }

  /* 18.8.10 cellXfs CT_CellXfs */
  const cellXF_uint = ["numFmtId", "fillId", "fontId", "borderId", "xfId"];
  const cellXF_bool = [
    "applyAlignment",
    "applyBorder",
    "applyFill",
    "applyFont",
    "applyNumberFormat",
    "applyProtection",
    "pivotButton",
    "quotePrefix"
  ];
  function parse_cellXfs(t, styles, opts) {
    styles.CellXf = [];
    let xf;
    let pass = false;
    (t[0].match(tagregex) || []).forEach(function(x) {
      const y = parsexmltag(x);
      let i = 0;
      switch (strip_ns(y[0])) {
        case "<cellXfs":
        case "<cellXfs>":
        case "<cellXfs/>":
        case "</cellXfs>":
          break;

        /* 18.8.45 xf CT_Xf */
        case "<xf":
        case "<xf/>":
          xf = y;
          delete xf[0];
          for (i = 0; i < cellXF_uint.length; ++i)
            if (xf[cellXF_uint[i]])
              xf[cellXF_uint[i]] = parseInt(xf[cellXF_uint[i]], 10);
          for (i = 0; i < cellXF_bool.length; ++i)
            if (xf[cellXF_bool[i]])
              xf[cellXF_bool[i]] = parsexmlbool(xf[cellXF_bool[i]]);
          if (xf.numFmtId > 0x188) {
            for (i = 0x188; i > 0x3c; --i)
              if (styles.NumberFmt[xf.numFmtId] == styles.NumberFmt[i]) {
                xf.numFmtId = i;
                break;
              }
          }
          styles.CellXf.push(xf);
          break;
        case "</xf>":
          break;

        /* 18.8.1 alignment CT_CellAlignment */
        case "<alignment":
        case "<alignment/>":
          var alignment = {};
          if (y.vertical) alignment.vertical = y.vertical;
          if (y.horizontal) alignment.horizontal = y.horizontal;
          if (y.textRotation != null) alignment.textRotation = y.textRotation;
          if (y.indent) alignment.indent = y.indent;
          if (y.wrapText) alignment.wrapText = parsexmlbool(y.wrapText);
          xf.alignment = alignment;
          break;
        case "</alignment>":
          break;

        /* 18.8.33 protection CT_CellProtection */
        case "<protection":
          break;
        case "</protection>":
        case "<protection/>":
          break;

        /* note: sometimes mc:AlternateContent appears bare */
        case "<AlternateContent":
          pass = true;
          break;
        case "</AlternateContent>":
          pass = false;
          break;

        /* 18.2.10 extLst CT_ExtensionList ? */
        case "<extLst":
        case "<extLst>":
        case "</extLst>":
          break;
        case "<ext":
          pass = true;
          break;
        case "</ext>":
          pass = false;
          break;
        default:
          if (opts && opts.WTF) {
            if (!pass) throw new Error(`unrecognized ${y[0]} in cellXfs`);
          }
      }
    });
  }

  /* 18.8 Styles CT_Stylesheet */
  const parse_sty_xml = (function make_pstyx() {
    const numFmtRegex = /<(?:\w+:)?numFmts([^>]*)>[\S\s]*?<\/(?:\w+:)?numFmts>/;
    const cellXfRegex = /<(?:\w+:)?cellXfs([^>]*)>[\S\s]*?<\/(?:\w+:)?cellXfs>/;
    const fillsRegex = /<(?:\w+:)?fills([^>]*)>[\S\s]*?<\/(?:\w+:)?fills>/;
    const fontsRegex = /<(?:\w+:)?fonts([^>]*)>[\S\s]*?<\/(?:\w+:)?fonts>/;
    const bordersRegex = /<(?:\w+:)?borders([^>]*)>[\S\s]*?<\/(?:\w+:)?borders>/;

    return function parse_sty_xml(data, themes, opts) {
      const styles = {};
      if (!data) return styles;
      data = data
        .replace(/<!--([\s\S]*?)-->/gm, "")
        .replace(/<!DOCTYPE[^\[]*\[[^\]]*\]>/gm, "");
      /* 18.8.39 styleSheet CT_Stylesheet */
      let t;

      /* 18.8.31 numFmts CT_NumFmts ? */
      if ((t = data.match(numFmtRegex))) parse_numFmts(t, styles, opts);

      /* 18.8.23 fonts CT_Fonts ? */
      if ((t = data.match(fontsRegex))) parse_fonts(t, styles, themes, opts);

      /* 18.8.21 fills CT_Fills ? */
      if ((t = data.match(fillsRegex))) parse_fills(t, styles, themes, opts);

      /* 18.8.5  borders CT_Borders ? */
      if ((t = data.match(bordersRegex)))
        parse_borders(t, styles, themes, opts);

      /* 18.8.9  cellStyleXfs CT_CellStyleXfs ? */
      /* 18.8.8  cellStyles CT_CellStyles ? */

      /* 18.8.10 cellXfs CT_CellXfs ? */
      if ((t = data.match(cellXfRegex))) parse_cellXfs(t, styles, opts);

      /* 18.8.15 dxfs CT_Dxfs ? */
      /* 18.8.42 tableStyles CT_TableStyles ? */
      /* 18.8.11 colors CT_Colors ? */
      /* 18.2.10 extLst CT_ExtensionList ? */

      return styles;
    };
  })();

  RELS.STY =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

  /* [MS-XLSB] 2.4.657 BrtFmt */
  function parse_BrtFmt(data, length) {
    const numFmtId = data.read_shift(2);
    const stFmtCode = parse_XLWideString(data, length - 2);
    return [numFmtId, stFmtCode];
  }

  /* [MS-XLSB] 2.4.659 BrtFont TODO */
  function parse_BrtFont(data, length, opts) {
    const out = {};

    out.sz = data.read_shift(2) / 20;

    const grbit = parse_FontFlags(data, 2, opts);
    if (grbit.fItalic) out.italic = 1;
    if (grbit.fCondense) out.condense = 1;
    if (grbit.fExtend) out.extend = 1;
    if (grbit.fShadow) out.shadow = 1;
    if (grbit.fOutline) out.outline = 1;
    if (grbit.fStrikeout) out.strike = 1;

    const bls = data.read_shift(2);
    if (bls === 0x02bc) out.bold = 1;

    switch (data.read_shift(2)) {
      /* case 0: out.vertAlign = "baseline"; break; */
      case 1:
        out.vertAlign = "superscript";
        break;
      case 2:
        out.vertAlign = "subscript";
        break;
    }

    const underline = data.read_shift(1);
    if (underline != 0) out.underline = underline;

    const family = data.read_shift(1);
    if (family > 0) out.family = family;

    const bCharSet = data.read_shift(1);
    if (bCharSet > 0) out.charset = bCharSet;

    data.l++;
    out.color = parse_BrtColor(data, 8);

    switch (data.read_shift(1)) {
      /* case 0: out.scheme = "none": break; */
      case 1:
        out.scheme = "major";
        break;
      case 2:
        out.scheme = "minor";
        break;
    }

    out.name = parse_XLWideString(data, length - 21);

    return out;
  }

  /* [MS-XLSB] 2.4.650 BrtFill */
  const XLSBFillPTNames = [
    "none",
    "solid",
    "mediumGray",
    "darkGray",
    "lightGray",
    "darkHorizontal",
    "darkVertical",
    "darkDown",
    "darkUp",
    "darkGrid",
    "darkTrellis",
    "lightHorizontal",
    "lightVertical",
    "lightDown",
    "lightUp",
    "lightGrid",
    "lightTrellis",
    "gray125",
    "gray0625"
  ];
  /* TODO: gradient fill representation */
  const parse_BrtFill = parsenoop;

  /* [MS-XLSB] 2.4.824 BrtXF */
  function parse_BrtXF(data, length) {
    const tgt = data.l + length;
    const ixfeParent = data.read_shift(2);
    const ifmt = data.read_shift(2);
    data.l = tgt;
    return { ixfe: ixfeParent, numFmtId: ifmt };
  }

  /* [MS-XLSB] 2.5.4 Blxf TODO */
  function write_Blxf(data, o) {
    if (!o) o = new_buf(10);
    o.write_shift(1, 0); /* dg */
    o.write_shift(1, 0);
    o.write_shift(4, 0); /* color */
    o.write_shift(4, 0); /* color */
    return o;
  }
  /* [MS-XLSB] 2.4.302 BrtBorder TODO */
  const parse_BrtBorder = parsenoop;

  /* [MS-XLSB] 2.1.7.50 Styles */
  function parse_sty_bin(data, themes, opts) {
    const styles = {};
    styles.NumberFmt = [];
    for (const y in SSF._table) styles.NumberFmt[y] = SSF._table[y];

    styles.CellXf = [];
    styles.Fonts = [];
    const state = [];
    let pass = false;
    recordhopper(data, function hopper_sty(val, R_n, RT) {
      switch (RT) {
        case 0x002c /* 'BrtFmt' */:
          styles.NumberFmt[val[0]] = val[1];
          SSF.load(val[1], val[0]);
          break;
        case 0x002b /* 'BrtFont' */:
          styles.Fonts.push(val);
          if (
            val.color.theme != null &&
            themes &&
            themes.themeElements &&
            themes.themeElements.clrScheme
          ) {
            val.color.rgb = rgb_tint(
              themes.themeElements.clrScheme[val.color.theme].rgb,
              val.color.tint || 0
            );
          }
          break;
        case 0x0401:
          /* 'BrtKnownFonts' */ break;
        case 0x002d /* 'BrtFill' */:
          break;
        case 0x002e /* 'BrtBorder' */:
          break;
        case 0x002f /* 'BrtXF' */:
          if (state[state.length - 1] == "BrtBeginCellXFs") {
            styles.CellXf.push(val);
          }
          break;
        case 0x0030: /* 'BrtStyle' */
        case 0x01fb: /* 'BrtDXF' */
        case 0x023c: /* 'BrtMRUColor' */
        case 0x01db /* 'BrtIndexedColor': */:
          break;

        case 0x0493: /* 'BrtDXF14' */
        case 0x0836: /* 'BrtDXF15' */
        case 0x046a: /* 'BrtSlicerStyleElement' */
        case 0x0200: /* 'BrtTableStyleElement' */
        case 0x082f: /* 'BrtTimelineStyleElement' */
        case 0x0c00 /* 'BrtUid' */:
          break;

        case 0x0023 /* 'BrtFRTBegin' */:
          pass = true;
          break;
        case 0x0024 /* 'BrtFRTEnd' */:
          pass = false;
          break;
        case 0x0025 /* 'BrtACBegin' */:
          state.push(R_n);
          pass = true;
          break;
        case 0x0026 /* 'BrtACEnd' */:
          state.pop();
          pass = false;
          break;

        default:
          if ((R_n || "").indexOf("Begin") > 0) state.push(R_n);
          else if ((R_n || "").indexOf("End") > 0) state.pop();
          else if (
            !pass ||
            (opts.WTF && state[state.length - 1] != "BrtACBegin")
          )
            throw new Error(`Unexpected record ${RT} ${R_n}`);
      }
    });
    return styles;
  }

  /* [MS-XLSB] 2.1.7.50 Styles */
  RELS.THEME =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

  /* Even though theme layout is dk1 lt1 dk2 lt2, true order is lt1 dk1 lt2 dk2 */
  const XLSXThemeClrScheme = [
    "</a:lt1>",
    "</a:dk1>",
    "</a:lt2>",
    "</a:dk2>",
    "</a:accent1>",
    "</a:accent2>",
    "</a:accent3>",
    "</a:accent4>",
    "</a:accent5>",
    "</a:accent6>",
    "</a:hlink>",
    "</a:folHlink>"
  ];
  /* 20.1.6.2 clrScheme CT_ColorScheme */
  function parse_clrScheme(t, themes, opts) {
    themes.themeElements.clrScheme = [];
    let color = {};
    (t[0].match(tagregex) || []).forEach(function(x) {
      const y = parsexmltag(x);
      switch (y[0]) {
        /* 20.1.6.2 clrScheme (Color Scheme) CT_ColorScheme */
        case "<a:clrScheme":
        case "</a:clrScheme>":
          break;

        /* 20.1.2.3.32 srgbClr CT_SRgbColor */
        case "<a:srgbClr":
          color.rgb = y.val;
          break;

        /* 20.1.2.3.33 sysClr CT_SystemColor */
        case "<a:sysClr":
          color.rgb = y.lastClr;
          break;

        /* 20.1.4.1.1 accent1 (Accent 1) */
        /* 20.1.4.1.2 accent2 (Accent 2) */
        /* 20.1.4.1.3 accent3 (Accent 3) */
        /* 20.1.4.1.4 accent4 (Accent 4) */
        /* 20.1.4.1.5 accent5 (Accent 5) */
        /* 20.1.4.1.6 accent6 (Accent 6) */
        /* 20.1.4.1.9 dk1 (Dark 1) */
        /* 20.1.4.1.10 dk2 (Dark 2) */
        /* 20.1.4.1.15 folHlink (Followed Hyperlink) */
        /* 20.1.4.1.19 hlink (Hyperlink) */
        /* 20.1.4.1.22 lt1 (Light 1) */
        /* 20.1.4.1.23 lt2 (Light 2) */
        case "<a:dk1>":
        case "</a:dk1>":
        case "<a:lt1>":
        case "</a:lt1>":
        case "<a:dk2>":
        case "</a:dk2>":
        case "<a:lt2>":
        case "</a:lt2>":
        case "<a:accent1>":
        case "</a:accent1>":
        case "<a:accent2>":
        case "</a:accent2>":
        case "<a:accent3>":
        case "</a:accent3>":
        case "<a:accent4>":
        case "</a:accent4>":
        case "<a:accent5>":
        case "</a:accent5>":
        case "<a:accent6>":
        case "</a:accent6>":
        case "<a:hlink>":
        case "</a:hlink>":
        case "<a:folHlink>":
        case "</a:folHlink>":
          if (y[0].charAt(1) === "/") {
            themes.themeElements.clrScheme[
              XLSXThemeClrScheme.indexOf(y[0])
            ] = color;
            color = {};
          } else {
            color.name = y[0].slice(3, y[0].length - 1);
          }
          break;

        default:
          if (opts && opts.WTF)
            throw new Error(`Unrecognized ${y[0]} in clrScheme`);
      }
    });
  }

  /* 20.1.4.1.18 fontScheme CT_FontScheme */
  function parse_fontScheme() {}

  /* 20.1.4.1.15 fmtScheme CT_StyleMatrix */
  function parse_fmtScheme() {}

  const clrsregex = /<a:clrScheme([^>]*)>[\s\S]*<\/a:clrScheme>/;
  const fntsregex = /<a:fontScheme([^>]*)>[\s\S]*<\/a:fontScheme>/;
  const fmtsregex = /<a:fmtScheme([^>]*)>[\s\S]*<\/a:fmtScheme>/;

  /* 20.1.6.10 themeElements CT_BaseStyles */
  function parse_themeElements(data, themes, opts) {
    themes.themeElements = {};

    let t;

    [
      /* clrScheme CT_ColorScheme */
      ["clrScheme", clrsregex, parse_clrScheme],
      /* fontScheme CT_FontScheme */
      ["fontScheme", fntsregex, parse_fontScheme],
      /* fmtScheme CT_StyleMatrix */
      ["fmtScheme", fmtsregex, parse_fmtScheme]
    ].forEach(function(m) {
      if (!(t = data.match(m[1])))
        throw new Error(`${m[0]} not found in themeElements`);
      m[2](t, themes, opts);
    });
  }

  const themeltregex = /<a:themeElements([^>]*)>[\s\S]*<\/a:themeElements>/;

  /* 14.2.7 Theme Part */
  function parse_theme_xml(data, opts) {
    /* 20.1.6.9 theme CT_OfficeStyleSheet */
    if (!data || data.length === 0) return parse_theme_xml(write_theme());

    let t;
    const themes = {};

    /* themeElements CT_BaseStyles */
    if (!(t = data.match(themeltregex)))
      throw new Error("themeElements not found in theme");
    parse_themeElements(t[0], themes, opts);
    themes.raw = data;
    return themes;
  }

  function write_theme(Themes, opts) {
    if (opts && opts.themeXLSX) return opts.themeXLSX;
    if (Themes && typeof Themes.raw === "string") return Themes.raw;
    const o = [XML_HEADER];
    o[o.length] =
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">';
    o[o.length] = "<a:themeElements>";

    o[o.length] = '<a:clrScheme name="Office">';
    o[o.length] =
      '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>';
    o[o.length] = '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>';
    o[o.length] = '<a:dk2><a:srgbClr val="1F497D"/></a:dk2>';
    o[o.length] = '<a:lt2><a:srgbClr val="EEECE1"/></a:lt2>';
    o[o.length] = '<a:accent1><a:srgbClr val="4F81BD"/></a:accent1>';
    o[o.length] = '<a:accent2><a:srgbClr val="C0504D"/></a:accent2>';
    o[o.length] = '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3>';
    o[o.length] = '<a:accent4><a:srgbClr val="8064A2"/></a:accent4>';
    o[o.length] = '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5>';
    o[o.length] = '<a:accent6><a:srgbClr val="F79646"/></a:accent6>';
    o[o.length] = '<a:hlink><a:srgbClr val="0000FF"/></a:hlink>';
    o[o.length] = '<a:folHlink><a:srgbClr val="800080"/></a:folHlink>';
    o[o.length] = "</a:clrScheme>";

    o[o.length] = '<a:fontScheme name="Office">';
    o[o.length] = "<a:majorFont>";
    o[o.length] = '<a:latin typeface="Cambria"/>';
    o[o.length] = '<a:ea typeface=""/>';
    o[o.length] = '<a:cs typeface=""/>';
    o[o.length] = '<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>';
    o[o.length] = '<a:font script="Hang" typeface="맑은 고딕"/>';
    o[o.length] = '<a:font script="Hans" typeface="宋体"/>';
    o[o.length] = '<a:font script="Hant" typeface="新細明體"/>';
    o[o.length] = '<a:font script="Arab" typeface="Times New Roman"/>';
    o[o.length] = '<a:font script="Hebr" typeface="Times New Roman"/>';
    o[o.length] = '<a:font script="Thai" typeface="Tahoma"/>';
    o[o.length] = '<a:font script="Ethi" typeface="Nyala"/>';
    o[o.length] = '<a:font script="Beng" typeface="Vrinda"/>';
    o[o.length] = '<a:font script="Gujr" typeface="Shruti"/>';
    o[o.length] = '<a:font script="Khmr" typeface="MoolBoran"/>';
    o[o.length] = '<a:font script="Knda" typeface="Tunga"/>';
    o[o.length] = '<a:font script="Guru" typeface="Raavi"/>';
    o[o.length] = '<a:font script="Cans" typeface="Euphemia"/>';
    o[o.length] = '<a:font script="Cher" typeface="Plantagenet Cherokee"/>';
    o[o.length] = '<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>';
    o[o.length] = '<a:font script="Tibt" typeface="Microsoft Himalaya"/>';
    o[o.length] = '<a:font script="Thaa" typeface="MV Boli"/>';
    o[o.length] = '<a:font script="Deva" typeface="Mangal"/>';
    o[o.length] = '<a:font script="Telu" typeface="Gautami"/>';
    o[o.length] = '<a:font script="Taml" typeface="Latha"/>';
    o[o.length] = '<a:font script="Syrc" typeface="Estrangelo Edessa"/>';
    o[o.length] = '<a:font script="Orya" typeface="Kalinga"/>';
    o[o.length] = '<a:font script="Mlym" typeface="Kartika"/>';
    o[o.length] = '<a:font script="Laoo" typeface="DokChampa"/>';
    o[o.length] = '<a:font script="Sinh" typeface="Iskoola Pota"/>';
    o[o.length] = '<a:font script="Mong" typeface="Mongolian Baiti"/>';
    o[o.length] = '<a:font script="Viet" typeface="Times New Roman"/>';
    o[o.length] = '<a:font script="Uigh" typeface="Microsoft Uighur"/>';
    o[o.length] = '<a:font script="Geor" typeface="Sylfaen"/>';
    o[o.length] = "</a:majorFont>";
    o[o.length] = "<a:minorFont>";
    o[o.length] = '<a:latin typeface="Calibri"/>';
    o[o.length] = '<a:ea typeface=""/>';
    o[o.length] = '<a:cs typeface=""/>';
    o[o.length] = '<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>';
    o[o.length] = '<a:font script="Hang" typeface="맑은 고딕"/>';
    o[o.length] = '<a:font script="Hans" typeface="宋体"/>';
    o[o.length] = '<a:font script="Hant" typeface="新細明體"/>';
    o[o.length] = '<a:font script="Arab" typeface="Arial"/>';
    o[o.length] = '<a:font script="Hebr" typeface="Arial"/>';
    o[o.length] = '<a:font script="Thai" typeface="Tahoma"/>';
    o[o.length] = '<a:font script="Ethi" typeface="Nyala"/>';
    o[o.length] = '<a:font script="Beng" typeface="Vrinda"/>';
    o[o.length] = '<a:font script="Gujr" typeface="Shruti"/>';
    o[o.length] = '<a:font script="Khmr" typeface="DaunPenh"/>';
    o[o.length] = '<a:font script="Knda" typeface="Tunga"/>';
    o[o.length] = '<a:font script="Guru" typeface="Raavi"/>';
    o[o.length] = '<a:font script="Cans" typeface="Euphemia"/>';
    o[o.length] = '<a:font script="Cher" typeface="Plantagenet Cherokee"/>';
    o[o.length] = '<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>';
    o[o.length] = '<a:font script="Tibt" typeface="Microsoft Himalaya"/>';
    o[o.length] = '<a:font script="Thaa" typeface="MV Boli"/>';
    o[o.length] = '<a:font script="Deva" typeface="Mangal"/>';
    o[o.length] = '<a:font script="Telu" typeface="Gautami"/>';
    o[o.length] = '<a:font script="Taml" typeface="Latha"/>';
    o[o.length] = '<a:font script="Syrc" typeface="Estrangelo Edessa"/>';
    o[o.length] = '<a:font script="Orya" typeface="Kalinga"/>';
    o[o.length] = '<a:font script="Mlym" typeface="Kartika"/>';
    o[o.length] = '<a:font script="Laoo" typeface="DokChampa"/>';
    o[o.length] = '<a:font script="Sinh" typeface="Iskoola Pota"/>';
    o[o.length] = '<a:font script="Mong" typeface="Mongolian Baiti"/>';
    o[o.length] = '<a:font script="Viet" typeface="Arial"/>';
    o[o.length] = '<a:font script="Uigh" typeface="Microsoft Uighur"/>';
    o[o.length] = '<a:font script="Geor" typeface="Sylfaen"/>';
    o[o.length] = "</a:minorFont>";
    o[o.length] = "</a:fontScheme>";

    o[o.length] = '<a:fmtScheme name="Office">';
    o[o.length] = "<a:fillStyleLst>";
    o[o.length] = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
    o[o.length] = '<a:gradFill rotWithShape="1">';
    o[o.length] = "<a:gsLst>";
    o[o.length] =
      '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>';
    o[o.length] =
      '<a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>';
    o[o.length] =
      '<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
    o[o.length] = "</a:gsLst>";
    o[o.length] = '<a:lin ang="16200000" scaled="1"/>';
    o[o.length] = "</a:gradFill>";
    o[o.length] = '<a:gradFill rotWithShape="1">';
    o[o.length] = "<a:gsLst>";
    o[o.length] =
      '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="100000"/><a:shade val="100000"/><a:satMod val="130000"/></a:schemeClr></a:gs>';
    o[o.length] =
      '<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="50000"/><a:shade val="100000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
    o[o.length] = "</a:gsLst>";
    o[o.length] = '<a:lin ang="16200000" scaled="0"/>';
    o[o.length] = "</a:gradFill>";
    o[o.length] = "</a:fillStyleLst>";
    o[o.length] = "<a:lnStyleLst>";
    o[o.length] =
      '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln>';
    o[o.length] =
      '<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>';
    o[o.length] =
      '<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>';
    o[o.length] = "</a:lnStyleLst>";
    o[o.length] = "<a:effectStyleLst>";
    o[o.length] = "<a:effectStyle>";
    o[o.length] = "<a:effectLst>";
    o[o.length] =
      '<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw>';
    o[o.length] = "</a:effectLst>";
    o[o.length] = "</a:effectStyle>";
    o[o.length] = "<a:effectStyle>";
    o[o.length] = "<a:effectLst>";
    o[o.length] =
      '<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw>';
    o[o.length] = "</a:effectLst>";
    o[o.length] = "</a:effectStyle>";
    o[o.length] = "<a:effectStyle>";
    o[o.length] = "<a:effectLst>";
    o[o.length] =
      '<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw>';
    o[o.length] = "</a:effectLst>";
    o[o.length] =
      '<a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d>';
    o[o.length] = '<a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d>';
    o[o.length] = "</a:effectStyle>";
    o[o.length] = "</a:effectStyleLst>";
    o[o.length] = "<a:bgFillStyleLst>";
    o[o.length] = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
    o[o.length] = '<a:gradFill rotWithShape="1">';
    o[o.length] = "<a:gsLst>";
    o[o.length] =
      '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
    o[o.length] =
      '<a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
    o[o.length] =
      '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs>';
    o[o.length] = "</a:gsLst>";
    o[o.length] =
      '<a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path>';
    o[o.length] = "</a:gradFill>";
    o[o.length] = '<a:gradFill rotWithShape="1">';
    o[o.length] = "<a:gsLst>";
    o[o.length] =
      '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs>';
    o[o.length] =
      '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs>';
    o[o.length] = "</a:gsLst>";
    o[o.length] =
      '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>';
    o[o.length] = "</a:gradFill>";
    o[o.length] = "</a:bgFillStyleLst>";
    o[o.length] = "</a:fmtScheme>";
    o[o.length] = "</a:themeElements>";

    o[o.length] = "<a:objectDefaults>";
    o[o.length] = "<a:spDef>";
    o[o.length] =
      '<a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></a:style>';
    o[o.length] = "</a:spDef>";
    o[o.length] = "<a:lnDef>";
    o[o.length] =
      '<a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></a:style>';
    o[o.length] = "</a:lnDef>";
    o[o.length] = "</a:objectDefaults>";
    o[o.length] = "<a:extraClrSchemeLst/>";
    o[o.length] = "</a:theme>";
    return o.join("");
  }
  /* [MS-XLS] 2.4.326 TODO: payload is a zip file */
  function parse_Theme(blob, length, opts) {
    const end = blob.l + length;
    const dwThemeVersion = blob.read_shift(4);
    if (dwThemeVersion === 124226) return;
    if (!opts.cellStyles || !jszip) {
      blob.l = end;
      return;
    }
    const data = blob.slice(blob.l);
    blob.l = end;
    let zip;
    try {
      zip = new jszip(data);
    } catch (e) {
      return;
    }
    const themeXML = getzipstr(zip, "theme/theme/theme1.xml", true);
    if (!themeXML) return;
    return parse_theme_xml(themeXML, opts);
  }

  /* 2.5.49 */
  function parse_ColorTheme(blob) {
    return blob.read_shift(4);
  }

  /* 2.5.155 */
  function parse_FullColorExt(blob) {
    const o = {};
    o.xclrType = blob.read_shift(2);
    o.nTintShade = blob.read_shift(2);
    switch (o.xclrType) {
      case 0:
        blob.l += 4;
        break;
      case 1:
        o.xclrValue = parse_IcvXF(blob, 4);
        break;
      case 2:
        o.xclrValue = parse_LongRGBA(blob, 4);
        break;
      case 3:
        o.xclrValue = parse_ColorTheme(blob, 4);
        break;
      case 4:
        blob.l += 4;
        break;
    }
    blob.l += 8;
    return o;
  }

  /* 2.5.164 TODO: read 7 bits */
  function parse_IcvXF(blob, length) {
    return parsenoop(blob, length);
  }

  /* 2.5.280 */
  function parse_XFExtGradient(blob, length) {
    return parsenoop(blob, length);
  }

  /* [MS-XLS] 2.5.108 */
  function parse_ExtProp(blob) {
    const extType = blob.read_shift(2);
    const cb = blob.read_shift(2) - 4;
    const o = [extType];
    switch (extType) {
      case 0x04:
      case 0x05:
      case 0x07:
      case 0x08:
      case 0x09:
      case 0x0a:
      case 0x0b:
      case 0x0d:
        o[1] = parse_FullColorExt(blob, cb);
        break;
      case 0x06:
        o[1] = parse_XFExtGradient(blob, cb);
        break;
      case 0x0e:
      case 0x0f:
        o[1] = blob.read_shift(cb === 1 ? 1 : 2);
        break;
      default:
        throw new Error(`Unrecognized ExtProp type: ${extType} ${cb}`);
    }
    return o;
  }

  /* 2.4.355 */
  function parse_XFExt(blob, length) {
    const end = blob.l + length;
    blob.l += 2;
    const ixfe = blob.read_shift(2);
    blob.l += 2;
    let cexts = blob.read_shift(2);
    const ext = [];
    while (cexts-- > 0) ext.push(parse_ExtProp(blob, end - blob.l));
    return { ixfe, ext };
  }

  /* xf is an XF, see parse_XFExt for xfext */
  function update_xfext(xf, xfext) {
    xfext.forEach(function(xfe) {
      switch (xfe[0] /* 2.5.108 extPropData */) {
        case 0x04:
          break; /* foreground color */
        case 0x05:
          break; /* background color */
        case 0x06:
          break; /* gradient fill */
        case 0x07:
          break; /* top cell border color */
        case 0x08:
          break; /* bottom cell border color */
        case 0x09:
          break; /* left cell border color */
        case 0x0a:
          break; /* right cell border color */
        case 0x0b:
          break; /* diagonal cell border color */
        case 0x0d /* text color */:
          break;
        case 0x0e:
          break; /* font scheme */
        case 0x0f:
          break; /* indentation level */
      }
    });
  }

  /* 18.6 Calculation Chain */
  function parse_cc_xml(data) {
    const d = [];
    if (!data) return d;
    let i = 1;
    (data.match(tagregex) || []).forEach(function(x) {
      const y = parsexmltag(x);
      switch (y[0]) {
        case "<?xml":
          break;
        /* 18.6.2  calcChain CT_CalcChain 1 */
        case "<calcChain":
        case "<calcChain>":
        case "</calcChain>":
          break;
        /* 18.6.1  c CT_CalcCell 1 */
        case "<c":
          delete y[0];
          if (y.i) i = y.i;
          else y.i = i;
          d.push(y);
          break;
      }
    });
    return d;
  }

  // function write_cc_xml(data, opts) { }

  /* [MS-XLSB] 2.6.4.1 */
  function parse_BrtCalcChainItem$(data) {
    const out = {};
    out.i = data.read_shift(4);
    const cell = {};
    cell.r = data.read_shift(4);
    cell.c = data.read_shift(4);
    out.r = encode_cell(cell);
    const flags = data.read_shift(1);
    if (flags & 0x2) out.l = "1";
    if (flags & 0x8) out.a = "1";
    return out;
  }

  /* 18.6 Calculation Chain */
  function parse_cc_bin(data, name, opts) {
    const out = [];
    const pass = false;
    recordhopper(data, function hopper_cc(val, R_n, RT) {
      switch (RT) {
        case 0x003f /* 'BrtCalcChainItem$' */:
          out.push(val);
          break;

        default:
          if ((R_n || "").indexOf("Begin") > 0) {
            /* empty */
          } else if ((R_n || "").indexOf("End") > 0) {
            /* empty */
          } else if (!pass || opts.WTF)
            throw new Error(`Unexpected record ${RT} ${R_n}`);
      }
    });
    return out;
  }

  // function write_cc_bin(data, opts) { }
  /* 18.14 Supplementary Workbook Data */
  function parse_xlink_xml() {
    // var opts = _opts || {};
    // if(opts.WTF) throw "XLSX External Link";
  }

  /* [MS-XLSB] 2.1.7.25 External Link */
  function parse_xlink_bin(data, rel, name, _opts) {
    if (!data) return data;
    const opts = _opts || {};

    let pass = false;
    const end = false;

    recordhopper(
      data,
      function xlink_parse(val, R_n, RT) {
        if (end) return;
        switch (RT) {
          case 0x0167: /* 'BrtSupTabs' */
          case 0x016b: /* 'BrtExternTableStart' */
          case 0x016c: /* 'BrtExternTableEnd' */
          case 0x016e: /* 'BrtExternRowHdr' */
          case 0x016f: /* 'BrtExternCellBlank' */
          case 0x0170: /* 'BrtExternCellReal' */
          case 0x0171: /* 'BrtExternCellBool' */
          case 0x0172: /* 'BrtExternCellError' */
          case 0x0173: /* 'BrtExternCellString' */
          case 0x01d8: /* 'BrtExternValueMeta' */
          case 0x0241: /* 'BrtSupNameStart' */
          case 0x0242: /* 'BrtSupNameValueStart' */
          case 0x0243: /* 'BrtSupNameValueEnd' */
          case 0x0244: /* 'BrtSupNameNum' */
          case 0x0245: /* 'BrtSupNameErr' */
          case 0x0246: /* 'BrtSupNameSt' */
          case 0x0247: /* 'BrtSupNameNil' */
          case 0x0248: /* 'BrtSupNameBool' */
          case 0x0249: /* 'BrtSupNameFmla' */
          case 0x024a: /* 'BrtSupNameBits' */
          case 0x024b /* 'BrtSupNameEnd' */:
            break;

          case 0x0023 /* 'BrtFRTBegin' */:
            pass = true;
            break;
          case 0x0024 /* 'BrtFRTEnd' */:
            pass = false;
            break;

          default:
            if ((R_n || "").indexOf("Begin") > 0) {
              /* empty */
            } else if ((R_n || "").indexOf("End") > 0) {
              /* empty */
            } else if (!pass || opts.WTF)
              throw new Error(`Unexpected record ${RT.toString(16)} ${R_n}`);
        }
      },
      opts
    );
  }
  /* 20.5 DrawingML - SpreadsheetML Drawing */
  RELS.IMG =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
  RELS.DRAW =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

  /* 20.5.2.35 wsDr CT_Drawing */
  function parse_drawing(data, rels) {
    if (!data) return "??";
    /*
	  Chartsheet Drawing:
	   - 20.5.2.35 wsDr CT_Drawing
	    - 20.5.2.1  absoluteAnchor CT_AbsoluteAnchor
	     - 20.5.2.16 graphicFrame CT_GraphicalObjectFrame
	      - 20.1.2.2.16 graphic CT_GraphicalObject
	       - 20.1.2.2.17 graphicData CT_GraphicalObjectData
          - chart reference
	   the actual type is based on the URI of the graphicData
		TODO: handle embedded charts and other types of graphics
	*/
    const id = (data.match(/<c:chart [^>]*r:id="([^"]*)"/) || ["", ""])[1];

    return rels["!id"][id].Target;
  }

  /* L.5.5.2 SpreadsheetML Comments + VML Schema */
  RELS.CMNT =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

  function sheet_insert_comments(sheet, comments) {
    const dense = Array.isArray(sheet);
    let cell;
    comments.forEach(function(comment) {
      const r = decode_cell(comment.ref);
      if (dense) {
        if (!sheet[r.r]) sheet[r.r] = [];
        cell = sheet[r.r][r.c];
      } else cell = sheet[comment.ref];
      if (!cell) {
        cell = { t: "z" };
        if (dense) sheet[r.r][r.c] = cell;
        else sheet[comment.ref] = cell;
        const range = safe_decode_range(sheet["!ref"] || "BDWGO1000001:A1");
        if (range.s.r > r.r) range.s.r = r.r;
        if (range.e.r < r.r) range.e.r = r.r;
        if (range.s.c > r.c) range.s.c = r.c;
        if (range.e.c < r.c) range.e.c = r.c;
        const encoded = encode_range(range);
        if (encoded !== sheet["!ref"]) sheet["!ref"] = encoded;
      }

      if (!cell.c) cell.c = [];
      const o = { a: comment.author, t: comment.t, r: comment.r };
      if (comment.h) o.h = comment.h;
      cell.c.push(o);
    });
  }

  /* 18.7 Comments */
  function parse_comments_xml(data, opts) {
    /* 18.7.6 CT_Comments */
    if (data.match(/<(?:\w+:)?comments *\/>/)) return [];
    const authors = [];
    const commentList = [];
    const authtag = data.match(
      /<(?:\w+:)?authors>([\s\S]*)<\/(?:\w+:)?authors>/
    );
    if (authtag && authtag[1])
      authtag[1].split(/<\/\w*:?author>/).forEach(function(x) {
        if (x === "" || x.trim() === "") return;
        const a = x.match(/<(?:\w+:)?author[^>]*>(.*)/);
        if (a) authors.push(a[1]);
      });
    const cmnttag = data.match(
      /<(?:\w+:)?commentList>([\s\S]*)<\/(?:\w+:)?commentList>/
    );
    if (cmnttag && cmnttag[1])
      cmnttag[1].split(/<\/\w*:?comment>/).forEach(function(x) {
        if (x === "" || x.trim() === "") return;
        const cm = x.match(/<(?:\w+:)?comment[^>]*>/);
        if (!cm) return;
        const y = parsexmltag(cm[0]);
        const comment = {
          author: (y.authorId && authors[y.authorId]) || "sheetjsghost",
          ref: y.ref,
          guid: y.guid
        };
        const cell = decode_cell(y.ref);
        if (opts.sheetRows && opts.sheetRows <= cell.r) return;
        const textMatch = x.match(/<(?:\w+:)?text>([\s\S]*)<\/(?:\w+:)?text>/);
        const rt = (!!textMatch &&
          !!textMatch[1] &&
          parse_si(textMatch[1])) || { r: "", t: "", h: "" };
        comment.r = rt.r;
        if (rt.r == "<t></t>") rt.t = rt.h = "";
        comment.t = rt.t.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
        if (opts.cellHTML) comment.h = rt.h;
        commentList.push(comment);
      });
    return commentList;
  }

  /* [MS-XLSB] 2.4.28 BrtBeginComment */
  function parse_BrtBeginComment(data) {
    const out = {};
    out.iauthor = data.read_shift(4);
    const rfx = parse_UncheckedRfX(data, 16);
    out.rfx = rfx.s;
    out.ref = encode_cell(rfx.s);
    data.l += 16; /* var guid = parse_GUID(data); */
    return out;
  }

  /* [MS-XLSB] 2.4.327 BrtCommentAuthor */
  const parse_BrtCommentAuthor = parse_XLWideString;

  /* [MS-XLSB] 2.1.7.8 Comments */
  function parse_comments_bin(data, opts) {
    const out = [];
    const authors = [];
    let c = {};
    let pass = false;
    recordhopper(data, function hopper_cmnt(val, R_n, RT) {
      switch (RT) {
        case 0x0278 /* 'BrtCommentAuthor' */:
          authors.push(val);
          break;
        case 0x027b /* 'BrtBeginComment' */:
          c = val;
          break;
        case 0x027d /* 'BrtCommentText' */:
          c.t = val.t;
          c.h = val.h;
          c.r = val.r;
          break;
        case 0x027c /* 'BrtEndComment' */:
          c.author = authors[c.iauthor];
          delete c.iauthor;
          if (opts.sheetRows && c.rfx && opts.sheetRows <= c.rfx.r) break;
          if (!c.t) c.t = "";
          delete c.rfx;
          out.push(c);
          break;

        case 0x0c00 /* 'BrtUid' */:
          break;

        case 0x0023 /* 'BrtFRTBegin' */:
          pass = true;
          break;
        case 0x0024 /* 'BrtFRTEnd' */:
          pass = false;
          break;
        case 0x0025:
          /* 'BrtACBegin' */ break;
        case 0x0026:
          /* 'BrtACEnd' */ break;

        default:
          if ((R_n || "").indexOf("Begin") > 0) {
            /* empty */
          } else if ((R_n || "").indexOf("End") > 0) {
            /* empty */
          } else if (!pass || opts.WTF)
            throw new Error(`Unexpected record ${RT} ${R_n}`);
      }
    });
    return out;
  }

  const CT_VBA = "application/vnd.ms-office.vbaProject";
  function make_vba_xls(cfb) {
    const newcfb = CFB.utils.cfb_new({ root: "R" });
    cfb.FullPaths.forEach(function(p, i) {
      if (p.slice(-1) === "/" || !p.match(/_VBA_PROJECT_CUR/)) return;
      const newpath = p
        .replace(/^[^\/]*/, "R")
        .replace(/\/_VBA_PROJECT_CUR\u0000*/, "");
      CFB.utils.cfb_add(newcfb, newpath, cfb.FileIndex[i].content);
    });
    return CFB.write(newcfb);
  }

  RELS.DS =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
  RELS.MS =
    "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";

  /* macro and dialog sheet stubs */
  function parse_ds_bin() {
    return { "!type": "dialog" };
  }
  function parse_ds_xml() {
    return { "!type": "dialog" };
  }
  function parse_ms_bin() {
    return { "!type": "macro" };
  }
  function parse_ms_xml() {
    return { "!type": "macro" };
  }
  /* TODO: it will be useful to parse the function str */
  var rc_to_a1 = (function() {
    const rcregex = /(^|[^A-Za-z_])R(\[?-?\d+\]|[1-9]\d*|)C(\[?-?\d+\]|[1-9]\d*|)(?![A-Za-z0-9_])/g;
    let rcbase = { r: 0, c: 0 };
    function rcfunc($$, $1, $2, $3) {
      let cRel = false;
      let rRel = false;

      if ($2.length == 0) rRel = true;
      else if ($2.charAt(0) == "[") {
        rRel = true;
        $2 = $2.slice(1, -1);
      }

      if ($3.length == 0) cRel = true;
      else if ($3.charAt(0) == "[") {
        cRel = true;
        $3 = $3.slice(1, -1);
      }

      let R = $2.length > 0 ? parseInt($2, 10) | 0 : 0;
      let C = $3.length > 0 ? parseInt($3, 10) | 0 : 0;

      if (cRel) C += rcbase.c;
      else --C;
      if (rRel) R += rcbase.r;
      else --R;
      return (
        $1 +
        (cRel ? "" : "$") +
        encode_col(C) +
        (rRel ? "" : "$") +
        encode_row(R)
      );
    }
    return function rc_to_a1(fstr, base) {
      rcbase = base;
      return fstr.replace(rcregex, rcfunc);
    };
  })();

  const crefregex = /(^|[^._A-Z0-9])([$]?)([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])([$]?)(10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})(?![_.\(A-Za-z0-9])/g;
  var a1_to_rc = (function() {
    return function a1_to_rc(fstr, base) {
      return fstr.replace(crefregex, function($0, $1, $2, $3, $4, $5) {
        const c = decode_col($3) - ($2 ? 0 : base.c);
        const r = decode_row($5) - ($4 ? 0 : base.r);
        const R = r == 0 ? "" : !$4 ? `[${r}]` : r + 1;
        const C = c == 0 ? "" : !$2 ? `[${c}]` : c + 1;
        return `${$1}R${R}C${C}`;
      });
    };
  })();

  /* no defined name can collide with a valid cell address A1:XFD1048576 ... except LOG10! */
  function shift_formula_str(f, delta) {
    return f.replace(crefregex, function($0, $1, $2, $3, $4, $5) {
      return (
        $1 +
        ($2 == "$" ? $2 + $3 : encode_col(decode_col($3) + delta.c)) +
        ($4 == "$" ? $4 + $5 : encode_row(decode_row($5) + delta.r))
      );
    });
  }

  function shift_formula_xlsx(f, range, cell) {
    const r = decode_range(range);
    const { s } = r;
    const c = decode_cell(cell);
    const delta = { r: c.r - s.r, c: c.c - s.c };
    return shift_formula_str(f, delta);
  }

  /* TODO: parse formula */
  function fuzzyfmla(f) {
    if (f.length == 1) return false;
    return true;
  }

  function _xlfn(f) {
    return f.replace(/_xlfn\./g, "");
  }
  function parseread1(blob) {
    blob.l += 1;
  }

  /* [MS-XLS] 2.5.51 */
  function parse_ColRelU(blob, length) {
    const c = blob.read_shift(length == 1 ? 1 : 2);
    return [c & 0x3fff, (c >> 14) & 1, (c >> 15) & 1];
  }

  /* [MS-XLS] 2.5.198.105 ; [MS-XLSB] 2.5.97.89 */
  function parse_RgceArea(blob, length, opts) {
    let w = 2;
    if (opts) {
      if (opts.biff >= 2 && opts.biff <= 5)
        return parse_RgceArea_BIFF2(blob, length, opts);
      else if (opts.biff == 12) w = 4;
    }
    const r = blob.read_shift(w);
    const R = blob.read_shift(w);
    const c = parse_ColRelU(blob, 2);
    const C = parse_ColRelU(blob, 2);
    return {
      s: { r, c: c[0], cRel: c[1], rRel: c[2] },
      e: { r: R, c: C[0], cRel: C[1], rRel: C[2] }
    };
  }
  /* BIFF 2-5 encodes flags in the row field */
  function parse_RgceArea_BIFF2(blob) {
    const r = parse_ColRelU(blob, 2);
    const R = parse_ColRelU(blob, 2);
    const c = blob.read_shift(1);
    const C = blob.read_shift(1);
    return {
      s: { r: r[0], c, cRel: r[1], rRel: r[2] },
      e: { r: R[0], c: C, cRel: R[1], rRel: R[2] }
    };
  }

  /* [MS-XLS] 2.5.198.105 ; [MS-XLSB] 2.5.97.90 */
  function parse_RgceAreaRel(blob, length, opts) {
    if (opts.biff < 8) return parse_RgceArea_BIFF2(blob, length, opts);
    const r = blob.read_shift(opts.biff == 12 ? 4 : 2);
    const R = blob.read_shift(opts.biff == 12 ? 4 : 2);
    const c = parse_ColRelU(blob, 2);
    const C = parse_ColRelU(blob, 2);
    return {
      s: { r, c: c[0], cRel: c[1], rRel: c[2] },
      e: { r: R, c: C[0], cRel: C[1], rRel: C[2] }
    };
  }

  /* [MS-XLS] 2.5.198.109 ; [MS-XLSB] 2.5.97.91 */
  function parse_RgceLoc(blob, length, opts) {
    if (opts && opts.biff >= 2 && opts.biff <= 5)
      return parse_RgceLoc_BIFF2(blob, length, opts);
    const r = blob.read_shift(opts && opts.biff == 12 ? 4 : 2);
    const c = parse_ColRelU(blob, 2);
    return { r, c: c[0], cRel: c[1], rRel: c[2] };
  }
  function parse_RgceLoc_BIFF2(blob) {
    const r = parse_ColRelU(blob, 2);
    const c = blob.read_shift(1);
    return { r: r[0], c, cRel: r[1], rRel: r[2] };
  }

  /* [MS-XLS] 2.5.198.107, 2.5.47 */
  function parse_RgceElfLoc(blob) {
    const r = blob.read_shift(2);
    const c = blob.read_shift(2);
    return {
      r,
      c: c & 0xff,
      fQuoted: !!(c & 0x4000),
      cRel: c >> 15,
      rRel: c >> 15
    };
  }

  /* [MS-XLS] 2.5.198.111 ; [MS-XLSB] 2.5.97.92 TODO */
  function parse_RgceLocRel(blob, length, opts) {
    const biff = opts && opts.biff ? opts.biff : 8;
    if (biff >= 2 && biff <= 5)
      return parse_RgceLocRel_BIFF2(blob, length, opts);
    let r = blob.read_shift(biff >= 12 ? 4 : 2);
    let cl = blob.read_shift(2);
    const cRel = (cl & 0x4000) >> 14;
    const rRel = (cl & 0x8000) >> 15;
    cl &= 0x3fff;
    if (rRel == 1) while (r > 0x7ffff) r -= 0x100000;
    if (cRel == 1) while (cl > 0x1fff) cl -= 0x4000;
    return { r, c: cl, cRel, rRel };
  }
  function parse_RgceLocRel_BIFF2(blob) {
    let rl = blob.read_shift(2);
    let c = blob.read_shift(1);
    const rRel = (rl & 0x8000) >> 15;
    const cRel = (rl & 0x4000) >> 14;
    rl &= 0x3fff;
    if (rRel == 1 && rl >= 0x2000) rl -= 0x4000;
    if (cRel == 1 && c >= 0x80) c -= 0x100;
    return { r: rl, c, cRel, rRel };
  }

  /* [MS-XLS] 2.5.198.27 ; [MS-XLSB] 2.5.97.18 */
  function parse_PtgArea(blob, length, opts) {
    const type = (blob[blob.l++] & 0x60) >> 5;
    const area = parse_RgceArea(
      blob,
      opts.biff >= 2 && opts.biff <= 5 ? 6 : 8,
      opts
    );
    return [type, area];
  }

  /* [MS-XLS] 2.5.198.28 ; [MS-XLSB] 2.5.97.19 */
  function parse_PtgArea3d(blob, length, opts) {
    const type = (blob[blob.l++] & 0x60) >> 5;
    const ixti = blob.read_shift(2, "i");
    let w = 8;
    if (opts)
      switch (opts.biff) {
        case 5:
          blob.l += 12;
          w = 6;
          break;
        case 12:
          w = 12;
          break;
      }
    const area = parse_RgceArea(blob, w, opts);
    return [type, ixti, area];
  }

  /* [MS-XLS] 2.5.198.29 ; [MS-XLSB] 2.5.97.20 */
  function parse_PtgAreaErr(blob, length, opts) {
    const type = (blob[blob.l++] & 0x60) >> 5;
    blob.l += opts && opts.biff > 8 ? 12 : opts.biff < 8 ? 6 : 8;
    return [type];
  }
  /* [MS-XLS] 2.5.198.30 ; [MS-XLSB] 2.5.97.21 */
  function parse_PtgAreaErr3d(blob, length, opts) {
    const type = (blob[blob.l++] & 0x60) >> 5;
    const ixti = blob.read_shift(2);
    let w = 8;
    if (opts)
      switch (opts.biff) {
        case 5:
          blob.l += 12;
          w = 6;
          break;
        case 12:
          w = 12;
          break;
      }
    blob.l += w;
    return [type, ixti];
  }

  /* [MS-XLS] 2.5.198.31 ; [MS-XLSB] 2.5.97.22 */
  function parse_PtgAreaN(blob, length, opts) {
    const type = (blob[blob.l++] & 0x60) >> 5;
    const area = parse_RgceAreaRel(blob, length - 1, opts);
    return [type, area];
  }

  /* [MS-XLS] 2.5.198.32 ; [MS-XLSB] 2.5.97.23 */
  function parse_PtgArray(blob, length, opts) {
    const type = (blob[blob.l++] & 0x60) >> 5;
    blob.l += opts.biff == 2 ? 6 : opts.biff == 12 ? 14 : 7;
    return [type];
  }

  /* [MS-XLS] 2.5.198.33 ; [MS-XLSB] 2.5.97.24 */
  function parse_PtgAttrBaxcel(blob) {
    const bitSemi = blob[blob.l + 1] & 0x01; /* 1 = volatile */
    const bitBaxcel = 1;
    blob.l += 4;
    return [bitSemi, bitBaxcel];
  }

  /* [MS-XLS] 2.5.198.34 ; [MS-XLSB] 2.5.97.25 */
  function parse_PtgAttrChoose(blob, length, opts) {
    blob.l += 2;
    const offset = blob.read_shift(opts && opts.biff == 2 ? 1 : 2);
    const o = [];
    /* offset is 1 less than the number of elements */
    for (let i = 0; i <= offset; ++i)
      o.push(blob.read_shift(opts && opts.biff == 2 ? 1 : 2));
    return o;
  }

  /* [MS-XLS] 2.5.198.35 ; [MS-XLSB] 2.5.97.26 */
  function parse_PtgAttrGoto(blob, length, opts) {
    const bitGoto = blob[blob.l + 1] & 0xff ? 1 : 0;
    blob.l += 2;
    return [bitGoto, blob.read_shift(opts && opts.biff == 2 ? 1 : 2)];
  }

  /* [MS-XLS] 2.5.198.36 ; [MS-XLSB] 2.5.97.27 */
  function parse_PtgAttrIf(blob, length, opts) {
    const bitIf = blob[blob.l + 1] & 0xff ? 1 : 0;
    blob.l += 2;
    return [bitIf, blob.read_shift(opts && opts.biff == 2 ? 1 : 2)];
  }

  /* [MS-XLSB] 2.5.97.28 */
  function parse_PtgAttrIfError(blob) {
    const bitIf = blob[blob.l + 1] & 0xff ? 1 : 0;
    blob.l += 2;
    return [bitIf, blob.read_shift(2)];
  }

  /* [MS-XLS] 2.5.198.37 ; [MS-XLSB] 2.5.97.29 */
  function parse_PtgAttrSemi(blob, length, opts) {
    const bitSemi = blob[blob.l + 1] & 0xff ? 1 : 0;
    blob.l += opts && opts.biff == 2 ? 3 : 4;
    return [bitSemi];
  }

  /* [MS-XLS] 2.5.198.40 ; [MS-XLSB] 2.5.97.32 */
  function parse_PtgAttrSpaceType(blob) {
    const type = blob.read_shift(1);
    const cch = blob.read_shift(1);
    return [type, cch];
  }

  /* [MS-XLS] 2.5.198.38 ; [MS-XLSB] 2.5.97.30 */
  function parse_PtgAttrSpace(blob) {
    blob.read_shift(2);
    return parse_PtgAttrSpaceType(blob, 2);
  }

  /* [MS-XLS] 2.5.198.39 ; [MS-XLSB] 2.5.97.31 */
  function parse_PtgAttrSpaceSemi(blob) {
    blob.read_shift(2);
    return parse_PtgAttrSpaceType(blob, 2);
  }

  /* [MS-XLS] 2.5.198.84 ; [MS-XLSB] 2.5.97.68 TODO */
  function parse_PtgRef(blob, length, opts) {
    // var ptg = blob[blob.l] & 0x1F;
    const type = (blob[blob.l] & 0x60) >> 5;
    blob.l += 1;
    const loc = parse_RgceLoc(blob, 0, opts);
    return [type, loc];
  }

  /* [MS-XLS] 2.5.198.88 ; [MS-XLSB] 2.5.97.72 TODO */
  function parse_PtgRefN(blob, length, opts) {
    const type = (blob[blob.l] & 0x60) >> 5;
    blob.l += 1;
    const loc = parse_RgceLocRel(blob, 0, opts);
    return [type, loc];
  }

  /* [MS-XLS] 2.5.198.85 ; [MS-XLSB] 2.5.97.69 TODO */
  function parse_PtgRef3d(blob, length, opts) {
    const type = (blob[blob.l] & 0x60) >> 5;
    blob.l += 1;
    const ixti = blob.read_shift(2); // XtiIndex
    if (opts && opts.biff == 5) blob.l += 12;
    const loc = parse_RgceLoc(blob, 0, opts); // TODO: or RgceLocRel
    return [type, ixti, loc];
  }

  /* [MS-XLS] 2.5.198.62 ; [MS-XLSB] 2.5.97.45 TODO */
  function parse_PtgFunc(blob, length, opts) {
    // var ptg = blob[blob.l] & 0x1F;
    const type = (blob[blob.l] & 0x60) >> 5;
    blob.l += 1;
    const iftab = blob.read_shift(opts && opts.biff <= 3 ? 1 : 2);
    return [FtabArgc[iftab], Ftab[iftab], type];
  }
  /* [MS-XLS] 2.5.198.63 ; [MS-XLSB] 2.5.97.46 TODO */
  function parse_PtgFuncVar(blob, length, opts) {
    const type = blob[blob.l++];
    const cparams = blob.read_shift(1);
    const tab =
      opts && opts.biff <= 3
        ? [type == 0x58 ? -1 : 0, blob.read_shift(1)]
        : parsetab(blob);
    return [cparams, (tab[0] === 0 ? Ftab : Cetab)[tab[1]]];
  }

  function parsetab(blob) {
    return [blob[blob.l + 1] >> 7, blob.read_shift(2) & 0x7fff];
  }

  /* [MS-XLS] 2.5.198.41 ; [MS-XLSB] 2.5.97.33 */
  function parse_PtgAttrSum(blob, length, opts) {
    blob.l += opts && opts.biff == 2 ? 3 : 4;
  }

  /* [MS-XLS] 2.5.198.58 ; [MS-XLSB] 2.5.97.40 */
  function parse_PtgExp(blob, length, opts) {
    blob.l++;
    if (opts && opts.biff == 12) return [blob.read_shift(4, "i"), 0];
    const row = blob.read_shift(2);
    const col = blob.read_shift(opts && opts.biff == 2 ? 1 : 2);
    return [row, col];
  }

  /* [MS-XLS] 2.5.198.57 ; [MS-XLSB] 2.5.97.39 */
  function parse_PtgErr(blob) {
    blob.l++;
    return BErr[blob.read_shift(1)];
  }

  /* [MS-XLS] 2.5.198.66 ; [MS-XLSB] 2.5.97.49 */
  function parse_PtgInt(blob) {
    blob.l++;
    return blob.read_shift(2);
  }

  /* [MS-XLS] 2.5.198.42 ; [MS-XLSB] 2.5.97.34 */
  function parse_PtgBool(blob) {
    blob.l++;
    return blob.read_shift(1) !== 0;
  }

  /* [MS-XLS] 2.5.198.79 ; [MS-XLSB] 2.5.97.63 */
  function parse_PtgNum(blob) {
    blob.l++;
    return parse_Xnum(blob, 8);
  }

  /* [MS-XLS] 2.5.198.89 ; [MS-XLSB] 2.5.97.74 */
  function parse_PtgStr(blob, length, opts) {
    blob.l++;
    return parse_ShortXLUnicodeString(blob, length - 1, opts);
  }

  /* [MS-XLS] 2.5.192.112 + 2.5.192.11{3,4,5,6,7} */
  /* [MS-XLSB] 2.5.97.93 + 2.5.97.9{4,5,6,7} */
  function parse_SerAr(blob, biff) {
    const val = [blob.read_shift(1)];
    if (biff == 12)
      switch (val[0]) {
        case 0x02:
          val[0] = 0x04;
          break; /* SerBool */
        case 0x04:
          val[0] = 0x10;
          break; /* SerErr */
        case 0x00:
          val[0] = 0x01;
          break; /* SerNum */
        case 0x01:
          val[0] = 0x02;
          break; /* SerStr */
      }
    switch (val[0]) {
      case 0x04 /* SerBool -- boolean */:
        val[1] = parsebool(blob, 1) ? "TRUE" : "FALSE";
        if (biff != 12) blob.l += 7;
        break;
      case 0x25: /* appears to be an alias */
      case 0x10 /* SerErr -- error */:
        val[1] = BErr[blob[blob.l]];
        blob.l += biff == 12 ? 4 : 8;
        break;
      case 0x00 /* SerNil -- honestly, I'm not sure how to reproduce this */:
        blob.l += 8;
        break;
      case 0x01 /* SerNum -- Xnum */:
        val[1] = parse_Xnum(blob, 8);
        break;
      case 0x02 /* SerStr -- XLUnicodeString (<256 chars) */:
        val[1] = parse_XLUnicodeString2(blob, 0, {
          biff: biff > 0 && biff < 8 ? 2 : biff
        });
        break;
      default:
        throw new Error(`Bad SerAr: ${val[0]}`); /* Unreachable */
    }
    return val;
  }

  /* [MS-XLS] 2.5.198.61 ; [MS-XLSB] 2.5.97.44 */
  function parse_PtgExtraMem(blob, cce, opts) {
    const count = blob.read_shift(opts.biff == 12 ? 4 : 2);
    const out = [];
    for (let i = 0; i != count; ++i)
      out.push((opts.biff == 12 ? parse_UncheckedRfX : parse_Ref8U)(blob, 8));
    return out;
  }

  /* [MS-XLS] 2.5.198.59 ; [MS-XLSB] 2.5.97.41 */
  function parse_PtgExtraArray(blob, length, opts) {
    let rows = 0;
    let cols = 0;
    if (opts.biff == 12) {
      rows = blob.read_shift(4); // DRw
      cols = blob.read_shift(4); // DCol
    } else {
      cols = 1 + blob.read_shift(1); // DColByteU
      rows = 1 + blob.read_shift(2); // DRw
    }
    if (opts.biff >= 2 && opts.biff < 8) {
      --rows;
      if (--cols == 0) cols = 0x100;
    }
    // $FlowIgnore
    for (var i = 0, o = []; i != rows && (o[i] = []); ++i)
      for (let j = 0; j != cols; ++j) o[i][j] = parse_SerAr(blob, opts.biff);
    return o;
  }

  /* [MS-XLS] 2.5.198.76 ; [MS-XLSB] 2.5.97.60 */
  function parse_PtgName(blob, length, opts) {
    const type = (blob.read_shift(1) >>> 5) & 0x03;
    const w = !opts || opts.biff >= 8 ? 4 : 2;
    const nameindex = blob.read_shift(w);
    switch (opts.biff) {
      case 2:
        blob.l += 5;
        break;
      case 3:
      case 4:
        blob.l += 8;
        break;
      case 5:
        blob.l += 12;
        break;
    }
    return [type, 0, nameindex];
  }

  /* [MS-XLS] 2.5.198.77 ; [MS-XLSB] 2.5.97.61 */
  function parse_PtgNameX(blob, length, opts) {
    if (opts.biff == 5) return parse_PtgNameX_BIFF5(blob, length, opts);
    const type = (blob.read_shift(1) >>> 5) & 0x03;
    const ixti = blob.read_shift(2); // XtiIndex
    const nameindex = blob.read_shift(4);
    return [type, ixti, nameindex];
  }
  function parse_PtgNameX_BIFF5(blob) {
    const type = (blob.read_shift(1) >>> 5) & 0x03;
    const ixti = blob.read_shift(2, "i"); // XtiIndex
    blob.l += 8;
    const nameindex = blob.read_shift(2);
    blob.l += 12;
    return [type, ixti, nameindex];
  }

  /* [MS-XLS] 2.5.198.70 ; [MS-XLSB] 2.5.97.54 */
  function parse_PtgMemArea(blob, length, opts) {
    const type = (blob.read_shift(1) >>> 5) & 0x03;
    blob.l += opts && opts.biff == 2 ? 3 : 4;
    const cce = blob.read_shift(opts && opts.biff == 2 ? 1 : 2);
    return [type, cce];
  }

  /* [MS-XLS] 2.5.198.72 ; [MS-XLSB] 2.5.97.56 */
  function parse_PtgMemFunc(blob, length, opts) {
    const type = (blob.read_shift(1) >>> 5) & 0x03;
    const cce = blob.read_shift(opts && opts.biff == 2 ? 1 : 2);
    return [type, cce];
  }

  /* [MS-XLS] 2.5.198.86 ; [MS-XLSB] 2.5.97.69 */
  function parse_PtgRefErr(blob, length, opts) {
    const type = (blob.read_shift(1) >>> 5) & 0x03;
    blob.l += 4;
    if (opts.biff < 8) blob.l--;
    if (opts.biff == 12) blob.l += 2;
    return [type];
  }

  /* [MS-XLS] 2.5.198.87 ; [MS-XLSB] 2.5.97.71 */
  function parse_PtgRefErr3d(blob, length, opts) {
    const type = (blob[blob.l++] & 0x60) >> 5;
    const ixti = blob.read_shift(2);
    let w = 4;
    if (opts)
      switch (opts.biff) {
        case 5:
          w = 15;
          break;
        case 12:
          w = 6;
          break;
      }
    blob.l += w;
    return [type, ixti];
  }

  /* [MS-XLS] 2.5.198.71 ; [MS-XLSB] 2.5.97.55 */
  const parse_PtgMemErr = parsenoop;
  /* [MS-XLS] 2.5.198.73  ; [MS-XLSB] 2.5.97.57 */
  const parse_PtgMemNoMem = parsenoop;
  /* [MS-XLS] 2.5.198.92 */
  const parse_PtgTbl = parsenoop;

  function parse_PtgElfLoc(blob, length, opts) {
    blob.l += 2;
    return [parse_RgceElfLoc(blob, 4, opts)];
  }
  function parse_PtgElfNoop(blob) {
    blob.l += 6;
    return [];
  }
  /* [MS-XLS] 2.5.198.46 */
  const parse_PtgElfCol = parse_PtgElfLoc;
  /* [MS-XLS] 2.5.198.47 */
  const parse_PtgElfColS = parse_PtgElfNoop;
  /* [MS-XLS] 2.5.198.48 */
  const parse_PtgElfColSV = parse_PtgElfNoop;
  /* [MS-XLS] 2.5.198.49 */
  const parse_PtgElfColV = parse_PtgElfLoc;
  /* [MS-XLS] 2.5.198.50 */
  function parse_PtgElfLel(blob) {
    blob.l += 2;
    return [parseuint16(blob), blob.read_shift(2) & 0x01];
  }
  /* [MS-XLS] 2.5.198.51 */
  const parse_PtgElfRadical = parse_PtgElfLoc;
  /* [MS-XLS] 2.5.198.52 */
  const parse_PtgElfRadicalLel = parse_PtgElfLel;
  /* [MS-XLS] 2.5.198.53 */
  const parse_PtgElfRadicalS = parse_PtgElfNoop;
  /* [MS-XLS] 2.5.198.54 */
  const parse_PtgElfRw = parse_PtgElfLoc;
  /* [MS-XLS] 2.5.198.55 */
  const parse_PtgElfRwV = parse_PtgElfLoc;

  /* [MS-XLSB] 2.5.97.52 TODO */
  const PtgListRT = [
    "Data",
    "All",
    "Headers",
    "??",
    "?Data2",
    "??",
    "?DataHeaders",
    "??",
    "Totals",
    "??",
    "??",
    "??",
    "?DataTotals",
    "??",
    "??",
    "??",
    "?Current"
  ];
  function parse_PtgList(blob) {
    blob.l += 2;
    const ixti = blob.read_shift(2);
    const flags = blob.read_shift(2);
    const idx = blob.read_shift(4);
    const c = blob.read_shift(2);
    const C = blob.read_shift(2);
    const rt = PtgListRT[(flags >> 2) & 0x1f];
    return { ixti, coltype: flags & 0x3, rt, idx, c, C };
  }
  /* [MS-XLS] 2.5.198.91 ; [MS-XLSB] 2.5.97.76 */
  function parse_PtgSxName(blob) {
    blob.l += 2;
    return [blob.read_shift(4)];
  }

  /* [XLS] old spec */
  function parse_PtgSheet(blob, length, opts) {
    blob.l += 5;
    blob.l += 2;
    blob.l += opts.biff == 2 ? 1 : 4;
    return ["PTGSHEET"];
  }
  function parse_PtgEndSheet(blob, length, opts) {
    blob.l += opts.biff == 2 ? 4 : 5;
    return ["PTGENDSHEET"];
  }
  function parse_PtgMemAreaN(blob) {
    const type = (blob.read_shift(1) >>> 5) & 0x03;
    const cce = blob.read_shift(2);
    return [type, cce];
  }
  function parse_PtgMemNoMemN(blob) {
    const type = (blob.read_shift(1) >>> 5) & 0x03;
    const cce = blob.read_shift(2);
    return [type, cce];
  }
  function parse_PtgAttrNoop(blob) {
    blob.l += 4;
    return [0, 0];
  }

  /* [MS-XLS] 2.5.198.25 ; [MS-XLSB] 2.5.97.16 */
  const PtgTypes = {
    0x01: { n: "PtgExp", f: parse_PtgExp },
    0x02: { n: "PtgTbl", f: parse_PtgTbl },
    0x03: { n: "PtgAdd", f: parseread1 },
    0x04: { n: "PtgSub", f: parseread1 },
    0x05: { n: "PtgMul", f: parseread1 },
    0x06: { n: "PtgDiv", f: parseread1 },
    0x07: { n: "PtgPower", f: parseread1 },
    0x08: { n: "PtgConcat", f: parseread1 },
    0x09: { n: "PtgLt", f: parseread1 },
    0x0a: { n: "PtgLe", f: parseread1 },
    0x0b: { n: "PtgEq", f: parseread1 },
    0x0c: { n: "PtgGe", f: parseread1 },
    0x0d: { n: "PtgGt", f: parseread1 },
    0x0e: { n: "PtgNe", f: parseread1 },
    0x0f: { n: "PtgIsect", f: parseread1 },
    0x10: { n: "PtgUnion", f: parseread1 },
    0x11: { n: "PtgRange", f: parseread1 },
    0x12: { n: "PtgUplus", f: parseread1 },
    0x13: { n: "PtgUminus", f: parseread1 },
    0x14: { n: "PtgPercent", f: parseread1 },
    0x15: { n: "PtgParen", f: parseread1 },
    0x16: { n: "PtgMissArg", f: parseread1 },
    0x17: { n: "PtgStr", f: parse_PtgStr },
    0x1a: { n: "PtgSheet", f: parse_PtgSheet },
    0x1b: { n: "PtgEndSheet", f: parse_PtgEndSheet },
    0x1c: { n: "PtgErr", f: parse_PtgErr },
    0x1d: { n: "PtgBool", f: parse_PtgBool },
    0x1e: { n: "PtgInt", f: parse_PtgInt },
    0x1f: { n: "PtgNum", f: parse_PtgNum },
    0x20: { n: "PtgArray", f: parse_PtgArray },
    0x21: { n: "PtgFunc", f: parse_PtgFunc },
    0x22: { n: "PtgFuncVar", f: parse_PtgFuncVar },
    0x23: { n: "PtgName", f: parse_PtgName },
    0x24: { n: "PtgRef", f: parse_PtgRef },
    0x25: { n: "PtgArea", f: parse_PtgArea },
    0x26: { n: "PtgMemArea", f: parse_PtgMemArea },
    0x27: { n: "PtgMemErr", f: parse_PtgMemErr },
    0x28: { n: "PtgMemNoMem", f: parse_PtgMemNoMem },
    0x29: { n: "PtgMemFunc", f: parse_PtgMemFunc },
    0x2a: { n: "PtgRefErr", f: parse_PtgRefErr },
    0x2b: { n: "PtgAreaErr", f: parse_PtgAreaErr },
    0x2c: { n: "PtgRefN", f: parse_PtgRefN },
    0x2d: { n: "PtgAreaN", f: parse_PtgAreaN },
    0x2e: { n: "PtgMemAreaN", f: parse_PtgMemAreaN },
    0x2f: { n: "PtgMemNoMemN", f: parse_PtgMemNoMemN },
    0x39: { n: "PtgNameX", f: parse_PtgNameX },
    0x3a: { n: "PtgRef3d", f: parse_PtgRef3d },
    0x3b: { n: "PtgArea3d", f: parse_PtgArea3d },
    0x3c: { n: "PtgRefErr3d", f: parse_PtgRefErr3d },
    0x3d: { n: "PtgAreaErr3d", f: parse_PtgAreaErr3d },
    0xff: {}
  };
  /* These are duplicated in the PtgTypes table */
  const PtgDupes = {
    0x40: 0x20,
    0x60: 0x20,
    0x41: 0x21,
    0x61: 0x21,
    0x42: 0x22,
    0x62: 0x22,
    0x43: 0x23,
    0x63: 0x23,
    0x44: 0x24,
    0x64: 0x24,
    0x45: 0x25,
    0x65: 0x25,
    0x46: 0x26,
    0x66: 0x26,
    0x47: 0x27,
    0x67: 0x27,
    0x48: 0x28,
    0x68: 0x28,
    0x49: 0x29,
    0x69: 0x29,
    0x4a: 0x2a,
    0x6a: 0x2a,
    0x4b: 0x2b,
    0x6b: 0x2b,
    0x4c: 0x2c,
    0x6c: 0x2c,
    0x4d: 0x2d,
    0x6d: 0x2d,
    0x4e: 0x2e,
    0x6e: 0x2e,
    0x4f: 0x2f,
    0x6f: 0x2f,
    0x58: 0x22,
    0x78: 0x22,
    0x59: 0x39,
    0x79: 0x39,
    0x5a: 0x3a,
    0x7a: 0x3a,
    0x5b: 0x3b,
    0x7b: 0x3b,
    0x5c: 0x3c,
    0x7c: 0x3c,
    0x5d: 0x3d,
    0x7d: 0x3d
  };
  (function() {
    for (const y in PtgDupes) PtgTypes[y] = PtgTypes[PtgDupes[y]];
  })();

  const Ptg18 = {
    0x01: { n: "PtgElfLel", f: parse_PtgElfLel },
    0x02: { n: "PtgElfRw", f: parse_PtgElfRw },
    0x03: { n: "PtgElfCol", f: parse_PtgElfCol },
    0x06: { n: "PtgElfRwV", f: parse_PtgElfRwV },
    0x07: { n: "PtgElfColV", f: parse_PtgElfColV },
    0x0a: { n: "PtgElfRadical", f: parse_PtgElfRadical },
    0x0b: { n: "PtgElfRadicalS", f: parse_PtgElfRadicalS },
    0x0d: { n: "PtgElfColS", f: parse_PtgElfColS },
    0x0f: { n: "PtgElfColSV", f: parse_PtgElfColSV },
    0x10: { n: "PtgElfRadicalLel", f: parse_PtgElfRadicalLel },
    0x19: { n: "PtgList", f: parse_PtgList },
    0x1d: { n: "PtgSxName", f: parse_PtgSxName },
    0xff: {}
  };
  const Ptg19 = {
    0x00: { n: "PtgAttrNoop", f: parse_PtgAttrNoop },
    0x01: { n: "PtgAttrSemi", f: parse_PtgAttrSemi },
    0x02: { n: "PtgAttrIf", f: parse_PtgAttrIf },
    0x04: { n: "PtgAttrChoose", f: parse_PtgAttrChoose },
    0x08: { n: "PtgAttrGoto", f: parse_PtgAttrGoto },
    0x10: { n: "PtgAttrSum", f: parse_PtgAttrSum },
    0x20: { n: "PtgAttrBaxcel", f: parse_PtgAttrBaxcel },
    0x40: { n: "PtgAttrSpace", f: parse_PtgAttrSpace },
    0x41: { n: "PtgAttrSpaceSemi", f: parse_PtgAttrSpaceSemi },
    0x80: { n: "PtgAttrIfError", f: parse_PtgAttrIfError },
    0xff: {}
  };
  Ptg19[0x21] = Ptg19[0x20];

  /* [MS-XLS] 2.5.198.103 ; [MS-XLSB] 2.5.97.87 */
  function parse_RgbExtra(blob, length, rgce, opts) {
    if (opts.biff < 8) return parsenoop(blob, length);
    const target = blob.l + length;
    const o = [];
    for (let i = 0; i !== rgce.length; ++i) {
      switch (rgce[i][0]) {
        case "PtgArray" /* PtgArray -> PtgExtraArray */:
          rgce[i][1] = parse_PtgExtraArray(blob, 0, opts);
          o.push(rgce[i][1]);
          break;
        case "PtgMemArea" /* PtgMemArea -> PtgExtraMem */:
          rgce[i][2] = parse_PtgExtraMem(blob, rgce[i][1], opts);
          o.push(rgce[i][2]);
          break;
        case "PtgExp" /* PtgExp -> PtgExtraCol */:
          if (opts && opts.biff == 12) {
            rgce[i][1][1] = blob.read_shift(4);
            o.push(rgce[i][1]);
          }
          break;
        case "PtgList": /* TODO: PtgList -> PtgExtraList */
        case "PtgElfRadicalS": /* TODO: PtgElfRadicalS -> PtgExtraElf */
        case "PtgElfColS": /* TODO: PtgElfColS -> PtgExtraElf */
        case "PtgElfColSV" /* TODO: PtgElfColSV -> PtgExtraElf */:
          throw `Unsupported ${rgce[i][0]}`;
        default:
          break;
      }
    }
    length = target - blob.l;
    /* note: this is technically an error but Excel disregards */
    // if(target !== blob.l && blob.l !== target - length) throw new Error(target + " != " + blob.l);
    if (length !== 0) o.push(parsenoop(blob, length));
    return o;
  }

  /* [MS-XLS] 2.5.198.104 ; [MS-XLSB] 2.5.97.88 */
  function parse_Rgce(blob, length, opts) {
    const target = blob.l + length;
    let R;
    let id;
    const ptgs = [];
    while (target != blob.l) {
      length = target - blob.l;
      id = blob[blob.l];
      R = PtgTypes[id];
      if (id === 0x18 || id === 0x19)
        R = (id === 0x18 ? Ptg18 : Ptg19)[blob[blob.l + 1]];
      if (!R || !R.f) {
        /* ptgs.push */ parsenoop(blob, length);
      } else {
        ptgs.push([R.n, R.f(blob, length, opts)]);
      }
    }
    return ptgs;
  }

  function stringify_array(f) {
    const o = [];
    for (let i = 0; i < f.length; ++i) {
      const x = f[i];
      const r = [];
      for (let j = 0; j < x.length; ++j) {
        const y = x[j];
        if (y)
          switch (y[0]) {
            // TODO: handle embedded quotes
            case 0x02:
              r.push(`"${y[1].replace(/"/g, '""')}"`);
              break;
            default:
              r.push(y[1]);
          }
        else r.push("");
      }
      o.push(r.join(","));
    }
    return o.join(";");
  }

  /* [MS-XLS] 2.2.2 ; [MS-XLSB] 2.2.2 TODO */
  const PtgBinOp = {
    PtgAdd: "+",
    PtgConcat: "&",
    PtgDiv: "/",
    PtgEq: "=",
    PtgGe: ">=",
    PtgGt: ">",
    PtgLe: "<=",
    PtgLt: "<",
    PtgMul: "*",
    PtgNe: "<>",
    PtgPower: "^",
    PtgSub: "-"
  };

  // List of invalid characters needs to be tested further
  const quoteCharacters = new RegExp(/[^\w\u4E00-\u9FFF\u3040-\u30FF]/);
  function formula_quote_sheet_name(sname, opts) {
    if (!sname && !(opts && opts.biff <= 5 && opts.biff >= 2))
      throw new Error("empty sheet name");
    if (quoteCharacters.test(sname)) return `'${sname}'`;
    return sname;
  }
  function get_ixti_raw(supbooks, ixti, opts) {
    if (!supbooks) return "SH33TJSERR0";
    if (opts.biff > 8 && (!supbooks.XTI || !supbooks.XTI[ixti]))
      return supbooks.SheetNames[ixti];
    if (!supbooks.XTI) return "SH33TJSERR6";
    const XTI = supbooks.XTI[ixti];
    if (opts.biff < 8) {
      if (ixti > 10000) ixti -= 65536;
      if (ixti < 0) ixti = -ixti;
      return ixti == 0 ? "" : supbooks.XTI[ixti - 1];
    }
    if (!XTI) return "SH33TJSERR1";
    let o = "";
    if (opts.biff > 8)
      switch (supbooks[XTI[0]][0]) {
        case 0x0165 /* 'BrtSupSelf' */:
          o = XTI[1] == -1 ? "#REF" : supbooks.SheetNames[XTI[1]];
          return XTI[1] == XTI[2] ? o : `${o}:${supbooks.SheetNames[XTI[2]]}`;
        case 0x0166 /* 'BrtSupSame' */:
          if (opts.SID != null) return supbooks.SheetNames[opts.SID];
          return `SH33TJSSAME${supbooks[XTI[0]][0]}`;
        case 0x0163: /* 'BrtSupBookSrc' */
        /* falls through */
        default:
          return `SH33TJSSRC${supbooks[XTI[0]][0]}`;
      }
    switch (supbooks[XTI[0]][0][0]) {
      case 0x0401:
        o =
          XTI[1] == -1 ? "#REF" : supbooks.SheetNames[XTI[1]] || "SH33TJSERR3";
        return XTI[1] == XTI[2] ? o : `${o}:${supbooks.SheetNames[XTI[2]]}`;
      case 0x3a01:
        return supbooks[XTI[0]]
          .slice(1)
          .map(function(name) {
            return name.Name;
          })
          .join(";;"); // return "SH33TJSERR8";
      default:
        if (!supbooks[XTI[0]][0][3]) return "SH33TJSERR2";
        o =
          XTI[1] == -1
            ? "#REF"
            : supbooks[XTI[0]][0][3][XTI[1]] || "SH33TJSERR4";
        return XTI[1] == XTI[2] ? o : `${o}:${supbooks[XTI[0]][0][3][XTI[2]]}`;
    }
  }
  function get_ixti(supbooks, ixti, opts) {
    return formula_quote_sheet_name(get_ixti_raw(supbooks, ixti, opts), opts);
  }
  function stringify_formula(
    formula /* Array<any> */,
    range,
    cell,
    supbooks,
    opts
  ) {
    const biff = (opts && opts.biff) || 8;
    const _range = /* range != null ? range : */ {
      s: { c: 0, r: 0 },
      e: { c: 0, r: 0 }
    };
    const stack = [];
    let e1;
    let e2;
    let c;
    let ixti = 0;
    let nameidx = 0;
    let r;
    let sname = "";
    if (!formula[0] || !formula[0][0]) return "";
    let last_sp = -1;
    let sp = "";
    for (let ff = 0, fflen = formula[0].length; ff < fflen; ++ff) {
      let f = formula[0][ff];
      switch (f[0]) {
        case "PtgUminus" /* [MS-XLS] 2.5.198.93 */:
          stack.push(`-${stack.pop()}`);
          break;
        case "PtgUplus" /* [MS-XLS] 2.5.198.95 */:
          stack.push(`+${stack.pop()}`);
          break;
        case "PtgPercent" /* [MS-XLS] 2.5.198.81 */:
          stack.push(`${stack.pop()}%`);
          break;

        case "PtgAdd": /* [MS-XLS] 2.5.198.26 */
        case "PtgConcat": /* [MS-XLS] 2.5.198.43 */
        case "PtgDiv": /* [MS-XLS] 2.5.198.45 */
        case "PtgEq": /* [MS-XLS] 2.5.198.56 */
        case "PtgGe": /* [MS-XLS] 2.5.198.64 */
        case "PtgGt": /* [MS-XLS] 2.5.198.65 */
        case "PtgLe": /* [MS-XLS] 2.5.198.68 */
        case "PtgLt": /* [MS-XLS] 2.5.198.69 */
        case "PtgMul": /* [MS-XLS] 2.5.198.75 */
        case "PtgNe": /* [MS-XLS] 2.5.198.78 */
        case "PtgPower": /* [MS-XLS] 2.5.198.82 */
        case "PtgSub" /* [MS-XLS] 2.5.198.90 */:
          e1 = stack.pop();
          e2 = stack.pop();
          if (last_sp >= 0) {
            switch (formula[0][last_sp][1][0]) {
              case 0:
                // $FlowIgnore
                sp = fill(" ", formula[0][last_sp][1][1]);
                break;
              case 1:
                // $FlowIgnore
                sp = fill("\r", formula[0][last_sp][1][1]);
                break;
              default:
                sp = "";
                // $FlowIgnore
                if (opts.WTF)
                  throw new Error(
                    `Unexpected PtgAttrSpaceType ${formula[0][last_sp][1][0]}`
                  );
            }
            e2 += sp;
            last_sp = -1;
          }
          stack.push(e2 + PtgBinOp[f[0]] + e1);
          break;

        case "PtgIsect" /* [MS-XLS] 2.5.198.67 */:
          e1 = stack.pop();
          e2 = stack.pop();
          stack.push(`${e2} ${e1}`);
          break;
        case "PtgUnion" /* [MS-XLS] 2.5.198.94 */:
          e1 = stack.pop();
          e2 = stack.pop();
          stack.push(`${e2},${e1}`);
          break;
        case "PtgRange" /* [MS-XLS] 2.5.198.83 */:
          e1 = stack.pop();
          e2 = stack.pop();
          stack.push(`${e2}:${e1}`);
          break;

        case "PtgAttrChoose" /* [MS-XLS] 2.5.198.34 */:
          break;
        case "PtgAttrGoto" /* [MS-XLS] 2.5.198.35 */:
          break;
        case "PtgAttrIf" /* [MS-XLS] 2.5.198.36 */:
          break;
        case "PtgAttrIfError" /* [MS-XLSB] 2.5.97.28 */:
          break;

        case "PtgRef" /* [MS-XLS] 2.5.198.84 */:
          c = shift_cell_xls(f[1][1], _range, opts);
          stack.push(encode_cell_xls(c, biff));
          break;
        case "PtgRefN" /* [MS-XLS] 2.5.198.88 */:
          c = cell ? shift_cell_xls(f[1][1], cell, opts) : f[1][1];
          stack.push(encode_cell_xls(c, biff));
          break;
        case "PtgRef3d" /* [MS-XLS] 2.5.198.85 */:
          ixti = f[1][1];
          c = shift_cell_xls(f[1][2], _range, opts);
          sname = get_ixti(supbooks, ixti, opts);
          var w = sname; /* IE9 fails on defined names */ // eslint-disable-line no-unused-vars
          stack.push(`${sname}!${encode_cell_xls(c, biff)}`);
          break;

        case "PtgFunc": /* [MS-XLS] 2.5.198.62 */
        case "PtgFuncVar" /* [MS-XLS] 2.5.198.63 */:
          /* f[1] = [argc, func, type] */
          var argc = f[1][0];
          var func = f[1][1];
          if (!argc) argc = 0;
          argc &= 0x7f;
          var args = argc == 0 ? [] : stack.slice(-argc);
          stack.length -= argc;
          if (func === "User") func = args.shift();
          stack.push(`${func}(${args.join(",")})`);
          break;

        case "PtgBool" /* [MS-XLS] 2.5.198.42 */:
          stack.push(f[1] ? "TRUE" : "FALSE");
          break;
        case "PtgInt" /* [MS-XLS] 2.5.198.66 */:
          stack.push(f[1]);
          break;
        case "PtgNum" /* [MS-XLS] 2.5.198.79 TODO: precision? */:
          stack.push(String(f[1]));
          break;
        case "PtgStr" /* [MS-XLS] 2.5.198.89 */:
          // $FlowIgnore
          stack.push(`"${f[1].replace(/"/g, '""')}"`);
          break;
        case "PtgErr" /* [MS-XLS] 2.5.198.57 */:
          stack.push(f[1]);
          break;
        case "PtgAreaN" /* [MS-XLS] 2.5.198.31 TODO */:
          r = shift_range_xls(f[1][1], cell ? { s: cell } : _range, opts);
          stack.push(encode_range_xls(r, opts));
          break;
        case "PtgArea" /* [MS-XLS] 2.5.198.27 TODO: fixed points */:
          r = shift_range_xls(f[1][1], _range, opts);
          stack.push(encode_range_xls(r, opts));
          break;
        case "PtgArea3d" /* [MS-XLS] 2.5.198.28 TODO */:
          ixti = f[1][1];
          r = f[1][2];
          sname = get_ixti(supbooks, ixti, opts);
          stack.push(`${sname}!${encode_range_xls(r, opts)}`);
          break;
        case "PtgAttrSum" /* [MS-XLS] 2.5.198.41 */:
          stack.push(`SUM(${stack.pop()})`);
          break;

        case "PtgAttrBaxcel": /* [MS-XLS] 2.5.198.33 */
        case "PtgAttrSemi" /* [MS-XLS] 2.5.198.37 */:
          break;

        case "PtgName" /* [MS-XLS] 2.5.198.76 ; [MS-XLSB] 2.5.97.60 TODO: revisions */:
          /* f[1] = type, 0, nameindex */
          nameidx = f[1][2];
          var lbl =
            (supbooks.names || [])[nameidx - 1] || (supbooks[0] || [])[nameidx];
          var name = lbl ? lbl.Name : `SH33TJSNAME${String(nameidx)}`;
          if (name in XLSXFutureFunctions) name = XLSXFutureFunctions[name];
          stack.push(name);
          break;

        case "PtgNameX" /* [MS-XLS] 2.5.198.77 ; [MS-XLSB] 2.5.97.61 TODO: revisions */:
          /* f[1] = type, ixti, nameindex */
          var bookidx = f[1][1];
          nameidx = f[1][2];
          var externbook;
          /* TODO: Properly handle missing values -- this should be using get_ixti_raw primarily */
          if (opts.biff <= 5) {
            if (bookidx < 0) bookidx = -bookidx;
            if (supbooks[bookidx]) externbook = supbooks[bookidx][nameidx];
          } else {
            let o = "";
            if (((supbooks[bookidx] || [])[0] || [])[0] == 0x3a01) {
              /* empty */
            } else if (((supbooks[bookidx] || [])[0] || [])[0] == 0x0401) {
              if (
                supbooks[bookidx][nameidx] &&
                supbooks[bookidx][nameidx].itab > 0
              ) {
                o = `${
                  supbooks.SheetNames[supbooks[bookidx][nameidx].itab - 1]
                }!`;
              }
            } else o = `${supbooks.SheetNames[nameidx - 1]}!`;
            if (supbooks[bookidx] && supbooks[bookidx][nameidx])
              o += supbooks[bookidx][nameidx].Name;
            else if (supbooks[0] && supbooks[0][nameidx])
              o += supbooks[0][nameidx].Name;
            else {
              const ixtidata = get_ixti_raw(supbooks, bookidx, opts).split(
                ";;"
              );
              if (ixtidata[nameidx - 1]) o = ixtidata[nameidx - 1];
              // TODO: confirm this is correct
              else o += "SH33TJSERRX";
            }
            stack.push(o);
            break;
          }
          if (!externbook) externbook = { Name: "SH33TJSERRY" };
          stack.push(externbook.Name);
          break;

        case "PtgParen" /* [MS-XLS] 2.5.198.80 */:
          var lp = "(";
          var rp = ")";
          if (last_sp >= 0) {
            sp = "";
            switch (formula[0][last_sp][1][0]) {
              // $FlowIgnore
              case 2:
                lp = fill(" ", formula[0][last_sp][1][1]) + lp;
                break;
              // $FlowIgnore
              case 3:
                lp = fill("\r", formula[0][last_sp][1][1]) + lp;
                break;
              // $FlowIgnore
              case 4:
                rp = fill(" ", formula[0][last_sp][1][1]) + rp;
                break;
              // $FlowIgnore
              case 5:
                rp = fill("\r", formula[0][last_sp][1][1]) + rp;
                break;
              default:
                // $FlowIgnore
                if (opts.WTF)
                  throw new Error(
                    `Unexpected PtgAttrSpaceType ${formula[0][last_sp][1][0]}`
                  );
            }
            last_sp = -1;
          }
          stack.push(lp + stack.pop() + rp);
          break;

        case "PtgRefErr" /* [MS-XLS] 2.5.198.86 */:
          stack.push("#REF!");
          break;

        case "PtgRefErr3d" /* [MS-XLS] 2.5.198.87 */:
          stack.push("#REF!");
          break;

        case "PtgExp" /* [MS-XLS] 2.5.198.58 TODO */:
          c = { c: f[1][1], r: f[1][0] };
          var q = { c: cell.c, r: cell.r };
          if (supbooks.sharedf[encode_cell(c)]) {
            const parsedf = supbooks.sharedf[encode_cell(c)];
            stack.push(stringify_formula(parsedf, _range, q, supbooks, opts));
          } else {
            let fnd = false;
            for (e1 = 0; e1 != supbooks.arrayf.length; ++e1) {
              /* TODO: should be something like range_has */
              e2 = supbooks.arrayf[e1];
              if (c.c < e2[0].s.c || c.c > e2[0].e.c) continue;
              if (c.r < e2[0].s.r || c.r > e2[0].e.r) continue;
              stack.push(stringify_formula(e2[1], _range, q, supbooks, opts));
              fnd = true;
              break;
            }
            if (!fnd) stack.push(f[1]);
          }
          break;

        case "PtgArray" /* [MS-XLS] 2.5.198.32 TODO */:
          stack.push(`{${stringify_array(f[1])}}`);
          break;

        case "PtgMemArea" /* [MS-XLS] 2.5.198.70 TODO: confirm this is a non-display */:
          // stack.push("(" + f[2].map(encode_range).join(",") + ")");
          break;

        case "PtgAttrSpace": /* [MS-XLS] 2.5.198.38 */
        case "PtgAttrSpaceSemi" /* [MS-XLS] 2.5.198.39 */:
          last_sp = ff;
          break;

        case "PtgTbl" /* [MS-XLS] 2.5.198.92 TODO */:
          break;

        case "PtgMemErr" /* [MS-XLS] 2.5.198.71 */:
          break;

        case "PtgMissArg" /* [MS-XLS] 2.5.198.74 */:
          stack.push("");
          break;

        case "PtgAreaErr" /* [MS-XLS] 2.5.198.29 */:
          stack.push("#REF!");
          break;

        case "PtgAreaErr3d" /* [MS-XLS] 2.5.198.30 */:
          stack.push("#REF!");
          break;

        case "PtgList" /* [MS-XLSB] 2.5.97.52 */:
          // $FlowIgnore
          stack.push(`Table${f[1].idx}[#${f[1].rt}]`);
          break;

        case "PtgMemAreaN":
        case "PtgMemNoMemN":
        case "PtgAttrNoop":
        case "PtgSheet":
        case "PtgEndSheet":
          break;

        case "PtgMemFunc" /* [MS-XLS] 2.5.198.72 TODO */:
          break;
        case "PtgMemNoMem" /* [MS-XLS] 2.5.198.73 TODO */:
          break;

        case "PtgElfCol": /* [MS-XLS] 2.5.198.46 */
        case "PtgElfColS": /* [MS-XLS] 2.5.198.47 */
        case "PtgElfColSV": /* [MS-XLS] 2.5.198.48 */
        case "PtgElfColV": /* [MS-XLS] 2.5.198.49 */
        case "PtgElfLel": /* [MS-XLS] 2.5.198.50 */
        case "PtgElfRadical": /* [MS-XLS] 2.5.198.51 */
        case "PtgElfRadicalLel": /* [MS-XLS] 2.5.198.52 */
        case "PtgElfRadicalS": /* [MS-XLS] 2.5.198.53 */
        case "PtgElfRw": /* [MS-XLS] 2.5.198.54 */
        case "PtgElfRwV" /* [MS-XLS] 2.5.198.55 */:
          throw new Error("Unsupported ELFs");

        case "PtgSxName" /* [MS-XLS] 2.5.198.91 TODO -- find a test case */:
          throw new Error(`Unrecognized Formula Token: ${String(f)}`);
        default:
          throw new Error(`Unrecognized Formula Token: ${String(f)}`);
      }
      const PtgNonDisp = ["PtgAttrSpace", "PtgAttrSpaceSemi", "PtgAttrGoto"];
      if (opts.biff != 3)
        if (last_sp >= 0 && PtgNonDisp.indexOf(formula[0][ff][0]) == -1) {
          f = formula[0][last_sp];
          let _left = true;
          switch (f[1][0]) {
            /* note: some bad XLSB files omit the PtgParen */
            case 4:
              _left = false;
            /* falls through */
            case 0:
              // $FlowIgnore
              sp = fill(" ", f[1][1]);
              break;
            case 5:
              _left = false;
            /* falls through */
            case 1:
              // $FlowIgnore
              sp = fill("\r", f[1][1]);
              break;
            default:
              sp = "";
              // $FlowIgnore
              if (opts.WTF)
                throw new Error(`Unexpected PtgAttrSpaceType ${f[1][0]}`);
          }
          stack.push((_left ? sp : "") + stack.pop() + (_left ? "" : sp));
          last_sp = -1;
        }
    }
    if (stack.length > 1 && opts.WTF) throw new Error("bad formula stack");
    return stack[0];
  }

  /* [MS-XLS] 2.5.198.1 TODO */
  function parse_ArrayParsedFormula(blob, length, opts) {
    const target = blob.l + length;
    const len = opts.biff == 2 ? 1 : 2;
    let rgcb;
    const cce = blob.read_shift(len); // length of rgce
    if (cce == 0xffff) return [[], parsenoop(blob, length - 2)];
    const rgce = parse_Rgce(blob, cce, opts);
    if (length !== cce + len)
      rgcb = parse_RgbExtra(blob, length - cce - len, rgce, opts);
    blob.l = target;
    return [rgce, rgcb];
  }

  /* [MS-XLS] 2.5.198.3 TODO */
  function parse_XLSCellParsedFormula(blob, length, opts) {
    const target = blob.l + length;
    const len = opts.biff == 2 ? 1 : 2;
    let rgcb;
    const cce = blob.read_shift(len); // length of rgce
    if (cce == 0xffff) return [[], parsenoop(blob, length - 2)];
    const rgce = parse_Rgce(blob, cce, opts);
    if (length !== cce + len)
      rgcb = parse_RgbExtra(blob, length - cce - len, rgce, opts);
    blob.l = target;
    return [rgce, rgcb];
  }

  /* [MS-XLS] 2.5.198.21 */
  function parse_NameParsedFormula(blob, length, opts, cce) {
    const target = blob.l + length;
    const rgce = parse_Rgce(blob, cce, opts);
    let rgcb;
    if (target !== blob.l)
      rgcb = parse_RgbExtra(blob, target - blob.l, rgce, opts);
    return [rgce, rgcb];
  }

  /* [MS-XLS] 2.5.198.118 TODO */
  function parse_SharedParsedFormula(blob, length, opts) {
    const target = blob.l + length;
    let rgcb;
    const cce = blob.read_shift(2); // length of rgce
    const rgce = parse_Rgce(blob, cce, opts);
    if (cce == 0xffff) return [[], parsenoop(blob, length - 2)];
    if (length !== cce + 2)
      rgcb = parse_RgbExtra(blob, target - cce - 2, rgce, opts);
    return [rgce, rgcb];
  }

  /* [MS-XLS] 2.5.133 TODO: how to emit empty strings? */
  function parse_FormulaValue(blob) {
    let b;
    if (__readUInt16LE(blob, blob.l + 6) !== 0xffff)
      return [parse_Xnum(blob), "n"];
    switch (blob[blob.l]) {
      case 0x00:
        blob.l += 8;
        return ["String", "s"];
      case 0x01:
        b = blob[blob.l + 2] === 0x1;
        blob.l += 8;
        return [b, "b"];
      case 0x02:
        b = blob[blob.l + 2];
        blob.l += 8;
        return [b, "e"];
      case 0x03:
        blob.l += 8;
        return ["", "s"];
    }
    return [];
  }
  function write_FormulaValue(value) {
    if (value == null) {
      // Blank String Value
      const o = new_buf(8);
      o.write_shift(1, 0x03);
      o.write_shift(1, 0);
      o.write_shift(2, 0);
      o.write_shift(2, 0);
      o.write_shift(2, 0xffff);
      return o;
    } else if (typeof value === "number") return write_Xnum(value);
    return write_Xnum(0);
  }

  /* [MS-XLS] 2.4.127 TODO */
  function parse_Formula(blob, length, opts) {
    const end = blob.l + length;
    const cell = parse_XLSCell(blob, 6);
    if (opts.biff == 2) ++blob.l;
    const val = parse_FormulaValue(blob, 8);
    const flags = blob.read_shift(1);
    if (opts.biff != 2) {
      blob.read_shift(1);
      if (opts.biff >= 5) {
        /* var chn = */ blob.read_shift(4);
      }
    }
    const cbf = parse_XLSCellParsedFormula(blob, end - blob.l, opts);
    return {
      cell,
      val: val[0],
      formula: cbf,
      shared: (flags >> 3) & 1,
      tt: val[1]
    };
  }
  function write_Formula(cell, R, C, opts, os) {
    // Cell
    const o1 = write_XLSCell(R, C, os);

    // FormulaValue
    const o2 = write_FormulaValue(cell.v);

    // flags + cache
    const o3 = new_buf(6);
    const flags = 0x01 | 0x20;
    o3.write_shift(2, flags);
    o3.write_shift(4, 0);

    // CellParsedFormula
    const bf = new_buf(cell.bf.length);
    for (let i = 0; i < cell.bf.length; ++i) bf[i] = cell.bf[i];

    const out = bconcat([o1, o2, o3, bf]);
    return out;
  }

  /* XLSB Parsed Formula records have the same shape */
  function parse_XLSBParsedFormula(data, length, opts) {
    const cce = data.read_shift(4);
    const rgce = parse_Rgce(data, cce, opts);
    const cb = data.read_shift(4);
    const rgcb = cb > 0 ? parse_RgbExtra(data, cb, rgce, opts) : null;
    return [rgce, rgcb];
  }

  /* [MS-XLSB] 2.5.97.1 ArrayParsedFormula */
  const parse_XLSBArrayParsedFormula = parse_XLSBParsedFormula;
  /* [MS-XLSB] 2.5.97.4 CellParsedFormula */
  const parse_XLSBCellParsedFormula = parse_XLSBParsedFormula;
  /* [MS-XLSB] 2.5.97.8 DVParsedFormula */
  // var parse_XLSBDVParsedFormula = parse_XLSBParsedFormula;
  /* [MS-XLSB] 2.5.97.9 FRTParsedFormula */
  // var parse_XLSBFRTParsedFormula = parse_XLSBParsedFormula2;
  /* [MS-XLSB] 2.5.97.12 NameParsedFormula */
  const parse_XLSBNameParsedFormula = parse_XLSBParsedFormula;
  /* [MS-XLSB] 2.5.97.98 SharedParsedFormula */
  const parse_XLSBSharedParsedFormula = parse_XLSBParsedFormula;
  /* [MS-XLS] 2.5.198.4 */
  var Cetab = {
    0x0000: "BEEP",
    0x0001: "OPEN",
    0x0002: "OPEN.LINKS",
    0x0003: "CLOSE.ALL",
    0x0004: "SAVE",
    0x0005: "SAVE.AS",
    0x0006: "FILE.DELETE",
    0x0007: "PAGE.SETUP",
    0x0008: "PRINT",
    0x0009: "PRINTER.SETUP",
    0x000a: "QUIT",
    0x000b: "NEW.WINDOW",
    0x000c: "ARRANGE.ALL",
    0x000d: "WINDOW.SIZE",
    0x000e: "WINDOW.MOVE",
    0x000f: "FULL",
    0x0010: "CLOSE",
    0x0011: "RUN",
    0x0016: "SET.PRINT.AREA",
    0x0017: "SET.PRINT.TITLES",
    0x0018: "SET.PAGE.BREAK",
    0x0019: "REMOVE.PAGE.BREAK",
    0x001a: "FONT",
    0x001b: "DISPLAY",
    0x001c: "PROTECT.DOCUMENT",
    0x001d: "PRECISION",
    0x001e: "A1.R1C1",
    0x001f: "CALCULATE.NOW",
    0x0020: "CALCULATION",
    0x0022: "DATA.FIND",
    0x0023: "EXTRACT",
    0x0024: "DATA.DELETE",
    0x0025: "SET.DATABASE",
    0x0026: "SET.CRITERIA",
    0x0027: "SORT",
    0x0028: "DATA.SERIES",
    0x0029: "TABLE",
    0x002a: "FORMAT.NUMBER",
    0x002b: "ALIGNMENT",
    0x002c: "STYLE",
    0x002d: "BORDER",
    0x002e: "CELL.PROTECTION",
    0x002f: "COLUMN.WIDTH",
    0x0030: "UNDO",
    0x0031: "CUT",
    0x0032: "COPY",
    0x0033: "PASTE",
    0x0034: "CLEAR",
    0x0035: "PASTE.SPECIAL",
    0x0036: "EDIT.DELETE",
    0x0037: "INSERT",
    0x0038: "FILL.RIGHT",
    0x0039: "FILL.DOWN",
    0x003d: "DEFINE.NAME",
    0x003e: "CREATE.NAMES",
    0x003f: "FORMULA.GOTO",
    0x0040: "FORMULA.FIND",
    0x0041: "SELECT.LAST.CELL",
    0x0042: "SHOW.ACTIVE.CELL",
    0x0043: "GALLERY.AREA",
    0x0044: "GALLERY.BAR",
    0x0045: "GALLERY.COLUMN",
    0x0046: "GALLERY.LINE",
    0x0047: "GALLERY.PIE",
    0x0048: "GALLERY.SCATTER",
    0x0049: "COMBINATION",
    0x004a: "PREFERRED",
    0x004b: "ADD.OVERLAY",
    0x004c: "GRIDLINES",
    0x004d: "SET.PREFERRED",
    0x004e: "AXES",
    0x004f: "LEGEND",
    0x0050: "ATTACH.TEXT",
    0x0051: "ADD.ARROW",
    0x0052: "SELECT.CHART",
    0x0053: "SELECT.PLOT.AREA",
    0x0054: "PATTERNS",
    0x0055: "MAIN.CHART",
    0x0056: "OVERLAY",
    0x0057: "SCALE",
    0x0058: "FORMAT.LEGEND",
    0x0059: "FORMAT.TEXT",
    0x005a: "EDIT.REPEAT",
    0x005b: "PARSE",
    0x005c: "JUSTIFY",
    0x005d: "HIDE",
    0x005e: "UNHIDE",
    0x005f: "WORKSPACE",
    0x0060: "FORMULA",
    0x0061: "FORMULA.FILL",
    0x0062: "FORMULA.ARRAY",
    0x0063: "DATA.FIND.NEXT",
    0x0064: "DATA.FIND.PREV",
    0x0065: "FORMULA.FIND.NEXT",
    0x0066: "FORMULA.FIND.PREV",
    0x0067: "ACTIVATE",
    0x0068: "ACTIVATE.NEXT",
    0x0069: "ACTIVATE.PREV",
    0x006a: "UNLOCKED.NEXT",
    0x006b: "UNLOCKED.PREV",
    0x006c: "COPY.PICTURE",
    0x006d: "SELECT",
    0x006e: "DELETE.NAME",
    0x006f: "DELETE.FORMAT",
    0x0070: "VLINE",
    0x0071: "HLINE",
    0x0072: "VPAGE",
    0x0073: "HPAGE",
    0x0074: "VSCROLL",
    0x0075: "HSCROLL",
    0x0076: "ALERT",
    0x0077: "NEW",
    0x0078: "CANCEL.COPY",
    0x0079: "SHOW.CLIPBOARD",
    0x007a: "MESSAGE",
    0x007c: "PASTE.LINK",
    0x007d: "APP.ACTIVATE",
    0x007e: "DELETE.ARROW",
    0x007f: "ROW.HEIGHT",
    0x0080: "FORMAT.MOVE",
    0x0081: "FORMAT.SIZE",
    0x0082: "FORMULA.REPLACE",
    0x0083: "SEND.KEYS",
    0x0084: "SELECT.SPECIAL",
    0x0085: "APPLY.NAMES",
    0x0086: "REPLACE.FONT",
    0x0087: "FREEZE.PANES",
    0x0088: "SHOW.INFO",
    0x0089: "SPLIT",
    0x008a: "ON.WINDOW",
    0x008b: "ON.DATA",
    0x008c: "DISABLE.INPUT",
    0x008e: "OUTLINE",
    0x008f: "LIST.NAMES",
    0x0090: "FILE.CLOSE",
    0x0091: "SAVE.WORKBOOK",
    0x0092: "DATA.FORM",
    0x0093: "COPY.CHART",
    0x0094: "ON.TIME",
    0x0095: "WAIT",
    0x0096: "FORMAT.FONT",
    0x0097: "FILL.UP",
    0x0098: "FILL.LEFT",
    0x0099: "DELETE.OVERLAY",
    0x009b: "SHORT.MENUS",
    0x009f: "SET.UPDATE.STATUS",
    0x00a1: "COLOR.PALETTE",
    0x00a2: "DELETE.STYLE",
    0x00a3: "WINDOW.RESTORE",
    0x00a4: "WINDOW.MAXIMIZE",
    0x00a6: "CHANGE.LINK",
    0x00a7: "CALCULATE.DOCUMENT",
    0x00a8: "ON.KEY",
    0x00a9: "APP.RESTORE",
    0x00aa: "APP.MOVE",
    0x00ab: "APP.SIZE",
    0x00ac: "APP.MINIMIZE",
    0x00ad: "APP.MAXIMIZE",
    0x00ae: "BRING.TO.FRONT",
    0x00af: "SEND.TO.BACK",
    0x00b9: "MAIN.CHART.TYPE",
    0x00ba: "OVERLAY.CHART.TYPE",
    0x00bb: "SELECT.END",
    0x00bc: "OPEN.MAIL",
    0x00bd: "SEND.MAIL",
    0x00be: "STANDARD.FONT",
    0x00bf: "CONSOLIDATE",
    0x00c0: "SORT.SPECIAL",
    0x00c1: "GALLERY.3D.AREA",
    0x00c2: "GALLERY.3D.COLUMN",
    0x00c3: "GALLERY.3D.LINE",
    0x00c4: "GALLERY.3D.PIE",
    0x00c5: "VIEW.3D",
    0x00c6: "GOAL.SEEK",
    0x00c7: "WORKGROUP",
    0x00c8: "FILL.GROUP",
    0x00c9: "UPDATE.LINK",
    0x00ca: "PROMOTE",
    0x00cb: "DEMOTE",
    0x00cc: "SHOW.DETAIL",
    0x00ce: "UNGROUP",
    0x00cf: "OBJECT.PROPERTIES",
    0x00d0: "SAVE.NEW.OBJECT",
    0x00d1: "SHARE",
    0x00d2: "SHARE.NAME",
    0x00d3: "DUPLICATE",
    0x00d4: "APPLY.STYLE",
    0x00d5: "ASSIGN.TO.OBJECT",
    0x00d6: "OBJECT.PROTECTION",
    0x00d7: "HIDE.OBJECT",
    0x00d8: "SET.EXTRACT",
    0x00d9: "CREATE.PUBLISHER",
    0x00da: "SUBSCRIBE.TO",
    0x00db: "ATTRIBUTES",
    0x00dc: "SHOW.TOOLBAR",
    0x00de: "PRINT.PREVIEW",
    0x00df: "EDIT.COLOR",
    0x00e0: "SHOW.LEVELS",
    0x00e1: "FORMAT.MAIN",
    0x00e2: "FORMAT.OVERLAY",
    0x00e3: "ON.RECALC",
    0x00e4: "EDIT.SERIES",
    0x00e5: "DEFINE.STYLE",
    0x00f0: "LINE.PRINT",
    0x00f3: "ENTER.DATA",
    0x00f9: "GALLERY.RADAR",
    0x00fa: "MERGE.STYLES",
    0x00fb: "EDITION.OPTIONS",
    0x00fc: "PASTE.PICTURE",
    0x00fd: "PASTE.PICTURE.LINK",
    0x00fe: "SPELLING",
    0x0100: "ZOOM",
    0x0103: "INSERT.OBJECT",
    0x0104: "WINDOW.MINIMIZE",
    0x0109: "SOUND.NOTE",
    0x010a: "SOUND.PLAY",
    0x010b: "FORMAT.SHAPE",
    0x010c: "EXTEND.POLYGON",
    0x010d: "FORMAT.AUTO",
    0x0110: "GALLERY.3D.BAR",
    0x0111: "GALLERY.3D.SURFACE",
    0x0112: "FILL.AUTO",
    0x0114: "CUSTOMIZE.TOOLBAR",
    0x0115: "ADD.TOOL",
    0x0116: "EDIT.OBJECT",
    0x0117: "ON.DOUBLECLICK",
    0x0118: "ON.ENTRY",
    0x0119: "WORKBOOK.ADD",
    0x011a: "WORKBOOK.MOVE",
    0x011b: "WORKBOOK.COPY",
    0x011c: "WORKBOOK.OPTIONS",
    0x011d: "SAVE.WORKSPACE",
    0x0120: "CHART.WIZARD",
    0x0121: "DELETE.TOOL",
    0x0122: "MOVE.TOOL",
    0x0123: "WORKBOOK.SELECT",
    0x0124: "WORKBOOK.ACTIVATE",
    0x0125: "ASSIGN.TO.TOOL",
    0x0127: "COPY.TOOL",
    0x0128: "RESET.TOOL",
    0x0129: "CONSTRAIN.NUMERIC",
    0x012a: "PASTE.TOOL",
    0x012e: "WORKBOOK.NEW",
    0x0131: "SCENARIO.CELLS",
    0x0132: "SCENARIO.DELETE",
    0x0133: "SCENARIO.ADD",
    0x0134: "SCENARIO.EDIT",
    0x0135: "SCENARIO.SHOW",
    0x0136: "SCENARIO.SHOW.NEXT",
    0x0137: "SCENARIO.SUMMARY",
    0x0138: "PIVOT.TABLE.WIZARD",
    0x0139: "PIVOT.FIELD.PROPERTIES",
    0x013a: "PIVOT.FIELD",
    0x013b: "PIVOT.ITEM",
    0x013c: "PIVOT.ADD.FIELDS",
    0x013e: "OPTIONS.CALCULATION",
    0x013f: "OPTIONS.EDIT",
    0x0140: "OPTIONS.VIEW",
    0x0141: "ADDIN.MANAGER",
    0x0142: "MENU.EDITOR",
    0x0143: "ATTACH.TOOLBARS",
    0x0144: "VBAActivate",
    0x0145: "OPTIONS.CHART",
    0x0148: "VBA.INSERT.FILE",
    0x014a: "VBA.PROCEDURE.DEFINITION",
    0x0150: "ROUTING.SLIP",
    0x0152: "ROUTE.DOCUMENT",
    0x0153: "MAIL.LOGON",
    0x0156: "INSERT.PICTURE",
    0x0157: "EDIT.TOOL",
    0x0158: "GALLERY.DOUGHNUT",
    0x015e: "CHART.TREND",
    0x0160: "PIVOT.ITEM.PROPERTIES",
    0x0162: "WORKBOOK.INSERT",
    0x0163: "OPTIONS.TRANSITION",
    0x0164: "OPTIONS.GENERAL",
    0x0172: "FILTER.ADVANCED",
    0x0175: "MAIL.ADD.MAILER",
    0x0176: "MAIL.DELETE.MAILER",
    0x0177: "MAIL.REPLY",
    0x0178: "MAIL.REPLY.ALL",
    0x0179: "MAIL.FORWARD",
    0x017a: "MAIL.NEXT.LETTER",
    0x017b: "DATA.LABEL",
    0x017c: "INSERT.TITLE",
    0x017d: "FONT.PROPERTIES",
    0x017e: "MACRO.OPTIONS",
    0x017f: "WORKBOOK.HIDE",
    0x0180: "WORKBOOK.UNHIDE",
    0x0181: "WORKBOOK.DELETE",
    0x0182: "WORKBOOK.NAME",
    0x0184: "GALLERY.CUSTOM",
    0x0186: "ADD.CHART.AUTOFORMAT",
    0x0187: "DELETE.CHART.AUTOFORMAT",
    0x0188: "CHART.ADD.DATA",
    0x0189: "AUTO.OUTLINE",
    0x018a: "TAB.ORDER",
    0x018b: "SHOW.DIALOG",
    0x018c: "SELECT.ALL",
    0x018d: "UNGROUP.SHEETS",
    0x018e: "SUBTOTAL.CREATE",
    0x018f: "SUBTOTAL.REMOVE",
    0x0190: "RENAME.OBJECT",
    0x019c: "WORKBOOK.SCROLL",
    0x019d: "WORKBOOK.NEXT",
    0x019e: "WORKBOOK.PREV",
    0x019f: "WORKBOOK.TAB.SPLIT",
    0x01a0: "FULL.SCREEN",
    0x01a1: "WORKBOOK.PROTECT",
    0x01a4: "SCROLLBAR.PROPERTIES",
    0x01a5: "PIVOT.SHOW.PAGES",
    0x01a6: "TEXT.TO.COLUMNS",
    0x01a7: "FORMAT.CHARTTYPE",
    0x01a8: "LINK.FORMAT",
    0x01a9: "TRACER.DISPLAY",
    0x01ae: "TRACER.NAVIGATE",
    0x01af: "TRACER.CLEAR",
    0x01b0: "TRACER.ERROR",
    0x01b1: "PIVOT.FIELD.GROUP",
    0x01b2: "PIVOT.FIELD.UNGROUP",
    0x01b3: "CHECKBOX.PROPERTIES",
    0x01b4: "LABEL.PROPERTIES",
    0x01b5: "LISTBOX.PROPERTIES",
    0x01b6: "EDITBOX.PROPERTIES",
    0x01b7: "PIVOT.REFRESH",
    0x01b8: "LINK.COMBO",
    0x01b9: "OPEN.TEXT",
    0x01ba: "HIDE.DIALOG",
    0x01bb: "SET.DIALOG.FOCUS",
    0x01bc: "ENABLE.OBJECT",
    0x01bd: "PUSHBUTTON.PROPERTIES",
    0x01be: "SET.DIALOG.DEFAULT",
    0x01bf: "FILTER",
    0x01c0: "FILTER.SHOW.ALL",
    0x01c1: "CLEAR.OUTLINE",
    0x01c2: "FUNCTION.WIZARD",
    0x01c3: "ADD.LIST.ITEM",
    0x01c4: "SET.LIST.ITEM",
    0x01c5: "REMOVE.LIST.ITEM",
    0x01c6: "SELECT.LIST.ITEM",
    0x01c7: "SET.CONTROL.VALUE",
    0x01c8: "SAVE.COPY.AS",
    0x01ca: "OPTIONS.LISTS.ADD",
    0x01cb: "OPTIONS.LISTS.DELETE",
    0x01cc: "SERIES.AXES",
    0x01cd: "SERIES.X",
    0x01ce: "SERIES.Y",
    0x01cf: "ERRORBAR.X",
    0x01d0: "ERRORBAR.Y",
    0x01d1: "FORMAT.CHART",
    0x01d2: "SERIES.ORDER",
    0x01d3: "MAIL.LOGOFF",
    0x01d4: "CLEAR.ROUTING.SLIP",
    0x01d5: "APP.ACTIVATE.MICROSOFT",
    0x01d6: "MAIL.EDIT.MAILER",
    0x01d7: "ON.SHEET",
    0x01d8: "STANDARD.WIDTH",
    0x01d9: "SCENARIO.MERGE",
    0x01da: "SUMMARY.INFO",
    0x01db: "FIND.FILE",
    0x01dc: "ACTIVE.CELL.FONT",
    0x01dd: "ENABLE.TIPWIZARD",
    0x01de: "VBA.MAKE.ADDIN",
    0x01e0: "INSERTDATATABLE",
    0x01e1: "WORKGROUP.OPTIONS",
    0x01e2: "MAIL.SEND.MAILER",
    0x01e5: "AUTOCORRECT",
    0x01e9: "POST.DOCUMENT",
    0x01eb: "PICKLIST",
    0x01ed: "VIEW.SHOW",
    0x01ee: "VIEW.DEFINE",
    0x01ef: "VIEW.DELETE",
    0x01fd: "SHEET.BACKGROUND",
    0x01fe: "INSERT.MAP.OBJECT",
    0x01ff: "OPTIONS.MENONO",
    0x0205: "MSOCHECKS",
    0x0206: "NORMAL",
    0x0207: "LAYOUT",
    0x0208: "RM.PRINT.AREA",
    0x0209: "CLEAR.PRINT.AREA",
    0x020a: "ADD.PRINT.AREA",
    0x020b: "MOVE.BRK",
    0x0221: "HIDECURR.NOTE",
    0x0222: "HIDEALL.NOTES",
    0x0223: "DELETE.NOTE",
    0x0224: "TRAVERSE.NOTES",
    0x0225: "ACTIVATE.NOTES",
    0x026c: "PROTECT.REVISIONS",
    0x026d: "UNPROTECT.REVISIONS",
    0x0287: "OPTIONS.ME",
    0x028d: "WEB.PUBLISH",
    0x029b: "NEWWEBQUERY",
    0x02a1: "PIVOT.TABLE.CHART",
    0x02f1: "OPTIONS.SAVE",
    0x02f3: "OPTIONS.SPELL",
    0x0328: "HIDEALL.INKANNOTS"
  };

  /* [MS-XLS] 2.5.198.17 */
  /* [MS-XLSB] 2.5.97.10 */
  var Ftab = {
    0x0000: "COUNT",
    0x0001: "IF",
    0x0002: "ISNA",
    0x0003: "ISERROR",
    0x0004: "SUM",
    0x0005: "AVERAGE",
    0x0006: "MIN",
    0x0007: "MAX",
    0x0008: "ROW",
    0x0009: "COLUMN",
    0x000a: "NA",
    0x000b: "NPV",
    0x000c: "STDEV",
    0x000d: "DOLLAR",
    0x000e: "FIXED",
    0x000f: "SIN",
    0x0010: "COS",
    0x0011: "TAN",
    0x0012: "ATAN",
    0x0013: "PI",
    0x0014: "SQRT",
    0x0015: "EXP",
    0x0016: "LN",
    0x0017: "LOG10",
    0x0018: "ABS",
    0x0019: "INT",
    0x001a: "SIGN",
    0x001b: "ROUND",
    0x001c: "LOOKUP",
    0x001d: "INDEX",
    0x001e: "REPT",
    0x001f: "MID",
    0x0020: "LEN",
    0x0021: "VALUE",
    0x0022: "TRUE",
    0x0023: "FALSE",
    0x0024: "AND",
    0x0025: "OR",
    0x0026: "NOT",
    0x0027: "MOD",
    0x0028: "DCOUNT",
    0x0029: "DSUM",
    0x002a: "DAVERAGE",
    0x002b: "DMIN",
    0x002c: "DMAX",
    0x002d: "DSTDEV",
    0x002e: "VAR",
    0x002f: "DVAR",
    0x0030: "TEXT",
    0x0031: "LINEST",
    0x0032: "TREND",
    0x0033: "LOGEST",
    0x0034: "GROWTH",
    0x0035: "GOTO",
    0x0036: "HALT",
    0x0037: "RETURN",
    0x0038: "PV",
    0x0039: "FV",
    0x003a: "NPER",
    0x003b: "PMT",
    0x003c: "RATE",
    0x003d: "MIRR",
    0x003e: "IRR",
    0x003f: "RAND",
    0x0040: "MATCH",
    0x0041: "DATE",
    0x0042: "TIME",
    0x0043: "DAY",
    0x0044: "MONTH",
    0x0045: "YEAR",
    0x0046: "WEEKDAY",
    0x0047: "HOUR",
    0x0048: "MINUTE",
    0x0049: "SECOND",
    0x004a: "NOW",
    0x004b: "AREAS",
    0x004c: "ROWS",
    0x004d: "COLUMNS",
    0x004e: "OFFSET",
    0x004f: "ABSREF",
    0x0050: "RELREF",
    0x0051: "ARGUMENT",
    0x0052: "SEARCH",
    0x0053: "TRANSPOSE",
    0x0054: "ERROR",
    0x0055: "STEP",
    0x0056: "TYPE",
    0x0057: "ECHO",
    0x0058: "SET.NAME",
    0x0059: "CALLER",
    0x005a: "DEREF",
    0x005b: "WINDOWS",
    0x005c: "SERIES",
    0x005d: "DOCUMENTS",
    0x005e: "ACTIVE.CELL",
    0x005f: "SELECTION",
    0x0060: "RESULT",
    0x0061: "ATAN2",
    0x0062: "ASIN",
    0x0063: "ACOS",
    0x0064: "CHOOSE",
    0x0065: "HLOOKUP",
    0x0066: "VLOOKUP",
    0x0067: "LINKS",
    0x0068: "INPUT",
    0x0069: "ISREF",
    0x006a: "GET.FORMULA",
    0x006b: "GET.NAME",
    0x006c: "SET.VALUE",
    0x006d: "LOG",
    0x006e: "EXEC",
    0x006f: "CHAR",
    0x0070: "LOWER",
    0x0071: "UPPER",
    0x0072: "PROPER",
    0x0073: "LEFT",
    0x0074: "RIGHT",
    0x0075: "EXACT",
    0x0076: "TRIM",
    0x0077: "REPLACE",
    0x0078: "SUBSTITUTE",
    0x0079: "CODE",
    0x007a: "NAMES",
    0x007b: "DIRECTORY",
    0x007c: "FIND",
    0x007d: "CELL",
    0x007e: "ISERR",
    0x007f: "ISTEXT",
    0x0080: "ISNUMBER",
    0x0081: "ISBLANK",
    0x0082: "T",
    0x0083: "N",
    0x0084: "FOPEN",
    0x0085: "FCLOSE",
    0x0086: "FSIZE",
    0x0087: "FREADLN",
    0x0088: "FREAD",
    0x0089: "FWRITELN",
    0x008a: "FWRITE",
    0x008b: "FPOS",
    0x008c: "DATEVALUE",
    0x008d: "TIMEVALUE",
    0x008e: "SLN",
    0x008f: "SYD",
    0x0090: "DDB",
    0x0091: "GET.DEF",
    0x0092: "REFTEXT",
    0x0093: "TEXTREF",
    0x0094: "INDIRECT",
    0x0095: "REGISTER",
    0x0096: "CALL",
    0x0097: "ADD.BAR",
    0x0098: "ADD.MENU",
    0x0099: "ADD.COMMAND",
    0x009a: "ENABLE.COMMAND",
    0x009b: "CHECK.COMMAND",
    0x009c: "RENAME.COMMAND",
    0x009d: "SHOW.BAR",
    0x009e: "DELETE.MENU",
    0x009f: "DELETE.COMMAND",
    0x00a0: "GET.CHART.ITEM",
    0x00a1: "DIALOG.BOX",
    0x00a2: "CLEAN",
    0x00a3: "MDETERM",
    0x00a4: "MINVERSE",
    0x00a5: "MMULT",
    0x00a6: "FILES",
    0x00a7: "IPMT",
    0x00a8: "PPMT",
    0x00a9: "COUNTA",
    0x00aa: "CANCEL.KEY",
    0x00ab: "FOR",
    0x00ac: "WHILE",
    0x00ad: "BREAK",
    0x00ae: "NEXT",
    0x00af: "INITIATE",
    0x00b0: "REQUEST",
    0x00b1: "POKE",
    0x00b2: "EXECUTE",
    0x00b3: "TERMINATE",
    0x00b4: "RESTART",
    0x00b5: "HELP",
    0x00b6: "GET.BAR",
    0x00b7: "PRODUCT",
    0x00b8: "FACT",
    0x00b9: "GET.CELL",
    0x00ba: "GET.WORKSPACE",
    0x00bb: "GET.WINDOW",
    0x00bc: "GET.DOCUMENT",
    0x00bd: "DPRODUCT",
    0x00be: "ISNONTEXT",
    0x00bf: "GET.NOTE",
    0x00c0: "NOTE",
    0x00c1: "STDEVP",
    0x00c2: "VARP",
    0x00c3: "DSTDEVP",
    0x00c4: "DVARP",
    0x00c5: "TRUNC",
    0x00c6: "ISLOGICAL",
    0x00c7: "DCOUNTA",
    0x00c8: "DELETE.BAR",
    0x00c9: "UNREGISTER",
    0x00cc: "USDOLLAR",
    0x00cd: "FINDB",
    0x00ce: "SEARCHB",
    0x00cf: "REPLACEB",
    0x00d0: "LEFTB",
    0x00d1: "RIGHTB",
    0x00d2: "MIDB",
    0x00d3: "LENB",
    0x00d4: "ROUNDUP",
    0x00d5: "ROUNDDOWN",
    0x00d6: "ASC",
    0x00d7: "DBCS",
    0x00d8: "RANK",
    0x00db: "ADDRESS",
    0x00dc: "DAYS360",
    0x00dd: "TODAY",
    0x00de: "VDB",
    0x00df: "ELSE",
    0x00e0: "ELSE.IF",
    0x00e1: "END.IF",
    0x00e2: "FOR.CELL",
    0x00e3: "MEDIAN",
    0x00e4: "SUMPRODUCT",
    0x00e5: "SINH",
    0x00e6: "COSH",
    0x00e7: "TANH",
    0x00e8: "ASINH",
    0x00e9: "ACOSH",
    0x00ea: "ATANH",
    0x00eb: "DGET",
    0x00ec: "CREATE.OBJECT",
    0x00ed: "VOLATILE",
    0x00ee: "LAST.ERROR",
    0x00ef: "CUSTOM.UNDO",
    0x00f0: "CUSTOM.REPEAT",
    0x00f1: "FORMULA.CONVERT",
    0x00f2: "GET.LINK.INFO",
    0x00f3: "TEXT.BOX",
    0x00f4: "INFO",
    0x00f5: "GROUP",
    0x00f6: "GET.OBJECT",
    0x00f7: "DB",
    0x00f8: "PAUSE",
    0x00fb: "RESUME",
    0x00fc: "FREQUENCY",
    0x00fd: "ADD.TOOLBAR",
    0x00fe: "DELETE.TOOLBAR",
    0x00ff: "User",
    0x0100: "RESET.TOOLBAR",
    0x0101: "EVALUATE",
    0x0102: "GET.TOOLBAR",
    0x0103: "GET.TOOL",
    0x0104: "SPELLING.CHECK",
    0x0105: "ERROR.TYPE",
    0x0106: "APP.TITLE",
    0x0107: "WINDOW.TITLE",
    0x0108: "SAVE.TOOLBAR",
    0x0109: "ENABLE.TOOL",
    0x010a: "PRESS.TOOL",
    0x010b: "REGISTER.ID",
    0x010c: "GET.WORKBOOK",
    0x010d: "AVEDEV",
    0x010e: "BETADIST",
    0x010f: "GAMMALN",
    0x0110: "BETAINV",
    0x0111: "BINOMDIST",
    0x0112: "CHIDIST",
    0x0113: "CHIINV",
    0x0114: "COMBIN",
    0x0115: "CONFIDENCE",
    0x0116: "CRITBINOM",
    0x0117: "EVEN",
    0x0118: "EXPONDIST",
    0x0119: "FDIST",
    0x011a: "FINV",
    0x011b: "FISHER",
    0x011c: "FISHERINV",
    0x011d: "FLOOR",
    0x011e: "GAMMADIST",
    0x011f: "GAMMAINV",
    0x0120: "CEILING",
    0x0121: "HYPGEOMDIST",
    0x0122: "LOGNORMDIST",
    0x0123: "LOGINV",
    0x0124: "NEGBINOMDIST",
    0x0125: "NORMDIST",
    0x0126: "NORMSDIST",
    0x0127: "NORMINV",
    0x0128: "NORMSINV",
    0x0129: "STANDARDIZE",
    0x012a: "ODD",
    0x012b: "PERMUT",
    0x012c: "POISSON",
    0x012d: "TDIST",
    0x012e: "WEIBULL",
    0x012f: "SUMXMY2",
    0x0130: "SUMX2MY2",
    0x0131: "SUMX2PY2",
    0x0132: "CHITEST",
    0x0133: "CORREL",
    0x0134: "COVAR",
    0x0135: "FORECAST",
    0x0136: "FTEST",
    0x0137: "INTERCEPT",
    0x0138: "PEARSON",
    0x0139: "RSQ",
    0x013a: "STEYX",
    0x013b: "SLOPE",
    0x013c: "TTEST",
    0x013d: "PROB",
    0x013e: "DEVSQ",
    0x013f: "GEOMEAN",
    0x0140: "HARMEAN",
    0x0141: "SUMSQ",
    0x0142: "KURT",
    0x0143: "SKEW",
    0x0144: "ZTEST",
    0x0145: "LARGE",
    0x0146: "SMALL",
    0x0147: "QUARTILE",
    0x0148: "PERCENTILE",
    0x0149: "PERCENTRANK",
    0x014a: "MODE",
    0x014b: "TRIMMEAN",
    0x014c: "TINV",
    0x014e: "MOVIE.COMMAND",
    0x014f: "GET.MOVIE",
    0x0150: "CONCATENATE",
    0x0151: "POWER",
    0x0152: "PIVOT.ADD.DATA",
    0x0153: "GET.PIVOT.TABLE",
    0x0154: "GET.PIVOT.FIELD",
    0x0155: "GET.PIVOT.ITEM",
    0x0156: "RADIANS",
    0x0157: "DEGREES",
    0x0158: "SUBTOTAL",
    0x0159: "SUMIF",
    0x015a: "COUNTIF",
    0x015b: "COUNTBLANK",
    0x015c: "SCENARIO.GET",
    0x015d: "OPTIONS.LISTS.GET",
    0x015e: "ISPMT",
    0x015f: "DATEDIF",
    0x0160: "DATESTRING",
    0x0161: "NUMBERSTRING",
    0x0162: "ROMAN",
    0x0163: "OPEN.DIALOG",
    0x0164: "SAVE.DIALOG",
    0x0165: "VIEW.GET",
    0x0166: "GETPIVOTDATA",
    0x0167: "HYPERLINK",
    0x0168: "PHONETIC",
    0x0169: "AVERAGEA",
    0x016a: "MAXA",
    0x016b: "MINA",
    0x016c: "STDEVPA",
    0x016d: "VARPA",
    0x016e: "STDEVA",
    0x016f: "VARA",
    0x0170: "BAHTTEXT",
    0x0171: "THAIDAYOFWEEK",
    0x0172: "THAIDIGIT",
    0x0173: "THAIMONTHOFYEAR",
    0x0174: "THAINUMSOUND",
    0x0175: "THAINUMSTRING",
    0x0176: "THAISTRINGLENGTH",
    0x0177: "ISTHAIDIGIT",
    0x0178: "ROUNDBAHTDOWN",
    0x0179: "ROUNDBAHTUP",
    0x017a: "THAIYEAR",
    0x017b: "RTD",

    0x017c: "CUBEVALUE",
    0x017d: "CUBEMEMBER",
    0x017e: "CUBEMEMBERPROPERTY",
    0x017f: "CUBERANKEDMEMBER",
    0x0180: "HEX2BIN",
    0x0181: "HEX2DEC",
    0x0182: "HEX2OCT",
    0x0183: "DEC2BIN",
    0x0184: "DEC2HEX",
    0x0185: "DEC2OCT",
    0x0186: "OCT2BIN",
    0x0187: "OCT2HEX",
    0x0188: "OCT2DEC",
    0x0189: "BIN2DEC",
    0x018a: "BIN2OCT",
    0x018b: "BIN2HEX",
    0x018c: "IMSUB",
    0x018d: "IMDIV",
    0x018e: "IMPOWER",
    0x018f: "IMABS",
    0x0190: "IMSQRT",
    0x0191: "IMLN",
    0x0192: "IMLOG2",
    0x0193: "IMLOG10",
    0x0194: "IMSIN",
    0x0195: "IMCOS",
    0x0196: "IMEXP",
    0x0197: "IMARGUMENT",
    0x0198: "IMCONJUGATE",
    0x0199: "IMAGINARY",
    0x019a: "IMREAL",
    0x019b: "COMPLEX",
    0x019c: "IMSUM",
    0x019d: "IMPRODUCT",
    0x019e: "SERIESSUM",
    0x019f: "FACTDOUBLE",
    0x01a0: "SQRTPI",
    0x01a1: "QUOTIENT",
    0x01a2: "DELTA",
    0x01a3: "GESTEP",
    0x01a4: "ISEVEN",
    0x01a5: "ISODD",
    0x01a6: "MROUND",
    0x01a7: "ERF",
    0x01a8: "ERFC",
    0x01a9: "BESSELJ",
    0x01aa: "BESSELK",
    0x01ab: "BESSELY",
    0x01ac: "BESSELI",
    0x01ad: "XIRR",
    0x01ae: "XNPV",
    0x01af: "PRICEMAT",
    0x01b0: "YIELDMAT",
    0x01b1: "INTRATE",
    0x01b2: "RECEIVED",
    0x01b3: "DISC",
    0x01b4: "PRICEDISC",
    0x01b5: "YIELDDISC",
    0x01b6: "TBILLEQ",
    0x01b7: "TBILLPRICE",
    0x01b8: "TBILLYIELD",
    0x01b9: "PRICE",
    0x01ba: "YIELD",
    0x01bb: "DOLLARDE",
    0x01bc: "DOLLARFR",
    0x01bd: "NOMINAL",
    0x01be: "EFFECT",
    0x01bf: "CUMPRINC",
    0x01c0: "CUMIPMT",
    0x01c1: "EDATE",
    0x01c2: "EOMONTH",
    0x01c3: "YEARFRAC",
    0x01c4: "COUPDAYBS",
    0x01c5: "COUPDAYS",
    0x01c6: "COUPDAYSNC",
    0x01c7: "COUPNCD",
    0x01c8: "COUPNUM",
    0x01c9: "COUPPCD",
    0x01ca: "DURATION",
    0x01cb: "MDURATION",
    0x01cc: "ODDLPRICE",
    0x01cd: "ODDLYIELD",
    0x01ce: "ODDFPRICE",
    0x01cf: "ODDFYIELD",
    0x01d0: "RANDBETWEEN",
    0x01d1: "WEEKNUM",
    0x01d2: "AMORDEGRC",
    0x01d3: "AMORLINC",
    0x01d4: "CONVERT",
    0x02d4: "SHEETJS",
    0x01d5: "ACCRINT",
    0x01d6: "ACCRINTM",
    0x01d7: "WORKDAY",
    0x01d8: "NETWORKDAYS",
    0x01d9: "GCD",
    0x01da: "MULTINOMIAL",
    0x01db: "LCM",
    0x01dc: "FVSCHEDULE",
    0x01dd: "CUBEKPIMEMBER",
    0x01de: "CUBESET",
    0x01df: "CUBESETCOUNT",
    0x01e0: "IFERROR",
    0x01e1: "COUNTIFS",
    0x01e2: "SUMIFS",
    0x01e3: "AVERAGEIF",
    0x01e4: "AVERAGEIFS"
  };
  var FtabArgc = {
    0x0002: 1 /* ISNA */,
    0x0003: 1 /* ISERROR */,
    0x000a: 0 /* NA */,
    0x000f: 1 /* SIN */,
    0x0010: 1 /* COS */,
    0x0011: 1 /* TAN */,
    0x0012: 1 /* ATAN */,
    0x0013: 0 /* PI */,
    0x0014: 1 /* SQRT */,
    0x0015: 1 /* EXP */,
    0x0016: 1 /* LN */,
    0x0017: 1 /* LOG10 */,
    0x0018: 1 /* ABS */,
    0x0019: 1 /* INT */,
    0x001a: 1 /* SIGN */,
    0x001b: 2 /* ROUND */,
    0x001e: 2 /* REPT */,
    0x001f: 3 /* MID */,
    0x0020: 1 /* LEN */,
    0x0021: 1 /* VALUE */,
    0x0022: 0 /* TRUE */,
    0x0023: 0 /* FALSE */,
    0x0026: 1 /* NOT */,
    0x0027: 2 /* MOD */,
    0x0028: 3 /* DCOUNT */,
    0x0029: 3 /* DSUM */,
    0x002a: 3 /* DAVERAGE */,
    0x002b: 3 /* DMIN */,
    0x002c: 3 /* DMAX */,
    0x002d: 3 /* DSTDEV */,
    0x002f: 3 /* DVAR */,
    0x0030: 2 /* TEXT */,
    0x0035: 1 /* GOTO */,
    0x003d: 3 /* MIRR */,
    0x003f: 0 /* RAND */,
    0x0041: 3 /* DATE */,
    0x0042: 3 /* TIME */,
    0x0043: 1 /* DAY */,
    0x0044: 1 /* MONTH */,
    0x0045: 1 /* YEAR */,
    0x0046: 1 /* WEEKDAY */,
    0x0047: 1 /* HOUR */,
    0x0048: 1 /* MINUTE */,
    0x0049: 1 /* SECOND */,
    0x004a: 0 /* NOW */,
    0x004b: 1 /* AREAS */,
    0x004c: 1 /* ROWS */,
    0x004d: 1 /* COLUMNS */,
    0x004f: 2 /* ABSREF */,
    0x0050: 2 /* RELREF */,
    0x0053: 1 /* TRANSPOSE */,
    0x0055: 0 /* STEP */,
    0x0056: 1 /* TYPE */,
    0x0059: 0 /* CALLER */,
    0x005a: 1 /* DEREF */,
    0x005e: 0 /* ACTIVE.CELL */,
    0x005f: 0 /* SELECTION */,
    0x0061: 2 /* ATAN2 */,
    0x0062: 1 /* ASIN */,
    0x0063: 1 /* ACOS */,
    0x0065: 3 /* HLOOKUP */,
    0x0066: 3 /* VLOOKUP */,
    0x0069: 1 /* ISREF */,
    0x006a: 1 /* GET.FORMULA */,
    0x006c: 2 /* SET.VALUE */,
    0x006f: 1 /* CHAR */,
    0x0070: 1 /* LOWER */,
    0x0071: 1 /* UPPER */,
    0x0072: 1 /* PROPER */,
    0x0075: 2 /* EXACT */,
    0x0076: 1 /* TRIM */,
    0x0077: 4 /* REPLACE */,
    0x0079: 1 /* CODE */,
    0x007e: 1 /* ISERR */,
    0x007f: 1 /* ISTEXT */,
    0x0080: 1 /* ISNUMBER */,
    0x0081: 1 /* ISBLANK */,
    0x0082: 1 /* T */,
    0x0083: 1 /* N */,
    0x0085: 1 /* FCLOSE */,
    0x0086: 1 /* FSIZE */,
    0x0087: 1 /* FREADLN */,
    0x0088: 2 /* FREAD */,
    0x0089: 2 /* FWRITELN */,
    0x008a: 2 /* FWRITE */,
    0x008c: 1 /* DATEVALUE */,
    0x008d: 1 /* TIMEVALUE */,
    0x008e: 3 /* SLN */,
    0x008f: 4 /* SYD */,
    0x0090: 4 /* DDB */,
    0x00a1: 1 /* DIALOG.BOX */,
    0x00a2: 1 /* CLEAN */,
    0x00a3: 1 /* MDETERM */,
    0x00a4: 1 /* MINVERSE */,
    0x00a5: 2 /* MMULT */,
    0x00ac: 1 /* WHILE */,
    0x00af: 2 /* INITIATE */,
    0x00b0: 2 /* REQUEST */,
    0x00b1: 3 /* POKE */,
    0x00b2: 2 /* EXECUTE */,
    0x00b3: 1 /* TERMINATE */,
    0x00b8: 1 /* FACT */,
    0x00ba: 1 /* GET.WORKSPACE */,
    0x00bd: 3 /* DPRODUCT */,
    0x00be: 1 /* ISNONTEXT */,
    0x00c3: 3 /* DSTDEVP */,
    0x00c4: 3 /* DVARP */,
    0x00c5: 1 /* TRUNC */,
    0x00c6: 1 /* ISLOGICAL */,
    0x00c7: 3 /* DCOUNTA */,
    0x00c9: 1 /* UNREGISTER */,
    0x00cf: 4 /* REPLACEB */,
    0x00d2: 3 /* MIDB */,
    0x00d3: 1 /* LENB */,
    0x00d4: 2 /* ROUNDUP */,
    0x00d5: 2 /* ROUNDDOWN */,
    0x00d6: 1 /* ASC */,
    0x00d7: 1 /* DBCS */,
    0x00e1: 0 /* END.IF */,
    0x00e5: 1 /* SINH */,
    0x00e6: 1 /* COSH */,
    0x00e7: 1 /* TANH */,
    0x00e8: 1 /* ASINH */,
    0x00e9: 1 /* ACOSH */,
    0x00ea: 1 /* ATANH */,
    0x00eb: 3 /* DGET */,
    0x00f4: 1 /* INFO */,
    0x00f7: 4 /* DB */,
    0x00fc: 2 /* FREQUENCY */,
    0x0101: 1 /* EVALUATE */,
    0x0105: 1 /* ERROR.TYPE */,
    0x010f: 1 /* GAMMALN */,
    0x0111: 4 /* BINOMDIST */,
    0x0112: 2 /* CHIDIST */,
    0x0113: 2 /* CHIINV */,
    0x0114: 2 /* COMBIN */,
    0x0115: 3 /* CONFIDENCE */,
    0x0116: 3 /* CRITBINOM */,
    0x0117: 1 /* EVEN */,
    0x0118: 3 /* EXPONDIST */,
    0x0119: 3 /* FDIST */,
    0x011a: 3 /* FINV */,
    0x011b: 1 /* FISHER */,
    0x011c: 1 /* FISHERINV */,
    0x011d: 2 /* FLOOR */,
    0x011e: 4 /* GAMMADIST */,
    0x011f: 3 /* GAMMAINV */,
    0x0120: 2 /* CEILING */,
    0x0121: 4 /* HYPGEOMDIST */,
    0x0122: 3 /* LOGNORMDIST */,
    0x0123: 3 /* LOGINV */,
    0x0124: 3 /* NEGBINOMDIST */,
    0x0125: 4 /* NORMDIST */,
    0x0126: 1 /* NORMSDIST */,
    0x0127: 3 /* NORMINV */,
    0x0128: 1 /* NORMSINV */,
    0x0129: 3 /* STANDARDIZE */,
    0x012a: 1 /* ODD */,
    0x012b: 2 /* PERMUT */,
    0x012c: 3 /* POISSON */,
    0x012d: 3 /* TDIST */,
    0x012e: 4 /* WEIBULL */,
    0x012f: 2 /* SUMXMY2 */,
    0x0130: 2 /* SUMX2MY2 */,
    0x0131: 2 /* SUMX2PY2 */,
    0x0132: 2 /* CHITEST */,
    0x0133: 2 /* CORREL */,
    0x0134: 2 /* COVAR */,
    0x0135: 3 /* FORECAST */,
    0x0136: 2 /* FTEST */,
    0x0137: 2 /* INTERCEPT */,
    0x0138: 2 /* PEARSON */,
    0x0139: 2 /* RSQ */,
    0x013a: 2 /* STEYX */,
    0x013b: 2 /* SLOPE */,
    0x013c: 4 /* TTEST */,
    0x0145: 2 /* LARGE */,
    0x0146: 2 /* SMALL */,
    0x0147: 2 /* QUARTILE */,
    0x0148: 2 /* PERCENTILE */,
    0x014b: 2 /* TRIMMEAN */,
    0x014c: 2 /* TINV */,
    0x0151: 2 /* POWER */,
    0x0156: 1 /* RADIANS */,
    0x0157: 1 /* DEGREES */,
    0x015a: 2 /* COUNTIF */,
    0x015b: 1 /* COUNTBLANK */,
    0x015e: 4 /* ISPMT */,
    0x015f: 3 /* DATEDIF */,
    0x0160: 1 /* DATESTRING */,
    0x0161: 2 /* NUMBERSTRING */,
    0x0168: 1 /* PHONETIC */,
    0x0170: 1 /* BAHTTEXT */,
    0x0171: 1 /* THAIDAYOFWEEK */,
    0x0172: 1 /* THAIDIGIT */,
    0x0173: 1 /* THAIMONTHOFYEAR */,
    0x0174: 1 /* THAINUMSOUND */,
    0x0175: 1 /* THAINUMSTRING */,
    0x0176: 1 /* THAISTRINGLENGTH */,
    0x0177: 1 /* ISTHAIDIGIT */,
    0x0178: 1 /* ROUNDBAHTDOWN */,
    0x0179: 1 /* ROUNDBAHTUP */,
    0x017a: 1 /* THAIYEAR */,
    0x017e: 3 /* CUBEMEMBERPROPERTY */,
    0x0181: 1 /* HEX2DEC */,
    0x0188: 1 /* OCT2DEC */,
    0x0189: 1 /* BIN2DEC */,
    0x018c: 2 /* IMSUB */,
    0x018d: 2 /* IMDIV */,
    0x018e: 2 /* IMPOWER */,
    0x018f: 1 /* IMABS */,
    0x0190: 1 /* IMSQRT */,
    0x0191: 1 /* IMLN */,
    0x0192: 1 /* IMLOG2 */,
    0x0193: 1 /* IMLOG10 */,
    0x0194: 1 /* IMSIN */,
    0x0195: 1 /* IMCOS */,
    0x0196: 1 /* IMEXP */,
    0x0197: 1 /* IMARGUMENT */,
    0x0198: 1 /* IMCONJUGATE */,
    0x0199: 1 /* IMAGINARY */,
    0x019a: 1 /* IMREAL */,
    0x019e: 4 /* SERIESSUM */,
    0x019f: 1 /* FACTDOUBLE */,
    0x01a0: 1 /* SQRTPI */,
    0x01a1: 2 /* QUOTIENT */,
    0x01a4: 1 /* ISEVEN */,
    0x01a5: 1 /* ISODD */,
    0x01a6: 2 /* MROUND */,
    0x01a8: 1 /* ERFC */,
    0x01a9: 2 /* BESSELJ */,
    0x01aa: 2 /* BESSELK */,
    0x01ab: 2 /* BESSELY */,
    0x01ac: 2 /* BESSELI */,
    0x01ae: 3 /* XNPV */,
    0x01b6: 3 /* TBILLEQ */,
    0x01b7: 3 /* TBILLPRICE */,
    0x01b8: 3 /* TBILLYIELD */,
    0x01bb: 2 /* DOLLARDE */,
    0x01bc: 2 /* DOLLARFR */,
    0x01bd: 2 /* NOMINAL */,
    0x01be: 2 /* EFFECT */,
    0x01bf: 6 /* CUMPRINC */,
    0x01c0: 6 /* CUMIPMT */,
    0x01c1: 2 /* EDATE */,
    0x01c2: 2 /* EOMONTH */,
    0x01d0: 2 /* RANDBETWEEN */,
    0x01d4: 3 /* CONVERT */,
    0x01dc: 2 /* FVSCHEDULE */,
    0x01df: 1 /* CUBESETCOUNT */,
    0x01e0: 2 /* IFERROR */,
    0xffff: 0
  };
  /* [MS-XLSX] 2.2.3 Functions */
  /* [MS-XLSB] 2.5.97.10 Ftab */
  var XLSXFutureFunctions = {
    "_xlfn.ACOT": "ACOT",
    "_xlfn.ACOTH": "ACOTH",
    "_xlfn.AGGREGATE": "AGGREGATE",
    "_xlfn.ARABIC": "ARABIC",
    "_xlfn.AVERAGEIF": "AVERAGEIF",
    "_xlfn.AVERAGEIFS": "AVERAGEIFS",
    "_xlfn.BASE": "BASE",
    "_xlfn.BETA.DIST": "BETA.DIST",
    "_xlfn.BETA.INV": "BETA.INV",
    "_xlfn.BINOM.DIST": "BINOM.DIST",
    "_xlfn.BINOM.DIST.RANGE": "BINOM.DIST.RANGE",
    "_xlfn.BINOM.INV": "BINOM.INV",
    "_xlfn.BITAND": "BITAND",
    "_xlfn.BITLSHIFT": "BITLSHIFT",
    "_xlfn.BITOR": "BITOR",
    "_xlfn.BITRSHIFT": "BITRSHIFT",
    "_xlfn.BITXOR": "BITXOR",
    "_xlfn.CEILING.MATH": "CEILING.MATH",
    "_xlfn.CEILING.PRECISE": "CEILING.PRECISE",
    "_xlfn.CHISQ.DIST": "CHISQ.DIST",
    "_xlfn.CHISQ.DIST.RT": "CHISQ.DIST.RT",
    "_xlfn.CHISQ.INV": "CHISQ.INV",
    "_xlfn.CHISQ.INV.RT": "CHISQ.INV.RT",
    "_xlfn.CHISQ.TEST": "CHISQ.TEST",
    "_xlfn.COMBINA": "COMBINA",
    "_xlfn.CONCAT": "CONCAT",
    "_xlfn.CONFIDENCE.NORM": "CONFIDENCE.NORM",
    "_xlfn.CONFIDENCE.T": "CONFIDENCE.T",
    "_xlfn.COT": "COT",
    "_xlfn.COTH": "COTH",
    "_xlfn.COUNTIFS": "COUNTIFS",
    "_xlfn.COVARIANCE.P": "COVARIANCE.P",
    "_xlfn.COVARIANCE.S": "COVARIANCE.S",
    "_xlfn.CSC": "CSC",
    "_xlfn.CSCH": "CSCH",
    "_xlfn.DAYS": "DAYS",
    "_xlfn.DECIMAL": "DECIMAL",
    "_xlfn.ECMA.CEILING": "ECMA.CEILING",
    "_xlfn.ERF.PRECISE": "ERF.PRECISE",
    "_xlfn.ERFC.PRECISE": "ERFC.PRECISE",
    "_xlfn.EXPON.DIST": "EXPON.DIST",
    "_xlfn.F.DIST": "F.DIST",
    "_xlfn.F.DIST.RT": "F.DIST.RT",
    "_xlfn.F.INV": "F.INV",
    "_xlfn.F.INV.RT": "F.INV.RT",
    "_xlfn.F.TEST": "F.TEST",
    "_xlfn.FILTERXML": "FILTERXML",
    "_xlfn.FLOOR.MATH": "FLOOR.MATH",
    "_xlfn.FLOOR.PRECISE": "FLOOR.PRECISE",
    "_xlfn.FORECAST.ETS": "FORECAST.ETS",
    "_xlfn.FORECAST.ETS.CONFINT": "FORECAST.ETS.CONFINT",
    "_xlfn.FORECAST.ETS.SEASONALITY": "FORECAST.ETS.SEASONALITY",
    "_xlfn.FORECAST.ETS.STAT": "FORECAST.ETS.STAT",
    "_xlfn.FORECAST.LINEAR": "FORECAST.LINEAR",
    "_xlfn.FORMULATEXT": "FORMULATEXT",
    "_xlfn.GAMMA": "GAMMA",
    "_xlfn.GAMMA.DIST": "GAMMA.DIST",
    "_xlfn.GAMMA.INV": "GAMMA.INV",
    "_xlfn.GAMMALN.PRECISE": "GAMMALN.PRECISE",
    "_xlfn.GAUSS": "GAUSS",
    "_xlfn.HYPGEOM.DIST": "HYPGEOM.DIST",
    "_xlfn.IFERROR": "IFERROR",
    "_xlfn.IFNA": "IFNA",
    "_xlfn.IFS": "IFS",
    "_xlfn.IMCOSH": "IMCOSH",
    "_xlfn.IMCOT": "IMCOT",
    "_xlfn.IMCSC": "IMCSC",
    "_xlfn.IMCSCH": "IMCSCH",
    "_xlfn.IMSEC": "IMSEC",
    "_xlfn.IMSECH": "IMSECH",
    "_xlfn.IMSINH": "IMSINH",
    "_xlfn.IMTAN": "IMTAN",
    "_xlfn.ISFORMULA": "ISFORMULA",
    "_xlfn.ISO.CEILING": "ISO.CEILING",
    "_xlfn.ISOWEEKNUM": "ISOWEEKNUM",
    "_xlfn.LOGNORM.DIST": "LOGNORM.DIST",
    "_xlfn.LOGNORM.INV": "LOGNORM.INV",
    "_xlfn.MAXIFS": "MAXIFS",
    "_xlfn.MINIFS": "MINIFS",
    "_xlfn.MODE.MULT": "MODE.MULT",
    "_xlfn.MODE.SNGL": "MODE.SNGL",
    "_xlfn.MUNIT": "MUNIT",
    "_xlfn.NEGBINOM.DIST": "NEGBINOM.DIST",
    "_xlfn.NETWORKDAYS.INTL": "NETWORKDAYS.INTL",
    "_xlfn.NIGBINOM": "NIGBINOM",
    "_xlfn.NORM.DIST": "NORM.DIST",
    "_xlfn.NORM.INV": "NORM.INV",
    "_xlfn.NORM.S.DIST": "NORM.S.DIST",
    "_xlfn.NORM.S.INV": "NORM.S.INV",
    "_xlfn.NUMBERVALUE": "NUMBERVALUE",
    "_xlfn.PDURATION": "PDURATION",
    "_xlfn.PERCENTILE.EXC": "PERCENTILE.EXC",
    "_xlfn.PERCENTILE.INC": "PERCENTILE.INC",
    "_xlfn.PERCENTRANK.EXC": "PERCENTRANK.EXC",
    "_xlfn.PERCENTRANK.INC": "PERCENTRANK.INC",
    "_xlfn.PERMUTATIONA": "PERMUTATIONA",
    "_xlfn.PHI": "PHI",
    "_xlfn.POISSON.DIST": "POISSON.DIST",
    "_xlfn.QUARTILE.EXC": "QUARTILE.EXC",
    "_xlfn.QUARTILE.INC": "QUARTILE.INC",
    "_xlfn.QUERYSTRING": "QUERYSTRING",
    "_xlfn.RANK.AVG": "RANK.AVG",
    "_xlfn.RANK.EQ": "RANK.EQ",
    "_xlfn.RRI": "RRI",
    "_xlfn.SEC": "SEC",
    "_xlfn.SECH": "SECH",
    "_xlfn.SHEET": "SHEET",
    "_xlfn.SHEETS": "SHEETS",
    "_xlfn.SKEW.P": "SKEW.P",
    "_xlfn.STDEV.P": "STDEV.P",
    "_xlfn.STDEV.S": "STDEV.S",
    "_xlfn.SUMIFS": "SUMIFS",
    "_xlfn.SWITCH": "SWITCH",
    "_xlfn.T.DIST": "T.DIST",
    "_xlfn.T.DIST.2T": "T.DIST.2T",
    "_xlfn.T.DIST.RT": "T.DIST.RT",
    "_xlfn.T.INV": "T.INV",
    "_xlfn.T.INV.2T": "T.INV.2T",
    "_xlfn.T.TEST": "T.TEST",
    "_xlfn.TEXTJOIN": "TEXTJOIN",
    "_xlfn.UNICHAR": "UNICHAR",
    "_xlfn.UNICODE": "UNICODE",
    "_xlfn.VAR.P": "VAR.P",
    "_xlfn.VAR.S": "VAR.S",
    "_xlfn.WEBSERVICE": "WEBSERVICE",
    "_xlfn.WEIBULL.DIST": "WEIBULL.DIST",
    "_xlfn.WORKDAY.INTL": "WORKDAY.INTL",
    "_xlfn.XOR": "XOR",
    "_xlfn.Z.TEST": "Z.TEST"
  };

  /* Part 3 TODO: actually parse formulae */
  function ods_to_csf_formula(f) {
    if (f.slice(0, 3) == "of:") f = f.slice(3);
    /* 5.2 Basic Expressions */
    if (f.charCodeAt(0) == 61) {
      f = f.slice(1);
      if (f.charCodeAt(0) == 61) f = f.slice(1);
    }
    f = f.replace(/COM\.MICROSOFT\./g, "");
    /* Part 3 Section 5.8 References */
    f = f.replace(/\[((?:\.[A-Z]+[0-9]+)(?::\.[A-Z]+[0-9]+)?)\]/g, function(
      $$,
      $1
    ) {
      return $1.replace(/\./g, "");
    });
    /* TODO: something other than this */
    f = f.replace(/\[.(#[A-Z]*[?!])\]/g, "$1");
    return f.replace(/[;~]/g, ",").replace(/\|/g, ";");
  }

  function ods_to_csf_3D(r) {
    const a = r.split(":");
    const s = a[0].split(".")[0];
    return [
      s,
      a[0].split(".")[1] +
        (a.length > 1 ? `:${a[1].split(".")[1] || a[1].split(".")[0]}` : "")
    ];
  }

  let strs = {}; // shared strings
  let _ssfopts = {}; // spreadsheet formatting options

  RELS.WS = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet"
  ];

  /* global Map */
  const browser_has_Map = typeof Map !== "undefined";

  function get_sst_id(sst, str, rev) {
    let i = 0;
    const len = sst.length;
    if (rev) {
      if (
        browser_has_Map
          ? rev.has(str)
          : Object.prototype.hasOwnProperty.call(rev, str)
      ) {
        const revarr = browser_has_Map ? rev.get(str) : rev[str];
        for (; i < revarr.length; ++i) {
          if (sst[revarr[i]].t === str) {
            sst.Count++;
            return revarr[i];
          }
        }
      }
    } else
      for (; i < len; ++i) {
        if (sst[i].t === str) {
          sst.Count++;
          return i;
        }
      }
    sst[len] = { t: str };
    sst.Count++;
    sst.Unique++;
    if (rev) {
      if (browser_has_Map) {
        if (!rev.has(str)) rev.set(str, []);
        rev.get(str).push(len);
      } else {
        if (!Object.prototype.hasOwnProperty.call(rev, str)) rev[str] = [];
        rev[str].push(len);
      }
    }
    return len;
  }

  function col_obj_w(C, col) {
    const p = { min: C + 1, max: C + 1 };
    /* wch (chars), wpx (pixels) */
    let wch = -1;
    if (col.MDW) MDW = col.MDW;
    if (col.width != null) p.customWidth = 1;
    else if (col.wpx != null) wch = px2char(col.wpx);
    else if (col.wch != null) wch = col.wch;
    if (wch > -1) {
      p.width = char2width(wch);
      p.customWidth = 1;
    } else if (col.width != null) p.width = col.width;
    if (col.hidden) p.hidden = true;
    return p;
  }

  function default_margins(margins, mode) {
    if (!margins) return;
    let defs = [0.7, 0.7, 0.75, 0.75, 0.3, 0.3];
    if (mode == "xlml") defs = [1, 1, 1, 1, 0.5, 0.5];
    if (margins.left == null) margins.left = defs[0];
    if (margins.right == null) margins.right = defs[1];
    if (margins.top == null) margins.top = defs[2];
    if (margins.bottom == null) margins.bottom = defs[3];
    if (margins.header == null) margins.header = defs[4];
    if (margins.footer == null) margins.footer = defs[5];
  }

  function get_cell_style(styles, cell, opts) {
    let z = opts.revssf[cell.z != null ? cell.z : "General"];
    let i = 0x3c;
    const len = styles.length;
    if (z == null && opts.ssf) {
      for (; i < 0x188; ++i)
        if (opts.ssf[i] == null) {
          SSF.load(cell.z, i);
          // $FlowIgnore
          opts.ssf[i] = cell.z;
          opts.revssf[cell.z] = z = i;
          break;
        }
    }
    for (i = 0; i != len; ++i) if (styles[i].numFmtId === z) return i;
    styles[len] = {
      numFmtId: z,
      fontId: 0,
      fillId: 0,
      borderId: 0,
      xfId: 0,
      applyNumberFormat: 1
    };
    return len;
  }

  function safe_format(p, fmtid, fillid, opts, themes, styles) {
    try {
      if (opts.cellNF) p.z = SSF._table[fmtid];
    } catch (e) {
      if (opts.WTF) throw e;
    }
    if (p.t === "z") return;
    if (p.t === "d" && typeof p.v === "string") p.v = parseDate(p.v);
    if (!opts || opts.cellText !== false)
      try {
        if (SSF._table[fmtid] == null)
          SSF.load(SSFImplicit[fmtid] || "General", fmtid);
        if (p.t === "e") p.w = p.w || BErr[p.v];
        else if (fmtid === 0) {
          if (p.t === "n") {
            if ((p.v | 0) === p.v) p.w = SSF._general_int(p.v);
            else p.w = SSF._general_num(p.v);
          } else if (p.t === "d") {
            const dd = datenum(p.v);
            if ((dd | 0) === dd) p.w = SSF._general_int(dd);
            else p.w = SSF._general_num(dd);
          } else if (p.v === undefined) return "";
          else p.w = SSF._general(p.v, _ssfopts);
        } else if (p.t === "d") p.w = SSF.format(fmtid, datenum(p.v), _ssfopts);
        else p.w = SSF.format(fmtid, p.v, _ssfopts);
      } catch (e) {
        if (opts.WTF) throw e;
      }
    if (!opts.cellStyles) return;
    if (fillid != null)
      try {
        p.s = styles.Fills[fillid];
        if (p.s.fgColor && p.s.fgColor.theme && !p.s.fgColor.rgb) {
          p.s.fgColor.rgb = rgb_tint(
            themes.themeElements.clrScheme[p.s.fgColor.theme].rgb,
            p.s.fgColor.tint || 0
          );
          if (opts.WTF)
            p.s.fgColor.raw_rgb =
              themes.themeElements.clrScheme[p.s.fgColor.theme].rgb;
        }
        if (p.s.bgColor && p.s.bgColor.theme) {
          p.s.bgColor.rgb = rgb_tint(
            themes.themeElements.clrScheme[p.s.bgColor.theme].rgb,
            p.s.bgColor.tint || 0
          );
          if (opts.WTF)
            p.s.bgColor.raw_rgb =
              themes.themeElements.clrScheme[p.s.bgColor.theme].rgb;
        }
      } catch (e) {
        if (opts.WTF && styles.Fills) throw e;
      }
  }

  function parse_ws_xml_dim(ws, s) {
    const d = safe_decode_range(s);
    if (d.s.r <= d.e.r && d.s.c <= d.e.c && d.s.r >= 0 && d.s.c >= 0)
      ws["!ref"] = encode_range(d);
  }
  const mergecregex = /<(?:\w:)?mergeCell ref="[A-Z0-9:]+"\s*[\/]?>/g;
  const sheetdataregex = /<(?:\w+:)?sheetData[^>]*>([\s\S]*)<\/(?:\w+:)?sheetData>/;
  const hlinkregex = /<(?:\w:)?hyperlink [^>]*>/gm;
  const dimregex = /"(\w*:\w*)"/;
  const colregex = /<(?:\w:)?col\b[^>]*[\/]?>/g;
  const afregex = /<(?:\w:)?autoFilter[^>]*([\/]|>([\s\S]*)<\/(?:\w:)?autoFilter)>/g;
  const marginregex = /<(?:\w:)?pageMargins[^>]*\/>/g;
  const sheetprregex = /<(?:\w:)?sheetPr\b(?:[^>a-z][^>]*)?\/>/;
  const svsregex = /<(?:\w:)?sheetViews[^>]*(?:[\/]|>([\s\S]*)<\/(?:\w:)?sheetViews)>/;

  /* 18.3 Worksheets */
  function parse_ws_xml(data, opts, idx, rels, wb, themes, styles) {
    if (!data) return data;
    if (!rels) rels = { "!id": {} };
    if (DENSE != null && opts.dense == null) opts.dense = DENSE;

    /* 18.3.1.99 worksheet CT_Worksheet */
    const s = opts.dense ? [] : {};
    const refguess = { s: { r: 2000000, c: 2000000 }, e: { r: 0, c: 0 } };

    let data1 = "";
    let data2 = "";
    const mtch = data.match(sheetdataregex);
    if (mtch) {
      data1 = data.slice(0, mtch.index);
      data2 = data.slice(mtch.index + mtch[0].length);
    } else data1 = data2 = data;

    /* 18.3.1.82 sheetPr CT_SheetPr */
    const sheetPr = data1.match(sheetprregex);
    if (sheetPr) parse_ws_xml_sheetpr(sheetPr[0], s, wb, idx);

    /* 18.3.1.35 dimension CT_SheetDimension */
    let ridx = (data1.match(/<(?:\w*:)?dimension/) || { index: -1 }).index;
    if (ridx > 0) {
      const ref = data1.slice(ridx, ridx + 50).match(dimregex);
      if (ref) parse_ws_xml_dim(s, ref[1]);
    }

    /* 18.3.1.88 sheetViews CT_SheetViews */
    const svs = data1.match(svsregex);
    if (svs && svs[1]) parse_ws_xml_sheetviews(svs[1], wb);

    /* 18.3.1.17 cols CT_Cols */
    const columns = [];
    if (opts.cellStyles) {
      /* 18.3.1.13 col CT_Col */
      const cols = data1.match(colregex);
      if (cols) parse_ws_xml_cols(columns, cols);
    }

    /* 18.3.1.80 sheetData CT_SheetData ? */
    if (mtch) parse_ws_xml_data(mtch[1], s, opts, refguess, themes, styles);

    /* 18.3.1.2  autoFilter CT_AutoFilter */
    const afilter = data2.match(afregex);
    if (afilter) s["!autofilter"] = parse_ws_xml_autofilter(afilter[0]);

    /* 18.3.1.55 mergeCells CT_MergeCells */
    const merges = [];
    const _merge = data2.match(mergecregex);
    if (_merge)
      for (ridx = 0; ridx != _merge.length; ++ridx)
        merges[ridx] = safe_decode_range(
          _merge[ridx].slice(_merge[ridx].indexOf('"') + 1)
        );

    /* 18.3.1.48 hyperlinks CT_Hyperlinks */
    const hlink = data2.match(hlinkregex);
    if (hlink) parse_ws_xml_hlinks(s, hlink, rels);

    /* 18.3.1.62 pageMargins CT_PageMargins */
    const margins = data2.match(marginregex);
    if (margins) s["!margins"] = parse_ws_xml_margins(parsexmltag(margins[0]));

    if (
      !s["!ref"] &&
      refguess.e.c >= refguess.s.c &&
      refguess.e.r >= refguess.s.r
    )
      s["!ref"] = encode_range(refguess);
    if (opts.sheetRows > 0 && s["!ref"]) {
      const tmpref = safe_decode_range(s["!ref"]);
      if (opts.sheetRows <= +tmpref.e.r) {
        tmpref.e.r = opts.sheetRows - 1;
        if (tmpref.e.r > refguess.e.r) tmpref.e.r = refguess.e.r;
        if (tmpref.e.r < tmpref.s.r) tmpref.s.r = tmpref.e.r;
        if (tmpref.e.c > refguess.e.c) tmpref.e.c = refguess.e.c;
        if (tmpref.e.c < tmpref.s.c) tmpref.s.c = tmpref.e.c;
        s["!fullref"] = s["!ref"];
        s["!ref"] = encode_range(tmpref);
      }
    }
    if (columns.length > 0) s["!cols"] = columns;
    if (merges.length > 0) s["!merges"] = merges;
    return s;
  }

  /* 18.3.1.82-3 sheetPr CT_ChartsheetPr / CT_SheetPr */
  function parse_ws_xml_sheetpr(sheetPr, s, wb, idx) {
    const data = parsexmltag(sheetPr);
    if (!wb.Sheets[idx]) wb.Sheets[idx] = {};
    if (data.codeName) wb.Sheets[idx].CodeName = data.codeName;
  }

  function parse_ws_xml_hlinks(s, data, rels) {
    const dense = Array.isArray(s);
    for (let i = 0; i != data.length; ++i) {
      const val = parsexmltag(utf8read(data[i]), true);
      if (!val.ref) return;
      let rel = ((rels || {})["!id"] || [])[val.id];
      if (rel) {
        val.Target = rel.Target;
        if (val.location) val.Target += `#${val.location}`;
      } else {
        val.Target = `#${val.location}`;
        rel = { Target: val.Target, TargetMode: "Internal" };
      }
      val.Rel = rel;
      if (val.tooltip) {
        val.Tooltip = val.tooltip;
        delete val.tooltip;
      }
      const rng = safe_decode_range(val.ref);
      for (let R = rng.s.r; R <= rng.e.r; ++R)
        for (let C = rng.s.c; C <= rng.e.c; ++C) {
          const addr = encode_cell({ c: C, r: R });
          if (dense) {
            if (!s[R]) s[R] = [];
            if (!s[R][C]) s[R][C] = { t: "z", v: undefined };
            s[R][C].l = val;
          } else {
            if (!s[addr]) s[addr] = { t: "z", v: undefined };
            s[addr].l = val;
          }
        }
    }
  }

  function parse_ws_xml_margins(margin) {
    const o = {};
    ["left", "right", "top", "bottom", "header", "footer"].forEach(function(k) {
      if (margin[k]) o[k] = parseFloat(margin[k]);
    });
    return o;
  }

  function parse_ws_xml_cols(columns, cols) {
    let seencol = false;
    for (let coli = 0; coli != cols.length; ++coli) {
      const coll = parsexmltag(cols[coli], true);
      if (coll.hidden) coll.hidden = parsexmlbool(coll.hidden);
      let colm = parseInt(coll.min, 10) - 1;
      const colM = parseInt(coll.max, 10) - 1;
      delete coll.min;
      delete coll.max;
      coll.width = +coll.width;
      if (!seencol && coll.width) {
        seencol = true;
        find_mdw_colw(coll.width);
      }
      process_col(coll);
      while (colm <= colM) columns[colm++] = dup(coll);
    }
  }

  function parse_ws_xml_autofilter(data) {
    const o = { ref: (data.match(/ref="([^"]*)"/) || [])[1] };
    return o;
  }

  /* 18.3.1.88 sheetViews CT_SheetViews */
  /* 18.3.1.87 sheetView CT_SheetView */
  const sviewregex = /<(?:\w:)?sheetView(?:[^>a-z][^>]*)?\/?>/;
  function parse_ws_xml_sheetviews(data, wb) {
    if (!wb.Views) wb.Views = [{}];
    (data.match(sviewregex) || []).forEach(function(r, i) {
      const tag = parsexmltag(r);
      // $FlowIgnore
      if (!wb.Views[i]) wb.Views[i] = {};
      // $FlowIgnore
      if (parsexmlbool(tag.rightToLeft)) wb.Views[i].RTL = true;
    });
  }

  var parse_ws_xml_data = (function() {
    const cellregex = /<(?:\w+:)?c[ >]/;
    const rowregex = /<\/(?:\w+:)?row>/;
    const rregex = /r=["']([^"']*)["']/;
    const isregex = /<(?:\w+:)?is>([\S\s]*?)<\/(?:\w+:)?is>/;
    const refregex = /ref=["']([^"']*)["']/;
    const match_v = matchtag("v");
    const match_f = matchtag("f");

    return function parse_ws_xml_data(sdata, s, opts, guess, themes, styles) {
      let ri = 0;
      let x = "";
      let cells = [];
      let cref = [];
      let idx = 0;
      let i = 0;
      let cc = 0;
      let d = "";
      let p;
      let tag;
      let tagr = 0;
      let tagc = 0;
      let sstr;
      let ftag;
      let fmtid = 0;
      let fillid = 0;
      const do_format = Array.isArray(styles.CellXf);
      let cf;
      const arrayf = [];
      const sharedf = [];
      const dense = Array.isArray(s);
      const rows = [];
      let rowobj = {};
      let rowrite = false;
      for (
        let marr = sdata.split(rowregex), mt = 0, marrlen = marr.length;
        mt != marrlen;
        ++mt
      ) {
        x = marr[mt].trim();
        const xlen = x.length;
        if (xlen === 0) continue;

        /* 18.3.1.73 row CT_Row */
        for (ri = 0; ri < xlen; ++ri) if (x.charCodeAt(ri) === 62) break;
        ++ri;
        tag = parsexmltag(x.slice(0, ri), true);
        tagr = tag.r != null ? parseInt(tag.r, 10) : tagr + 1;
        tagc = -1;
        if (opts.sheetRows && opts.sheetRows < tagr) continue;
        if (guess.s.r > tagr - 1) guess.s.r = tagr - 1;
        if (guess.e.r < tagr - 1) guess.e.r = tagr - 1;

        if (opts && opts.cellStyles) {
          rowobj = {};
          rowrite = false;
          if (tag.ht) {
            rowrite = true;
            rowobj.hpt = parseFloat(tag.ht);
            rowobj.hpx = pt2px(rowobj.hpt);
          }
          if (tag.hidden == "1") {
            rowrite = true;
            rowobj.hidden = true;
          }
          if (tag.outlineLevel != null) {
            rowrite = true;
            rowobj.level = +tag.outlineLevel;
          }
          if (rowrite) rows[tagr - 1] = rowobj;
        }

        /* 18.3.1.4 c CT_Cell */
        cells = x.slice(ri).split(cellregex);
        for (var rslice = 0; rslice != cells.length; ++rslice)
          if (cells[rslice].trim().charAt(0) != "<") break;
        cells = cells.slice(rslice);
        for (ri = 0; ri != cells.length; ++ri) {
          x = cells[ri].trim();
          if (x.length === 0) continue;
          cref = x.match(rregex);
          idx = ri;
          i = 0;
          cc = 0;
          x = `<c ${x.slice(0, 1) == "<" ? ">" : ""}${x}`;
          if (cref != null && cref.length === 2) {
            idx = 0;
            d = cref[1];
            for (i = 0; i != d.length; ++i) {
              if ((cc = d.charCodeAt(i) - 64) < 1 || cc > 26) break;
              idx = 26 * idx + cc;
            }
            --idx;
            tagc = idx;
          } else ++tagc;
          for (i = 0; i != x.length; ++i) if (x.charCodeAt(i) === 62) break;
          ++i;
          tag = parsexmltag(x.slice(0, i), true);
          if (!tag.r) tag.r = encode_cell({ r: tagr - 1, c: tagc });
          d = x.slice(i);
          p = { t: "" };

          if ((cref = d.match(match_v)) != null && cref[1] !== "")
            p.v = unescapexml(cref[1]);
          if (opts.cellFormula) {
            if ((cref = d.match(match_f)) != null && cref[1] !== "") {
              /* TODO: match against XLSXFutureFunctions */
              p.f = unescapexml(utf8read(cref[1]));
              if (!opts.xlfn) p.f = _xlfn(p.f);
              if (cref[0].indexOf('t="array"') > -1) {
                p.F = (d.match(refregex) || [])[1];
                if (p.F.indexOf(":") > -1)
                  arrayf.push([safe_decode_range(p.F), p.F]);
              } else if (cref[0].indexOf('t="shared"') > -1) {
                // TODO: parse formula
                ftag = parsexmltag(cref[0]);
                let ___f = unescapexml(utf8read(cref[1]));
                if (!opts.xlfn) ___f = _xlfn(___f);
                sharedf[parseInt(ftag.si, 10)] = [ftag, ___f, tag.r];
              }
            } else if ((cref = d.match(/<f[^>]*\/>/))) {
              ftag = parsexmltag(cref[0]);
              if (sharedf[ftag.si])
                p.f = shift_formula_xlsx(
                  sharedf[ftag.si][1],
                  sharedf[ftag.si][2] /* [0].ref */,
                  tag.r
                );
            }
            /* TODO: factor out contains logic */
            const _tag = decode_cell(tag.r);
            for (i = 0; i < arrayf.length; ++i)
              if (_tag.r >= arrayf[i][0].s.r && _tag.r <= arrayf[i][0].e.r)
                if (_tag.c >= arrayf[i][0].s.c && _tag.c <= arrayf[i][0].e.c)
                  p.F = arrayf[i][1];
          }

          if (tag.t == null && p.v === undefined) {
            if (p.f || p.F) {
              p.v = 0;
              p.t = "n";
            } else if (!opts.sheetStubs) continue;
            else p.t = "z";
          } else p.t = tag.t || "n";
          if (guess.s.c > tagc) guess.s.c = tagc;
          if (guess.e.c < tagc) guess.e.c = tagc;
          /* 18.18.11 t ST_CellType */
          switch (p.t) {
            case "n":
              if (p.v == "" || p.v == null) {
                if (!opts.sheetStubs) continue;
                p.t = "z";
              } else p.v = parseFloat(p.v);
              break;
            case "s":
              if (typeof p.v === "undefined") {
                if (!opts.sheetStubs) continue;
                p.t = "z";
              } else {
                sstr = strs[parseInt(p.v, 10)];
                p.v = sstr.t;
                p.r = sstr.r;
                if (opts.cellHTML) p.h = sstr.h;
              }
              break;
            case "str":
              p.t = "s";
              p.v = p.v != null ? utf8read(p.v) : "";
              if (opts.cellHTML) p.h = escapehtml(p.v);
              break;
            case "inlineStr":
              cref = d.match(isregex);
              p.t = "s";
              if (cref != null && (sstr = parse_si(cref[1]))) {
                p.v = sstr.t;
                if (opts.cellHTML) p.h = sstr.h;
              } else p.v = "";
              break;
            case "b":
              p.v = parsexmlbool(p.v);
              break;
            case "d":
              if (opts.cellDates) p.v = parseDate(p.v, 1);
              else {
                p.v = datenum(parseDate(p.v, 1));
                p.t = "n";
              }
              break;
            /* error string in .w, number in .v */
            case "e":
              if (!opts || opts.cellText !== false) p.w = p.v;
              p.v = RBErr[p.v];
              break;
          }
          /* formatting */
          fmtid = fillid = 0;
          cf = null;
          if (do_format && tag.s !== undefined) {
            cf = styles.CellXf[tag.s];
            if (cf != null) {
              if (cf.numFmtId != null) fmtid = cf.numFmtId;
              if (opts.cellStyles) {
                if (cf.fillId != null) fillid = cf.fillId;
              }
            }
          }
          safe_format(p, fmtid, fillid, opts, themes, styles);
          if (
            opts.cellDates &&
            do_format &&
            p.t == "n" &&
            SSF.is_date(SSF._table[fmtid])
          ) {
            p.t = "d";
            p.v = numdate(p.v);
          }
          if (dense) {
            const _r = decode_cell(tag.r);
            if (!s[_r.r]) s[_r.r] = [];
            s[_r.r][_r.c] = p;
          } else s[tag.r] = p;
        }
      }
      if (rows.length > 0) s["!rows"] = rows;
    };
  })();

  /* [MS-XLSB] 2.4.726 BrtRowHdr */
  function parse_BrtRowHdr(data, length) {
    const z = {};
    const tgt = data.l + length;
    z.r = data.read_shift(4);
    data.l += 4; // TODO: ixfe
    const miyRw = data.read_shift(2);
    data.l += 1; // TODO: top/bot padding
    const flags = data.read_shift(1);
    data.l = tgt;
    if (flags & 0x07) z.level = flags & 0x07;
    if (flags & 0x10) z.hidden = true;
    if (flags & 0x20) z.hpt = miyRw / 20;
    return z;
  }

  /* [MS-XLSB] 2.4.820 BrtWsDim */
  const parse_BrtWsDim = parse_UncheckedRfX;

  /* [MS-XLSB] 2.4.821 BrtWsFmtInfo */
  function parse_BrtWsFmtInfo() {}
  // function write_BrtWsFmtInfo(ws, o) { }

  /* [MS-XLSB] 2.4.823 BrtWsProp */
  function parse_BrtWsProp(data, length) {
    const z = {};
    /* TODO: pull flags */
    data.l += 19;
    z.name = parse_XLSBCodeName(data, length - 19);
    return z;
  }

  /* [MS-XLSB] 2.4.306 BrtCellBlank */
  function parse_BrtCellBlank(data) {
    const cell = parse_XLSBCell(data);
    return [cell];
  }
  function write_BrtCellBlank(cell, ncell, o) {
    if (o == null) o = new_buf(8);
    return write_XLSBCell(ncell, o);
  }

  /* [MS-XLSB] 2.4.307 BrtCellBool */
  function parse_BrtCellBool(data) {
    const cell = parse_XLSBCell(data);
    const fBool = data.read_shift(1);
    return [cell, fBool, "b"];
  }
  function write_BrtCellBool(cell, ncell, o) {
    if (o == null) o = new_buf(9);
    write_XLSBCell(ncell, o);
    o.write_shift(1, cell.v ? 1 : 0);
    return o;
  }

  /* [MS-XLSB] 2.4.308 BrtCellError */
  function parse_BrtCellError(data) {
    const cell = parse_XLSBCell(data);
    const bError = data.read_shift(1);
    return [cell, bError, "e"];
  }

  /* [MS-XLSB] 2.4.311 BrtCellIsst */
  function parse_BrtCellIsst(data) {
    const cell = parse_XLSBCell(data);
    const isst = data.read_shift(4);
    return [cell, isst, "s"];
  }
  function write_BrtCellIsst(cell, ncell, o) {
    if (o == null) o = new_buf(12);
    write_XLSBCell(ncell, o);
    o.write_shift(4, ncell.v);
    return o;
  }

  /* [MS-XLSB] 2.4.313 BrtCellReal */
  function parse_BrtCellReal(data) {
    const cell = parse_XLSBCell(data);
    const value = parse_Xnum(data);
    return [cell, value, "n"];
  }
  function write_BrtCellReal(cell, ncell, o) {
    if (o == null) o = new_buf(16);
    write_XLSBCell(ncell, o);
    write_Xnum(cell.v, o);
    return o;
  }

  /* [MS-XLSB] 2.4.314 BrtCellRk */
  function parse_BrtCellRk(data) {
    const cell = parse_XLSBCell(data);
    const value = parse_RkNumber(data);
    return [cell, value, "n"];
  }
  function write_BrtCellRk(cell, ncell, o) {
    if (o == null) o = new_buf(12);
    write_XLSBCell(ncell, o);
    write_RkNumber(cell.v, o);
    return o;
  }

  /* [MS-XLSB] 2.4.317 BrtCellSt */
  function parse_BrtCellSt(data) {
    const cell = parse_XLSBCell(data);
    const value = parse_XLWideString(data);
    return [cell, value, "str"];
  }
  function write_BrtCellSt(cell, ncell, o) {
    if (o == null) o = new_buf(12 + 4 * cell.v.length);
    write_XLSBCell(ncell, o);
    write_XLWideString(cell.v, o);
    return o.length > o.l ? o.slice(0, o.l) : o;
  }

  /* [MS-XLSB] 2.4.653 BrtFmlaBool */
  function parse_BrtFmlaBool(data, length, opts) {
    const end = data.l + length;
    const cell = parse_XLSBCell(data);
    cell.r = opts["!row"];
    const value = data.read_shift(1);
    const o = [cell, value, "b"];
    if (opts.cellFormula) {
      data.l += 2;
      const formula = parse_XLSBCellParsedFormula(data, end - data.l, opts);
      o[3] = stringify_formula(
        formula,
        null /* range */,
        cell,
        opts.supbooks,
        opts
      ); /* TODO */
    } else data.l = end;
    return o;
  }

  /* [MS-XLSB] 2.4.654 BrtFmlaError */
  function parse_BrtFmlaError(data, length, opts) {
    const end = data.l + length;
    const cell = parse_XLSBCell(data);
    cell.r = opts["!row"];
    const value = data.read_shift(1);
    const o = [cell, value, "e"];
    if (opts.cellFormula) {
      data.l += 2;
      const formula = parse_XLSBCellParsedFormula(data, end - data.l, opts);
      o[3] = stringify_formula(
        formula,
        null /* range */,
        cell,
        opts.supbooks,
        opts
      ); /* TODO */
    } else data.l = end;
    return o;
  }

  /* [MS-XLSB] 2.4.655 BrtFmlaNum */
  function parse_BrtFmlaNum(data, length, opts) {
    const end = data.l + length;
    const cell = parse_XLSBCell(data);
    cell.r = opts["!row"];
    const value = parse_Xnum(data);
    const o = [cell, value, "n"];
    if (opts.cellFormula) {
      data.l += 2;
      const formula = parse_XLSBCellParsedFormula(data, end - data.l, opts);
      o[3] = stringify_formula(
        formula,
        null /* range */,
        cell,
        opts.supbooks,
        opts
      ); /* TODO */
    } else data.l = end;
    return o;
  }

  /* [MS-XLSB] 2.4.656 BrtFmlaString */
  function parse_BrtFmlaString(data, length, opts) {
    const end = data.l + length;
    const cell = parse_XLSBCell(data);
    cell.r = opts["!row"];
    const value = parse_XLWideString(data);
    const o = [cell, value, "str"];
    if (opts.cellFormula) {
      data.l += 2;
      const formula = parse_XLSBCellParsedFormula(data, end - data.l, opts);
      o[3] = stringify_formula(
        formula,
        null /* range */,
        cell,
        opts.supbooks,
        opts
      ); /* TODO */
    } else data.l = end;
    return o;
  }

  /* [MS-XLSB] 2.4.682 BrtMergeCell */
  const parse_BrtMergeCell = parse_UncheckedRfX;
  const write_BrtMergeCell = write_UncheckedRfX;
  /* [MS-XLSB] 2.4.107 BrtBeginMergeCells */
  function write_BrtBeginMergeCells(cnt, o) {
    if (o == null) o = new_buf(4);
    o.write_shift(4, cnt);
    return o;
  }

  /* [MS-XLSB] 2.4.662 BrtHLink */
  function parse_BrtHLink(data, length) {
    const end = data.l + length;
    const rfx = parse_UncheckedRfX(data, 16);
    const relId = parse_XLNullableWideString(data);
    const loc = parse_XLWideString(data);
    const tooltip = parse_XLWideString(data);
    const display = parse_XLWideString(data);
    data.l = end;
    const o = { rfx, relId, loc, display };
    if (tooltip) o.Tooltip = tooltip;
    return o;
  }
  function write_BrtHLink(l, rId) {
    const o = new_buf(
      50 + 4 * (l[1].Target.length + (l[1].Tooltip || "").length)
    );
    write_UncheckedRfX({ s: decode_cell(l[0]), e: decode_cell(l[0]) }, o);
    write_RelID(`rId${rId}`, o);
    const locidx = l[1].Target.indexOf("#");
    const loc = locidx == -1 ? "" : l[1].Target.slice(locidx + 1);
    write_XLWideString(loc || "", o);
    write_XLWideString(l[1].Tooltip || "", o);
    write_XLWideString("", o);
    return o.slice(0, o.l);
  }

  /* [MS-XLSB] 2.4.692 BrtPane */
  function parse_BrtPane(/* data, length, opts */) {}

  /* [MS-XLSB] 2.4.6 BrtArrFmla */
  function parse_BrtArrFmla(data, length, opts) {
    const end = data.l + length;
    const rfx = parse_RfX(data, 16);
    const fAlwaysCalc = data.read_shift(1);
    const o = [rfx];
    o[2] = fAlwaysCalc;
    if (opts.cellFormula) {
      const formula = parse_XLSBArrayParsedFormula(data, end - data.l, opts);
      o[1] = formula;
    } else data.l = end;
    return o;
  }

  /* [MS-XLSB] 2.4.750 BrtShrFmla */
  function parse_BrtShrFmla(data, length, opts) {
    const end = data.l + length;
    const rfx = parse_UncheckedRfX(data, 16);
    const o = [rfx];
    if (opts.cellFormula) {
      const formula = parse_XLSBSharedParsedFormula(data, end - data.l, opts);
      o[1] = formula;
      data.l = end;
    } else data.l = end;
    return o;
  }

  /* [MS-XLSB] 2.4.323 BrtColInfo */
  /* TODO: once XLS ColInfo is set, combine the functions */
  function write_BrtColInfo(C, col, o) {
    if (o == null) o = new_buf(18);
    const p = col_obj_w(C, col);
    o.write_shift(-4, C);
    o.write_shift(-4, C);
    o.write_shift(4, (p.width || 10) * 256);
    o.write_shift(4, 0 /* ixfe */); // style
    let flags = 0;
    if (col.hidden) flags |= 0x01;
    if (typeof p.width === "number") flags |= 0x02;
    if (col.level) flags |= col.level << 8;
    o.write_shift(2, flags); // bit flag
    return o;
  }

  /* [MS-XLSB] 2.4.678 BrtMargins */
  const BrtMarginKeys = ["left", "right", "top", "bottom", "header", "footer"];
  function parse_BrtMargins(data) {
    const margins = {};
    BrtMarginKeys.forEach(function(k) {
      margins[k] = parse_Xnum(data, 8);
    });
    return margins;
  }

  /* [MS-XLSB] 2.4.299 BrtBeginWsView */
  function parse_BrtBeginWsView(data) {
    const f = data.read_shift(2);
    data.l += 28;
    return { RTL: f & 0x20 };
  }
  function write_BrtBeginWsView(ws, Workbook, o) {
    if (o == null) o = new_buf(30);
    let f = 0x39c;
    if ((((Workbook || {}).Views || [])[0] || {}).RTL) f |= 0x20;
    o.write_shift(2, f); // bit flag
    o.write_shift(4, 0);
    o.write_shift(4, 0); // view first row
    o.write_shift(4, 0); // view first col
    o.write_shift(1, 0); // gridline color ICV
    o.write_shift(1, 0);
    o.write_shift(2, 0);
    o.write_shift(2, 100); // zoom scale
    o.write_shift(2, 0);
    o.write_shift(2, 0);
    o.write_shift(2, 0);
    o.write_shift(4, 0); // workbook view id
    return o;
  }

  /* [MS-XLSB] 2.4.309 BrtCellIgnoreEC */
  function write_BrtCellIgnoreEC(ref) {
    const o = new_buf(24);
    o.write_shift(4, 4);
    o.write_shift(4, 1);
    write_UncheckedRfX(ref, o);
    return o;
  }

  /* [MS-XLSB] 2.4.748 BrtSheetProtection */
  function write_BrtSheetProtection(sp, o) {
    if (o == null) o = new_buf(16 * 4 + 2);
    o.write_shift(
      2,
      sp.password ? crypto_CreatePasswordVerifier_Method1(sp.password) : 0
    );
    o.write_shift(4, 1); // this record should not be written if no protection
    [
      ["objects", false], // fObjects
      ["scenarios", false], // fScenarios
      ["formatCells", true], // fFormatCells
      ["formatColumns", true], // fFormatColumns
      ["formatRows", true], // fFormatRows
      ["insertColumns", true], // fInsertColumns
      ["insertRows", true], // fInsertRows
      ["insertHyperlinks", true], // fInsertHyperlinks
      ["deleteColumns", true], // fDeleteColumns
      ["deleteRows", true], // fDeleteRows
      ["selectLockedCells", false], // fSelLockedCells
      ["sort", true], // fSort
      ["autoFilter", true], // fAutoFilter
      ["pivotTables", true], // fPivotTables
      ["selectUnlockedCells", false] // fSelUnlockedCells
    ].forEach(function(n) {
      if (n[1]) o.write_shift(4, sp[n[0]] != null && !sp[n[0]] ? 1 : 0);
      else o.write_shift(4, sp[n[0]] != null && sp[n[0]] ? 0 : 1);
    });
    return o;
  }

  function parse_BrtDVal(/* data, length, opts */) {}
  function parse_BrtDVal14(/* data, length, opts */) {}
  /* [MS-XLSB] 2.1.7.61 Worksheet */
  function parse_ws_bin(data, _opts, idx, rels, wb, themes, styles) {
    if (!data) return data;
    const opts = _opts || {};
    if (!rels) rels = { "!id": {} };
    if (DENSE != null && opts.dense == null) opts.dense = DENSE;
    const s = opts.dense ? [] : {};

    let ref;
    const refguess = { s: { r: 2000000, c: 2000000 }, e: { r: 0, c: 0 } };

    const state = [];
    let pass = false;
    let end = false;
    let row;
    let p;
    let cf;
    let R;
    let C;
    let addr;
    let sstr;
    let rr;
    let cell;
    const merges = [];
    opts.biff = 12;
    opts["!row"] = 0;

    let ai = 0;
    let af = false;

    const arrayf = [];
    const sharedf = {};
    const supbooks = opts.supbooks || wb.supbooks || [[]];
    supbooks.sharedf = sharedf;
    supbooks.arrayf = arrayf;
    supbooks.SheetNames =
      wb.SheetNames ||
      wb.Sheets.map(function(x) {
        return x.name;
      });
    if (!opts.supbooks) {
      opts.supbooks = supbooks;
      if (wb.Names)
        for (let i = 0; i < wb.Names.length; ++i)
          supbooks[0][i + 1] = wb.Names[i];
    }

    const colinfo = [];
    const rowinfo = [];
    let seencol = false;

    recordhopper(
      data,
      function ws_parse(val, R_n, RT) {
        if (end) return;
        switch (RT) {
          case 0x0094 /* 'BrtWsDim' */:
            ref = val;
            break;
          case 0x0000 /* 'BrtRowHdr' */:
            row = val;
            if (opts.sheetRows && opts.sheetRows <= row.r) end = true;
            rr = encode_row((R = row.r));
            opts["!row"] = row.r;
            if (val.hidden || val.hpt || val.level != null) {
              if (val.hpt) val.hpx = pt2px(val.hpt);
              rowinfo[val.r] = val;
            }
            break;

          case 0x0002: /* 'BrtCellRk' */
          case 0x0003: /* 'BrtCellError' */
          case 0x0004: /* 'BrtCellBool' */
          case 0x0005: /* 'BrtCellReal' */
          case 0x0006: /* 'BrtCellSt' */
          case 0x0007: /* 'BrtCellIsst' */
          case 0x0008: /* 'BrtFmlaString' */
          case 0x0009: /* 'BrtFmlaNum' */
          case 0x000a: /* 'BrtFmlaBool' */
          case 0x000b /* 'BrtFmlaError' */:
            p = { t: val[2] };
            switch (val[2]) {
              case "n":
                p.v = val[1];
                break;
              case "s":
                sstr = strs[val[1]];
                p.v = sstr.t;
                p.r = sstr.r;
                break;
              case "b":
                p.v = !!val[1];
                break;
              case "e":
                p.v = val[1];
                if (opts.cellText !== false) p.w = BErr[p.v];
                break;
              case "str":
                p.t = "s";
                p.v = val[1];
                break;
            }
            if ((cf = styles.CellXf[val[0].iStyleRef]))
              safe_format(p, cf.numFmtId, null, opts, themes, styles);
            C = val[0].c;
            if (opts.dense) {
              if (!s[R]) s[R] = [];
              s[R][C] = p;
            } else s[encode_col(C) + rr] = p;
            if (opts.cellFormula) {
              af = false;
              for (ai = 0; ai < arrayf.length; ++ai) {
                const aii = arrayf[ai];
                if (row.r >= aii[0].s.r && row.r <= aii[0].e.r)
                  if (C >= aii[0].s.c && C <= aii[0].e.c) {
                    p.F = encode_range(aii[0]);
                    af = true;
                  }
              }
              if (!af && val.length > 3) p.f = val[3];
            }
            if (refguess.s.r > row.r) refguess.s.r = row.r;
            if (refguess.s.c > C) refguess.s.c = C;
            if (refguess.e.r < row.r) refguess.e.r = row.r;
            if (refguess.e.c < C) refguess.e.c = C;
            if (
              opts.cellDates &&
              cf &&
              p.t == "n" &&
              SSF.is_date(SSF._table[cf.numFmtId])
            ) {
              const _d = SSF.parse_date_code(p.v);
              if (_d) {
                p.t = "d";
                p.v = new Date(_d.y, _d.m - 1, _d.d, _d.H, _d.M, _d.S, _d.u);
              }
            }
            break;

          case 0x0001 /* 'BrtCellBlank' */:
            if (!opts.sheetStubs || pass) break;
            p = { t: "z", v: undefined };
            C = val[0].c;
            if (opts.dense) {
              if (!s[R]) s[R] = [];
              s[R][C] = p;
            } else s[encode_col(C) + rr] = p;
            if (refguess.s.r > row.r) refguess.s.r = row.r;
            if (refguess.s.c > C) refguess.s.c = C;
            if (refguess.e.r < row.r) refguess.e.r = row.r;
            if (refguess.e.c < C) refguess.e.c = C;
            break;

          case 0x00b0 /* 'BrtMergeCell' */:
            merges.push(val);
            break;

          case 0x01ee /* 'BrtHLink' */:
            var rel = rels["!id"][val.relId];
            if (rel) {
              val.Target = rel.Target;
              if (val.loc) val.Target += `#${val.loc}`;
              val.Rel = rel;
            } else if (val.relId == "") {
              val.Target = `#${val.loc}`;
            }
            for (R = val.rfx.s.r; R <= val.rfx.e.r; ++R)
              for (C = val.rfx.s.c; C <= val.rfx.e.c; ++C) {
                if (opts.dense) {
                  if (!s[R]) s[R] = [];
                  if (!s[R][C]) s[R][C] = { t: "z", v: undefined };
                  s[R][C].l = val;
                } else {
                  addr = encode_cell({ c: C, r: R });
                  if (!s[addr]) s[addr] = { t: "z", v: undefined };
                  s[addr].l = val;
                }
              }
            break;

          case 0x01aa /* 'BrtArrFmla' */:
            if (!opts.cellFormula) break;
            arrayf.push(val);
            cell = opts.dense ? s[R][C] : s[encode_col(C) + rr];
            cell.f = stringify_formula(
              val[1],
              refguess,
              { r: row.r, c: C },
              supbooks,
              opts
            );
            cell.F = encode_range(val[0]);
            break;
          case 0x01ab /* 'BrtShrFmla' */:
            if (!opts.cellFormula) break;
            sharedf[encode_cell(val[0].s)] = val[1];
            cell = opts.dense ? s[R][C] : s[encode_col(C) + rr];
            cell.f = stringify_formula(
              val[1],
              refguess,
              { r: row.r, c: C },
              supbooks,
              opts
            );
            break;

          /* identical to 'ColInfo' in XLS */
          case 0x003c /* 'BrtColInfo' */:
            if (!opts.cellStyles) break;
            while (val.e >= val.s) {
              colinfo[val.e--] = {
                width: val.w / 256,
                hidden: !!(val.flags & 0x01),
                level: val.level
              };
              if (!seencol) {
                seencol = true;
                find_mdw_colw(val.w / 256);
              }
              process_col(colinfo[val.e + 1]);
            }
            break;

          case 0x00a1 /* 'BrtBeginAFilter' */:
            s["!autofilter"] = { ref: encode_range(val) };
            break;

          case 0x01dc /* 'BrtMargins' */:
            s["!margins"] = val;
            break;

          case 0x0093 /* 'BrtWsProp' */:
            if (!wb.Sheets[idx]) wb.Sheets[idx] = {};
            if (val.name) wb.Sheets[idx].CodeName = val.name;
            break;

          case 0x0089 /* 'BrtBeginWsView' */:
            if (!wb.Views) wb.Views = [{}];
            if (!wb.Views[0]) wb.Views[0] = {};
            if (val.RTL) wb.Views[0].RTL = true;
            break;

          case 0x01e5 /* 'BrtWsFmtInfo' */:
            break;

          case 0x0040: /* 'BrtDVal' */
          case 0x041d /* 'BrtDVal14' */:
            break;

          case 0x0097 /* 'BrtPane' */:
            break;
          case 0x00af: /* 'BrtAFilterDateGroupItem' */
          case 0x0284: /* 'BrtActiveX' */
          case 0x0271: /* 'BrtBigName' */
          case 0x0232: /* 'BrtBkHim' */
          case 0x018c: /* 'BrtBrk' */
          case 0x0458: /* 'BrtCFIcon' */
          case 0x047a: /* 'BrtCFRuleExt' */
          case 0x01d7: /* 'BrtCFVO' */
          case 0x041a: /* 'BrtCFVO14' */
          case 0x0289: /* 'BrtCellIgnoreEC' */
          case 0x0451: /* 'BrtCellIgnoreEC14' */
          case 0x0031: /* 'BrtCellMeta' */
          case 0x024d: /* 'BrtCellSmartTagProperty' */
          case 0x025f: /* 'BrtCellWatch' */
          case 0x0234: /* 'BrtColor' */
          case 0x041f: /* 'BrtColor14' */
          case 0x00a8: /* 'BrtColorFilter' */
          case 0x00ae: /* 'BrtCustomFilter' */
          case 0x049c: /* 'BrtCustomFilter14' */
          case 0x01f3: /* 'BrtDRef' */
          case 0x01fb: /* 'BrtDXF' */
          case 0x0226: /* 'BrtDrawing' */
          case 0x00ab: /* 'BrtDynamicFilter' */
          case 0x00a7: /* 'BrtFilter' */
          case 0x0499: /* 'BrtFilter14' */
          case 0x00a9: /* 'BrtIconFilter' */
          case 0x049d: /* 'BrtIconFilter14' */
          case 0x0227: /* 'BrtLegacyDrawing' */
          case 0x0228: /* 'BrtLegacyDrawingHF' */
          case 0x0295: /* 'BrtListPart' */
          case 0x027f: /* 'BrtOleObject' */
          case 0x01de: /* 'BrtPageSetup' */
          case 0x0219: /* 'BrtPhoneticInfo' */
          case 0x01dd: /* 'BrtPrintOptions' */
          case 0x0218: /* 'BrtRangeProtection' */
          case 0x044f: /* 'BrtRangeProtection14' */
          case 0x02a8: /* 'BrtRangeProtectionIso' */
          case 0x0450: /* 'BrtRangeProtectionIso14' */
          case 0x0400: /* 'BrtRwDescent' */
          case 0x0098: /* 'BrtSel' */
          case 0x0297: /* 'BrtSheetCalcProp' */
          case 0x0217: /* 'BrtSheetProtection' */
          case 0x02a6: /* 'BrtSheetProtectionIso' */
          case 0x01f8: /* 'BrtSlc' */
          case 0x0413: /* 'BrtSparkline' */
          case 0x01ac: /* 'BrtTable' */
          case 0x00aa: /* 'BrtTop10Filter' */
          case 0x0c00: /* 'BrtUid' */
          case 0x0032: /* 'BrtValueMeta' */
          case 0x0816: /* 'BrtWebExtension' */
          case 0x0415 /* 'BrtWsFmtInfoEx14' */:
            break;

          case 0x0023 /* 'BrtFRTBegin' */:
            pass = true;
            break;
          case 0x0024 /* 'BrtFRTEnd' */:
            pass = false;
            break;
          case 0x0025 /* 'BrtACBegin' */:
            state.push(R_n);
            pass = true;
            break;
          case 0x0026 /* 'BrtACEnd' */:
            state.pop();
            pass = false;
            break;

          default:
            if ((R_n || "").indexOf("Begin") > 0) {
              /* empty */
            } else if ((R_n || "").indexOf("End") > 0) {
              /* empty */
            } else if (!pass || opts.WTF)
              throw new Error(`Unexpected record ${RT} ${R_n}`);
        }
      },
      opts
    );

    delete opts.supbooks;
    delete opts["!row"];

    if (
      !s["!ref"] &&
      (refguess.s.r < 2000000 ||
        (ref && (ref.e.r > 0 || ref.e.c > 0 || ref.s.r > 0 || ref.s.c > 0)))
    )
      s["!ref"] = encode_range(ref || refguess);
    if (opts.sheetRows && s["!ref"]) {
      const tmpref = safe_decode_range(s["!ref"]);
      if (opts.sheetRows <= +tmpref.e.r) {
        tmpref.e.r = opts.sheetRows - 1;
        if (tmpref.e.r > refguess.e.r) tmpref.e.r = refguess.e.r;
        if (tmpref.e.r < tmpref.s.r) tmpref.s.r = tmpref.e.r;
        if (tmpref.e.c > refguess.e.c) tmpref.e.c = refguess.e.c;
        if (tmpref.e.c < tmpref.s.c) tmpref.s.c = tmpref.e.c;
        s["!fullref"] = s["!ref"];
        s["!ref"] = encode_range(tmpref);
      }
    }
    if (merges.length > 0) s["!merges"] = merges;
    if (colinfo.length > 0) s["!cols"] = colinfo;
    if (rowinfo.length > 0) s["!rows"] = rowinfo;
    return s;
  }

  /* TODO: something useful -- this is a stub */
  function write_ws_bin_cell(ba, cell, R, C, opts, ws) {
    if (cell.v === undefined) return;
    let vv = "";
    switch (cell.t) {
      case "b":
        vv = cell.v ? "1" : "0";
        break;
      case "d": // no BrtCellDate :(
        cell = dup(cell);
        cell.z = cell.z || SSF._table[14];
        cell.v = datenum(parseDate(cell.v));
        cell.t = "n";
        break;
      /* falls through */
      case "n":
      case "e":
        vv = `${cell.v}`;
        break;
      default:
        vv = cell.v;
        break;
    }
    const o = { r: R, c: C };
    /* TODO: cell style */
    o.s = get_cell_style(opts.cellXfs, cell, opts);
    if (cell.l) ws["!links"].push([encode_cell(o), cell.l]);
    if (cell.c) ws["!comments"].push([encode_cell(o), cell.c]);
    switch (cell.t) {
      case "s":
      case "str":
        if (opts.bookSST) {
          vv = get_sst_id(opts.Strings, cell.v, opts.revStrings);
          o.t = "s";
          o.v = vv;
          write_record(ba, "BrtCellIsst", write_BrtCellIsst(cell, o));
        } else {
          o.t = "str";
          write_record(ba, "BrtCellSt", write_BrtCellSt(cell, o));
        }
        return;
      case "n":
        /* TODO: determine threshold for Real vs RK */
        if (cell.v == (cell.v | 0) && cell.v > -1000 && cell.v < 1000)
          write_record(ba, "BrtCellRk", write_BrtCellRk(cell, o));
        else write_record(ba, "BrtCellReal", write_BrtCellReal(cell, o));
        return;
      case "b":
        o.t = "b";
        write_record(ba, "BrtCellBool", write_BrtCellBool(cell, o));
        return;
      case "e":
        /* TODO: error */ o.t = "e";
        break;
    }
    write_record(ba, "BrtCellBlank", write_BrtCellBlank(cell, o));
  }

  RELS.CHART =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
  RELS.CHARTEX =
    "http://schemas.microsoft.com/office/2014/relationships/chartEx";

  function parse_Cache(data) {
    const col = [];
    const num = data.match(/^<c:numCache>/);
    let f;

    /* 21.2.2.150 pt CT_NumVal */
    (data.match(/<c:pt idx="(\d*)">(.*?)<\/c:pt>/gm) || []).forEach(function(
      pt
    ) {
      const q = pt.match(/<c:pt idx="(\d*?)"><c:v>(.*)<\/c:v><\/c:pt>/);
      if (!q) return;
      col[+q[1]] = num ? +q[2] : q[2];
    });

    /* 21.2.2.71 formatCode CT_Xstring */
    const nf = unescapexml(
      (data.match(/<c:formatCode>([\s\S]*?)<\/c:formatCode>/) || [
        "",
        "General"
      ])[1]
    );

    (data.match(/<c:f>(.*?)<\/c:f>/gm) || []).forEach(function(F) {
      f = F.replace(/<.*?>/g, "");
    });

    return [col, nf, f];
  }

  /* 21.2 DrawingML - Charts */
  function parse_chart(data, name, opts, rels, wb, csheet) {
    const cs = csheet || { "!type": "chart" };
    if (!data) return csheet;
    /* 21.2.2.27 chart CT_Chart */

    let C = 0;
    let R = 0;
    let col = "A";
    const refguess = { s: { r: 2000000, c: 2000000 }, e: { r: 0, c: 0 } };

    /* 21.2.2.120 numCache CT_NumData */
    (data.match(/<c:numCache>[\s\S]*?<\/c:numCache>/gm) || []).forEach(function(
      nc
    ) {
      const cache = parse_Cache(nc);
      refguess.s.r = refguess.s.c = 0;
      refguess.e.c = C;
      col = encode_col(C);
      cache[0].forEach(function(n, i) {
        cs[col + encode_row(i)] = { t: "n", v: n, z: cache[1] };
        R = i;
      });
      if (refguess.e.r < R) refguess.e.r = R;
      ++C;
    });
    if (C > 0) cs["!ref"] = encode_range(refguess);
    return cs;
  }
  RELS.CS =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";

  /* 18.3 Worksheets also covers Chartsheets */
  function parse_cs_xml(data, opts, idx, rels, wb) {
    if (!data) return data;
    /* 18.3.1.12 chartsheet CT_ChartSheet */
    if (!rels) rels = { "!id": {} };
    const s = { "!type": "chart", "!drawel": null, "!rel": "" };
    let m;

    /* 18.3.1.83 sheetPr CT_ChartsheetPr */
    const sheetPr = data.match(sheetprregex);
    if (sheetPr) parse_ws_xml_sheetpr(sheetPr[0], s, wb, idx);

    /* 18.3.1.36 drawing CT_Drawing */
    if ((m = data.match(/drawing r:id="(.*?)"/))) s["!rel"] = m[1];

    if (rels["!id"][s["!rel"]]) s["!drawel"] = rels["!id"][s["!rel"]];
    return s;
  }

  /* [MS-XLSB] 2.4.331 BrtCsProp */
  function parse_BrtCsProp(data, length) {
    data.l += 10;
    const name = parse_XLWideString(data, length - 10);
    return { name };
  }

  /* [MS-XLSB] 2.1.7.7 Chart Sheet */
  function parse_cs_bin(data, opts, idx, rels, wb) {
    if (!data) return data;
    if (!rels) rels = { "!id": {} };
    const s = { "!type": "chart", "!drawel": null, "!rel": "" };
    const state = [];
    let pass = false;
    recordhopper(
      data,
      function cs_parse(val, R_n, RT) {
        switch (RT) {
          case 0x0226 /* 'BrtDrawing' */:
            s["!rel"] = val;
            break;

          case 0x028b /* 'BrtCsProp' */:
            if (!wb.Sheets[idx]) wb.Sheets[idx] = {};
            if (val.name) wb.Sheets[idx].CodeName = val.name;
            break;

          case 0x0232: /* 'BrtBkHim' */
          case 0x028c: /* 'BrtCsPageSetup' */
          case 0x029d: /* 'BrtCsProtection' */
          case 0x02a7: /* 'BrtCsProtectionIso' */
          case 0x0227: /* 'BrtLegacyDrawing' */
          case 0x0228: /* 'BrtLegacyDrawingHF' */
          case 0x01dc: /* 'BrtMargins' */
          case 0x0c00 /* 'BrtUid' */:
            break;

          case 0x0023 /* 'BrtFRTBegin' */:
            pass = true;
            break;
          case 0x0024 /* 'BrtFRTEnd' */:
            pass = false;
            break;
          case 0x0025 /* 'BrtACBegin' */:
            state.push(R_n);
            break;
          case 0x0026 /* 'BrtACEnd' */:
            state.pop();
            break;

          default:
            if ((R_n || "").indexOf("Begin") > 0) state.push(R_n);
            else if ((R_n || "").indexOf("End") > 0) state.pop();
            else if (!pass || opts.WTF)
              throw new Error(`Unexpected record ${RT} ${R_n}`);
        }
      },
      opts
    );

    if (rels["!id"][s["!rel"]]) s["!drawel"] = rels["!id"][s["!rel"]];
    return s;
  }
  /* 18.2.28 (CT_WorkbookProtection) Defaults */
  const WBPropsDef = [
    ["allowRefreshQuery", false, "bool"],
    ["autoCompressPictures", true, "bool"],
    ["backupFile", false, "bool"],
    ["checkCompatibility", false, "bool"],
    ["CodeName", ""],
    ["date1904", false, "bool"],
    ["defaultThemeVersion", 0, "int"],
    ["filterPrivacy", false, "bool"],
    ["hidePivotFieldList", false, "bool"],
    ["promptedSolutions", false, "bool"],
    ["publishItems", false, "bool"],
    ["refreshAllConnections", false, "bool"],
    ["saveExternalLinkValues", true, "bool"],
    ["showBorderUnselectedTables", true, "bool"],
    ["showInkAnnotation", true, "bool"],
    ["showObjects", "all"],
    ["showPivotChartFilter", false, "bool"],
    ["updateLinks", "userSet"]
  ];

  /* 18.2.30 (CT_BookView) Defaults */
  const WBViewDef = [
    ["activeTab", 0, "int"],
    ["autoFilterDateGrouping", true, "bool"],
    ["firstSheet", 0, "int"],
    ["minimized", false, "bool"],
    ["showHorizontalScroll", true, "bool"],
    ["showSheetTabs", true, "bool"],
    ["showVerticalScroll", true, "bool"],
    ["tabRatio", 600, "int"],
    ["visibility", "visible"]
    // window{Height,Width}, {x,y}Window
  ];

  /* 18.2.19 (CT_Sheet) Defaults */
  const SheetDef = [
    // ['state', 'visible']
  ];

  /* 18.2.2  (CT_CalcPr) Defaults */
  const CalcPrDef = [
    ["calcCompleted", "true"],
    ["calcMode", "auto"],
    ["calcOnSave", "true"],
    ["concurrentCalc", "true"],
    ["fullCalcOnLoad", "false"],
    ["fullPrecision", "true"],
    ["iterate", "false"],
    ["iterateCount", "100"],
    ["iterateDelta", "0.001"],
    ["refMode", "A1"]
  ];

  /* 18.2.3 (CT_CustomWorkbookView) Defaults */
  /* var CustomWBViewDef = [
	['autoUpdate', 'false'],
	['changesSavedWin', 'false'],
	['includeHiddenRowCol', 'true'],
	['includePrintSettings', 'true'],
	['maximized', 'false'],
	['minimized', 'false'],
	['onlySync', 'false'],
	['personalView', 'false'],
	['showComments', 'commIndicator'],
	['showFormulaBar', 'true'],
	['showHorizontalScroll', 'true'],
	['showObjects', 'all'],
	['showSheetTabs', 'true'],
	['showStatusbar', 'true'],
	['showVerticalScroll', 'true'],
	['tabRatio', '600'],
	['xWindow', '0'],
	['yWindow', '0']
]; */

  function push_defaults_array(target, defaults) {
    for (let j = 0; j != target.length; ++j) {
      const w = target[j];
      for (let i = 0; i != defaults.length; ++i) {
        const z = defaults[i];
        if (w[z[0]] == null) w[z[0]] = z[1];
        else
          switch (z[2]) {
            case "bool":
              if (typeof w[z[0]] === "string") w[z[0]] = parsexmlbool(w[z[0]]);
              break;
            case "int":
              if (typeof w[z[0]] === "string") w[z[0]] = parseInt(w[z[0]], 10);
              break;
          }
      }
    }
  }
  function push_defaults(target, defaults) {
    for (let i = 0; i != defaults.length; ++i) {
      const z = defaults[i];
      if (target[z[0]] == null) target[z[0]] = z[1];
      else
        switch (z[2]) {
          case "bool":
            if (typeof target[z[0]] === "string")
              target[z[0]] = parsexmlbool(target[z[0]]);
            break;
          case "int":
            if (typeof target[z[0]] === "string")
              target[z[0]] = parseInt(target[z[0]], 10);
            break;
        }
    }
  }

  function parse_wb_defaults(wb) {
    push_defaults(wb.WBProps, WBPropsDef);
    push_defaults(wb.CalcPr, CalcPrDef);

    push_defaults_array(wb.WBView, WBViewDef);
    push_defaults_array(wb.Sheets, SheetDef);

    _ssfopts.date1904 = parsexmlbool(wb.WBProps.date1904);
  }

  function safe1904(wb) {
    /* TODO: store date1904 somewhere else */
    if (!wb.Workbook) return "false";
    if (!wb.Workbook.WBProps) return "false";
    return parsexmlbool(wb.Workbook.WBProps.date1904) ? "true" : "false";
  }

  const badchars = "][*?/\\".split("");
  /* 18.2 Workbook */
  const wbnsregex = /<\w+:workbook/;
  function parse_wb_xml(data, opts) {
    if (!data) throw new Error("Could not find file");
    const wb = {
      AppVersion: {},
      WBProps: {},
      WBView: [],
      Sheets: [],
      CalcPr: {},
      Names: [],
      xmlns: ""
    };
    let pass = false;
    let xmlns = "xmlns";
    let dname = {};
    let dnstart = 0;
    data.replace(tagregex, function xml_wb(x, idx) {
      const y = parsexmltag(x);
      switch (strip_ns(y[0])) {
        case "<?xml":
          break;

        /* 18.2.27 workbook CT_Workbook 1 */
        case "<workbook":
          if (x.match(wbnsregex)) xmlns = `xmlns${x.match(/<(\w+):/)[1]}`;
          wb.xmlns = y[xmlns];
          break;
        case "</workbook>":
          break;

        /* 18.2.13 fileVersion CT_FileVersion ? */
        case "<fileVersion":
          delete y[0];
          wb.AppVersion = y;
          break;
        case "<fileVersion/>":
        case "</fileVersion>":
          break;

        /* 18.2.12 fileSharing CT_FileSharing ? */
        case "<fileSharing":
          break;
        case "<fileSharing/>":
          break;

        /* 18.2.28 workbookPr CT_WorkbookPr ? */
        case "<workbookPr":
        case "<workbookPr/>":
          WBPropsDef.forEach(function(w) {
            if (y[w[0]] == null) return;
            switch (w[2]) {
              case "bool":
                wb.WBProps[w[0]] = parsexmlbool(y[w[0]]);
                break;
              case "int":
                wb.WBProps[w[0]] = parseInt(y[w[0]], 10);
                break;
              default:
                wb.WBProps[w[0]] = y[w[0]];
            }
          });
          if (y.codeName) wb.WBProps.CodeName = utf8read(y.codeName);
          break;
        case "</workbookPr>":
          break;

        /* 18.2.29 workbookProtection CT_WorkbookProtection ? */
        case "<workbookProtection":
          break;
        case "<workbookProtection/>":
          break;

        /* 18.2.1  bookViews CT_BookViews ? */
        case "<bookViews":
        case "<bookViews>":
        case "</bookViews>":
          break;
        /* 18.2.30   workbookView CT_BookView + */
        case "<workbookView":
        case "<workbookView/>":
          delete y[0];
          wb.WBView.push(y);
          break;
        case "</workbookView>":
          break;

        /* 18.2.20 sheets CT_Sheets 1 */
        case "<sheets":
        case "<sheets>":
        case "</sheets>":
          break; // aggregate sheet
        /* 18.2.19   sheet CT_Sheet + */
        case "<sheet":
          switch (y.state) {
            case "hidden":
              y.Hidden = 1;
              break;
            case "veryHidden":
              y.Hidden = 2;
              break;
            default:
              y.Hidden = 0;
          }
          delete y.state;
          y.name = unescapexml(utf8read(y.name));
          delete y[0];
          wb.Sheets.push(y);
          break;
        case "</sheet>":
          break;

        /* 18.2.15 functionGroups CT_FunctionGroups ? */
        case "<functionGroups":
        case "<functionGroups/>":
          break;
        /* 18.2.14   functionGroup CT_FunctionGroup + */
        case "<functionGroup":
          break;

        /* 18.2.9  externalReferences CT_ExternalReferences ? */
        case "<externalReferences":
        case "</externalReferences>":
        case "<externalReferences>":
          break;
        /* 18.2.8    externalReference CT_ExternalReference + */
        case "<externalReference":
          break;

        /* 18.2.6  definedNames CT_DefinedNames ? */
        case "<definedNames/>":
          break;
        case "<definedNames>":
        case "<definedNames":
          pass = true;
          break;
        case "</definedNames>":
          pass = false;
          break;
        /* 18.2.5    definedName CT_DefinedName + */
        case "<definedName":
          {
            dname = {};
            dname.Name = utf8read(y.name);
            if (y.comment) dname.Comment = y.comment;
            if (y.localSheetId) dname.Sheet = +y.localSheetId;
            if (parsexmlbool(y.hidden || "0")) dname.Hidden = true;
            dnstart = idx + x.length;
          }
          break;
        case "</definedName>":
          {
            dname.Ref = unescapexml(utf8read(data.slice(dnstart, idx)));
            wb.Names.push(dname);
          }
          break;
        case "<definedName/>":
          break;

        /* 18.2.2  calcPr CT_CalcPr ? */
        case "<calcPr":
          delete y[0];
          wb.CalcPr = y;
          break;
        case "<calcPr/>":
          delete y[0];
          wb.CalcPr = y;
          break;
        case "</calcPr>":
          break;

        /* 18.2.16 oleSize CT_OleSize ? (ref required) */
        case "<oleSize":
          break;

        /* 18.2.4  customWorkbookViews CT_CustomWorkbookViews ? */
        case "<customWorkbookViews>":
        case "</customWorkbookViews>":
        case "<customWorkbookViews":
          break;
        /* 18.2.3  customWorkbookView CT_CustomWorkbookView + */
        case "<customWorkbookView":
        case "</customWorkbookView>":
          break;

        /* 18.2.18 pivotCaches CT_PivotCaches ? */
        case "<pivotCaches>":
        case "</pivotCaches>":
        case "<pivotCaches":
          break;
        /* 18.2.17 pivotCache CT_PivotCache ? */
        case "<pivotCache":
          break;

        /* 18.2.21 smartTagPr CT_SmartTagPr ? */
        case "<smartTagPr":
        case "<smartTagPr/>":
          break;

        /* 18.2.23 smartTagTypes CT_SmartTagTypes ? */
        case "<smartTagTypes":
        case "<smartTagTypes>":
        case "</smartTagTypes>":
          break;
        /* 18.2.22 smartTagType CT_SmartTagType ? */
        case "<smartTagType":
          break;

        /* 18.2.24 webPublishing CT_WebPublishing ? */
        case "<webPublishing":
        case "<webPublishing/>":
          break;

        /* 18.2.11 fileRecoveryPr CT_FileRecoveryPr ? */
        case "<fileRecoveryPr":
        case "<fileRecoveryPr/>":
          break;

        /* 18.2.26 webPublishObjects CT_WebPublishObjects ? */
        case "<webPublishObjects>":
        case "<webPublishObjects":
        case "</webPublishObjects>":
          break;
        /* 18.2.25 webPublishObject CT_WebPublishObject ? */
        case "<webPublishObject":
          break;

        /* 18.2.10 extLst CT_ExtensionList ? */
        case "<extLst":
        case "<extLst>":
        case "</extLst>":
        case "<extLst/>":
          break;
        /* 18.2.7  ext CT_Extension + */
        case "<ext":
          pass = true;
          break; // TODO: check with versions of excel
        case "</ext>":
          pass = false;
          break;

        /* Others */
        case "<ArchID":
          break;
        case "<AlternateContent":
        case "<AlternateContent>":
          pass = true;
          break;
        case "</AlternateContent>":
          pass = false;
          break;

        /* TODO */
        case "<revisionPtr":
          break;

        default:
          if (!pass && opts.WTF)
            throw new Error(`unrecognized ${y[0]} in workbook`);
      }
      return x;
    });
    if (XMLNS.main.indexOf(wb.xmlns) === -1)
      throw new Error(`Unknown Namespace: ${wb.xmlns}`);

    parse_wb_defaults(wb);

    return wb;
  }

  /* [MS-XLSB] 2.4.304 BrtBundleSh */
  function parse_BrtBundleSh(data, length) {
    const z = {};
    z.Hidden = data.read_shift(4); // hsState ST_SheetState
    z.iTabID = data.read_shift(4);
    z.strRelID = parse_RelID(data, length - 8);
    z.name = parse_XLWideString(data);
    return z;
  }
  function write_BrtBundleSh(data, o) {
    if (!o) o = new_buf(127);
    o.write_shift(4, data.Hidden);
    o.write_shift(4, data.iTabID);
    write_RelID(data.strRelID, o);
    write_XLWideString(data.name.slice(0, 31), o);
    return o.length > o.l ? o.slice(0, o.l) : o;
  }

  /* [MS-XLSB] 2.4.815 BrtWbProp */
  function parse_BrtWbProp(data, length) {
    const o = {};
    const flags = data.read_shift(4);
    o.defaultThemeVersion = data.read_shift(4);
    const strName = length > 8 ? parse_XLWideString(data) : "";
    if (strName.length > 0) o.CodeName = strName;
    o.autoCompressPictures = !!(flags & 0x10000);
    o.backupFile = !!(flags & 0x40);
    o.checkCompatibility = !!(flags & 0x1000);
    o.date1904 = !!(flags & 0x01);
    o.filterPrivacy = !!(flags & 0x08);
    o.hidePivotFieldList = !!(flags & 0x400);
    o.promptedSolutions = !!(flags & 0x10);
    o.publishItems = !!(flags & 0x800);
    o.refreshAllConnections = !!(flags & 0x40000);
    o.saveExternalLinkValues = !!(flags & 0x80);
    o.showBorderUnselectedTables = !!(flags & 0x04);
    o.showInkAnnotation = !!(flags & 0x20);
    o.showObjects = ["all", "placeholders", "none"][(flags >> 13) & 0x03];
    o.showPivotChartFilter = !!(flags & 0x8000);
    o.updateLinks = ["userSet", "never", "always"][(flags >> 8) & 0x03];
    return o;
  }

  function parse_BrtFRTArchID$(data, length) {
    const o = {};
    data.read_shift(4);
    o.ArchID = data.read_shift(4);
    data.l += length - 8;
    return o;
  }

  /* [MS-XLSB] 2.4.687 BrtName */
  function parse_BrtName(data, length, opts) {
    const end = data.l + length;
    data.l += 4; // var flags = data.read_shift(4);
    data.l += 1; // var chKey = data.read_shift(1);
    const itab = data.read_shift(4);
    const name = parse_XLNameWideString(data);
    const formula = parse_XLSBNameParsedFormula(data, 0, opts);
    const comment = parse_XLNullableWideString(data);
    // if(0 /* fProc */) {
    // unusedstring1: XLNullableWideString
    // description: XLNullableWideString
    // helpTopic: XLNullableWideString
    // unusedstring2: XLNullableWideString
    // }
    data.l = end;
    const out = { Name: name, Ptg: formula };
    if (itab < 0xfffffff) out.Sheet = itab;
    if (comment) out.Comment = comment;
    return out;
  }

  /* [MS-XLSB] 2.1.7.61 Workbook */
  function parse_wb_bin(data, opts) {
    const wb = {
      AppVersion: {},
      WBProps: {},
      WBView: [],
      Sheets: [],
      CalcPr: {},
      xmlns: ""
    };
    const state = [];
    let pass = false;

    if (!opts) opts = {};
    opts.biff = 12;

    const Names = [];
    const supbooks = [[]];
    supbooks.SheetNames = [];
    supbooks.XTI = [];

    recordhopper(
      data,
      function hopper_wb(val, R_n, RT) {
        switch (RT) {
          case 0x009c /* 'BrtBundleSh' */:
            supbooks.SheetNames.push(val.name);
            wb.Sheets.push(val);
            break;

          case 0x0099 /* 'BrtWbProp' */:
            wb.WBProps = val;
            break;

          case 0x0027 /* 'BrtName' */:
            if (val.Sheet != null) opts.SID = val.Sheet;
            val.Ref = stringify_formula(val.Ptg, null, null, supbooks, opts);
            delete opts.SID;
            delete val.Ptg;
            Names.push(val);
            break;
          case 0x040c:
            /* 'BrtNameExt' */ break;

          case 0x0165: /* 'BrtSupSelf' */
          case 0x0166: /* 'BrtSupSame' */
          case 0x0163: /* 'BrtSupBookSrc' */
          case 0x029b /* 'BrtSupAddin' */:
            if (!supbooks[0].length) supbooks[0] = [RT, val];
            else supbooks.push([RT, val]);
            supbooks[supbooks.length - 1].XTI = [];
            break;
          case 0x016a /* 'BrtExternSheet' */:
            if (supbooks.length === 0) {
              supbooks[0] = [];
              supbooks[0].XTI = [];
            }
            supbooks[supbooks.length - 1].XTI = supbooks[
              supbooks.length - 1
            ].XTI.concat(val);
            supbooks.XTI = supbooks.XTI.concat(val);
            break;
          case 0x0169 /* 'BrtPlaceholderName' */:
            break;

          /* case 'BrtModelTimeGroupingCalcCol' */
          case 0x0c00: /* 'BrtUid' */
          case 0x0c01: /* 'BrtRevisionPtr' */
          case 0x0817: /* 'BrtAbsPath15' */
          case 0x0216: /* 'BrtBookProtection' */
          case 0x02a5: /* 'BrtBookProtectionIso' */
          case 0x009e: /* 'BrtBookView' */
          case 0x009d: /* 'BrtCalcProp' */
          case 0x0262: /* 'BrtCrashRecErr' */
          case 0x0802: /* 'BrtDecoupledPivotCacheID' */
          case 0x009b: /* 'BrtFileRecover' */
          case 0x0224: /* 'BrtFileSharing' */
          case 0x02a4: /* 'BrtFileSharingIso' */
          case 0x0080: /* 'BrtFileVersion' */
          case 0x0299: /* 'BrtFnGroup' */
          case 0x0850: /* 'BrtModelRelationship' */
          case 0x084d: /* 'BrtModelTable' */
          case 0x0225: /* 'BrtOleSize' */
          case 0x0805: /* 'BrtPivotTableRef' */
          case 0x0254: /* 'BrtSmartTagType' */
          case 0x081c: /* 'BrtTableSlicerCacheID' */
          case 0x081b: /* 'BrtTableSlicerCacheIDs' */
          case 0x0822: /* 'BrtTimelineCachePivotCacheID' */
          case 0x018d: /* 'BrtUserBookView' */
          case 0x009a: /* 'BrtWbFactoid' */
          case 0x045d: /* 'BrtWbProp14' */
          case 0x0229: /* 'BrtWebOpt' */
          case 0x082b /* 'BrtWorkBookPr15' */:
            break;

          case 0x0023 /* 'BrtFRTBegin' */:
            state.push(R_n);
            pass = true;
            break;
          case 0x0024 /* 'BrtFRTEnd' */:
            state.pop();
            pass = false;
            break;
          case 0x0025 /* 'BrtACBegin' */:
            state.push(R_n);
            pass = true;
            break;
          case 0x0026 /* 'BrtACEnd' */:
            state.pop();
            pass = false;
            break;

          case 0x0010:
            /* 'BrtFRTArchID$' */ break;

          default:
            if ((R_n || "").indexOf("Begin") > 0) {
              /* empty */
            } else if ((R_n || "").indexOf("End") > 0) {
              /* empty */
            } else if (
              !pass ||
              (opts.WTF &&
                state[state.length - 1] != "BrtACBegin" &&
                state[state.length - 1] != "BrtFRTBegin")
            )
              throw new Error(`Unexpected record ${RT} ${R_n}`);
        }
      },
      opts
    );

    parse_wb_defaults(wb);

    // $FlowIgnore
    wb.Names = Names;

    wb.supbooks = supbooks;
    return wb;
  }

  /* [MS-XLSB] 2.4.305 BrtCalcProp */
  /* function write_BrtCalcProp(data, o) {
	if(!o) o = new_buf(26);
	o.write_shift(4,0); // force recalc
	o.write_shift(4,1);
	o.write_shift(4,0);
	write_Xnum(0, o);
	o.write_shift(-4, 1023);
	o.write_shift(1, 0x33);
	o.write_shift(1, 0x00);
	return o;
} */

  /* [MS-XLSB] 2.4.646 BrtFileRecover */
  /* function write_BrtFileRecover(data, o) {
	if(!o) o = new_buf(1);
	o.write_shift(1,0);
	return o;
} */

  /* [MS-XLSB] 2.1.7.61 Workbook */
  function parse_wb(data, name, opts) {
    if (name.slice(-4) === ".bin") return parse_wb_bin(data, opts);
    return parse_wb_xml(data, opts);
  }

  function parse_ws(data, name, idx, opts, rels, wb, themes, styles) {
    if (name.slice(-4) === ".bin")
      return parse_ws_bin(data, opts, idx, rels, wb, themes, styles);
    return parse_ws_xml(data, opts, idx, rels, wb, themes, styles);
  }

  function parse_cs(data, name, idx, opts, rels, wb, themes, styles) {
    if (name.slice(-4) === ".bin")
      return parse_cs_bin(data, opts, idx, rels, wb, themes, styles);
    return parse_cs_xml(data, opts, idx, rels, wb, themes, styles);
  }

  function parse_ms(data, name, idx, opts, rels, wb, themes, styles) {
    if (name.slice(-4) === ".bin")
      return parse_ms_bin(data, opts, idx, rels, wb, themes, styles);
    return parse_ms_xml(data, opts, idx, rels, wb, themes, styles);
  }

  function parse_ds(data, name, idx, opts, rels, wb, themes, styles) {
    if (name.slice(-4) === ".bin")
      return parse_ds_bin(data, opts, idx, rels, wb, themes, styles);
    return parse_ds_xml(data, opts, idx, rels, wb, themes, styles);
  }

  function parse_sty(data, name, themes, opts) {
    if (name.slice(-4) === ".bin") return parse_sty_bin(data, themes, opts);
    return parse_sty_xml(data, themes, opts);
  }

  function parse_theme(data, name, opts) {
    return parse_theme_xml(data, opts);
  }

  function parse_sst(data, name, opts) {
    if (name.slice(-4) === ".bin") return parse_sst_bin(data, opts);
    return parse_sst_xml(data, opts);
  }

  function parse_cmnt(data, name, opts) {
    if (name.slice(-4) === ".bin") return parse_comments_bin(data, opts);
    return parse_comments_xml(data, opts);
  }

  function parse_cc(data, name, opts) {
    if (name.slice(-4) === ".bin") return parse_cc_bin(data, name, opts);
    return parse_cc_xml(data, name, opts);
  }

  function parse_xlink(data, rel, name, opts) {
    if (name.slice(-4) === ".bin")
      return parse_xlink_bin(data, rel, name, opts);
    return parse_xlink_xml(data, rel, name, opts);
  }

  /*
function write_cc(data, name:string, opts) {
	return (name.slice(-4)===".bin" ? write_cc_bin : write_cc_xml)(data, opts);
}
*/
  const attregexg2 = /([\w:]+)=((?:")([^"]*)(?:")|(?:')([^']*)(?:'))/g;
  const attregex2 = /([\w:]+)=((?:")(?:[^"]*)(?:")|(?:')(?:[^']*)(?:'))/;
  function xlml_parsexmltag(tag, skip_root) {
    const words = tag.split(/\s+/);
    const z = [];
    if (!skip_root) z[0] = words[0];
    if (words.length === 1) return z;
    const m = tag.match(attregexg2);
    let y;
    let j;
    let w;
    let i;
    if (m)
      for (i = 0; i != m.length; ++i) {
        y = m[i].match(attregex2);
        if ((j = y[1].indexOf(":")) === -1)
          z[y[1]] = y[2].slice(1, y[2].length - 1);
        else {
          if (y[1].slice(0, 6) === "xmlns:") w = `xmlns${y[1].slice(6)}`;
          else w = y[1].slice(j + 1);
          z[w] = y[2].slice(1, y[2].length - 1);
        }
      }
    return z;
  }
  function xlml_parsexmltagobj(tag) {
    const words = tag.split(/\s+/);
    const z = {};
    if (words.length === 1) return z;
    const m = tag.match(attregexg2);
    let y;
    let j;
    let w;
    let i;
    if (m)
      for (i = 0; i != m.length; ++i) {
        y = m[i].match(attregex2);
        if ((j = y[1].indexOf(":")) === -1)
          z[y[1]] = y[2].slice(1, y[2].length - 1);
        else {
          if (y[1].slice(0, 6) === "xmlns:") w = `xmlns${y[1].slice(6)}`;
          else w = y[1].slice(j + 1);
          z[w] = y[2].slice(1, y[2].length - 1);
        }
      }
    return z;
  }

  // ----

  function xlml_format(format, value) {
    const fmt = XLMLFormatMap[format] || unescapexml(format);
    if (fmt === "General") return SSF._general(value);
    return SSF.format(fmt, value);
  }

  function xlml_set_custprop(Custprops, key, cp, val) {
    let oval = val;
    switch ((cp[0].match(/dt:dt="([\w.]+)"/) || ["", ""])[1]) {
      case "boolean":
        oval = parsexmlbool(val);
        break;
      case "i2":
      case "int":
        oval = parseInt(val, 10);
        break;
      case "r4":
      case "float":
        oval = parseFloat(val);
        break;
      case "date":
      case "dateTime.tz":
        oval = parseDate(val);
        break;
      case "i8":
      case "string":
      case "fixed":
      case "uuid":
      case "bin.base64":
        break;
      default:
        throw new Error(`bad custprop:${cp[0]}`);
    }
    Custprops[unescapexml(key)] = oval;
  }

  function safe_format_xlml(cell, nf, o) {
    if (cell.t === "z") return;
    if (!o || o.cellText !== false)
      try {
        if (cell.t === "e") {
          cell.w = cell.w || BErr[cell.v];
        } else if (nf === "General") {
          if (cell.t === "n") {
            if ((cell.v | 0) === cell.v) cell.w = SSF._general_int(cell.v);
            else cell.w = SSF._general_num(cell.v);
          } else cell.w = SSF._general(cell.v);
        } else cell.w = xlml_format(nf || "General", cell.v);
      } catch (e) {
        if (o.WTF) throw e;
      }
    try {
      const z = XLMLFormatMap[nf] || nf || "General";
      if (o.cellNF) cell.z = z;
      if (o.cellDates && cell.t == "n" && SSF.is_date(z)) {
        const _d = SSF.parse_date_code(cell.v);
        if (_d) {
          cell.t = "d";
          cell.v = new Date(_d.y, _d.m - 1, _d.d, _d.H, _d.M, _d.S, _d.u);
        }
      }
    } catch (e) {
      if (o.WTF) throw e;
    }
  }

  function process_style_xlml(styles, stag, opts) {
    if (opts.cellStyles) {
      if (stag.Interior) {
        const I = stag.Interior;
        if (I.Pattern)
          I.patternType = XLMLPatternTypeMap[I.Pattern] || I.Pattern;
      }
    }
    styles[stag.ID] = stag;
  }

  /* TODO: there must exist some form of OSP-blessed spec */
  function parse_xlml_data(
    xml,
    ss,
    data,
    cell,
    base,
    styles,
    csty,
    row,
    arrayf,
    o
  ) {
    let nf = "General";
    let sid = cell.StyleID;
    const S = {};
    o = o || {};
    const interiors = [];
    let i = 0;
    if (sid === undefined && row) sid = row.StyleID;
    if (sid === undefined && csty) sid = csty.StyleID;
    while (styles[sid] !== undefined) {
      if (styles[sid].nf) nf = styles[sid].nf;
      if (styles[sid].Interior) interiors.push(styles[sid].Interior);
      if (!styles[sid].Parent) break;
      sid = styles[sid].Parent;
    }
    switch (data.Type) {
      case "Boolean":
        cell.t = "b";
        cell.v = parsexmlbool(xml);
        break;
      case "String":
        cell.t = "s";
        cell.r = xlml_fixstr(unescapexml(xml));
        cell.v =
          xml.indexOf("<") > -1
            ? unescapexml(ss || xml).replace(/<.*?>/g, "")
            : cell.r; // todo: BR etc
        break;
      case "DateTime":
        if (xml.slice(-1) != "Z") xml += "Z";
        cell.v =
          (parseDate(xml) - new Date(Date.UTC(1899, 11, 30))) /
          (24 * 60 * 60 * 1000);
        if (cell.v !== cell.v) cell.v = unescapexml(xml);
        else if (cell.v < 60) cell.v -= 1;
        if (!nf || nf == "General") nf = "yyyy-mm-dd";
      /* falls through */
      case "Number":
        if (cell.v === undefined) cell.v = +xml;
        if (!cell.t) cell.t = "n";
        break;
      case "Error":
        cell.t = "e";
        cell.v = RBErr[xml];
        if (o.cellText !== false) cell.w = xml;
        break;
      default:
        if (xml == "" && ss == "") {
          cell.t = "z";
        } else {
          cell.t = "s";
          cell.v = xlml_fixstr(ss || xml);
        }
        break;
    }
    safe_format_xlml(cell, nf, o);
    if (o.cellFormula !== false) {
      if (cell.Formula) {
        let fstr = unescapexml(cell.Formula);
        /* strictly speaking, the leading = is required but some writers omit */
        if (fstr.charCodeAt(0) == 61 /* = */) fstr = fstr.slice(1);
        cell.f = rc_to_a1(fstr, base);
        delete cell.Formula;
        if (cell.ArrayRange == "RC") cell.F = rc_to_a1("RC:RC", base);
        else if (cell.ArrayRange) {
          cell.F = rc_to_a1(cell.ArrayRange, base);
          arrayf.push([safe_decode_range(cell.F), cell.F]);
        }
      } else {
        for (i = 0; i < arrayf.length; ++i)
          if (base.r >= arrayf[i][0].s.r && base.r <= arrayf[i][0].e.r)
            if (base.c >= arrayf[i][0].s.c && base.c <= arrayf[i][0].e.c)
              cell.F = arrayf[i][1];
      }
    }
    if (o.cellStyles) {
      interiors.forEach(function(x) {
        if (!S.patternType && x.patternType) S.patternType = x.patternType;
      });
      cell.s = S;
    }
    if (cell.StyleID !== undefined) cell.ixfe = cell.StyleID;
  }

  function xlml_clean_comment(comment) {
    comment.t = comment.v || "";
    comment.t = comment.t.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    comment.v = comment.w = comment.ixfe = undefined;
  }

  function xlml_normalize(d) {
    if (has_buf && Buffer.isBuffer(d)) return d.toString("utf8");
    if (typeof d === "string") return d;
    /* duktape */
    if (typeof Uint8Array !== "undefined" && d instanceof Uint8Array)
      return utf8read(a2s(ab2a(d)));
    throw new Error("Bad input format: expected Buffer or string");
  }

  /* TODO: Everything */
  /* UOS uses CJK in tags */
  var xlmlregex = /<(\/?)([^\s?><!\/:]*:|)([^\s?<>:\/]+)(?:[\s?:\/][^>]*)?>/gm;

  function parse_xlml_xml(d, _opts) {
    const opts = _opts || {};
    make_ssf(SSF);
    let str = debom(xlml_normalize(d));
    if (
      opts.type == "binary" ||
      opts.type == "array" ||
      opts.type == "base64"
    ) {
      if (typeof cptable !== "undefined")
        str = cptable.utils.decode(65001, char_codes(str));
      else str = utf8read(str);
    }
    const opening = str.slice(0, 1024).toLowerCase();
    let ishtml = false;
    if (opening.indexOf("<?xml") == -1)
      ["html", "table", "head", "meta", "script", "style", "div"].forEach(
        function(tag) {
          if (opening.indexOf(`<${tag}`) >= 0) ishtml = true;
        }
      );
    if (ishtml) return HTML_.to_workbook(str, opts);
    let Rn;
    const state = [];
    let tmp;
    if (DENSE != null && opts.dense == null) opts.dense = DENSE;
    const sheets = {};
    const sheetnames = [];
    let cursheet = opts.dense ? [] : {};
    let sheetname = "";
    let table = {};
    let cell = {};
    let row = {}; // eslint-disable-line no-unused-vars
    let dtag = xlml_parsexmltag('<Data ss:Type="String">');
    let didx = 0;
    let c = 0;
    let r = 0;
    let refguess = { s: { r: 2000000, c: 2000000 }, e: { r: 0, c: 0 } };
    const styles = {};
    let stag = {};
    let ss = "";
    let fidx = 0;
    let merges = [];
    const Props = {};
    const Custprops = {};
    let pidx = 0;
    let cp = [];
    let comments = [];
    let comment = {};
    let cstys = [];
    let csty;
    let seencol = false;
    let arrayf = [];
    let rowinfo = [];
    let rowobj = {};
    let cc = 0;
    let rr = 0;
    const Workbook = { Sheets: [], WBProps: { date1904: false } };
    let wsprops = {};
    xlmlregex.lastIndex = 0;
    str = str.replace(/<!--([\s\S]*?)-->/gm, "");
    let raw_Rn3 = "";
    while ((Rn = xlmlregex.exec(str)))
      switch ((Rn[3] = (raw_Rn3 = Rn[3]).toLowerCase())) {
        case "data" /* case 'Data' */:
          if (raw_Rn3 == "data") {
            if (Rn[1] === "/") {
              if ((tmp = state.pop())[0] !== Rn[3])
                throw new Error(`Bad state: ${tmp.join("|")}`);
            } else if (Rn[0].charAt(Rn[0].length - 2) !== "/")
              state.push([Rn[3], true]);
            break;
          }
          if (state[state.length - 1][1]) break;
          if (Rn[1] === "/")
            parse_xlml_data(
              str.slice(didx, Rn.index),
              ss,
              dtag,
              state[state.length - 1][0] == /* "Comment" */ "comment"
                ? comment
                : cell,
              { c, r },
              styles,
              cstys[c],
              row,
              arrayf,
              opts
            );
          else {
            ss = "";
            dtag = xlml_parsexmltag(Rn[0]);
            didx = Rn.index + Rn[0].length;
          }
          break;
        case "cell" /* case 'Cell' */:
          if (Rn[1] === "/") {
            if (comments.length > 0) cell.c = comments;
            if (
              (!opts.sheetRows || opts.sheetRows > r) &&
              cell.v !== undefined
            ) {
              if (opts.dense) {
                if (!cursheet[r]) cursheet[r] = [];
                cursheet[r][c] = cell;
              } else cursheet[encode_col(c) + encode_row(r)] = cell;
            }
            if (cell.HRef) {
              cell.l = { Target: cell.HRef };
              if (cell.HRefScreenTip) cell.l.Tooltip = cell.HRefScreenTip;
              delete cell.HRef;
              delete cell.HRefScreenTip;
            }
            if (cell.MergeAcross || cell.MergeDown) {
              cc = c + (parseInt(cell.MergeAcross, 10) | 0);
              rr = r + (parseInt(cell.MergeDown, 10) | 0);
              merges.push({ s: { c, r }, e: { c: cc, r: rr } });
            }
            if (!opts.sheetStubs) {
              if (cell.MergeAcross) c = cc + 1;
              else ++c;
            } else if (cell.MergeAcross || cell.MergeDown) {
              for (let cma = c; cma <= cc; ++cma) {
                for (let cmd = r; cmd <= rr; ++cmd) {
                  if (cma > c || cmd > r) {
                    if (opts.dense) {
                      if (!cursheet[cmd]) cursheet[cmd] = [];
                      cursheet[cmd][cma] = { t: "z" };
                    } else
                      cursheet[encode_col(cma) + encode_row(cmd)] = { t: "z" };
                  }
                }
              }
              c = cc + 1;
            } else ++c;
          } else {
            cell = xlml_parsexmltagobj(Rn[0]);
            if (cell.Index) c = +cell.Index - 1;
            if (c < refguess.s.c) refguess.s.c = c;
            if (c > refguess.e.c) refguess.e.c = c;
            if (Rn[0].slice(-2) === "/>") ++c;
            comments = [];
          }
          break;
        case "row" /* case 'Row' */:
          if (Rn[1] === "/" || Rn[0].slice(-2) === "/>") {
            if (r < refguess.s.r) refguess.s.r = r;
            if (r > refguess.e.r) refguess.e.r = r;
            if (Rn[0].slice(-2) === "/>") {
              row = xlml_parsexmltag(Rn[0]);
              if (row.Index) r = +row.Index - 1;
            }
            c = 0;
            ++r;
          } else {
            row = xlml_parsexmltag(Rn[0]);
            if (row.Index) r = +row.Index - 1;
            rowobj = {};
            if (row.AutoFitHeight == "0" || row.Height) {
              rowobj.hpx = parseInt(row.Height, 10);
              rowobj.hpt = px2pt(rowobj.hpx);
              rowinfo[r] = rowobj;
            }
            if (row.Hidden == "1") {
              rowobj.hidden = true;
              rowinfo[r] = rowobj;
            }
          }
          break;
        case "worksheet" /* TODO: read range from FullRows/FullColumns */ /* case 'Worksheet' */:
          if (Rn[1] === "/") {
            if ((tmp = state.pop())[0] !== Rn[3])
              throw new Error(`Bad state: ${tmp.join("|")}`);
            sheetnames.push(sheetname);
            if (refguess.s.r <= refguess.e.r && refguess.s.c <= refguess.e.c) {
              cursheet["!ref"] = encode_range(refguess);
              if (opts.sheetRows && opts.sheetRows <= refguess.e.r) {
                cursheet["!fullref"] = cursheet["!ref"];
                refguess.e.r = opts.sheetRows - 1;
                cursheet["!ref"] = encode_range(refguess);
              }
            }
            if (merges.length) cursheet["!merges"] = merges;
            if (cstys.length > 0) cursheet["!cols"] = cstys;
            if (rowinfo.length > 0) cursheet["!rows"] = rowinfo;
            sheets[sheetname] = cursheet;
          } else {
            refguess = { s: { r: 2000000, c: 2000000 }, e: { r: 0, c: 0 } };
            r = c = 0;
            state.push([Rn[3], false]);
            tmp = xlml_parsexmltag(Rn[0]);
            sheetname = unescapexml(tmp.Name);
            cursheet = opts.dense ? [] : {};
            merges = [];
            arrayf = [];
            rowinfo = [];
            wsprops = { name: sheetname, Hidden: 0 };
            Workbook.Sheets.push(wsprops);
          }
          break;
        case "table" /* case 'Table' */:
          if (Rn[1] === "/") {
            if ((tmp = state.pop())[0] !== Rn[3])
              throw new Error(`Bad state: ${tmp.join("|")}`);
          } else if (Rn[0].slice(-2) == "/>") break;
          else {
            table = xlml_parsexmltag(Rn[0]);
            state.push([Rn[3], false]);
            cstys = [];
            seencol = false;
          }
          break;

        case "style" /* case 'Style' */:
          if (Rn[1] === "/") process_style_xlml(styles, stag, opts);
          else stag = xlml_parsexmltag(Rn[0]);
          break;

        case "numberformat" /* case 'NumberFormat' */:
          stag.nf = unescapexml(xlml_parsexmltag(Rn[0]).Format || "General");
          if (XLMLFormatMap[stag.nf]) stag.nf = XLMLFormatMap[stag.nf];
          for (var ssfidx = 0; ssfidx != 0x188; ++ssfidx)
            if (SSF._table[ssfidx] == stag.nf) break;
          if (ssfidx == 0x188)
            for (ssfidx = 0x39; ssfidx != 0x188; ++ssfidx)
              if (SSF._table[ssfidx] == null) {
                SSF.load(stag.nf, ssfidx);
                break;
              }
          break;

        case "column" /* case 'Column' */:
          if (state[state.length - 1][0] !== /* 'Table' */ "table") break;
          csty = xlml_parsexmltag(Rn[0]);
          if (csty.Hidden) {
            csty.hidden = true;
            delete csty.Hidden;
          }
          if (csty.Width) csty.wpx = parseInt(csty.Width, 10);
          if (!seencol && csty.wpx > 10) {
            seencol = true;
            MDW = DEF_MDW; // find_mdw_wpx(csty.wpx);
            for (let _col = 0; _col < cstys.length; ++_col)
              if (cstys[_col]) process_col(cstys[_col]);
          }
          if (seencol) process_col(csty);
          cstys[csty.Index - 1 || cstys.length] = csty;
          for (let i = 0; i < +csty.Span; ++i) cstys[cstys.length] = dup(csty);
          break;

        case "namedrange" /* case 'NamedRange' */:
          if (Rn[1] === "/") break;
          if (!Workbook.Names) Workbook.Names = [];
          var _NamedRange = parsexmltag(Rn[0]);
          var _DefinedName = {
            Name: _NamedRange.Name,
            Ref: rc_to_a1(_NamedRange.RefersTo.slice(1), { r: 0, c: 0 })
          };
          if (Workbook.Sheets.length > 0)
            _DefinedName.Sheet = Workbook.Sheets.length - 1;
          Workbook.Names.push(_DefinedName);
          break;

        case "namedcell" /* case 'NamedCell' */:
          break;
        case "b" /* case 'B' */:
          break;
        case "i" /* case 'I' */:
          break;
        case "u" /* case 'U' */:
          break;
        case "s" /* case 'S' */:
          break;
        case "em" /* case 'EM' */:
          break;
        case "h2" /* case 'H2' */:
          break;
        case "h3" /* case 'H3' */:
          break;
        case "sub" /* case 'Sub' */:
          break;
        case "sup" /* case 'Sup' */:
          break;
        case "span" /* case 'Span' */:
          break;
        case "alignment" /* case 'Alignment' */:
          break;
        case "borders" /* case 'Borders' */:
          break;
        case "border" /* case 'Border' */:
          break;
        case "font" /* case 'Font' */:
          if (Rn[0].slice(-2) === "/>") break;
          else if (Rn[1] === "/") ss += str.slice(fidx, Rn.index);
          else fidx = Rn.index + Rn[0].length;
          break;
        case "interior" /* case 'Interior' */:
          if (!opts.cellStyles) break;
          stag.Interior = xlml_parsexmltag(Rn[0]);
          break;
        case "protection" /* case 'Protection' */:
          break;

        case "author" /* case 'Author' */:
        case "title" /* case 'Title' */:
        case "description" /* case 'Description' */:
        case "created" /* case 'Created' */:
        case "keywords" /* case 'Keywords' */:
        case "subject" /* case 'Subject' */:
        case "category" /* case 'Category' */:
        case "company" /* case 'Company' */:
        case "lastauthor" /* case 'LastAuthor' */:
        case "lastsaved" /* case 'LastSaved' */:
        case "lastprinted" /* case 'LastPrinted' */:
        case "version" /* case 'Version' */:
        case "revision" /* case 'Revision' */:
        case "totaltime" /* case 'TotalTime' */:
        case "hyperlinkbase" /* case 'HyperlinkBase' */:
        case "manager" /* case 'Manager' */:
        case "contentstatus" /* case 'ContentStatus' */:
        case "identifier" /* case 'Identifier' */:
        case "language" /* case 'Language' */:
        case "appname" /* case 'AppName' */:
          if (Rn[0].slice(-2) === "/>") break;
          else if (Rn[1] === "/")
            xlml_set_prop(Props, raw_Rn3, str.slice(pidx, Rn.index));
          else pidx = Rn.index + Rn[0].length;
          break;
        case "paragraphs" /* case 'Paragraphs' */:
          break;

        case "styles" /* case 'Styles' */:
        case "workbook" /* case 'Workbook' */:
          if (Rn[1] === "/") {
            if ((tmp = state.pop())[0] !== Rn[3])
              throw new Error(`Bad state: ${tmp.join("|")}`);
          } else state.push([Rn[3], false]);
          break;

        case "comment" /* case 'Comment' */:
          if (Rn[1] === "/") {
            if ((tmp = state.pop())[0] !== Rn[3])
              throw new Error(`Bad state: ${tmp.join("|")}`);
            xlml_clean_comment(comment);
            comments.push(comment);
          } else {
            state.push([Rn[3], false]);
            tmp = xlml_parsexmltag(Rn[0]);
            comment = { a: tmp.Author };
          }
          break;

        case "autofilter" /* case 'AutoFilter' */:
          if (Rn[1] === "/") {
            if ((tmp = state.pop())[0] !== Rn[3])
              throw new Error(`Bad state: ${tmp.join("|")}`);
          } else if (Rn[0].charAt(Rn[0].length - 2) !== "/") {
            const AutoFilter = xlml_parsexmltag(Rn[0]);
            cursheet["!autofilter"] = {
              ref: rc_to_a1(AutoFilter.Range).replace(/\$/g, "")
            };
            state.push([Rn[3], true]);
          }
          break;

        case "name" /* case 'Name' */:
          break;

        case "datavalidation" /* case 'DataValidation' */:
          if (Rn[1] === "/") {
            if ((tmp = state.pop())[0] !== Rn[3])
              throw new Error(`Bad state: ${tmp.join("|")}`);
          } else if (Rn[0].charAt(Rn[0].length - 2) !== "/")
            state.push([Rn[3], true]);
          break;

        case "pixelsperinch" /* case 'PixelsPerInch' */:
          break;
        case "componentoptions" /* case 'ComponentOptions' */:
        case "documentproperties" /* case 'DocumentProperties' */:
        case "customdocumentproperties" /* case 'CustomDocumentProperties' */:
        case "officedocumentsettings" /* case 'OfficeDocumentSettings' */:
        case "pivottable" /* case 'PivotTable' */:
        case "pivotcache" /* case 'PivotCache' */:
        case "names" /* case 'Names' */:
        case "mapinfo" /* case 'MapInfo' */:
        case "pagebreaks" /* case 'PageBreaks' */:
        case "querytable" /* case 'QueryTable' */:
        case "sorting" /* case 'Sorting' */:
        case "schema" /* case 'Schema' */: // case 'data' /*case 'data'*/:
        case "conditionalformatting" /* case 'ConditionalFormatting' */:
        case "smarttagtype" /* case 'SmartTagType' */:
        case "smarttags" /* case 'SmartTags' */:
        case "excelworkbook" /* case 'ExcelWorkbook' */:
        case "workbookoptions" /* case 'WorkbookOptions' */:
        case "worksheetoptions" /* case 'WorksheetOptions' */:
          if (Rn[1] === "/") {
            if ((tmp = state.pop())[0] !== Rn[3])
              throw new Error(`Bad state: ${tmp.join("|")}`);
          } else if (Rn[0].charAt(Rn[0].length - 2) !== "/")
            state.push([Rn[3], true]);
          break;

        default:
          /* FODS file root is <office:document> */
          if (state.length == 0 && Rn[3] == "document")
            return parse_fods(str, opts);
          /* UOS file root is <uof:UOF> */
          if (state.length == 0 && Rn[3] == "uof" /* "UOF" */)
            return parse_fods(str, opts);

          var seen = true;
          switch (state[state.length - 1][0]) {
            /* OfficeDocumentSettings */
            case "officedocumentsettings" /* case 'OfficeDocumentSettings' */:
              switch (Rn[3]) {
                case "allowpng" /* case 'AllowPNG' */:
                  break;
                case "removepersonalinformation" /* case 'RemovePersonalInformation' */:
                  break;
                case "downloadcomponents" /* case 'DownloadComponents' */:
                  break;
                case "locationofcomponents" /* case 'LocationOfComponents' */:
                  break;
                case "colors" /* case 'Colors' */:
                  break;
                case "color" /* case 'Color' */:
                  break;
                case "index" /* case 'Index' */:
                  break;
                case "rgb" /* case 'RGB' */:
                  break;
                case "targetscreensize" /* case 'TargetScreenSize' */:
                  break;
                case "readonlyrecommended" /* case 'ReadOnlyRecommended' */:
                  break;
                default:
                  seen = false;
              }
              break;

            /* ComponentOptions */
            case "componentoptions" /* case 'ComponentOptions' */:
              switch (Rn[3]) {
                case "toolbar" /* case 'Toolbar' */:
                  break;
                case "hideofficelogo" /* case 'HideOfficeLogo' */:
                  break;
                case "spreadsheetautofit" /* case 'SpreadsheetAutoFit' */:
                  break;
                case "label" /* case 'Label' */:
                  break;
                case "caption" /* case 'Caption' */:
                  break;
                case "maxheight" /* case 'MaxHeight' */:
                  break;
                case "maxwidth" /* case 'MaxWidth' */:
                  break;
                case "nextsheetnumber" /* case 'NextSheetNumber' */:
                  break;
                default:
                  seen = false;
              }
              break;

            /* ExcelWorkbook */
            case "excelworkbook" /* case 'ExcelWorkbook' */:
              switch (Rn[3]) {
                case "date1904" /* case 'Date1904' */:
                  Workbook.WBProps.date1904 = true;
                  break;
                case "windowheight" /* case 'WindowHeight' */:
                  break;
                case "windowwidth" /* case 'WindowWidth' */:
                  break;
                case "windowtopx" /* case 'WindowTopX' */:
                  break;
                case "windowtopy" /* case 'WindowTopY' */:
                  break;
                case "tabratio" /* case 'TabRatio' */:
                  break;
                case "protectstructure" /* case 'ProtectStructure' */:
                  break;
                case "protectwindow" /* case 'ProtectWindow' */:
                  break;
                case "protectwindows" /* case 'ProtectWindows' */:
                  break;
                case "activesheet" /* case 'ActiveSheet' */:
                  break;
                case "displayinknotes" /* case 'DisplayInkNotes' */:
                  break;
                case "firstvisiblesheet" /* case 'FirstVisibleSheet' */:
                  break;
                case "supbook" /* case 'SupBook' */:
                  break;
                case "sheetname" /* case 'SheetName' */:
                  break;
                case "sheetindex" /* case 'SheetIndex' */:
                  break;
                case "sheetindexfirst" /* case 'SheetIndexFirst' */:
                  break;
                case "sheetindexlast" /* case 'SheetIndexLast' */:
                  break;
                case "dll" /* case 'Dll' */:
                  break;
                case "acceptlabelsinformulas" /* case 'AcceptLabelsInFormulas' */:
                  break;
                case "donotsavelinkvalues" /* case 'DoNotSaveLinkValues' */:
                  break;
                case "iteration" /* case 'Iteration' */:
                  break;
                case "maxiterations" /* case 'MaxIterations' */:
                  break;
                case "maxchange" /* case 'MaxChange' */:
                  break;
                case "path" /* case 'Path' */:
                  break;
                case "xct" /* case 'Xct' */:
                  break;
                case "count" /* case 'Count' */:
                  break;
                case "selectedsheets" /* case 'SelectedSheets' */:
                  break;
                case "calculation" /* case 'Calculation' */:
                  break;
                case "uncalced" /* case 'Uncalced' */:
                  break;
                case "startupprompt" /* case 'StartupPrompt' */:
                  break;
                case "crn" /* case 'Crn' */:
                  break;
                case "externname" /* case 'ExternName' */:
                  break;
                case "formula" /* case 'Formula' */:
                  break;
                case "colfirst" /* case 'ColFirst' */:
                  break;
                case "collast" /* case 'ColLast' */:
                  break;
                case "wantadvise" /* case 'WantAdvise' */:
                  break;
                case "boolean" /* case 'Boolean' */:
                  break;
                case "error" /* case 'Error' */:
                  break;
                case "text" /* case 'Text' */:
                  break;
                case "ole" /* case 'OLE' */:
                  break;
                case "noautorecover" /* case 'NoAutoRecover' */:
                  break;
                case "publishobjects" /* case 'PublishObjects' */:
                  break;
                case "donotcalculatebeforesave" /* case 'DoNotCalculateBeforeSave' */:
                  break;
                case "number" /* case 'Number' */:
                  break;
                case "refmoder1c1" /* case 'RefModeR1C1' */:
                  break;
                case "embedsavesmarttags" /* case 'EmbedSaveSmartTags' */:
                  break;
                default:
                  seen = false;
              }
              break;

            /* WorkbookOptions */
            case "workbookoptions" /* case 'WorkbookOptions' */:
              switch (Rn[3]) {
                case "owcversion" /* case 'OWCVersion' */:
                  break;
                case "height" /* case 'Height' */:
                  break;
                case "width" /* case 'Width' */:
                  break;
                default:
                  seen = false;
              }
              break;

            /* WorksheetOptions */
            case "worksheetoptions" /* case 'WorksheetOptions' */:
              switch (Rn[3]) {
                case "visible" /* case 'Visible' */:
                  if (Rn[0].slice(-2) === "/>") {
                    /* empty */
                  } else if (Rn[1] === "/")
                    switch (str.slice(pidx, Rn.index)) {
                      case "SheetHidden":
                        wsprops.Hidden = 1;
                        break;
                      case "SheetVeryHidden":
                        wsprops.Hidden = 2;
                        break;
                    }
                  else pidx = Rn.index + Rn[0].length;
                  break;
                case "header" /* case 'Header' */:
                  if (!cursheet["!margins"])
                    default_margins((cursheet["!margins"] = {}), "xlml");
                  cursheet["!margins"].header = parsexmltag(Rn[0]).Margin;
                  break;
                case "footer" /* case 'Footer' */:
                  if (!cursheet["!margins"])
                    default_margins((cursheet["!margins"] = {}), "xlml");
                  cursheet["!margins"].footer = parsexmltag(Rn[0]).Margin;
                  break;
                case "pagemargins" /* case 'PageMargins' */:
                  var pagemargins = parsexmltag(Rn[0]);
                  if (!cursheet["!margins"])
                    default_margins((cursheet["!margins"] = {}), "xlml");
                  if (pagemargins.Top)
                    cursheet["!margins"].top = pagemargins.Top;
                  if (pagemargins.Left)
                    cursheet["!margins"].left = pagemargins.Left;
                  if (pagemargins.Right)
                    cursheet["!margins"].right = pagemargins.Right;
                  if (pagemargins.Bottom)
                    cursheet["!margins"].bottom = pagemargins.Bottom;
                  break;
                case "displayrighttoleft" /* case 'DisplayRightToLeft' */:
                  if (!Workbook.Views) Workbook.Views = [];
                  if (!Workbook.Views[0]) Workbook.Views[0] = {};
                  Workbook.Views[0].RTL = true;
                  break;

                case "freezepanes" /* case 'FreezePanes' */:
                  break;
                case "frozennosplit" /* case 'FrozenNoSplit' */:
                  break;

                case "splithorizontal" /* case 'SplitHorizontal' */:
                case "splitvertical" /* case 'SplitVertical' */:
                  break;

                case "donotdisplaygridlines" /* case 'DoNotDisplayGridlines' */:
                  break;

                case "toprowbottompane" /* case 'TopRowBottomPane' */:
                  break;
                case "leftcolumnrightpane" /* case 'LeftColumnRightPane' */:
                  break;

                case "unsynced" /* case 'Unsynced' */:
                  break;
                case "print" /* case 'Print' */:
                  break;
                case "panes" /* case 'Panes' */:
                  break;
                case "scale" /* case 'Scale' */:
                  break;
                case "pane" /* case 'Pane' */:
                  break;
                case "number" /* case 'Number' */:
                  break;
                case "layout" /* case 'Layout' */:
                  break;
                case "pagesetup" /* case 'PageSetup' */:
                  break;
                case "selected" /* case 'Selected' */:
                  break;
                case "protectobjects" /* case 'ProtectObjects' */:
                  break;
                case "enableselection" /* case 'EnableSelection' */:
                  break;
                case "protectscenarios" /* case 'ProtectScenarios' */:
                  break;
                case "validprinterinfo" /* case 'ValidPrinterInfo' */:
                  break;
                case "horizontalresolution" /* case 'HorizontalResolution' */:
                  break;
                case "verticalresolution" /* case 'VerticalResolution' */:
                  break;
                case "numberofcopies" /* case 'NumberofCopies' */:
                  break;
                case "activerow" /* case 'ActiveRow' */:
                  break;
                case "activecol" /* case 'ActiveCol' */:
                  break;
                case "activepane" /* case 'ActivePane' */:
                  break;
                case "toprowvisible" /* case 'TopRowVisible' */:
                  break;
                case "leftcolumnvisible" /* case 'LeftColumnVisible' */:
                  break;
                case "fittopage" /* case 'FitToPage' */:
                  break;
                case "rangeselection" /* case 'RangeSelection' */:
                  break;
                case "papersizeindex" /* case 'PaperSizeIndex' */:
                  break;
                case "pagelayoutzoom" /* case 'PageLayoutZoom' */:
                  break;
                case "pagebreakzoom" /* case 'PageBreakZoom' */:
                  break;
                case "filteron" /* case 'FilterOn' */:
                  break;
                case "fitwidth" /* case 'FitWidth' */:
                  break;
                case "fitheight" /* case 'FitHeight' */:
                  break;
                case "commentslayout" /* case 'CommentsLayout' */:
                  break;
                case "zoom" /* case 'Zoom' */:
                  break;
                case "lefttoright" /* case 'LeftToRight' */:
                  break;
                case "gridlines" /* case 'Gridlines' */:
                  break;
                case "allowsort" /* case 'AllowSort' */:
                  break;
                case "allowfilter" /* case 'AllowFilter' */:
                  break;
                case "allowinsertrows" /* case 'AllowInsertRows' */:
                  break;
                case "allowdeleterows" /* case 'AllowDeleteRows' */:
                  break;
                case "allowinsertcols" /* case 'AllowInsertCols' */:
                  break;
                case "allowdeletecols" /* case 'AllowDeleteCols' */:
                  break;
                case "allowinserthyperlinks" /* case 'AllowInsertHyperlinks' */:
                  break;
                case "allowformatcells" /* case 'AllowFormatCells' */:
                  break;
                case "allowsizecols" /* case 'AllowSizeCols' */:
                  break;
                case "allowsizerows" /* case 'AllowSizeRows' */:
                  break;
                case "nosummaryrowsbelowdetail" /* case 'NoSummaryRowsBelowDetail' */:
                  break;
                case "tabcolorindex" /* case 'TabColorIndex' */:
                  break;
                case "donotdisplayheadings" /* case 'DoNotDisplayHeadings' */:
                  break;
                case "showpagelayoutzoom" /* case 'ShowPageLayoutZoom' */:
                  break;
                case "nosummarycolumnsrightdetail" /* case 'NoSummaryColumnsRightDetail' */:
                  break;
                case "blackandwhite" /* case 'BlackAndWhite' */:
                  break;
                case "donotdisplayzeros" /* case 'DoNotDisplayZeros' */:
                  break;
                case "displaypagebreak" /* case 'DisplayPageBreak' */:
                  break;
                case "rowcolheadings" /* case 'RowColHeadings' */:
                  break;
                case "donotdisplayoutline" /* case 'DoNotDisplayOutline' */:
                  break;
                case "noorientation" /* case 'NoOrientation' */:
                  break;
                case "allowusepivottables" /* case 'AllowUsePivotTables' */:
                  break;
                case "zeroheight" /* case 'ZeroHeight' */:
                  break;
                case "viewablerange" /* case 'ViewableRange' */:
                  break;
                case "selection" /* case 'Selection' */:
                  break;
                case "protectcontents" /* case 'ProtectContents' */:
                  break;
                default:
                  seen = false;
              }
              break;

            /* PivotTable */
            case "pivottable" /* case 'PivotTable' */:
            case "pivotcache" /* case 'PivotCache' */:
              switch (Rn[3]) {
                case "immediateitemsondrop" /* case 'ImmediateItemsOnDrop' */:
                  break;
                case "showpagemultipleitemlabel" /* case 'ShowPageMultipleItemLabel' */:
                  break;
                case "compactrowindent" /* case 'CompactRowIndent' */:
                  break;
                case "location" /* case 'Location' */:
                  break;
                case "pivotfield" /* case 'PivotField' */:
                  break;
                case "orientation" /* case 'Orientation' */:
                  break;
                case "layoutform" /* case 'LayoutForm' */:
                  break;
                case "layoutsubtotallocation" /* case 'LayoutSubtotalLocation' */:
                  break;
                case "layoutcompactrow" /* case 'LayoutCompactRow' */:
                  break;
                case "position" /* case 'Position' */:
                  break;
                case "pivotitem" /* case 'PivotItem' */:
                  break;
                case "datatype" /* case 'DataType' */:
                  break;
                case "datafield" /* case 'DataField' */:
                  break;
                case "sourcename" /* case 'SourceName' */:
                  break;
                case "parentfield" /* case 'ParentField' */:
                  break;
                case "ptlineitems" /* case 'PTLineItems' */:
                  break;
                case "ptlineitem" /* case 'PTLineItem' */:
                  break;
                case "countofsameitems" /* case 'CountOfSameItems' */:
                  break;
                case "item" /* case 'Item' */:
                  break;
                case "itemtype" /* case 'ItemType' */:
                  break;
                case "ptsource" /* case 'PTSource' */:
                  break;
                case "cacheindex" /* case 'CacheIndex' */:
                  break;
                case "consolidationreference" /* case 'ConsolidationReference' */:
                  break;
                case "filename" /* case 'FileName' */:
                  break;
                case "reference" /* case 'Reference' */:
                  break;
                case "nocolumngrand" /* case 'NoColumnGrand' */:
                  break;
                case "norowgrand" /* case 'NoRowGrand' */:
                  break;
                case "blanklineafteritems" /* case 'BlankLineAfterItems' */:
                  break;
                case "hidden" /* case 'Hidden' */:
                  break;
                case "subtotal" /* case 'Subtotal' */:
                  break;
                case "basefield" /* case 'BaseField' */:
                  break;
                case "mapchilditems" /* case 'MapChildItems' */:
                  break;
                case "function" /* case 'Function' */:
                  break;
                case "refreshonfileopen" /* case 'RefreshOnFileOpen' */:
                  break;
                case "printsettitles" /* case 'PrintSetTitles' */:
                  break;
                case "mergelabels" /* case 'MergeLabels' */:
                  break;
                case "defaultversion" /* case 'DefaultVersion' */:
                  break;
                case "refreshname" /* case 'RefreshName' */:
                  break;
                case "refreshdate" /* case 'RefreshDate' */:
                  break;
                case "refreshdatecopy" /* case 'RefreshDateCopy' */:
                  break;
                case "versionlastrefresh" /* case 'VersionLastRefresh' */:
                  break;
                case "versionlastupdate" /* case 'VersionLastUpdate' */:
                  break;
                case "versionupdateablemin" /* case 'VersionUpdateableMin' */:
                  break;
                case "versionrefreshablemin" /* case 'VersionRefreshableMin' */:
                  break;
                case "calculation" /* case 'Calculation' */:
                  break;
                default:
                  seen = false;
              }
              break;

            /* PageBreaks */
            case "pagebreaks" /* case 'PageBreaks' */:
              switch (Rn[3]) {
                case "colbreaks" /* case 'ColBreaks' */:
                  break;
                case "colbreak" /* case 'ColBreak' */:
                  break;
                case "rowbreaks" /* case 'RowBreaks' */:
                  break;
                case "rowbreak" /* case 'RowBreak' */:
                  break;
                case "colstart" /* case 'ColStart' */:
                  break;
                case "colend" /* case 'ColEnd' */:
                  break;
                case "rowend" /* case 'RowEnd' */:
                  break;
                default:
                  seen = false;
              }
              break;

            /* AutoFilter */
            case "autofilter" /* case 'AutoFilter' */:
              switch (Rn[3]) {
                case "autofiltercolumn" /* case 'AutoFilterColumn' */:
                  break;
                case "autofiltercondition" /* case 'AutoFilterCondition' */:
                  break;
                case "autofilterand" /* case 'AutoFilterAnd' */:
                  break;
                case "autofilteror" /* case 'AutoFilterOr' */:
                  break;
                default:
                  seen = false;
              }
              break;

            /* QueryTable */
            case "querytable" /* case 'QueryTable' */:
              switch (Rn[3]) {
                case "id" /* case 'Id' */:
                  break;
                case "autoformatfont" /* case 'AutoFormatFont' */:
                  break;
                case "autoformatpattern" /* case 'AutoFormatPattern' */:
                  break;
                case "querysource" /* case 'QuerySource' */:
                  break;
                case "querytype" /* case 'QueryType' */:
                  break;
                case "enableredirections" /* case 'EnableRedirections' */:
                  break;
                case "refreshedinxl9" /* case 'RefreshedInXl9' */:
                  break;
                case "urlstring" /* case 'URLString' */:
                  break;
                case "htmltables" /* case 'HTMLTables' */:
                  break;
                case "connection" /* case 'Connection' */:
                  break;
                case "commandtext" /* case 'CommandText' */:
                  break;
                case "refreshinfo" /* case 'RefreshInfo' */:
                  break;
                case "notitles" /* case 'NoTitles' */:
                  break;
                case "nextid" /* case 'NextId' */:
                  break;
                case "columninfo" /* case 'ColumnInfo' */:
                  break;
                case "overwritecells" /* case 'OverwriteCells' */:
                  break;
                case "donotpromptforfile" /* case 'DoNotPromptForFile' */:
                  break;
                case "textwizardsettings" /* case 'TextWizardSettings' */:
                  break;
                case "source" /* case 'Source' */:
                  break;
                case "number" /* case 'Number' */:
                  break;
                case "decimal" /* case 'Decimal' */:
                  break;
                case "thousandseparator" /* case 'ThousandSeparator' */:
                  break;
                case "trailingminusnumbers" /* case 'TrailingMinusNumbers' */:
                  break;
                case "formatsettings" /* case 'FormatSettings' */:
                  break;
                case "fieldtype" /* case 'FieldType' */:
                  break;
                case "delimiters" /* case 'Delimiters' */:
                  break;
                case "tab" /* case 'Tab' */:
                  break;
                case "comma" /* case 'Comma' */:
                  break;
                case "autoformatname" /* case 'AutoFormatName' */:
                  break;
                case "versionlastedit" /* case 'VersionLastEdit' */:
                  break;
                case "versionlastrefresh" /* case 'VersionLastRefresh' */:
                  break;
                default:
                  seen = false;
              }
              break;

            case "datavalidation" /* case 'DataValidation' */:
              switch (Rn[3]) {
                case "range" /* case 'Range' */:
                  break;

                case "type" /* case 'Type' */:
                  break;
                case "min" /* case 'Min' */:
                  break;
                case "max" /* case 'Max' */:
                  break;
                case "sort" /* case 'Sort' */:
                  break;
                case "descending" /* case 'Descending' */:
                  break;
                case "order" /* case 'Order' */:
                  break;
                case "casesensitive" /* case 'CaseSensitive' */:
                  break;
                case "value" /* case 'Value' */:
                  break;
                case "errorstyle" /* case 'ErrorStyle' */:
                  break;
                case "errormessage" /* case 'ErrorMessage' */:
                  break;
                case "errortitle" /* case 'ErrorTitle' */:
                  break;
                case "inputmessage" /* case 'InputMessage' */:
                  break;
                case "inputtitle" /* case 'InputTitle' */:
                  break;
                case "combohide" /* case 'ComboHide' */:
                  break;
                case "inputhide" /* case 'InputHide' */:
                  break;
                case "condition" /* case 'Condition' */:
                  break;
                case "qualifier" /* case 'Qualifier' */:
                  break;
                case "useblank" /* case 'UseBlank' */:
                  break;
                case "value1" /* case 'Value1' */:
                  break;
                case "value2" /* case 'Value2' */:
                  break;
                case "format" /* case 'Format' */:
                  break;

                case "cellrangelist" /* case 'CellRangeList' */:
                  break;
                default:
                  seen = false;
              }
              break;

            case "sorting" /* case 'Sorting' */:
            case "conditionalformatting" /* case 'ConditionalFormatting' */:
              switch (Rn[3]) {
                case "range" /* case 'Range' */:
                  break;
                case "type" /* case 'Type' */:
                  break;
                case "min" /* case 'Min' */:
                  break;
                case "max" /* case 'Max' */:
                  break;
                case "sort" /* case 'Sort' */:
                  break;
                case "descending" /* case 'Descending' */:
                  break;
                case "order" /* case 'Order' */:
                  break;
                case "casesensitive" /* case 'CaseSensitive' */:
                  break;
                case "value" /* case 'Value' */:
                  break;
                case "errorstyle" /* case 'ErrorStyle' */:
                  break;
                case "errormessage" /* case 'ErrorMessage' */:
                  break;
                case "errortitle" /* case 'ErrorTitle' */:
                  break;
                case "cellrangelist" /* case 'CellRangeList' */:
                  break;
                case "inputmessage" /* case 'InputMessage' */:
                  break;
                case "inputtitle" /* case 'InputTitle' */:
                  break;
                case "combohide" /* case 'ComboHide' */:
                  break;
                case "inputhide" /* case 'InputHide' */:
                  break;
                case "condition" /* case 'Condition' */:
                  break;
                case "qualifier" /* case 'Qualifier' */:
                  break;
                case "useblank" /* case 'UseBlank' */:
                  break;
                case "value1" /* case 'Value1' */:
                  break;
                case "value2" /* case 'Value2' */:
                  break;
                case "format" /* case 'Format' */:
                  break;
                default:
                  seen = false;
              }
              break;

            /* MapInfo (schema) */
            case "mapinfo" /* case 'MapInfo' */:
            case "schema" /* case 'Schema' */:
            case "data" /* case 'data' */:
              switch (Rn[3]) {
                case "map" /* case 'Map' */:
                  break;
                case "entry" /* case 'Entry' */:
                  break;
                case "range" /* case 'Range' */:
                  break;
                case "xpath" /* case 'XPath' */:
                  break;
                case "field" /* case 'Field' */:
                  break;
                case "xsdtype" /* case 'XSDType' */:
                  break;
                case "filteron" /* case 'FilterOn' */:
                  break;
                case "aggregate" /* case 'Aggregate' */:
                  break;
                case "elementtype" /* case 'ElementType' */:
                  break;
                case "attributetype" /* case 'AttributeType' */:
                  break;
                /* These are from xsd (XML Schema Definition) */
                case "schema" /* case 'schema' */:
                case "element" /* case 'element' */:
                case "complextype" /* case 'complexType' */:
                case "datatype" /* case 'datatype' */:
                case "all" /* case 'all' */:
                case "attribute" /* case 'attribute' */:
                case "extends" /* case 'extends' */:
                  break;

                case "row" /* case 'row' */:
                  break;
                default:
                  seen = false;
              }
              break;

            /* SmartTags (can be anything) */
            case "smarttags" /* case 'SmartTags' */:
              break;

            default:
              seen = false;
              break;
          }
          if (seen) break;
          /* CustomDocumentProperties */
          if (Rn[3].match(/!\[CDATA/)) break;
          if (!state[state.length - 1][1])
            throw `Unrecognized tag: ${Rn[3]}|${state.join("|")}`;
          if (
            state[state.length - 1][0] ===
            /* 'CustomDocumentProperties' */ "customdocumentproperties"
          ) {
            if (Rn[0].slice(-2) === "/>") break;
            else if (Rn[1] === "/")
              xlml_set_custprop(
                Custprops,
                raw_Rn3,
                cp,
                str.slice(pidx, Rn.index)
              );
            else {
              cp = Rn;
              pidx = Rn.index + Rn[0].length;
            }
            break;
          }
          if (opts.WTF) throw `Unrecognized tag: ${Rn[3]}|${state.join("|")}`;
      }
    const out = {};
    if (!opts.bookSheets && !opts.bookProps) out.Sheets = sheets;
    out.SheetNames = sheetnames;
    out.Workbook = Workbook;
    out.SSF = SSF.get_table();
    out.Props = Props;
    out.Custprops = Custprops;
    return out;
  }

  function parse_xlml(data, opts) {
    fix_read_opts((opts = opts || {}));
    switch (opts.type || "base64") {
      case "base64":
        return parse_xlml_xml(Base64.decode(data), opts);
      case "binary":
      case "buffer":
      case "file":
        return parse_xlml_xml(data, opts);
      case "array":
        return parse_xlml_xml(a2s(data), opts);
    }
  }

  /* [MS-OLEDS] 2.3.8 CompObjStream */
  function parse_compobj(obj) {
    const v = {};
    const o = obj.content;
    /* [MS-OLEDS] 2.3.7 CompObjHeader -- All fields MUST be ignored */
    o.l = 28;

    v.AnsiUserType = o.read_shift(0, "lpstr-ansi");
    v.AnsiClipboardFormat = parse_ClipboardFormatOrAnsiString(o);

    if (o.length - o.l <= 4) return v;

    let m = o.read_shift(4);
    if (m == 0 || m > 40) return v;
    o.l -= 4;
    v.Reserved1 = o.read_shift(0, "lpstr-ansi");

    if (o.length - o.l <= 4) return v;
    m = o.read_shift(4);
    if (m !== 0x71b239f4) return v;
    v.UnicodeClipboardFormat = parse_ClipboardFormatOrUnicodeString(o);

    m = o.read_shift(4);
    if (m == 0 || m > 40) return v;
    o.l -= 4;
    v.Reserved2 = o.read_shift(0, "lpwstr");
  }

  /*
	Continue logic for:
	- 2.4.58 Continue
	- 2.4.59 ContinueBigName
	- 2.4.60 ContinueFrt
	- 2.4.61 ContinueFrt11
	- 2.4.62 ContinueFrt12
*/
  function slurp(R, blob, length, opts) {
    let l = length;
    const bufs = [];
    const d = blob.slice(blob.l, blob.l + l);
    if (opts && opts.enc && opts.enc.insitu)
      switch (R.n) {
        case "BOF":
        case "FilePass":
        case "FileLock":
        case "InterfaceHdr":
        case "RRDInfo":
        case "RRDHead":
        case "UsrExcl":
          break;
        default:
          if (d.length === 0) break;
          opts.enc.insitu(d);
      }
    bufs.push(d);
    blob.l += l;
    let next = XLSRecordEnum[__readUInt16LE(blob, blob.l)];
    let start = 0;
    while (next != null && next.n.slice(0, 8) === "Continue") {
      l = __readUInt16LE(blob, blob.l + 2);
      start = blob.l + 4;
      if (next.n == "ContinueFrt") start += 4;
      else if (next.n.slice(0, 11) == "ContinueFrt") start += 12;
      bufs.push(blob.slice(start, blob.l + 4 + l));
      blob.l += 4 + l;
      next = XLSRecordEnum[__readUInt16LE(blob, blob.l)];
    }
    const b = bconcat(bufs);
    prep_blob(b, 0);
    let ll = 0;
    b.lens = [];
    for (let j = 0; j < bufs.length; ++j) {
      b.lens.push(ll);
      ll += bufs[j].length;
    }
    return R.f(b, b.length, opts);
  }

  function safe_format_xf(p, opts, date1904) {
    if (p.t === "z") return;
    if (!p.XF) return;
    let fmtid = 0;
    try {
      fmtid = p.z || p.XF.numFmtId || 0;
      if (opts.cellNF) p.z = SSF._table[fmtid];
    } catch (e) {
      if (opts.WTF) throw e;
    }
    if (!opts || opts.cellText !== false)
      try {
        if (p.t === "e") {
          p.w = p.w || BErr[p.v];
        } else if (fmtid === 0 || fmtid == "General") {
          if (p.t === "n") {
            if ((p.v | 0) === p.v) p.w = SSF._general_int(p.v);
            else p.w = SSF._general_num(p.v);
          } else p.w = SSF._general(p.v);
        } else p.w = SSF.format(fmtid, p.v, { date1904: !!date1904 });
      } catch (e) {
        if (opts.WTF) throw e;
      }
    if (
      opts.cellDates &&
      fmtid &&
      p.t == "n" &&
      SSF.is_date(SSF._table[fmtid] || String(fmtid))
    ) {
      const _d = SSF.parse_date_code(p.v);
      if (_d) {
        p.t = "d";
        p.v = new Date(_d.y, _d.m - 1, _d.d, _d.H, _d.M, _d.S, _d.u);
      }
    }
  }

  function make_cell(val, ixfe, t) {
    return { v: val, ixfe, t };
  }

  // 2.3.2
  function parse_workbook(blob, options) {
    const wb = { opts: {} };
    const Sheets = {};
    if (DENSE != null && options.dense == null) options.dense = DENSE;
    let out = options.dense ? [] : {};
    const Directory = {};
    let range = {};
    let last_formula = null;
    let sst = [];
    let cur_sheet = "";
    let Preamble = {};
    let lastcell;
    let last_cell = "";
    let cc;
    let cmnt;
    let rngC;
    let rngR;
    const sharedf = {};
    let arrayf = [];
    let temp_val;
    let country;
    let cell_valid = true;
    const XFs = []; /* XF records */
    let palette = [];
    const Workbook = { Sheets: [], WBProps: { date1904: false }, Views: [{}] };
    let wsprops = {};
    const get_rgb = function getrgb(icv) {
      if (icv < 8) return XLSIcv[icv];
      if (icv < 64) return palette[icv - 8] || XLSIcv[icv];
      return XLSIcv[icv];
    };
    const process_cell_style = function pcs(cell, line, options) {
      const xfd = line.XF.data;
      if (!xfd || !xfd.patternType || !options || !options.cellStyles) return;
      line.s = {};
      line.s.patternType = xfd.patternType;
      let t;
      if ((t = rgb2Hex(get_rgb(xfd.icvFore)))) {
        line.s.fgColor = { rgb: t };
      }
      if ((t = rgb2Hex(get_rgb(xfd.icvBack)))) {
        line.s.bgColor = { rgb: t };
      }
    };
    const addcell = function addcell(cell, line, options) {
      if (file_depth > 1) return;
      if (options.sheetRows && cell.r >= options.sheetRows) cell_valid = false;
      if (!cell_valid) return;
      if (options.cellStyles && line.XF && line.XF.data)
        process_cell_style(cell, line, options);
      delete line.ixfe;
      delete line.XF;
      lastcell = cell;
      last_cell = encode_cell(cell);
      if (!range || !range.s || !range.e)
        range = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
      if (cell.r < range.s.r) range.s.r = cell.r;
      if (cell.c < range.s.c) range.s.c = cell.c;
      if (cell.r + 1 > range.e.r) range.e.r = cell.r + 1;
      if (cell.c + 1 > range.e.c) range.e.c = cell.c + 1;
      if (options.cellFormula && line.f) {
        for (let afi = 0; afi < arrayf.length; ++afi) {
          if (arrayf[afi][0].s.c > cell.c || arrayf[afi][0].s.r > cell.r)
            continue;
          if (arrayf[afi][0].e.c < cell.c || arrayf[afi][0].e.r < cell.r)
            continue;
          line.F = encode_range(arrayf[afi][0]);
          if (arrayf[afi][0].s.c != cell.c || arrayf[afi][0].s.r != cell.r)
            delete line.f;
          if (line.f)
            line.f = `${stringify_formula(
              arrayf[afi][1],
              range,
              cell,
              supbooks,
              opts
            )}`;
          break;
        }
      }
      {
        if (options.dense) {
          if (!out[cell.r]) out[cell.r] = [];
          out[cell.r][cell.c] = line;
        } else out[last_cell] = line;
      }
    };
    var opts = {
      enc: false, // encrypted
      sbcch: 0, // cch in the preceding SupBook
      snames: [], // sheetnames
      sharedf, // shared formulae by address
      arrayf, // array formulae array
      rrtabid: [], // RRTabId
      lastuser: "", // Last User from WriteAccess
      biff: 8, // BIFF version
      codepage: 0, // CP from CodePage record
      winlocked: 0, // fLockWn from WinProtect
      cellStyles: !!options && !!options.cellStyles,
      WTF: !!options && !!options.wtf
    };
    if (options.password) opts.password = options.password;
    let themes;
    let merges = [];
    let objects = [];
    let colinfo = [];
    let rowinfo = [];
    // eslint-disable-next-line no-unused-vars
    let defwidth = 0;
    let defheight = 0; // twips / MDW respectively
    let seencol = false;
    var supbooks = []; // 1-indexed, will hold extern names
    supbooks.SheetNames = opts.snames;
    supbooks.sharedf = opts.sharedf;
    supbooks.arrayf = opts.arrayf;
    supbooks.names = [];
    supbooks.XTI = [];
    let last_Rn = "";
    var file_depth = 0; /* TODO: make a real stack */
    let BIFF2Fmt = 0;
    const BIFF2FmtTable = [];
    const FilterDatabases = []; /* TODO: sort out supbooks and process elsewhere */
    let last_lbl;

    /* explicit override for some broken writers */
    opts.codepage = 1200;
    set_cp(1200);
    let seen_codepage = false;
    while (blob.l < blob.length - 1) {
      const s = blob.l;
      const RecordType = blob.read_shift(2);
      if (RecordType === 0 && last_Rn === "EOF") break;
      let length = blob.l === blob.length ? 0 : blob.read_shift(2);
      const R = XLSRecordEnum[RecordType];
      // console.log(RecordType.toString(16), RecordType, R, blob.l, length, blob.length);
      // if(!R) console.log(blob.slice(blob.l, blob.l + length));
      if (R && R.f) {
        if (options.bookSheets) {
          if (last_Rn === "BoundSheet8" && R.n !== "BoundSheet8") break;
        }
        last_Rn = R.n;
        if (R.r === 2 || R.r == 12) {
          const rt = blob.read_shift(2);
          length -= 2;
          if (
            !opts.enc &&
            rt !== RecordType &&
            (((rt & 0xff) << 8) | (rt >> 8)) !== RecordType
          )
            throw new Error(`rt mismatch: ${rt}!=${RecordType}`);
          if (R.r == 12) {
            blob.l += 10;
            length -= 10;
          } // skip FRT
        }
        // console.error(R,blob.l,length,blob.length);
        let val = {};
        if (R.n === "EOF") val = R.f(blob, length, opts);
        else val = slurp(R, blob, length, opts);
        const Rn = R.n;
        if (file_depth == 0 && Rn != "BOF") continue;
        /* nested switch statements to workaround V8 128 limit */
        switch (Rn) {
          /* Workbook Options */
          case "Date1904":
            wb.opts.Date1904 = Workbook.WBProps.date1904 = val;
            break;
          case "WriteProtect":
            wb.opts.WriteProtect = true;
            break;
          case "FilePass":
            if (!opts.enc) blob.l = 0;
            opts.enc = val;
            if (!options.password)
              throw new Error("File is password-protected");
            if (val.valid == null)
              throw new Error("Encryption scheme unsupported");
            if (!val.valid) throw new Error("Password is incorrect");
            break;
          case "WriteAccess":
            opts.lastuser = val;
            break;
          case "FileSharing":
            break; // TODO
          case "CodePage":
            var cpval = Number(val);
            /* overrides based on test cases */
            switch (cpval) {
              case 0x5212:
                cpval = 1200;
                break;
              case 0x8000:
                cpval = 10000;
                break;
              case 0x8001:
                cpval = 1252;
                break;
            }
            set_cp((opts.codepage = cpval));
            seen_codepage = true;
            break;
          case "RRTabId":
            opts.rrtabid = val;
            break;
          case "WinProtect":
            opts.winlocked = val;
            break;
          case "Template":
            break; // TODO
          case "BookBool":
            break; // TODO
          case "UsesELFs":
            break;
          case "MTRSettings":
            break;
          case "RefreshAll":
          case "CalcCount":
          case "CalcDelta":
          case "CalcIter":
          case "CalcMode":
          case "CalcPrecision":
          case "CalcSaveRecalc":
            wb.opts[Rn] = val;
            break;
          case "CalcRefMode":
            opts.CalcRefMode = val;
            break; // TODO: implement R1C1
          case "Uncalced":
            break;
          case "ForceFullCalculation":
            wb.opts.FullCalc = val;
            break;
          case "WsBool":
            if (val.fDialog) out["!type"] = "dialog";
            break; // TODO
          case "XF":
            XFs.push(val);
            break;
          case "ExtSST":
            break; // TODO
          case "BookExt":
            break; // TODO
          case "RichTextStream":
            break;
          case "BkHim":
            break;

          case "SupBook":
            supbooks.push([val]);
            supbooks[supbooks.length - 1].XTI = [];
            break;
          case "ExternName":
            supbooks[supbooks.length - 1].push(val);
            break;
          case "Index":
            break; // TODO
          case "Lbl":
            last_lbl = {
              Name: val.Name,
              Ref: stringify_formula(val.rgce, range, null, supbooks, opts)
            };
            if (val.itab > 0) last_lbl.Sheet = val.itab - 1;
            supbooks.names.push(last_lbl);
            if (!supbooks[0]) {
              supbooks[0] = [];
              supbooks[0].XTI = [];
            }
            supbooks[supbooks.length - 1].push(val);
            if (val.Name == "_xlnm._FilterDatabase" && val.itab > 0)
              if (
                val.rgce &&
                val.rgce[0] &&
                val.rgce[0][0] &&
                val.rgce[0][0][0] == "PtgArea3d"
              )
                FilterDatabases[val.itab - 1] = {
                  ref: encode_range(val.rgce[0][0][1][2])
                };
            break;
          case "ExternCount":
            opts.ExternCount = val;
            break;
          case "ExternSheet":
            if (supbooks.length == 0) {
              supbooks[0] = [];
              supbooks[0].XTI = [];
            }
            supbooks[supbooks.length - 1].XTI = supbooks[
              supbooks.length - 1
            ].XTI.concat(val);
            supbooks.XTI = supbooks.XTI.concat(val);
            break;
          case "NameCmt":
            /* TODO: search for correct name */
            if (opts.biff < 8) break;
            if (last_lbl != null) last_lbl.Comment = val[1];
            break;

          case "Protect":
            out["!protect"] = val;
            break; /* for sheet or book */
          case "Password":
            if (val !== 0 && opts.WTF)
              console.error(`Password verifier: ${val}`);
            break;
          case "Prot4Rev":
          case "Prot4RevPass":
            break; /* TODO: Revision Control */

          case "BoundSheet8":
            {
              Directory[val.pos] = val;
              opts.snames.push(val.name);
            }
            break;
          case "EOF":
            {
              if (--file_depth) break;
              if (range.e) {
                if (range.e.r > 0 && range.e.c > 0) {
                  range.e.r--;
                  range.e.c--;
                  out["!ref"] = encode_range(range);
                  if (options.sheetRows && options.sheetRows <= range.e.r) {
                    const tmpri = range.e.r;
                    range.e.r = options.sheetRows - 1;
                    out["!fullref"] = out["!ref"];
                    out["!ref"] = encode_range(range);
                    range.e.r = tmpri;
                  }
                  range.e.r++;
                  range.e.c++;
                }
                if (merges.length > 0) out["!merges"] = merges;
                if (objects.length > 0) out["!objects"] = objects;
                if (colinfo.length > 0) out["!cols"] = colinfo;
                if (rowinfo.length > 0) out["!rows"] = rowinfo;
                Workbook.Sheets.push(wsprops);
              }
              if (cur_sheet === "") Preamble = out;
              else Sheets[cur_sheet] = out;
              out = options.dense ? [] : {};
            }
            break;
          case "BOF":
            {
              if (opts.biff === 8)
                opts.biff =
                  {
                    0x0009: 2,
                    0x0209: 3,
                    0x0409: 4
                  }[RecordType] ||
                  {
                    0x0200: 2,
                    0x0300: 3,
                    0x0400: 4,
                    0x0500: 5,
                    0x0600: 8,
                    0x0002: 2,
                    0x0007: 2
                  }[val.BIFFVer] ||
                  8;
              if (opts.biff == 8 && val.BIFFVer == 0 && val.dt == 16)
                opts.biff = 2;
              if (file_depth++) break;
              cell_valid = true;
              out = options.dense ? [] : {};

              if (opts.biff < 8 && !seen_codepage) {
                seen_codepage = true;
                set_cp((opts.codepage = options.codepage || 1252));
              }
              if (opts.biff < 5) {
                if (cur_sheet === "") cur_sheet = "Sheet1";
                range = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
                /* fake BoundSheet8 */
                const fakebs8 = { pos: blob.l - length, name: cur_sheet };
                Directory[fakebs8.pos] = fakebs8;
                opts.snames.push(cur_sheet);
              } else cur_sheet = (Directory[s] || { name: "" }).name;
              if (val.dt == 0x20) out["!type"] = "chart";
              if (val.dt == 0x40) out["!type"] = "macro";
              merges = [];
              objects = [];
              opts.arrayf = arrayf = [];
              colinfo = [];
              rowinfo = [];
              defwidth = defheight = 0;
              seencol = false;
              wsprops = {
                Hidden: (Directory[s] || { hs: 0 }).hs,
                name: cur_sheet
              };
            }
            break;

          case "Number":
          case "BIFF2NUM":
          case "BIFF2INT":
            {
              if (out["!type"] == "chart")
                if (
                  options.dense
                    ? (out[val.r] || [])[val.c]
                    : out[encode_cell({ c: val.c, r: val.r })]
                )
                  ++val.c;
              temp_val = {
                ixfe: val.ixfe,
                XF: XFs[val.ixfe] || {},
                v: val.val,
                t: "n"
              };
              if (BIFF2Fmt > 0)
                temp_val.z = BIFF2FmtTable[(temp_val.ixfe >> 8) & 0x1f];
              safe_format_xf(temp_val, options, wb.opts.Date1904);
              addcell({ c: val.c, r: val.r }, temp_val, options);
            }
            break;
          case "BoolErr":
            {
              temp_val = {
                ixfe: val.ixfe,
                XF: XFs[val.ixfe],
                v: val.val,
                t: val.t
              };
              if (BIFF2Fmt > 0)
                temp_val.z = BIFF2FmtTable[(temp_val.ixfe >> 8) & 0x1f];
              safe_format_xf(temp_val, options, wb.opts.Date1904);
              addcell({ c: val.c, r: val.r }, temp_val, options);
            }
            break;
          case "RK":
            {
              temp_val = {
                ixfe: val.ixfe,
                XF: XFs[val.ixfe],
                v: val.rknum,
                t: "n"
              };
              if (BIFF2Fmt > 0)
                temp_val.z = BIFF2FmtTable[(temp_val.ixfe >> 8) & 0x1f];
              safe_format_xf(temp_val, options, wb.opts.Date1904);
              addcell({ c: val.c, r: val.r }, temp_val, options);
            }
            break;
          case "MulRk":
            {
              for (let j = val.c; j <= val.C; ++j) {
                const ixfe = val.rkrec[j - val.c][0];
                temp_val = {
                  ixfe,
                  XF: XFs[ixfe],
                  v: val.rkrec[j - val.c][1],
                  t: "n"
                };
                if (BIFF2Fmt > 0)
                  temp_val.z = BIFF2FmtTable[(temp_val.ixfe >> 8) & 0x1f];
                safe_format_xf(temp_val, options, wb.opts.Date1904);
                addcell({ c: j, r: val.r }, temp_val, options);
              }
            }
            break;
          case "Formula":
            {
              if (val.val == "String") {
                last_formula = val;
                break;
              }
              temp_val = make_cell(val.val, val.cell.ixfe, val.tt);
              temp_val.XF = XFs[temp_val.ixfe];
              if (options.cellFormula) {
                const _f = val.formula;
                if (_f && _f[0] && _f[0][0] && _f[0][0][0] == "PtgExp") {
                  const _fr = _f[0][0][1][0];
                  const _fc = _f[0][0][1][1];
                  const _fe = encode_cell({ r: _fr, c: _fc });
                  if (sharedf[_fe])
                    temp_val.f = `${stringify_formula(
                      val.formula,
                      range,
                      val.cell,
                      supbooks,
                      opts
                    )}`;
                  else
                    temp_val.F = (
                      (options.dense ? (out[_fr] || [])[_fc] : out[_fe]) || {}
                    ).F;
                } else
                  temp_val.f = `${stringify_formula(
                    val.formula,
                    range,
                    val.cell,
                    supbooks,
                    opts
                  )}`;
              }
              if (BIFF2Fmt > 0)
                temp_val.z = BIFF2FmtTable[(temp_val.ixfe >> 8) & 0x1f];
              safe_format_xf(temp_val, options, wb.opts.Date1904);
              addcell(val.cell, temp_val, options);
              last_formula = val;
            }
            break;
          case "String":
            {
              if (last_formula) {
                /* technically always true */
                last_formula.val = val;
                temp_val = make_cell(val, last_formula.cell.ixfe, "s");
                temp_val.XF = XFs[temp_val.ixfe];
                if (options.cellFormula) {
                  temp_val.f = `${stringify_formula(
                    last_formula.formula,
                    range,
                    last_formula.cell,
                    supbooks,
                    opts
                  )}`;
                }
                if (BIFF2Fmt > 0)
                  temp_val.z = BIFF2FmtTable[(temp_val.ixfe >> 8) & 0x1f];
                safe_format_xf(temp_val, options, wb.opts.Date1904);
                addcell(last_formula.cell, temp_val, options);
                last_formula = null;
              } else throw new Error("String record expects Formula");
            }
            break;
          case "Array":
            {
              arrayf.push(val);
              const _arraystart = encode_cell(val[0].s);
              cc = options.dense
                ? (out[val[0].s.r] || [])[val[0].s.c]
                : out[_arraystart];
              if (options.cellFormula && cc) {
                if (!last_formula) break; /* technically unreachable */
                if (!_arraystart || !cc) break;
                cc.f = `${stringify_formula(
                  val[1],
                  range,
                  val[0],
                  supbooks,
                  opts
                )}`;
                cc.F = encode_range(val[0]);
              }
            }
            break;
          case "ShrFmla":
            {
              if (!cell_valid) break;
              if (!options.cellFormula) break;
              if (last_cell) {
                /* TODO: capture range */
                if (!last_formula) break; /* technically unreachable */
                sharedf[encode_cell(last_formula.cell)] = val[0];
                cc = options.dense
                  ? (out[last_formula.cell.r] || [])[last_formula.cell.c]
                  : out[encode_cell(last_formula.cell)];
                (cc || {}).f = `${stringify_formula(
                  val[0],
                  range,
                  lastcell,
                  supbooks,
                  opts
                )}`;
              }
            }
            break;
          case "LabelSst":
            temp_val = make_cell(sst[val.isst].t, val.ixfe, "s");
            if (sst[val.isst].h) temp_val.h = sst[val.isst].h;
            temp_val.XF = XFs[temp_val.ixfe];
            if (BIFF2Fmt > 0)
              temp_val.z = BIFF2FmtTable[(temp_val.ixfe >> 8) & 0x1f];
            safe_format_xf(temp_val, options, wb.opts.Date1904);
            addcell({ c: val.c, r: val.r }, temp_val, options);
            break;
          case "Blank":
            if (options.sheetStubs) {
              temp_val = { ixfe: val.ixfe, XF: XFs[val.ixfe], t: "z" };
              if (BIFF2Fmt > 0)
                temp_val.z = BIFF2FmtTable[(temp_val.ixfe >> 8) & 0x1f];
              safe_format_xf(temp_val, options, wb.opts.Date1904);
              addcell({ c: val.c, r: val.r }, temp_val, options);
            }
            break;
          case "MulBlank":
            if (options.sheetStubs) {
              for (let _j = val.c; _j <= val.C; ++_j) {
                const _ixfe = val.ixfe[_j - val.c];
                temp_val = { ixfe: _ixfe, XF: XFs[_ixfe], t: "z" };
                if (BIFF2Fmt > 0)
                  temp_val.z = BIFF2FmtTable[(temp_val.ixfe >> 8) & 0x1f];
                safe_format_xf(temp_val, options, wb.opts.Date1904);
                addcell({ c: _j, r: val.r }, temp_val, options);
              }
            }
            break;
          case "RString":
          case "Label":
          case "BIFF2STR":
            temp_val = make_cell(val.val, val.ixfe, "s");
            temp_val.XF = XFs[temp_val.ixfe];
            if (BIFF2Fmt > 0)
              temp_val.z = BIFF2FmtTable[(temp_val.ixfe >> 8) & 0x1f];
            safe_format_xf(temp_val, options, wb.opts.Date1904);
            addcell({ c: val.c, r: val.r }, temp_val, options);
            break;

          case "Dimensions":
            {
              if (file_depth === 1) range = val; /* TODO: stack */
            }
            break;
          case "SST":
            {
              sst = val;
            }
            break;
          case "Format":
            {
              /* val = [id, fmt] */
              if (opts.biff == 4) {
                BIFF2FmtTable[BIFF2Fmt++] = val[1];
                for (var b4idx = 0; b4idx < BIFF2Fmt + 163; ++b4idx)
                  if (SSF._table[b4idx] == val[1]) break;
                if (b4idx >= 163) SSF.load(val[1], BIFF2Fmt + 163);
              } else SSF.load(val[1], val[0]);
            }
            break;
          case "BIFF2FORMAT":
            {
              BIFF2FmtTable[BIFF2Fmt++] = val;
              for (var b2idx = 0; b2idx < BIFF2Fmt + 163; ++b2idx)
                if (SSF._table[b2idx] == val) break;
              if (b2idx >= 163) SSF.load(val, BIFF2Fmt + 163);
            }
            break;

          case "MergeCells":
            merges = merges.concat(val);
            break;

          case "Obj":
            objects[val.cmo[0]] = opts.lastobj = val;
            break;
          case "TxO":
            opts.lastobj.TxO = val;
            break;
          case "ImData":
            opts.lastobj.ImData = val;
            break;

          case "HLink":
            {
              for (rngR = val[0].s.r; rngR <= val[0].e.r; ++rngR)
                for (rngC = val[0].s.c; rngC <= val[0].e.c; ++rngC) {
                  cc = options.dense
                    ? (out[rngR] || [])[rngC]
                    : out[encode_cell({ c: rngC, r: rngR })];
                  if (cc) cc.l = val[1];
                }
            }
            break;
          case "HLinkTooltip":
            {
              for (rngR = val[0].s.r; rngR <= val[0].e.r; ++rngR)
                for (rngC = val[0].s.c; rngC <= val[0].e.c; ++rngC) {
                  cc = options.dense
                    ? (out[rngR] || [])[rngC]
                    : out[encode_cell({ c: rngC, r: rngR })];
                  if (cc && cc.l) cc.l.Tooltip = val[1];
                }
            }
            break;

          /* Comments */
          case "Note":
            {
              if (opts.biff <= 5 && opts.biff >= 2) break; /* TODO: BIFF5 */
              cc = options.dense
                ? (out[val[0].r] || [])[val[0].c]
                : out[encode_cell(val[0])];
              const noteobj = objects[val[2]];
              if (!cc) {
                if (options.dense) {
                  if (!out[val[0].r]) out[val[0].r] = [];
                  cc = out[val[0].r][val[0].c] = { t: "z" };
                } else {
                  cc = out[encode_cell(val[0])] = { t: "z" };
                }
                range.e.r = Math.max(range.e.r, val[0].r);
                range.s.r = Math.min(range.s.r, val[0].r);
                range.e.c = Math.max(range.e.c, val[0].c);
                range.s.c = Math.min(range.s.c, val[0].c);
              }
              if (!cc.c) cc.c = [];
              cmnt = { a: val[1], t: noteobj.TxO.t };
              cc.c.push(cmnt);
            }
            break;

          default:
            switch (R.n /* nested */) {
              case "ClrtClient":
                break;
              case "XFExt":
                update_xfext(XFs[val.ixfe], val.ext);
                break;

              case "DefColWidth":
                defwidth = val;
                break;
              case "DefaultRowHeight":
                defheight = val[1];
                break; // TODO: flags

              case "ColInfo":
                {
                  if (!opts.cellStyles) break;
                  while (val.e >= val.s) {
                    colinfo[val.e--] = { width: val.w / 256 };
                    if (!seencol) {
                      seencol = true;
                      find_mdw_colw(val.w / 256);
                    }
                    process_col(colinfo[val.e + 1]);
                  }
                }
                break;
              case "Row":
                {
                  const rowobj = {};
                  if (val.level != null) {
                    rowinfo[val.r] = rowobj;
                    rowobj.level = val.level;
                  }
                  if (val.hidden) {
                    rowinfo[val.r] = rowobj;
                    rowobj.hidden = true;
                  }
                  if (val.hpt) {
                    rowinfo[val.r] = rowobj;
                    rowobj.hpt = val.hpt;
                    rowobj.hpx = pt2px(val.hpt);
                  }
                }
                break;

              case "LeftMargin":
              case "RightMargin":
              case "TopMargin":
              case "BottomMargin":
                if (!out["!margins"]) default_margins((out["!margins"] = {}));
                out["!margins"][Rn.slice(0, -6).toLowerCase()] = val;
                break;

              case "Setup": // TODO
                if (!out["!margins"]) default_margins((out["!margins"] = {}));
                out["!margins"].header = val.header;
                out["!margins"].footer = val.footer;
                break;

              case "Window2": // TODO
                // $FlowIgnore
                if (val.RTL) Workbook.Views[0].RTL = true;
                break;

              case "Header":
                break; // TODO
              case "Footer":
                break; // TODO
              case "HCenter":
                break; // TODO
              case "VCenter":
                break; // TODO
              case "Pls":
                break; // TODO
              case "GCW":
                break;
              case "LHRecord":
                break;
              case "DBCell":
                break; // TODO
              case "EntExU2":
                break; // TODO
              case "SxView":
                break; // TODO
              case "Sxvd":
                break; // TODO
              case "SXVI":
                break; // TODO
              case "SXVDEx":
                break; // TODO
              case "SxIvd":
                break; // TODO
              case "SXString":
                break; // TODO
              case "Sync":
                break;
              case "Addin":
                break;
              case "SXDI":
                break; // TODO
              case "SXLI":
                break; // TODO
              case "SXEx":
                break; // TODO
              case "QsiSXTag":
                break; // TODO
              case "Selection":
                break;
              case "Feat":
                break;
              case "FeatHdr":
              case "FeatHdr11":
                break;
              case "Feature11":
              case "Feature12":
              case "List12":
                break;
              case "Country":
                country = val;
                break;
              case "RecalcId":
                break;
              case "DxGCol":
                break; // TODO: htmlify
              case "Fbi":
              case "Fbi2":
              case "GelFrame":
                break;
              case "Font":
                break; // TODO
              case "XFCRC":
                break; // TODO
              case "Style":
                break; // TODO
              case "StyleExt":
                break; // TODO
              case "Palette":
                palette = val;
                break;
              case "Theme":
                themes = val;
                break;
              /* Protection */
              case "ScenarioProtect":
                break;
              case "ObjProtect":
                break;

              /* Conditional Formatting */
              case "CondFmt12":
                break;

              /* Table */
              case "Table":
                break; // TODO
              case "TableStyles":
                break; // TODO
              case "TableStyle":
                break; // TODO
              case "TableStyleElement":
                break; // TODO

              /* PivotTable */
              case "SXStreamID":
                break; // TODO
              case "SXVS":
                break; // TODO
              case "DConRef":
                break; // TODO
              case "SXAddl":
                break; // TODO
              case "DConBin":
                break; // TODO
              case "DConName":
                break; // TODO
              case "SXPI":
                break; // TODO
              case "SxFormat":
                break; // TODO
              case "SxSelect":
                break; // TODO
              case "SxRule":
                break; // TODO
              case "SxFilt":
                break; // TODO
              case "SxItm":
                break; // TODO
              case "SxDXF":
                break; // TODO

              /* Scenario Manager */
              case "ScenMan":
                break;

              /* Data Consolidation */
              case "DCon":
                break;

              /* Watched Cell */
              case "CellWatch":
                break;

              /* Print Settings */
              case "PrintRowCol":
                break;
              case "PrintGrid":
                break;
              case "PrintSize":
                break;

              case "XCT":
                break;
              case "CRN":
                break;

              case "Scl":
                {
                  // console.log("Zoom Level:", val[0]/val[1],val);
                }
                break;
              case "SheetExt":
                {
                  /* empty */
                }
                break;
              case "SheetExtOptional":
                {
                  /* empty */
                }
                break;

              /* VBA */
              case "ObNoMacros":
                {
                  /* empty */
                }
                break;
              case "ObProj":
                {
                  /* empty */
                }
                break;
              case "CodeName":
                {
                  if (!cur_sheet)
                    Workbook.WBProps.CodeName = val || "ThisWorkbook";
                  else wsprops.CodeName = val || wsprops.name;
                }
                break;
              case "GUIDTypeLib":
                {
                  /* empty */
                }
                break;

              case "WOpt":
                break; // TODO: WTF?
              case "PhoneticInfo":
                break;

              case "OleObjectSize":
                break;

              /* Differential Formatting */
              case "DXF":
              case "DXFN":
              case "DXFN12":
              case "DXFN12List":
              case "DXFN12NoCB":
                break;

              /* Data Validation */
              case "Dv":
              case "DVal":
                break;

              /* Data Series */
              case "BRAI":
              case "Series":
              case "SeriesText":
                break;

              /* Data Connection */
              case "DConn":
                break;
              case "DbOrParamQry":
                break;
              case "DBQueryExt":
                break;

              case "OleDbConn":
                break;
              case "ExtString":
                break;

              /* Formatting */
              case "IFmtRecord":
                break;
              case "CondFmt":
              case "CF":
              case "CF12":
              case "CFEx":
                break;

              /* Explicitly Ignored */
              case "Excel9File":
                break;
              case "Units":
                break;
              case "InterfaceHdr":
              case "Mms":
              case "InterfaceEnd":
              case "DSF":
                break;
              case "BuiltInFnGroupCount":
                /* 2.4.30 0x0E or 0x10 but excel 2011 generates 0x11? */ break;
              /* View Stuff */
              case "Window1":
              case "HideObj":
              case "GridSet":
              case "Guts":
              case "UserBView":
              case "UserSViewBegin":
              case "UserSViewEnd":
                break;
              case "Pane":
                break;
              default:
                switch (R.n /* nested */) {
                  /* Chart */
                  case "Dat":
                  case "Begin":
                  case "End":
                  case "StartBlock":
                  case "EndBlock":
                  case "Frame":
                  case "Area":
                  case "Axis":
                  case "AxisLine":
                  case "Tick":
                    break;
                  case "AxesUsed":
                  case "CrtLayout12":
                  case "CrtLayout12A":
                  case "CrtLink":
                  case "CrtLine":
                  case "CrtMlFrt":
                  case "CrtMlFrtContinue":
                    break;
                  case "LineFormat":
                  case "AreaFormat":
                  case "Chart":
                  case "Chart3d":
                  case "Chart3DBarShape":
                  case "ChartFormat":
                  case "ChartFrtInfo":
                    break;
                  case "PlotArea":
                  case "PlotGrowth":
                    break;
                  case "SeriesList":
                  case "SerParent":
                  case "SerAuxTrend":
                    break;
                  case "DataFormat":
                  case "SerToCrt":
                  case "FontX":
                    break;
                  case "CatSerRange":
                  case "AxcExt":
                  case "SerFmt":
                    break;
                  case "ShtProps":
                    break;
                  case "DefaultText":
                  case "Text":
                  case "CatLab":
                    break;
                  case "DataLabExtContents":
                    break;
                  case "Legend":
                  case "LegendException":
                    break;
                  case "Pie":
                  case "Scatter":
                    break;
                  case "PieFormat":
                  case "MarkerFormat":
                    break;
                  case "StartObject":
                  case "EndObject":
                    break;
                  case "AlRuns":
                  case "ObjectLink":
                    break;
                  case "SIIndex":
                    break;
                  case "AttachedLabel":
                  case "YMult":
                    break;

                  /* Chart Group */
                  case "Line":
                  case "Bar":
                    break;
                  case "Surf":
                    break;

                  /* Axis Group */
                  case "AxisParent":
                    break;
                  case "Pos":
                    break;
                  case "ValueRange":
                    break;

                  /* Pivot Chart */
                  case "SXViewEx9":
                    break; // TODO
                  case "SXViewLink":
                    break;
                  case "PivotChartBits":
                    break;
                  case "SBaseRef":
                    break;
                  case "TextPropsStream":
                    break;

                  /* Chart Misc */
                  case "LnExt":
                    break;
                  case "MkrExt":
                    break;
                  case "CrtCoopt":
                    break;

                  /* Query Table */
                  case "Qsi":
                  case "Qsif":
                  case "Qsir":
                  case "QsiSXTag":
                    break;
                  case "TxtQry":
                    break;

                  /* Filter */
                  case "FilterMode":
                    break;
                  case "AutoFilter":
                  case "AutoFilterInfo":
                    break;
                  case "AutoFilter12":
                    break;
                  case "DropDownObjIds":
                    break;
                  case "Sort":
                    break;
                  case "SortData":
                    break;

                  /* Drawing */
                  case "ShapePropsStream":
                    break;
                  case "MsoDrawing":
                  case "MsoDrawingGroup":
                  case "MsoDrawingSelection":
                    break;
                  /* Pub Stuff */
                  case "WebPub":
                  case "AutoWebPub":
                    break;

                  /* Print Stuff */
                  case "HeaderFooter":
                  case "HFPicture":
                  case "PLV":
                  case "HorizontalPageBreaks":
                  case "VerticalPageBreaks":
                    break;
                  /* Behavioral */
                  case "Backup":
                  case "CompressPictures":
                  case "Compat12":
                    break;

                  /* Should not Happen */
                  case "Continue":
                  case "ContinueFrt12":
                    break;

                  /* Future Records */
                  case "FrtFontList":
                  case "FrtWrapper":
                    break;

                  default:
                    switch (R.n /* nested */) {
                      /* BIFF5 records */
                      case "TabIdConf":
                      case "Radar":
                      case "RadarArea":
                      case "DropBar":
                      case "Intl":
                      case "CoordList":
                      case "SerAuxErrBar":
                        break;

                      /* BIFF2-4 records */
                      case "BIFF2FONTCLR":
                      case "BIFF2FMTCNT":
                      case "BIFF2FONTXTRA":
                        break;
                      case "BIFF2XF":
                      case "BIFF3XF":
                      case "BIFF4XF":
                        break;
                      case "BIFF4FMTCNT":
                      case "BIFF2ROW":
                      case "BIFF2WINDOW2":
                        break;

                      /* Miscellaneous */
                      case "SCENARIO":
                      case "DConBin":
                      case "PicF":
                      case "DataLabExt":
                      case "Lel":
                      case "BopPop":
                      case "BopPopCustom":
                      case "RealTimeData":
                      case "Name":
                        break;
                      case "LHNGraph":
                      case "FnGroupName":
                      case "AddMenu":
                      case "LPr":
                        break;
                      case "ListObj":
                      case "ListField":
                        break;
                      case "RRSort":
                        break;
                      case "BigName":
                        break;
                      case "ToolbarHdr":
                      case "ToolbarEnd":
                        break;
                      case "DDEObjName":
                        break;
                      case "FRTArchId$":
                        break;
                      default:
                        if (options.WTF) throw `Unrecognized Record ${R.n}`;
                    }
                }
            }
        }
      } else blob.l += length;
    }
    wb.SheetNames = keys(Directory)
      .sort(function(a, b) {
        return Number(a) - Number(b);
      })
      .map(function(x) {
        return Directory[x].name;
      });
    if (!options.bookSheets) wb.Sheets = Sheets;
    if (wb.Sheets)
      FilterDatabases.forEach(function(r, i) {
        wb.Sheets[wb.SheetNames[i]]["!autofilter"] = r;
      });
    wb.Preamble = Preamble;
    wb.Strings = sst;
    wb.SSF = SSF.get_table();
    if (opts.enc) wb.Encryption = opts.enc;
    if (themes) wb.Themes = themes;
    wb.Metadata = {};
    if (country !== undefined) wb.Metadata.Country = country;
    if (supbooks.names.length > 0) Workbook.Names = supbooks.names;
    wb.Workbook = Workbook;
    return wb;
  }

  /* TODO: split props */
  const PSCLSID = {
    SI: "e0859ff2f94f6810ab9108002b27b3d9",
    DSI: "02d5cdd59c2e1b10939708002b2cf9ae",
    UDI: "05d5cdd59c2e1b10939708002b2cf9ae"
  };
  function parse_xls_props(cfb, props, o) {
    /* [MS-OSHARED] 2.3.3.2.2 Document Summary Information Property Set */
    const DSI = CFB.find(cfb, "!DocumentSummaryInformation");
    if (DSI && DSI.size > 0)
      try {
        const DocSummary = parse_PropertySetStream(
          DSI,
          DocSummaryPIDDSI,
          PSCLSID.DSI
        );
        for (const d in DocSummary) props[d] = DocSummary[d];
      } catch (e) {
        if (o.WTF) throw e; /* empty */
      }

    /* [MS-OSHARED] 2.3.3.2.1 Summary Information Property Set */
    const SI = CFB.find(cfb, "!SummaryInformation");
    if (SI && SI.size > 0)
      try {
        const Summary = parse_PropertySetStream(SI, SummaryPIDSI, PSCLSID.SI);
        for (const s in Summary) if (props[s] == null) props[s] = Summary[s];
      } catch (e) {
        if (o.WTF) throw e; /* empty */
      }

    if (props.HeadingPairs && props.TitlesOfParts) {
      load_props_pairs(props.HeadingPairs, props.TitlesOfParts, props, o);
      delete props.HeadingPairs;
      delete props.TitlesOfParts;
    }
  }

  function parse_xlscfb(cfb, options) {
    if (!options) options = {};
    fix_read_opts(options);
    reset_cp();
    if (options.codepage) set_ansi(options.codepage);
    let CompObj;
    let WB;
    if (cfb.FullPaths) {
      if (CFB.find(cfb, "/encryption"))
        throw new Error("File is password-protected");
      CompObj = CFB.find(cfb, "!CompObj");
      WB = CFB.find(cfb, "/Workbook") || CFB.find(cfb, "/Book");
    } else {
      switch (options.type) {
        case "base64":
          cfb = s2a(Base64.decode(cfb));
          break;
        case "binary":
          cfb = s2a(cfb);
          break;
        case "buffer":
          break;
        case "array":
          if (!Array.isArray(cfb)) cfb = Array.prototype.slice.call(cfb);
          break;
      }
      prep_blob(cfb, 0);
      WB = { content: cfb };
    }
    let WorkbookP;

    let _data;
    if (CompObj) parse_compobj(CompObj);
    if (options.bookProps && !options.bookSheets) WorkbookP = {};
    else {
      const T = has_buf ? "buffer" : "array";
      if (WB && WB.content) WorkbookP = parse_workbook(WB.content, options);
      /* Quattro Pro 7-8 */ else if (
        (_data = CFB.find(cfb, "PerfectOffice_MAIN")) &&
        _data.content
      )
        WorkbookP = WK_.to_workbook(
          _data.content,
          ((options.type = T), options)
        );
      /* Quattro Pro 9 */ else if (
        (_data = CFB.find(cfb, "NativeContent_MAIN")) &&
        _data.content
      )
        WorkbookP = WK_.to_workbook(
          _data.content,
          ((options.type = T), options)
        );
      else throw new Error("Cannot find Workbook stream");
      if (
        options.bookVBA &&
        cfb.FullPaths &&
        CFB.find(cfb, "/_VBA_PROJECT_CUR/VBA/dir")
      )
        WorkbookP.vbaraw = make_vba_xls(cfb);
    }

    const props = {};
    if (cfb.FullPaths) parse_xls_props(cfb, props, options);

    WorkbookP.Props = WorkbookP.Custprops = props; /* TODO: split up properties */
    if (options.bookFiles) WorkbookP.cfb = cfb;
    /* WorkbookP.CompObjP = CompObjP; // TODO: storage? */
    return WorkbookP;
  }

  /* [MS-XLSB] 2.3 Record Enumeration */
  var XLSBRecordEnum = {
    0x0000: { n: "BrtRowHdr", f: parse_BrtRowHdr },
    0x0001: { n: "BrtCellBlank", f: parse_BrtCellBlank },
    0x0002: { n: "BrtCellRk", f: parse_BrtCellRk },
    0x0003: { n: "BrtCellError", f: parse_BrtCellError },
    0x0004: { n: "BrtCellBool", f: parse_BrtCellBool },
    0x0005: { n: "BrtCellReal", f: parse_BrtCellReal },
    0x0006: { n: "BrtCellSt", f: parse_BrtCellSt },
    0x0007: { n: "BrtCellIsst", f: parse_BrtCellIsst },
    0x0008: { n: "BrtFmlaString", f: parse_BrtFmlaString },
    0x0009: { n: "BrtFmlaNum", f: parse_BrtFmlaNum },
    0x000a: { n: "BrtFmlaBool", f: parse_BrtFmlaBool },
    0x000b: { n: "BrtFmlaError", f: parse_BrtFmlaError },
    0x0010: { n: "BrtFRTArchID$", f: parse_BrtFRTArchID$ },
    0x0013: { n: "BrtSSTItem", f: parse_RichStr },
    0x0014: { n: "BrtPCDIMissing" },
    0x0015: { n: "BrtPCDINumber" },
    0x0016: { n: "BrtPCDIBoolean" },
    0x0017: { n: "BrtPCDIError" },
    0x0018: { n: "BrtPCDIString" },
    0x0019: { n: "BrtPCDIDatetime" },
    0x001a: { n: "BrtPCDIIndex" },
    0x001b: { n: "BrtPCDIAMissing" },
    0x001c: { n: "BrtPCDIANumber" },
    0x001d: { n: "BrtPCDIABoolean" },
    0x001e: { n: "BrtPCDIAError" },
    0x001f: { n: "BrtPCDIAString" },
    0x0020: { n: "BrtPCDIADatetime" },
    0x0021: { n: "BrtPCRRecord" },
    0x0022: { n: "BrtPCRRecordDt" },
    0x0023: { n: "BrtFRTBegin" },
    0x0024: { n: "BrtFRTEnd" },
    0x0025: { n: "BrtACBegin" },
    0x0026: { n: "BrtACEnd" },
    0x0027: { n: "BrtName", f: parse_BrtName },
    0x0028: { n: "BrtIndexRowBlock" },
    0x002a: { n: "BrtIndexBlock" },
    0x002b: { n: "BrtFont", f: parse_BrtFont },
    0x002c: { n: "BrtFmt", f: parse_BrtFmt },
    0x002d: { n: "BrtFill", f: parse_BrtFill },
    0x002e: { n: "BrtBorder", f: parse_BrtBorder },
    0x002f: { n: "BrtXF", f: parse_BrtXF },
    0x0030: { n: "BrtStyle" },
    0x0031: { n: "BrtCellMeta" },
    0x0032: { n: "BrtValueMeta" },
    0x0033: { n: "BrtMdb" },
    0x0034: { n: "BrtBeginFmd" },
    0x0035: { n: "BrtEndFmd" },
    0x0036: { n: "BrtBeginMdx" },
    0x0037: { n: "BrtEndMdx" },
    0x0038: { n: "BrtBeginMdxTuple" },
    0x0039: { n: "BrtEndMdxTuple" },
    0x003a: { n: "BrtMdxMbrIstr" },
    0x003b: { n: "BrtStr" },
    0x003c: { n: "BrtColInfo", f: parse_ColInfo },
    0x003e: { n: "BrtCellRString" },
    0x003f: { n: "BrtCalcChainItem$", f: parse_BrtCalcChainItem$ },
    0x0040: { n: "BrtDVal", f: parse_BrtDVal },
    0x0041: { n: "BrtSxvcellNum" },
    0x0042: { n: "BrtSxvcellStr" },
    0x0043: { n: "BrtSxvcellBool" },
    0x0044: { n: "BrtSxvcellErr" },
    0x0045: { n: "BrtSxvcellDate" },
    0x0046: { n: "BrtSxvcellNil" },
    0x0080: { n: "BrtFileVersion" },
    0x0081: { n: "BrtBeginSheet" },
    0x0082: { n: "BrtEndSheet" },
    0x0083: { n: "BrtBeginBook", f: parsenoop, p: 0 },
    0x0084: { n: "BrtEndBook" },
    0x0085: { n: "BrtBeginWsViews" },
    0x0086: { n: "BrtEndWsViews" },
    0x0087: { n: "BrtBeginBookViews" },
    0x0088: { n: "BrtEndBookViews" },
    0x0089: { n: "BrtBeginWsView", f: parse_BrtBeginWsView },
    0x008a: { n: "BrtEndWsView" },
    0x008b: { n: "BrtBeginCsViews" },
    0x008c: { n: "BrtEndCsViews" },
    0x008d: { n: "BrtBeginCsView" },
    0x008e: { n: "BrtEndCsView" },
    0x008f: { n: "BrtBeginBundleShs" },
    0x0090: { n: "BrtEndBundleShs" },
    0x0091: { n: "BrtBeginSheetData" },
    0x0092: { n: "BrtEndSheetData" },
    0x0093: { n: "BrtWsProp", f: parse_BrtWsProp },
    0x0094: { n: "BrtWsDim", f: parse_BrtWsDim, p: 16 },
    0x0097: { n: "BrtPane", f: parse_BrtPane },
    0x0098: { n: "BrtSel" },
    0x0099: { n: "BrtWbProp", f: parse_BrtWbProp },
    0x009a: { n: "BrtWbFactoid" },
    0x009b: { n: "BrtFileRecover" },
    0x009c: { n: "BrtBundleSh", f: parse_BrtBundleSh },
    0x009d: { n: "BrtCalcProp" },
    0x009e: { n: "BrtBookView" },
    0x009f: { n: "BrtBeginSst", f: parse_BrtBeginSst },
    0x00a0: { n: "BrtEndSst" },
    0x00a1: { n: "BrtBeginAFilter", f: parse_UncheckedRfX },
    0x00a2: { n: "BrtEndAFilter" },
    0x00a3: { n: "BrtBeginFilterColumn" },
    0x00a4: { n: "BrtEndFilterColumn" },
    0x00a5: { n: "BrtBeginFilters" },
    0x00a6: { n: "BrtEndFilters" },
    0x00a7: { n: "BrtFilter" },
    0x00a8: { n: "BrtColorFilter" },
    0x00a9: { n: "BrtIconFilter" },
    0x00aa: { n: "BrtTop10Filter" },
    0x00ab: { n: "BrtDynamicFilter" },
    0x00ac: { n: "BrtBeginCustomFilters" },
    0x00ad: { n: "BrtEndCustomFilters" },
    0x00ae: { n: "BrtCustomFilter" },
    0x00af: { n: "BrtAFilterDateGroupItem" },
    0x00b0: { n: "BrtMergeCell", f: parse_BrtMergeCell },
    0x00b1: { n: "BrtBeginMergeCells" },
    0x00b2: { n: "BrtEndMergeCells" },
    0x00b3: { n: "BrtBeginPivotCacheDef" },
    0x00b4: { n: "BrtEndPivotCacheDef" },
    0x00b5: { n: "BrtBeginPCDFields" },
    0x00b6: { n: "BrtEndPCDFields" },
    0x00b7: { n: "BrtBeginPCDField" },
    0x00b8: { n: "BrtEndPCDField" },
    0x00b9: { n: "BrtBeginPCDSource" },
    0x00ba: { n: "BrtEndPCDSource" },
    0x00bb: { n: "BrtBeginPCDSRange" },
    0x00bc: { n: "BrtEndPCDSRange" },
    0x00bd: { n: "BrtBeginPCDFAtbl" },
    0x00be: { n: "BrtEndPCDFAtbl" },
    0x00bf: { n: "BrtBeginPCDIRun" },
    0x00c0: { n: "BrtEndPCDIRun" },
    0x00c1: { n: "BrtBeginPivotCacheRecords" },
    0x00c2: { n: "BrtEndPivotCacheRecords" },
    0x00c3: { n: "BrtBeginPCDHierarchies" },
    0x00c4: { n: "BrtEndPCDHierarchies" },
    0x00c5: { n: "BrtBeginPCDHierarchy" },
    0x00c6: { n: "BrtEndPCDHierarchy" },
    0x00c7: { n: "BrtBeginPCDHFieldsUsage" },
    0x00c8: { n: "BrtEndPCDHFieldsUsage" },
    0x00c9: { n: "BrtBeginExtConnection" },
    0x00ca: { n: "BrtEndExtConnection" },
    0x00cb: { n: "BrtBeginECDbProps" },
    0x00cc: { n: "BrtEndECDbProps" },
    0x00cd: { n: "BrtBeginECOlapProps" },
    0x00ce: { n: "BrtEndECOlapProps" },
    0x00cf: { n: "BrtBeginPCDSConsol" },
    0x00d0: { n: "BrtEndPCDSConsol" },
    0x00d1: { n: "BrtBeginPCDSCPages" },
    0x00d2: { n: "BrtEndPCDSCPages" },
    0x00d3: { n: "BrtBeginPCDSCPage" },
    0x00d4: { n: "BrtEndPCDSCPage" },
    0x00d5: { n: "BrtBeginPCDSCPItem" },
    0x00d6: { n: "BrtEndPCDSCPItem" },
    0x00d7: { n: "BrtBeginPCDSCSets" },
    0x00d8: { n: "BrtEndPCDSCSets" },
    0x00d9: { n: "BrtBeginPCDSCSet" },
    0x00da: { n: "BrtEndPCDSCSet" },
    0x00db: { n: "BrtBeginPCDFGroup" },
    0x00dc: { n: "BrtEndPCDFGroup" },
    0x00dd: { n: "BrtBeginPCDFGItems" },
    0x00de: { n: "BrtEndPCDFGItems" },
    0x00df: { n: "BrtBeginPCDFGRange" },
    0x00e0: { n: "BrtEndPCDFGRange" },
    0x00e1: { n: "BrtBeginPCDFGDiscrete" },
    0x00e2: { n: "BrtEndPCDFGDiscrete" },
    0x00e3: { n: "BrtBeginPCDSDTupleCache" },
    0x00e4: { n: "BrtEndPCDSDTupleCache" },
    0x00e5: { n: "BrtBeginPCDSDTCEntries" },
    0x00e6: { n: "BrtEndPCDSDTCEntries" },
    0x00e7: { n: "BrtBeginPCDSDTCEMembers" },
    0x00e8: { n: "BrtEndPCDSDTCEMembers" },
    0x00e9: { n: "BrtBeginPCDSDTCEMember" },
    0x00ea: { n: "BrtEndPCDSDTCEMember" },
    0x00eb: { n: "BrtBeginPCDSDTCQueries" },
    0x00ec: { n: "BrtEndPCDSDTCQueries" },
    0x00ed: { n: "BrtBeginPCDSDTCQuery" },
    0x00ee: { n: "BrtEndPCDSDTCQuery" },
    0x00ef: { n: "BrtBeginPCDSDTCSets" },
    0x00f0: { n: "BrtEndPCDSDTCSets" },
    0x00f1: { n: "BrtBeginPCDSDTCSet" },
    0x00f2: { n: "BrtEndPCDSDTCSet" },
    0x00f3: { n: "BrtBeginPCDCalcItems" },
    0x00f4: { n: "BrtEndPCDCalcItems" },
    0x00f5: { n: "BrtBeginPCDCalcItem" },
    0x00f6: { n: "BrtEndPCDCalcItem" },
    0x00f7: { n: "BrtBeginPRule" },
    0x00f8: { n: "BrtEndPRule" },
    0x00f9: { n: "BrtBeginPRFilters" },
    0x00fa: { n: "BrtEndPRFilters" },
    0x00fb: { n: "BrtBeginPRFilter" },
    0x00fc: { n: "BrtEndPRFilter" },
    0x00fd: { n: "BrtBeginPNames" },
    0x00fe: { n: "BrtEndPNames" },
    0x00ff: { n: "BrtBeginPName" },
    0x0100: { n: "BrtEndPName" },
    0x0101: { n: "BrtBeginPNPairs" },
    0x0102: { n: "BrtEndPNPairs" },
    0x0103: { n: "BrtBeginPNPair" },
    0x0104: { n: "BrtEndPNPair" },
    0x0105: { n: "BrtBeginECWebProps" },
    0x0106: { n: "BrtEndECWebProps" },
    0x0107: { n: "BrtBeginEcWpTables" },
    0x0108: { n: "BrtEndECWPTables" },
    0x0109: { n: "BrtBeginECParams" },
    0x010a: { n: "BrtEndECParams" },
    0x010b: { n: "BrtBeginECParam" },
    0x010c: { n: "BrtEndECParam" },
    0x010d: { n: "BrtBeginPCDKPIs" },
    0x010e: { n: "BrtEndPCDKPIs" },
    0x010f: { n: "BrtBeginPCDKPI" },
    0x0110: { n: "BrtEndPCDKPI" },
    0x0111: { n: "BrtBeginDims" },
    0x0112: { n: "BrtEndDims" },
    0x0113: { n: "BrtBeginDim" },
    0x0114: { n: "BrtEndDim" },
    0x0115: { n: "BrtIndexPartEnd" },
    0x0116: { n: "BrtBeginStyleSheet" },
    0x0117: { n: "BrtEndStyleSheet" },
    0x0118: { n: "BrtBeginSXView" },
    0x0119: { n: "BrtEndSXVI" },
    0x011a: { n: "BrtBeginSXVI" },
    0x011b: { n: "BrtBeginSXVIs" },
    0x011c: { n: "BrtEndSXVIs" },
    0x011d: { n: "BrtBeginSXVD" },
    0x011e: { n: "BrtEndSXVD" },
    0x011f: { n: "BrtBeginSXVDs" },
    0x0120: { n: "BrtEndSXVDs" },
    0x0121: { n: "BrtBeginSXPI" },
    0x0122: { n: "BrtEndSXPI" },
    0x0123: { n: "BrtBeginSXPIs" },
    0x0124: { n: "BrtEndSXPIs" },
    0x0125: { n: "BrtBeginSXDI" },
    0x0126: { n: "BrtEndSXDI" },
    0x0127: { n: "BrtBeginSXDIs" },
    0x0128: { n: "BrtEndSXDIs" },
    0x0129: { n: "BrtBeginSXLI" },
    0x012a: { n: "BrtEndSXLI" },
    0x012b: { n: "BrtBeginSXLIRws" },
    0x012c: { n: "BrtEndSXLIRws" },
    0x012d: { n: "BrtBeginSXLICols" },
    0x012e: { n: "BrtEndSXLICols" },
    0x012f: { n: "BrtBeginSXFormat" },
    0x0130: { n: "BrtEndSXFormat" },
    0x0131: { n: "BrtBeginSXFormats" },
    0x0132: { n: "BrtEndSxFormats" },
    0x0133: { n: "BrtBeginSxSelect" },
    0x0134: { n: "BrtEndSxSelect" },
    0x0135: { n: "BrtBeginISXVDRws" },
    0x0136: { n: "BrtEndISXVDRws" },
    0x0137: { n: "BrtBeginISXVDCols" },
    0x0138: { n: "BrtEndISXVDCols" },
    0x0139: { n: "BrtEndSXLocation" },
    0x013a: { n: "BrtBeginSXLocation" },
    0x013b: { n: "BrtEndSXView" },
    0x013c: { n: "BrtBeginSXTHs" },
    0x013d: { n: "BrtEndSXTHs" },
    0x013e: { n: "BrtBeginSXTH" },
    0x013f: { n: "BrtEndSXTH" },
    0x0140: { n: "BrtBeginISXTHRws" },
    0x0141: { n: "BrtEndISXTHRws" },
    0x0142: { n: "BrtBeginISXTHCols" },
    0x0143: { n: "BrtEndISXTHCols" },
    0x0144: { n: "BrtBeginSXTDMPS" },
    0x0145: { n: "BrtEndSXTDMPs" },
    0x0146: { n: "BrtBeginSXTDMP" },
    0x0147: { n: "BrtEndSXTDMP" },
    0x0148: { n: "BrtBeginSXTHItems" },
    0x0149: { n: "BrtEndSXTHItems" },
    0x014a: { n: "BrtBeginSXTHItem" },
    0x014b: { n: "BrtEndSXTHItem" },
    0x014c: { n: "BrtBeginMetadata" },
    0x014d: { n: "BrtEndMetadata" },
    0x014e: { n: "BrtBeginEsmdtinfo" },
    0x014f: { n: "BrtMdtinfo" },
    0x0150: { n: "BrtEndEsmdtinfo" },
    0x0151: { n: "BrtBeginEsmdb" },
    0x0152: { n: "BrtEndEsmdb" },
    0x0153: { n: "BrtBeginEsfmd" },
    0x0154: { n: "BrtEndEsfmd" },
    0x0155: { n: "BrtBeginSingleCells" },
    0x0156: { n: "BrtEndSingleCells" },
    0x0157: { n: "BrtBeginList" },
    0x0158: { n: "BrtEndList" },
    0x0159: { n: "BrtBeginListCols" },
    0x015a: { n: "BrtEndListCols" },
    0x015b: { n: "BrtBeginListCol" },
    0x015c: { n: "BrtEndListCol" },
    0x015d: { n: "BrtBeginListXmlCPr" },
    0x015e: { n: "BrtEndListXmlCPr" },
    0x015f: { n: "BrtListCCFmla" },
    0x0160: { n: "BrtListTrFmla" },
    0x0161: { n: "BrtBeginExternals" },
    0x0162: { n: "BrtEndExternals" },
    0x0163: { n: "BrtSupBookSrc", f: parse_RelID },
    0x0165: { n: "BrtSupSelf" },
    0x0166: { n: "BrtSupSame" },
    0x0167: { n: "BrtSupTabs" },
    0x0168: { n: "BrtBeginSupBook" },
    0x0169: { n: "BrtPlaceholderName" },
    0x016a: { n: "BrtExternSheet", f: parse_ExternSheet },
    0x016b: { n: "BrtExternTableStart" },
    0x016c: { n: "BrtExternTableEnd" },
    0x016e: { n: "BrtExternRowHdr" },
    0x016f: { n: "BrtExternCellBlank" },
    0x0170: { n: "BrtExternCellReal" },
    0x0171: { n: "BrtExternCellBool" },
    0x0172: { n: "BrtExternCellError" },
    0x0173: { n: "BrtExternCellString" },
    0x0174: { n: "BrtBeginEsmdx" },
    0x0175: { n: "BrtEndEsmdx" },
    0x0176: { n: "BrtBeginMdxSet" },
    0x0177: { n: "BrtEndMdxSet" },
    0x0178: { n: "BrtBeginMdxMbrProp" },
    0x0179: { n: "BrtEndMdxMbrProp" },
    0x017a: { n: "BrtBeginMdxKPI" },
    0x017b: { n: "BrtEndMdxKPI" },
    0x017c: { n: "BrtBeginEsstr" },
    0x017d: { n: "BrtEndEsstr" },
    0x017e: { n: "BrtBeginPRFItem" },
    0x017f: { n: "BrtEndPRFItem" },
    0x0180: { n: "BrtBeginPivotCacheIDs" },
    0x0181: { n: "BrtEndPivotCacheIDs" },
    0x0182: { n: "BrtBeginPivotCacheID" },
    0x0183: { n: "BrtEndPivotCacheID" },
    0x0184: { n: "BrtBeginISXVIs" },
    0x0185: { n: "BrtEndISXVIs" },
    0x0186: { n: "BrtBeginColInfos" },
    0x0187: { n: "BrtEndColInfos" },
    0x0188: { n: "BrtBeginRwBrk" },
    0x0189: { n: "BrtEndRwBrk" },
    0x018a: { n: "BrtBeginColBrk" },
    0x018b: { n: "BrtEndColBrk" },
    0x018c: { n: "BrtBrk" },
    0x018d: { n: "BrtUserBookView" },
    0x018e: { n: "BrtInfo" },
    0x018f: { n: "BrtCUsr" },
    0x0190: { n: "BrtUsr" },
    0x0191: { n: "BrtBeginUsers" },
    0x0193: { n: "BrtEOF" },
    0x0194: { n: "BrtUCR" },
    0x0195: { n: "BrtRRInsDel" },
    0x0196: { n: "BrtRREndInsDel" },
    0x0197: { n: "BrtRRMove" },
    0x0198: { n: "BrtRREndMove" },
    0x0199: { n: "BrtRRChgCell" },
    0x019a: { n: "BrtRREndChgCell" },
    0x019b: { n: "BrtRRHeader" },
    0x019c: { n: "BrtRRUserView" },
    0x019d: { n: "BrtRRRenSheet" },
    0x019e: { n: "BrtRRInsertSh" },
    0x019f: { n: "BrtRRDefName" },
    0x01a0: { n: "BrtRRNote" },
    0x01a1: { n: "BrtRRConflict" },
    0x01a2: { n: "BrtRRTQSIF" },
    0x01a3: { n: "BrtRRFormat" },
    0x01a4: { n: "BrtRREndFormat" },
    0x01a5: { n: "BrtRRAutoFmt" },
    0x01a6: { n: "BrtBeginUserShViews" },
    0x01a7: { n: "BrtBeginUserShView" },
    0x01a8: { n: "BrtEndUserShView" },
    0x01a9: { n: "BrtEndUserShViews" },
    0x01aa: { n: "BrtArrFmla", f: parse_BrtArrFmla },
    0x01ab: { n: "BrtShrFmla", f: parse_BrtShrFmla },
    0x01ac: { n: "BrtTable" },
    0x01ad: { n: "BrtBeginExtConnections" },
    0x01ae: { n: "BrtEndExtConnections" },
    0x01af: { n: "BrtBeginPCDCalcMems" },
    0x01b0: { n: "BrtEndPCDCalcMems" },
    0x01b1: { n: "BrtBeginPCDCalcMem" },
    0x01b2: { n: "BrtEndPCDCalcMem" },
    0x01b3: { n: "BrtBeginPCDHGLevels" },
    0x01b4: { n: "BrtEndPCDHGLevels" },
    0x01b5: { n: "BrtBeginPCDHGLevel" },
    0x01b6: { n: "BrtEndPCDHGLevel" },
    0x01b7: { n: "BrtBeginPCDHGLGroups" },
    0x01b8: { n: "BrtEndPCDHGLGroups" },
    0x01b9: { n: "BrtBeginPCDHGLGroup" },
    0x01ba: { n: "BrtEndPCDHGLGroup" },
    0x01bb: { n: "BrtBeginPCDHGLGMembers" },
    0x01bc: { n: "BrtEndPCDHGLGMembers" },
    0x01bd: { n: "BrtBeginPCDHGLGMember" },
    0x01be: { n: "BrtEndPCDHGLGMember" },
    0x01bf: { n: "BrtBeginQSI" },
    0x01c0: { n: "BrtEndQSI" },
    0x01c1: { n: "BrtBeginQSIR" },
    0x01c2: { n: "BrtEndQSIR" },
    0x01c3: { n: "BrtBeginDeletedNames" },
    0x01c4: { n: "BrtEndDeletedNames" },
    0x01c5: { n: "BrtBeginDeletedName" },
    0x01c6: { n: "BrtEndDeletedName" },
    0x01c7: { n: "BrtBeginQSIFs" },
    0x01c8: { n: "BrtEndQSIFs" },
    0x01c9: { n: "BrtBeginQSIF" },
    0x01ca: { n: "BrtEndQSIF" },
    0x01cb: { n: "BrtBeginAutoSortScope" },
    0x01cc: { n: "BrtEndAutoSortScope" },
    0x01cd: { n: "BrtBeginConditionalFormatting" },
    0x01ce: { n: "BrtEndConditionalFormatting" },
    0x01cf: { n: "BrtBeginCFRule" },
    0x01d0: { n: "BrtEndCFRule" },
    0x01d1: { n: "BrtBeginIconSet" },
    0x01d2: { n: "BrtEndIconSet" },
    0x01d3: { n: "BrtBeginDatabar" },
    0x01d4: { n: "BrtEndDatabar" },
    0x01d5: { n: "BrtBeginColorScale" },
    0x01d6: { n: "BrtEndColorScale" },
    0x01d7: { n: "BrtCFVO" },
    0x01d8: { n: "BrtExternValueMeta" },
    0x01d9: { n: "BrtBeginColorPalette" },
    0x01da: { n: "BrtEndColorPalette" },
    0x01db: { n: "BrtIndexedColor" },
    0x01dc: { n: "BrtMargins", f: parse_BrtMargins },
    0x01dd: { n: "BrtPrintOptions" },
    0x01de: { n: "BrtPageSetup" },
    0x01df: { n: "BrtBeginHeaderFooter" },
    0x01e0: { n: "BrtEndHeaderFooter" },
    0x01e1: { n: "BrtBeginSXCrtFormat" },
    0x01e2: { n: "BrtEndSXCrtFormat" },
    0x01e3: { n: "BrtBeginSXCrtFormats" },
    0x01e4: { n: "BrtEndSXCrtFormats" },
    0x01e5: { n: "BrtWsFmtInfo", f: parse_BrtWsFmtInfo },
    0x01e6: { n: "BrtBeginMgs" },
    0x01e7: { n: "BrtEndMGs" },
    0x01e8: { n: "BrtBeginMGMaps" },
    0x01e9: { n: "BrtEndMGMaps" },
    0x01ea: { n: "BrtBeginMG" },
    0x01eb: { n: "BrtEndMG" },
    0x01ec: { n: "BrtBeginMap" },
    0x01ed: { n: "BrtEndMap" },
    0x01ee: { n: "BrtHLink", f: parse_BrtHLink },
    0x01ef: { n: "BrtBeginDCon" },
    0x01f0: { n: "BrtEndDCon" },
    0x01f1: { n: "BrtBeginDRefs" },
    0x01f2: { n: "BrtEndDRefs" },
    0x01f3: { n: "BrtDRef" },
    0x01f4: { n: "BrtBeginScenMan" },
    0x01f5: { n: "BrtEndScenMan" },
    0x01f6: { n: "BrtBeginSct" },
    0x01f7: { n: "BrtEndSct" },
    0x01f8: { n: "BrtSlc" },
    0x01f9: { n: "BrtBeginDXFs" },
    0x01fa: { n: "BrtEndDXFs" },
    0x01fb: { n: "BrtDXF" },
    0x01fc: { n: "BrtBeginTableStyles" },
    0x01fd: { n: "BrtEndTableStyles" },
    0x01fe: { n: "BrtBeginTableStyle" },
    0x01ff: { n: "BrtEndTableStyle" },
    0x0200: { n: "BrtTableStyleElement" },
    0x0201: { n: "BrtTableStyleClient" },
    0x0202: { n: "BrtBeginVolDeps" },
    0x0203: { n: "BrtEndVolDeps" },
    0x0204: { n: "BrtBeginVolType" },
    0x0205: { n: "BrtEndVolType" },
    0x0206: { n: "BrtBeginVolMain" },
    0x0207: { n: "BrtEndVolMain" },
    0x0208: { n: "BrtBeginVolTopic" },
    0x0209: { n: "BrtEndVolTopic" },
    0x020a: { n: "BrtVolSubtopic" },
    0x020b: { n: "BrtVolRef" },
    0x020c: { n: "BrtVolNum" },
    0x020d: { n: "BrtVolErr" },
    0x020e: { n: "BrtVolStr" },
    0x020f: { n: "BrtVolBool" },
    0x0210: { n: "BrtBeginCalcChain$" },
    0x0211: { n: "BrtEndCalcChain$" },
    0x0212: { n: "BrtBeginSortState" },
    0x0213: { n: "BrtEndSortState" },
    0x0214: { n: "BrtBeginSortCond" },
    0x0215: { n: "BrtEndSortCond" },
    0x0216: { n: "BrtBookProtection" },
    0x0217: { n: "BrtSheetProtection" },
    0x0218: { n: "BrtRangeProtection" },
    0x0219: { n: "BrtPhoneticInfo" },
    0x021a: { n: "BrtBeginECTxtWiz" },
    0x021b: { n: "BrtEndECTxtWiz" },
    0x021c: { n: "BrtBeginECTWFldInfoLst" },
    0x021d: { n: "BrtEndECTWFldInfoLst" },
    0x021e: { n: "BrtBeginECTwFldInfo" },
    0x0224: { n: "BrtFileSharing" },
    0x0225: { n: "BrtOleSize" },
    0x0226: { n: "BrtDrawing", f: parse_RelID },
    0x0227: { n: "BrtLegacyDrawing" },
    0x0228: { n: "BrtLegacyDrawingHF" },
    0x0229: { n: "BrtWebOpt" },
    0x022a: { n: "BrtBeginWebPubItems" },
    0x022b: { n: "BrtEndWebPubItems" },
    0x022c: { n: "BrtBeginWebPubItem" },
    0x022d: { n: "BrtEndWebPubItem" },
    0x022e: { n: "BrtBeginSXCondFmt" },
    0x022f: { n: "BrtEndSXCondFmt" },
    0x0230: { n: "BrtBeginSXCondFmts" },
    0x0231: { n: "BrtEndSXCondFmts" },
    0x0232: { n: "BrtBkHim" },
    0x0234: { n: "BrtColor" },
    0x0235: { n: "BrtBeginIndexedColors" },
    0x0236: { n: "BrtEndIndexedColors" },
    0x0239: { n: "BrtBeginMRUColors" },
    0x023a: { n: "BrtEndMRUColors" },
    0x023c: { n: "BrtMRUColor" },
    0x023d: { n: "BrtBeginDVals" },
    0x023e: { n: "BrtEndDVals" },
    0x0241: { n: "BrtSupNameStart" },
    0x0242: { n: "BrtSupNameValueStart" },
    0x0243: { n: "BrtSupNameValueEnd" },
    0x0244: { n: "BrtSupNameNum" },
    0x0245: { n: "BrtSupNameErr" },
    0x0246: { n: "BrtSupNameSt" },
    0x0247: { n: "BrtSupNameNil" },
    0x0248: { n: "BrtSupNameBool" },
    0x0249: { n: "BrtSupNameFmla" },
    0x024a: { n: "BrtSupNameBits" },
    0x024b: { n: "BrtSupNameEnd" },
    0x024c: { n: "BrtEndSupBook" },
    0x024d: { n: "BrtCellSmartTagProperty" },
    0x024e: { n: "BrtBeginCellSmartTag" },
    0x024f: { n: "BrtEndCellSmartTag" },
    0x0250: { n: "BrtBeginCellSmartTags" },
    0x0251: { n: "BrtEndCellSmartTags" },
    0x0252: { n: "BrtBeginSmartTags" },
    0x0253: { n: "BrtEndSmartTags" },
    0x0254: { n: "BrtSmartTagType" },
    0x0255: { n: "BrtBeginSmartTagTypes" },
    0x0256: { n: "BrtEndSmartTagTypes" },
    0x0257: { n: "BrtBeginSXFilters" },
    0x0258: { n: "BrtEndSXFilters" },
    0x0259: { n: "BrtBeginSXFILTER" },
    0x025a: { n: "BrtEndSXFilter" },
    0x025b: { n: "BrtBeginFills" },
    0x025c: { n: "BrtEndFills" },
    0x025d: { n: "BrtBeginCellWatches" },
    0x025e: { n: "BrtEndCellWatches" },
    0x025f: { n: "BrtCellWatch" },
    0x0260: { n: "BrtBeginCRErrs" },
    0x0261: { n: "BrtEndCRErrs" },
    0x0262: { n: "BrtCrashRecErr" },
    0x0263: { n: "BrtBeginFonts" },
    0x0264: { n: "BrtEndFonts" },
    0x0265: { n: "BrtBeginBorders" },
    0x0266: { n: "BrtEndBorders" },
    0x0267: { n: "BrtBeginFmts" },
    0x0268: { n: "BrtEndFmts" },
    0x0269: { n: "BrtBeginCellXFs" },
    0x026a: { n: "BrtEndCellXFs" },
    0x026b: { n: "BrtBeginStyles" },
    0x026c: { n: "BrtEndStyles" },
    0x0271: { n: "BrtBigName" },
    0x0272: { n: "BrtBeginCellStyleXFs" },
    0x0273: { n: "BrtEndCellStyleXFs" },
    0x0274: { n: "BrtBeginComments" },
    0x0275: { n: "BrtEndComments" },
    0x0276: { n: "BrtBeginCommentAuthors" },
    0x0277: { n: "BrtEndCommentAuthors" },
    0x0278: { n: "BrtCommentAuthor", f: parse_BrtCommentAuthor },
    0x0279: { n: "BrtBeginCommentList" },
    0x027a: { n: "BrtEndCommentList" },
    0x027b: { n: "BrtBeginComment", f: parse_BrtBeginComment },
    0x027c: { n: "BrtEndComment" },
    0x027d: { n: "BrtCommentText", f: parse_BrtCommentText },
    0x027e: { n: "BrtBeginOleObjects" },
    0x027f: { n: "BrtOleObject" },
    0x0280: { n: "BrtEndOleObjects" },
    0x0281: { n: "BrtBeginSxrules" },
    0x0282: { n: "BrtEndSxRules" },
    0x0283: { n: "BrtBeginActiveXControls" },
    0x0284: { n: "BrtActiveX" },
    0x0285: { n: "BrtEndActiveXControls" },
    0x0286: { n: "BrtBeginPCDSDTCEMembersSortBy" },
    0x0288: { n: "BrtBeginCellIgnoreECs" },
    0x0289: { n: "BrtCellIgnoreEC" },
    0x028a: { n: "BrtEndCellIgnoreECs" },
    0x028b: { n: "BrtCsProp", f: parse_BrtCsProp },
    0x028c: { n: "BrtCsPageSetup" },
    0x028d: { n: "BrtBeginUserCsViews" },
    0x028e: { n: "BrtEndUserCsViews" },
    0x028f: { n: "BrtBeginUserCsView" },
    0x0290: { n: "BrtEndUserCsView" },
    0x0291: { n: "BrtBeginPcdSFCIEntries" },
    0x0292: { n: "BrtEndPCDSFCIEntries" },
    0x0293: { n: "BrtPCDSFCIEntry" },
    0x0294: { n: "BrtBeginListParts" },
    0x0295: { n: "BrtListPart" },
    0x0296: { n: "BrtEndListParts" },
    0x0297: { n: "BrtSheetCalcProp" },
    0x0298: { n: "BrtBeginFnGroup" },
    0x0299: { n: "BrtFnGroup" },
    0x029a: { n: "BrtEndFnGroup" },
    0x029b: { n: "BrtSupAddin" },
    0x029c: { n: "BrtSXTDMPOrder" },
    0x029d: { n: "BrtCsProtection" },
    0x029f: { n: "BrtBeginWsSortMap" },
    0x02a0: { n: "BrtEndWsSortMap" },
    0x02a1: { n: "BrtBeginRRSort" },
    0x02a2: { n: "BrtEndRRSort" },
    0x02a3: { n: "BrtRRSortItem" },
    0x02a4: { n: "BrtFileSharingIso" },
    0x02a5: { n: "BrtBookProtectionIso" },
    0x02a6: { n: "BrtSheetProtectionIso" },
    0x02a7: { n: "BrtCsProtectionIso" },
    0x02a8: { n: "BrtRangeProtectionIso" },
    0x02a9: { n: "BrtDValList" },
    0x0400: { n: "BrtRwDescent" },
    0x0401: { n: "BrtKnownFonts" },
    0x0402: { n: "BrtBeginSXTupleSet" },
    0x0403: { n: "BrtEndSXTupleSet" },
    0x0404: { n: "BrtBeginSXTupleSetHeader" },
    0x0405: { n: "BrtEndSXTupleSetHeader" },
    0x0406: { n: "BrtSXTupleSetHeaderItem" },
    0x0407: { n: "BrtBeginSXTupleSetData" },
    0x0408: { n: "BrtEndSXTupleSetData" },
    0x0409: { n: "BrtBeginSXTupleSetRow" },
    0x040a: { n: "BrtEndSXTupleSetRow" },
    0x040b: { n: "BrtSXTupleSetRowItem" },
    0x040c: { n: "BrtNameExt" },
    0x040d: { n: "BrtPCDH14" },
    0x040e: { n: "BrtBeginPCDCalcMem14" },
    0x040f: { n: "BrtEndPCDCalcMem14" },
    0x0410: { n: "BrtSXTH14" },
    0x0411: { n: "BrtBeginSparklineGroup" },
    0x0412: { n: "BrtEndSparklineGroup" },
    0x0413: { n: "BrtSparkline" },
    0x0414: { n: "BrtSXDI14" },
    0x0415: { n: "BrtWsFmtInfoEx14" },
    0x0416: { n: "BrtBeginConditionalFormatting14" },
    0x0417: { n: "BrtEndConditionalFormatting14" },
    0x0418: { n: "BrtBeginCFRule14" },
    0x0419: { n: "BrtEndCFRule14" },
    0x041a: { n: "BrtCFVO14" },
    0x041b: { n: "BrtBeginDatabar14" },
    0x041c: { n: "BrtBeginIconSet14" },
    0x041d: { n: "BrtDVal14", f: parse_BrtDVal14 },
    0x041e: { n: "BrtBeginDVals14" },
    0x041f: { n: "BrtColor14" },
    0x0420: { n: "BrtBeginSparklines" },
    0x0421: { n: "BrtEndSparklines" },
    0x0422: { n: "BrtBeginSparklineGroups" },
    0x0423: { n: "BrtEndSparklineGroups" },
    0x0425: { n: "BrtSXVD14" },
    0x0426: { n: "BrtBeginSXView14" },
    0x0427: { n: "BrtEndSXView14" },
    0x0428: { n: "BrtBeginSXView16" },
    0x0429: { n: "BrtEndSXView16" },
    0x042a: { n: "BrtBeginPCD14" },
    0x042b: { n: "BrtEndPCD14" },
    0x042c: { n: "BrtBeginExtConn14" },
    0x042d: { n: "BrtEndExtConn14" },
    0x042e: { n: "BrtBeginSlicerCacheIDs" },
    0x042f: { n: "BrtEndSlicerCacheIDs" },
    0x0430: { n: "BrtBeginSlicerCacheID" },
    0x0431: { n: "BrtEndSlicerCacheID" },
    0x0433: { n: "BrtBeginSlicerCache" },
    0x0434: { n: "BrtEndSlicerCache" },
    0x0435: { n: "BrtBeginSlicerCacheDef" },
    0x0436: { n: "BrtEndSlicerCacheDef" },
    0x0437: { n: "BrtBeginSlicersEx" },
    0x0438: { n: "BrtEndSlicersEx" },
    0x0439: { n: "BrtBeginSlicerEx" },
    0x043a: { n: "BrtEndSlicerEx" },
    0x043b: { n: "BrtBeginSlicer" },
    0x043c: { n: "BrtEndSlicer" },
    0x043d: { n: "BrtSlicerCachePivotTables" },
    0x043e: { n: "BrtBeginSlicerCacheOlapImpl" },
    0x043f: { n: "BrtEndSlicerCacheOlapImpl" },
    0x0440: { n: "BrtBeginSlicerCacheLevelsData" },
    0x0441: { n: "BrtEndSlicerCacheLevelsData" },
    0x0442: { n: "BrtBeginSlicerCacheLevelData" },
    0x0443: { n: "BrtEndSlicerCacheLevelData" },
    0x0444: { n: "BrtBeginSlicerCacheSiRanges" },
    0x0445: { n: "BrtEndSlicerCacheSiRanges" },
    0x0446: { n: "BrtBeginSlicerCacheSiRange" },
    0x0447: { n: "BrtEndSlicerCacheSiRange" },
    0x0448: { n: "BrtSlicerCacheOlapItem" },
    0x0449: { n: "BrtBeginSlicerCacheSelections" },
    0x044a: { n: "BrtSlicerCacheSelection" },
    0x044b: { n: "BrtEndSlicerCacheSelections" },
    0x044c: { n: "BrtBeginSlicerCacheNative" },
    0x044d: { n: "BrtEndSlicerCacheNative" },
    0x044e: { n: "BrtSlicerCacheNativeItem" },
    0x044f: { n: "BrtRangeProtection14" },
    0x0450: { n: "BrtRangeProtectionIso14" },
    0x0451: { n: "BrtCellIgnoreEC14" },
    0x0457: { n: "BrtList14" },
    0x0458: { n: "BrtCFIcon" },
    0x0459: { n: "BrtBeginSlicerCachesPivotCacheIDs" },
    0x045a: { n: "BrtEndSlicerCachesPivotCacheIDs" },
    0x045b: { n: "BrtBeginSlicers" },
    0x045c: { n: "BrtEndSlicers" },
    0x045d: { n: "BrtWbProp14" },
    0x045e: { n: "BrtBeginSXEdit" },
    0x045f: { n: "BrtEndSXEdit" },
    0x0460: { n: "BrtBeginSXEdits" },
    0x0461: { n: "BrtEndSXEdits" },
    0x0462: { n: "BrtBeginSXChange" },
    0x0463: { n: "BrtEndSXChange" },
    0x0464: { n: "BrtBeginSXChanges" },
    0x0465: { n: "BrtEndSXChanges" },
    0x0466: { n: "BrtSXTupleItems" },
    0x0468: { n: "BrtBeginSlicerStyle" },
    0x0469: { n: "BrtEndSlicerStyle" },
    0x046a: { n: "BrtSlicerStyleElement" },
    0x046b: { n: "BrtBeginStyleSheetExt14" },
    0x046c: { n: "BrtEndStyleSheetExt14" },
    0x046d: { n: "BrtBeginSlicerCachesPivotCacheID" },
    0x046e: { n: "BrtEndSlicerCachesPivotCacheID" },
    0x046f: { n: "BrtBeginConditionalFormattings" },
    0x0470: { n: "BrtEndConditionalFormattings" },
    0x0471: { n: "BrtBeginPCDCalcMemExt" },
    0x0472: { n: "BrtEndPCDCalcMemExt" },
    0x0473: { n: "BrtBeginPCDCalcMemsExt" },
    0x0474: { n: "BrtEndPCDCalcMemsExt" },
    0x0475: { n: "BrtPCDField14" },
    0x0476: { n: "BrtBeginSlicerStyles" },
    0x0477: { n: "BrtEndSlicerStyles" },
    0x0478: { n: "BrtBeginSlicerStyleElements" },
    0x0479: { n: "BrtEndSlicerStyleElements" },
    0x047a: { n: "BrtCFRuleExt" },
    0x047b: { n: "BrtBeginSXCondFmt14" },
    0x047c: { n: "BrtEndSXCondFmt14" },
    0x047d: { n: "BrtBeginSXCondFmts14" },
    0x047e: { n: "BrtEndSXCondFmts14" },
    0x0480: { n: "BrtBeginSortCond14" },
    0x0481: { n: "BrtEndSortCond14" },
    0x0482: { n: "BrtEndDVals14" },
    0x0483: { n: "BrtEndIconSet14" },
    0x0484: { n: "BrtEndDatabar14" },
    0x0485: { n: "BrtBeginColorScale14" },
    0x0486: { n: "BrtEndColorScale14" },
    0x0487: { n: "BrtBeginSxrules14" },
    0x0488: { n: "BrtEndSxrules14" },
    0x0489: { n: "BrtBeginPRule14" },
    0x048a: { n: "BrtEndPRule14" },
    0x048b: { n: "BrtBeginPRFilters14" },
    0x048c: { n: "BrtEndPRFilters14" },
    0x048d: { n: "BrtBeginPRFilter14" },
    0x048e: { n: "BrtEndPRFilter14" },
    0x048f: { n: "BrtBeginPRFItem14" },
    0x0490: { n: "BrtEndPRFItem14" },
    0x0491: { n: "BrtBeginCellIgnoreECs14" },
    0x0492: { n: "BrtEndCellIgnoreECs14" },
    0x0493: { n: "BrtDxf14" },
    0x0494: { n: "BrtBeginDxF14s" },
    0x0495: { n: "BrtEndDxf14s" },
    0x0499: { n: "BrtFilter14" },
    0x049a: { n: "BrtBeginCustomFilters14" },
    0x049c: { n: "BrtCustomFilter14" },
    0x049d: { n: "BrtIconFilter14" },
    0x049e: { n: "BrtPivotCacheConnectionName" },
    0x0800: { n: "BrtBeginDecoupledPivotCacheIDs" },
    0x0801: { n: "BrtEndDecoupledPivotCacheIDs" },
    0x0802: { n: "BrtDecoupledPivotCacheID" },
    0x0803: { n: "BrtBeginPivotTableRefs" },
    0x0804: { n: "BrtEndPivotTableRefs" },
    0x0805: { n: "BrtPivotTableRef" },
    0x0806: { n: "BrtSlicerCacheBookPivotTables" },
    0x0807: { n: "BrtBeginSxvcells" },
    0x0808: { n: "BrtEndSxvcells" },
    0x0809: { n: "BrtBeginSxRow" },
    0x080a: { n: "BrtEndSxRow" },
    0x080c: { n: "BrtPcdCalcMem15" },
    0x0813: { n: "BrtQsi15" },
    0x0814: { n: "BrtBeginWebExtensions" },
    0x0815: { n: "BrtEndWebExtensions" },
    0x0816: { n: "BrtWebExtension" },
    0x0817: { n: "BrtAbsPath15" },
    0x0818: { n: "BrtBeginPivotTableUISettings" },
    0x0819: { n: "BrtEndPivotTableUISettings" },
    0x081b: { n: "BrtTableSlicerCacheIDs" },
    0x081c: { n: "BrtTableSlicerCacheID" },
    0x081d: { n: "BrtBeginTableSlicerCache" },
    0x081e: { n: "BrtEndTableSlicerCache" },
    0x081f: { n: "BrtSxFilter15" },
    0x0820: { n: "BrtBeginTimelineCachePivotCacheIDs" },
    0x0821: { n: "BrtEndTimelineCachePivotCacheIDs" },
    0x0822: { n: "BrtTimelineCachePivotCacheID" },
    0x0823: { n: "BrtBeginTimelineCacheIDs" },
    0x0824: { n: "BrtEndTimelineCacheIDs" },
    0x0825: { n: "BrtBeginTimelineCacheID" },
    0x0826: { n: "BrtEndTimelineCacheID" },
    0x0827: { n: "BrtBeginTimelinesEx" },
    0x0828: { n: "BrtEndTimelinesEx" },
    0x0829: { n: "BrtBeginTimelineEx" },
    0x082a: { n: "BrtEndTimelineEx" },
    0x082b: { n: "BrtWorkBookPr15" },
    0x082c: { n: "BrtPCDH15" },
    0x082d: { n: "BrtBeginTimelineStyle" },
    0x082e: { n: "BrtEndTimelineStyle" },
    0x082f: { n: "BrtTimelineStyleElement" },
    0x0830: { n: "BrtBeginTimelineStylesheetExt15" },
    0x0831: { n: "BrtEndTimelineStylesheetExt15" },
    0x0832: { n: "BrtBeginTimelineStyles" },
    0x0833: { n: "BrtEndTimelineStyles" },
    0x0834: { n: "BrtBeginTimelineStyleElements" },
    0x0835: { n: "BrtEndTimelineStyleElements" },
    0x0836: { n: "BrtDxf15" },
    0x0837: { n: "BrtBeginDxfs15" },
    0x0838: { n: "brtEndDxfs15" },
    0x0839: { n: "BrtSlicerCacheHideItemsWithNoData" },
    0x083a: { n: "BrtBeginItemUniqueNames" },
    0x083b: { n: "BrtEndItemUniqueNames" },
    0x083c: { n: "BrtItemUniqueName" },
    0x083d: { n: "BrtBeginExtConn15" },
    0x083e: { n: "BrtEndExtConn15" },
    0x083f: { n: "BrtBeginOledbPr15" },
    0x0840: { n: "BrtEndOledbPr15" },
    0x0841: { n: "BrtBeginDataFeedPr15" },
    0x0842: { n: "BrtEndDataFeedPr15" },
    0x0843: { n: "BrtTextPr15" },
    0x0844: { n: "BrtRangePr15" },
    0x0845: { n: "BrtDbCommand15" },
    0x0846: { n: "BrtBeginDbTables15" },
    0x0847: { n: "BrtEndDbTables15" },
    0x0848: { n: "BrtDbTable15" },
    0x0849: { n: "BrtBeginDataModel" },
    0x084a: { n: "BrtEndDataModel" },
    0x084b: { n: "BrtBeginModelTables" },
    0x084c: { n: "BrtEndModelTables" },
    0x084d: { n: "BrtModelTable" },
    0x084e: { n: "BrtBeginModelRelationships" },
    0x084f: { n: "BrtEndModelRelationships" },
    0x0850: { n: "BrtModelRelationship" },
    0x0851: { n: "BrtBeginECTxtWiz15" },
    0x0852: { n: "BrtEndECTxtWiz15" },
    0x0853: { n: "BrtBeginECTWFldInfoLst15" },
    0x0854: { n: "BrtEndECTWFldInfoLst15" },
    0x0855: { n: "BrtBeginECTWFldInfo15" },
    0x0856: { n: "BrtFieldListActiveItem" },
    0x0857: { n: "BrtPivotCacheIdVersion" },
    0x0858: { n: "BrtSXDI15" },
    0x0859: { n: "BrtBeginModelTimeGroupings" },
    0x085a: { n: "BrtEndModelTimeGroupings" },
    0x085b: { n: "BrtBeginModelTimeGrouping" },
    0x085c: { n: "BrtEndModelTimeGrouping" },
    0x085d: { n: "BrtModelTimeGroupingCalcCol" },
    0x0c00: { n: "BrtUid" },
    0x0c01: { n: "BrtRevisionPtr" },
    0x13e7: { n: "BrtBeginCalcFeatures" },
    0x13e8: { n: "BrtEndCalcFeatures" },
    0x13e9: { n: "BrtCalcFeature" },
    0xffff: { n: "" }
  };

  var XLSBRE = evert_key(XLSBRecordEnum, "n");

  /* [MS-XLS] 2.3 Record Enumeration */
  var XLSRecordEnum = {
    0x0003: { n: "BIFF2NUM", f: parse_BIFF2NUM },
    0x0004: { n: "BIFF2STR", f: parse_BIFF2STR },
    0x0006: { n: "Formula", f: parse_Formula },
    0x0009: { n: "BOF", f: parse_BOF },
    0x000a: { n: "EOF", f: parsenoop2 },
    0x000c: { n: "CalcCount", f: parseuint16 },
    0x000d: { n: "CalcMode", f: parseuint16 },
    0x000e: { n: "CalcPrecision", f: parsebool },
    0x000f: { n: "CalcRefMode", f: parsebool },
    0x0010: { n: "CalcDelta", f: parse_Xnum },
    0x0011: { n: "CalcIter", f: parsebool },
    0x0012: { n: "Protect", f: parsebool },
    0x0013: { n: "Password", f: parseuint16 },
    0x0014: { n: "Header", f: parse_XLHeaderFooter },
    0x0015: { n: "Footer", f: parse_XLHeaderFooter },
    0x0017: { n: "ExternSheet", f: parse_ExternSheet },
    0x0018: { n: "Lbl", f: parse_Lbl },
    0x0019: { n: "WinProtect", f: parsebool },
    0x001a: { n: "VerticalPageBreaks" },
    0x001b: { n: "HorizontalPageBreaks" },
    0x001c: { n: "Note", f: parse_Note },
    0x001d: { n: "Selection" },
    0x0022: { n: "Date1904", f: parsebool },
    0x0023: { n: "ExternName", f: parse_ExternName },
    0x0024: { n: "COLWIDTH" },
    0x0026: { n: "LeftMargin", f: parse_Xnum },
    0x0027: { n: "RightMargin", f: parse_Xnum },
    0x0028: { n: "TopMargin", f: parse_Xnum },
    0x0029: { n: "BottomMargin", f: parse_Xnum },
    0x002a: { n: "PrintRowCol", f: parsebool },
    0x002b: { n: "PrintGrid", f: parsebool },
    0x002f: { n: "FilePass", f: parse_FilePass },
    0x0031: { n: "Font", f: parse_Font },
    0x0033: { n: "PrintSize", f: parseuint16 },
    0x003c: { n: "Continue" },
    0x003d: { n: "Window1", f: parse_Window1 },
    0x0040: { n: "Backup", f: parsebool },
    0x0041: { n: "Pane", f: parse_Pane },
    0x0042: { n: "CodePage", f: parseuint16 },
    0x004d: { n: "Pls" },
    0x0050: { n: "DCon" },
    0x0051: { n: "DConRef" },
    0x0052: { n: "DConName" },
    0x0055: { n: "DefColWidth", f: parseuint16 },
    0x0059: { n: "XCT" },
    0x005a: { n: "CRN" },
    0x005b: { n: "FileSharing" },
    0x005c: { n: "WriteAccess", f: parse_WriteAccess },
    0x005d: { n: "Obj", f: parse_Obj },
    0x005e: { n: "Uncalced" },
    0x005f: { n: "CalcSaveRecalc", f: parsebool },
    0x0060: { n: "Template" },
    0x0061: { n: "Intl" },
    0x0063: { n: "ObjProtect", f: parsebool },
    0x007d: { n: "ColInfo", f: parse_ColInfo },
    0x0080: { n: "Guts", f: parse_Guts },
    0x0081: { n: "WsBool", f: parse_WsBool },
    0x0082: { n: "GridSet", f: parseuint16 },
    0x0083: { n: "HCenter", f: parsebool },
    0x0084: { n: "VCenter", f: parsebool },
    0x0085: { n: "BoundSheet8", f: parse_BoundSheet8 },
    0x0086: { n: "WriteProtect" },
    0x008c: { n: "Country", f: parse_Country },
    0x008d: { n: "HideObj", f: parseuint16 },
    0x0090: { n: "Sort" },
    0x0092: { n: "Palette", f: parse_Palette },
    0x0097: { n: "Sync" },
    0x0098: { n: "LPr" },
    0x0099: { n: "DxGCol" },
    0x009a: { n: "FnGroupName" },
    0x009b: { n: "FilterMode" },
    0x009c: { n: "BuiltInFnGroupCount", f: parseuint16 },
    0x009d: { n: "AutoFilterInfo" },
    0x009e: { n: "AutoFilter" },
    0x00a0: { n: "Scl", f: parse_Scl },
    0x00a1: { n: "Setup", f: parse_Setup },
    0x00ae: { n: "ScenMan" },
    0x00af: { n: "SCENARIO" },
    0x00b0: { n: "SxView" },
    0x00b1: { n: "Sxvd" },
    0x00b2: { n: "SXVI" },
    0x00b4: { n: "SxIvd" },
    0x00b5: { n: "SXLI" },
    0x00b6: { n: "SXPI" },
    0x00b8: { n: "DocRoute" },
    0x00b9: { n: "RecipName" },
    0x00bd: { n: "MulRk", f: parse_MulRk },
    0x00be: { n: "MulBlank", f: parse_MulBlank },
    0x00c1: { n: "Mms", f: parsenoop2 },
    0x00c5: { n: "SXDI" },
    0x00c6: { n: "SXDB" },
    0x00c7: { n: "SXFDB" },
    0x00c8: { n: "SXDBB" },
    0x00c9: { n: "SXNum" },
    0x00ca: { n: "SxBool", f: parsebool },
    0x00cb: { n: "SxErr" },
    0x00cc: { n: "SXInt" },
    0x00cd: { n: "SXString" },
    0x00ce: { n: "SXDtr" },
    0x00cf: { n: "SxNil" },
    0x00d0: { n: "SXTbl" },
    0x00d1: { n: "SXTBRGIITM" },
    0x00d2: { n: "SxTbpg" },
    0x00d3: { n: "ObProj" },
    0x00d5: { n: "SXStreamID" },
    0x00d7: { n: "DBCell" },
    0x00d8: { n: "SXRng" },
    0x00d9: { n: "SxIsxoper" },
    0x00da: { n: "BookBool", f: parseuint16 },
    0x00dc: { n: "DbOrParamQry" },
    0x00dd: { n: "ScenarioProtect", f: parsebool },
    0x00de: { n: "OleObjectSize" },
    0x00e0: { n: "XF", f: parse_XF },
    0x00e1: { n: "InterfaceHdr", f: parse_InterfaceHdr },
    0x00e2: { n: "InterfaceEnd", f: parsenoop2 },
    0x00e3: { n: "SXVS" },
    0x00e5: { n: "MergeCells", f: parse_MergeCells },
    0x00e9: { n: "BkHim" },
    0x00eb: { n: "MsoDrawingGroup" },
    0x00ec: { n: "MsoDrawing" },
    0x00ed: { n: "MsoDrawingSelection" },
    0x00ef: { n: "PhoneticInfo" },
    0x00f0: { n: "SxRule" },
    0x00f1: { n: "SXEx" },
    0x00f2: { n: "SxFilt" },
    0x00f4: { n: "SxDXF" },
    0x00f5: { n: "SxItm" },
    0x00f6: { n: "SxName" },
    0x00f7: { n: "SxSelect" },
    0x00f8: { n: "SXPair" },
    0x00f9: { n: "SxFmla" },
    0x00fb: { n: "SxFormat" },
    0x00fc: { n: "SST", f: parse_SST },
    0x00fd: { n: "LabelSst", f: parse_LabelSst },
    0x00ff: { n: "ExtSST", f: parse_ExtSST },
    0x0100: { n: "SXVDEx" },
    0x0103: { n: "SXFormula" },
    0x0122: { n: "SXDBEx" },
    0x0137: { n: "RRDInsDel" },
    0x0138: { n: "RRDHead" },
    0x013b: { n: "RRDChgCell" },
    0x013d: { n: "RRTabId", f: parseuint16a },
    0x013e: { n: "RRDRenSheet" },
    0x013f: { n: "RRSort" },
    0x0140: { n: "RRDMove" },
    0x014a: { n: "RRFormat" },
    0x014b: { n: "RRAutoFmt" },
    0x014d: { n: "RRInsertSh" },
    0x014e: { n: "RRDMoveBegin" },
    0x014f: { n: "RRDMoveEnd" },
    0x0150: { n: "RRDInsDelBegin" },
    0x0151: { n: "RRDInsDelEnd" },
    0x0152: { n: "RRDConflict" },
    0x0153: { n: "RRDDefName" },
    0x0154: { n: "RRDRstEtxp" },
    0x015f: { n: "LRng" },
    0x0160: { n: "UsesELFs", f: parsebool },
    0x0161: { n: "DSF", f: parsenoop2 },
    0x0191: { n: "CUsr" },
    0x0192: { n: "CbUsr" },
    0x0193: { n: "UsrInfo" },
    0x0194: { n: "UsrExcl" },
    0x0195: { n: "FileLock" },
    0x0196: { n: "RRDInfo" },
    0x0197: { n: "BCUsrs" },
    0x0198: { n: "UsrChk" },
    0x01a9: { n: "UserBView" },
    0x01aa: { n: "UserSViewBegin" },
    0x01ab: { n: "UserSViewEnd" },
    0x01ac: { n: "RRDUserView" },
    0x01ad: { n: "Qsi" },
    0x01ae: { n: "SupBook", f: parse_SupBook },
    0x01af: { n: "Prot4Rev", f: parsebool },
    0x01b0: { n: "CondFmt" },
    0x01b1: { n: "CF" },
    0x01b2: { n: "DVal" },
    0x01b5: { n: "DConBin" },
    0x01b6: { n: "TxO", f: parse_TxO },
    0x01b7: { n: "RefreshAll", f: parsebool },
    0x01b8: { n: "HLink", f: parse_HLink },
    0x01b9: { n: "Lel" },
    0x01ba: { n: "CodeName", f: parse_XLUnicodeString },
    0x01bb: { n: "SXFDBType" },
    0x01bc: { n: "Prot4RevPass", f: parseuint16 },
    0x01bd: { n: "ObNoMacros" },
    0x01be: { n: "Dv" },
    0x01c0: { n: "Excel9File", f: parsenoop2 },
    0x01c1: { n: "RecalcId", f: parse_RecalcId, r: 2 },
    0x01c2: { n: "EntExU2", f: parsenoop2 },
    0x0200: { n: "Dimensions", f: parse_Dimensions },
    0x0201: { n: "Blank", f: parse_Blank },
    0x0203: { n: "Number", f: parse_Number },
    0x0204: { n: "Label", f: parse_Label },
    0x0205: { n: "BoolErr", f: parse_BoolErr },
    0x0206: { n: "Formula", f: parse_Formula },
    0x0207: { n: "String", f: parse_String },
    0x0208: { n: "Row", f: parse_Row },
    0x020b: { n: "Index" },
    0x0221: { n: "Array", f: parse_Array },
    0x0225: { n: "DefaultRowHeight", f: parse_DefaultRowHeight },
    0x0236: { n: "Table" },
    0x023e: { n: "Window2", f: parse_Window2 },
    0x027e: { n: "RK", f: parse_RK },
    0x0293: { n: "Style" },
    0x0406: { n: "Formula", f: parse_Formula },
    0x0418: { n: "BigName" },
    0x041e: { n: "Format", f: parse_Format },
    0x043c: { n: "ContinueBigName" },
    0x04bc: { n: "ShrFmla", f: parse_ShrFmla },
    0x0800: { n: "HLinkTooltip", f: parse_HLinkTooltip },
    0x0801: { n: "WebPub" },
    0x0802: { n: "QsiSXTag" },
    0x0803: { n: "DBQueryExt" },
    0x0804: { n: "ExtString" },
    0x0805: { n: "TxtQry" },
    0x0806: { n: "Qsir" },
    0x0807: { n: "Qsif" },
    0x0808: { n: "RRDTQSIF" },
    0x0809: { n: "BOF", f: parse_BOF },
    0x080a: { n: "OleDbConn" },
    0x080b: { n: "WOpt" },
    0x080c: { n: "SXViewEx" },
    0x080d: { n: "SXTH" },
    0x080e: { n: "SXPIEx" },
    0x080f: { n: "SXVDTEx" },
    0x0810: { n: "SXViewEx9" },
    0x0812: { n: "ContinueFrt" },
    0x0813: { n: "RealTimeData" },
    0x0850: { n: "ChartFrtInfo" },
    0x0851: { n: "FrtWrapper" },
    0x0852: { n: "StartBlock" },
    0x0853: { n: "EndBlock" },
    0x0854: { n: "StartObject" },
    0x0855: { n: "EndObject" },
    0x0856: { n: "CatLab" },
    0x0857: { n: "YMult" },
    0x0858: { n: "SXViewLink" },
    0x0859: { n: "PivotChartBits" },
    0x085a: { n: "FrtFontList" },
    0x0862: { n: "SheetExt" },
    0x0863: { n: "BookExt", r: 12 },
    0x0864: { n: "SXAddl" },
    0x0865: { n: "CrErr" },
    0x0866: { n: "HFPicture" },
    0x0867: { n: "FeatHdr", f: parsenoop2 },
    0x0868: { n: "Feat" },
    0x086a: { n: "DataLabExt" },
    0x086b: { n: "DataLabExtContents" },
    0x086c: { n: "CellWatch" },
    0x0871: { n: "FeatHdr11" },
    0x0872: { n: "Feature11" },
    0x0874: { n: "DropDownObjIds" },
    0x0875: { n: "ContinueFrt11" },
    0x0876: { n: "DConn" },
    0x0877: { n: "List12" },
    0x0878: { n: "Feature12" },
    0x0879: { n: "CondFmt12" },
    0x087a: { n: "CF12" },
    0x087b: { n: "CFEx" },
    0x087c: { n: "XFCRC", f: parse_XFCRC, r: 12 },
    0x087d: { n: "XFExt", f: parse_XFExt, r: 12 },
    0x087e: { n: "AutoFilter12" },
    0x087f: { n: "ContinueFrt12" },
    0x0884: { n: "MDTInfo" },
    0x0885: { n: "MDXStr" },
    0x0886: { n: "MDXTuple" },
    0x0887: { n: "MDXSet" },
    0x0888: { n: "MDXProp" },
    0x0889: { n: "MDXKPI" },
    0x088a: { n: "MDB" },
    0x088b: { n: "PLV" },
    0x088c: { n: "Compat12", f: parsebool, r: 12 },
    0x088d: { n: "DXF" },
    0x088e: { n: "TableStyles", r: 12 },
    0x088f: { n: "TableStyle" },
    0x0890: { n: "TableStyleElement" },
    0x0892: { n: "StyleExt" },
    0x0893: { n: "NamePublish" },
    0x0894: { n: "NameCmt", f: parse_NameCmt, r: 12 },
    0x0895: { n: "SortData" },
    0x0896: { n: "Theme", f: parse_Theme, r: 12 },
    0x0897: { n: "GUIDTypeLib" },
    0x0898: { n: "FnGrp12" },
    0x0899: { n: "NameFnGrp12" },
    0x089a: { n: "MTRSettings", f: parse_MTRSettings, r: 12 },
    0x089b: { n: "CompressPictures", f: parsenoop2 },
    0x089c: { n: "HeaderFooter" },
    0x089d: { n: "CrtLayout12" },
    0x089e: { n: "CrtMlFrt" },
    0x089f: { n: "CrtMlFrtContinue" },
    0x08a3: { n: "ForceFullCalculation", f: parse_ForceFullCalculation },
    0x08a4: { n: "ShapePropsStream" },
    0x08a5: { n: "TextPropsStream" },
    0x08a6: { n: "RichTextStream" },
    0x08a7: { n: "CrtLayout12A" },
    0x1001: { n: "Units" },
    0x1002: { n: "Chart" },
    0x1003: { n: "Series" },
    0x1006: { n: "DataFormat" },
    0x1007: { n: "LineFormat" },
    0x1009: { n: "MarkerFormat" },
    0x100a: { n: "AreaFormat" },
    0x100b: { n: "PieFormat" },
    0x100c: { n: "AttachedLabel" },
    0x100d: { n: "SeriesText" },
    0x1014: { n: "ChartFormat" },
    0x1015: { n: "Legend" },
    0x1016: { n: "SeriesList" },
    0x1017: { n: "Bar" },
    0x1018: { n: "Line" },
    0x1019: { n: "Pie" },
    0x101a: { n: "Area" },
    0x101b: { n: "Scatter" },
    0x101c: { n: "CrtLine" },
    0x101d: { n: "Axis" },
    0x101e: { n: "Tick" },
    0x101f: { n: "ValueRange" },
    0x1020: { n: "CatSerRange" },
    0x1021: { n: "AxisLine" },
    0x1022: { n: "CrtLink" },
    0x1024: { n: "DefaultText" },
    0x1025: { n: "Text" },
    0x1026: { n: "FontX", f: parseuint16 },
    0x1027: { n: "ObjectLink" },
    0x1032: { n: "Frame" },
    0x1033: { n: "Begin" },
    0x1034: { n: "End" },
    0x1035: { n: "PlotArea" },
    0x103a: { n: "Chart3d" },
    0x103c: { n: "PicF" },
    0x103d: { n: "DropBar" },
    0x103e: { n: "Radar" },
    0x103f: { n: "Surf" },
    0x1040: { n: "RadarArea" },
    0x1041: { n: "AxisParent" },
    0x1043: { n: "LegendException" },
    0x1044: { n: "ShtProps", f: parse_ShtProps },
    0x1045: { n: "SerToCrt" },
    0x1046: { n: "AxesUsed" },
    0x1048: { n: "SBaseRef" },
    0x104a: { n: "SerParent" },
    0x104b: { n: "SerAuxTrend" },
    0x104e: { n: "IFmtRecord" },
    0x104f: { n: "Pos" },
    0x1050: { n: "AlRuns" },
    0x1051: { n: "BRAI" },
    0x105b: { n: "SerAuxErrBar" },
    0x105c: { n: "ClrtClient", f: parse_ClrtClient },
    0x105d: { n: "SerFmt" },
    0x105f: { n: "Chart3DBarShape" },
    0x1060: { n: "Fbi" },
    0x1061: { n: "BopPop" },
    0x1062: { n: "AxcExt" },
    0x1063: { n: "Dat" },
    0x1064: { n: "PlotGrowth" },
    0x1065: { n: "SIIndex" },
    0x1066: { n: "GelFrame" },
    0x1067: { n: "BopPopCustom" },
    0x1068: { n: "Fbi2" },

    0x0000: { n: "Dimensions", f: parse_Dimensions },
    0x0002: { n: "BIFF2INT", f: parse_BIFF2INT },
    0x0005: { n: "BoolErr", f: parse_BoolErr },
    0x0007: { n: "String", f: parse_BIFF2STRING },
    0x0008: { n: "BIFF2ROW" },
    0x000b: { n: "Index" },
    0x0016: { n: "ExternCount", f: parseuint16 },
    0x001e: { n: "BIFF2FORMAT", f: parse_BIFF2Format },
    0x001f: { n: "BIFF2FMTCNT" } /* 16-bit cnt of BIFF2FORMAT records */,
    0x0020: { n: "BIFF2COLINFO" },
    0x0021: { n: "Array", f: parse_Array },
    0x0025: { n: "DefaultRowHeight", f: parse_DefaultRowHeight },
    0x0032: { n: "BIFF2FONTXTRA", f: parse_BIFF2FONTXTRA },
    0x0034: { n: "DDEObjName" },
    0x003e: { n: "BIFF2WINDOW2" },
    0x0043: { n: "BIFF2XF" },
    0x0045: { n: "BIFF2FONTCLR" },
    0x0056: { n: "BIFF4FMTCNT" } /* 16-bit cnt, similar to BIFF2 */,
    0x007e: { n: "RK" } /* Not necessarily same as 0x027e */,
    0x007f: { n: "ImData", f: parse_ImData },
    0x0087: { n: "Addin" },
    0x0088: { n: "Edg" },
    0x0089: { n: "Pub" },
    0x0091: { n: "Sub" },
    0x0094: { n: "LHRecord" },
    0x0095: { n: "LHNGraph" },
    0x0096: { n: "Sound" },
    0x00a9: { n: "CoordList" },
    0x00ab: { n: "GCW" },
    0x00bc: { n: "ShrFmla" } /* Not necessarily same as 0x04bc */,
    0x00bf: { n: "ToolbarHdr" },
    0x00c0: { n: "ToolbarEnd" },
    0x00c2: { n: "AddMenu" },
    0x00c3: { n: "DelMenu" },
    0x00d6: { n: "RString", f: parse_RString },
    0x00df: { n: "UDDesc" },
    0x00ea: { n: "TabIdConf" },
    0x0162: { n: "XL5Modify" },
    0x01a5: { n: "FileSharing2" },
    0x0209: { n: "BOF", f: parse_BOF },
    0x0218: { n: "Lbl", f: parse_Lbl },
    0x0223: { n: "ExternName", f: parse_ExternName },
    0x0231: { n: "Font" },
    0x0243: { n: "BIFF3XF" },
    0x0409: { n: "BOF", f: parse_BOF },
    0x0443: { n: "BIFF4XF" },
    0x086d: { n: "FeatInfo" },
    0x0873: { n: "FeatInfo11" },
    0x0881: { n: "SXAddl12" },
    0x08c0: { n: "AutoWebPub" },
    0x08c1: { n: "ListObj" },
    0x08c2: { n: "ListField" },
    0x08c3: { n: "ListDV" },
    0x08c4: { n: "ListCondFmt" },
    0x08c5: { n: "ListCF" },
    0x08c6: { n: "FMQry" },
    0x08c7: { n: "FMSQry" },
    0x08c8: { n: "PLV" },
    0x08c9: { n: "LnExt" },
    0x08ca: { n: "MkrExt" },
    0x08cb: { n: "CrtCoopt" },
    0x08d6: { n: "FRTArchId$", r: 12 },

    0x7262: {}
  };

  const XLSRE = evert_key(XLSRecordEnum, "n");
  function write_biff_rec(ba, type, payload, length) {
    const t = +type || +XLSRE[type];
    if (isNaN(t)) return;
    const len = length || (payload || []).length || 0;
    const o = ba.next(4);
    o.write_shift(2, t);
    o.write_shift(2, len);
    if (len > 0 && is_buf(payload)) ba.push(payload);
  }

  function write_biff_continue(ba, type, payload, length) {
    const len = length || (payload || []).length || 0;
    if (len <= 8224) return write_biff_rec(ba, type, payload, len);
    const t = +type || +XLSRE[type];
    if (isNaN(t)) return;
    const parts = payload.parts || [];
    let sidx = 0;
    let i = 0;
    let w = 0;
    while (w + (parts[sidx] || 8224) <= 8224) {
      w += parts[sidx] || 8224;
      sidx++;
    }
    let o = ba.next(4);
    o.write_shift(2, t);
    o.write_shift(2, w);
    ba.push(payload.slice(i, i + w));
    i += w;
    while (i < len) {
      o = ba.next(4);
      o.write_shift(2, 0x3c); // TODO: figure out correct continue type
      w = 0;
      while (w + (parts[sidx] || 8224) <= 8224) {
        w += parts[sidx] || 8224;
        sidx++;
      }
      o.write_shift(2, w);
      ba.push(payload.slice(i, i + w));
      i += w;
    }
  }

  function write_BIFF2Cell(out, r, c) {
    if (!out) out = new_buf(7);
    out.write_shift(2, r);
    out.write_shift(2, c);
    out.write_shift(2, 0);
    out.write_shift(1, 0);
    return out;
  }

  function write_BIFF2BERR(r, c, val, t) {
    const out = new_buf(9);
    write_BIFF2Cell(out, r, c);
    if (t == "e") {
      out.write_shift(1, val);
      out.write_shift(1, 1);
    } else {
      out.write_shift(1, val ? 1 : 0);
      out.write_shift(1, 0);
    }
    return out;
  }

  /* TODO: codepage, large strings */
  function write_BIFF2LABEL(r, c, val) {
    const out = new_buf(8 + 2 * val.length);
    write_BIFF2Cell(out, r, c);
    out.write_shift(1, val.length);
    out.write_shift(val.length, val, "sbcs");
    return out.l < out.length ? out.slice(0, out.l) : out;
  }

  function write_ws_biff2_cell(ba, cell, R, C) {
    if (cell.v != null)
      switch (cell.t) {
        case "d":
        case "n":
          var v = cell.t == "d" ? datenum(parseDate(cell.v)) : cell.v;
          if (v == (v | 0) && v >= 0 && v < 65536)
            write_biff_rec(ba, 0x0002, write_BIFF2INT(R, C, v));
          else write_biff_rec(ba, 0x0003, write_BIFF2NUM(R, C, v));
          return;
        case "b":
        case "e":
          write_biff_rec(ba, 0x0005, write_BIFF2BERR(R, C, cell.v, cell.t));
          return;
        /* TODO: codepage, sst */
        case "s":
        case "str":
          write_biff_rec(ba, 0x0004, write_BIFF2LABEL(R, C, cell.v));
          return;
      }
    write_biff_rec(ba, 0x0001, write_BIFF2Cell(null, R, C));
  }

  function write_ws_biff2(ba, ws, idx, opts) {
    const dense = Array.isArray(ws);
    const range = safe_decode_range(ws["!ref"] || "A1");
    let ref;
    let rr = "";
    const cols = [];
    if (range.e.c > 0xff || range.e.r > 0x3fff) {
      if (opts.WTF)
        throw new Error(
          `Range ${ws["!ref"] || "A1"} exceeds format limit A1:IV16384`
        );
      range.e.c = Math.min(range.e.c, 0xff);
      range.e.r = Math.min(range.e.c, 0x3fff);
      ref = encode_range(range);
    }
    for (let R = range.s.r; R <= range.e.r; ++R) {
      rr = encode_row(R);
      for (let C = range.s.c; C <= range.e.c; ++C) {
        if (R === range.s.r) cols[C] = encode_col(C);
        ref = cols[C] + rr;
        const cell = dense ? (ws[R] || [])[C] : ws[ref];
        if (!cell) continue;
        /* write cell */
        write_ws_biff2_cell(ba, cell, R, C, opts);
      }
    }
  }

  /* Based on test files */
  function write_biff2_buf(wb, opts) {
    const o = opts || {};
    if (DENSE != null && o.dense == null) o.dense = DENSE;
    const ba = buf_array();
    let idx = 0;
    for (let i = 0; i < wb.SheetNames.length; ++i)
      if (wb.SheetNames[i] == o.sheet) idx = i;
    if (idx == 0 && !!o.sheet && wb.SheetNames[0] != o.sheet)
      throw new Error(`Sheet not found: ${o.sheet}`);
    write_biff_rec(ba, 0x0009, write_BOF(wb, 0x10, o));
    /* ... */
    write_ws_biff2(ba, wb.Sheets[wb.SheetNames[idx]], idx, o, wb);
    /* ... */
    write_biff_rec(ba, 0x000a);
    return ba.end();
  }

  function write_FONTS_biff8(ba, data, opts) {
    write_biff_rec(
      ba,
      "Font",
      write_Font(
        {
          sz: 12,
          color: { theme: 1 },
          name: "Arial",
          family: 2,
          scheme: "minor"
        },
        opts
      )
    );
  }

  function write_FMTS_biff8(ba, NF, opts) {
    if (!NF) return;
    [
      [5, 8],
      [23, 26],
      [41, 44],
      [/* 63 */ 50, /* 66],[164, */ 392]
    ].forEach(function(r) {
      for (let i = r[0]; i <= r[1]; ++i)
        if (NF[i] != null)
          write_biff_rec(ba, "Format", write_Format(i, NF[i], opts));
    });
  }

  function write_FEAT(ba, ws) {
    /* [MS-XLS] 2.4.112 */
    let o = new_buf(19);
    o.write_shift(4, 0x867);
    o.write_shift(4, 0);
    o.write_shift(4, 0);
    o.write_shift(2, 3);
    o.write_shift(1, 1);
    o.write_shift(4, 0);
    write_biff_rec(ba, "FeatHdr", o);
    /* [MS-XLS] 2.4.111 */
    o = new_buf(39);
    o.write_shift(4, 0x868);
    o.write_shift(4, 0);
    o.write_shift(4, 0);
    o.write_shift(2, 3);
    o.write_shift(1, 0);
    o.write_shift(4, 0);
    o.write_shift(2, 1);
    o.write_shift(4, 4);
    o.write_shift(2, 0);
    write_Ref8U(safe_decode_range(ws["!ref"] || "A1"), o);
    o.write_shift(4, 4);
    write_biff_rec(ba, "Feat", o);
  }

  function write_CELLXFS_biff8(ba, opts) {
    for (let i = 0; i < 16; ++i)
      write_biff_rec(ba, "XF", write_XF({ numFmtId: 0, style: true }, 0, opts));
    opts.cellXfs.forEach(function(c) {
      write_biff_rec(ba, "XF", write_XF(c, 0, opts));
    });
  }

  function write_ws_biff8_hlinks(ba, ws) {
    for (let R = 0; R < ws["!links"].length; ++R) {
      const HL = ws["!links"][R];
      write_biff_rec(ba, "HLink", write_HLink(HL));
      if (HL[1].Tooltip)
        write_biff_rec(ba, "HLinkTooltip", write_HLinkTooltip(HL));
    }
    delete ws["!links"];
  }

  function write_ws_biff8_cell(ba, cell, R, C, opts) {
    const os = 16 + get_cell_style(opts.cellXfs, cell, opts);
    if (cell.v == null && !cell.bf) {
      write_biff_rec(ba, "Blank", write_XLSCell(R, C, os));
      return;
    }
    if (cell.bf)
      write_biff_rec(ba, "Formula", write_Formula(cell, R, C, opts, os));
    else
      switch (cell.t) {
        case "d":
        case "n":
          var v = cell.t == "d" ? datenum(parseDate(cell.v)) : cell.v;
          /* TODO: emit RK as appropriate */
          write_biff_rec(ba, "Number", write_Number(R, C, v, os, opts));
          break;
        case "b":
        case "e":
          write_biff_rec(
            ba,
            0x0205,
            write_BoolErr(R, C, cell.v, os, opts, cell.t)
          );
          break;
        /* TODO: codepage, sst */
        case "s":
        case "str":
          if (opts.bookSST) {
            const isst = get_sst_id(opts.Strings, cell.v, opts.revStrings);
            write_biff_rec(
              ba,
              "LabelSst",
              write_LabelSst(R, C, isst, os, opts)
            );
          } else
            write_biff_rec(ba, "Label", write_Label(R, C, cell.v, os, opts));
          break;
        default:
          write_biff_rec(ba, "Blank", write_XLSCell(R, C, os));
      }
  }

  /* [MS-XLS] 2.1.7.20.5 */
  function write_ws_biff8(idx, opts, wb) {
    const ba = buf_array();
    const s = wb.SheetNames[idx];
    const ws = wb.Sheets[s] || {};
    const _WB = (wb || {}).Workbook || {};
    const _sheet = (_WB.Sheets || [])[idx] || {};
    const dense = Array.isArray(ws);
    const b8 = opts.biff == 8;
    let ref;
    let rr = "";
    const cols = [];
    const range = safe_decode_range(ws["!ref"] || "A1");
    const MAX_ROWS = b8 ? 65536 : 16384;
    if (range.e.c > 0xff || range.e.r >= MAX_ROWS) {
      if (opts.WTF)
        throw new Error(
          `Range ${ws["!ref"] || "A1"} exceeds format limit A1:IV16384`
        );
      range.e.c = Math.min(range.e.c, 0xff);
      range.e.r = Math.min(range.e.c, MAX_ROWS - 1);
    }

    write_biff_rec(ba, 0x0809, write_BOF(wb, 0x10, opts));
    /* [Uncalced] Index */
    write_biff_rec(ba, "CalcMode", writeuint16(1));
    write_biff_rec(ba, "CalcCount", writeuint16(100));
    write_biff_rec(ba, "CalcRefMode", writebool(true));
    write_biff_rec(ba, "CalcIter", writebool(false));
    write_biff_rec(ba, "CalcDelta", write_Xnum(0.001));
    write_biff_rec(ba, "CalcSaveRecalc", writebool(true));
    write_biff_rec(ba, "PrintRowCol", writebool(false));
    write_biff_rec(ba, "PrintGrid", writebool(false));
    write_biff_rec(ba, "GridSet", writeuint16(1));
    write_biff_rec(ba, "Guts", write_Guts([0, 0]));
    /* DefaultRowHeight WsBool [Sync] [LPr] [HorizontalPageBreaks] [VerticalPageBreaks] */
    /* Header (string) */
    /* Footer (string) */
    write_biff_rec(ba, "HCenter", writebool(false));
    write_biff_rec(ba, "VCenter", writebool(false));
    /* ... */
    write_biff_rec(ba, 0x200, write_Dimensions(range, opts));
    /* ... */

    if (b8) ws["!links"] = [];
    for (let R = range.s.r; R <= range.e.r; ++R) {
      rr = encode_row(R);
      for (let C = range.s.c; C <= range.e.c; ++C) {
        if (R === range.s.r) cols[C] = encode_col(C);
        ref = cols[C] + rr;
        const cell = dense ? (ws[R] || [])[C] : ws[ref];
        if (!cell) continue;
        /* write cell */
        write_ws_biff8_cell(ba, cell, R, C, opts);
        if (b8 && cell.l) ws["!links"].push([ref, cell.l]);
      }
    }
    const cname = _sheet.CodeName || _sheet.name || s;
    /* ... */
    if (b8 && _WB.Views)
      write_biff_rec(ba, "Window2", write_Window2(_WB.Views[0]));
    /* ... */
    if (b8 && (ws["!merges"] || []).length)
      write_biff_rec(ba, "MergeCells", write_MergeCells(ws["!merges"]));
    /* [LRng] *QUERYTABLE [PHONETICINFO] CONDFMTS */
    if (b8) write_ws_biff8_hlinks(ba, ws);
    /* [DVAL] */
    write_biff_rec(ba, "CodeName", write_XLUnicodeString(cname, opts));
    /* *WebPub *CellWatch [SheetExt] */
    if (b8) write_FEAT(ba, ws);
    /* *FEAT11 *RECORD12 */
    write_biff_rec(ba, "EOF");
    return ba.end();
  }

  /* [MS-XLS] 2.1.7.20.3 */
  function write_biff8_global(wb, bufs, opts) {
    const A = buf_array();
    const _WB = (wb || {}).Workbook || {};
    const _sheets = _WB.Sheets || [];
    const _wb = _WB.WBProps || {};
    const b8 = opts.biff == 8;
    const b5 = opts.biff == 5;
    write_biff_rec(A, 0x0809, write_BOF(wb, 0x05, opts));
    if (opts.bookType == "xla") write_biff_rec(A, "Addin");
    write_biff_rec(A, "InterfaceHdr", b8 ? writeuint16(0x04b0) : null);
    write_biff_rec(A, "Mms", writezeroes(2));
    if (b5) write_biff_rec(A, "ToolbarHdr");
    if (b5) write_biff_rec(A, "ToolbarEnd");
    write_biff_rec(A, "InterfaceEnd");
    write_biff_rec(A, "WriteAccess", write_WriteAccess("SheetJS", opts));
    /* [FileSharing] */
    write_biff_rec(A, "CodePage", writeuint16(b8 ? 0x04b0 : 0x04e4));
    /* *2047 Lel */
    if (b8) write_biff_rec(A, "DSF", writeuint16(0));
    if (b8) write_biff_rec(A, "Excel9File");
    write_biff_rec(A, "RRTabId", write_RRTabId(wb.SheetNames.length));
    if (b8 && wb.vbaraw) write_biff_rec(A, "ObProj");
    /* [ObNoMacros] */
    if (b8 && wb.vbaraw) {
      const cname = _wb.CodeName || "ThisWorkbook";
      write_biff_rec(A, "CodeName", write_XLUnicodeString(cname, opts));
    }
    write_biff_rec(A, "BuiltInFnGroupCount", writeuint16(0x11));
    /* *FnGroupName *FnGrp12 */
    /* *Lbl */
    /* [OleObjectSize] */
    write_biff_rec(A, "WinProtect", writebool(false));
    write_biff_rec(A, "Protect", writebool(false));
    write_biff_rec(A, "Password", writeuint16(0));
    if (b8) write_biff_rec(A, "Prot4Rev", writebool(false));
    if (b8) write_biff_rec(A, "Prot4RevPass", writeuint16(0));
    write_biff_rec(A, "Window1", write_Window1(opts));
    write_biff_rec(A, "Backup", writebool(false));
    write_biff_rec(A, "HideObj", writeuint16(0));
    write_biff_rec(A, "Date1904", writebool(safe1904(wb) == "true"));
    write_biff_rec(A, "CalcPrecision", writebool(true));
    if (b8) write_biff_rec(A, "RefreshAll", writebool(false));
    write_biff_rec(A, "BookBool", writeuint16(0));
    /* ... */
    write_FONTS_biff8(A, wb, opts);
    write_FMTS_biff8(A, wb.SSF, opts);
    write_CELLXFS_biff8(A, opts);
    /* ... */
    if (b8) write_biff_rec(A, "UsesELFs", writebool(false));
    const a = A.end();

    const C = buf_array();
    /* METADATA [MTRSettings] [ForceFullCalculation] */
    if (b8) write_biff_rec(C, "Country", write_Country());
    /* *SUPBOOK *LBL *RTD [RecalcId] *HFPicture *MSODRAWINGGROUP */

    /* BIFF8: [SST *Continue] ExtSST */
    if (b8 && opts.Strings)
      write_biff_continue(C, "SST", write_SST(opts.Strings, opts));

    /* *WebPub [WOpt] [CrErr] [BookExt] *FeatHdr *DConn [THEME] [CompressPictures] [Compat12] [GUIDTypeLib] */
    write_biff_rec(C, "EOF");
    const c = C.end();

    const B = buf_array();
    let blen = 0;
    let j = 0;
    for (j = 0; j < wb.SheetNames.length; ++j)
      blen += (b8 ? 12 : 11) + (b8 ? 2 : 1) * wb.SheetNames[j].length;
    let start = a.length + blen + c.length;
    for (j = 0; j < wb.SheetNames.length; ++j) {
      const _sheet = _sheets[j] || {};
      write_biff_rec(
        B,
        "BoundSheet8",
        write_BoundSheet8(
          { pos: start, hs: _sheet.Hidden || 0, dt: 0, name: wb.SheetNames[j] },
          opts
        )
      );
      start += bufs[j].length;
    }
    /* 1*BoundSheet8 */
    const b = B.end();
    if (blen != b.length) throw new Error(`BS8 ${blen} != ${b.length}`);

    const out = [];
    if (a.length) out.push(a);
    if (b.length) out.push(b);
    if (c.length) out.push(c);
    return __toBuffer([out]);
  }

  /* [MS-XLS] 2.1.7.20 Workbook Stream */
  function write_biff8_buf(wb, opts) {
    const o = opts || {};
    const bufs = [];

    if (wb && !wb.SSF) {
      wb.SSF = SSF.get_table();
    }
    if (wb && wb.SSF) {
      make_ssf(SSF);
      SSF.load_table(wb.SSF);
      // $FlowIgnore
      o.revssf = evert_num(wb.SSF);
      o.revssf[wb.SSF[65535]] = 0;
      o.ssf = wb.SSF;
    }

    o.Strings = [];
    o.Strings.Count = 0;
    o.Strings.Unique = 0;
    fix_write_opts(o);

    o.cellXfs = [];
    get_cell_style(o.cellXfs, {}, { revssf: { General: 0 } });

    if (!wb.Props) wb.Props = {};

    for (let i = 0; i < wb.SheetNames.length; ++i)
      bufs[bufs.length] = write_ws_biff8(i, o, wb);
    bufs.unshift(write_biff8_global(wb, bufs, o));
    return __toBuffer([bufs]);
  }

  /* note: browser DOM element cannot see mso- style attrs, must parse */
  var HTML_ = (function() {
    function html_to_sheet(str, _opts) {
      const opts = _opts || {};
      if (DENSE != null && opts.dense == null) opts.dense = DENSE;
      const ws = opts.dense ? [] : {};
      str = str.replace(/<!--.*?-->/g, "");
      const mtch = str.match(/<table/i);
      if (!mtch) throw new Error("Invalid HTML: could not find <table>");
      const mtch2 = str.match(/<\/table/i);
      let i = mtch.index;
      let j = (mtch2 && mtch2.index) || str.length;
      const rows = split_regex(str.slice(i, j), /(:?<tr[^>]*>)/i, "<tr>");
      let R = -1;
      let C = 0;
      let RS = 0;
      let CS = 0;
      const range = { s: { r: 10000000, c: 10000000 }, e: { r: 0, c: 0 } };
      const merges = [];
      for (i = 0; i < rows.length; ++i) {
        const row = rows[i].trim();
        const hd = row.slice(0, 3).toLowerCase();
        if (hd == "<tr") {
          ++R;
          if (opts.sheetRows && opts.sheetRows <= R) {
            --R;
            break;
          }
          C = 0;
          continue;
        }
        if (hd != "<td" && hd != "<th") continue;
        const cells = row.split(/<\/t[dh]>/i);
        for (j = 0; j < cells.length; ++j) {
          const cell = cells[j].trim();
          if (!cell.match(/<t[dh]/i)) continue;
          let m = cell;
          let cc = 0;
          /* TODO: parse styles etc */
          while (m.charAt(0) == "<" && (cc = m.indexOf(">")) > -1)
            m = m.slice(cc + 1);
          for (let midx = 0; midx < merges.length; ++midx) {
            const _merge = merges[midx];
            if (_merge.s.c == C && _merge.s.r < R && R <= _merge.e.r) {
              C = _merge.e.c + 1;
              midx = -1;
            }
          }
          const tag = parsexmltag(cell.slice(0, cell.indexOf(">")));
          CS = tag.colspan ? +tag.colspan : 1;
          if ((RS = +tag.rowspan) > 1 || CS > 1)
            merges.push({
              s: { r: R, c: C },
              e: { r: R + (RS || 1) - 1, c: C + CS - 1 }
            });
          const _t = tag.t || "";
          /* TODO: generate stub cells */
          if (!m.length) {
            C += CS;
            continue;
          }
          m = htmldecode(m);
          if (range.s.r > R) range.s.r = R;
          if (range.e.r < R) range.e.r = R;
          if (range.s.c > C) range.s.c = C;
          if (range.e.c < C) range.e.c = C;
          if (!m.length) continue;
          let o = { t: "s", v: m };
          if (opts.raw || !m.trim().length || _t == "s") {
          } else if (m === "TRUE") o = { t: "b", v: true };
          else if (m === "FALSE") o = { t: "b", v: false };
          else if (!isNaN(fuzzynum(m))) o = { t: "n", v: fuzzynum(m) };
          else if (!isNaN(fuzzydate(m).getDate())) {
            o = { t: "d", v: parseDate(m) };
            if (!opts.cellDates) o = { t: "n", v: datenum(o.v) };
            o.z = opts.dateNF || SSF._table[14];
          }
          if (opts.dense) {
            if (!ws[R]) ws[R] = [];
            ws[R][C] = o;
          } else ws[encode_cell({ r: R, c: C })] = o;
          C += CS;
        }
      }
      ws["!ref"] = encode_range(range);
      if (merges.length) ws["!merges"] = merges;
      return ws;
    }
    function html_to_book(str, opts) {
      return sheet_to_workbook(html_to_sheet(str, opts), opts);
    }
    function make_html_row(ws, r, R, o) {
      const M = ws["!merges"] || [];
      const oo = [];
      for (let C = r.s.c; C <= r.e.c; ++C) {
        let RS = 0;
        let CS = 0;
        for (let j = 0; j < M.length; ++j) {
          if (M[j].s.r > R || M[j].s.c > C) continue;
          if (M[j].e.r < R || M[j].e.c < C) continue;
          if (M[j].s.r < R || M[j].s.c < C) {
            RS = -1;
            break;
          }
          RS = M[j].e.r - M[j].s.r + 1;
          CS = M[j].e.c - M[j].s.c + 1;
          break;
        }
        if (RS < 0) continue;
        const coord = encode_cell({ r: R, c: C });
        const cell = o.dense ? (ws[R] || [])[C] : ws[coord];
        /* TODO: html entities */
        let w =
          (cell &&
            cell.v != null &&
            (cell.h ||
              escapehtml(cell.w || (format_cell(cell), cell.w) || ""))) ||
          "";
        const sp = {};
        if (RS > 1) sp.rowspan = RS;
        if (CS > 1) sp.colspan = CS;
        sp.t = (cell && cell.t) || "z";
        if (o.editable) w = `<span contenteditable="true">${w}</span>`;
        sp.id = `${o.id || "sjs"}-${coord}`;
        if (sp.t != "z") {
          sp.v = cell.v;
          if (cell.z != null) sp.z = cell.z;
        }
        oo.push(writextag("td", w, sp));
      }
      const preamble = "<tr>";
      return `${preamble + oo.join("")}</tr>`;
    }
    function make_html_preamble(ws, R, o) {
      const out = [];
      return `${out.join("")}<table${o && o.id ? ` id="${o.id}"` : ""}>`;
    }
    const _BEGIN =
      '<html><head><meta charset="utf-8"/><title>SheetJS Table Export</title></head><body>';
    const _END = "</body></html>";
    function sheet_to_html(ws, opts /* , wb:?Workbook */) {
      const o = opts || {};
      const header = o.header != null ? o.header : _BEGIN;
      const footer = o.footer != null ? o.footer : _END;
      const out = [header];
      const r = decode_range(ws["!ref"]);
      o.dense = Array.isArray(ws);
      out.push(make_html_preamble(ws, r, o));
      for (let R = r.s.r; R <= r.e.r; ++R) out.push(make_html_row(ws, r, R, o));
      out.push(`</table>${footer}`);
      return out.join("");
    }
    return {
      to_workbook: html_to_book,
      to_sheet: html_to_sheet,
      _row: make_html_row,
      BEGIN: _BEGIN,
      END: _END,
      _preamble: make_html_preamble,
      from_sheet: sheet_to_html
    };
  })();

  function is_dom_element_hidden(element) {
    let display = "";
    const get_computed_style = get_get_computed_style_function(element);
    if (get_computed_style)
      display = get_computed_style(element).getPropertyValue("display");
    if (!display) display = element.style.display; // Fallback for cases when getComputedStyle is not available (e.g. an old browser or some Node.js environments) or doesn't work (e.g. if the element is not inserted to a document)
    return display === "none";
  }

  /* global getComputedStyle */
  function get_get_computed_style_function(element) {
    // The proper getComputedStyle implementation is the one defined in the element window
    if (
      element.ownerDocument.defaultView &&
      typeof element.ownerDocument.defaultView.getComputedStyle === "function"
    )
      return element.ownerDocument.defaultView.getComputedStyle;
    // If it is not available, try to get one from the global namespace
    if (typeof getComputedStyle === "function") return getComputedStyle;
    return null;
  }
  /* OpenDocument */
  const parse_content_xml = (function() {
    const parse_text_p = function(text) {
      /* 6.1.2 White Space Characters */
      const fixed = text
        .replace(/[\t\r\n]/g, " ")
        .trim()
        .replace(/ +/g, " ")
        .replace(/<text:s\/>/g, " ")
        .replace(/<text:s text:c="(\d+)"\/>/g, function($$, $1) {
          return Array(parseInt($1, 10) + 1).join(" ");
        })
        .replace(/<text:tab[^>]*\/>/g, "\t")
        .replace(/<text:line-break\/>/g, "\n");
      const v = unescapexml(fixed.replace(/<[^>]*>/g, ""));

      return [v];
    };

    const number_formats = {
      /* ods name: [short ssf fmt, long ssf fmt] */
      day: ["d", "dd"],
      month: ["m", "mm"],
      year: ["y", "yy"],
      hours: ["h", "hh"],
      minutes: ["m", "mm"],
      seconds: ["s", "ss"],
      "am-pm": ["A/P", "AM/PM"],
      "day-of-week": ["ddd", "dddd"],
      era: ["e", "ee"],
      /* there is no native representation of LO "Q" format */
      quarter: ["\\Qm", 'm\\"th quarter"']
    };

    return function pcx(d, _opts) {
      const opts = _opts || {};
      if (DENSE != null && opts.dense == null) opts.dense = DENSE;
      let str = xlml_normalize(d);
      const state = [];
      let tmp;
      let tag;
      let NFtag = { name: "" };
      let NF = "";
      let pidx = 0;
      let sheetag;
      let rowtag;
      const Sheets = {};
      const SheetNames = [];
      let ws = opts.dense ? [] : {};
      let Rn;
      let q;
      let ctag = { value: "" };
      let textp = "";
      let textpidx = 0;
      let textptag;
      let textR = [];
      let R = -1;
      let C = -1;
      const range = { s: { r: 1000000, c: 10000000 }, e: { r: 0, c: 0 } };
      let row_ol = 0;
      const number_format_map = {};
      let merges = [];
      let mrange = {};
      let mR = 0;
      let mC = 0;
      let rowinfo = [];
      let rowpeat = 1;
      let colpeat = 1;
      const arrayf = [];
      const WB = { Names: [] };
      let atag = {};
      let _Ref = ["", ""];
      let comments = [];
      let comment = {};
      let creator = "";
      let creatoridx = 0;
      let isstub = false;
      let intable = false;
      let i = 0;
      xlmlregex.lastIndex = 0;
      str = str
        .replace(/<!--([\s\S]*?)-->/gm, "")
        .replace(/<!DOCTYPE[^\[]*\[[^\]]*\]>/gm, "");
      while ((Rn = xlmlregex.exec(str)))
        switch ((Rn[3] = Rn[3].replace(/_.*$/, ""))) {
          case "table":
          case "工作表": // 9.1.2 <table:table>
            if (Rn[1] === "/") {
              if (range.e.c >= range.s.c && range.e.r >= range.s.r)
                ws["!ref"] = encode_range(range);
              if (opts.sheetRows > 0 && opts.sheetRows <= range.e.r) {
                ws["!fullref"] = ws["!ref"];
                range.e.r = opts.sheetRows - 1;
                ws["!ref"] = encode_range(range);
              }
              if (merges.length) ws["!merges"] = merges;
              if (rowinfo.length) ws["!rows"] = rowinfo;
              sheetag.name = sheetag["名称"] || sheetag.name;
              if (typeof JSON !== "undefined") JSON.stringify(sheetag);
              SheetNames.push(sheetag.name);
              Sheets[sheetag.name] = ws;
              intable = false;
            } else if (Rn[0].charAt(Rn[0].length - 2) !== "/") {
              sheetag = parsexmltag(Rn[0], false);
              R = C = -1;
              range.s.r = range.s.c = 10000000;
              range.e.r = range.e.c = 0;
              ws = opts.dense ? [] : {};
              merges = [];
              rowinfo = [];
              intable = true;
            }
            break;

          case "table-row-group": // 9.1.9 <table:table-row-group>
            if (Rn[1] === "/") --row_ol;
            else ++row_ol;
            break;
          case "table-row":
          case "行": // 9.1.3 <table:table-row>
            if (Rn[1] === "/") {
              R += rowpeat;
              rowpeat = 1;
              break;
            }
            rowtag = parsexmltag(Rn[0], false);
            if (rowtag["行号"]) R = rowtag["行号"] - 1;
            else if (R == -1) R = 0;
            rowpeat = +rowtag["number-rows-repeated"] || 1;
            /* TODO: remove magic */
            if (rowpeat < 10)
              for (i = 0; i < rowpeat; ++i)
                if (row_ol > 0) rowinfo[R + i] = { level: row_ol };
            C = -1;
            break;
          case "covered-table-cell": // 9.1.5 <table:covered-table-cell>
            if (Rn[1] !== "/") ++C;
            if (opts.sheetStubs) {
              if (opts.dense) {
                if (!ws[R]) ws[R] = [];
                ws[R][C] = { t: "z" };
              } else ws[encode_cell({ r: R, c: C })] = { t: "z" };
            }
            textp = "";
            textR = [];
            break; /* stub */
          case "table-cell":
          case "数据":
            if (Rn[0].charAt(Rn[0].length - 2) === "/") {
              ++C;
              ctag = parsexmltag(Rn[0], false);
              colpeat = parseInt(ctag["number-columns-repeated"] || "1", 10);
              q = { t: "z", v: null };
              if (ctag.formula && opts.cellFormula != false)
                q.f = ods_to_csf_formula(unescapexml(ctag.formula));
              if ((ctag["数据类型"] || ctag["value-type"]) == "string") {
                q.t = "s";
                q.v = unescapexml(ctag["string-value"] || "");
                if (opts.dense) {
                  if (!ws[R]) ws[R] = [];
                  ws[R][C] = q;
                } else {
                  ws[encode_cell({ r: R, c: C })] = q;
                }
              }
              C += colpeat - 1;
            } else if (Rn[1] !== "/") {
              ++C;
              colpeat = 1;
              const rptR = rowpeat ? R + rowpeat - 1 : R;
              if (C > range.e.c) range.e.c = C;
              if (C < range.s.c) range.s.c = C;
              if (R < range.s.r) range.s.r = R;
              if (rptR > range.e.r) range.e.r = rptR;
              ctag = parsexmltag(Rn[0], false);
              comments = [];
              comment = {};
              q = { t: ctag["数据类型"] || ctag["value-type"], v: null };
              if (opts.cellFormula) {
                if (ctag.formula) ctag.formula = unescapexml(ctag.formula);
                if (
                  ctag["number-matrix-columns-spanned"] &&
                  ctag["number-matrix-rows-spanned"]
                ) {
                  mR = parseInt(ctag["number-matrix-rows-spanned"], 10) || 0;
                  mC = parseInt(ctag["number-matrix-columns-spanned"], 10) || 0;
                  mrange = {
                    s: { r: R, c: C },
                    e: { r: R + mR - 1, c: C + mC - 1 }
                  };
                  q.F = encode_range(mrange);
                  arrayf.push([mrange, q.F]);
                }
                if (ctag.formula) q.f = ods_to_csf_formula(ctag.formula);
                else
                  for (i = 0; i < arrayf.length; ++i)
                    if (R >= arrayf[i][0].s.r && R <= arrayf[i][0].e.r)
                      if (C >= arrayf[i][0].s.c && C <= arrayf[i][0].e.c)
                        q.F = arrayf[i][1];
              }
              if (
                ctag["number-columns-spanned"] ||
                ctag["number-rows-spanned"]
              ) {
                mR = parseInt(ctag["number-rows-spanned"], 10) || 0;
                mC = parseInt(ctag["number-columns-spanned"], 10) || 0;
                mrange = {
                  s: { r: R, c: C },
                  e: { r: R + mR - 1, c: C + mC - 1 }
                };
                merges.push(mrange);
              }

              /* 19.675.2 table:number-columns-repeated */
              if (ctag["number-columns-repeated"])
                colpeat = parseInt(ctag["number-columns-repeated"], 10);

              /* 19.385 office:value-type */
              switch (q.t) {
                case "boolean":
                  q.t = "b";
                  q.v = parsexmlbool(ctag["boolean-value"]);
                  break;
                case "float":
                  q.t = "n";
                  q.v = parseFloat(ctag.value);
                  break;
                case "percentage":
                  q.t = "n";
                  q.v = parseFloat(ctag.value);
                  break;
                case "currency":
                  q.t = "n";
                  q.v = parseFloat(ctag.value);
                  break;
                case "date":
                  q.t = "d";
                  q.v = parseDate(ctag["date-value"]);
                  if (!opts.cellDates) {
                    q.t = "n";
                    q.v = datenum(q.v);
                  }
                  q.z = "m/d/yy";
                  break;
                case "time":
                  q.t = "n";
                  q.v = parse_isodur(ctag["time-value"]) / 86400;
                  break;
                case "number":
                  q.t = "n";
                  q.v = parseFloat(ctag["数据数值"]);
                  break;
                default:
                  if (q.t === "string" || q.t === "text" || !q.t) {
                    q.t = "s";
                    if (ctag["string-value"] != null) {
                      textp = unescapexml(ctag["string-value"]);
                      textR = [];
                    }
                  } else throw new Error(`Unsupported value type ${q.t}`);
              }
            } else {
              isstub = false;
              if (q.t === "s") {
                q.v = textp || "";
                if (textR.length) q.R = textR;
                isstub = textpidx == 0;
              }
              if (atag.Target) q.l = atag;
              if (comments.length > 0) {
                q.c = comments;
                comments = [];
              }
              if (textp && opts.cellText !== false) q.w = textp;
              if (isstub) {
                q.t = "z";
                delete q.v;
              }
              if (!isstub || opts.sheetStubs) {
                if (!(opts.sheetRows && opts.sheetRows <= R)) {
                  for (let rpt = 0; rpt < rowpeat; ++rpt) {
                    colpeat = parseInt(
                      ctag["number-columns-repeated"] || "1",
                      10
                    );
                    if (opts.dense) {
                      if (!ws[R + rpt]) ws[R + rpt] = [];
                      ws[R + rpt][C] = rpt == 0 ? q : dup(q);
                      while (--colpeat > 0) ws[R + rpt][C + colpeat] = dup(q);
                    } else {
                      ws[encode_cell({ r: R + rpt, c: C })] = q;
                      while (--colpeat > 0)
                        ws[encode_cell({ r: R + rpt, c: C + colpeat })] = dup(
                          q
                        );
                    }
                    if (range.e.c <= C) range.e.c = C;
                  }
                }
              }
              colpeat = parseInt(ctag["number-columns-repeated"] || "1", 10);
              C += colpeat - 1;
              colpeat = 0;
              q = {};
              textp = "";
              textR = [];
            }
            atag = {};
            break; // 9.1.4 <table:table-cell>

          /* pure state */
          case "document": // TODO: <office:document> is the root for FODS
          case "document-content":
          case "电子表格文档": // 3.1.3.2 <office:document-content>
          case "spreadsheet":
          case "主体": // 3.7 <office:spreadsheet>
          case "scripts": // 3.12 <office:scripts>
          case "styles": // TODO <office:styles>
          case "font-face-decls": // 3.14 <office:font-face-decls>
            if (Rn[1] === "/") {
              if ((tmp = state.pop())[0] !== Rn[3]) throw `Bad state: ${tmp}`;
            } else if (Rn[0].charAt(Rn[0].length - 2) !== "/")
              state.push([Rn[3], true]);
            break;

          case "annotation": // 14.1 <office:annotation>
            if (Rn[1] === "/") {
              if ((tmp = state.pop())[0] !== Rn[3]) throw `Bad state: ${tmp}`;
              comment.t = textp;
              if (textR.length) comment.R = textR;
              comment.a = creator;
              comments.push(comment);
            } else if (Rn[0].charAt(Rn[0].length - 2) !== "/") {
              state.push([Rn[3], false]);
            }
            creator = "";
            creatoridx = 0;
            textp = "";
            textpidx = 0;
            textR = [];
            break;

          case "creator": // 4.3.2.7 <dc:creator>
            if (Rn[1] === "/") {
              creator = str.slice(creatoridx, Rn.index);
            } else creatoridx = Rn.index + Rn[0].length;
            break;

          /* ignore state */
          case "meta":
          case "元数据": // TODO: <office:meta> <uof:元数据> FODS/UOF
          case "settings": // TODO: <office:settings>
          case "config-item-set": // TODO: <office:config-item-set>
          case "config-item-map-indexed": // TODO: <office:config-item-map-indexed>
          case "config-item-map-entry": // TODO: <office:config-item-map-entry>
          case "config-item-map-named": // TODO: <office:config-item-map-entry>
          case "shapes": // 9.2.8 <table:shapes>
          case "frame": // 10.4.2 <draw:frame>
          case "text-box": // 10.4.3 <draw:text-box>
          case "image": // 10.4.4 <draw:image>
          case "data-pilot-tables": // 9.6.2 <table:data-pilot-tables>
          case "list-style": // 16.30 <text:list-style>
          case "form": // 13.13 <form:form>
          case "dde-links": // 9.8 <table:dde-links>
          case "event-listeners": // TODO
          case "chart": // TODO
            if (Rn[1] === "/") {
              if ((tmp = state.pop())[0] !== Rn[3]) throw `Bad state: ${tmp}`;
            } else if (Rn[0].charAt(Rn[0].length - 2) !== "/")
              state.push([Rn[3], false]);
            textp = "";
            textpidx = 0;
            textR = [];
            break;

          case "scientific-number": // TODO: <number:scientific-number>
            break;
          case "currency-symbol": // TODO: <number:currency-symbol>
            break;
          case "currency-style": // TODO: <number:currency-style>
            break;
          case "number-style": // 16.27.2 <number:number-style>
          case "percentage-style": // 16.27.9 <number:percentage-style>
          case "date-style": // 16.27.10 <number:date-style>
          case "time-style": // 16.27.18 <number:time-style>
            if (Rn[1] === "/") {
              number_format_map[NFtag.name] = NF;
              if ((tmp = state.pop())[0] !== Rn[3]) throw `Bad state: ${tmp}`;
            } else if (Rn[0].charAt(Rn[0].length - 2) !== "/") {
              NF = "";
              NFtag = parsexmltag(Rn[0], false);
              state.push([Rn[3], true]);
            }
            break;

          case "script":
            break; // 3.13 <office:script>
          case "libraries":
            break; // TODO: <ooo:libraries>
          case "automatic-styles":
            break; // 3.15.3 <office:automatic-styles>
          case "master-styles":
            break; // TODO: <office:master-styles>

          case "default-style": // TODO: <style:default-style>
          case "page-layout":
            break; // TODO: <style:page-layout>
          case "style": // 16.2 <style:style>
            break;
          case "map":
            break; // 16.3 <style:map>
          case "font-face":
            break; // 16.21 <style:font-face>

          case "paragraph-properties":
            break; // 17.6 <style:paragraph-properties>
          case "table-properties":
            break; // 17.15 <style:table-properties>
          case "table-column-properties":
            break; // 17.16 <style:table-column-properties>
          case "table-row-properties":
            break; // 17.17 <style:table-row-properties>
          case "table-cell-properties":
            break; // 17.18 <style:table-cell-properties>

          case "number": // 16.27.3 <number:number>
            switch (state[state.length - 1][0]) {
              case "time-style":
              case "date-style":
                tag = parsexmltag(Rn[0], false);
                NF += number_formats[Rn[3]][tag.style === "long" ? 1 : 0];
                break;
            }
            break;

          case "fraction":
            break; // TODO 16.27.6 <number:fraction>

          case "day": // 16.27.11 <number:day>
          case "month": // 16.27.12 <number:month>
          case "year": // 16.27.13 <number:year>
          case "era": // 16.27.14 <number:era>
          case "day-of-week": // 16.27.15 <number:day-of-week>
          case "week-of-year": // 16.27.16 <number:week-of-year>
          case "quarter": // 16.27.17 <number:quarter>
          case "hours": // 16.27.19 <number:hours>
          case "minutes": // 16.27.20 <number:minutes>
          case "seconds": // 16.27.21 <number:seconds>
          case "am-pm": // 16.27.22 <number:am-pm>
            switch (state[state.length - 1][0]) {
              case "time-style":
              case "date-style":
                tag = parsexmltag(Rn[0], false);
                NF += number_formats[Rn[3]][tag.style === "long" ? 1 : 0];
                break;
            }
            break;

          case "boolean-style":
            break; // 16.27.23 <number:boolean-style>
          case "boolean":
            break; // 16.27.24 <number:boolean>
          case "text-style":
            break; // 16.27.25 <number:text-style>
          case "text": // 16.27.26 <number:text>
            if (Rn[0].slice(-2) === "/>") break;
            else if (Rn[1] === "/")
              switch (state[state.length - 1][0]) {
                case "number-style":
                case "date-style":
                case "time-style":
                  NF += str.slice(pidx, Rn.index);
                  break;
              }
            else pidx = Rn.index + Rn[0].length;
            break;

          case "named-range": // 9.4.12 <table:named-range>
            tag = parsexmltag(Rn[0], false);
            _Ref = ods_to_csf_3D(tag["cell-range-address"]);
            var nrange = { Name: tag.name, Ref: `${_Ref[0]}!${_Ref[1]}` };
            if (intable) nrange.Sheet = SheetNames.length;
            WB.Names.push(nrange);
            break;

          case "text-content":
            break; // 16.27.27 <number:text-content>
          case "text-properties":
            break; // 16.27.27 <style:text-properties>
          case "embedded-text":
            break; // 16.27.4 <number:embedded-text>

          case "body":
          case "电子表格":
            break; // 3.3 16.9.6 19.726.3

          case "forms":
            break; // 12.25.2 13.2
          case "table-column":
            break; // 9.1.6 <table:table-column>
          case "table-header-rows":
            break; // 9.1.7 <table:table-header-rows>
          case "table-rows":
            break; // 9.1.12 <table:table-rows>
          /* TODO: outline levels */
          case "table-column-group":
            break; // 9.1.10 <table:table-column-group>
          case "table-header-columns":
            break; // 9.1.11 <table:table-header-columns>
          case "table-columns":
            break; // 9.1.12 <table:table-columns>

          case "null-date":
            break; // 9.4.2 <table:null-date> TODO: date1904

          case "graphic-properties":
            break; // 17.21 <style:graphic-properties>
          case "calculation-settings":
            break; // 9.4.1 <table:calculation-settings>
          case "named-expressions":
            break; // 9.4.11 <table:named-expressions>
          case "label-range":
            break; // 9.4.9 <table:label-range>
          case "label-ranges":
            break; // 9.4.10 <table:label-ranges>
          case "named-expression":
            break; // 9.4.13 <table:named-expression>
          case "sort":
            break; // 9.4.19 <table:sort>
          case "sort-by":
            break; // 9.4.20 <table:sort-by>
          case "sort-groups":
            break; // 9.4.22 <table:sort-groups>

          case "tab":
            break; // 6.1.4 <text:tab>
          case "line-break":
            break; // 6.1.5 <text:line-break>
          case "span":
            break; // 6.1.7 <text:span>
          case "p":
          case "文本串": // 5.1.3 <text:p>
            if (Rn[1] === "/" && (!ctag || !ctag["string-value"])) {
              const ptp = parse_text_p(str.slice(textpidx, Rn.index), textptag);
              textp = (textp.length > 0 ? `${textp}\n` : "") + ptp[0];
            } else {
              textptag = parsexmltag(Rn[0], false);
              textpidx = Rn.index + Rn[0].length;
            }
            break; // <text:p>
          case "s":
            break; // <text:s>

          case "database-range": // 9.4.15 <table:database-range>
            if (Rn[1] === "/") break;
            try {
              _Ref = ods_to_csf_3D(parsexmltag(Rn[0])["target-range-address"]);
              Sheets[_Ref[0]]["!autofilter"] = { ref: _Ref[1] };
            } catch (e) {
              /* empty */
            }
            break;

          case "date":
            break; // <*:date>

          case "object":
            break; // 10.4.6.2 <draw:object>
          case "title":
          case "标题":
            break; // <*:title> OR <uof:标题>
          case "desc":
            break; // <*:desc>
          case "binary-data":
            break; // 10.4.5 TODO: b64 blob

          /* 9.2 Advanced Tables */
          case "table-source":
            break; // 9.2.6
          case "scenario":
            break; // 9.2.6

          case "iteration":
            break; // 9.4.3 <table:iteration>
          case "content-validations":
            break; // 9.4.4 <table:
          case "content-validation":
            break; // 9.4.5 <table:
          case "help-message":
            break; // 9.4.6 <table:
          case "error-message":
            break; // 9.4.7 <table:
          case "database-ranges":
            break; // 9.4.14 <table:database-ranges>
          case "filter":
            break; // 9.5.2 <table:filter>
          case "filter-and":
            break; // 9.5.3 <table:filter-and>
          case "filter-or":
            break; // 9.5.4 <table:filter-or>
          case "filter-condition":
            break; // 9.5.5 <table:filter-condition>

          case "list-level-style-bullet":
            break; // 16.31 <text:
          case "list-level-style-number":
            break; // 16.32 <text:
          case "list-level-properties":
            break; // 17.19 <style:

          /* 7.3 Document Fields */
          case "sender-firstname": // 7.3.6.2
          case "sender-lastname": // 7.3.6.3
          case "sender-initials": // 7.3.6.4
          case "sender-title": // 7.3.6.5
          case "sender-position": // 7.3.6.6
          case "sender-email": // 7.3.6.7
          case "sender-phone-private": // 7.3.6.8
          case "sender-fax": // 7.3.6.9
          case "sender-company": // 7.3.6.10
          case "sender-phone-work": // 7.3.6.11
          case "sender-street": // 7.3.6.12
          case "sender-city": // 7.3.6.13
          case "sender-postal-code": // 7.3.6.14
          case "sender-country": // 7.3.6.15
          case "sender-state-or-province": // 7.3.6.16
          case "author-name": // 7.3.7.1
          case "author-initials": // 7.3.7.2
          case "chapter": // 7.3.8
          case "file-name": // 7.3.9
          case "template-name": // 7.3.9
          case "sheet-name": // 7.3.9
            break;

          case "event-listener":
            break;
          /* TODO: FODS Properties */
          case "initial-creator":
          case "creation-date":
          case "print-date":
          case "generator":
          case "document-statistic":
          case "user-defined":
          case "editing-duration":
          case "editing-cycles":
            break;

          /* TODO: FODS Config */
          case "config-item":
            break;

          /* TODO: style tokens */
          case "page-number":
            break; // TODO <text:page-number>
          case "page-count":
            break; // TODO <text:page-count>
          case "time":
            break; // TODO <text:time>

          /* 9.3 Advanced Table Cells */
          case "cell-range-source":
            break; // 9.3.1 <table:
          case "detective":
            break; // 9.3.2 <table:
          case "operation":
            break; // 9.3.3 <table:
          case "highlighted-range":
            break; // 9.3.4 <table:

          /* 9.6 Data Pilot Tables <table: */
          case "data-pilot-table": // 9.6.3
          case "source-cell-range": // 9.6.5
          case "source-service": // 9.6.6
          case "data-pilot-field": // 9.6.7
          case "data-pilot-level": // 9.6.8
          case "data-pilot-subtotals": // 9.6.9
          case "data-pilot-subtotal": // 9.6.10
          case "data-pilot-members": // 9.6.11
          case "data-pilot-member": // 9.6.12
          case "data-pilot-display-info": // 9.6.13
          case "data-pilot-sort-info": // 9.6.14
          case "data-pilot-layout-info": // 9.6.15
          case "data-pilot-field-reference": // 9.6.16
          case "data-pilot-groups": // 9.6.17
          case "data-pilot-group": // 9.6.18
          case "data-pilot-group-member": // 9.6.19
            break;

          /* 10.3 Drawing Shapes */
          case "rect": // 10.3.2
            break;

          /* 14.6 DDE Connections */
          case "dde-connection-decls": // 14.6.2 <text:
          case "dde-connection-decl": // 14.6.3 <text:
          case "dde-link": // 14.6.4 <table:
          case "dde-source": // 14.6.5 <office:
            break;

          case "properties":
            break; // 13.7 <form:properties>
          case "property":
            break; // 13.8 <form:property>

          case "a": // 6.1.8 hyperlink
            if (Rn[1] !== "/") {
              atag = parsexmltag(Rn[0], false);
              if (!atag.href) break;
              atag.Target = atag.href;
              delete atag.href;
              if (
                atag.Target.charAt(0) == "#" &&
                atag.Target.indexOf(".") > -1
              ) {
                _Ref = ods_to_csf_3D(atag.Target.slice(1));
                atag.Target = `#${_Ref[0]}!${_Ref[1]}`;
              }
            }
            break;

          /* non-standard */
          case "table-protection":
            break;
          case "data-pilot-grand-total":
            break; // <table:
          case "office-document-common-attrs":
            break; // bare
          default:
            switch (Rn[2]) {
              case "dc:": // TODO: properties
              case "calcext:": // ignore undocumented extensions
              case "loext:": // ignore undocumented extensions
              case "ooo:": // ignore undocumented extensions
              case "chartooo:": // ignore undocumented extensions
              case "draw:": // TODO: drawing
              case "style:": // TODO: styles
              case "chart:": // TODO: charts
              case "form:": // TODO: forms
              case "uof:": // TODO: uof
              case "表:": // TODO: uof
              case "字:": // TODO: uof
                break;
              default:
                if (opts.WTF) throw new Error(Rn);
            }
        }
      const out = {
        Sheets,
        SheetNames,
        Workbook: WB
      };
      if (opts.bookSheets) delete out.Sheets;
      return out;
    };
  })();

  function parse_ods(zip, opts) {
    opts = opts || {};
    const ods = !!safegetzipfile(zip, "objectdata");
    if (ods) parse_manifest(getzipdata(zip, "META-INF/manifest.xml"), opts);
    const content = getzipstr(zip, "content.xml");
    if (!content)
      throw new Error(`Missing content.xml in ${ods ? "ODS" : "UOF"} file`);
    const wb = parse_content_xml(ods ? content : utf8read(content), opts);
    if (safegetzipfile(zip, "meta.xml"))
      wb.Props = parse_core_props(getzipdata(zip, "meta.xml"));
    return wb;
  }
  function parse_fods(data, opts) {
    return parse_content_xml(data, opts);
  }

  function fix_opts_func(defaults) {
    return function fix_opts(opts) {
      for (let i = 0; i != defaults.length; ++i) {
        const d = defaults[i];
        if (opts[d[0]] === undefined) opts[d[0]] = d[1];
        if (d[2] === "n") opts[d[0]] = Number(opts[d[0]]);
      }
    };
  }

  var fix_read_opts = fix_opts_func([
    ["cellNF", false] /* emit cell number format string as .z */,
    ["cellHTML", true] /* emit html string as .h */,
    ["cellFormula", true] /* emit formulae as .f */,
    ["cellStyles", false] /* emits style/theme as .s */,
    ["cellText", true] /* emit formatted text as .w */,
    ["cellDates", false] /* emit date cells with type `d` */,

    ["sheetStubs", false] /* emit empty cells */,
    ["sheetRows", 0, "n"] /* read n rows (0 = read all rows) */,

    ["bookDeps", false] /* parse calculation chains */,
    ["bookSheets", false] /* only try to get sheet names (no Sheets) */,
    ["bookProps", false] /* only try to get properties (no Sheets) */,
    ["bookFiles", false] /* include raw file structure (keys, files, cfb) */,
    ["bookVBA", false] /* include vba raw data (vbaraw) */,

    ["password", ""] /* password */,
    ["WTF", false] /* WTF mode (throws errors) */
  ]);

  var fix_write_opts = fix_opts_func([
    ["cellDates", false] /* write date cells with type `d` */,

    ["bookSST", false] /* Generate Shared String Table */,

    ["bookType", "xlsx"] /* Type of workbook (xlsx/m/b) */,

    ["compression", false] /* Use file compression */,

    ["WTF", false] /* WTF mode (throws errors) */
  ]);
  function get_sheet_type(n) {
    if (RELS.WS.indexOf(n) > -1) return "sheet";
    if (RELS.CS && n == RELS.CS) return "chart";
    if (RELS.DS && n == RELS.DS) return "dialog";
    if (RELS.MS && n == RELS.MS) return "macro";
    return n && n.length ? n : "sheet";
  }
  function safe_parse_wbrels(wbrels, sheets) {
    if (!wbrels) return 0;
    try {
      wbrels = sheets.map(function pwbr(w) {
        if (!w.id) w.id = w.strRelID;
        return [
          w.name,
          wbrels["!id"][w.id].Target,
          get_sheet_type(wbrels["!id"][w.id].Type)
        ];
      });
    } catch (e) {
      return null;
    }
    return !wbrels || wbrels.length === 0 ? null : wbrels;
  }

  function safe_parse_sheet(
    zip,
    path,
    relsPath,
    sheet,
    idx,
    sheetRels,
    sheets,
    stype,
    opts,
    wb,
    themes,
    styles
  ) {
    try {
      sheetRels[sheet] = parse_rels(getzipstr(zip, relsPath, true), path);
      const data = getzipdata(zip, path);
      let _ws;
      switch (stype) {
        case "sheet":
          _ws = parse_ws(
            data,
            path,
            idx,
            opts,
            sheetRels[sheet],
            wb,
            themes,
            styles
          );
          break;
        case "chart":
          _ws = parse_cs(
            data,
            path,
            idx,
            opts,
            sheetRels[sheet],
            wb,
            themes,
            styles
          );
          if (!_ws || !_ws["!drawel"]) break;
          var dfile = resolve_path(_ws["!drawel"].Target, path);
          var drelsp = get_rels_path(dfile);
          var draw = parse_drawing(
            getzipstr(zip, dfile, true),
            parse_rels(getzipstr(zip, drelsp, true), dfile)
          );
          var chartp = resolve_path(draw, dfile);
          var crelsp = get_rels_path(chartp);
          _ws = parse_chart(
            getzipstr(zip, chartp, true),
            chartp,
            opts,
            parse_rels(getzipstr(zip, crelsp, true), chartp),
            wb,
            _ws
          );
          break;
        case "macro":
          _ws = parse_ms(
            data,
            path,
            idx,
            opts,
            sheetRels[sheet],
            wb,
            themes,
            styles
          );
          break;
        case "dialog":
          _ws = parse_ds(
            data,
            path,
            idx,
            opts,
            sheetRels[sheet],
            wb,
            themes,
            styles
          );
          break;
        default:
          throw new Error(`Unrecognized sheet type ${stype}`);
      }
      sheets[sheet] = _ws;

      /* scan rels for comments */
      let comments = [];
      if (sheetRels && sheetRels[sheet])
        keys(sheetRels[sheet]).forEach(function(n) {
          if (sheetRels[sheet][n].Type == RELS.CMNT) {
            const dfile = resolve_path(sheetRels[sheet][n].Target, path);
            comments = parse_cmnt(getzipdata(zip, dfile, true), dfile, opts);
            if (!comments || !comments.length) return;
            sheet_insert_comments(_ws, comments);
          }
        });
    } catch (e) {
      if (opts.WTF) throw e;
    }
  }

  function strip_front_slash(x) {
    return x.charAt(0) == "/" ? x.slice(1) : x;
  }

  function parse_zip(zip, opts) {
    make_ssf(SSF);
    opts = opts || {};
    fix_read_opts(opts);

    /* OpenDocument Part 3 Section 2.2.1 OpenDocument Package */
    if (safegetzipfile(zip, "META-INF/manifest.xml"))
      return parse_ods(zip, opts);
    /* UOC */
    if (safegetzipfile(zip, "objectdata.xml")) return parse_ods(zip, opts);
    /* Numbers */
    if (safegetzipfile(zip, "Index/Document.iwa"))
      throw new Error("Unsupported NUMBERS file");

    const entries = zipentries(zip);
    const dir = parse_ct(getzipstr(zip, "[Content_Types].xml"));
    let xlsb = false;
    let sheets;
    let binname;
    if (dir.workbooks.length === 0) {
      binname = "xl/workbook.xml";
      if (getzipdata(zip, binname, true)) dir.workbooks.push(binname);
    }
    if (dir.workbooks.length === 0) {
      binname = "xl/workbook.bin";
      if (!getzipdata(zip, binname, true))
        throw new Error("Could not find workbook");
      dir.workbooks.push(binname);
      xlsb = true;
    }
    if (dir.workbooks[0].slice(-3) == "bin") xlsb = true;

    let themes = {};
    let styles = {};
    if (!opts.bookSheets && !opts.bookProps) {
      strs = [];
      if (dir.sst)
        try {
          strs = parse_sst(
            getzipdata(zip, strip_front_slash(dir.sst)),
            dir.sst,
            opts
          );
        } catch (e) {
          if (opts.WTF) throw e;
        }

      if (opts.cellStyles && dir.themes.length)
        themes = parse_theme(
          getzipstr(zip, dir.themes[0].replace(/^\//, ""), true) || "",
          dir.themes[0],
          opts
        );

      if (dir.style)
        styles = parse_sty(
          getzipdata(zip, strip_front_slash(dir.style)),
          dir.style,
          themes,
          opts
        );
    }

    /* var externbooks = */ dir.links.forEach(function(link) {
      try {
        const rels = parse_rels(
          getzipstr(zip, get_rels_path(strip_front_slash(link))),
          link
        );
        return parse_xlink(
          getzipdata(zip, strip_front_slash(link)),
          rels,
          link,
          opts
        );
      } catch (e) {}
    });

    const wb = parse_wb(
      getzipdata(zip, strip_front_slash(dir.workbooks[0])),
      dir.workbooks[0],
      opts
    );

    let props = {};
    let propdata = "";

    if (dir.coreprops.length) {
      propdata = getzipdata(zip, strip_front_slash(dir.coreprops[0]), true);
      if (propdata) props = parse_core_props(propdata);
      if (dir.extprops.length !== 0) {
        propdata = getzipdata(zip, strip_front_slash(dir.extprops[0]), true);
        if (propdata) parse_ext_props(propdata, props, opts);
      }
    }

    let custprops = {};
    if (!opts.bookSheets || opts.bookProps) {
      if (dir.custprops.length !== 0) {
        propdata = getzipstr(zip, strip_front_slash(dir.custprops[0]), true);
        if (propdata) custprops = parse_cust_props(propdata, opts);
      }
    }

    let out = {};
    if (opts.bookSheets || opts.bookProps) {
      if (wb.Sheets)
        sheets = wb.Sheets.map(function pluck(x) {
          return x.name;
        });
      else if (props.Worksheets && props.SheetNames.length > 0)
        sheets = props.SheetNames;
      if (opts.bookProps) {
        out.Props = props;
        out.Custprops = custprops;
      }
      if (opts.bookSheets && typeof sheets !== "undefined")
        out.SheetNames = sheets;
      if (opts.bookSheets ? out.SheetNames : opts.bookProps) return out;
    }
    sheets = {};

    let deps = {};
    if (opts.bookDeps && dir.calcchain)
      deps = parse_cc(
        getzipdata(zip, strip_front_slash(dir.calcchain)),
        dir.calcchain,
        opts
      );

    let i = 0;
    const sheetRels = {};
    let path;
    let relsPath;

    {
      const wbsheets = wb.Sheets;
      props.Worksheets = wbsheets.length;
      props.SheetNames = [];
      for (let j = 0; j != wbsheets.length; ++j) {
        props.SheetNames[j] = wbsheets[j].name;
      }
    }

    const wbext = xlsb ? "bin" : "xml";
    const wbrelsi = dir.workbooks[0].lastIndexOf("/");
    let wbrelsfile = `${dir.workbooks[0].slice(
      0,
      wbrelsi + 1
    )}_rels/${dir.workbooks[0].slice(wbrelsi + 1)}.rels`.replace(/^\//, "");
    if (!safegetzipfile(zip, wbrelsfile))
      wbrelsfile = `xl/_rels/workbook.${wbext}.rels`;
    let wbrels = parse_rels(getzipstr(zip, wbrelsfile, true), wbrelsfile);
    if (wbrels) wbrels = safe_parse_wbrels(wbrels, wb.Sheets);

    /* Numbers iOS hack */
    const nmode = getzipdata(zip, "xl/worksheets/sheet.xml", true) ? 1 : 0;
    wsloop: for (i = 0; i != props.Worksheets; ++i) {
      let stype = "sheet";
      if (wbrels && wbrels[i]) {
        path = `xl/${wbrels[i][1].replace(/[\/]?xl\//, "")}`;
        if (!safegetzipfile(zip, path)) path = wbrels[i][1];
        if (!safegetzipfile(zip, path))
          path = wbrelsfile.replace(/_rels\/.*$/, "") + wbrels[i][1];
        stype = wbrels[i][2];
      } else {
        path = `xl/worksheets/sheet${i + 1 - nmode}.${wbext}`;
        path = path.replace(/sheet0\./, "sheet.");
      }
      relsPath = path.replace(/^(.*)(\/)([^\/]*)$/, "$1/_rels/$3.rels");
      if (opts && opts.sheets != null)
        switch (typeof opts.sheets) {
          case "number":
            if (i != opts.sheets) continue wsloop;
            break;
          case "string":
            if (props.SheetNames[i].toLowerCase() != opts.sheets.toLowerCase())
              continue wsloop;
            break;
          default:
            if (Array.isArray && Array.isArray(opts.sheets)) {
              let snjseen = false;
              for (let snj = 0; snj != opts.sheets.length; ++snj) {
                if (
                  typeof opts.sheets[snj] === "number" &&
                  opts.sheets[snj] == i
                )
                  snjseen = 1;
                if (
                  typeof opts.sheets[snj] === "string" &&
                  opts.sheets[snj].toLowerCase() ==
                    props.SheetNames[i].toLowerCase()
                )
                  snjseen = 1;
              }
              if (!snjseen) continue wsloop;
            }
        }
      safe_parse_sheet(
        zip,
        path,
        relsPath,
        props.SheetNames[i],
        i,
        sheetRels,
        sheets,
        stype,
        opts,
        wb,
        themes,
        styles
      );
    }

    out = {
      Directory: dir,
      Workbook: wb,
      Props: props,
      Custprops: custprops,
      Deps: deps,
      Sheets: sheets,
      SheetNames: props.SheetNames,
      Strings: strs,
      Styles: styles,
      Themes: themes,
      SSF: SSF.get_table()
    };
    if (opts && opts.bookFiles) {
      out.keys = entries;
      out.files = zip.files;
    }
    if (opts && opts.bookVBA) {
      if (dir.vba.length > 0)
        out.vbaraw = getzipdata(zip, strip_front_slash(dir.vba[0]), true);
      else if (dir.defaults && dir.defaults.bin === CT_VBA)
        out.vbaraw = getzipdata(zip, "xl/vbaProject.bin", true);
    }
    return out;
  }

  /* [MS-OFFCRYPTO] 2.1.1 */
  function parse_xlsxcfb(cfb, _opts) {
    const opts = _opts || {};
    let f = "Workbook";
    let data = CFB.find(cfb, f);
    try {
      f = "/!DataSpaces/Version";
      data = CFB.find(cfb, f);
      if (!data || !data.content)
        throw new Error(`ECMA-376 Encrypted file missing ${f}`);
      /* var version = */ parse_DataSpaceVersionInfo(data.content);

      /* 2.3.4.1 */
      f = "/!DataSpaces/DataSpaceMap";
      data = CFB.find(cfb, f);
      if (!data || !data.content)
        throw new Error(`ECMA-376 Encrypted file missing ${f}`);
      const dsm = parse_DataSpaceMap(data.content);
      if (
        dsm.length !== 1 ||
        dsm[0].comps.length !== 1 ||
        dsm[0].comps[0].t !== 0 ||
        dsm[0].name !== "StrongEncryptionDataSpace" ||
        dsm[0].comps[0].v !== "EncryptedPackage"
      )
        throw new Error(`ECMA-376 Encrypted file bad ${f}`);

      /* 2.3.4.2 */
      f = "/!DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace";
      data = CFB.find(cfb, f);
      if (!data || !data.content)
        throw new Error(`ECMA-376 Encrypted file missing ${f}`);
      const seds = parse_DataSpaceDefinition(data.content);
      if (seds.length != 1 || seds[0] != "StrongEncryptionTransform")
        throw new Error(`ECMA-376 Encrypted file bad ${f}`);

      /* 2.3.4.3 */
      f = "/!DataSpaces/TransformInfo/StrongEncryptionTransform/!Primary";
      data = CFB.find(cfb, f);
      if (!data || !data.content)
        throw new Error(`ECMA-376 Encrypted file missing ${f}`);
      /* var hdr = */ parse_Primary(data.content);
    } catch (e) {}

    f = "/EncryptionInfo";
    data = CFB.find(cfb, f);
    if (!data || !data.content)
      throw new Error(`ECMA-376 Encrypted file missing ${f}`);
    const einfo = parse_EncryptionInfo(data.content);

    /* 2.3.4.4 */
    f = "/EncryptedPackage";
    data = CFB.find(cfb, f);
    if (!data || !data.content)
      throw new Error(`ECMA-376 Encrypted file missing ${f}`);

    /* global decrypt_agile */
    if (einfo[0] == 0x04 && typeof decrypt_agile !== "undefined")
      return decrypt_agile(einfo[1], data.content, opts.password || "", opts);
    /* global decrypt_std76 */
    if (einfo[0] == 0x02 && typeof decrypt_std76 !== "undefined")
      return decrypt_std76(einfo[1], data.content, opts.password || "", opts);
    throw new Error("File is password-protected");
  }

  function firstbyte(f, o) {
    let x = "";
    switch ((o || {}).type || "base64") {
      case "buffer":
        return [f[0], f[1], f[2], f[3], f[4], f[5], f[6], f[7]];
      case "base64":
        x = Base64.decode(f.slice(0, 12));
        break;
      case "binary":
        x = f;
        break;
      case "array":
        return [f[0], f[1], f[2], f[3], f[4], f[5], f[6], f[7]];
      default:
        throw new Error(`Unrecognized type ${(o && o.type) || "undefined"}`);
    }
    return [
      x.charCodeAt(0),
      x.charCodeAt(1),
      x.charCodeAt(2),
      x.charCodeAt(3),
      x.charCodeAt(4),
      x.charCodeAt(5),
      x.charCodeAt(6),
      x.charCodeAt(7)
    ];
  }

  function read_cfb(cfb, opts) {
    if (CFB.find(cfb, "EncryptedPackage")) return parse_xlsxcfb(cfb, opts);
    return parse_xlscfb(cfb, opts);
  }

  function read_zip(data, opts) {
    let zip;
    const d = data;
    const o = opts || {};
    if (!o.type)
      o.type = has_buf && Buffer.isBuffer(data) ? "buffer" : "base64";
    zip = zip_read(d, o);
    return parse_zip(zip, o);
  }

  function read_plaintext(data, o) {
    let i = 0;
    main: while (i < data.length)
      switch (data.charCodeAt(i)) {
        case 0x0a:
        case 0x0d:
        case 0x20:
          ++i;
          break;
        case 0x3c:
          return parse_xlml(data.slice(i), o);
        default:
          break main;
      }
    return PRN.to_workbook(data, o);
  }

  function read_plaintext_raw(data, o) {
    let str = "";
    const bytes = firstbyte(data, o);
    switch (o.type) {
      case "base64":
        str = Base64.decode(data);
        break;
      case "binary":
        str = data;
        break;
      case "buffer":
        str = data.toString("binary");
        break;
      case "array":
        str = cc2str(data);
        break;
      default:
        throw new Error(`Unrecognized type ${o.type}`);
    }
    if (bytes[0] == 0xef && bytes[1] == 0xbb && bytes[2] == 0xbf)
      str = utf8read(str);
    return read_plaintext(str, o);
  }

  function read_utf16(data, o) {
    let d = data;
    if (o.type == "base64") d = Base64.decode(d);
    d = cptable.utils.decode(1200, d.slice(2), "str");
    o.type = "binary";
    return read_plaintext(d, o);
  }

  function bstrify(data) {
    return !data.match(/[^\x00-\x7F]/) ? data : utf8write(data);
  }

  function read_prn(data, d, o, str) {
    if (str) {
      o.type = "string";
      return PRN.to_workbook(data, o);
    }
    return PRN.to_workbook(d, o);
  }

  function readSync(data, opts) {
    reset_cp();
    if (typeof ArrayBuffer !== "undefined" && data instanceof ArrayBuffer)
      return readSync(new Uint8Array(data), opts);
    let d = data;
    let n = [0, 0, 0, 0];
    let str = false;
    let o = opts || {};
    if (o.cellStyles) {
      o.cellNF = true;
      o.sheetStubs = true;
    }
    _ssfopts = {};
    if (o.dateNF) _ssfopts.dateNF = o.dateNF;
    if (!o.type)
      o.type = has_buf && Buffer.isBuffer(data) ? "buffer" : "base64";
    if (o.type == "file") {
      o.type = has_buf ? "buffer" : "binary";
      d = read_binary(data);
    }
    if (o.type == "string") {
      str = true;
      o.type = "binary";
      o.codepage = 65001;
      d = bstrify(data);
    }
    if (
      o.type == "array" &&
      typeof Uint8Array !== "undefined" &&
      data instanceof Uint8Array &&
      typeof ArrayBuffer !== "undefined"
    ) {
      // $FlowIgnore
      const ab = new ArrayBuffer(3);
      const vu = new Uint8Array(ab);
      vu.foo = "bar";
      // $FlowIgnore
      if (!vu.foo) {
        o = dup(o);
        o.type = "array";
        return readSync(ab2a(d), o);
      }
    }
    switch ((n = firstbyte(d, o))[0]) {
      case 0xd0:
        if (
          n[1] === 0xcf &&
          n[2] === 0x11 &&
          n[3] === 0xe0 &&
          n[4] === 0xa1 &&
          n[5] === 0xb1 &&
          n[6] === 0x1a &&
          n[7] === 0xe1
        )
          return read_cfb(CFB.read(d, o), o);
        break;
      case 0x09:
        if (n[1] <= 0x04) return parse_xlscfb(d, o);
        break;
      case 0x3c:
        return parse_xlml(d, o);
      case 0x49:
        if (n[1] === 0x44) return read_wb_ID(d, o);
        break;
      case 0x54:
        if (n[1] === 0x41 && n[2] === 0x42 && n[3] === 0x4c)
          return DIF.to_workbook(d, o);
        break;
      case 0x50:
        return n[1] === 0x4b && n[2] < 0x09 && n[3] < 0x09
          ? read_zip(d, o)
          : read_prn(data, d, o, str);
      case 0xef:
        return n[3] === 0x3c ? parse_xlml(d, o) : read_prn(data, d, o, str);
      case 0xff:
        if (n[1] === 0xfe) {
          return read_utf16(d, o);
        }
        break;
      case 0x00:
        if (n[1] === 0x00 && n[2] >= 0x02 && n[3] === 0x00)
          return WK_.to_workbook(d, o);
        break;
      case 0x03:
      case 0x83:
      case 0x8b:
      case 0x8c:
        return DBF.to_workbook(d, o);
      case 0x7b:
        if (n[1] === 0x5c && n[2] === 0x72 && n[3] === 0x74)
          return RTF.to_workbook(d, o);
        break;
      case 0x0a:
      case 0x0d:
      case 0x20:
        return read_plaintext_raw(d, o);
    }
    if (DBF.versions.indexOf(n[0]) > -1 && n[2] <= 12 && n[3] <= 31)
      return DBF.to_workbook(d, o);
    return read_prn(data, d, o, str);
  }

  function make_json_row(sheet, r, R, cols, header, hdr, dense, o) {
    const rr = encode_row(R);
    const { defval } = o;
    const raw = o.raw || !Object.prototype.hasOwnProperty.call(o, "raw");
    let isempty = true;
    const row = header === 1 ? [] : {};
    if (header !== 1) {
      if (Object.defineProperty)
        try {
          Object.defineProperty(row, "__rowNum__", {
            value: R,
            enumerable: false
          });
        } catch (e) {
          row.__rowNum__ = R;
        }
      else row.__rowNum__ = R;
    }
    if (!dense || sheet[R])
      for (let C = r.s.c; C <= r.e.c; ++C) {
        const val = dense ? sheet[R][C] : sheet[cols[C] + rr];
        if (val === undefined || val.t === undefined) {
          if (defval === undefined) continue;
          if (hdr[C] != null) {
            row[hdr[C]] = defval;
          }
          continue;
        }
        let { v } = val;
        switch (val.t) {
          case "z":
            if (v == null) break;
            continue;
          case "e":
            v = void 0;
            break;
          case "s":
          case "d":
          case "b":
          case "n":
            break;
          default:
            throw new Error(`unrecognized type ${val.t}`);
        }
        if (hdr[C] != null) {
          if (v == null) {
            if (defval !== undefined) row[hdr[C]] = defval;
            else if (raw && v === null) row[hdr[C]] = null;
            else continue;
          } else {
            row[hdr[C]] =
              raw || (o.rawNumbers && val.t == "n")
                ? v
                : format_cell(val, v, o);
          }
          if (v != null) isempty = false;
        }
      }
    return { row, isempty };
  }

  function sheet_to_json(sheet, opts) {
    if (sheet == null || sheet["!ref"] == null) return [];
    let val = { t: "n", v: 0 };
    let header = 0;
    let offset = 1;
    const hdr = [];
    let v = 0;
    let vv = "";
    let r = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
    const o = opts || {};
    const range = o.range != null ? o.range : sheet["!ref"];
    if (o.header === 1) header = 1;
    else if (o.header === "A") header = 2;
    else if (Array.isArray(o.header)) header = 3;
    else if (o.header == null) header = 0;
    switch (typeof range) {
      case "string":
        r = safe_decode_range(range);
        break;
      case "number":
        r = safe_decode_range(sheet["!ref"]);
        r.s.r = range;
        break;
      default:
        r = range;
    }
    if (header > 0) offset = 0;
    const rr = encode_row(r.s.r);
    const cols = [];
    const out = [];
    let outi = 0;
    let counter = 0;
    const dense = Array.isArray(sheet);
    let R = r.s.r;
    let C = 0;
    let CC = 0;
    if (dense && !sheet[R]) sheet[R] = [];
    for (C = r.s.c; C <= r.e.c; ++C) {
      cols[C] = encode_col(C);
      val = dense ? sheet[R][C] : sheet[cols[C] + rr];
      switch (header) {
        case 1:
          hdr[C] = C - r.s.c;
          break;
        case 2:
          hdr[C] = cols[C];
          break;
        case 3:
          hdr[C] = o.header[C - r.s.c];
          break;
        default:
          if (val == null) val = { w: "__EMPTY", t: "s" };
          vv = v = format_cell(val, null, o);
          counter = 0;
          for (CC = 0; CC < hdr.length; ++CC)
            if (hdr[CC] == vv) vv = `${v}_${++counter}`;
          hdr[C] = vv;
      }
    }
    for (R = r.s.r + offset; R <= r.e.r; ++R) {
      const row = make_json_row(sheet, r, R, cols, header, hdr, dense, o);
      if (
        row.isempty === false ||
        (header === 1 ? o.blankrows !== false : !!o.blankrows)
      )
        out[outi++] = row.row;
    }
    out.length = outi;
    return out;
  }

  const qreg = /"/g;

  XLSX.read = readSync;
}
make_xlsx_lib(module.exports);
