// iPOS WhatsApp Daily Sales Report
// Engine : Windows Script Host JScript (cscript.exe) - built into Windows XP+
// Usage  : cscript //nologo send_report.js
// No external dependencies required.

// ---- DEBUG MODE ----
// Set to true to print raw SQL, raw output, and skip sending to WhatsApp.
// Set to false for normal production use.
var DEBUG = true

var env = WScript.CreateObject("WScript.Shell").Environment("Process");
var fso = WScript.CreateObject("Scripting.FileSystemObject");
var SCRIPT_DIR = fso.GetParentFolderName(WScript.ScriptFullName);

// ---- Config (set by config.bat before calling this script) ----
var DB_HOST = env("DB_HOST") || "192.168.1.120";
var DB_PORT = env("DB_PORT") || "5444";
var DB_NAME = env("DB_NAME") || "i4_SlametStore";
var DB_USER = env("DB_USER");
var DB_PASS = env("DB_PASS");
var PSQL = env("PSQL") || "D:\\Kasir Slamet\\ipgsql\\bin\\psql.exe";
var WHAPI_TOKEN = env("WHAPI_TOKEN") || "";
var WHAPI_GATE = env("WHAPI_GATE") || "";

// Optional overrides:
//   REPORT_DATE=YYYY-MM-DD       run for an exact date
//   REPORT_DAYS_AGO=N            run for N days ago (1 = yesterday)
// Both empty/invalid => yesterday (production default).
var REPORT_DATE = env("REPORT_DATE") || "";
var REPORT_DAYS_AGO = parseInt(env("REPORT_DAYS_AGO") || "0", 10);
var REPORT_DATE_OBJ;
var DATE_FILTER;
if (/^\d{4}-\d{2}-\d{2}$/.test(REPORT_DATE)) {
    var _dp = REPORT_DATE.split("-");
    REPORT_DATE_OBJ = new Date(parseInt(_dp[0], 10), parseInt(_dp[1], 10) - 1, parseInt(_dp[2], 10));
    DATE_FILTER = "tanggal >= DATE '" + REPORT_DATE + "' + interval '1 minute' "
        + "AND tanggal < DATE '" + REPORT_DATE + "' + interval '1 day'";
} else if (REPORT_DAYS_AGO > 0) {
    REPORT_DATE_OBJ = new Date();
    REPORT_DATE_OBJ.setDate(REPORT_DATE_OBJ.getDate() - REPORT_DAYS_AGO);
    DATE_FILTER = "tanggal >= CURRENT_DATE - interval '" + REPORT_DAYS_AGO + " day' + interval '1 minute' "
        + "AND tanggal < CURRENT_DATE - interval '" + (REPORT_DAYS_AGO - 1) + " day'";
} else {
    REPORT_DATE_OBJ = new Date();
    REPORT_DATE_OBJ.setDate(REPORT_DATE_OBJ.getDate() - 1);
    DATE_FILTER = "tanggal >= CURRENT_DATE - interval '1 day' + interval '1 minute' "
        + "AND tanggal < CURRENT_DATE";
}

// Exclude transactions from ADMIN cashier (user1 is the operator).
// Override via env EXCLUDE_USER to a different name, or set EXCLUDE_USER=NONE to disable.
// (env() returns "" for unset vars in WSH, so we default empty -> "ADMIN".)
var EXCLUDE_USER = trim(env("EXCLUDE_USER") || "ADMIN");
if (EXCLUDE_USER.toUpperCase() === "NONE") EXCLUDE_USER = "";
var USER_FILTER = EXCLUDE_USER
    ? "UPPER(COALESCE(user1,'')) <> '" + EXCLUDE_USER.replace(/'/g, "''").toUpperCase() + "'"
    : "";

// Build recipients list from WHAPI_TO_1 .. WHAPI_TO_N (set WHAPI_TO_COUNT in config.bat).
// Falls back to plain WHAPI_TO for backwards compatibility.
var WHAPI_RECIPIENTS = [];
var _rCount = parseInt(env("WHAPI_TO_COUNT") || "0");
if (_rCount > 0) {
    for (var _i = 1; _i <= _rCount; _i++) {
        var _num = env("WHAPI_TO_" + _i);
        if (_num) WHAPI_RECIPIENTS.push(_num);
    }
} else {
    var _single = env("WHAPI_TO") || "";
    if (_single) WHAPI_RECIPIENTS.push(_single);
}

// ---- Logging ----
var logPath = SCRIPT_DIR + "\\report_log.txt";

function log(msg) {
    var line = "[" + new Date().toLocaleString() + "] " + msg;
    try {
        var ts = fso.OpenTextFile(logPath, 8, true); // 8 = append, true = create if missing
        ts.WriteLine(line);
        ts.Close();
    } catch (e) { /* ignore log errors */ }
    WScript.Echo(line);
}

function dbg(msg) {
    if (!DEBUG) return;
    var line = "[DBG] " + msg;
    WScript.Echo(line);
    try {
        var ts = fso.OpenTextFile(logPath, 8, true);
        ts.WriteLine(line);
        ts.Close();
    } catch (e) { }
}

// ---- Helpers ----
function formatRp(n) {
    var num = Math.round(parseFloat(n) || 0);
    if (num === 0) return "Rp 0";
    var str = num.toString();
    var result = "";
    var count = 0;
    for (var i = str.length - 1; i >= 0; i--) {
        if (count > 0 && count % 3 === 0) result = "." + result;
        result = str.charAt(i) + result;  // charAt — str[i] not supported in old JScript
        count++;
    }
    return "Rp " + result;
}

function trim(s) {
    return String(s).replace(/^\s+|\s+$/g, "");
}

function pad2(n) {
    return n < 10 ? "0" + n : "" + n;
}

function formatTanggal(d) {
    var days = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
    var months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    return days[d.getDay()] + ", " + d.getDate() + " " + months[d.getMonth()] + " " + d.getFullYear();
}

function jsonEscape(str) {
    return String(str)
        .replace(/\\/g, "\\\\")
        .replace(/"/g, '\\"')
        .replace(/\n/g, "\\n")
        .replace(/\r/g, "")
        .replace(/\t/g, " ");
}

// ---- Run psql query, return stdout as string ----
// Uses Shell.Run (truly synchronous) + temp output file to avoid the
// WSH Shell.Exec AtEndOfStream premature-EOF bug.
// Temp files go to %TEMP% (always a local C:\ path) so psql can read them
// even when this script lives on a mapped/network drive.
function runPsql(sql) {
    var shell = WScript.CreateObject("WScript.Shell");
    shell.Environment("Process").Item("PGPASSWORD") = DB_PASS;

    var tmpDir = shell.ExpandEnvironmentStrings("%TEMP%");
    var tmpSql = tmpDir + "\\ipos_query.sql";
    var tmpOut = tmpDir + "\\ipos_out.txt";

    // Write SQL to temp file
    try {
        var tf = fso.CreateTextFile(tmpSql, true);
        tf.Write(sql);
        tf.Close();
    } catch (e) {
        throw new Error("Cannot write temp SQL file: " + e.message);
    }

    var cmd = '"' + PSQL + '"'
        + ' -h ' + DB_HOST
        + ' -p ' + DB_PORT
        + ' -U ' + DB_USER
        + ' -d ' + DB_NAME
        + ' -t -A'               // | is default separator in -A mode; omit -F to avoid | quoting issues
        + ' -f "' + tmpSql + '"'
        + ' >"' + tmpOut + '" 2>&1';

    // cmd.exe /c strips the FIRST and LAST " in the argument string.
    // Wrapping the whole command in an extra pair of outer quotes means after
    // stripping, the inner "psql.exe" ... paths are left properly quoted.
    var comspec = shell.ExpandEnvironmentStrings("%ComSpec%");
    var fullCmd = comspec + ' /c "' + cmd + '"';

    dbg("--- runPsql ---");
    dbg("SQL : " + sql);
    dbg("CMD : " + fullCmd);

    // Shell.Run with bWaitOnReturn=true blocks until the process exits
    var exitCode = shell.Run(fullCmd, 0, true); // 0=hidden window, true=wait
    dbg("Exit: " + exitCode);

    // Read the output file
    var out = "";
    try {
        if (fso.FileExists(tmpOut)) {
            var f = fso.OpenTextFile(tmpOut, 1); // 1=ForReading
            if (!f.AtEndOfStream) out = f.ReadAll();
            f.Close();
        } else {
            dbg("WARN: output file not created: " + tmpOut);
        }
    } catch (e) {
        dbg("Error reading output file: " + e.message);
    }

    try { fso.DeleteFile(tmpSql); } catch (e) { }
    try { fso.DeleteFile(tmpOut); } catch (e) { }

    var result = trim(out);
    dbg("RAW(" + result.length + "): [" + result + "]");
    return result;
}

// ---- Detect psql connection-level errors in output ----
// Only connection errors are retried; auth/query errors are not.
function isPsqlConnError(s) {
    return /could not connect to server|Connection timed out|Connection refused/i.test(s);
}

// ---- Run psql with auto-retry every 60 s until DB is reachable ----
function runPsqlRetry(sql) {
    var attempt = 0;
    while (true) {
        attempt++;
        var out = runPsql(sql);
        if (!isPsqlConnError(out)) return out;
        log("DB unreachable (attempt " + attempt + "). Retrying in 60 seconds...");
        WScript.Sleep(15000);
    }
}

// ---- Upload a JSON string to catbox.moe (or litterbox for ephemeral) ----
// Target is chosen by mode:
//   "permanent" -> https://catbox.moe/user/api.php          (no expiry)
//   "72h"/"24h"/"12h"/"1h" -> litterbox (temporary bucket)
// Returns the resulting URL. Uses multipart/form-data with a text body.
function uploadJsonCatbox(jsonText, filename, mode) {
    var url, fields;
    if (mode === "permanent" || !mode) {
        url = "https://catbox.moe/user/api.php";
        fields = { reqtype: "fileupload" };
    } else {
        url = "https://litterbox.catbox.moe/resources/internals/api.php";
        fields = { reqtype: "fileupload", time: mode };
    }

    var boundary = "----iposReport" + new Date().getTime();
    var CRLF = "\r\n";
    var body = "";
    for (var k in fields) {
        body += "--" + boundary + CRLF
            + 'Content-Disposition: form-data; name="' + k + '"' + CRLF + CRLF
            + fields[k] + CRLF;
    }
    body += "--" + boundary + CRLF
        + 'Content-Disposition: form-data; name="fileToUpload"; filename="' + filename + '"' + CRLF
        + "Content-Type: application/json" + CRLF + CRLF
        + jsonText + CRLF
        + "--" + boundary + "--" + CRLF;

    var http = WScript.CreateObject("WinHttp.WinHttpRequest.5.1");
    http.Open("POST", url, false);
    http.SetRequestHeader("Content-Type", "multipart/form-data; boundary=" + boundary);
    http.SetRequestHeader("User-Agent", "iPOS-Reporter/1.0");
    http.Send(body);

    if (http.Status < 200 || http.Status >= 300) {
        throw new Error("Upload HTTP " + http.Status + " - " + String(http.ResponseText).substring(0, 200));
    }
    var resp = trim(http.ResponseText);
    if (!/^https?:\/\//i.test(resp)) {
        throw new Error("Unexpected upload response: " + resp.substring(0, 200));
    }
    return resp;
}

// ---- Send WhatsApp message to a single recipient via WHAPI ----
function sendWhatsapp(message, to) {
    if (!WHAPI_TOKEN) throw new Error("WHAPI_TOKEN is not set in config.bat");
    if (!WHAPI_GATE) throw new Error("WHAPI_GATE is not set in config.bat");
    if (!to) throw new Error("No recipient number provided");

    var gate = WHAPI_GATE.replace(/\s/g, "").replace(/\/$/, "");
    var toDigits = to.replace(/\D/g, "");
    var safe = message.length > 4090 ? message.substring(0, 4087) + "..." : message;
    var body = '{"to":"' + toDigits + '","body":"' + jsonEscape(safe) + '"}';

    var http = WScript.CreateObject("WinHttp.WinHttpRequest.5.1");
    http.Open("POST", gate + "/messages/text", false);
    http.SetRequestHeader("Authorization", "Bearer " + WHAPI_TOKEN);
    http.SetRequestHeader("Content-Type", "application/json; charset=utf-8");
    http.SetRequestHeader("Accept", "application/json");
    http.Send(body);

    if (http.Status < 200 || http.Status >= 300) {
        throw new Error("WHAPI HTTP " + http.Status + " - " + http.ResponseText.substring(0, 300));
    }
    return http.ResponseText;
}

// ================================================================
//  MAIN
// ================================================================
try {
    log("=== iPOS WhatsApp Report START === (DEBUG=" + DEBUG + ")");
    log("Excluding user1: " + (EXCLUDE_USER || "(none)"));
    dbg("Config: " + DB_USER + "@" + DB_HOST + ":" + DB_PORT + "/" + DB_NAME);
    dbg("PSQL  : " + PSQL);
    dbg("PSQL exists: " + fso.FileExists(PSQL));

    // 1. Store name
    log("Query 1: tbl_kantor");
    var tokoOut = runPsqlRetry("SELECT namakantor, alamat FROM tbl_kantor LIMIT 1");
    var tokoParts = tokoOut ? tokoOut.split("|") : [];
    var namaToko = tokoParts[0] || "";
    var alamatToko = tokoParts[1] || "";
    log("Store: [" + namaToko + "] [" + alamatToko + "]");

    // 2. Daily summary
    log("Query 2: daily summary");
    var sumOut = runPsqlRetry(
        "SELECT COUNT(*), " +
        "COALESCE(SUM(totalitem),0), " +
        "COALESCE(SUM(subtotal),0), " +
        "COALESCE(SUM(potfaktur),0), " +
        "COALESCE(SUM(pajak),0), " +
        "COALESCE(SUM(totalakhir),0) " +
        "FROM tbl_ikhd " +
        "WHERE " + DATE_FILTER + (USER_FILTER ? " AND " + USER_FILTER : "")
    );
    log("sumOut raw: [" + sumOut + "]");
    var sp = sumOut ? sumOut.split("|") : ["0", "0", "0", "0", "0", "0"];
    var totalTrx = parseInt(sp[0]) || 0;
    var totalItem = parseInt(sp[1]) || 0;
    var subtotal = parseFloat(sp[2]) || 0;
    var totalDiskon = parseFloat(sp[3]) || 0;
    var totalPajak = parseFloat(sp[4]) || 0;
    var totalAkhir = parseFloat(sp[5]) || 0;

    // Quick sanity: count ALL rows to confirm psql is working at all
    log("Query 2b: total row count sanity check");
    var totalRowsOut = runPsqlRetry("SELECT COUNT(*) FROM tbl_ikhd");
    log("Total rows in tbl_ikhd (all time): " + totalRowsOut);

    // Check what the earliest and latest tanggal look like
    log("Query 2c: date range sanity check");
    var dateRangeOut = runPsqlRetry("SELECT MIN(tanggal), MAX(tanggal), CURRENT_DATE FROM tbl_ikhd LIMIT 1");
    log("Date range (min|max|today): " + dateRangeOut);

    // 3. Payment method breakdown
    log("Query 3: payment breakdown");
    var bayarOut = runPsqlRetry(
        "SELECT COALESCE(carabayar,'Lainnya'), COUNT(*), COALESCE(SUM(totalakhir),0) " +
        "FROM tbl_ikhd " +
        "WHERE " + DATE_FILTER + (USER_FILTER ? " AND " + USER_FILTER : "") + " " +
        "GROUP BY carabayar ORDER BY 3 DESC"
    );
    log("bayarOut raw: [" + bayarOut + "]");

    // 4. Hourly breakdown
    log("Query 4: hourly");
    var hourlyOut = runPsqlRetry(
        "SELECT EXTRACT(HOUR FROM tanggal)::int, COUNT(*), COALESCE(SUM(totalakhir),0) " +
        "FROM tbl_ikhd " +
        "WHERE " + DATE_FILTER + (USER_FILTER ? " AND " + USER_FILTER : "") + " " +
        "GROUP BY 1 ORDER BY 1"
    );
    log("hourlyOut raw: [" + hourlyOut + "]");

    // 5. Top 5 customers
    log("Query 5: top customers");
    var topOut = runPsqlRetry(
        "SELECT COALESCE(s.nama, h.kodesupel, 'Umum'), COUNT(*), COALESCE(SUM(h.totalakhir),0) " +
        "FROM tbl_ikhd h " +
        "LEFT JOIN tbl_supel s ON s.kode = h.kodesupel " +
        "WHERE " + DATE_FILTER.replace(/tanggal/g, "h.tanggal") + " " +
        "AND h.kodesupel IS NOT NULL AND h.kodesupel <> '' " +
        (USER_FILTER ? "AND " + USER_FILTER.replace(/user1/g, "h.user1") + " " : "") +
        "GROUP BY s.nama, h.kodesupel ORDER BY 3 DESC LIMIT 5"
    );
    log("topOut raw: [" + topOut + "]");

    // 6. Full raw records (header + line items) via CSV COPY.
    // This Postgres is 8.4 — no row_to_json/json_agg, so we parse CSV client-side.
    log("Query 6: raw header records (tbl_ikhd)");
    var hdrCsv = runPsqlRetry(
        "COPY (SELECT * FROM tbl_ikhd WHERE " + DATE_FILTER +
        (USER_FILTER ? " AND " + USER_FILTER : "") +
        " ORDER BY tanggal, notransaksi) TO STDOUT WITH CSV HEADER"
    );
    log("hdrCsv length: " + hdrCsv.length);

    log("Query 7: raw item records (tbl_ikdt + tbl_item.namaitem)");
    var dtlCsv = runPsqlRetry(
        "COPY (SELECT d.*, i.namaitem FROM tbl_ikdt d " +
        "JOIN tbl_ikhd h USING (notransaksi) " +
        "LEFT JOIN tbl_item i ON i.kodeitem = d.kodeitem " +
        "WHERE " + DATE_FILTER.replace(/tanggal/g, "h.tanggal") +
        (USER_FILTER ? " AND " + USER_FILTER.replace(/user1/g, "h.user1") : "") +
        " ORDER BY d.notransaksi, d.nobaris) TO STDOUT WITH CSV HEADER"
    );
    log("dtlCsv length: " + dtlCsv.length);

    log("Data fetched. Transaksi=" + totalTrx + " Total=" + formatRp(totalAkhir));

    // ---- Build full JSON + upload FIRST so the viewer URL can be embedded in the WhatsApp message ----
    var reportDateStr = REPORT_DATE_OBJ.getFullYear() + "-"
        + pad2(REPORT_DATE_OBJ.getMonth() + 1) + "-"
        + pad2(REPORT_DATE_OBJ.getDate());

    var rawRecordsJson = buildRecordsJson(hdrCsv, dtlCsv);

    var paymentJson = rowsToJsonArray(bayarOut, function (p) {
        return '{"carabayar":' + jsonStr(p[0] || "Lainnya")
            + ',"count":' + (parseInt(p[1], 10) || 0)
            + ',"total":' + jsonNum(p[2]) + '}';
    });
    var hourlyJson = rowsToJsonArray(hourlyOut, function (h) {
        return '{"hour":' + (parseInt(h[0], 10) || 0)
            + ',"count":' + (parseInt(h[1], 10) || 0)
            + ',"total":' + jsonNum(h[2]) + '}';
    });
    var topJson = rowsToJsonArray(topOut, function (t) {
        return '{"nama":' + jsonStr(t[0])
            + ',"count":' + (parseInt(t[1], 10) || 0)
            + ',"total":' + jsonNum(t[2]) + '}';
    });

    var jsonOut = "{"
        + '"generatedAt":' + jsonStr(new Date().toString())
        + ',"reportDate":' + jsonStr(reportDateStr)
        + ',"store":{"nama":' + jsonStr(namaToko) + ',"alamat":' + jsonStr(alamatToko) + "}"
        + ',"summary":{'
        + '"totalTrx":' + totalTrx
        + ',"totalItem":' + totalItem
        + ',"subtotal":' + jsonNum(subtotal)
        + ',"totalDiskon":' + jsonNum(totalDiskon)
        + ',"totalPajak":' + jsonNum(totalPajak)
        + ',"totalAkhir":' + jsonNum(totalAkhir)
        + "}"
        + ',"paymentBreakdown":' + paymentJson
        + ',"hourly":' + hourlyJson
        + ',"topCustomers":' + topJson
        + ',"records":' + (rawRecordsJson || "[]")
        + "}";

    var jsonPath = SCRIPT_DIR + "\\report_" + reportDateStr + ".json";
    try {
        var jf = fso.CreateTextFile(jsonPath, true);
        jf.Write(jsonOut);
        jf.Close();
        log("JSON exported: " + jsonPath + " (" + jsonOut.length + " bytes)");
    } catch (e) {
        log("JSON write failed: " + e.message);
    }

    // Upload and capture URL. Disabled via UPLOAD_REPORT=0; retention via UPLOAD_MODE.
    var reportUrl = "";
    var viewerUrl = "";
    var VIEWER_BASE = env("VIEWER_URL") || "https://carinama1.github.io/report/report_viewer.html";
    var doUpload = (env("UPLOAD_REPORT") || "1") !== "0";
    if (doUpload) {
        var uploadMode = env("UPLOAD_MODE") || "permanent";
        try {
            reportUrl = uploadJsonCatbox(jsonOut, "report_" + reportDateStr + ".json", uploadMode);
            viewerUrl = VIEWER_BASE + "?url=" + reportUrl;
            log("Uploaded (" + uploadMode + "): " + reportUrl);
            log("Viewer : " + viewerUrl);
            try {
                var uf = fso.OpenTextFile(SCRIPT_DIR + "\\report_urls.txt", 8, true);
                uf.WriteLine("[" + new Date().toLocaleString() + "] " + reportDateStr
                    + " (" + uploadMode + ") -> " + viewerUrl);
                uf.Close();
            } catch (eW) { }
        } catch (e) {
            log("Upload failed: " + e.message);
        }
    }

    // ---- Build message ----
    var lines = [];

    lines.push("*LAPORAN PENJUALAN HARIAN*");
    if (namaToko) lines.push(namaToko);
    if (alamatToko) lines.push(alamatToko);
    lines.push("Tanggal: " + formatTanggal(REPORT_DATE_OBJ));
    lines.push("");

    lines.push("*Ringkasan*");
    lines.push("Total Transaksi : " + totalTrx + " transaksi");
    lines.push("Total Item      : " + totalItem + " item");
    lines.push("Subtotal        : " + formatRp(subtotal));
    if (totalDiskon > 0) lines.push("Diskon          : -" + formatRp(totalDiskon));
    if (totalPajak > 0) lines.push("Pajak           : " + formatRp(totalPajak));
    lines.push("*Total Penjualan : " + formatRp(totalAkhir) + "*");
    lines.push("");

    if (bayarOut) {
        lines.push("*Rincian Pembayaran*");
        var bayarRows = bayarOut.split("\n");
        for (var i = 0; i < bayarRows.length; i++) {
            if (!bayarRows[i]) continue;
            var p = bayarRows[i].split("|");
            var cara = p[0] || "Lainnya";
            while (cara.length < 12) cara += " ";
            lines.push(cara + ": " + formatRp(p[2]) + " (" + p[1] + "x)");
        }
        lines.push("");
    }

    if (hourlyOut) {
        lines.push("*Penjualan Per Jam*");
        var hourRows = hourlyOut.split("\n");
        for (var i = 0; i < hourRows.length; i++) {
            if (!hourRows[i]) continue;
            var h = hourRows[i].split("|");
            var jam = parseInt(h[0]);
            lines.push(pad2(jam) + ":00-" + pad2(jam + 1) + ":00 : " + formatRp(h[2]) + " (" + h[1] + "x)");
        }
        lines.push("");
    }

    if (topOut) {
        lines.push("*Top Pelanggan*");
        var topRows = topOut.split("\n");
        for (var i = 0; i < topRows.length; i++) {
            if (!topRows[i]) continue;
            var t = topRows[i].split("|");
            lines.push((i + 1) + ". " + t[0] + " - " + formatRp(t[2]) + " (" + t[1] + "x)");
        }
        lines.push("");
    }

    if (viewerUrl) {
        lines.push("*Lihat Detail Laporan:*");
        lines.push(viewerUrl);
        lines.push("");
    }

    lines.push("_Digenerate otomatis oleh iPOS Reporter_");

    var message = lines.join("\n");
    log("Message length: " + message.length + " chars.");

    if (DEBUG) {
        log("=== DEBUG MODE: WhatsApp send SKIPPED ===");
        log("Recipients (" + WHAPI_RECIPIENTS.length + "): " + WHAPI_RECIPIENTS.join(", "));
        log("--- Message preview ---");
        WScript.Echo(message);
        log("--- End of message ---");
    } else {
        if (WHAPI_RECIPIENTS.length === 0) {
            throw new Error("No recipients configured. Set WHAPI_TO_1 (and WHAPI_TO_COUNT) in config.bat");
        }
        for (var ri = 0; ri < WHAPI_RECIPIENTS.length; ri++) {
            var resp = sendWhatsapp(message, WHAPI_RECIPIENTS[ri]);
            log("Sent to " + WHAPI_RECIPIENTS[ri] + " [" + (ri + 1) + "/" + WHAPI_RECIPIENTS.length + "]. Response: " + resp.substring(0, 80));
        }

        // Cleanup: remove local JSON now that it's uploaded and sent.
        // Kept in DEBUG mode so developers can inspect the payload.
        try {
            if (fso.FileExists(jsonPath)) {
                fso.DeleteFile(jsonPath);
                log("Cleaned up: " + jsonPath);
            }
        } catch (eC) {
            log("Cleanup failed for " + jsonPath + ": " + eC.message);
        }
    }

    // ---- JSON export helpers (function declarations are hoisted within
    //      this try block, so they're usable from the upload block above) ----
    function jsonStr(s) { return '"' + jsonEscape(s == null ? "" : String(s)) + '"'; }
    function jsonNum(n) { var x = parseFloat(n); return isNaN(x) ? "0" : String(x); }

    // Minimal CSV parser (RFC 4180-ish) — handles quoted fields, "" escapes, CRLF.
    function parseCsv(text) {
        var rows = [], row = [], field = "", inQ = false, i = 0, n = text.length;
        while (i < n) {
            var c = text.charAt(i);
            if (inQ) {
                if (c === '"') {
                    if (text.charAt(i + 1) === '"') { field += '"'; i += 2; continue; }
                    inQ = false; i++; continue;
                }
                field += c; i++; continue;
            }
            if (c === '"') { inQ = true; i++; continue; }
            if (c === ',') { row.push(field); field = ""; i++; continue; }
            if (c === '\r') { i++; continue; }
            if (c === '\n') { row.push(field); rows.push(row); row = []; field = ""; i++; continue; }
            field += c; i++;
        }
        if (field !== "" || row.length > 0) { row.push(field); rows.push(row); }
        return rows;
    }

    // Value looks numeric? (int, decimal, negative)
    // Reject leading-zero integers like "004960" — those are IDs/codes, keep as strings.
    function looksNumeric(v) {
        if (v === "" || !/^-?\d+(\.\d+)?$/.test(v)) return false;
        if (/^-?0\d/.test(v)) return false;
        return true;
    }
    function rowToJsonObj(cols, vals) {
        var parts = [];
        for (var k = 0; k < cols.length; k++) {
            var v = vals[k];
            var out;
            if (v === undefined || v === "") out = "null";
            else if (looksNumeric(v)) out = v;
            else out = jsonStr(v);
            parts.push(jsonStr(cols[k]) + ":" + out);
        }
        return "{" + parts.join(",") + "}";
    }

    function rowsToJsonArray(out, mapFn) {
        if (!out) return "[]";
        var rows = out.split("\n");
        var parts = [];
        for (var i = 0; i < rows.length; i++) {
            if (!rows[i]) continue;
            parts.push(mapFn(rows[i].split("|")));
        }
        return "[" + parts.join(",") + "]";
    }

    // Build nested records: array of { header:{...}, items:[{...},...] }
    function buildRecordsJson(hdrCsv, dtlCsv) {
        var hdrRows = hdrCsv ? parseCsv(hdrCsv) : [];
        var dtlRows = dtlCsv ? parseCsv(dtlCsv) : [];
        if (hdrRows.length === 0) return "[]";

        var hdrCols = hdrRows[0];
        var dtlCols = dtlRows.length > 0 ? dtlRows[0] : [];

        // Find notransaksi column index in detail rows so we can group by it
        var dtlKeyIdx = -1;
        for (var c = 0; c < dtlCols.length; c++) {
            if (dtlCols[c] === "notransaksi") { dtlKeyIdx = c; break; }
        }

        // Group detail rows by notransaksi
        var dtlByKey = {};
        for (var r = 1; r < dtlRows.length; r++) {
            if (dtlKeyIdx < 0) break;
            var key = dtlRows[r][dtlKeyIdx];
            if (!dtlByKey[key]) dtlByKey[key] = [];
            dtlByKey[key].push(dtlRows[r]);
        }

        // Find notransaksi column in header
        var hdrKeyIdx = -1;
        for (var c2 = 0; c2 < hdrCols.length; c2++) {
            if (hdrCols[c2] === "notransaksi") { hdrKeyIdx = c2; break; }
        }

        var out = [];
        for (var r2 = 1; r2 < hdrRows.length; r2++) {
            var hRow = hdrRows[r2];
            var hKey = hdrKeyIdx >= 0 ? hRow[hdrKeyIdx] : "";
            var items = dtlByKey[hKey] || [];
            var itemParts = [];
            for (var ii = 0; ii < items.length; ii++) {
                itemParts.push(rowToJsonObj(dtlCols, items[ii]));
            }
            out.push('{"header":' + rowToJsonObj(hdrCols, hRow)
                + ',"items":[' + itemParts.join(",") + "]}");
        }
        return "[" + out.join(",") + "]";
    }

    log("=== DONE ===");
    WScript.Quit(0);

} catch (e) {
    log("ERROR: " + e.message);
    WScript.Quit(1);
}
