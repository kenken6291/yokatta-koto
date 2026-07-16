/**
 * 良かったこと日記SNS - GAS バックエンド
 *
 * スプレッドシート構成:
 *  シート「Users」    : Email, PasswordHash, Salt, Nickname, MustChangePassword, AgreedTerms, CreatedAt, UpdatedAt
 *  シート「Entries」  : Id, Email, EntryDate, Item1, Item1Public, Item2, Item2Public, Item3, Item3Public, CreatedAt
 *
 * スクリプトプロパティ (PropertiesService) に設定必須:
 *  SPREADSHEET_ID : このアプリ専用のスプレッドシートID
 *  PEPPER         : パスワードハッシュ用の秘密文字列(自分で決めた長いランダム文字列)
 *
 * デプロイ: 「ウェブアプリとして新しいバージョンをデプロイ」を忘れずに。
 */

const SS_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const PEPPER = PropertiesService.getScriptProperties().getProperty('PEPPER');
const SESSION_TTL_SEC = 60 * 60 * 24 * 7; // 7日間

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ ok: false, error: 'リクエストの形式が不正です' });
  }

  const action = body.action;
  try {
    switch (action) {
      case 'register': return jsonOut(register(body));
      case 'login': return jsonOut(login(body));
      case 'logout': return jsonOut(logout(body));
      case 'changePassword': return jsonOut(changePassword(body));
      case 'updateProfile': return jsonOut(updateProfile(body));
      case 'createEntry': return jsonOut(createEntry(body));
      case 'getMyEntries': return jsonOut(getMyEntries(body));
      case 'getEntries': return jsonOut(getEntries(body));
      case 'updateEntrySharing': return jsonOut(updateEntrySharing(body));
      case 'deleteEntry': return jsonOut(deleteEntry(body));
      case 'me': return jsonOut(me(body));
      default: return jsonOut({ ok: false, error: '不明なactionです' });
    }
  } catch (err) {
    return jsonOut({ ok: false, error: 'サーバーエラー: ' + err.message });
  }
}

function jsonOut(obj) {
  // CORS対策: text/plain で返す(Ken標準パターン)
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.TEXT);
}

/* ---------------- シート取得 ---------------- */

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === 'Users') {
      sheet.appendRow(['Email', 'PasswordHash', 'Salt', 'Nickname', 'MustChangePassword', 'AgreedTerms', 'CreatedAt', 'UpdatedAt']);
    } else if (name === 'Entries') {
      sheet.appendRow(['Id', 'Email', 'EntryDate', 'Item1', 'Item1Public', 'Item2', 'Item2Public', 'Item3', 'Item3Public', 'CreatedAt']);
    }
  }
  return sheet;
}

function usersData() {
  const sheet = getSheet('Users');
  const values = sheet.getDataRange().getValues();
  values.shift(); // header
  return values;
}

function entriesData() {
  const sheet = getSheet('Entries');
  const values = sheet.getDataRange().getValues();
  values.shift();
  return values;
}

function findUserRow(email) {
  const sheet = getSheet('Users');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).toLowerCase() === String(email).toLowerCase()) {
      return { rowIndex: i + 1, row: values[i], sheet: sheet };
    }
  }
  return null;
}

/* ---------------- ハッシュ / パスワード ---------------- */

function randomToken(len) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789';
  let out = '';
  for (let i = 0; i < len; i++) {
    out += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return out;
}

function hashPassword(password, salt) {
  const raw = password + salt + PEPPER;
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

/* ---------------- セッション ---------------- */

function createSession(email) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('session_' + token, email, SESSION_TTL_SEC);
  return token;
}

function sessionEmail(token) {
  if (!token) return null;
  return CacheService.getScriptCache().get('session_' + token);
}

function requireAuth(body) {
  const email = sessionEmail(body.token);
  if (!email) throw new Error('ログインが必要です。再度ログインしてください。');
  const found = findUserRow(email);
  if (!found) throw new Error('ユーザーが見つかりません');
  return found;
}

/* ---------------- 会員登録 / ログイン ---------------- */

function register(body) {
  const email = String(body.email || '').trim().toLowerCase();
  const nickname = String(body.nickname || '').trim();
  const agreed = !!body.agreed;

  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { ok: false, error: 'メールアドレスの形式が正しくありません' };
  }
  if (!nickname) {
    return { ok: false, error: 'ニックネームを入力してください' };
  }
  if (!agreed) {
    return { ok: false, error: '注意事項・免責事項への同意が必要です' };
  }
  if (findUserRow(email)) {
    return { ok: false, error: 'このメールアドレスは既に登録されています' };
  }

  const tempPassword = randomToken(10);
  const salt = randomToken(16);
  const hash = hashPassword(tempPassword, salt);
  const now = new Date();

  getSheet('Users').appendRow([email, hash, salt, nickname, true, true, now, now]);

  try {
    MailApp.sendEmail({
      to: email,
      subject: '【今日の良かったこと】仮パスワードのお知らせ',
      body:
        nickname + ' 様\n\n' +
        '「今日の良かったこと」にご登録いただきありがとうございます。\n\n' +
        '仮パスワード: ' + tempPassword + '\n\n' +
        '初回ログイン時にパスワードの変更をお願いします。\n\n' +
        '※このメールに心当たりがない場合は破棄してください。'
    });
  } catch (err) {
    return { ok: false, error: 'メール送信に失敗しました: ' + err.message };
  }

  return { ok: true, message: '仮パスワードをメールに送信しました' };
}

function login(body) {
  const email = String(body.email || '').trim().toLowerCase();
  const password = String(body.password || '');
  const found = findUserRow(email);
  if (!found) return { ok: false, error: 'メールアドレスまたはパスワードが違います' };

  const row = found.row;
  const salt = row[2];
  const hash = row[1];
  if (hashPassword(password, salt) !== hash) {
    return { ok: false, error: 'メールアドレスまたはパスワードが違います' };
  }

  const token = createSession(email);
  return {
    ok: true,
    token: token,
    nickname: row[3],
    mustChangePassword: !!row[4]
  };
}

function logout(body) {
  if (body.token) CacheService.getScriptCache().remove('session_' + body.token);
  return { ok: true };
}

function me(body) {
  const found = requireAuth(body);
  return { ok: true, email: found.row[0], nickname: found.row[3], mustChangePassword: !!found.row[4] };
}

function changePassword(body) {
  const found = requireAuth(body);
  const newPassword = String(body.newPassword || '');
  if (newPassword.length < 6) {
    return { ok: false, error: 'パスワードは6文字以上にしてください' };
  }
  const salt = randomToken(16);
  const hash = hashPassword(newPassword, salt);
  const sheet = found.sheet;
  sheet.getRange(found.rowIndex, 2).setValue(hash);   // PasswordHash
  sheet.getRange(found.rowIndex, 3).setValue(salt);   // Salt
  sheet.getRange(found.rowIndex, 5).setValue(false);  // MustChangePassword
  sheet.getRange(found.rowIndex, 8).setValue(new Date());
  return { ok: true };
}

function updateProfile(body) {
  const found = requireAuth(body);
  const nickname = String(body.nickname || '').trim();
  if (!nickname) return { ok: false, error: 'ニックネームを入力してください' };
  found.sheet.getRange(found.rowIndex, 4).setValue(nickname);
  found.sheet.getRange(found.rowIndex, 8).setValue(new Date());
  return { ok: true, nickname: nickname };
}

/* ---------------- 日記エントリ ---------------- */

function createEntry(body) {
  const found = requireAuth(body);
  const email = found.row[0];
  const items = body.items || [];
  if (items.length !== 3) {
    return { ok: false, error: '3つの項目が必要です' };
  }
  for (const it of items) {
    if (!it.text || !String(it.text).trim()) {
      return { ok: false, error: '空欄の項目があります' };
    }
  }

  const id = Utilities.getUuid();
  const now = new Date();
  const entryDate = body.entryDate ? new Date(body.entryDate) : now;

  getSheet('Entries').appendRow([
    id, email, entryDate,
    String(items[0].text).trim(), !!items[0].isPublic,
    String(items[1].text).trim(), !!items[1].isPublic,
    String(items[2].text).trim(), !!items[2].isPublic,
    now
  ]);

  return { ok: true, id: id };
}

function rowToMyEntry(row) {
  return {
    id: row[0],
    entryDate: row[2],
    items: [
      { text: row[3], isPublic: !!row[4] },
      { text: row[5], isPublic: !!row[6] },
      { text: row[7], isPublic: !!row[8] }
    ],
    createdAt: row[9]
  };
}

function getMyEntries(body) {
  const found = requireAuth(body);
  const email = found.row[0];
  const rows = entriesData().filter(r => String(r[1]).toLowerCase() === email);
  rows.sort((a, b) => new Date(b[9]) - new Date(a[9]));
  return { ok: true, entries: rows.map(rowToMyEntry) };
}

function nicknameOf(email) {
  const found = findUserRow(email);
  return found ? found.row[3] : '(退会済み)';
}

// 他会員の公開項目 + 自分の全項目を、指定モードで返す
function getEntries(body) {
  const found = requireAuth(body);
  const myEmail = found.row[0];
  const mode = body.mode === 'random' ? 'random' : 'new';
  const limit = Math.min(Math.max(parseInt(body.limit || 20, 10), 1), 100);

  const rows = entriesData();
  const feed = [];

  rows.forEach(row => {
    const email = String(row[1]).toLowerCase();
    const isOwn = email === myEmail;
    const items = [
      { text: row[3], isPublic: !!row[4] },
      { text: row[5], isPublic: !!row[6] },
      { text: row[7], isPublic: !!row[8] }
    ];
    const visibleItems = isOwn ? items : items.filter(it => it.isPublic);
    if (visibleItems.length === 0) return; // 表示できる項目が無ければスキップ

    feed.push({
      id: row[0],
      isOwn: isOwn,
      nickname: isOwn ? found.row[3] : nicknameOf(email),
      entryDate: row[2],
      items: visibleItems,
      createdAt: row[9]
    });
  });

  if (mode === 'random') {
    for (let i = feed.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [feed[i], feed[j]] = [feed[j], feed[i]];
    }
  } else {
    feed.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  }

  return { ok: true, entries: feed.slice(0, limit) };
}

function updateEntrySharing(body) {
  const found = requireAuth(body);
  const email = found.row[0];
  const id = body.id;
  const sheet = getSheet('Entries');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      if (String(values[i][1]).toLowerCase() !== email) {
        return { ok: false, error: '権限がありません' };
      }
      sheet.getRange(i + 1, 5).setValue(!!body.item1Public);
      sheet.getRange(i + 1, 7).setValue(!!body.item2Public);
      sheet.getRange(i + 1, 9).setValue(!!body.item3Public);
      return { ok: true };
    }
  }
  return { ok: false, error: '記録が見つかりません' };
}

function deleteEntry(body) {
  const found = requireAuth(body);
  const email = found.row[0];
  const id = body.id;
  const sheet = getSheet('Entries');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      if (String(values[i][1]).toLowerCase() !== email) {
        return { ok: false, error: '権限がありません' };
      }
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: '記録が見つかりません' };
}
