/**
 * 良かったこと日記SNS - GAS バックエンド
 *
 * スプレッドシート構成:
 *  シート「Users」    : Email, PasswordHash, Salt, Nickname, MustChangePassword, AgreedTerms, CreatedAt, UpdatedAt
 *  シート「Entries」  : Id, Email, EntryDate, Item1, Item1Public, Item2, Item2Public, Item3, Item3Public, CreatedAt, UpdatedAt
 *  シート「Comments」 : Id, EntryId, ParentId, Email, Text, CreatedAt, UpdatedAt  (ParentId空欄=記録への直接コメント/あり=返信)
 *  シート「Likes」    : Id, TargetType(entry|comment), TargetId, Email, CreatedAt
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

/**
 * ★最初に1回だけ、このプロジェクトを開いて手動で実行する関数★
 * (Apps Scriptエディタ上部の関数選択で setup を選び、▷実行 をクリック)
 *
 * これを実行すると:
 *  1. スプレッドシート/メール送信の権限許可画面が表示される(許可する)
 *  2. Users, Entries, Comments, Likes シートが見出し付きで作成される
 *  3. スクリプトプロパティ(SPREADSHEET_ID / PEPPER)が設定されているか確認する
 * 実行後、実行ログに「OK」と出れば準備完了。その後「デプロイ」に進んでください。
 */
function setup() {
  if (!SS_ID) {
    throw new Error('スクリプトプロパティ SPREADSHEET_ID が未設定です。歯車アイコン→スクリプトプロパティで設定してください。');
  }
  if (!PEPPER) {
    throw new Error('スクリプトプロパティ PEPPER が未設定です。歯車アイコン→スクリプトプロパティで設定してください。');
  }
  getSheet('Users');
  getSheet('Entries');
  getSheet('Comments');
  getSheet('Likes');
  Logger.log('OK: シート作成・権限確認が完了しました。ウェブアプリとしてデプロイしてください。');
}

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
      case 'requestPasswordReset': return jsonOut(requestPasswordReset(body));
      case 'logout': return jsonOut(logout(body));
      case 'changePassword': return jsonOut(changePassword(body));
      case 'updateProfile': return jsonOut(updateProfile(body));
      case 'createEntry': return jsonOut(createEntry(body));
      case 'updateEntry': return jsonOut(updateEntry(body));
      case 'getMyEntries': return jsonOut(getMyEntries(body));
      case 'getEntries': return jsonOut(getEntries(body));
      case 'deleteEntry': return jsonOut(deleteEntry(body));
      case 'addComment': return jsonOut(addComment(body));
      case 'updateComment': return jsonOut(updateComment(body));
      case 'deleteComment': return jsonOut(deleteComment(body));
      case 'getComments': return jsonOut(getComments(body));
      case 'toggleLike': return jsonOut(toggleLike(body));
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
      sheet.appendRow(['Id', 'Email', 'EntryDate', 'Item1', 'Item1Public', 'Item2', 'Item2Public', 'Item3', 'Item3Public', 'CreatedAt', 'UpdatedAt']);
    } else if (name === 'Comments') {
      sheet.appendRow(['Id', 'EntryId', 'ParentId', 'Email', 'Text', 'CreatedAt', 'UpdatedAt']);
    } else if (name === 'Likes') {
      sheet.appendRow(['Id', 'TargetType', 'TargetId', 'Email', 'CreatedAt']);
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

function findEntryRow(id) {
  const sheet = getSheet('Entries');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      return { rowIndex: i + 1, row: values[i], sheet: sheet };
    }
  }
  return null;
}

function findCommentRow(id) {
  const sheet = getSheet('Comments');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
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

function requestPasswordReset(body) {
  const email = String(body.email || '').trim().toLowerCase();
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { ok: false, error: 'メールアドレスの形式が正しくありません' };
  }

  const found = findUserRow(email);
  // メールアドレスの存在有無を外部に漏らさないよう、見つからない場合も同じ成功メッセージを返す
  if (!found) {
    return { ok: true, message: 'このメールアドレスが登録されていれば、仮パスワードを送信しました' };
  }

  const tempPassword = randomToken(10);
  const salt = randomToken(16);
  const hash = hashPassword(tempPassword, salt);

  found.sheet.getRange(found.rowIndex, 2).setValue(hash);   // PasswordHash
  found.sheet.getRange(found.rowIndex, 3).setValue(salt);   // Salt
  found.sheet.getRange(found.rowIndex, 5).setValue(true);   // MustChangePassword
  found.sheet.getRange(found.rowIndex, 8).setValue(new Date());

  const nickname = found.row[3];

  try {
    MailApp.sendEmail({
      to: email,
      subject: '【今日の良かったこと】仮パスワード再発行のお知らせ',
      body:
        nickname + ' 様\n\n' +
        'パスワード再発行のリクエストを受け付けました。\n\n' +
        '仮パスワード: ' + tempPassword + '\n\n' +
        'ログイン後、新しいパスワードの設定をお願いします。\n\n' +
        '※このリクエストに心当たりがない場合は、このメールを破棄してください。' +
        'パスワードは変更されていません。'
    });
  } catch (err) {
    return { ok: false, error: 'メール送信に失敗しました: ' + err.message };
  }

  return { ok: true, message: 'このメールアドレスが登録されていれば、仮パスワードを送信しました' };
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
    now, now
  ]);

  return { ok: true, id: id };
}

function updateEntry(body) {
  const found = requireAuth(body);
  const email = found.row[0];
  const entryFound = findEntryRow(body.id);
  if (!entryFound) return { ok: false, error: '記録が見つかりません' };
  if (String(entryFound.row[1]).toLowerCase() !== email) {
    return { ok: false, error: '権限がありません' };
  }
  const items = body.items || [];
  if (items.length !== 3) {
    return { ok: false, error: '3つの項目が必要です' };
  }
  for (const it of items) {
    if (!it.text || !String(it.text).trim()) {
      return { ok: false, error: '空欄の項目があります' };
    }
  }
  const sheet = entryFound.sheet;
  const r = entryFound.rowIndex;
  sheet.getRange(r, 4).setValue(String(items[0].text).trim());
  sheet.getRange(r, 5).setValue(!!items[0].isPublic);
  sheet.getRange(r, 6).setValue(String(items[1].text).trim());
  sheet.getRange(r, 7).setValue(!!items[1].isPublic);
  sheet.getRange(r, 8).setValue(String(items[2].text).trim());
  sheet.getRange(r, 9).setValue(!!items[2].isPublic);
  sheet.getRange(r, 11).setValue(new Date()); // UpdatedAt
  return { ok: true };
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
    createdAt: row[9],
    updatedAt: row[10]
  };
}

function nicknameOf(email) {
  const found = findUserRow(email);
  return found ? found.row[3] : '(退会済み)';
}

/* ---------------- いいね ---------------- */

function getLikeInfo(targetType, targetIds, myEmail) {
  const map = {};
  targetIds.forEach(id => { map[id] = { count: 0, likedByMe: false }; });
  const rows = getSheet('Likes').getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r[1] === targetType && map.hasOwnProperty(r[2])) {
      map[r[2]].count++;
      if (String(r[3]).toLowerCase() === myEmail) map[r[2]].likedByMe = true;
    }
  }
  return map;
}

function removeLikesForTargets(targetType, targetIds) {
  if (!targetIds.length) return;
  const sheet = getSheet('Likes');
  const values = sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i][1] === targetType && targetIds.indexOf(values[i][2]) !== -1) {
      sheet.deleteRow(i + 1);
    }
  }
}

function entryVisibleToViewer(entryRow, viewerEmail) {
  const email = String(entryRow[1]).toLowerCase();
  if (email === viewerEmail) return true;
  return !!entryRow[4] || !!entryRow[6] || !!entryRow[8];
}

function toggleLike(body) {
  const found = requireAuth(body);
  const email = found.row[0];
  const targetType = body.targetType;
  const targetId = body.targetId;
  if (targetType !== 'entry' && targetType !== 'comment') {
    return { ok: false, error: '不正な対象です' };
  }

  if (targetType === 'entry') {
    const entryFound = findEntryRow(targetId);
    if (!entryFound) return { ok: false, error: '記録が見つかりません' };
    if (!entryVisibleToViewer(entryFound.row, email)) return { ok: false, error: 'この記録にはいいねできません' };
  } else {
    const commentFound = findCommentRow(targetId);
    if (!commentFound) return { ok: false, error: 'コメントが見つかりません' };
    const entryFound = findEntryRow(commentFound.row[1]);
    if (!entryFound || !entryVisibleToViewer(entryFound.row, email)) {
      return { ok: false, error: 'いいねできません' };
    }
  }

  const sheet = getSheet('Likes');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][1] === targetType && values[i][2] === targetId && String(values[i][3]).toLowerCase() === email) {
      sheet.deleteRow(i + 1);
      return { ok: true, liked: false };
    }
  }
  sheet.appendRow([Utilities.getUuid(), targetType, targetId, email, new Date()]);
  return { ok: true, liked: true };
}

/* ---------------- 記録の一覧取得 ---------------- */

function getCommentCounts(entryIds) {
  const map = {};
  entryIds.forEach(id => { map[id] = 0; });
  const rows = getSheet('Comments').getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const eid = String(rows[i][1]);
    if (map.hasOwnProperty(eid)) map[eid]++;
  }
  return map;
}

function getMyEntries(body) {
  const found = requireAuth(body);
  const email = found.row[0];
  const rows = entriesData().filter(r => String(r[1]).toLowerCase() === email);
  rows.sort((a, b) => new Date(b[9]) - new Date(a[9]));
  const entries = rows.map(rowToMyEntry);
  const ids = entries.map(e => e.id);
  const likeMap = getLikeInfo('entry', ids, email);
  const commentCounts = getCommentCounts(ids);
  entries.forEach(e => {
    const li = likeMap[e.id] || { count: 0, likedByMe: false };
    e.likeCount = li.count;
    e.likedByMe = li.likedByMe;
    e.commentCount = commentCounts[e.id] || 0;
  });
  return { ok: true, entries: entries };
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
      createdAt: row[9],
      updatedAt: row[10]
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

  const page = feed.slice(0, limit);
  const ids = page.map(e => e.id);
  const likeMap = getLikeInfo('entry', ids, myEmail);
  const commentCounts = getCommentCounts(ids);
  page.forEach(e => {
    const li = likeMap[e.id] || { count: 0, likedByMe: false };
    e.likeCount = li.count;
    e.likedByMe = li.likedByMe;
    e.commentCount = commentCounts[e.id] || 0;
  });

  return { ok: true, entries: page };
}

function cascadeDeleteEntryRelated(entryId) {
  const cSheet = getSheet('Comments');
  const cValues = cSheet.getDataRange().getValues();
  const commentIds = [];
  for (let i = cValues.length - 1; i >= 1; i--) {
    if (String(cValues[i][1]) === entryId) {
      commentIds.push(cValues[i][0]);
      cSheet.deleteRow(i + 1);
    }
  }
  removeLikesForTargets('entry', [entryId]);
  if (commentIds.length) removeLikesForTargets('comment', commentIds);
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
      cascadeDeleteEntryRelated(id);
      return { ok: true };
    }
  }
  return { ok: false, error: '記録が見つかりません' };
}

/* ---------------- コメント / 返信 ---------------- */

function addComment(body) {
  const found = requireAuth(body);
  const email = found.row[0];
  const entryId = body.entryId;
  const text = String(body.text || '').trim();
  const parentId = body.parentId ? String(body.parentId) : '';

  if (!text) return { ok: false, error: 'コメントを入力してください' };
  if (text.length > 500) return { ok: false, error: 'コメントは500文字以内にしてください' };

  const entryFound = findEntryRow(entryId);
  if (!entryFound) return { ok: false, error: '記録が見つかりません' };
  if (!entryVisibleToViewer(entryFound.row, email)) return { ok: false, error: 'この記録にはコメントできません' };

  if (parentId) {
    const parentFound = findCommentRow(parentId);
    if (!parentFound) return { ok: false, error: '返信先のコメントが見つかりません' };
  }

  const id = Utilities.getUuid();
  const now = new Date();
  getSheet('Comments').appendRow([id, entryId, parentId, email, text, now, now]);
  return { ok: true, id: id };
}

function updateComment(body) {
  const found = requireAuth(body);
  const email = found.row[0];
  const commentFound = findCommentRow(body.id);
  if (!commentFound) return { ok: false, error: 'コメントが見つかりません' };
  if (String(commentFound.row[3]).toLowerCase() !== email) return { ok: false, error: '権限がありません' };

  const text = String(body.text || '').trim();
  if (!text) return { ok: false, error: 'コメントを入力してください' };
  if (text.length > 500) return { ok: false, error: 'コメントは500文字以内にしてください' };

  commentFound.sheet.getRange(commentFound.rowIndex, 5).setValue(text);
  commentFound.sheet.getRange(commentFound.rowIndex, 7).setValue(new Date());
  return { ok: true };
}

function deleteComment(body) {
  const found = requireAuth(body);
  const email = found.row[0];
  const commentFound = findCommentRow(body.id);
  if (!commentFound) return { ok: false, error: 'コメントが見つかりません' };
  if (String(commentFound.row[3]).toLowerCase() !== email) return { ok: false, error: '権限がありません' };

  const sheet = getSheet('Comments');
  const values = sheet.getDataRange().getValues();
  const idsToDelete = [body.id];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][2]) === body.id) idsToDelete.push(values[i][0]);
  }
  for (let i = values.length - 1; i >= 1; i--) {
    if (idsToDelete.indexOf(values[i][0]) !== -1) {
      sheet.deleteRow(i + 1);
    }
  }
  removeLikesForTargets('comment', idsToDelete);
  return { ok: true };
}

function getComments(body) {
  const found = requireAuth(body);
  const myEmail = found.row[0];
  const entryId = body.entryId;

  const entryFound = findEntryRow(entryId);
  if (!entryFound) return { ok: false, error: '記録が見つかりません' };
  if (!entryVisibleToViewer(entryFound.row, myEmail)) return { ok: false, error: 'この記録は閲覧できません' };

  const rows = getSheet('Comments').getDataRange().getValues();
  rows.shift();
  const all = rows.filter(r => String(r[1]) === entryId).map(r => {
    const email = String(r[3]).toLowerCase();
    return {
      id: r[0], entryId: r[1], parentId: r[2] || null,
      email: email,
      nickname: nicknameOf(email),
      text: r[4], createdAt: r[5], updatedAt: r[6],
      isOwn: email === myEmail
    };
  });

  const likeMap = getLikeInfo('comment', all.map(c => c.id), myEmail);
  all.forEach(c => {
    const li = likeMap[c.id] || { count: 0, likedByMe: false };
    c.likeCount = li.count;
    c.likedByMe = li.likedByMe;
  });

  const top = all.filter(c => !c.parentId).sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));
  const repliesMap = {};
  all.filter(c => c.parentId).forEach(c => {
    if (!repliesMap[c.parentId]) repliesMap[c.parentId] = [];
    repliesMap[c.parentId].push(c);
  });
  top.forEach(c => {
    c.replies = (repliesMap[c.id] || []).sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));
  });

  return { ok: true, comments: top };
}
