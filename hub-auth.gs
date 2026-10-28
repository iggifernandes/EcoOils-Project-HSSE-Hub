/**
 * EcoOils HSSE Hub - Team PIN service (free, runs in your own Google account).
 * - No PIN is stored in this code. The PIN is set only through the "Forgot PIN / set PIN" email flow.
 * - A 6-digit code is emailed to the account that owns this script.
 * - The PIN is kept only as a salted hash in Script Properties.
 * - 5 wrong PINs lock logins for 15 minutes.
 * - Also stores which tools ICs can see (the 'Share with ICs' switches in the hub).
 * Deploy: Extensions > Apps Script > paste this > Deploy > New deployment > Web app
 *         Execute as: Me   |   Who has access: Anyone   > copy the Web app URL into AUTH_API in the hub.
 */
function doPost(e) {
  var out;
  try { out = handle_(JSON.parse(e.postData.contents)); } catch (err) { out = { ok: false, error: 'bad request' }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() { return ContentService.createTextOutput('EcoOils hub auth OK'); }

function hash_(s, salt) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + s, Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 0xff).toString(16)).slice(-2); }).join('');
}
function handle_(r) {
  var P = PropertiesService.getScriptProperties(), C = CacheService.getScriptCache();
  if (r.a === 'login') {
    if (!P.getProperty('pinHash')) return { ok: false, notset: true };
    if (Number(C.get('fails') || 0) >= 5) return { ok: false, locked: true };
    if (hash_(String(r.pin || ''), P.getProperty('salt')) === P.getProperty('pinHash')) {
      C.remove('fails');
      var tok = Utilities.getUuid(); C.put('tok_' + tok, '1', 21600);   // session token, valid 6 hours
      return { ok: true, tok: tok };
    }
    C.put('fails', String(Number(C.get('fails') || 0) + 1), 900);
    return { ok: false };
  }
  if (r.a === 'shared') {                                 // public: which tools ICs can see
    var k = P.getProperty('shared');
    return k ? { ok: true, keys: JSON.parse(k) } : { ok: true, keys: null };
  }
  if (r.a === 'setshared') {                               // team only: needs a valid login token
    if (!r.tok || !C.get('tok_' + r.tok) || !Array.isArray(r.keys)) return { ok: false };
    var clean = r.keys.filter(function (x) { return /^[a-z0-9_]{1,30}$/.test(x); }).slice(0, 100);
    P.setProperty('shared', JSON.stringify(clean));
    return { ok: true };
  }
  if (r.a === 'forgot') {
    if (C.get('codeSent')) return { ok: true };            // one email per minute
    var code = ('' + (100000 + Math.floor(Math.random() * 900000)));
    P.setProperty('codeHash', hash_(code, 'code'));
    P.setProperty('codeExp', String(Date.now() + 15 * 60 * 1000));
    C.put('codeSent', '1', 60);
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'EcoOils Hub PIN code',
      'Your hub PIN reset code is ' + code + '. It expires in 15 minutes. If you did not ask for it, ignore this email.');
    return { ok: true };
  }
  if (r.a === 'reset') {
    var okCode = P.getProperty('codeHash') && Date.now() < Number(P.getProperty('codeExp') || 0) &&
                 hash_(String(r.code || ''), 'code') === P.getProperty('codeHash');
    if (!okCode || !/^\d{4,8}$/.test(String(r.pin || ''))) return { ok: false };
    var salt = Utilities.getUuid();
    P.setProperty('salt', salt); P.setProperty('pinHash', hash_(String(r.pin), salt));
    P.deleteProperty('codeHash'); P.deleteProperty('codeExp'); C.remove('fails');
    return { ok: true };
  }
  return { ok: false };
}
