/**
 * Budget app for kids and parents.
 *
 * Expected Google Sheets tabs:
 * 1) Users columns: username | password | role | childId | displayName
 *    - role values: kid, parent
 *    - For kid, childId should be unique (e.g. KID001)
 *    - For parent, childId can be empty
 * 2) Operations columns: timestamp | childId | amount | type | description | createdBy
 */

const SHEET_USERS = 'Users';
const SHEET_OPERATIONS = 'Operations';
const SESSION_TTL_SECONDS = 60 * 60 * 8; // 8 hours

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Kids Budget App')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(fileName) {
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

function login(payload) {
  validatePayload(payload, ['username', 'password']);
  const username = String(payload.username).trim();
  const password = String(payload.password);

  const users = getUsers_();
  const found = users.find((u) => u.username === username && u.password === password);

  if (!found) {
    return { success: false, message: 'Invalid username or password.' };
  }

  const session = createSession_(found);
  return {
    success: true,
    sessionToken: session.token,
    user: {
      username: found.username,
      role: found.role,
      displayName: found.displayName,
      childId: found.childId
    }
  };
}

function getDashboardData(sessionToken, selectedChildId) {
  const session = getSession_(sessionToken);
  if (!session) {
    throw new Error('Session expired. Please login again.');
  }

  let childId = session.childId;

  if (session.role === 'parent') {
    const target = String(selectedChildId || '').trim();
    if (!target) {
      return {
        user: session,
        kids: getKids_(),
        selectedChildId: '',
        balance: 0,
        operations: []
      };
    }
    childId = target;
  }

  const ops = getOperationsByChildId_(childId);
  const balance = ops.reduce((sum, op) => sum + op.amount, 0);

  return {
    user: session,
    kids: session.role === 'parent' ? getKids_() : [],
    selectedChildId: childId,
    balance: roundTo2_(balance),
    operations: ops
  };
}

function createOperation(sessionToken, payload) {
  const session = getSession_(sessionToken);
  if (!session) {
    throw new Error('Session expired. Please login again.');
  }
  if (session.role !== 'parent') {
    throw new Error('Only parents can create operations.');
  }

  validatePayload(payload, ['childId', 'amount', 'type', 'description']);

  const childId = String(payload.childId).trim();
  const amountInput = Number(payload.amount);
  const type = String(payload.type).trim().toLowerCase();
  const description = String(payload.description).trim();

  if (!childId) {
    throw new Error('Child ID is required.');
  }
  if (!Number.isFinite(amountInput) || amountInput <= 0) {
    throw new Error('Amount must be a positive number.');
  }
  if (type !== 'incoming' && type !== 'outgoing') {
    throw new Error('Operation type must be incoming or outgoing.');
  }
  if (!description) {
    throw new Error('Description is required.');
  }

  const signedAmount = type === 'incoming' ? amountInput : -Math.abs(amountInput);

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = getSheet_(SHEET_OPERATIONS);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['timestamp', 'childId', 'amount', 'type', 'description', 'createdBy']);
    }
    sheet.appendRow([
      new Date(),
      childId,
      roundTo2_(signedAmount),
      type,
      description,
      session.username
    ]);
  } finally {
    lock.releaseLock();
  }

  return { success: true };
}

function logout(sessionToken) {
  if (!sessionToken) {
    return { success: true };
  }
  const cache = CacheService.getScriptCache();
  cache.remove('session:' + sessionToken);
  return { success: true };
}

function validatePayload(payload, requiredFields) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Invalid request payload.');
  }
  requiredFields.forEach((field) => {
    if (payload[field] === undefined || payload[field] === null || String(payload[field]).trim() === '') {
      throw new Error('Missing required field: ' + field);
    }
  });
}

function createSession_(user) {
  const token = Utilities.getUuid();
  const session = {
    username: user.username,
    role: user.role,
    displayName: user.displayName,
    childId: user.childId || ''
  };
  CacheService.getScriptCache().put('session:' + token, JSON.stringify(session), SESSION_TTL_SECONDS);
  return { token: token, session: session };
}

function getSession_(token) {
  const t = String(token || '').trim();
  if (!t) {
    return null;
  }
  const raw = CacheService.getScriptCache().get('session:' + t);
  if (!raw) {
    return null;
  }
  return JSON.parse(raw);
}

function getUsers_() {
  const sheet = getSheet_(SHEET_USERS);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return [];
  }

  return values.slice(1).map((row) => ({
    username: String(row[0] || '').trim(),
    password: String(row[1] || ''),
    role: String(row[2] || '').trim().toLowerCase(),
    childId: String(row[3] || '').trim(),
    displayName: String(row[4] || '').trim() || String(row[0] || '').trim()
  })).filter((u) => u.username && u.password && (u.role === 'kid' || u.role === 'parent'));
}

function getKids_() {
  return getUsers_()
    .filter((u) => u.role === 'kid' && u.childId)
    .map((u) => ({ childId: u.childId, displayName: u.displayName }));
}

function getOperationsByChildId_(childId) {
  const sheet = getSheet_(SHEET_OPERATIONS);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return [];
  }

  return values
    .slice(1)
    .filter((row) => String(row[1] || '').trim() === childId)
    .map((row) => ({
      timestamp: row[0] instanceof Date ? row[0].toISOString() : String(row[0] || ''),
      childId: String(row[1] || '').trim(),
      amount: Number(row[2] || 0),
      type: String(row[3] || '').trim().toLowerCase(),
      description: String(row[4] || '').trim(),
      createdBy: String(row[5] || '').trim()
    }))
    .sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp)));
}

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error('Missing sheet: ' + name);
  }
  return sheet;
}

function roundTo2_(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}
