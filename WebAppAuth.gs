/** Простая серверная авторизация Web App. */
const WEB_APP_SESSION_PREFIX_ = 'document-web-session:';
const WEB_APP_INVALID_SESSION_MESSAGE_ = 'Сессия недействительна. Войдите снова.';
const WEB_APP_ACCESS_REVOKED_MESSAGE_ = 'Доступ к приложению прекращён. Войдите снова.';

function webAppLogin(login, password) {
  const normalizedLogin = String(login == null ? '' : login).trim();
  const suppliedPassword = String(password == null ? '' : password);
  if (!normalizedLogin || !suppliedPassword) {
    throw new Error('Введите логин и пароль.');
  }

  const users = webAppReadUsers_().filter(function (user) {
    return user.login === normalizedLogin;
  });
  if (users.length > 1) {
    throw new Error(
      'Невозможно выполнить вход: в справочнике пользователей ' +
      'найден повторяющийся логин.'
    );
  }
  if (users.length !== 1 || users[0].password !== suppliedPassword) {
    throw new Error('Неверный логин или пароль.');
  }
  if (users[0].access !== SYSTEM_CONFIG.VALUES.WEB_ACCESS_ALLOWED) {
    throw new Error('Доступ к приложению запрещён.');
  }

  const publicUser = webAppPublicUser_(users[0]);
  const session = { login: users[0].login };
  const token = Utilities.getUuid() + Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty(
    WEB_APP_SESSION_PREFIX_ + token,
    JSON.stringify(session)
  );
  return { token: token, user: publicUser };
}

function webAppLogout(sessionToken) {
  const token = String(sessionToken == null ? '' : sessionToken);
  if (token) {
    PropertiesService.getScriptProperties().deleteProperty(
      WEB_APP_SESSION_PREFIX_ + token
    );
  }
  return { ok: true };
}

function webAppRequireSession_(sessionToken) {
  const token = String(sessionToken == null ? '' : sessionToken);
  const properties = PropertiesService.getScriptProperties();
  const key = WEB_APP_SESSION_PREFIX_ + token;
  const serialized = token && properties.getProperty(
    key
  );
  if (!serialized) throw new Error(WEB_APP_INVALID_SESSION_MESSAGE_);

  let session;
  try {
    session = JSON.parse(serialized);
  } catch (error) {
    properties.deleteProperty(key);
    throw new Error(WEB_APP_INVALID_SESSION_MESSAGE_);
  }
  const login = String(session && session.login != null
    ? session.login : '').trim();
  if (!login) {
    properties.deleteProperty(key);
    throw new Error(WEB_APP_INVALID_SESSION_MESSAGE_);
  }

  const matches = webAppReadUsers_().filter(function (user) {
    return user.login === login;
  });
  if (
    matches.length !== 1 ||
    matches[0].access !== SYSTEM_CONFIG.VALUES.WEB_ACCESS_ALLOWED
  ) {
    properties.deleteProperty(key);
    throw new Error(WEB_APP_ACCESS_REVOKED_MESSAGE_);
  }
  return webAppPublicUser_(matches[0]);
}

function webAppReadUsers_() {
  assertSystemSheetsStructure_(['WEB_USERS']);
  const context = getSystemSheetContext_('WEB_USERS');
  const rows = webAppReadRows_(context);
  const indexes = webAppIndexes_(context, context.config.requiredHeaders);
  return rows.reduce(function (users, row) {
    const login = String(row[indexes[H.WEB_LOGIN]] == null
      ? '' : row[indexes[H.WEB_LOGIN]]).trim();
    if (!login) return users;
    users.push({
      login: login,
      password: String(row[indexes[H.WEB_PASSWORD]] == null
        ? '' : row[indexes[H.WEB_PASSWORD]]),
      fullName: String(row[indexes[H.WEB_FULL_NAME]] == null
        ? '' : row[indexes[H.WEB_FULL_NAME]]).trim(),
      contact: String(row[indexes[H.WEB_CONTACT]] == null
        ? '' : row[indexes[H.WEB_CONTACT]]).trim(),
      access: String(row[indexes[H.WEB_ACCESS]] == null
        ? '' : row[indexes[H.WEB_ACCESS]]).trim()
    });
    return users;
  }, []);
}

function webAppPublicUser_(user) {
  return {
    login: user.login,
    fullName: user.fullName,
    contact: user.contact,
    actor: user.contact || user.login
  };
}
