/** Простая серверная авторизация Web App. */
const WEB_APP_SESSION_SECONDS_ = 21600;
const WEB_APP_SESSION_PREFIX_ = 'document-web-session:';

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

  const session = webAppPublicUser_(users[0]);
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put(
    WEB_APP_SESSION_PREFIX_ + token,
    JSON.stringify(session),
    WEB_APP_SESSION_SECONDS_
  );
  return { token: token, user: session };
}

function webAppLogout(sessionToken) {
  const token = String(sessionToken == null ? '' : sessionToken);
  if (token) CacheService.getScriptCache().remove(WEB_APP_SESSION_PREFIX_ + token);
  return { ok: true };
}

function webAppRequireSession_(sessionToken) {
  const token = String(sessionToken == null ? '' : sessionToken);
  const serialized = token && CacheService.getScriptCache().get(
    WEB_APP_SESSION_PREFIX_ + token
  );
  if (!serialized) throw new Error('Сессия истекла. Войдите снова.');
  const user = JSON.parse(serialized);
  CacheService.getScriptCache().put(
    WEB_APP_SESSION_PREFIX_ + token,
    serialized,
    WEB_APP_SESSION_SECONDS_
  );
  return user;
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
