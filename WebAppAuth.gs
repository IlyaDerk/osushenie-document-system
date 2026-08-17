/** Простая сессия Web App: токен хранится только в Script Cache. */
const WEB_APP_SESSION_PREFIX_ = 'document-web-session:';
const WEB_APP_SESSION_SECONDS_ = 21600;

function webAppLogin(login, password) {
  const normalizedLogin = String(login == null ? '' : login).trim();
  const suppliedPassword = String(password == null ? '' : password);
  if (!normalizedLogin || !suppliedPassword) {
    return { ok: false, message: 'Введите логин и пароль.' };
  }
  try {
    assertSystemSheetsStructure_(['WEB_USERS']);
    const users = webAppReadUsers_().filter(function (user) {
      return user.login === normalizedLogin;
    });
    if (users.length > 1) {
      return {
        ok: false,
        message: 'Невозможно выполнить вход: в справочнике ' +
          'пользователей найден повторяющийся логин.'
      };
    }
    if (users.length !== 1 || users[0].password !== suppliedPassword) {
      return { ok: false, message: 'Неверный логин или пароль.' };
    }
    if (webAppNormalize_(users[0].access) !== webAppNormalize_('Да')) {
      return { ok: false, message: 'Доступ к Web-приложению закрыт.' };
    }
    const user = webAppPublicUser_(users[0]);
    const token = Utilities.getUuid() + Utilities.getUuid();
    CacheService.getScriptCache().put(
      WEB_APP_SESSION_PREFIX_ + token,
      JSON.stringify(user),
      WEB_APP_SESSION_SECONDS_
    );
    return { ok: true, sessionToken: token, user: user };
  } catch (error) {
    return { ok: false, message: webAppSafeError_(error) };
  }
}

function webAppResumeSession(sessionToken) {
  try {
    return { ok: true, user: webAppRequireSession_(sessionToken) };
  } catch (error) {
    return { ok: false, sessionExpired: true, message: error.message };
  }
}

function webAppLogout(sessionToken) {
  const token = String(sessionToken == null ? '' : sessionToken);
  if (token) CacheService.getScriptCache().remove(WEB_APP_SESSION_PREFIX_ + token);
  return { ok: true };
}

function webAppRequireSession_(sessionToken) {
  const token = String(sessionToken == null ? '' : sessionToken).trim();
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
  const context = getSystemSheetContext_('WEB_USERS');
  const rows = readCreationSheetValues_(context);
  const indexes = {
    login: creationColumnIndex_(context, H.WEB_LOGIN),
    password: creationColumnIndex_(context, H.WEB_PASSWORD),
    fullName: creationColumnIndex_(context, H.WEB_FULL_NAME),
    contact: creationColumnIndex_(context, H.WEB_CONTACT),
    access: creationColumnIndex_(context, H.WEB_ACCESS)
  };
  return rows.reduce(function (users, row) {
    const login = String(row[indexes.login] == null ? '' : row[indexes.login]).trim();
    if (!login) return users;
    users.push({
      login: login,
      password: String(row[indexes.password] == null ? '' : row[indexes.password]),
      fullName: String(row[indexes.fullName] == null ? '' : row[indexes.fullName]).trim(),
      contact: String(row[indexes.contact] == null ? '' : row[indexes.contact]).trim(),
      access: String(row[indexes.access] == null ? '' : row[indexes.access]).trim()
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

function webAppNormalize_(value) {
  return String(value == null ? '' : value).replace(/\u00a0/g, ' ').trim().toLowerCase();
}

function webAppSafeError_(error) {
  const message = String(error && error.message ? error.message : error);
  return message || 'Не удалось выполнить операцию.';
}
