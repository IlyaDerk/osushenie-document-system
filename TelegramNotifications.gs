/** Stage 3 Telegram formatting, configuration, delivery planning and transport. */
const WORKFLOW_TELEGRAM_ = Object.freeze({
  MAX_MESSAGE_LENGTH: 3900,
  TEST_MARKER: '🧪 ТЕСТ — уведомления по документообороту',
  PROPERTIES: Object.freeze({
    ENABLED: 'WORKFLOW_NOTIFICATIONS_ENABLED',
    CHAT_ID: 'WORKFLOW_TG_CHAT_ID',
    BOT_TOKEN: 'WORKFLOW_TG_BOT_TOKEN'
  })
});

function readWorkflowNotificationConfig_() {
  const properties = PropertiesService.getScriptProperties();
  const enabledText = String(properties.getProperty(WORKFLOW_TELEGRAM_.PROPERTIES.ENABLED) || '').trim().toLowerCase();
  return {
    enabled: enabledText === 'true',
    enabledConfigured: enabledText !== '',
    chatId: String(properties.getProperty(WORKFLOW_TELEGRAM_.PROPERTIES.CHAT_ID) || '').trim(),
    botToken: String(properties.getProperty(WORKFLOW_TELEGRAM_.PROPERTIES.BOT_TOKEN) || '').trim()
  };
}

function requireWorkflowNotificationConfig_(config, requireToken) {
  if (!config.chatId) throw new Error('Не настроен Script Property WORKFLOW_TG_CHAT_ID.');
  if (requireToken && !config.botToken) {
    throw new Error('Не настроен Script Property WORKFLOW_TG_BOT_TOKEN.');
  }
  return config;
}

function redactWorkflowNotificationSecret_(value, secret) {
  let text = String(value == null ? '' : value);
  const token = String(secret || '');
  if (token) text = text.split(token).join('[REDACTED]');
  return text.replace(/https:\/\/api\.telegram\.org\/bot[^/\s]+/gi,
    'https://api.telegram.org/bot[REDACTED]');
}

function telegramDisplayDate_(dateKey) {
  const match = String(dateKey || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return String(dateKey || '').trim() || 'не указано';
  return match[3] + '.' + match[2] + '.' + match[1];
}

function telegramCandidateDate_(value) {
  if (value && typeof value.getTime === 'function' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, WORKFLOW_NOTIFICATION_.TIMEZONE, 'dd.MM.yyyy');
  }
  return telegramDisplayDate_(value);
}

function telegramRussianDays_(number) {
  const n = Math.abs(Number(number)), tail = n % 100, last = n % 10;
  if (tail >= 11 && tail <= 14) return n + ' дней';
  if (last === 1) return n + ' день';
  if (last >= 2 && last <= 4) return n + ' дня';
  return n + ' дней';
}

function telegramDeadlineText_(daysRemaining) {
  const days = Number(daysRemaining);
  if (days === 0) return 'Срок: сегодня';
  if (days < 0) return 'Просрочено: ' + telegramRussianDays_(-days);
  return 'Срок: через ' + telegramRussianDays_(days);
}

function formatTelegramDocumentBlock_(candidate, index) {
  const type = String(candidate.documentType || '').trim() || 'Документ';
  const number = String(candidate.documentNumber || '').trim();
  const label = [type, number].filter(function (value) { return value; }).join(' ');
  const action = String(candidate.action || '').trim() ||
    WORKFLOW_NOTIFICATION_.ACTION_NOT_CONFIGURED;
  return String(index) + '. ' + label + ' - ' + action + '. До ' +
    telegramDisplayDate_(candidate.controlDateKey);
}

function telegramObjectHeading_(object) {
  const name = String(object.objectName || '').trim();
  const id = String(object.objectId || '').trim();
  return '🏗 Объект: ' + (name || (id ? '№ ' + id : 'без названия'));
}

function telegramHeader_(businessDate, testMode, partNumber, partCount) {
  const lines = [];
  if (testMode) lines.push(WORKFLOW_TELEGRAM_.TEST_MARKER);
  if (partCount > 1) lines.push('Часть ' + partNumber + '/' + partCount);
  return lines.join('\n');
}

/** Pure deterministic formatter/splitter. */
function buildTelegramPhysicalMessages_(model, options) {
  const input = options || {}, testMode = input.testMode === true;
  const footer = input.operatorCardUrl ? 'Ссылка на таблицу: ' + input.operatorCardUrl : '';
  const objects = (model && model.objects) || [];
  if (!objects.length && !testMode) return [];
  const headerReserve = telegramHeader_(input.businessDate, testMode, 9999, 9999).length + 2;
  const footerReserve = footer ? footer.length + 2 : 0;
  const bodyLimit = WORKFLOW_TELEGRAM_.MAX_MESSAGE_LENGTH - headerReserve - footerReserve;
  if (bodyLimit < 1) throw new Error('Ссылка карточки операциониста слишком длинная.');
  let segments = [];
  objects.forEach(function (object) {
    const heading = telegramObjectHeading_(object);
    const documents = (object.documents || []).map(function (candidate, index) {
      return formatTelegramDocumentBlock_(candidate, index + 1);
    });
    const whole = [heading].concat(documents).join('\n');
    if (whole.length <= bodyLimit) {
      segments.push({ text: whole, references: (object.documents || []).slice() });
      return;
    }
    documents.forEach(function (documentText, index) {
      const block = heading + '\n' + documentText;
      const candidate = object.documents[index];
      if (block.length > bodyLimit) {
        throw new Error('Telegram document block exceeds safe limit: documentId=' +
          String(candidate.documentId || '').trim() + ', objectId=' +
          String(candidate.objectId || '').trim() + ', length=' + block.length +
          ', limit=' + bodyLimit + '.');
      }
      segments.push({ text: block, references: [candidate] });
    });
  });
  if (!segments.length) segments = [{ text: 'На текущую дату уведомлений нет.', references: [] }];
  const bodies = [];
  segments.forEach(function (segment) {
    const current = bodies[bodies.length - 1], combined = current ? current.text + '\n\n' + segment.text : segment.text;
    if (!current || combined.length > bodyLimit) {
      bodies.push({ text: segment.text, references: segment.references.slice() });
    } else {
      current.text = combined;
      current.references = current.references.concat(segment.references);
    }
  });
  const count = bodies.length;
  return bodies.map(function (body, index) {
    const header = telegramHeader_(input.businessDate, testMode, index + 1, count);
    let text = header ? header + '\n\n' + body.text : body.text;
    if (index === count - 1 && footer) text += '\n\n' + footer;
    if (text.length > WORKFLOW_TELEGRAM_.MAX_MESSAGE_LENGTH) throw new Error('Telegram message exceeds safe limit.');
    return { partNumber: index + 1, partCount: count, text: text, references: body.references };
  });
}

function hashTelegramMessage_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    String(text), Utilities.Charset.UTF_8);
  return bytes.map(function (byte) { return ('0' + ((byte + 256) % 256).toString(16)).slice(-2); }).join('');
}

function buildTelegramDeliveryPlan_(model, options) {
  const input = options || {};
  return buildTelegramPhysicalMessages_(model, input).map(function (part) {
    return {
      deliveryId: generateNotificationId_('NOTIFICATION_DELIVERY'), attemptNumber: 1,
      channel: WORKFLOW_NOTIFICATION_HISTORY_.CHANNEL_TELEGRAM,
      target: String(input.target || '').trim(), messageText: part.text,
      messageHash: hashTelegramMessage_(part.text), partNumber: part.partNumber,
      partCount: part.partCount, references: part.references
    };
  });
}

function telegramFailure_(status, code, text, token) {
  let result = 'UNKNOWN_DELIVERY_OUTCOME', retry = false;
  const numeric = Number(code);
  if (numeric === 429 || numeric >= 500) { result = 'FAILED_TECHNICAL'; retry = true; }
  else if (numeric === 401 || numeric === 403) result = 'FAILED_CONFIGURATION';
  else if (numeric >= 400 && numeric < 500) result = 'FAILED_PERMANENT';
  return { sent: false, result: result, telegramMessageId: '', httpStatus: status || '',
    errorCode: code == null ? '' : String(code), errorText: redactWorkflowNotificationSecret_(text, token),
    retryEligible: retry };
}

/** One invocation performs at most one fetch and never retries. */
function sendTelegramMessage_(botToken, chatId, text) {
  const endpoint = 'https://api.telegram.org/bot' + botToken + '/sendMessage';
  let response;
  try {
    response = UrlFetchApp.fetch(endpoint, { method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ chat_id: String(chatId), text: String(text) }),
      muteHttpExceptions: true });
  } catch (error) {
    return telegramFailure_('', '', error && error.message ? error.message : error, botToken);
  }
  const status = response.getResponseCode(), body = response.getContentText();
  let parsed;
  try { parsed = JSON.parse(body); } catch (error) {
    return status >= 400
      ? telegramFailure_(status, status, 'Telegram returned an invalid error response.', botToken)
      : telegramFailure_(status, '', 'Telegram returned an ambiguous response.', botToken);
  }
  if (status >= 200 && status < 300 && parsed.ok === true && parsed.result &&
      parsed.result.message_id != null) {
    return { sent: true, result: 'SENT', telegramMessageId: String(parsed.result.message_id),
      httpStatus: status, errorCode: '', errorText: '', retryEligible: false };
  }
  const code = parsed && parsed.error_code != null ? parsed.error_code :
    (status >= 400 ? status : '');
  if (!code) return telegramFailure_(status, '', 'Telegram delivery outcome is ambiguous.', botToken);
  return telegramFailure_(status, code, parsed.description || 'Telegram request failed.', botToken);
}
