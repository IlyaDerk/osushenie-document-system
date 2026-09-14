# Notification Engine — Stage 4

## Scope

Stage 4 adds production orchestration, durable delivery lifecycle records, one bounded technical retry, stale validation, and notification trigger management. It does not enable production, deploy code, create live triggers, or change the 22-column notification-history contract.

## Daily production flow

`runWorkflowNotificationsDaily()` reads notification configuration first. Any value other than the exact normalized `true` returns `{ mode: "PRODUCTION", status: "DISABLED" }` before sheet validation, sheet reads, history access, or Telegram transport.

When enabled, the job requires the configured Telegram group chat ID and bot token, calculates the business date in `Europe/Moscow`, batch-reads documents and active workflow rules, and delegates eligibility to `evaluateWorkflowNotification_()`. Candidates receive the configured string target, `TELEGRAM` channel, and Moscow notification date.

Final business dedup runs in `reserveWorkflowNotificationPlan_()` under the shared document lock. Its pure callback builds the existing grouped model and Stage 3 Telegram delivery plan, maps every candidate to exactly one physical part, and produces business reservations plus attempt-1 prepared records. The repository performs one contiguous `setValues()`, calls `SpreadsheetApp.flush()`, and only then lets the lock return. Telegram is called sequentially after lock release. A write or flush failure therefore prevents every transport call.

Each transport outcome uses a result record ID allocated before the network request. Result persistence is idempotent by that exact ID and is locally retried up to three times without repeating Telegram. Unconfirmed persistence is returned as a redacted reconciliation diagnostic; the existing business reservation still blocks a duplicate daily send.

Invalid rule-notification permissions are persisted as aggregate `CONFIG_WARNING` records. Missing rule/action warnings may be stored on their business reservations. Neither changes the business key.

## Retry flow

`retryWorkflowNotificationDeliveries()` has the same disabled-first safety. Enabled runs select only attempt-1 `FAILED_TECHNICAL` outcomes at least 60 minutes old with an original production prepared record and no terminal outcome or attempt 2. `SENT`, permanent/configuration failures, unknown outcomes, stale skips, invalid timestamps, and every existing attempt-2 prepared/result record are ineligible.

Retry uses the exact attempt-1 message text and SHA-256 hash. It never reformats or creates a partial delivery. All reservations linked to the physical delivery are checked against one current batch read. Notification control ends only when the current document is both `Подписан с обеих сторон` and `В офисе`; only that combined completed state is stale (neither status nor location alone ends control). The whole part is also stale when any reserved document is missing, duplicated, inactive, has a changed status, blank/invalid transfer date, invalid nonblank status-change date, an unstarted workflow cycle, ambiguous matching rules, or a matching rule whose notification permission is `Нет`, blank, or invalid. A missing rule alone and a blank action are not stale.

A stale part atomically reserves an attempt-2 `SKIPPED_STALE_BEFORE_RETRY` result. A valid part atomically reserves `DELIVERY_PREPARED` attempt 2 with the original payload and hash. Both paths re-read history under the lock, and an existing attempt 2 wins. Only a successfully flushed prepared record permits the single retry transport call. Attempt-2 outcomes use the same idempotent result persistence. There is no attempt 3.

## Triggers

`setupWorkflowNotificationTriggers()` removes only project triggers whose handlers exactly equal `runWorkflowNotificationsDaily` or `retryWorkflowNotificationDeliveries`, then creates one daily trigger near 08:00 in `Europe/Moscow` and one shared 15-minute retry trigger. Repeated setup remains idempotent and unrelated triggers are untouched. `removeWorkflowNotificationTriggers()` removes only those same two exact handlers. Trigger setup never changes Script Properties or the production enabled switch.

## Operational boundary

No notification function writes to `Документы объектов`. Telegram recipients remain the configured group chat ID; employee Telegram nicknames are not transport targets. Enabling production, trigger creation in the live Apps Script project, deployment, and smoke testing remain manual owner actions.
