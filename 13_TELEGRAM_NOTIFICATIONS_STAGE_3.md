# Notification Engine — Stage 3

Stage 3 adds manual Telegram test delivery and a production-style read-only preview.
It does not add production scheduling, triggers, retries, direct employee messages, or
document writes.

## Configuration

The only notification Script Properties are `WORKFLOW_NOTIFICATIONS_ENABLED`,
`WORKFLOW_TG_CHAT_ID`, and `WORKFLOW_TG_BOT_TOKEN`. Only the trimmed, case-insensitive
value `true` enables future production operation; missing or unknown values fail closed.
Dry run requires the chat ID only. Manual test send requires both chat ID and token and
is intentionally independent of the production switch.

## Telegram output

Reports are deterministic plain text grouped by the Stage 1 object model. Physical
messages are capped at 3900 JavaScript characters, receive final multipart headings,
and only the last part receives the Operator Card link. SHA-256 is calculated over the
exact final physical text. Object blocks may split only between documents. A document
block that cannot fit a physical message fails closed before history or network effects;
one business event is therefore never referenced by multiple physical deliveries.

## Test delivery lifecycle

All `TEST_DELIVERY/PREPARED` rows are appended before the first network call. Each part
is then sent once, in order, followed by a corresponding `TEST_DELIVERY` outcome. A
failure to persist prepared history prevents all sends. A failure to persist outcomes
is returned as a warning and never causes a resend. Test records neither read nor write
production reservation/prepared state.

Telegram transport performs one JSON `sendMessage` POST per invocation. Explicit 429
and 5xx responses are technical/retry-eligible; 401/403 are configuration failures;
other 4xx responses are permanent. Ambiguous 2xx responses and network exceptions are
unknown outcomes and are not retry-eligible. Stage 3 only classifies retry eligibility;
it never retries.
