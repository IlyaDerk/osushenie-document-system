# Notification Engine — Stage 2

## Scope

Stage 2 adds notification state storage and a read-only preview. It does not
send messages, format or split Telegram payloads, configure credentials,
create triggers, or retry deliveries.

## History contract

`SystemCore.gs` defines the required system sheet `История уведомлений`, with
headers on row 2 and append-only data beginning on row 3. Its 22 columns are
the approved notification-history contract. The deployment process creates
the physical sheet manually; application code fails closed if it is absent or
its required headers are missing or duplicated.

All repository reads and writes are header-driven. Existing rows are never
updated, deleted, cleared, sorted, or repaired. A state transition is a new
row.

## Deduplication and atomic reservation

The production key is the canonical combination:

`document ID | yyyy-MM-dd | event | channel | target`

Only `BUSINESS_RESERVATION` creates a production key. A reservation therefore
blocks another production reservation even if no later delivery result exists
or a later result reports failure. `TEST_DELIVERY` never participates in this
index.

`reserveWorkflowNotificationPlan_()` accepts preliminary business candidates
and a pure delivery-builder callback. Under the common document lock it:

1. batch-reads current history;
2. folds the latest production dedup state in memory;
3. gives only accepted events to the pure builder;
4. validates final reservations and prepared deliveries;
5. appends all rows through one contiguous `setValues()` call.

This callback design allows the future formatter to rebuild exact payloads
after the final locked dedup decision, without persisting an intermediate
reservation and without retaining an empty or stale prepared delivery.

## Dry run

`dryRunWorkflowNotifications()` batch-reads documents, active rules, and
history; uses the Moscow business date and Stage 1 evaluation; previews
production dedup against the deterministic `DRY_RUN_GROUP_TARGET`; aggregates
warnings; and returns counts, event counts, skips, dedup details, and the
grouped Stage 1 model.

The dry run only reads spreadsheets. It does not write history or documents,
access Script Properties, create triggers, call a transport, or build Telegram
text.
