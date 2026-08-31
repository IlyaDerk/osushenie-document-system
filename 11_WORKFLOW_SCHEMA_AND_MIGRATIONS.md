# Workflow schema and data migrations

This change prepares data contracts after owner approval. It does not install
columns in Google Sheets and does not implement a notification engine.

## Statuses and KS2/KS3

New documents start in `На подготовке`. The data-only status migration changes
only `Ожидает заполнения` to `На подготовке` and `Передан заказчику` to
`На согласовании у заказчика`. `Подписан у заказчика` remains an ordinary
dictionary status and has no special Manager Summary meaning.

The existing KS-2 type is renamed to `КС2/КС3` while retaining its existing
type ID. No KS-3 ID is created or restored. The migration identifies rows only
by the explicitly supplied retained ID. It changes a number only when the
entire value matches the legacy generated form `<legacy name> №<positive
integer>`; manual numbers are preserved.

Public entry points are:

* `migrateDocumentStatuses()`;
* `migrateKs2Ks3DocumentType(retainedTypeId, legacyTypeName)`;
* `migrateDocumentWorkflowData(retainedTypeId, legacyTypeName)`.

All three use the shared document lock, resolve columns from `SystemCore.gs`,
write change and operation history, return counters, and are idempotent. The
combined and KS entry points intentionally require the retained ID rather than
guessing it from names.

## Contracts

`Документы объектов` has 36 required headers at row 3 (data begins at row 4).
`Карточка операциониста` has 31 required headers at row 5 (data begins at row
6). Both add `Когда передан` and `Дней на реализацию` immediately after `Кто
передал`. They are read-only in the Operator Card field map in this first
change; edit/save UX belongs to the follow-up change.

`Справочник клиентов` has ten headers at row 4 (data begins at row 5), adding
`Адрес клиента` after `Телефон`. One row remains one contact, `ID клиента`
remains the contact ID, and no company entity is introduced.

`Справочник сотрудников` adds `Участвует в документообороте` (`Да`/`Нет`). The
pure eligibility helper combines an active employee with `Да`. This flag is
reserved for future `У кого документ` and the employee part of `Кто передал`;
it does not constrain foremen or Web users and does not change ST/CL IDs.

`Справочник условий и действий` has headers at row 4 and data from row 5:
`ID правила`, `Тип документа`, `Статус документа`, `Где документ`, `Дней на
реализацию`, `Действие`, `Активно`. The reader returns only active (`Да`) rules
and rejects an invalid active rule deadline. This is a mandatory system sheet:
the common `validateSystemStructure()` check requires both the sheet and its
exact `WORKFLOW_RULES` header contract even when it contains no data rows. The
code does not create the sheet or change its visibility; the owner hides it
manually in the working spreadsheet.

This is a deployment prerequisite, not a rules-data prerequisite: create the
sheet manually before installing the new code regardless of whether any rule
rows have been prepared. Do not rely on the code to provision an empty sheet.

## Validation and future deadline calculation

`Когда передан` accepts blank or a valid `Date`. `Дней на реализацию` accepts
blank or an integer greater than or equal to zero; zero is filled and valid.
The pure effective-days helper applies `manual → matching active rule → 7` and
never writes a fallback into a document. A control date is returned only when
`Когда передан` is filled.

No automatic object synchronization writes either new document-owned field.
No Telegram integration, scheduler, trigger, workflow engine, dynamic address
dropdown, or new Operator Card sorting is included.

## Manual Sheet work after merge

Before deploying code or running migrations on a copy, the owner must:

1. add `Когда передан` and `Дней на реализацию` in the documented positions on
   `Документы объектов` and `Карточка операциониста`;
2. add `Адрес клиента` after `Телефон` on `Справочник клиентов`;
3. add `Участвует в документообороте` on `Справочник сотрудников` and fill
   every applicable row with `Да` or `Нет`;
4. obligatorily create `Справочник условий и действий` with its seven exact
   headers before installing or running the new code, even if it initially has
   no rules; hide the sheet manually for ordinary users if required;
5. add `На подготовке`, `На согласовании у заказчика`, and `Подписан у
   заказчика` as required on `Справочник для КО`, retaining any legacy values
   until the status migration has run;
6. rename the existing KS-2 row in `Справочник документов` to `КС2/КС3`
   without changing its retained type ID and without creating KS-3;
7. run the combined migration with that retained ID, review its report and both
   histories, then rerun it to confirm zero changes.

These are manual production-sheet changes; no migration inserts, moves,
renames, or deletes physical columns or headers.
