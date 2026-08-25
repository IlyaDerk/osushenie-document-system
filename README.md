# Пакет задания для Codex

## Содержимое

- `01_SHEET_STRUCTURE.md` — точная структура 9 листов и заголовков.
- `02_CREATION_MODULE_SPEC.md` — утверждённая бизнес-логика модуля.
- `03_ACCEPTANCE_TESTS.md` — 15 сценариев приёмочного тестирования.
- `04_ARCHITECTURE_GUARD.md` — запрет изменения архитектуры без согласования.
- `CODEX_PROMPT.txt` — готовый промт для Codex.
- `SOURCE_FILES_REQUIRED.md` — список актуальных файлов Apps Script, которые нужно добавить.

## Как использовать

1. Экспортируйте из Apps Script актуальные:
   - `SystemCore.gs`
   - `Code.gs`
   - `OperatorSidebar.html`
2. Поместите их рядом с файлами этого пакета.
3. Передайте весь набор Codex.
4. Вставьте содержимое `CODEX_PROMPT.txt` как основное задание.
5. Проверяйте результат сначала на копии Google-таблицы.

## Карточка операциониста: чтение

Рабочая read-only загрузка карточки реализована в `OperatorCard.gs`, публичные точки входа находятся в `Code.gs`, интерфейс — в `OperatorSidebar.html`. Подробный контракт, автоматические проверки и шесть сценариев ручной приёмки описаны в [05_OPERATOR_CARD_READ_SPEC.md](05_OPERATOR_CARD_READ_SPEC.md).

## Карточка объектов

Generated read-only отчёт реализован изолированно в `ObjectCardReport.gs`.
Архитектура snapshot, правила полного учёта и инструкция ручного hourly trigger
описаны в [09_OBJECT_CARD_REPORT_SPEC.md](09_OBJECT_CARD_REPORT_SPEC.md).
