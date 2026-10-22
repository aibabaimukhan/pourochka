# Поурочный план на Gemini — установка

## 1. Apps Script (бэкенд + данные)
1. Создайте Google Таблицу → **Расширения → Apps Script**.
2. Вставьте содержимое `Code.gs` вместо кода по умолчанию, сохраните.
3. **Project Settings → Script properties → Add**:
   - `GEMINI_API_KEY` — ключ из Google AI Studio
   - `ADMIN_PASSWORD` — пароль дашборда
   - `GEMINI_MODEL` — (необязательно) по умолчанию `gemini-3.6-flash`
4. Выберите функцию `setup` → **Run** → выдайте доступы (создаст листы `Log` и `Blocked`).
5. **Deploy → New deployment → Web app**: Execute as **Me**, Who has access **Anyone** → скопируйте URL (`.../exec`).
   После любых правок `Code.gs`: Deploy → Manage deployments → Edit → New version.

## 2. GitHub Pages
1. В `config.js` замените `ВСТАВЬТЕ_URL_ВЕБ_ПРИЛОЖЕНИЯ` на URL из шага 1.5.
2. Загрузите в репозиторий: `index.html`, `admin.html`, `config.js`, `style.css`, `docbuilder.js`, `template.docx`.
3. Settings → Pages → Deploy from branch → `main` / root.
4. Преподаватели: `https://<логин>.github.io/<репо>/`, дашборд: `.../admin.html`.

## Как работает
- Gemini возвращает только текст (JSON); `.docx` собирается в браузере из `template.docx` и сразу скачивается, нигде не хранится.
- В таблице хранится только: время, ФИО, группа, тема, дата урока, тип, язык, модуль, статус.
- Блокировка: запись «Иванов» блокирует любого, у кого в ФИО есть слово «Иванов»; «Иванов А.Б.» — только тех, у кого есть все три слова.
- Ключ Gemini и пароль лежат только в Script properties, в код сайта они не попадают.
