# Zhosik Nails — Онлайн-запись и CRM

Сервис для маникюрной студии. Работает на Google Sheets + GitHub Pages. Без серверов.

## Настройка

### 1. Google Таблица
Создайте таблицу, скопируйте ID из URL (между /d/ и /edit).

### 2. Apps Script
- Откройте Таблицу → Расширения → Apps Script
- Вставьте код из `backend/api.gs`
- В 1-й строке замените `SPREADSHEET_ID` на ваш
- Развернуть → Новое развертывание → Веб-приложение → Все
- Скопируйте URL

### 3. HTML
В 4 файлах (index.html, check.html, master.html, admin.html) замените `API_URL` на ваш URL.

### 4. GitHub Pages
Загрузите все файлы в репозиторий, включите Pages в Settings.

## Ссылки
- Запись: `index.html`
- Проверка: `check.html`
- Мастер: `master.html?access=zhosik-jasmine-secret`
- CRM: `admin.html?access=zhosik-jasmine-secret`