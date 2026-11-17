# Инструкция по работе с Git для курсового проекта N1 (Ktpt.CourseOrder)

**Дата создания:** 17 ноября 2026
**Проект:** `F:\курсовая Копейкин Андрей\Копейкин Андрей\Курсовой проект`
**Репозиторий:** уже создан локально, ветка `main`, первый коммит выполнен.

---

## 0. Что уже сделано (можно сразу переходить к разделу 5)

| Задача | Статус |
|---|---|
| Установлен Git 2.56.0 (portable, MinGit, `.git-tools\mingit`) | ✅ |
| Создан репозиторий в корне проекта | ✅ |
| Ветка `main` создана (вместо `master`) | ✅ |
| Настроены `user.name` / `user.email` | ✅ |
| Единый `.gitignore` в корне (190 строк) | ✅ |
| Добавлен `.gitattributes` для переводов строк | ✅ |
| Удалён пустой вложенный репозиторий `scr/Ktpt.CourseOrder/.git` | ✅ |
| Удалены дубликаты `Cursework.gitignore` и `scr/Ktpt.CourseOrder/.gitignore` | ✅ |
| Первый коммит: 210 файлов, 19,4 МБ | ✅ |
| Рабочее дерево чистое (`git status` — пусто) | ✅ |
| **Осталось только:** создать репозиторий на GitHub/GitLab и выполнить `push` | ⬜ |

---

## 1. Как вызвать Git в этом проекте

Git установлен **portable** (без прав администратора), поэтому `git` не прописан в `PATH`.
Используйте либо полный путь, либо переменную:

```powershell
# Вариант 1 — полный путь (работает всегда)
$git = "F:\курсовая Копейкин Андрей\Копейкин Андрей\Курсовой проект\.git-tools\mingit\cmd\git.exe"
& $git status

# Вариант 2 — задать переменную один раз в сессии PowerShell
$git = "F:\курсовая Копейкин Андрей\Копейкин Андрей\Курсовой проект\.git-tools\mingit\cmd\git.exe"
& $git log --oneline
```

### Как сделать `git` доступным одной командой (рекомендуется)

Добавьте папку MinGit в `PATH` **для пользователя**:

```powershell
$mingit = "F:\курсовая Копейкин Андрей\Копейкин Андрей\Курсовой проект\.git-tools\mingit\cmd"
$current = [Environment]::GetEnvironmentVariable('Path','User')
if ($current -notlike "*$mingit*") {
    [Environment]::SetEnvironmentVariable('Path', "$current;$mingit", 'User')
    Write-Output "Добавлено в PATH. Перезапустите терминал."
} else {
    Write-Output "Уже в PATH."
}
```

> После перезапуска терминала команда `git status` будет работать без путей.
> Каталог `.git-tools/` добавлен в `.gitignore`, поэтому в репозиторий он не попадёт.

### Альтернатива: установить Git «обычно»

Если нужен полноценный Git с графическим клиентом (GitHub Desktop) и контекстным менюем:

```powershell
winget install --id Git.Git -e
```

Установщик запросит права администратора — подтвердите UAC-вводом вручную.
После установки проверьте: `git --version`.

---

## 2. Текущее состояние репозитория

```
проект/
├── .git/                 ← репозиторий (210 файлов в коммите)
├── .gitattributes        ← нормализация переводов строк
├── .gitignore            ← единые правила игнорирования
├── .git-tools/           ← portable Git (в репозиторий НЕ входит)
├── docs/                 ← скриншоты, диаграммы, бэкапы документов
├── scr/                  ← исходный код C#/.NET 10 (Avalonia)
│   ├── Ktpt.CourseOrder/       ← основное приложение (+ .sln)
│   ├── Ktpt.CourseOrder.Auth/  ← авторизация
│   └── Ktpt.CourseOrder.Data/  ← доступ к данным (EF Core)
├── sql/                  ← SQL-скрипты
├── tools/                ← PowerShell-скрипты сборки документов
├── *.drawio              ← диаграммы (исходники)
├── Пояснительная записка.docx
└── Презентация.pptx
```

**Что НЕ попало в репозиторий (и правильно):**

- `scr/*/bin/`, `scr/*/obj/` — результаты сборки (≈ 690 МБ из 717 МБ!)
- `Копейкин Андрей (14).zip` — дублирующий архив
- `~$*` — временные файлы Office
- `.git-tools/` — portable Git

> **Важно:** если вы когда-нибудь добавите в `.gitignore` новые правила, уже
> отслеживаемые файлы продолжат отслеживаться. Снять отслеживание:
> `git rm --cached <путь>`.

---

## 3. Проверка, что всё настроено правильно

Выполните в PowerShell из корня проекта:

```powershell
Set-Location "F:\курсовая Копейкин Андрей\Копейкин Андрей\Курсовой проект"
$git = "$PWD\.git-tools\mingit\cmd\git.exe"

& $git status              # должно быть пусто
& $git log --oneline       # должен быть 1 коммит
& $git branch              # должна быть ветка main
& $git remote -v           # пока пусто — это нормально, см. раздел 5
```

Ожидаемый вывод `git log --oneline`:

```
d69c1d4 Initial commit: курсовой проект N1 (Ktpt.CourseOrder)
```

---

## 4. Ежедневная работа (после первого push)

### Основной цикл: изменение → коммит → push

```powershell
$git = "$PWD\.git-tools\mingit\cmd\git.exe"

& $git status                    # 1. что изменилось
& $git add .                     # 2. добавить всё в индекс
& $git commit -m "Описание изменения"   # 3. зафиксировать
& $git push                      # 4. отправить на сервер
```

### Частые команды

```powershell
& $git status                    # состояние рабочего дерева
& $git diff                      # что конкретно изменилось (незакоммиченное)
& $git log --oneline --graph     # история коммитов
& $git log --stat                # история со списком файлов
& $git show <commit>             # подробности одного коммита
& $git restore <файл>            # отменить изменения в файле (не коммитить)
& $git restore --staged <файл>   # убрать файл из индекса, оставить в файле
```

### Отмена последнего коммита

```powershell
& $git reset --soft HEAD~1   # отменить коммит, файлы останутся в индексе
& $git reset --hard HEAD~1   # отменить коммит и все изменения (ОСТОРОЖНО)
```

### Отправить одну ветку и установить её как upstream (один раз)

```powershell
& $git push -u origin main
```

После этого достаточно `& $git push`.

### Забрать изменения с сервера

```powershell
& $git pull
```

> Если при `pull` возникает конфликт — сообщите об этом, не пробуйте разрешать
> его вслепую. Безопаснее сначала посмотреть конфликт:
> `& $git diff --name-only --diff-filter=U`

---

## 5. Публикация на GitHub / GitLab (главный оставшийся шаг)

### 5.1. Создайте пустой репозиторий на сервере

На GitHub: **New repository → НЕ ставить галочки** «Add a README», «Add .gitignore», «Add license»
(иначе будут конфликты при первом push).

На GitLab: **New project → Create blank project**.

### 5.2. Подключите удалённый репозиторий

```powershell
Set-Location "F:\курсовая Копейкин Андрей\Копейкин Андрей\Курсовой проект"
$git = "$PWD\.git-tools\mingit\cmd\git.exe"

# Для GitHub:
& $git remote add origin https://github.com/<ВАШ_ЛОГИН>/<ИМЯ_РЕПОЗИТОРИЯ>.git

# Для GitLab:
& $git remote add origin https://gitlab.com/<ВАШ_ЛОГИН>/<ИМЯ_РЕПОЗИТОРИЯ>.git

& $git remote -v     # проверка
```

### 5.3. Отправьте код

```powershell
& $git push -u origin main
```

GitHub/GitLab попросят авторизацию. Выберите **«Открыть в браузере»** или введите
**Personal Access Token** вместо пароля (GitHub с 2021 г. не принимает пароль
для Git-операций по HTTPS).

> **Токен для GitHub:** Settings → Developer settings → Personal access tokens →
> Tokens (classic) → Generate new token → отметить `repo` → скопировать токен.
> **Токен для GitLab:** User Settings → Access Tokens → добавить токен со scope `write_repository`.

### 5.4. Проверка

```powershell
& $git status              # должно быть: "nothing to commit, working tree clean"
& $git log --oneline -1     # последний коммит
& $git remote -v            # адрес remote
```

Откройте адрес репозитория в браузере — там должны быть все 210 файлов.

---

## 6. Что делать, если что-то пошло не так

| Проблема | Решение |
|---|---|
| `git` не найден | Используйте полный путь `$PWD\.git-tools\mingit\cmd\git.exe` (раздел 1) |
| `fatal: detected dubious ownership` | `& $git config --global --add safe.directory "*"` — уже выполнено |
| `Updates were rejected` | Сначала `& $git pull --rebase`, затем `& $git push` |
| `The remote end hung up unexpectedly` | Проверьте интернет и авторизацию; повторите `push` |
| `fatal: Authentication failed` | Используйте Personal Access Token вместо пароля |
| `Please tell me who you are` | `& $git config user.name "Имя"` и `& $git config user.email "почта"` |
| Случайно добавили лишнее (например, `bin/`) | `& $git rm -r --cached <путь>` + добавить правило в `.gitignore` |
| Забыли `.gitignore` и уже запушили большой файл | См. раздел 7 |

---

## 7. Если большой файл уже попал в историю

В репозиторий уже попал файл размером 3,4 МБ (`Копейкин Андрей (14).zip`).
Это ниже лимита GitHub (100 МБ), поэтому проблем нет. Но если в будущем
случится превышение лимита, исправление выглядит так:

```powershell
# Переписываем историю, удаляя файл из последнего коммита
& $git filter-branch --force --index-filter `
    "git rm --cached --ignore-unmatch 'путь/к/файлу'" HEAD

# Принудительная отправка
& $git push --force origin main
```

> Предупреждение: `filter-branch` переписывает историю. Делайте только до того,
> как с репозиторием работали другие люди.

---

## 8. Рекомендации по работе с репозиторием

1. **Коммитьте часто и логично.** Один коммит — одно законченное изменение:
   «Добавил экран корзины», «Исправил расчёт скидки», «Обновил диаграмму архитектуры».
2. **Сообщения коммитов — на русском или английском**, осмысленные.
   Не пишите `111`, `asdf`, `fix`.
3. **Не коммитьте результаты сборки** — они уже исключены в `.gitignore`.
4. **Перед крупным этапом работы** сделайте коммит, чтобы было к чему вернуться:
   `& $git checkout -b feature/orders-history`
5. **Не храните в Git пароли и строки подключения к БД.**
   Если они попали в историю — см. раздел 7.
6. **Резервная копия**: репозиторий на GitHub — это и есть бэкап всех версий.

---

## 9. Частые ошибки новичков

| Ошибка | Почему это плохо | Как правильно |
|---|---|---|
| `git add .` в папке с `bin/` | В коммит попадут 690 МБ | Уже исправлено — `bin/obj` в `.gitignore` |
| Коммит всех файлов одной кучей | Невозможно найти нужную версию | Дробите на логичные коммиты |
| `git push --force` без необходимости | Ломает историю у коллег | Используйте `--force-with-lease` |
| Удаление `.git` вручную | Теряется вся история | `git rm` не удаляет историю, а файл из индекса |
| Хранение `connectionString` с паролем | Утечка учётных данных | Держите в `.gitignore` или в переменных окружения |

---

## 10. Шпаргалка

```powershell
# Переменная (в начале каждой сессии)
$git = "F:\курсовая Копейкин Андрей\Копейкин Андрей\Курсовой проект\.git-tools\mingit\cmd\git.exe"

# Первичная настройка (уже выполнена)
& $git config --global --add safe.directory "*"
& $git config user.name "Andrey Kopeykin"
& $git config user.email "kopeykin@example.com"
& $git config core.quotePath false      # чтобы кириллица в выводе не экранировалась

# Основной цикл
& $git status
& $git add .
& $git commit -m "Комментарий"
& $git push

# Публикация (разово)
& $git remote add origin https://github.com/<Логин>/<Репозиторий>.git
& $git push -u origin main
```

---

## Приложение: структура проекта для справки

| Путь | Содержимое |
|---|---|
| `scr/Ktpt.CourseOrder/` | Основное приложение (Avalonia UI, ViewModels, Views, Services) |
| `scr/Ktpt.CourseOrder.Auth/` | Модуль авторизации |
| `scr/Ktpt.CourseOrder.Data/` | Модуль доступа к данным (Entity Framework Core) |
| `scr/Ktpt.CourseOrder/Ktpt.CourseOrder.sln` | Solution-файл, открывать в Visual Studio / Rider |
| `sql/` | `SQLCreateTableTable.sql`, `insert_sqlserver.sql` — схема и данные |
| `tools/` | PowerShell-скрипты генерации DOCX/PPTX (сборка через `dotnet build`) |
| `docs/diagrams/` | PNG-диаграммы (архитектура, ER, слои, алгоритмы) |
| `docs/deck_png/` | PNG-слайды презентации |
| `docs/screens/` | Скриншоты работающего приложения |
| `*.drawio` | Исходники диаграмм (draw.io) |
