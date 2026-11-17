# Инструкция по работе с Git для курсового проекта N1 (Ktpt.CourseOrder)

**Дата создания:** 17 ноября 2026
**Обновлено:** 17 ноября 2026 — проект перенесён, опубликован на GitHub
**Проект:** `C:\Projects\Curse-work-N1`
**Репозиторий на GitHub:** https://github.com/COODELC/Curse-work-N1

> **Почему проект перенесён.** Исходный путь `F:\курсовая Копейкин Андрей\...`
> содержал пробелы и кириллицу. Git запускает credential helper через `sh`,
> а `sh` не может разрешить такой путь — push не проходил. Короткий путь без
> пробелов и кириллицы проблему снимает.

---

## 0. Что уже сделано

| Задача | Статус |
|---|---|
| Установлен Git 2.56.0 (portable, MinGit, `.git-tools\mingit`) | ✅ |
| Установлен GitHub CLI 2.63.2 (portable, `.git-tools\gh\bin\gh.exe`) | ✅ |
| Проект перенесён в `C:\Projects\Curse-work-N1` | ✅ |
| Создан репозиторий, ветка `main` | ✅ |
| Настроены `user.name` / `user.email` | ✅ |
| Единый `.gitignore` в корне (190 строк) | ✅ |
| Добавлен `.gitattributes` для переводов строк | ✅ |
| Удалён пустой вложенный репозиторий `scr/Ktpt.CourseOrder/.git` | ✅ |
| Удалены дубликаты `Cursework.gitignore` и `scr/Ktpt.CourseOrder/.gitignore` | ✅ |
| Выполнена авторизация GitHub CLI под ником **COODELC** | ✅ |
| Remote `origin` → `https://github.com/COODELC/Curse-work-N1.git` | ✅ |
| **Код опубликован: 4 коммита, 214 файлов** | ✅ |
| Рабочее дерево чистое, upstream = `origin/main` | ✅ |

---

## 1. Как вызвать Git в этом проекте

Git установлен **portable** (без прав администратора), поэтому `git` не прописан в `PATH`.
Используйте либо полный путь, либо переменную:

```powershell
# Вариант 1 — полный путь (работает всегда)
$git = "C:\Projects\Curse-work-N1\.git-tools\mingit\cmd\git.exe"
& $git status

# Вариант 2 — задать переменную один раз в сессии PowerShell
$git = "C:\Projects\Curse-work-N1\.git-tools\mingit\cmd\git.exe"
& $git log --oneline
```

### Как сделать `git` доступным одной командой (рекомендуется)

```powershell
$mingit = "C:\Projects\Curse-work-N1\.git-tools\mingit\cmd"
$gh     = "C:\Projects\Curse-work-N1\.git-tools\gh\bin"
$current = [Environment]::GetEnvironmentVariable('Path','User')
foreach ($p in @($mingit, $gh)) {
    if ($current -notlike "*$p*") { $current = "$current;$p" }
}
[Environment]::SetEnvironmentVariable('Path', $current, 'User')
Write-Output "Добавлено в PATH. Перезапустите терминал."
```

> После перезапуска терминала `git status` и `gh auth status` будут работать без путей.
> Каталог `.git-tools/` добавлен в `.gitignore`, поэтому в репозиторий он не попадает.

---

## 2. Текущее состояние репозитория

```
C:\Projects\Curse-work-N1\
├── .git\                  ← репозиторий (214 файлов, 4 коммита)
├── .gitattributes         ← нормализация переводов строк
├── .gitignore             ← единые правила игнорирования
├── .git-tools\            ← portable Git + GitHub CLI (в репозиторий НЕ входит)
├── docs\                  ← скриншоты, диаграммы, бэкапы документов
├── scr\                   ← исходный код C#/.NET 10 (Avalonia)
│   ├── Ktpt.CourseOrder\       ← основное приложение (+ .sln)
│   ├── Ktpt.CourseOrder.Auth\  ← авторизация
│   └── Ktpt.CourseOrder.Data\  ← доступ к данным (EF Core)
├── sql\                   ← SQL-скрипты
├── tools\                 ← PowerShell-скрипты сборки документов
├── *.drawio               ← диаграммы (исходники)
├── Пояснительная записка.docx
└── Презентация.pptx
```

**Что НЕ попало в репозиторий (и правильно):**

- `scr/*/bin/`, `scr/*/obj/` — результаты сборки (≈ 690 МБ из 717 МБ!)
- `Копейкин Андрей (14).zip` — дублирующий архив
- `~$*` — временные файлы Office
- `.git-tools/` — portable Git и GitHub CLI

---

## 3. Проверка, что всё настроено правильно

```powershell
Set-Location "C:\Projects\Curse-work-N1"
$git = "$PWD\.git-tools\mingit\cmd\git.exe"

& $git status              # должно быть пусто
& $git log --oneline       # 4 коммита
& $git branch              # main
& $git remote -v           # COODELC/Curse-work-N1
```

Ожидаемый вывод `git log --oneline`:

```
60702a4 docs: README в UTF-8 со структурой проекта и командой сборки
f95a997 fix: не игнорировать .md файлы; добавить инструкцию по Git
c9b3ae7 docs: add .gitattributes and Git usage instruction
d69c1d4 Initial commit: курсовой проект N1 (Ktpt.CourseOrder)
```

---

## 4. Ежедневная работа

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
& $git diff                      # что изменилось (незакоммиченное)
& $git log --oneline --graph     # история коммитов
& $git log --stat                # история со списком файлов
& $git show <commit>             # подробности одного коммита
& $git restore <файл>            # отменить изменения в файле
& $git restore --staged <файл>   # убрать файл из индекса, оставить в файле
```

### Отмена последнего коммита

```powershell
& $git reset --soft HEAD~1   # отменить коммит, файлы останутся в индексе
& $git reset --hard HEAD~1   # отменить коммит и все изменения (ОСТОРОЖНО)
```

### Забрать изменения с сервера

```powershell
& $git pull
```

> При конфликте посмотрите конфликт: `& $git diff --name-only --diff-filter=U`

---

## 5. Публикация на GitHub (уже выполнена)

Репозиторий создан под ником **COODELC**: https://github.com/COODELC/Curse-work-N1

```powershell
& $git remote -v
# origin  https://github.com/COODELC/Curse-work-N1.git (fetch)
# origin  https://github.com/COODELC/Curse-work-N1.git (push)
```

Upstream уже настроен — достаточно `& $git push`.

**Если потребуется создать новый репозиторий:**

```powershell
& $git remote add origin https://github.com/<Логин>/<Репозиторий>.git
& $git push -u origin main
```

---

## 6. Что делать, если что-то пошло не так

| Проблема | Решение |
|---|---|
| `git` не найден | Используйте полный путь `$PWD\.git-tools\mingit\cmd\git.exe` |
| `gh: command not found` при push | Добавьте `.git-tools\gh\bin` в `PATH` (раздел 1) |
| `fatal: detected dubious ownership` | `& $git config --global --add safe.directory "*"` — уже выполнено |
| `CRYPT_E_REVOCATION_OFFLINE` | `& $git config --global http.schannelCheckRevocation false` — уже выполнено |
| `Updates were rejected` | Сначала `& $git pull --rebase`, затем `& $git push` |
| `fatal: Authentication failed` | Токен истёк: `& "$PWD\.git-tools\gh\bin\gh.exe" auth login` |
| `Please tell me who you are` | `& $git config user.name "Имя"` и `& $git config user.email "почта"` |
| Случайно добавили лишнее (например, `bin/`) | `& $git rm -r --cached <путь>` + правило в `.gitignore` |

### Проверка авторизации

```powershell
& "$PWD\.git-tools\gh\bin\gh.exe" auth status
```

Ожидаемый вывод:

```
github.com
  ✓ Logged in to github.com account COODELC (keyring)
  - Token scopes: 'gist', 'read:org', 'repo'
```

### Переавторизация (если токен истёк)

```powershell
& "$PWD\.git-tools\mingit\cmd\git.exe" config --global credential.helper `
    "!C:/Projects/Curse-work-N1/.git-tools/gh/bin/gh.exe auth git-credential"
& "$PWD\.git-tools\gh\bin\gh.exe" auth login --hostname github.com --git-protocol https --web
```

> **Внимание:** в helper должен стоять **абсолютный путь со слэшами** (`C:/...`).
> Путь с пробелами или кириллицей `sh` не разрешит — push упадёт с
> `gh: command not found`.

---

## 7. Если большой файл попал в историю

Лимит GitHub — 100 МБ на файл. Сейчас максимальный файл в репозитории — 3,5 МБ
(`Презентация_курсовой_проект_Копейкин.pptx`), запас большой. Если лимит всё же
превышен:

```powershell
& $git filter-branch --force --index-filter `
    "git rm --cached --ignore-unmatch 'путь/к/файлу'" HEAD
& $git push --force origin main
```

> `filter-branch` переписывает историю. Делайте только до того, как с репозиторием
> работали другие люди.

---

## 8. Рекомендации по работе с репозиторием

1. **Коммитьте часто и логично.** Один коммит — одно законченное изменение:
   «Добавил экран корзины», «Исправил расчёт скидки», «Обновил диаграмму архитектуры».
2. **Сообщения коммитов — осмысленные.** Не пишите `111`, `asdf`, `fix`.
3. **Не коммитьте результаты сборки** — они уже исключены в `.gitignore`.
4. **Перед крупным этапом работы** сделайте коммит или ветку:
   `& $git checkout -b feature/orders-history`
5. **Не храните в Git пароли и строки подключения к БД.**
6. **Репозиторий на GitHub — это и есть бэкап всех версий.**

---

## 9. Частые ошибки новичков

| Ошибка | Почему это плохо | Как правильно |
|---|---|---|
| `git add .` в папке с `bin/` | В коммит попадут 690 МБ | Уже исправлено — `bin/obj` в `.gitignore` |
| Коммит всех файлов одной кучей | Невозможно найти нужную версию | Дробите на логичные коммиты |
| `git push --force` без необходимости | Ломает историю у коллег | Используйте `--force-with-lease` |
| Удаление `.git` вручную | Теряется вся история | `git rm` убирает файл из индекса, а не из истории |
| Хранение `connectionString` с паролем | Утечка учётных данных | Держите в `.gitignore` или в переменных окружения |
| Путь с пробелами/кириллицей для `gh` | `sh` не находит helper → push не работает | Короткий путь + абсолютный путь в helper |

---

## 10. Шпаргалка

```powershell
# Переменные (в начале каждой сессии)
Set-Location "C:\Projects\Curse-work-N1"
$git = "$PWD\.git-tools\mingit\cmd\git.exe"
$gh  = "$PWD\.git-tools\gh\bin\gh.exe"

# Основной цикл
& $git status
& $git add .
& $git commit -m "Комментарий"
& $git push

# Проверки
& $git log --oneline
& $git remote -v
& $gh auth status
```

---

## Приложение: структура проекта для справки

| Путь | Содержимое |
|---|---|
| `scr/Ktpt.CourseOrder/` | Основное приложение (Avalonia UI, ViewModels, Views, Services) |
| `scr/Ktpt.CourseOrder.Auth/` | Модуль авторизации |
| `scr/Ktpt.CourseOrder.Data/` | Модуль доступа к данным (Entity Framework Core) |
| `scr/Ktpt.CourseOrder/Ktpt.CourseOrder.sln` | Solution-файл, открывать в Visual Studio / Rider |
| `sql/` | `SQLCreateTable.sql`, `insert_sqlserver.sql` — схема и данные |
| `tools/` | PowerShell-скрипты генерации DOCX/PPTX (сборка через `dotnet build`) |
| `docs/diagrams/` | PNG-диаграммы (архитектура, ER, слои, алгоритмы) |
| `docs/deck_png/` | PNG-слайды презентации |
| `docs/screens/` | Скриншоты работающего приложения |
| `*.drawio` | Исходники диаграмм (draw.io) |

### Сборка и запуск

```powershell
dotnet build scr\Ktpt.CourseOrder\Ktpt.CourseOrder.sln
dotnet run --project scr\Ktpt.CourseOrder\Ktpt.CourseOrder.csproj
```
