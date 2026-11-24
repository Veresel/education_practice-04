# -*- coding: utf-8 -*-
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml.ns import qn

root = Path(r"D:\учеб 04 практ")
ev = root / "evidence"
template = Path(r"C:\Users\nikol\Downloads\Otchyot o vyipolnenii proekta.docx")
out = root / "Отчёт о выполнении проекта.docx"
doc = Document(template)

def font(run, name="Times New Roman", size=12):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)

def set_para(p, text):
    p.clear()
    r = p.add_run(text)
    font(r)
    return r

def set_cell(cell, text):
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(text)
    font(r)

STAGES = [
    "Подготовка исходного состояния",
    "D1. Вход зарегистрированного пользователя",
    "D2. Возврат без потери заявки",
    "D3. Совместный отбор и числовая сортировка",
    "Отказ и исправление ввода",
    "Подтверждение и сохранение",
    "Дополнительный сценарий своего варианта",
    "Передача и ограничения",
    "Инструкция пользователя",
]

TEXTS = {
    "Подготовка исходного состояния": {
        "data": "Стенд варианта 19, baseline.json. Ожидание до запуска: Subject = «Ремонт измерителей качества воды»; пять Id P1–P5; исполнения BNC и DIN; P1.BNC=5, P1.DIN=3; Orders пуст; пользователь Гость.",
        "act": "Распакован Shop-variant-19.zip. Из папки stands\\Shop-source выполнены Copy-Item baseline.json → work-data.json и dotnet run --project .\\Shop.csproj -- .\\work-data.json.",
        "fact": "Над каталогом указана область варианта 19. Пять позиций, исходные суммарные остатки 8/3/3/3/3. Кнопка «Войти» доступна. После исправления Catalog числовой порядок: 90, 200, 1000, 1000, 1000 (P4, P3, P1, P2, P5).",
        "out": "Контрольное состояние совпадает с карточкой и JSON. Расхождений нет.",
    },
    "D1. Вход зарегистрированного пользователя": {
        "data": "nobody отсутствует в Users и должен получить отказ. client → Иванов Иван Иванович. buyer → Петрова Анна Сергеевна. Пустой логин — отказ. Сравнение логина чувствительно к регистру, Trim снимает крайние пробелы.",
        "act": "На исходном стенде nobody → «Войти». Причина: Authenticate использовал ?? users.FirstOrDefault(). Метод заменён на FirstOrDefault без запасного пользователя. Повтор nobody, затем client и buyer.",
        "fact": "До правки nobody получил ФИО Иванова и доступ к поиску. После правки отказ, пользователь Гость, поле nobody сохранено. client даёт своё ФИО; buyer — другое. Пустой логин отклонён.",
        "out": "Причина D1 — подмена неизвестного логина первым пользователем. Регрессия зарегистрированных логинов проходит.",
    },
    "D2. Возврат без потери заявки": {
        "data": "client, P1, исполнение BNC (остаток 5), количество 1. Ожидание: одна строка P1/BNC/1 до и после «Назад к каталогу» и открытия P5. Stock без подтверждения не меняется.",
        "act": "Исходный Back вызывал draft.Clear(). Очистка удалена; команды отмены и выхода не менялись. Повтор добавления P1 и возврата. Регрессия: отмена «Нет»/«Да».",
        "fact": "До правки после возврата таблица заявки пуста (статус ещё упоминает P1). После правки в карточке P5 остаётся P1 / BNC / 1. «Нет» сохраняет строки, «Да» очищает, JSON Stock без изменений.",
        "out": "Причина D2 — очистка коллекции при навигации. Переход экрана больше не отменяет заявку.",
    },
    "D3. Совместный отбор и числовая сортировка": {
        "data": "Запрос «сервисный», категория «Электроды». До запуска: P1 и P5 подходят; P2 той же категории без слова; P3 со словом, но категория «Модули обработки». Пустой поиск, все категории, возрастание: P4(90), P3(200), P1, P2, P5(1000).",
        "act": "Исходный Catalog связывал условия через || и сортировал DiscountPrice.ToString(). Заменено на && между категорией и текстом, || внутри имени/описания, OrderBy по decimal и ThenBy Id.",
        "fact": "До правки в выборку попали P2 и P3, 1000 стояло раньше 200. После правки при тех же входах только P1 и P5. Полный каталог: 90 → 200 → 1000. Пробелы по краям и пустой результат проверены.",
        "out": "Совместный отбор и числовой порядок соответствуют данным варианта 19.",
    },
    "Отказ и исправление ввода": {
        "data": "Чистая копия, client, P1/BNC, остаток 5. Сначала 3 (успех), затем ещё 3: отказ, так как 3+3>5. Поле, исполнение и первая строка сохраняются. Исправление на 2 даёт две строки суммарно 5.",
        "act": "Добавление 3, повтор 3, замена на 2. Отдельно: 0, -1, abc, 1,5 — отказ с именем поля «Количество», затем ввод 1.",
        "fact": "Сообщение: уменьшить значение с учётом уже добавленных позиций. Строка 3/BNC осталась, Stock=5, Orders пуст. После 2 в заявке две строки, списания нет до подтверждения.",
        "out": "Отказ не уничтожает заявку и логин. OrderService не переписывался.",
    },
    "Подтверждение и сохранение": {
        "data": "Независимая чистая копия. 2×P1/BNC и 1×P3/BNC. Итог 2200. После записи: один Order, Login=client; P1.BNC=3, P1.DIN=3; P3.BNC=1, P3.DIN=1.",
        "act": "Добавление через каталог, подтверждение, просмотр work-data.json, повторный запуск той же командой.",
        "fact": "Сообщение об оформлении и обновлённый JSON получены вместе. Каталог: P1 остаток 6, P3 остаток 2. После перезапуска 11.11.2026 те же суммы. Дата в Orders — фактическое время подтверждения.",
        "out": "Готовая операция сохранилась. Следующий независимый тест начинается с копии baseline.json.",
    },
    "Дополнительный сценарий своего варианта": {
        "data": "Карточка варианта 19: поиск по части названия. Первое слово P1 «Электрод» → P1, P2, P5; P3 и P4 исключены. Полное «Электрод pH» → P1 и P5, P2 исключён. «электрод ph» даёт тот же набор.",
        "act": "Чистая копия, client, все категории, запросы «Электрод», «Электрод pH», «электрод ph».",
        "fact": "Короткий запрос вернул три электрода. Длинный оставил P1 и P5 и убрал P2. Регистр состав не изменил.",
        "out": "Contains ищет подстроку; разные Id с похожими именами не сливаются. Сценарий карточки выполнен.",
    },
    "Передача и ограничения": {
        "data": "В репозитории ожидаются Shop.csproj, исходники, baseline.json и чистый work-data.json. Меняются только три метода RepairLogic.",
        "act": "Подготовлены README, .gitignore, отчёт со снимками. OrderService, Confirm_Click и модель не изменялись.",
        "fact": "Authenticate возвращает null при отказе; Catalog — новый список; Back не очищает draft. JSON — локальный стенд, не БД.",
        "out": "Граница УП04 соблюдена. URL GitHub вставляется после публикации репозитория.",
    },
    "Инструкция пользователя": {
        "data": "Инструкция должна повторять запуск и проверки на данных варианта 19 без образца «Контур».",
        "act": "Файл stands/Shop-source/README.md: восстановление JSON, вход, поиск, возврат, отказ количества, отмена, подтверждение.",
        "fact": "Команды совпадают с методичкой. Ключи Stock — BNC и DIN.",
        "out": "Инструкция пригодна для проверяющего. ФИО и группу на титуле нужно вписать вручную.",
    },
}

FIELDS = {
    "Исходные данные и ожидаемый результат": "data",
    "Выполненные действия": "act",
    "Фактический результат и проверка": "fact",
    "Вывод по этапу и объяснение расхождений": "out",
}

# Title fields
for p in doc.paragraphs:
    t = p.text.strip()
    if t.startswith("Номер варианта:"):
        set_para(p, "Номер варианта: 19")
    elif t.startswith("Предметная область:"):
        set_para(p, "Предметная область: Ремонт измерителей качества воды")

# Tables 1–3
t1 = doc.tables[0]
set_cell(t1.rows[1].cells[1], "Вариант 19. Ремонт измерителей качества воды")
set_cell(t1.rows[2].cells[1], "Архив Shop-variant-19.zip; папка stands\\Shop-source; JSON своего варианта")
set_cell(t1.rows[3].cells[1], "baseline.json — контрольная копия; work-data.json — рабочий файл запусков; Orders исходно пуст")
set_cell(t1.rows[4].cells[1], "P1 Электрод pH / Электроды / 1000 / BNC→5 DIN→3; P2 Электрод проводимости / Электроды / 1000 / BNC→2 DIN→1; P3 Преобразователь сигнала / Модули обработки / 200 / BNC→2 DIN→1; P4 Модуль индикации / Модули обработки / 90 / BNC→2 DIN→1; P5 Электрод pH — резервный узел / Электроды / 1000 / BNC→2 DIN→1")

t2 = doc.tables[1]
labels = ["Windows", "Версия .NET SDK", "Visual Studio или редактор C#", "Редактор отчёта"]
vers = ["Windows 11, сборка 26300", "10.0.401", "Visual Studio Code / Cursor", "Microsoft Word, бланк задания"]
need = [
    "Классическое WPF-приложение, запуск из Shop-source",
    "Целевая платформа net10.0-windows; dotnet run",
    "Правки только RepairLogic.cs",
    "Собственные снимки варианта 19",
]
for i in range(1, 5):
    if not t2.rows[i].cells[0].text.strip():
        set_cell(t2.rows[i].cells[0], labels[i - 1])
    set_cell(t2.rows[i].cells[1], vers[i - 1])
    set_cell(t2.rows[i].cells[2], need[i - 1])

t3 = doc.tables[2]
set_cell(t3.rows[1].cells[0], "Контур")
set_cell(t3.rows[1].cells[1], "Ремонт измерителей качества воды")
set_cell(t3.rows[1].cells[2], "Subject, заголовок каталога")
set_cell(t3.rows[2].cells[0], "48 В / 54 В, Модули питания")
set_cell(t3.rows[2].cells[1], "BNC / DIN, Электроды, Модули обработки")
set_cell(t3.rows[2].cells[2], "D2, D3, заявка, доп. сценарий")
set_cell(t3.rows[3].cells[0], "Демонстрационный JSON")
set_cell(t3.rows[3].cells[1], "baseline.json / work-data.json варианта 19")
set_cell(t3.rows[3].cells[2], "Все проверки")
set_cell(t3.rows[4].cells[0], "Ожидания образца")
set_cell(t3.rows[4].cells[1], "Пересчёт по своим Id, ценам и ключам Stock")
set_cell(t3.rows[4].cells[2], "Этапы отчёта и таблица 4")

# Stage blanks
stage = None
waiting = None
for p in doc.paragraphs:
    t = p.text.strip()
    if t in TEXTS:
        stage = t
        waiting = None
        continue
    if t in FIELDS and stage:
        waiting = FIELDS[t]
        continue
    if waiting and stage and (t == "" or set(t.replace(" ", "")) <= set("_")):
        set_para(p, TEXTS[stage][waiting])
        waiting = None
        continue
    if t.startswith("Подтверждения") or t.startswith("Зафиксировать") or t.startswith("Рисунок"):
        waiting = None

# Tail after table 4
mode = None
for p in doc.paragraphs:
    t = p.text.strip()
    if t.startswith("Состав передаваемых материалов"):
        mode = "files"
        continue
    if t.startswith("Ссылка на репозиторий"):
        mode = "repo"
        continue
    if t.startswith("Невыполненные требования"):
        mode = "miss"
        continue
    if mode and (t == "" or set(t.replace(" ", "")) <= set("_")):
        if mode == "files":
            set_para(p, "stands/Shop-source (Shop.csproj, RepairLogic.cs, OrderService.cs, модели, baseline.json, data.json, чистый work-data.json), README.md, отчёт DOCX, снимки evidence/*.png. Исключены bin, obj, .vs.")
        elif mode == "repo":
            set_para(p, "Репозиторий ещё не опубликован. После загрузки вставить полный URL, например https://github.com/<логин>/up04-v19, и проверить доступ преподавателя.")
        else:
            set_para(p, "Не опубликована ссылка GitHub. На титуле не заполнены ФИО и группа. Аттестационный лист, индивидуальное задание и презентация в этом комплекте не заполнялись: в запросе были стенд, assignment.pdf и бланк отчёта.")
        mode = None

# Images into 1x1 placeholder tables
images = [
    ev / "01-control-state.png",
    ev / "02-d1-nobody-before.png",
    ev / "03-d1-nobody-after.png",
    ev / "04-d1-client-after.png",
    ev / "05-d2-cart-before-back.png",
    ev / "06-d2-cart-after-back-before-fix.png",
    ev / "07-d2-cart-after-back-after-fix.png",
    ev / "08-d3-filter-before.png",
    ev / "09-d3-filter-after.png",
    ev / "10-quantity-reject.png",
    ev / "11-quantity-continue.png",
    ev / "12-two-lines-before-confirm.png",
    ev / "13-confirm-result.png",
    ev / "14-after-restart.png",
    ev / "15-extra-partial-name.png",
]
ph = [t for t in doc.tables if len(t.rows) == 1 and len(t.columns) == 1]
assert len(ph) >= 15, len(ph)
for tbl, img in zip(ph[:15], images):
    cell = tbl.cell(0, 0)
    for p in cell.paragraphs:
        p.clear()
    run = cell.paragraphs[0].add_run()
    run.add_picture(str(img), width=Inches(6.0))

# extra 15b after last placeholder
extra = ev / "15b-extra-full-name.png"
if extra.exists():
    run = ph[14].cell(0, 0).paragraphs[0].add_run()
    run.add_break()
    run2 = ph[14].cell(0, 0).paragraphs[0].add_run()
    run2.add_picture(str(extra), width=Inches(6.0))

# Table 4
t4 = doc.tables[-1]
rows = [
    "Исходное состояние и копии данных",
    "D1: вход и регрессия",
    "D2: возврат без потери заявки",
    "D3: совместный отбор и сортировка",
    "Отказ количества и продолжение",
    "Готовая операция и сохранение",
    "Дополнительный сценарий варианта",
    "Контракт RepairLogic и OrderService",
    "Инструкция пользователя",
    "Чистая копия, отчёт и репозиторий",
]
data = [
    ("Пять позиций, пустые Orders, P1 BNC=5 DIN=3", "Каталог области 19, 5 строк, Гость", "Совпадает с baseline.json"),
    ("nobody/пустой — отказ; client — Иванов", "nobody отклонён, поле сохранено; client вошёл", "D1 устранён"),
    ("P1/BNC/1 сохраняется после возврата и открытия P5", "Строка на месте; отмена «Нет»/«Да» работает отдельно", "D2 устранён"),
    ("«сервисный»+«Электроды» → P1, P5; сортировка decimal", "Исходный OR пропускал P2/P3; после правки только P1 и P5", "D3 устранён"),
    ("3+3>5 — отказ, исправление на 2", "Первая строка и поле сохранены, затем строки 3 и 2", "OrderService без переписывания"),
    ("2×P1 BNC + 1×P3 BNC; P1.BNC=3, P3.BNC=1", "Итог 2200; после перезапуска P1=6, P3=2", "Сохранение подтверждено"),
    ("«Электрод» vs «Электрод pH»", "P1,P2,P5 / затем P1,P5 без P2; регистр не влияет", "Сценарий карточки выполнен"),
    ("Только RepairLogic; OrderService готов", "Authenticate/Back/Catalog исправлены по образцу методички", "Граница соблюдена"),
    ("README из Shop-source", "Описаны вход, поиск, возврат, отмена, подтверждение", "Совпадает со стендом"),
    ("Чистый JSON, отчёт, репозиторий", "work-data восстановлен; отчёт со снимками", "URL GitHub — после публикации"),
]
for i, (exp, fact, concl) in enumerate(data, start=1):
    if not t4.rows[i].cells[0].text.strip():
        set_cell(t4.rows[i].cells[0], rows[i - 1])
    set_cell(t4.rows[i].cells[1], exp)
    set_cell(t4.rows[i].cells[2], fact)
    set_cell(t4.rows[i].cells[3], concl)

doc.save(out)
print("saved", out, out.stat().st_size)
