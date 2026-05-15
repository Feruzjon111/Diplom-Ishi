(() => {
    const supportedLanguages = ["uz", "ru", "en"];
    const translations = {
        "e-Praktika": {
            ru: "e-Praktika",
            en: "e-Praktika",
        },
        "Bosh sahifa": {
            ru: "Главная",
            en: "Home",
        },
        "Bosh sahifaga": {
            ru: "На главную",
            en: "To home",
        },
        "Bosh sahifaga qaytish": {
            ru: "Вернуться на главную",
            en: "Back to home",
        },
        "Yuklash": {
            ru: "Загрузка",
            en: "Upload",
        },
        "Yuklash bo'limi": {
            ru: "Раздел загрузки",
            en: "Upload section",
        },
        "Yuklash bo‘limi": {
            ru: "Раздел загрузки",
            en: "Upload section",
        },
        "Profil": {
            ru: "Профиль",
            en: "Profile",
        },
        "Sozlamalar": {
            ru: "Настройки",
            en: "Settings",
        },
        "Akkount sozlamalari": {
            ru: "Настройки аккаунта",
            en: "Account settings",
        },
        "Fayl yuklash": {
            ru: "Загрузить файл",
            en: "Upload file",
        },
        "Chiqish": {
            ru: "Выйти",
            en: "Logout",
        },
        "Kirish": {
            ru: "Войти",
            en: "Login",
        },
        "Qayta kirish": {
            ru: "Войти снова",
            en: "Login again",
        },
        "Ro'yxatdan o'tish": {
            ru: "Регистрация",
            en: "Register",
        },
        "Ro‘yxatdan o‘tish": {
            ru: "Регистрация",
            en: "Register",
        },
        "Dashboard": {
            ru: "Панель",
            en: "Dashboard",
        },
        "Rejimni almashtirish": {
            ru: "Сменить режим",
            en: "Toggle theme",
        },
        "Kunduzgi/tungi rejim": {
            ru: "Светлая/темная тема",
            en: "Light/dark mode",
        },
        "Boshqaruv paneli": {
            ru: "Панель управления",
            en: "Control panel",
        },
        "Amaliyot hujjatlarini bitta joydan boshqaring": {
            ru: "Управляйте документами практики в одном месте",
            en: "Manage internship documents in one place",
        },
        "Talabalarni yuklang, hujjatlarni tayyorlang va profilingiz orqali barcha bo‘limlarga o‘ting.": {
            ru: "Загружайте студентов, готовьте документы и переходите ко всем разделам через профиль.",
            en: "Upload students, prepare documents, and access all sections from your profile.",
        },
        "Profilni ko‘rish": {
            ru: "Открыть профиль",
            en: "View profile",
        },
        "Talabalar": {
            ru: "Студенты",
            en: "Students",
        },
        "Korxonalar": {
            ru: "Организации",
            en: "Organizations",
        },
        "Hujjatlar": {
            ru: "Документы",
            en: "Documents",
        },
        "Asosiy bo‘limlar": {
            ru: "Основные разделы",
            en: "Main sections",
        },
        "Asosiy bo'limlar": {
            ru: "Основные разделы",
            en: "Main sections",
        },
        "Excel yoki Word yuklash": {
            ru: "Загрузить Excel или Word",
            en: "Upload Excel or Word",
        },
        "ZIP yuklab olish": {
            ru: "Скачать ZIP",
            en: "Download ZIP",
        },
        "Yuklangan talabalar": {
            ru: "Загруженные студенты",
            en: "Loaded students",
        },
        "Hozircha talabalar ma'lumotlari yuklanmagan.": {
            ru: "Данные студентов пока не загружены.",
            en: "No student data has been uploaded yet.",
        },
        "Oldingi": {
            ru: "Назад",
            en: "Previous",
        },
        "Keyingi": {
            ru: "Далее",
            en: "Next",
        },
        "Import markazi": {
            ru: "Центр импорта",
            en: "Import center",
        },
        "Talabalar ma'lumotlarini yuklash": {
            ru: "Загрузка данных студентов",
            en: "Upload student data",
        },
        "Excel yoki Word fayldagi ro'yxatni tizimga yuklang. Excel namuna ichida kurs, boshlanish sanasi va tugash sanasi ham bo'ladi, keyin hujjatlar yaratish va ZIP eksport qilish mumkin bo'ladi.": {
            ru: "Загрузите список из Excel или Word в систему. В шаблоне Excel также будут курс, дата начала и дата окончания, после чего можно создавать документы и экспортировать ZIP.",
            en: "Upload a list from an Excel or Word file. The Excel sample also includes course, start date, and end date, then documents can be generated and exported as ZIP.",
        },
        "Foydalanuvchi": {
            ru: "Пользователь",
            en: "User",
        },
        "Formatlar": {
            ru: "Форматы",
            en: "Formats",
        },
        "Yuklash bo'yicha ko'rsatma": {
            ru: "Инструкция по загрузке",
            en: "Upload guide",
        },
        "1. Talabalar faylini tanlang": {
            ru: "1. Выберите файл студентов",
            en: "1. Select the student file",
        },
        "2. Excel ichidagi barcha ustunlarni to'ldiring": {
            ru: "2. Заполните все столбцы в Excel",
            en: "2. Fill in all columns in Excel",
        },
        "3. Ma'lumotlarni tizimga yuboring": {
            ru: "3. Отправьте данные в систему",
            en: "3. Send the data to the system",
        },
        "4. Keyingi bosqichda hujjatlarni yuklab oling": {
            ru: "4. Download documents in the next step",
            en: "4. Download the documents in the next step",
        },
        "Namuna Excel yuklab olish": {
            ru: "Скачать пример Excel",
            en: "Download sample Excel",
        },
        "shablon manbasi": {
            ru: "источник шаблона",
            en: "template source",
        },
        "Fayl hozircha topilmadi": {
            ru: "Файл пока не найден",
            en: "File not found yet",
        },
        "Amaliyot hujjatlari importi": {
            ru: "Импорт документов практики",
            en: "Internship document import",
        },
        "Talabalar fayli (.xlsx yoki .docx)": {
            ru: "Файл студентов (.xlsx или .docx)",
            en: "Student file (.xlsx or .docx)",
        },
        "XLSX yuklasangiz kurs va sanalar Excel ichidan olinadi. DOCX yuklashda esa oxirgi saqlangan qiymatlar ishlatiladi.": {
            ru: "При загрузке XLSX курс и даты берутся из Excel. При загрузке DOCX используются последние сохраненные значения.",
            en: "When uploading XLSX, the course and dates are taken from Excel. When uploading DOCX, the last saved values are used.",
        },
        "Ma'lumotlarni yuklash": {
            ru: "Загрузить данные",
            en: "Upload data",
        },
        "Ma'lumotlar muvaffaqiyatli yuklandi. ZIP fayl avtomatik yuklanadi.": {
            ru: "Данные успешно загружены. ZIP-файл будет скачан автоматически.",
            en: "Data uploaded successfully. The ZIP file will download automatically.",
        },
        "Hujjatlarni ZIP holatda yuklab olish": {
            ru: "Скачать документы в ZIP",
            en: "Download documents as ZIP",
        },
        "Profil markazi": {
            ru: "Центр профиля",
            en: "Profile center",
        },
        "Foydalanuvchi profili": {
            ru: "Профиль пользователя",
            en: "User profile",
        },
        "Profilingiz orqali akkaunt ma'lumotlari va tezkor sahifa o'tishlar bir joyda jamlangan.": {
            ru: "В профиле собраны данные аккаунта и быстрые переходы по страницам.",
            en: "Your profile keeps account details and quick page links in one place.",
        },
        "Tahrirlash": {
            ru: "Редактировать",
            en: "Edit",
        },
        "Status": {
            ru: "Статус",
            en: "Status",
        },
        "Faol": {
            ru: "Активен",
            en: "Active",
        },
        "Asosiy ma'lumotlar": {
            ru: "Основные данные",
            en: "Main information",
        },
        "Foydalanuvchi nomi:": {
            ru: "Имя пользователя:",
            en: "Username:",
        },
        "Ism:": {
            ru: "Имя:",
            en: "First name:",
        },
        "Familiya:": {
            ru: "Фамилия:",
            en: "Last name:",
        },
        "Email:": {
            ru: "Email:",
            en: "Email:",
        },
        "Yo'q": {
            ru: "Нет",
            en: "No",
        },
        "Kiritilmagan": {
            ru: "Не указано",
            en: "Not provided",
        },
        "Tezkor amallar": {
            ru: "Быстрые действия",
            en: "Quick actions",
        },
        "Profilni tahrirlash": {
            ru: "Редактировать профиль",
            en: "Edit profile",
        },
        "Talabalar faylini yuklash": {
            ru: "Загрузить файл студентов",
            en: "Upload student file",
        },
        "Profil ma'lumotlari va parolni shu sahifadan yangilashingiz mumkin.": {
            ru: "На этой странице можно обновить данные профиля и пароль.",
            en: "You can update profile information and password on this page.",
        },
        "Panel": {
            ru: "Панель",
            en: "Panel",
        },
        "Foydalanuvchi nomi": {
            ru: "Имя пользователя",
            en: "Username",
        },
        "Ism": {
            ru: "Имя",
            en: "First name",
        },
        "Familiya": {
            ru: "Фамилия",
            en: "Last name",
        },
        "Ma'lumotlarni saqlash": {
            ru: "Сохранить данные",
            en: "Save information",
        },
        "Parolni yangilash": {
            ru: "Обновить пароль",
            en: "Update password",
        },
        "Joriy parol": {
            ru: "Текущий пароль",
            en: "Current password",
        },
        "Hozirgi parolingiz": {
            ru: "Ваш текущий пароль",
            en: "Your current password",
        },
        "Yangi parol": {
            ru: "Новый пароль",
            en: "New password",
        },
        "Kamida 8 ta belgi": {
            ru: "Минимум 8 символов",
            en: "At least 8 characters",
        },
        "Yangi parolni tasdiqlang": {
            ru: "Подтвердите новый пароль",
            en: "Confirm new password",
        },
        "Parolni qayta kiriting": {
            ru: "Введите пароль еще раз",
            en: "Enter the password again",
        },
        "Parolni saqlash": {
            ru: "Сохранить пароль",
            en: "Save password",
        },
        "Tezkor bo'limlar": {
            ru: "Быстрые разделы",
            en: "Quick sections",
        },
        "Profil sahifasi": {
            ru: "Страница профиля",
            en: "Profile page",
        },
        "Universitet tizimi": {
            ru: "Университетская система",
            en: "University system",
        },
        "Tizimga kirish": {
            ru: "Вход в систему",
            en: "Sign in",
        },
        "Akkauntingiz orqali boshqaruv paneli, yuklash bo‘limi, profil va hujjatlar sahifalariga o‘ting.": {
            ru: "Через аккаунт перейдите к панели управления, загрузке, профилю и страницам документов.",
            en: "Use your account to access the control panel, upload section, profile, and document pages.",
        },
        "Talabalar amaliyoti hujjatlarini boshqarish uchun yagona ishchi muhit.": {
            ru: "Единая рабочая среда для управления документами практики студентов.",
            en: "A single workspace for managing student internship documents.",
        },
        "Foydalanuvchi nomi va parol orqali tizimga kiring.": {
            ru: "Войдите в систему с помощью имени пользователя и пароля.",
            en: "Sign in with your username and password.",
        },
        "Masalan: dekanat_admin": {
            ru: "Например: dekanat_admin",
            en: "Example: dekanat_admin",
        },
        "Parol": {
            ru: "Пароль",
            en: "Password",
        },
        "Parolingizni kiriting": {
            ru: "Введите пароль",
            en: "Enter your password",
        },
        "Meni eslab qolish": {
            ru: "Запомнить меня",
            en: "Remember me",
        },
        "Akkauntingiz yo‘qmi?": {
            ru: "Нет аккаунта?",
            en: "Do not have an account?",
        },
        "Akkauntingiz yo'qmi?": {
            ru: "Нет аккаунта?",
            en: "Do not have an account?",
        },
        "Yangi foydalanuvchi": {
            ru: "Новый пользователь",
            en: "New user",
        },
        "Universitet muhitiga mos boshqaruv panelidan foydalanish uchun akkaunt yarating.": {
            ru: "Создайте аккаунт, чтобы пользоваться панелью управления для университетской среды.",
            en: "Create an account to use the management panel for a university environment.",
        },
        "Profil sahifasi avtomatik yaratiladi": {
            ru: "Страница профиля создается автоматически",
            en: "The profile page is created automatically",
        },
        "Dashboard orqali barcha bo‘limlar ochiladi": {
            ru: "Все разделы открываются через панель",
            en: "All sections are opened through the dashboard",
        },
        "Keyinchalik yuklash va hujjatlar bo‘limiga o‘tasiz": {
            ru: "Позже вы перейдете к разделам загрузки и документов",
            en: "Later you can go to the upload and documents sections",
        },
        "Yangi akkaunt yaratish": {
            ru: "Создать новый аккаунт",
            en: "Create a new account",
        },
        "Hisob ochish uchun quyidagi ma'lumotlarni kiriting.": {
            ru: "Введите следующие данные, чтобы создать аккаунт.",
            en: "Enter the following information to create an account.",
        },
        "F.I.Sh.": {
            ru: "Ф.И.О.",
            en: "Full name",
        },
        "Masalan: Asadbek Karimov": {
            ru: "Например: Асадбек Каримов",
            en: "Example: Asadbek Karimov",
        },
        "Masalan: dekanat_operator": {
            ru: "Например: dekanat_operator",
            en: "Example: dekanat_operator",
        },
        "Masalan: operator@university.uz": {
            ru: "Например: operator@university.uz",
            en: "Example: operator@university.uz",
        },
        "Parolda harf va raqam ishlatish tavsiya etiladi.": {
            ru: "Рекомендуется использовать буквы и цифры в пароле.",
            en: "Using letters and numbers in the password is recommended.",
        },
        "Parolni tasdiqlang": {
            ru: "Подтвердите пароль",
            en: "Confirm password",
        },
        "Ro‘yxatdan o‘tishni yakunlash": {
            ru: "Завершить регистрацию",
            en: "Complete registration",
        },
        "Akkauntingiz bormi?": {
            ru: "Уже есть аккаунт?",
            en: "Already have an account?",
        },
        "Amaliyot hujjatlarini tez, tartibli va qulay boshqaring": {
            ru: "Управляйте документами практики быстро, удобно и организованно",
            en: "Manage internship documents quickly, neatly, and conveniently",
        },
        "Tizim qanday ishlaydi": {
            ru: "Как работает система",
            en: "How the system works",
        },
        "Talabalarni yuklash": {
            ru: "Загрузка студентов",
            en: "Uploading students",
        },
        "Hujjatlar yaratish": {
            ru: "Создание документов",
            en: "Creating documents",
        },
        "Kimlar uchun mo‘ljallangan": {
            ru: "Для кого предназначено",
            en: "Who it is for",
        },
        "Platforma afzalliklari": {
            ru: "Преимущества платформы",
            en: "Platform advantages",
        },
        "Tez import": {
            ru: "Быстрый импорт",
            en: "Fast import",
        },
        "Zamonaviy dizayn": {
            ru: "Современный дизайн",
            en: "Modern design",
        },
        "API tayyor": {
            ru: "API готов",
            en: "API ready",
        },
        "Admin panel": {
            ru: "Админ-панель",
            en: "Admin panel",
        },
        "Sahifa topilmadi": {
            ru: "Страница не найдена",
            en: "Page not found",
        },
        "Ortga qaytish": {
            ru: "Назад",
            en: "Go back",
        },
        "Serverda xatolik yuz berdi": {
            ru: "На сервере произошла ошибка",
            en: "A server error occurred",
        },
        "Qayta yuklash": {
            ru: "Перезагрузить",
            en: "Reload",
        },
        "Kirish mumkin emas": {
            ru: "Доступ запрещен",
            en: "Access denied",
        },
        "Sahifani qayta yuklash": {
            ru: "Перезагрузить страницу",
            en: "Reload page",
        },
    };

    const reverseLookup = new Map();
    for (const [uz, values] of Object.entries(translations)) {
        reverseLookup.set(normalizeText(uz), uz);
        for (const lang of ["ru", "en"]) {
            reverseLookup.set(normalizeText(values[lang]), uz);
        }
    }

    function normalizeText(value) {
        return String(value || "").replace(/\s+/g, " ").trim();
    }

    function getCurrentLanguage() {
        const savedLanguage = localStorage.getItem("appLanguage") || "uz";
        return supportedLanguages.includes(savedLanguage) ? savedLanguage : "uz";
    }

    function splitOuterWhitespace(value) {
        const start = value.match(/^\s*/)[0];
        const end = value.match(/\s*$/)[0];
        return {
            start,
            end,
            core: value.slice(start.length, value.length - end.length),
        };
    }

    function translateCore(core, language) {
        const normalized = normalizeText(core);
        if (!normalized) {
            return core;
        }

        const countMatch = normalized.match(/^(.*?)(\s*\(\d+\))$/);
        if (countMatch) {
            const prefixKey = reverseLookup.get(normalizeText(countMatch[1]));
            if (prefixKey && translations[prefixKey]) {
                return (language === "uz" ? prefixKey : translations[prefixKey][language]) + countMatch[2];
            }
        }

        const key = reverseLookup.get(normalized);
        if (!key || !translations[key]) {
            return core;
        }

        return language === "uz" ? key : translations[key][language];
    }

    function translateValue(value, language) {
        const parts = splitOuterWhitespace(value);
        return parts.start + translateCore(parts.core, language) + parts.end;
    }

    function translateTextNodes(language) {
        const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
        const ignoredParents = new Set(["SCRIPT", "STYLE", "TEXTAREA", "CODE", "PRE"]);
        const nodes = [];

        while (walker.nextNode()) {
            const node = walker.currentNode;
            if (node.parentElement && ignoredParents.has(node.parentElement.tagName)) {
                continue;
            }
            if (!normalizeText(node.nodeValue)) {
                continue;
            }
            nodes.push(node);
        }

        for (const node of nodes) {
            node.nodeValue = translateValue(node.nodeValue, language);
        }
    }

    function translateAttributes(language) {
        const attributeNames = ["placeholder", "title", "aria-label", "value"];
        for (const element of document.querySelectorAll("[placeholder], [title], [aria-label], input[type='submit'][value]")) {
            for (const attributeName of attributeNames) {
                if (!element.hasAttribute(attributeName)) {
                    continue;
                }
                if (attributeName === "value" && element.tagName !== "INPUT") {
                    continue;
                }
                element.setAttribute(attributeName, translateValue(element.getAttribute(attributeName), language));
            }
        }
    }

    function updateLanguageButtons(language) {
        for (const button of document.querySelectorAll("[data-lang]")) {
            const active = button.dataset.lang === language;
            button.classList.toggle("active", active);
            button.setAttribute("aria-pressed", String(active));
        }
    }

    function ensureLanguageStyles() {
        if (document.getElementById("app-language-styles")) {
            return;
        }

        const style = document.createElement("style");
        style.id = "app-language-styles";
        style.textContent = `
            .language-switcher {
                display: inline-flex;
                align-items: center;
                gap: 4px;
                padding: 4px;
                border: 1px solid var(--line, rgba(148, 163, 184, 0.24));
                border-radius: 999px;
                background: var(--surface, rgba(255, 255, 255, 0.8));
            }
            .lang-option {
                min-width: 38px;
                min-height: 34px;
                border: 0;
                border-radius: 999px;
                padding: 7px 10px;
                color: var(--text, #10233b);
                background: transparent;
                box-shadow: none;
                font: inherit;
                font-weight: 800;
                cursor: pointer;
            }
            .lang-option:hover,
            .lang-option.active {
                color: #fff;
                background: linear-gradient(135deg, #2563eb, #7c3aed);
            }
            .auth-language-switcher {
                position: fixed;
                top: 22px;
                right: 82px;
                z-index: 10;
            }
        `;
        document.head.appendChild(style);
    }

    function setAppLanguage(language) {
        const nextLanguage = supportedLanguages.includes(language) ? language : "uz";
        localStorage.setItem("appLanguage", nextLanguage);
        document.documentElement.lang = nextLanguage;
        document.title = translateValue(document.title, nextLanguage);
        translateTextNodes(nextLanguage);
        translateAttributes(nextLanguage);
        updateLanguageButtons(nextLanguage);
    }

    window.setAppLanguage = setAppLanguage;

    window.addEventListener("DOMContentLoaded", () => {
        ensureLanguageStyles();
        const language = getCurrentLanguage();
        setAppLanguage(language);

        document.addEventListener("click", (event) => {
            const button = event.target.closest("[data-lang]");
            if (!button) {
                return;
            }
            setAppLanguage(button.dataset.lang);
        });
    });
})();
