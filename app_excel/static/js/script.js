function applyTheme(theme) {
    document.body.classList.toggle("dark-mode", theme === "dark");
    localStorage.setItem("theme", theme);
}

function toggleDarkMode() {
    const nextTheme = document.body.classList.contains("dark-mode") ? "light" : "dark";
    applyTheme(nextTheme);
}

function initAiChat() {
    const widget = document.querySelector(".ai-chat-widget");
    if (!widget) {
        return;
    }

    const toggle = widget.querySelector(".ai-chat-toggle");
    const panel = widget.querySelector(".ai-chat-panel");
    const closeButton = widget.querySelector(".ai-chat-close");
    const form = widget.querySelector(".ai-chat-form");
    const input = widget.querySelector(".ai-chat-input");
    const sendButton = widget.querySelector(".ai-chat-send");
    const messages = widget.querySelector(".ai-chat-messages");
    const chatUrl = widget.dataset.chatUrl;
    const csrfToken = widget.dataset.chatCsrf;
    const history = [];

    function setOpen(isOpen) {
        widget.classList.toggle("open", isOpen);
        toggle.setAttribute("aria-expanded", String(isOpen));
        panel.setAttribute("aria-hidden", String(!isOpen));
        if (isOpen) {
            setTimeout(() => input.focus(), 50);
        }
    }

    function getPhoneHref(text) {
        const digits = text.replace(/\D/g, "");
        if (digits.length === 9) {
            return `tel:+998${digits}`;
        }
        if (digits.length >= 10) {
            return `tel:+${digits}`;
        }
        return "";
    }

    function getLinkForToken(token) {
        if (token.startsWith("http://") || token.startsWith("https://")) {
            return token;
        }
        if (token.startsWith("/")) {
            return token;
        }
        if (/^@[A-Za-z0-9_]{5,32}$/.test(token)) {
            return `https://t.me/${token.slice(1)}`;
        }
        return getPhoneHref(token);
    }

    function setMessageText(message, text) {
        const pattern = /(https?:\/\/[^\s)]+)|(\/[A-Za-z0-9_/-]+\/)|(@[A-Za-z0-9_]{5,32})|(\+?\d[\d\s-]{7,}\d)/g;
        let lastIndex = 0;
        message.replaceChildren();

        for (const match of text.matchAll(pattern)) {
            const token = match[0];
            const start = match.index;
            const href = getLinkForToken(token);

            if (start > lastIndex) {
                message.appendChild(document.createTextNode(text.slice(lastIndex, start)));
            }

            if (href) {
                const link = document.createElement("a");
                link.href = href;
                link.textContent = token;
                if (href.startsWith("http")) {
                    link.target = "_blank";
                    link.rel = "noopener noreferrer";
                }
                message.appendChild(link);
            } else {
                message.appendChild(document.createTextNode(token));
            }

            lastIndex = start + token.length;
        }

        if (lastIndex < text.length) {
            message.appendChild(document.createTextNode(text.slice(lastIndex)));
        }
    }

    function appendMessage(role, text, isError = false) {
        const message = document.createElement("div");
        message.className = `ai-chat-message ${role}${isError ? " error" : ""}`;
        setMessageText(message, text);
        messages.appendChild(message);
        messages.scrollTop = messages.scrollHeight;
        return message;
    }

    function remember(role, content) {
        history.push({ role, content });
        if (history.length > 10) {
            history.shift();
        }
    }

    async function sendMessage(text) {
        appendMessage("user", text);
        remember("user", text);
        const waiting = appendMessage("bot", "Javob tayyorlanmoqda...");
        sendButton.disabled = true;
        input.disabled = true;

        try {
            const response = await fetch(chatUrl, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "X-CSRFToken": csrfToken,
                },
                body: JSON.stringify({
                    message: text,
                    history: history.slice(0, -1),
                }),
            });
            const data = await response.json();
            const reply = data.reply || data.error || "Javob topilmadi.";
            setMessageText(waiting, reply);
            waiting.classList.toggle("error", !response.ok);
            remember("assistant", reply);
        } catch (error) {
            const reply = "Server bilan bog'lanishda muammo bo'ldi. Qayta urinib ko'ring.";
            setMessageText(waiting, reply);
            waiting.classList.add("error");
            remember("assistant", reply);
        } finally {
            sendButton.disabled = false;
            input.disabled = false;
            input.focus();
        }
    }

    toggle.addEventListener("click", () => setOpen(!widget.classList.contains("open")));
    closeButton.addEventListener("click", () => setOpen(false));

    form.addEventListener("submit", (event) => {
        event.preventDefault();
        const text = input.value.trim();
        if (!text) {
            return;
        }
        input.value = "";
        sendMessage(text);
    });

    input.addEventListener("keydown", (event) => {
        if (event.key === "Enter" && !event.shiftKey) {
            event.preventDefault();
            form.requestSubmit();
        }
    });
}

window.addEventListener("DOMContentLoaded", () => {
    const savedTheme = localStorage.getItem("theme") || "light";
    applyTheme(savedTheme);

    const themeToggle = document.getElementById("themeToggle");
    if (themeToggle) {
        themeToggle.addEventListener("click", toggleDarkMode);
    }

    const dropdown = document.getElementById("profileDropdown");
    const menuButton = document.getElementById("profileMenuButton");

    if (dropdown && menuButton) {
        menuButton.addEventListener("click", (event) => {
            event.stopPropagation();
            const isOpen = dropdown.classList.toggle("open");
            menuButton.setAttribute("aria-expanded", String(isOpen));
        });

        document.addEventListener("click", (event) => {
            if (!dropdown.contains(event.target)) {
                dropdown.classList.remove("open");
                menuButton.setAttribute("aria-expanded", "false");
            }
        });
    }

    initAiChat();
});
