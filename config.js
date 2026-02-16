// Автоматическая конфигурация проектов
const PROJECTS_CONFIG = {
    flask: {
        name: "Flask API",
        path: "http://localhost:5000",      // Прямая ссылка на порт
        port: 5000,
        healthEndpoint: "/health",
        icon: "🐍"
    },
    react: {
        name: "React App", 
        path: "http://localhost:3002",      // Прямая ссылка на порт (мы сменили 3000 на 3002)
        port: 3002,
        healthEndpoint: "/",
        icon: "⚛️"
    }
};

// Функция для динамического обновления статуса (если используется в index.html)
async function updateProjectStatus() {
    for (const [key, config] of Object.entries(PROJECTS_CONFIG)) {
        try {
            // Используем mode: 'no-cors' для проверки доступности с другого порта
            await fetch(`${config.path}${config.healthEndpoint}`, { 
                method: 'HEAD',
                mode: 'no-cors',
                cache: 'no-store'
            });
            
            const badge = document.querySelector(`#${key}-status`);
            if(badge) {
                badge.textContent = '● Онлайн';
                badge.style.color = '#10b981'; // Зеленый
                badge.style.background = '#dcfce7';
            }
        } catch {
            const badge = document.querySelector(`#${key}-status`);
            if(badge) {
                badge.textContent = '● Офлайн';
                badge.style.color = '#ef4444'; // Красный
                badge.style.background = '#fee2e2';
            }
        }
    }
}