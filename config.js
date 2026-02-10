// Автоматическая конфигурация проектов
const PROJECTS_CONFIG = {
    flask: {
        name: "Flask API",
        path: "/flask",
        port: 5000,
        healthEndpoint: "/health",
        icon: "🐍"
    },
    react: {
        name: "React App", 
        path: "/react",
        port: 3000,
        healthEndpoint: "/",
        icon: "⚛️"
    }
};

// Функция для динамического обновления статуса
async function updateProjectStatus() {
    for (const [key, config] of Object.entries(PROJECTS_CONFIG)) {
        try {
            await fetch(`${config.path}${config.healthEndpoint}`);
            document.querySelector(`#${key}-status`).textContent = '● Онлайн';
            document.querySelector(`#${key}-status`).style.color = '#10b981';
        } catch {
            document.querySelector(`#${key}-status`).textContent = '● Офлайн';
            document.querySelector(`#${key}-status`).style.color = '#ef4444';
        }
    }
}