document.addEventListener('DOMContentLoaded', function() {
    // Обработка всех форм
    const forms = document.querySelectorAll('form');
    forms.forEach(form => {
        form.addEventListener('submit', async function(e) {
            e.preventDefault();

            const formData = new FormData(this);
            const submitBtn = this.querySelector('button[type="submit"]');
            const originalText = submitBtn.textContent;

            submitBtn.disabled = true;
            submitBtn.textContent = 'Обработка...';

            try {
                const response = await fetch(form.action || window.location.pathname, {
                    method: 'POST',
                    body: formData
                });

                const result = await response.json();

                if (result.success) {
                    alert('Операция выполнена успешно!');
                    if (result.result_dir) {
                        alert(`Результаты сохранены в: ${result.result_dir}`);
                    }
                } else {
                    alert(`Ошибка: ${result.error}`);
                }
            } catch (error) {
                alert('Ошибка сети или сервера');
            } finally {
                submitBtn.disabled = false;
                submitBtn.textContent = originalText;
            }
        });
    });

    // Переключение синхронизации Google
    const googleToggle = document.getElementById('googleSyncToggle');
    if (googleToggle) {
        googleToggle.addEventListener('change', async function() {
            const response = await fetch('/toggle_google_sync', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ enabled: this.checked })
            });

            const result = await response.json();
            if (!result.success) {
                alert('Ошибка при изменении настроек синхронизации');
            }
        });
    }
});