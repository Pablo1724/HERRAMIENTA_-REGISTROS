document.addEventListener('DOMContentLoaded', () => {
    const searchInput = document.getElementById('career-search');
    
    // Si el input no existe en esta página, no ejecutamos nada
    if (!searchInput) return;

    searchInput.addEventListener('input', (e) => {
        const term = e.target.value.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        const cards = document.querySelectorAll('.career-card-v3');
        let visibles = 0;

        cards.forEach(card => {
            // Buscamos en el título y en el tag de la facultad
            const title = card.querySelector('h3').textContent.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
            const tag = card.querySelector('.card-tag').textContent.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
            
            if (title.includes(term) || tag.includes(term)) {
                card.style.display = "flex"; // Volver a mostrar
                card.style.animation = "fadeIn 0.3s ease";
                visibles++;
            } else {
                card.style.display = "none"; // Ocultar
            }
        });

        // Manejo de estado vacío (No hay resultados)
        toggleEmptyState(visibles);
    });
});

function toggleEmptyState(count) {
    let msg = document.getElementById('no-results');
    const container = document.querySelector('.career-grid-professional');

    if (count === 0) {
        if (!msg) {
            msg = document.createElement('div');
            msg.id = 'no-results';
            msg.innerHTML = `
                <div class="empty-state-box">
                    <span class="empty-icon">📂</span>
                    <h3>No se encontraron carreras</h3>
                    <p>Intenta con otros términos de búsqueda.</p>
                </div>`;
            container.after(msg);
        }
    } else if (msg) {
        msg.remove();
    }
}