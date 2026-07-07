// Funciones adicionales para el panel de administración
// (El JavaScript principal ya está incluido en admin.html)

// Función para exportar a CSV (opcional)
function exportarDatos() {
    fetch('/api/datos')
        .then(response => response.json())
        .then(data => {
            let csv = 'Estudiante,Bloque A,Bloque B,Nota Final\n';
            for (const [estudiante, nota] of Object.entries(data.final)) {
                // Usar coma como separador decimal en el CSV
                const bloqueA = data.bloqueA[estudiante].toFixed(2).replace('.', ',');
                const bloqueB = data.bloqueB[estudiante].toFixed(2).replace('.', ',');
                const final = nota.toFixed(2).replace('.', ',');
                csv += `${estudiante},${bloqueA},${bloqueB},${final}\n`;
            }
            
            const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'calificaciones.csv';
            a.click();
            window.URL.revokeObjectURL(url);
        });
}

// Función para imprimir el dashboard
function imprimirDashboard() {
    window.print();
}

// Detectar si hay cambios sin guardar (opcional)
let hayCambiosSinGuardar = false;

document.addEventListener('DOMContentLoaded', function() {
    console.log('📊 Sistema de Calificaciones - I-2026');
    console.log('✅ Formato decimal con coma activado');
    console.log('✅ Notas en base 20');
    
    // Agregar botón de exportar en el admin si existe
    const adminActions = document.querySelector('.admin-actions');
    if (adminActions) {
        const exportBtn = document.createElement('button');
        exportBtn.className = 'btn-reload';
        exportBtn.innerHTML = '<i class="fas fa-file-export"></i> Exportar CSV';
        exportBtn.onclick = exportarDatos;
        adminActions.appendChild(exportBtn);
    }
// Función mejorada de guardar con feedback visual
function guardarCambio(element) {
    const estudiante = element.dataset.estudiante;
    const campo = element.dataset.campo;
    const indice = element.dataset.indice;
    
    const valorTexto = element.value;
    const valor = parseCommaNumber(valorTexto);
    
    if (isNaN(valor) || valor < 0 || valor > 20) {
        element.classList.add('error');
        showToast('❌ Nota inválida (debe ser 0-20)', 'error');
        return;
    }
    
    // Mostrar indicador de carga
    element.classList.add('loading');
    showToast('💾 Guardando...', 'info');
    
    fetch('/api/actualizar', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            estudiante: estudiante,
            campo: campo,
            valor: valor,
            indice: indice ? parseInt(indice) : null
        })
    })
    .then(response => response.json())
    .then(data => {
        element.classList.remove('loading');
        
        if (data.success) {
            element.classList.add('success');
            showToast('✅ ¡Guardado exitosamente!', 'success');
            setTimeout(() => element.classList.remove('success'), 2000);
        } else {
            element.classList.add('error');
            showToast('❌ Error: ' + data.error, 'error');
        }
    })
    .catch(error => {
        element.classList.remove('loading');
        element.classList.add('error');
        showToast('❌ Error al guardar', 'error');
    });
}

// Sistema de notificaciones tipo toast
function showToast(message, type = 'info') {
    const colors = {
        success: '#4CAF50',
        error: '#f44336',
        info: '#2196F3',
        warning: '#FF9800'
    };
    
    const toast = document.createElement('div');
    toast.style.cssText = `
        position: fixed;
        bottom: 20px;
        right: 20px;
        background: ${colors[type]};
        color: white;
        padding: 16px 24px;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3);
        z-index: 9999;
        font-weight: 500;
        animation: slideIn 0.3s ease;
        max-width: 400px;
    `;
    toast.textContent = message;
    
    document.body.appendChild(toast);
    
    setTimeout(() => {
        toast.style.animation = 'slideOut 0.3s ease';
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}
});
