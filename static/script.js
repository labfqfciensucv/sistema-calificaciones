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
});
