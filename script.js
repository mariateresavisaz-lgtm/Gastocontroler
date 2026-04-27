// Variables globales
let movimientos = JSON.parse(localStorage.getItem("movimientos")) || [];
let movimientosFiltrados = [];
let chartGastos = null;
let chartIngresos = null;

// Inicializar
document.addEventListener("DOMContentLoaded", function() {
  mostrarMovimientos();
  inicializarTema();
  // Agregar event listeners a los filtros
  document.getElementById("filtroFechaInicio").addEventListener("change", aplicarFiltros);
  document.getElementById("filtroFechaFin").addEventListener("change", aplicarFiltros);
  document.getElementById("filtroCategoria").addEventListener("change", aplicarFiltros);
  document.getElementById("filtroTipo").addEventListener("change", aplicarFiltros);
});

// ============================================
// TEMA OSCURO/CLARO
// ============================================

function toggleTema() {
  document.body.classList.toggle("dark-mode");
  const temaDark = document.body.classList.contains("dark-mode");
  localStorage.setItem("tema", temaDark ? "dark" : "light");
  
  const icono = document.getElementById("iconoTema");
  icono.textContent = temaDark ? "☀️" : "🌙";
  
  // Actualizar gráficos
  if (chartGastos) chartGastos.destroy();
  if (chartIngresos) chartIngresos.destroy();
  mostrarMovimientos();
}

function inicializarTema() {
  const temaSaved = localStorage.getItem("tema");
  const temaOscuro = temaSaved === "dark";
  
  if (temaOscuro) {
    document.body.classList.add("dark-mode");
    document.getElementById("iconoTema").textContent = "☀️";
  }
}

// ============================================
// GUARDAR Y CARGAR DATOS
// ============================================

function guardarDatos() {
  localStorage.setItem("movimientos", JSON.stringify(movimientos));
}

// ============================================
// AGREGAR MOVIMIENTO MANUAL
// ============================================

function agregarMovimiento() {
  const fecha = document.getElementById("fecha").value;
  const tipo = document.getElementById("tipo").value;
  const categoria = document.getElementById("categoria").value;
  const descripcion = document.getElementById("descripcion").value;
  const importe = parseFloat(document.getElementById("importe").value);

  if (!fecha || !descripcion || isNaN(importe) || importe <= 0) {
    alert("❌ Por favor, completa todos los campos correctamente.");
    return;
  }

  const movimiento = {
    id: Date.now() + Math.random(),
    fecha,
    tipo,
    categoria,
    descripcion,
    importe
  };

  movimientos.push(movimiento);
  guardarDatos();
  mostrarMovimientos();

  // Limpiar formulario
  document.getElementById("fecha").value = "";
  document.getElementById("descripcion").value = "";
  document.getElementById("importe").value = "";
  
  alert("✅ Movimiento añadido correctamente");
}

// ============================================
// IMPORTAR EXCEL
// ============================================

function importarExcel() {
  const archivo = document.getElementById("archivoExcel").files[0];

  if (!archivo) {
    alert("❌ Selecciona un archivo Excel primero.");
    return;
  }

  const lector = new FileReader();

  lector.onload = function(e) {
    try {
      const datos = new Uint8Array(e.target.result);
      const libro = XLSX.read(datos, { type: "array" });
      const hoja = libro.Sheets[libro.SheetNames[0]];
      const filas = XLSX.utils.sheet_to_json(hoja);

      let importados = 0;

      filas.forEach(fila => {
        const fecha = fila.Fecha || fila.fecha || fila.FECHA || fila["Fecha operación"] || fila["FECHA OPERACIÓN"];
        const descripcion = fila.Descripcion || fila.Descripción || fila.descripcion || fila.Concepto || fila.concepto || fila.CONCEPTO || fila.Detalle || fila.detalle;
        const importeBruto = fila.Importe || fila.importe || fila.IMPORTE || fila.Cantidad || fila.cantidad || fila.CANTIDAD || fila.Euros || fila.euros;

        const importe = limpiarImporte(importeBruto);

        if (fecha && descripcion && !isNaN(importe)) {
          const tipo = importe >= 0 ? "ingreso" : "gasto";

          const movimiento = {
            id: Date.now() + Math.random(),
            fecha: convertirFecha(fecha),
            tipo: tipo,
            categoria: detectarCategoria(descripcion),
            descripcion: descripcion,
            importe: Math.abs(importe)
          };

          movimientos.push(movimiento);
          importados++;
        }
      });

      guardarDatos();
      mostrarMovimientos();
      document.getElementById("archivoExcel").value = "";

      alert(`✅ Se han importado ${importados} movimientos correctamente.`);
    } catch (error) {
      alert("❌ Error al leer el archivo. Asegúrate de que es un Excel válido.");
      console.error(error);
    }
  };

  lector.readAsArrayBuffer(archivo);
}

function limpiarImporte(valor) {
  if (typeof valor === "number") {
    return valor;
  }

  if (!valor) {
    return NaN;
  }

  valor = valor.toString();
  valor = valor.replace("€", "");
  valor = valor.replace(/\s/g, "");
  valor = valor.replace(/\./g, "");
  valor = valor.replace(",", ".");

  return parseFloat(valor);
}

function convertirFecha(fecha) {
  if (typeof fecha === "number") {
    const fechaExcel = XLSX.SSF.parse_date_code(fecha);
    return `${fechaExcel.y}-${String(fechaExcel.m).padStart(2, "0")}-${String(fechaExcel.d).padStart(2, "0")}`;
  }

  if (typeof fecha === "string") {
    const partes = fecha.split("/");

    if (partes.length === 3) {
      return `${partes[2]}-${partes[1].padStart(2, "0")}-${partes[0].padStart(2, "0")}`;
    }
  }

  return fecha;
}

function detectarCategoria(texto) {
  texto = texto.toLowerCase();

  if (
    texto.includes("mercadona") ||
    texto.includes("carrefour") ||
    texto.includes("lidl") ||
    texto.includes("aldi") ||
    texto.includes("supermercado") ||
    texto.includes("dia") ||
    texto.includes("alcampo")
  ) {
    return "Alimentación";
  }

  if (
    texto.includes("renfe") ||
    texto.includes("metro") ||
    texto.includes("gasolina") ||
    texto.includes("repsol") ||
    texto.includes("cepsa") ||
    texto.includes("uber") ||
    texto.includes("cabify") ||
    texto.includes("autobús") ||
    texto.includes("transporte")
  ) {
    return "Transporte";
  }

  if (
    texto.includes("netflix") ||
    texto.includes("spotify") ||
    texto.includes("amazon prime") ||
    texto.includes("hbo") ||
    texto.includes("disney") ||
    texto.includes("suscripción")
  ) {
    return "Suscripciones";
  }

  if (
    texto.includes("farmacia") ||
    texto.includes("hospital") ||
    texto.includes("clínica") ||
    texto.includes("clinica") ||
    texto.includes("dentista") ||
    texto.includes("médico")
  ) {
    return "Salud";
  }

  if (
    texto.includes("zara") ||
    texto.includes("amazon") ||
    texto.includes("shein") ||
    texto.includes("el corte ingles") ||
    texto.includes("primark") ||
    texto.includes("h&m") ||
    texto.includes("compra")
  ) {
    return "Compras";
  }

  if (
    texto.includes("restaurante") ||
    texto.includes("bar") ||
    texto.includes("cine") ||
    texto.includes("glovo") ||
    texto.includes("just eat") ||
    texto.includes("telepizza") ||
    texto.includes("ocio") ||
    texto.includes("entretenimiento")
  ) {
    return "Ocio";
  }

  if (
    texto.includes("alquiler") ||
    texto.includes("hipoteca") ||
    texto.includes("luz") ||
    texto.includes("agua") ||
    texto.includes("gas") ||
    texto.includes("iberdrola") ||
    texto.includes("endesa") ||
    texto.includes("vivienda")
  ) {
    return "Vivienda";
  }

  if (
    texto.includes("nómina") ||
    texto.includes("nomina") ||
    texto.includes("salario") ||
    texto.includes("transferencia recibida") ||
    texto.includes("sueldo")
  ) {
    return "Nómina";
  }

  return "Otros";
}

// ============================================
// FILTROS
// ============================================

function aplicarFiltros() {
  const fechaInicio = document.getElementById("filtroFechaInicio").value;
  const fechaFin = document.getElementById("filtroFechaFin").value;
  const categoria = document.getElementById("filtroCategoria").value;
  const tipo = document.getElementById("filtroTipo").value;

  movimientosFiltrados = movimientos.filter(mov => {
    let cumple = true;

    if (fechaInicio && mov.fecha < fechaInicio) cumple = false;
    if (fechaFin && mov.fecha > fechaFin) cumple = false;
    if (categoria && mov.categoria !== categoria) cumple = false;
    if (tipo && mov.tipo !== tipo) cumple = false;

    return cumple;
  });

  mostrarMovimientos();
}

function limpiarFiltros() {
  document.getElementById("filtroFechaInicio").value = "";
  document.getElementById("filtroFechaFin").value = "";
  document.getElementById("filtroCategoria").value = "";
  document.getElementById("filtroTipo").value = "";
  
  movimientosFiltrados = [];
  mostrarMovimientos();
}

// ============================================
// ELIMINAR Y BORRAR
// ============================================

function eliminarMovimiento(id) {
  const confirmar = confirm("¿Estás seguro de que quieres eliminar este movimiento?");
  if (confirmar) {
    movimientos = movimientos.filter(mov => mov.id !== id);
    guardarDatos();
    mostrarMovimientos();
  }
}

function borrarTodo() {
  const confirmar = confirm("⚠️ ¿Seguro que quieres borrar TODOS los datos? Esta acción no se puede deshacer.");
  if (confirmar) {
    movimientos = [];
    guardarDatos();
    mostrarMovimientos();
    alert("✅ Todos los datos han sido eliminados.");
  }
}

// ============================================
// EXPORTAR EXCEL
// ============================================

function exportarExcel() {
  if (movimientos.length === 0) {
    alert("❌ No hay movimientos para exportar.");
    return;
  }

  const datos = movimientos.map(mov => ({
    Fecha: mov.fecha,
    Tipo: mov.tipo,
    Categoría: mov.categoria,
    Descripción: mov.descripcion,
    Importe: mov.importe
  }));

  const hoja = XLSX.utils.json_to_sheet(datos);
  const libro = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(libro, hoja, "Movimientos");

  const fecha = new Date().toISOString().split("T")[0];
  XLSX.writeFile(libro, `Gastocontroler_${fecha}.xlsx`);
  
  alert("✅ Archivo descargado correctamente.");
}

// ============================================
// MOSTRAR MOVIMIENTOS Y ACTUALIZAR
// ============================================

function mostrarMovimientos() {
  const tabla = document.getElementById("tablaMovimientos");
  tabla.innerHTML = "";

  // Usar movimientos filtrados si hay filtros, sino mostrar todos
  const movimientosAMostrar = movimientosFiltrados.length > 0 ? movimientosFiltrados : movimientos;

  let totalIngresos = 0;
  let totalGastos = 0;
  let categorias = {};

  movimientosAMostrar.forEach(movimiento => {
    if (movimiento.tipo === "ingreso") {
      totalIngresos += movimiento.importe;
    } else {
      totalGastos += movimiento.importe;

      if (!categorias[movimiento.categoria]) {
        categorias[movimiento.categoria] = 0;
      }

      categorias[movimiento.categoria] += movimiento.importe;
    }

    tabla.innerHTML += `
      <tr>
        <td>${movimiento.fecha}</td>
        <td class="${movimiento.tipo}">${movimiento.tipo === "ingreso" ? "💰 Ingreso" : "📉 Gasto"}</td>
        <td>${movimiento.categoria}</td>
        <td>${movimiento.descripcion}</td>
        <td>${movimiento.importe.toFixed(2)} €</td>
        <td>
          <button class="btn-eliminar" onclick="eliminarMovimiento(${movimiento.id})">
            ❌
          </button>
        </td>
      </tr>
    `;
  });

  document.getElementById("totalIngresos").textContent = totalIngresos.toFixed(2) + " €";
  document.getElementById("totalGastos").textContent = totalGastos.toFixed(2) + " €";
  document.getElementById("saldo").textContent = (totalIngresos - totalGastos).toFixed(2) + " €";

  mostrarResumenCategorias(categorias);
  mostrarGraficos(categorias, totalIngresos, totalGastos);
}

function mostrarResumenCategorias(categorias) {
  const contenedor = document.getElementById("resumenCategorias");
  contenedor.innerHTML = "";

  const categoriasOrdenadas = Object.entries(categorias)
    .sort((a, b) => b[1] - a[1]);

  if (categoriasOrdenadas.length === 0) {
    contenedor.innerHTML = "<p>📊 No hay gastos registrados todavía.</p>";
    return;
  }

  categoriasOrdenadas.forEach(([categoria, monto]) => {
    const porcentaje = ((monto / Object.values(categorias).reduce((a, b) => a + b, 0)) * 100).toFixed(1);
    contenedor.innerHTML += `
      <div class="categoria-resumen">
        <strong>${categoria}:</strong>
        <span>${monto.toFixed(2)} € (${porcentaje}%)</span>
      </div>
    `;
  });
}

// ============================================
// GRÁFICOS
// ============================================

function mostrarGraficos(categorias, totalIngresos, totalGastos) {
  const coloresTema = document.body.classList.contains("dark-mode") 
    ? { text: "#e0e0e0", grid: "#444", bg: "#2d2d2d" }
    : { text: "#333", grid: "#ddd", bg: "#fff" };

  // Gráfico de gastos por categoría
  const ctxGastos = document.getElementById("graficoGastosCategorias");
  if (ctxGastos) {
    if (chartGastos) chartGastos.destroy();

    const colores = [
      "#FF6384", "#36A2EB", "#FFCE56", "#4BC0C0", "#9966FF",
      "#FF9F40", "#FF6384", "#C9CBCF", "#4BC0C0", "#FF6384"
    ];

    chartGastos = new Chart(ctxGastos, {
      type: "doughnut",
      data: {
        labels: Object.keys(categorias),
        datasets: [{
          data: Object.values(categorias),
          backgroundColor: colores.slice(0, Object.keys(categorias).length),
          borderColor: coloresTema.bg,
          borderWidth: 2
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: true,
        plugins: {
          legend: {
            labels: { color: coloresTema.text, font: { size: 12 } }
          }
        }
      }
    });
  }

  // Gráfico de ingresos vs gastos
  const ctxIngresos = document.getElementById("graficoIngresosGastos");
  if (ctxIngresos) {
    if (chartIngresos) chartIngresos.destroy();

    chartIngresos = new Chart(ctxIngresos, {
      type: "bar",
      data: {
        labels: ["Ingresos", "Gastos"],
        datasets: [{
          label: "Cantidad (€)",
          data: [totalIngresos, totalGastos],
          backgroundColor: ["#27ae60", "#e74c3c"],
          borderColor: ["#1e8449", "#c0392b"],
          borderWidth: 1
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: true,
        scales: {
          y: {
            ticks: { color: coloresTema.text },
            grid: { color: coloresTema.grid }
          },
          x: {
            ticks: { color: coloresTema.text },
            grid: { color: coloresTema.grid }
          }
        },
        plugins: {
          legend: {
            labels: { color: coloresTema.text }
          }
        }
      }
    });
  }
}
