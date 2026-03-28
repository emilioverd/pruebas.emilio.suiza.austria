# Acumulador automático de Excels (Python + Power Query M)

Este repositorio incluye **dos versiones del mismo enfoque** para acumular automáticamente múltiples archivos Excel:

1. `acumulador_excels.py` (Python)
2. `acumulador_power_query.m` (M para Power Query en Excel)

---

## 1) Opción Python

### Requisitos

```bash
pip install pandas openpyxl
```

### Ejemplo de uso

```bash
python acumulador_excels.py --origen "C:/Datos/Excels" --salida "C:/Datos/Salida/acumulado.xlsx"
```

### Parámetros útiles

- `--patron`: filtra nombres de archivo. Ej: `Ventas_*.xlsx`
- `--hoja`: nombre o índice de la hoja (por defecto `0`)

Ejemplo:

```bash
python acumulador_excels.py \
  --origen "C:/Datos/Excels" \
  --salida "C:/Datos/Salida/acumulado.xlsx" \
  --patron "Ventas_*.xlsx" \
  --hoja "Hoja1"
```

---

## 2) Opción Power Query (M) en Excel

1. Abre Excel.
2. Ve a **Datos > Obtener datos > Consulta en blanco**.
3. Abre el **Editor avanzado**.
4. Pega el contenido de `acumulador_power_query.m`.
5. Ajusta:
   - `CarpetaOrigen`
   - `Extension`
   - `NombreHoja`
6. Carga la consulta.

Con esto, cada vez que agregues nuevos archivos en la carpeta, solo debes pulsar **Actualizar todo** para acumularlos de nuevo.
