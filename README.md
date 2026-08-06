# AutomatizaciÃ³n de Reporte de Multas SUNAT

Proyecto Python para generar un Excel independiente por empresa, consultar SUNAT con Playwright y llenar la plantilla `Base.xlsx` sin cambiar su formato base.

## Arquitectura propuesta

- `main.py`: punto de entrada del programa.
- `sunat_multas/cli.py`: argumentos, ciclo por empresa y manejo de errores.
- `sunat_multas/excel_io.py`: lectura de `LISTA`, copia de `Base.xlsx` y escritura con `openpyxl`.
- `sunat_multas/sunat_client.py`: login, navegaciÃ³n, consulta y extracciÃ³n desde SUNAT con Playwright.
- `sunat_multas/config.py`: lectura de `config.yaml`.
- `sunat_multas/logging_setup.py`: logs e incidencias.
- `logs/incidentes.csv`: empresas con error de autenticaciÃ³n, sin registros o errores de extracciÃ³n.
- `output/`: archivos finales por empresa.

## Flujo implementado

1. Lee `lista_reporte_de_multas_sector_b.xlsx`, hoja `LISTA`.
2. Valida que existan `ITEM`, `NOMBRE`, `RUC`, `USUARIO` y `CLAVE`.
3. Copia `Base.xlsx` a `output/<NOMBRE EMPRESA>.xlsx`.
4. Ingresa a SUNAT con RUC, usuario y clave.
5. Navega por `Manifiesto de Carga de Ingreso` -> `Consultas` -> `Consulta Manifiesto Desconsolidado`.
6. Espera 10 segundos para que la consulta termine de cargar.
7. Consulta por rango de fechas.
8. Extrae la tabla de resultados.
9. Clasifica registros en `IMPO118`, `IMPO235`, `EXPO118` o `EXPO235`.
10. Inserta solo valores en las columnas existentes.
11. Registra incidencias y continÃºa con la siguiente empresa.

## InstalaciÃ³n

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
playwright install chromium
```

## Uso

Primero prueba sin abrir SUNAT:

```powershell
python main.py --mes 8 --anio 2026 --dry-run
```

Luego ejecuta el flujo completo:

```powershell
python main.py --mes 8 --anio 2026
```

Si no indicas `--mes`, el programa solicita el número del mes al iniciar. Para el mes `07` del año `2026`, consulta desde `01/07/2026` hasta `31/07/2026`.

## Ajuste necesario de SUNAT

SUNAT cambia con frecuencia nombres de campos, iframes y botones. Por eso `config.yaml` deja los selectores en un solo lugar. Si la automatizaciÃ³n no encuentra un campo o botÃ³n, actualiza el selector correspondiente:

- `ruc_input`
- `ruc_mode_button`
- `user_input`
- `password_input`
- `login_button`
- `menu_search_input`
- `date_from_input`
- `date_to_input`
- `search_button`
- `result_table`

Puedes separar alternativas con `|`.

## Consideraciones importantes

- El rango `--desde` y `--hasta` es obligatorio para evitar consultas ambiguas y reducir el riesgo de omitir registros.
- Si SUNAT muestra CAPTCHA, 2FA o una validaciÃ³n humana, el script registra la incidencia y sigue con la siguiente empresa.
- Las credenciales se leen desde el Excel original; no se escriben en logs.
- La plantilla se conserva porque cada archivo se genera con una copia directa de `Base.xlsx`.
## Ejecutar una sola empresa

Puedes probar una empresa antes de procesar todas:

```powershell
python main.py --desde 2026-08-01 --hasta 2026-08-31 --item 1 --dry-run
```

Qué hace cada parte:

- `python`: ejecuta Python.
- `main.py`: abre el programa de automatización.
- `--desde 2026-08-01`: usa el 1 de agosto de 2026 como fecha inicial.
- `--hasta 2026-08-31`: usa el 31 de agosto de 2026 como fecha final.
- `--item 1`: procesa solo la fila cuyo `ITEM` sea `1` en la hoja `LISTA`.
- `--dry-run`: prueba lectura y creación de Excel, pero no entra a SUNAT.

Para entrar a SUNAT con una sola empresa:

```powershell
python main.py --desde 2026-08-01 --hasta 2026-08-31 --item 1
```

También puedes filtrar por RUC o por parte del nombre:

```powershell
python main.py --desde 2026-08-01 --hasta 2026-08-31 --ruc 20602733808
python main.py --desde 2026-08-01 --hasta 2026-08-31 --nombre "PRADIVO"
```

