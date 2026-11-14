# GESTEM en C# — estado del traslado

Se construye en el PC de prueba con el compilador de .NET Framework 4.8 (csc, C# 5) — sin instalar nada.
`all.bat` = compilar (`build.bat`) + prueba de igualdad (`test.bat`).
La versión PowerShell (GESTEM.ps1) sigue siendo la que usa la gestoría hasta el final.

## Reglas
- C# 5: sin `$"..."`, sin `?.`, sin `nameof`, sin propiedades con valor inicial.
- Cada parte se compara con la función de PowerShell real (tools/pstest.ps1 carga las funciones de GESTEM.ps1).
- Después de copiar archivos al PC, comprobar el tamaño (la copia puede llegar vieja).

## Hecho
- [x] Compilar en el PC (csc 4.8), la ventana carga main.xaml (XAML de 1.7.3 sin cambios)
- [x] Datos fijos exportados de GESTEM.ps1: data/tramites.json, text.json, lists.json, admin_tr.json (tools/dump.ps1)
- [x] Etapa 1 — datos: Json, Ps (comportamientos de PowerShell), Config, Store (leer/guardar/bloqueo/copia diaria),
      Resolve-Record, New-RowObject, Build-TramiteData.
      Prueba 13/11/2026 con los datos de la copia real: 309 clientes en 12 trámites, resumen idéntico y
      los 9 archivos guardados idénticos a los de PowerShell.

- [x] Etapa 2a — leer el xlsx sin Excel: Xlsx (Read-XlsxSheets, Read-XlsxValidationLists), Hand (pestañas hechas a mano,
      Sync-HandmadeTabs, Update-ListsFromFile). Prueba 14/11/2026: idéntico (9 pestañas a mano, 309 clientes campo a campo).
      Ojo: se usan Hashtable con CurrentCultureIgnoreCase para recorrer en el mismo orden que @{} de PowerShell.
- [x] Etapa 2b — escribir directo en el xlsx: ExcelFile (Get-ClientOps, Save-ClientsToFile, Backup-Excel).
      Prueba 14/11/2026: editar, cliente nuevo, lote de 2, pasada SUM y trámite sin pestaña -> los .xlsx resultantes
      son idénticos byte a byte en todas las entradas.

- [x] Etapa 2c — Excel por COM (ExcelCom con IDispatch por reflexión, no «dynamic»: dynamic fallaba al cerrar y volver a abrir Excel),
      Data (Load-Data, Refresh-FromExcel, Sync-PendingToExcel, Save-ClientStore…). Prueba 14/11/2026: leer pestaña, guardar con Excel,
      cliente nuevo, crear pestaña AUTÓNOMO -> mismas celdas (valores y fórmulas) en las 16 pestañas, y Excel abre los dos archivos.

## Encontrado al trasladar (arreglado en PowerShell 1.7.4)
- La escritura directa dejaba calcChain.xml apuntando a celdas sin fórmula -> Excel no abría el archivo. Ahora se quita calcChain.xml.
- Textos con espacios al principio/final: SetAttribute('space', ns, …) escribía «d1p1:space» -> Excel no abría el archivo. Ahora xml:space.
- Repair-ExcelFile: una vez por PC arregla los archivos ya dañados (config ExcelRepaired).
- New-TramiteSheet nunca funcionaba en PowerShell (tabla de nº de expediente) -> ahora celda a celda; formatos sin bloquear.
- Pruebas: comtest.bat (¿Excel abre los archivos?), cleanup.bat (cierra Excel ocultos de las pruebas).
- Ojo: device_commit_files puede copiar una versión vieja si se reutiliza la misma ruta de origen: usar una carpeta nueva cada vez.
- Ojo: ejecutar GESTEM.ps1 desde otra carpeta cambia los accesos directos (Repair-Shortcuts) -> fix_shortcuts.bat.

## Pendiente
- [ ] (hecho arriba) Etapa 2c — Excel por COM (respaldo cuando el archivo está abierto): Open-Book/Close-Book/Invoke-Excel, Get-Sheet,
      New-TramiteSheet, Restore-Formula, Set-CellValue, Save-Client, Read-TramiteData, Import-TramiteFromExcel,
      Sync-RecordToExcel, Save-PendingBatchToFile, Sync-PendingToExcel, Repair-TotalsFormulas, Refresh-FromExcel, Load-Data
- [ ] Etapa 3 — ventana: lista, ficha, menú de clic derecho, deshacer, avisos, atajos, documentos
- [ ] Etapa 4 — Resumen (estadísticas y gráfico), Ajustes, sesión/usuarios, admin, control remoto
- [ ] Etapa 5 — facturas PDF, exportar CSV/PDF, copia externa, actualizaciones
- [ ] Etapa 6 — comparar las dos versiones lado a lado con una copia de los datos
- [ ] Etapa 7 — última actualización de PowerShell que instala GESTEM.exe
