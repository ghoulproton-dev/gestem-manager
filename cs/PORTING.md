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

## Pendiente
- [ ] Etapa 2 — Excel: leer xlsx (pestañas a mano, listas), escribir directo (Save-ClientsToFile), COM como respaldo, copias
- [ ] Etapa 3 — ventana: lista, ficha, menú de clic derecho, deshacer, avisos, atajos, documentos
- [ ] Etapa 4 — Resumen (estadísticas y gráfico), Ajustes, sesión/usuarios, admin, control remoto
- [ ] Etapa 5 — facturas PDF, exportar CSV/PDF, copia externa, actualizaciones
- [ ] Etapa 6 — comparar las dos versiones lado a lado con una copia de los datos
- [ ] Etapa 7 — última actualización de PowerShell que instala GESTEM.exe
