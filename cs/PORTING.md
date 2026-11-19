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

- [x] Etapa 2d — al día con PowerShell 1.9.2 en los datos: pstest.ps1 y dump.ps1 leen el GESTEM.ps1 del repo (bloques por marcas,
      toolspsblocks.ps1), data*.json exportados de 1.9.2, store con uid/borrado/borradosMano, huella y pendientes de Firebase (Fb.cs),
      campos Local/ReadOnly, requerimiento (fecha límite, avisos, fechas mm/dd) y regularización, Merge-StoreDeleted, Excel solo se lee
      (Data.ExcelWrites = false). test.bat compara con toolsmpare.ps1 -> testesult.txt (store por contenido, sin «guardado»).
      Prueba 19/11/2026: resumen (1278 líneas), los 15 archivos de datos y firebase.json idénticos.
      Arreglado en pstest.ps1: $script:HandRowBase se ponía a mitad de la prueba (la 1.ª pasada no daba uid).

- [x] Etapa 3a — ventana (Win.cs): main.xaml de 1.9.2, paneles y navegación, trámite (con caché), lista de clientes con búsqueda,
      filtros (estado, deuda, solo míos), línea de estado, aviso (toast), aviso de requerimientos, Ctrl+F, carga al abrir.
      Prueba: GESTEM.exe --uitest (en test.bat) -> testII_summary.txt: la lista de los 14 trámites = sus datos, búsqueda y deuda bien.
      Falta en 3a (shortcut en el código): textos de los demás paneles, nº de documentos en la lista, sesión, clave, Find-ExcelAuto.

- [x] Etapa 3b — ficha (WinForm.cs): construir según el trámite, rellenar/leer, tarjeta con IVA (tasas DGT aparte), margen y comisión,
      traducción (precio = traductora + 3 € × nº docs), fecha límite de requerimientos, aviso de NIE repetido, conflicto con otro PC,
      guardar, editar (doble clic / Enter / botón), Ctrl+S, Ctrl+N, Esc, aviso al cerrar con cambios.
      Prueba: pstest.ps1 abre la MISMA XAML con las funciones de PowerShell; los dos abren cada cliente en la ficha:
      valores leídos y textos de la tarjeta idénticos (658 líneas). Config guarda ya los ajustes 1.8–1.9 (clave, FbSync…).

## Encontrado al trasladar (arreglado en PowerShell 1.7.4)
- Clientes repetidos (AUT-2026-001 y AUT-001) en la copia real: una versión antigua importó pestañas hechas a mano como si fueran
  de GESTEM. C# los quitaba al abrir (Hand.StaleImports, solo por nombre y by=Excel), pero eso borraba también los clientes
  adoptados de la pestaña a mano al unirse a Firebase (mano-…, by=Excel): 171 de 274 en los datos de verdad. Quitado (19/11/2026):
  C# no quita nada al abrir, como PowerShell. selftest -> test\cs_keep.txt (KEEP OK, línea en result.txt): los clientes con uid,
  adoptados mano- incluidos, salen iguales de dos aperturas. Los datos de verdad en Firestore no tienen esas copias antiguas.
- La ficha recuerda su trámite (formT): al cambiar de trámite, los avisos tardíos de los campos viejos ya no rompen.
- La escritura directa dejaba calcChain.xml apuntando a celdas sin fórmula -> Excel no abría el archivo. Ahora se quita calcChain.xml.
- Textos con espacios al principio/final: SetAttribute('space', ns, …) escribía «d1p1:space» -> Excel no abría el archivo. Ahora xml:space.
- Repair-ExcelFile: una vez por PC arregla los archivos ya dañados (config ExcelRepaired).
- New-TramiteSheet nunca funcionaba en PowerShell (tabla de nº de expediente) -> ahora celda a celda; formatos sin bloquear.
- Pruebas: comtest.bat (¿Excel abre los archivos?), cleanup.bat (cierra Excel ocultos de las pruebas).
- Ojo: device_commit_files puede copiar una versión vieja si se reutiliza la misma ruta de origen: usar una carpeta nueva cada vez.
- Ojo: ejecutar GESTEM.ps1 desde otra carpeta cambia los accesos directos (Repair-Shortcuts) -> fix_shortcuts.bat.

## Añadido en PowerShell después de 1.7.4 (1.8.0 – 1.9.2) — hay que trasladarlo
Referencia: GESTEM.ps1 1.9.2 (commit 2e86b24). La prueba de igualdad compara con ESTA versión, no con la 1.7.4.

### Datos (afecta a la prueba de igualdad: va antes que las pantallas)
- Formato de «GESTEM datos\<trámite>.json» (Read-Store / Write-Store): cada cliente lleva `uid` (id para Firebase;
  los de antes salen de `<trámite>-<nº>`, repetido -> GUID), `updated`, `by`, `borrado {by, at}` (eliminado = marca, no se quita);
  el archivo lleva `adoptados` y `borradosMano` (clientes de pestañas a mano eliminados: uid, key, name, on, by, at).
  Huella H = Get-FbHash(nº + «BORRADO» si eliminado, campos sin `Local`); si cambia al guardar -> firebase.json `Pend`.
- Campos `Local` (los «Contraseña…»): se guardan solo en el PC, nunca se suben (Get-FbFields).
- Campos solo de GESTEM (Col vacío): colab, comtipo, comval, compag, ndocs.
- Trámites nuevos: `requerimiento` (pestaña REQUERIMIENTOS, fechas mm/dd, `flim` = Fecha límite de solo lectura =
  10 días hábiles desde apertura o notificación: Get-ReqDeadline / Update-ReqDeadline dentro de Resolve-Record,
  Get-ReqAlerts para el aviso de la barra lateral) y `regularizacion` (REGULARIZACION EXTRAORDINARIA).
- Merge-StoreDeleted: lo eliminado en GESTEM no vuelve al reimportar el Excel.
- Text18 (textos nuevos, se mezclan en Text) -> volver a exportar data\*.json (tools\dump.ps1).

### Excel solo se importa (1.9.0)
- GESTEM ya no escribe en el Excel: Sync-RecordToExcel devuelve $true y Sync-PendingToExcel 0 sin hacer nada; botón «volver a importar» oculto.
  El Excel solo se lee al empezar un trámite vacío. Lo escrito en las etapas 2b/2c (escribir en el xlsx / por COM) queda
  solo para Repair-ExcelFile y New-TramiteSheet; las pantallas NO deben llamar a la escritura.
- Find-ExcelAuto (1.8.1): si el Excel cambió de sitio o el PC es nuevo, se busca «GESTEM datos» y el Excel de al lado.

### 1.8 — pedido de la gestoría (18/11)
- Tasas DGT del canje (Cfg.TasasCanje «B=33; C=100», Get-CanjeTasa / Get-TasasFor, Split-Iva con tasas,
  línea «Tasas DGT (suplido)» sin IVA en factura y recibo).
- Margen y % en la ficha (Get-Margin = precio − tasas − gastos − comisión; Get-Commission: % o fijo).
- Traducción: precio = traductora + 3 € × nº docs (Cfg.TradProfit) con aviso.
- Página Colaboradores (Get-ColabItems, Get-CollaboratorNames, KPIs, Load-ColabPanel, Update-ColabView, Open-ColabItem,
  Mark-ColabPaid / Undo-ColabPaid, Set-ColabComState).

### Clave de GESTEM (1.8.2)
- Get/Set-GestemKey (DPAPI CurrentUser en Cfg.UpdateKey), Test-GestemKey, Ask-GestemKey (una vez por PC al abrir;
  si se cancela, se vuelve a pedir), Get-RepoFile (API de GitHub con la clave, repositorio privado) para las actualizaciones.
- La misma clave es la contraseña de Firebase.

### Firebase (1.9.0)
- Firestore REST (Connect-Firebase: email/contraseña, token 1 h; Invoke-Fs / Invoke-FbJson), proyecto en HANDOVER.md
  (no copiar los datos de la cuenta a este archivo: el repo es público).
- Documento `gestorias/<g>/clientes/<uid>`: t, id, datos (JSON sin campos Local), deleted, by, pc, updated, mod (hora del servidor).
- Estado por PC: %APPDATA%\GESTEM\firebase.json (lastMod, joined ''/'push'/'yes', Pend) — Get-FbState / Save-FbState / Add-FbPend.
- Sincronización cada 20 s en segundo plano (Step-FbSync desde Invoke-LiveRefresh, $script:FbSyncScript en un runspace);
  Receive-FbDocs aplica lo que llega (FbApplying: no se reenvía), Get-FbRow / Get-FbWantRow (nº repetido creado sin internet
  -> el de este PC pasa al siguiente libre + aviso).
- Primera unión (Join-FbDocs): bajar todo primero; mismo uid/nº + mismo NIE (o nombre, Test-FbSamePerson / Test-SameNameKey)
  = mismo cliente, gana el último cambio; mismo nº otra persona -> siguiente nº libre + aviso; clientes de pestañas a mano
  adoptados una vez (`mano-<trámite>-<md5>`, Get-HandDelUid); solo locales se suben con exists=false (409 -> volver a unir).
- Interruptor Set-FbSyncSwitch (Ajustes) + %APPDATA%\GESTEM\firebase_apagado.txt. Panel admin: Update-AdmFbState.
- Pruebas de PowerShell: tests\fbtest.ps1 (sin red), con red solo en `_tests/pruebas` (ver HANDOVER.md).

### Eliminar cliente (1.9.0)
- Remove-Client / Set-ClientDeleted (marca `borrado`; los de pestañas a mano por nombre en `borradosMano`, uid `hand-…`),
  Undo-RemoveClient (aviso con «Deshacer» y Ctrl+Z); se sincroniza como deleted=true.

### Carpeta de documentos compartida (1.9.1 – 1.9.2)
- Cfg.DocsDir en Ajustes (Choose-DocsFolder, Set-DocsRoot: desde la carpeta local mueve, desde otra solo copia).
- Al abrir (Initialize-DocsFolder): sin carpeta y con Firebase -> busca «GESTEM documentos» en Get-CloudRoots
  (Find-CloudDocsDir, profundidad 3) y mueve una vez los de este PC (Merge-DocsFolder: nunca pisa; mismo nombre -> «nombre (2)»;
  idénticos -> uno). Get-LocalDocsRoot = junto al Excel.
- La carpeta sigue al cliente renumerado (Repair-DocsFolders / Repair-AllDocsFolders: tras Join-FbDocs, Receive-FbDocs y al abrir;
  por nombre único, nunca si es ambiguo); Get-ClientDocsDir prefiere «nº - nombre» exacto.
- Pruebas: tests\docsfolder.ps1, tests\docsrenum.ps1.

### Ventana (1.8 – 1.9)
- La XAML cambió (página Colaboradores, DocsDirBox, aviso de requerimientos, Eliminar cliente…): main.xaml se vuelve a sacar
  del bloque `$xaml = @'…'@` de GESTEM.ps1 1.9.2 en la etapa 3.
- Arreglo 1.9.0: lista vacía de clientes (Refresh-ClientGrid / Update-StatusLine usaban `-not $script:Rows`) — no repetirlo.

## Pendiente
- [x] Etapa 2d — ver «Hecho»
- [x] Etapa 3c — menú del clic derecho (WinMenu.cs): estados, «pagado», asignar responsable (varios a la vez), copiar nombre/NIE/teléfono,
      eliminar cliente (marca borrado; pestaña a mano: borradosMano), Ctrl+Z (un dato, responsable, eliminar). Prueba: pstest.ps1 usa
      las funciones reales con la hora fija (Get-Date) y pinta el menú de cada cliente (2375 líneas iguales, datos y firebase.json iguales).
      Recibo, documentos, carpeta y abrir Excel salen desactivados hasta su etapa.
      El submenú «Estado» salía como una sola opción con todos los estados juntos en PowerShell 1.9.2; arreglado en 1.9.3.
- [x] Etapa 3d — documentos (Docs.cs, WinDocs.cs): carpeta «GESTEM documentos\<pestaña>\<nº> - <nombre>», panel (lista, abrir,
      añadir con diálogo o arrastrando, quitar a la papelera, abrir carpeta), nº de documentos en la lista (clip), carpetas que
      quedaron con el nº de otro cliente se arreglan al abrir (Repair-AllDocsFolders). Prueba: tools\docsfixture.ps1 crea la misma
      carpeta para los dos; árbol, cuentas, panel y añadir idénticos. Sin probar (diálogos): quitar, abrir, diálogo de añadir.
      Pendiente: elegir carpeta en Ajustes (etapa 4), buscarla en la nube (5b), copia a la carpeta externa (5).
      Versión del C#: 1.9.3.
- [x] MODO PRUEBA (hasta terminar el traslado): GESTEM.exe sin --real trabaja con una copia (cs\data\prueba: config, registro,
      Excel y «GESTEM datos» copiados la primera vez; borrar la carpeta para copiar de nuevo) y documentos de prueba (cs\data\docs).
      App.Guard para toda escritura fuera de las carpetas de prueba (datos, documentos, Excel, config, copias, firebase.json);
      el registro de verdad no se toca. La ventana lo dice en el título. Prueba: selftest -> test\cs_guard.txt (GUARD OK).
      Quitar el modo prueba (o dejar --real por defecto) en la etapa 7.
- [x] Etapa 4 — Resumen (estadísticas y gráfico), Ajustes, sesión/usuarios, admin, Colaboradores, margen (3b). El control remoto
      (control.json de GitHub) pasa a la etapa 5: necesita Get-RepoFile, que llega con las actualizaciones.
  - [x] 4a Resumen (WinStats.cs): filtros, tabla por trámite o por estado, tarjetas, gráfica día/semana/mes. Prueba: SF/ST/CH en
        el resumen (4 alcances x 4 periodos x 3 personas, y los 3 modos de la gráfica) + «UI resumen» (gráfica dibujada en la ventana).
        Ps.Sort copia el Sort-Object de PS 5.1 (no estable: los empates salen en un orden concreto); también en la lista de documentos.
        Ojo: los @{} de los trámites ya vienen del JSON en el orden en que PS los recorre; meterlos en otro Hashtable lo cambia.
        Arreglado de paso: ExcelCom.Close comprobaba FullName (en OneDrive es https://…) -> ahora la ruta del Excel abierto;
        compare.ps1 ya no da IGUAL si pstest.ps1 falla sin dejar ps_summary.txt.
  - [x] 4b Ajustes y empresa (WinSettings.cs, Invoices.cs): panel, guardar, logo (copia en la carpeta de datos), nombre en la
        ventana, cambiar la carpeta de documentos (Set-DocsRoot), nº de la próxima factura (facturas.json), icono como recurso.
        Prueba: SP/SC/SD en el resumen + facturas.json + «UI ajustes». Pendiente: copia externa (Backup-External, «copiar ahora»),
        «buscar OneDrive» y «buscar actualizaciones» (etapas 5/5b; botones ocultos). Repair-Shortcuts no se traslada (etapa 7).
  - [x] 4c Sesión (WinSession.cs): quién usa GESTEM (lista de nombres, la última persona entra sola con «¿No eres tú?»), sin atajos
        mientras se elige. Prueba: SS en el resumen. Pendiente: «actividad GESTEM - PC.txt» en la carpeta externa (etapa 5).
  - [x] 4d Administrador (WinAdmin.cs; Ctrl+Mayús+Alt+A): contraseña (admin.dat cifrado con DPAPI), bloqueos de Ajustes, registro
        de actividad y de errores, textos traducidos, cambiar contraseña, pantalla de pausa. En modo prueba admin.dat se copia (solo
        se lee el de verdad). Prueba: AH/AL/AT/AD/AR en el resumen (también un PC que no es el del creador) + «UI admin».
        Pendiente: leer control.json de GitHub (Check-Remote, con Get-RepoFile en la etapa 5), «buscar actualización» (5),
        encender/apagar Firebase desde el panel (Set-FbSyncSwitch, 5b; la casilla no se puede cambiar).
        pstest: psblocks.ps1 carga también TrText/TrTip, OwnerPrints, LockKeys y AppVersion de GESTEM.ps1.
  - [x] 4e Colaboradores (WinColab.cs): lista de colaboradores/abogados con trámites y comisión por pagar, KPIs, «solo por pagar»,
        marcar comisión pagada (solo en los datos de GESTEM, no en el Excel) con «Deshacer», doble clic abre la ficha en su trámite.
        Prueba: CO texts/CO0–CO6 en el resumen y los .json (marcar 3 y deshacer) + «UI colaboradores». En pstest, Load-Data -All
        rehace la caché solo para este bloque y los textos salen de las líneas reales de Apply-Language.
- [ ] Etapa 5 — copia externa, actualizaciones + control.json, carpeta de documentos en la nube, Firebase (por este orden);
      después facturas PDF (con tasas DGT) y exportar CSV/PDF.
  - [x] 5a Copia externa (ExtBackup.cs): Excel del día (60 días), ajustes, documentos (solo lo nuevo), datos y «actividad GESTEM -
        PC.txt» sin nombres de clientes; al abrir (y la primera vez busca la nube de la empresa), al cerrar, «copiar ahora», al cambiar
        la carpeta en Ajustes; «buscar OneDrive». Como Get-ChildItem sin -Force: sin archivos/carpetas ocultos (desktop.ini).
        Modo prueba: la «nube» es cs\data\nube (App.CloudRootsOverride) y la carpeta externa de la config se cambia a
        cs\data\nube\GESTEM copias. Prueba: EX/EXF/EXA/EX2 + «UI copia externa».
  - [x] 5b Actualizaciones + control.json (Updater.cs): clave de GESTEM (pedirla una vez, DPAPI), leer version.txt en segundo plano,
        «buscar actualización» (Ajustes y admin), control.json (mensaje una vez, pausa remota, copia sin conexión, 404 la borra).
        Solo LEE de GitHub. Con versión nueva solo avisa: instalar un GESTEM.exe nuevo queda para la etapa 7 (hace falta decidir dónde
        se publica el .exe); nunca toca GESTEM.ps1 ni version.txt. Las pruebas (Quiet) no van a internet.
        Prueba: RC0–RC7/RV/RU/RK con un GitHub de mentira (Updater.FetchHook / Get-RepoFile de la prueba).
  - [x] 5c Carpeta de documentos en la nube (ExtBackup.cs: FindCloudDocsDir; MainWin.InitializeDocsFolder): al abrir, sin carpeta elegida
        y con la sincronización encendida, usa la «GESTEM documentos» más reciente de la nube (hasta 3 carpetas de hondo, sin ocultas)
        y pasa allí los documentos de este PC. En modo prueba nunca (carpeta de prueba fija, sincronización apagada).
        Prueba: DC find/off/on/none con dos «nubes» de prueba. ToastHook va ahora en ShowToast (todos los avisos, como pstest).
        build.bat compila a GESTEM.new.exe y solo cambia GESTEM.exe si compila y arranca (--smoke); si falla, queda el anterior.
  - [x] 5d Firebase (FbSync.cs): unión, envío, recibir, interruptor de admin (solo con --real), actualización cada 20 s. En modo prueba nunca sincroniza; Fb.CheckName no deja usar nada fuera de gestorias/elmahi/_tests/ sin --real; los campos Local (contraseñas) nunca se suben. Igualdad sin red en test.bat (FB S0–S3). Red de verdad aparte: `GESTEM.exe --fbtest <carpeta> elmahi/_tests/<nombre>` (datos inventados, empieza vacía, lo borra todo al final, escribe cs_fb.txt y cs_fb_paths.txt; F10/F11: la vuelta de 20 s de la ventana, Fb.Step, con el interruptor apagado y encendido; Fb.TestRun solo lo pone --fbtest). Restos de una prueba cortada: `GESTEM.exe --fbclean <carpeta>` (busca con una consulta limitada a _tests/, borra solo ahí; cs_fbclean.txt).
- [x] Etapa 6 — las dos versiones lado a lado con una copia de los datos (tools\side.ps1: prep / start / dump ps|cs / stop).
      PowerShell 1.9.3 corre desde una copia del script con su propio APPDATA (cs\test\side), C# en modo prueba; Firebase apagado
      en los dos. dump lee por UI Automation las 14 listas de trámites, la ficha del 1.er cliente de cada uno, resumen, colaboradores
      y ajustes. Prueba 19/11/2026: 1035 líneas idénticas. Diferencias encontradas:
      1) línea de estado sin deudas: PowerShell «(-)», C# «(0,00 €)» -> arreglado en C#.
      2) C# quitaba al abrir los «repetidos» de pestañas a mano (Homologación 14/7, Autónomo 34/17…): esa purga ya no existe.
         Repetido 19/11/2026 con cs\data\prueba nuevo (copia real completa) y sin igualar los datos: 1073 líneas idénticas.
      Ojo: abrir el GESTEM.ps1 copiado cambia los accesos directos -> fix_shortcuts.bat después.
- [ ] Etapa 7 — última actualización de PowerShell que instala GESTEM.exe
  Antes (sección 0 del plan):
  - [x] config.json: C# guarda tal cual las claves que no conoce (PowerShell las pierde: Load-Config solo copia las suyas).
        Prueba: «solo C#» en result.txt (cs_only.txt).
  - [x] Al iniciar sesión, «actividad GESTEM - PC.txt» en la carpeta externa (Start-Session). Prueba: SSA en el resumen.
