#  GESTEM MANAGER  -  Gestión de clientes para gestoría
#  Creado por Yassine Belali  ·  2026
#  Escribe directamente en el Excel de la gestoría usando Microsoft Excel.
#  Requiere: Windows 10/11, Windows PowerShell 5.1 y Microsoft Excel instalado.

$ErrorActionPreference = 'Stop'
trap {
    $msg = "$($_.Exception.Message)`r`n$($_.InvocationInfo.PositionMessage)"
    try { $d = Join-Path $env:APPDATA 'GESTEM'; New-Item -ItemType Directory -Force -Path $d | Out-Null
          Add-Content -Path (Join-Path $d 'gestem.log') -Value ("{0:yyyy-MM-dd HH:mm:ss}  FATAL  {1}" -f (Get-Date), $msg) -Encoding UTF8 } catch {}
    try { Add-Type -AssemblyName PresentationFramework; [void][Windows.MessageBox]::Show($msg, 'GESTEM MANAGER - error', 'OK', 'Error') } catch {}
    exit 1
}
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

$script:AppName    = 'GESTEM MANAGER'
$script:AppVersion = '1.1'
$script:UpdateBase = 'https://raw.githubusercontent.com/ghoulproton-dev/gestem-manager/main'
$script:AppDir     = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:DataDir    = Join-Path $env:APPDATA 'GESTEM'
$script:CfgPath    = Join-Path $script:DataDir 'config.json'
$script:BackupDir  = Join-Path $script:DataDir 'copias'
$script:LogPath    = Join-Path $script:DataDir 'gestem.log'
$script:ES         = [Globalization.CultureInfo]::GetCultureInfo('es-ES')
New-Item -ItemType Directory -Force -Path $script:DataDir, $script:BackupDir | Out-Null

function Write-Log([string]$msg) {
    try { Add-Content -Path $script:LogPath -Value ("{0:yyyy-MM-dd HH:mm:ss}  {1}" -f (Get-Date), $msg) -Encoding UTF8 } catch {}
}

# =====================================================================
#  TRÁMITES  (para añadir Homologación, Extranjería... se añade aquí otro bloque)
# =====================================================================
$script:Nacionalidades = @('MARROQUI','ARGELINA','TUNECINA','MAURITANA','SENEGALESA','MALIENSE','GAMBIANA',
    'GUINEANA','NIGERIANA','GHANESA','EGIPCIA','SIRIA','PAKISTANI','INDIA','BANGLADESI','CHINA','FILIPINA',
    'UCRANIANA','RUSA','GEORGIANA','RUMANA','BULGARA','COLOMBIANA','VENEZOLANA','ECUATORIANA','PERUANA',
    'BOLIVIANA','PARAGUAYA','URUGUAYA','ARGENTINA','CHILENA','BRASILEÑA','DOMINICANA','CUBANA','HONDUREÑA',
    'NICARAGÜENSE','SALVADOREÑA','MEXICANA','TURCA','ESPAÑOLA','FRANCESA','ITALIANA','PORTUGUESA','BRITANICA','OTRA')
$script:FormasPago = @('CNTA CHAY','CNTA SAFA','EFECTIVO','EFECTIVO CHAY','EFECTIVO SAFA','TPV','BIZUM','TRANSFERENCIA')

$script:Tramites = @(
    @{
        Id        = 'canje'
        Title     = @{ es = 'Canje de permiso (DGT)'; en = 'Driving licence exchange (DGT)'; fr = 'Échange de permis (DGT)'; ar = 'استبدال رخصة القيادة (DGT)' }
        Sheet     = 'DGT ACTUALIZADO'
        FirstRow  = 11
        LastRow   = 1009
        LastCol   = 'T'
        IdCol     = 'A'
        NameCol   = 'B'
        DocCol    = 'D'
        StatusCol = 'N'
        PayCol    = 'M'
        PendCol   = 'L'
        DaysCol   = 'T'
        TipoCol   = 'E'
        DefaultNationality = 'MARROQUI'      # lo que pone sola la fórmula del Excel
        StatusColors = @{ 'Documentación' = '#F4B183'; 'Pendiente Cita' = '#F8CBAD'; 'Verificación Origen' = '#FFF2CC';
                          'En Tramitación' = '#FFE699'; 'Permiso Provisional' = '#C6E0B4'; 'Finalizado' = '#A9D08E'; 'Incidencia' = '#FF9B9B' }
        Fields = @(
            @{ Key='name'; Col='B'; Type='Upper'; Section='client'; Span=2; Required=$true
               Label=@{ es='Nombre y apellidos'; en='Full name'; fr='Nom et prénom'; ar='الاسم الكامل' } }
            @{ Key='doc'; Col='D'; Type='Upper'; Section='client'
               Label=@{ es='Documento (NIE / DNI)'; en='ID document (NIE / DNI)'; fr="Pièce d'identité (NIE / DNI)"; ar='وثيقة الهوية (NIE / DNI)' } }
            @{ Key='nat'; Col='C'; Type='ComboEdit'; Section='client'; Role='nationality'; List='nationalities'
               Label=@{ es='Nacionalidad'; en='Nationality'; fr='Nationalité'; ar='الجنسية' } }
            @{ Key='cert'; Col='P'; Type='Text'; Section='client'
               Label=@{ es='Certificado digital'; en='Digital certificate'; fr='Certificat numérique'; ar='الشهادة الرقمية' } }
            @{ Key='tipo'; Col='E'; Type='Combo'; Section='proc'; List='excel:E'
               Label=@{ es='Tipo de permiso'; en='Licence type'; fr='Type de permis'; ar='نوع الرخصة' } }
            @{ Key='pasos'; Col='F'; Type='Combo'; Section='proc'; List='excel:F'
               Label=@{ es='Pasos'; en='Steps'; fr='Étapes'; ar='الخطوات' } }
            @{ Key='loc'; Col='G'; Type='Text'; Section='proc'
               Label=@{ es='Localizador'; en='Reference code'; fr='Code de suivi'; ar='رمز التتبع' } }
            @{ Key='fsol'; Col='O'; Type='Date'; Section='proc'; DefaultToday=$true
               Label=@{ es='Fecha de solicitud'; en='Application date'; fr='Date de demande'; ar='تاريخ الطلب' } }
            @{ Key='status'; Col='N'; Type='Pills'; Section='proc'; Span=2; List='excel:N'; Colored=$true
               Label=@{ es='Estado del trámite'; en='Procedure status'; fr='État de la démarche'; ar='حالة الإجراء' } }
            @{ Key='fent'; Col='Q'; Type='Date'; Section='proc'
               Label=@{ es='Fecha de entrega'; en='Delivery date'; fr='Date de remise'; ar='تاريخ التسليم' } }
            @{ Key='price'; Col='J'; Type='Money'; Section='pay'; Role='price'
               Label=@{ es='Precio total (€)'; en='Total price (€)'; fr='Prix total (€)'; ar='السعر الإجمالي (€)' } }
            @{ Key='paystate'; Col='M'; Type='Pills'; Section='pay'; List='excel:M'; Role='paystate'; Default='Pendiente'
               Label=@{ es='Estado de pago'; en='Payment status'; fr='État du paiement'; ar='حالة الدفع' } }
            @{ Key='paid'; Col='K'; Type='Money'; Section='pay'; Role='paid'
               Label=@{ es='Ha pagado (€) — solo si es parcial'; en='Amount paid (€) — partial only'; fr='Montant payé (€) — si partiel'; ar='المبلغ المدفوع (€) — عند الدفع الجزئي' } }
            @{ Key='forma'; Col='S'; Type='ComboEdit'; Section='pay'; List='payments'
               Label=@{ es='Forma de pago'; en='Payment method'; fr='Mode de paiement'; ar='طريقة الدفع' } }
            @{ Key='obs'; Col='R'; Type='Multiline'; Section='notes'; Span=2
               Label=@{ es='Observaciones'; en='Notes'; fr='Observations'; ar='ملاحظات' } }
        )
    }
)

# =====================================================================
#  IDIOMAS
# =====================================================================
$script:Langs = [ordered]@{ es = 'Español'; en = 'English'; fr = 'Français'; ar = 'العربية' }
$script:Text = @{
 es = @{
  navNew='Nuevo cliente'; navClients='Clientes'; navSettings='Ajustes'; navAbout='Acerca de'
  secclient='Datos del cliente'; secproc='Trámite'; secpay='Pago'; secnotes='Notas'
  newTitle='Nuevo cliente'; editTitle='Editar cliente {0}'
  btnSave='Guardar cliente'; btnSaveEdit='Guardar cambios'; btnClear='Vaciar formulario'; btnCancel='Cancelar'
  sumPrice='Precio'; sumPaid='Pagado'; sumPend='Pendiente'; cardEmpty='Nombre del cliente'; cardNew='Nuevo'
  dupWarn='Atención: este documento ya está en {0} ({1}).'
  dupAsk='Este documento ya está en {0} ({1}).{2}¿Guardar igualmente?'
  reqName='Escribe el nombre del cliente.'
  saved='Cliente guardado como {0}.'; savedEdit='Cambios guardados en {0}.'
  search='Buscar (nombre, NIE, ID, localizador)'; filterAll='Todos los estados'; onlyDebt='Solo con pago pendiente'
  btnEdit='Editar cliente'; btnOpenExcel='Abrir Excel'; btnRefresh='Actualizar'
  colId='ID'; colName='Cliente'; colDoc='NIE / DNI'; colTipo='Permiso'; colStatus='Estado'; colPay='Pago'; colPend='Pendiente'; colDays='Días'
  setTitle='Ajustes'; setFile='Archivo Excel de clientes'; btnChange='Cambiar…'; setLang='Idioma'
  setNat='Nacionalidad por defecto en clientes nuevos'; setBackup='Hacer una copia de seguridad del Excel antes de cada guardado'
  setOpen='Abrir Excel después de guardar un cliente'; btnBackups='Abrir carpeta de copias'; btnSaveSettings='Guardar ajustes'
  setNote='Los precios y las listas (tipos de permiso, pasos) se cambian en la pestaña AJUSTES del Excel.'
  settingsSaved='Ajustes guardados.'; chooseFile='Elige el archivo Excel de clientes'
  welcome='Para empezar, elige tu archivo Excel de clientes.'
  loading='Leyendo el Excel…'; savingMsg='Guardando en el Excel…'; ready='{0} clientes · {1} con pago pendiente ({2})'
  noFile='Elige tu archivo Excel en Ajustes.'; tramite='Trámite'
  errNoFile='No encuentro el archivo Excel. Elígelo en Ajustes.'
  errNoSheet='Este Excel no tiene la pestaña «{0}». Elige el archivo correcto en Ajustes.'
  errLocked='El Excel está abierto en otro ordenador o en otro programa. Ciérralo allí y vuelve a intentarlo.'
  errBusy='Excel está ocupado: si estás escribiendo en una celda, pulsa Enter en Excel y vuelve a intentarlo.'
  errFull='No quedan filas libres en el Excel.'; errNoExcel='GESTEM MANAGER necesita Microsoft Excel instalado en este ordenador.'
  errGeneric='No se pudo completar la operación:'; selectFirst='Selecciona un cliente de la lista.'
  aboutText='Programa de gestión de clientes para gestorías: canjes de permiso de conducir y otros trámites. Guarda todo directamente en tu Excel.'
  aboutBy='Creado por Yassine Belali'; aboutVer='Versión {0}'
  partialOnly='(solo si el pago es Parcial)'; none='(ninguno)'
  updating='Actualizando GESTEM a la versión {0}…'; updated='GESTEM actualizado a la versión {0}.'; upToDate='Ya tienes la última versión ({0}).'; updateFail='No se pudo buscar la actualización (¿hay internet?).'; btnUpdate='Buscar actualizaciones'
 }
 en = @{
  navNew='New client'; navClients='Clients'; navSettings='Settings'; navAbout='About'
  secclient='Client details'; secproc='Procedure'; secpay='Payment'; secnotes='Notes'
  newTitle='New client'; editTitle='Edit client {0}'
  btnSave='Save client'; btnSaveEdit='Save changes'; btnClear='Clear form'; btnCancel='Cancel'
  sumPrice='Price'; sumPaid='Paid'; sumPend='Outstanding'; cardEmpty='Client name'; cardNew='New'
  dupWarn='Note: this document is already in {0} ({1}).'
  dupAsk='This document is already in {0} ({1}).{2}Save anyway?'
  reqName="Enter the client's name."
  saved='Client saved as {0}.'; savedEdit='Changes saved in {0}.'
  search='Search (name, NIE, ID, reference)'; filterAll='All statuses'; onlyDebt='Outstanding payment only'
  btnEdit='Edit client'; btnOpenExcel='Open Excel'; btnRefresh='Refresh'
  colId='ID'; colName='Client'; colDoc='NIE / DNI'; colTipo='Licence'; colStatus='Status'; colPay='Payment'; colPend='Outstanding'; colDays='Days'
  setTitle='Settings'; setFile='Client Excel file'; btnChange='Change…'; setLang='Language'
  setNat='Default nationality for new clients'; setBackup='Back up the Excel file before every save'
  setOpen='Open Excel after saving a client'; btnBackups='Open backups folder'; btnSaveSettings='Save settings'
  setNote='Prices and lists (licence types, steps) are changed in the AJUSTES tab of the Excel file.'
  settingsSaved='Settings saved.'; chooseFile='Choose the client Excel file'
  welcome='To start, choose your client Excel file.'
  loading='Reading the Excel file…'; savingMsg='Saving to Excel…'; ready='{0} clients · {1} with outstanding payment ({2})'
  noFile='Choose your Excel file in Settings.'; tramite='Procedure'
  errNoFile="Can't find the Excel file. Choose it in Settings."
  errNoSheet='This Excel file has no «{0}» tab. Choose the right file in Settings.'
  errLocked='The Excel file is open on another computer or program. Close it there and try again.'
  errBusy="Excel is busy: if you're typing in a cell, press Enter in Excel and try again."
  errFull='There are no free rows left in the Excel file.'; errNoExcel='GESTEM MANAGER needs Microsoft Excel installed on this computer.'
  errGeneric="The operation couldn't be completed:"; selectFirst='Select a client in the list.'
  aboutText='Client management for administrative offices: driving licence exchanges and other procedures. Everything is saved straight into your Excel file.'
  aboutBy='Created by Yassine Belali'; aboutVer='Version {0}'
  partialOnly='(only if payment is Partial)'; none='(none)'
  updating='Updating GESTEM to version {0}…'; updated='GESTEM updated to version {0}.'; upToDate='You already have the latest version ({0}).'; updateFail="Couldn't check for updates (is there internet?)."; btnUpdate='Check for updates'
 }
 fr = @{
  navNew='Nouveau client'; navClients='Clients'; navSettings='Paramètres'; navAbout='À propos'
  secclient='Informations du client'; secproc='Démarche'; secpay='Paiement'; secnotes='Notes'
  newTitle='Nouveau client'; editTitle='Modifier le client {0}'
  btnSave='Enregistrer le client'; btnSaveEdit='Enregistrer les modifications'; btnClear='Vider le formulaire'; btnCancel='Annuler'
  sumPrice='Prix'; sumPaid='Payé'; sumPend='Reste à payer'; cardEmpty='Nom du client'; cardNew='Nouveau'
  dupWarn='Attention : ce document est déjà dans {0} ({1}).'
  dupAsk='Ce document est déjà dans {0} ({1}).{2}Enregistrer quand même ?'
  reqName='Saisissez le nom du client.'
  saved='Client enregistré sous {0}.'; savedEdit='Modifications enregistrées dans {0}.'
  search='Rechercher (nom, NIE, ID, code)'; filterAll='Tous les états'; onlyDebt='Paiement en attente seulement'
  btnEdit='Modifier le client'; btnOpenExcel='Ouvrir Excel'; btnRefresh='Actualiser'
  colId='ID'; colName='Client'; colDoc='NIE / DNI'; colTipo='Permis'; colStatus='État'; colPay='Paiement'; colPend='Reste'; colDays='Jours'
  setTitle='Paramètres'; setFile='Fichier Excel des clients'; btnChange='Changer…'; setLang='Langue'
  setNat='Nationalité par défaut des nouveaux clients'; setBackup="Faire une copie de sécurité de l'Excel avant chaque enregistrement"
  setOpen="Ouvrir Excel après l'enregistrement d'un client"; btnBackups='Ouvrir le dossier des copies'; btnSaveSettings='Enregistrer les paramètres'
  setNote="Les prix et les listes (types de permis, étapes) se modifient dans l'onglet AJUSTES de l'Excel."
  settingsSaved='Paramètres enregistrés.'; chooseFile='Choisissez le fichier Excel des clients'
  welcome='Pour commencer, choisissez votre fichier Excel des clients.'
  loading="Lecture de l'Excel…"; savingMsg="Enregistrement dans l'Excel…"; ready='{0} clients · {1} avec paiement en attente ({2})'
  noFile='Choisissez votre fichier Excel dans Paramètres.'; tramite='Démarche'
  errNoFile='Fichier Excel introuvable. Choisissez-le dans Paramètres.'
  errNoSheet="Cet Excel n'a pas d'onglet « {0} ». Choisissez le bon fichier dans Paramètres."
  errLocked="L'Excel est ouvert sur un autre ordinateur ou programme. Fermez-le puis réessayez."
  errBusy="Excel est occupé : si vous écrivez dans une cellule, appuyez sur Entrée dans Excel puis réessayez."
  errFull="Il n'y a plus de lignes libres dans l'Excel."; errNoExcel='GESTEM MANAGER a besoin de Microsoft Excel installé sur cet ordinateur.'
  errGeneric="L'opération n'a pas pu aboutir :"; selectFirst='Sélectionnez un client dans la liste.'
  aboutText="Gestion des clients pour gestorías : échanges de permis de conduire et autres démarches. Tout est enregistré directement dans votre Excel."
  aboutBy='Créé par Yassine Belali'; aboutVer='Version {0}'
  partialOnly='(seulement si le paiement est Partiel)'; none='(aucun)'
  updating='Mise à jour de GESTEM vers la version {0}…'; updated='GESTEM mis à jour vers la version {0}.'; upToDate='Vous avez déjà la dernière version ({0}).'; updateFail="Impossible de chercher la mise à jour (internet ?)."; btnUpdate='Rechercher des mises à jour'
 }
 ar = @{
  navNew='عميل جديد'; navClients='العملاء'; navSettings='الإعدادات'; navAbout='حول البرنامج'
  secclient='بيانات العميل'; secproc='الإجراء'; secpay='الدفع'; secnotes='ملاحظات'
  newTitle='عميل جديد'; editTitle='تعديل العميل {0}'
  btnSave='حفظ العميل'; btnSaveEdit='حفظ التغييرات'; btnClear='مسح النموذج'; btnCancel='إلغاء'
  sumPrice='السعر'; sumPaid='المدفوع'; sumPend='المتبقي'; cardEmpty='اسم العميل'; cardNew='جديد'
  dupWarn='تنبيه: هذه الوثيقة موجودة في {0} ({1}).'
  dupAsk='هذه الوثيقة موجودة في {0} ({1}).{2}هل تريد الحفظ رغم ذلك؟'
  reqName='اكتب اسم العميل.'
  saved='تم حفظ العميل برقم {0}.'; savedEdit='تم حفظ التغييرات في {0}.'
  search='بحث (الاسم، NIE، الرقم، الرمز)'; filterAll='كل الحالات'; onlyDebt='الدفع المتبقي فقط'
  btnEdit='تعديل العميل'; btnOpenExcel='فتح Excel'; btnRefresh='تحديث'
  colId='الرقم'; colName='العميل'; colDoc='NIE / DNI'; colTipo='الرخصة'; colStatus='الحالة'; colPay='الدفع'; colPend='المتبقي'; colDays='الأيام'
  setTitle='الإعدادات'; setFile='ملف Excel للعملاء'; btnChange='تغيير…'; setLang='اللغة'
  setNat='الجنسية الافتراضية للعملاء الجدد'; setBackup='عمل نسخة احتياطية من ملف Excel قبل كل حفظ'
  setOpen='فتح Excel بعد حفظ العميل'; btnBackups='فتح مجلد النسخ'; btnSaveSettings='حفظ الإعدادات'
  setNote='تُغيَّر الأسعار والقوائم (أنواع الرخص، الخطوات) في تبويب AJUSTES داخل ملف Excel.'
  settingsSaved='تم حفظ الإعدادات.'; chooseFile='اختر ملف Excel للعملاء'
  welcome='للبدء، اختر ملف Excel الخاص بالعملاء.'
  loading='جارٍ قراءة ملف Excel…'; savingMsg='جارٍ الحفظ في Excel…'; ready='{0} عميل · {1} بدفع متبقٍ ({2})'
  noFile='اختر ملف Excel من الإعدادات.'; tramite='الإجراء'
  errNoFile='لم أجد ملف Excel. اختره من الإعدادات.'
  errNoSheet='ملف Excel هذا لا يحتوي على تبويب «{0}». اختر الملف الصحيح من الإعدادات.'
  errLocked='ملف Excel مفتوح على جهاز أو برنامج آخر. أغلقه هناك ثم حاول مرة أخرى.'
  errBusy='برنامج Excel مشغول: إذا كنت تكتب في خلية، اضغط Enter في Excel ثم حاول مرة أخرى.'
  errFull='لا توجد صفوف فارغة في ملف Excel.'; errNoExcel='يحتاج GESTEM MANAGER إلى برنامج Microsoft Excel مثبتًا على هذا الحاسوب.'
  errGeneric='تعذّر إتمام العملية:'; selectFirst='اختر عميلًا من القائمة.'
  aboutText='برنامج لإدارة عملاء مكاتب الخدمات الإدارية: استبدال رخص القيادة وإجراءات أخرى. يُحفظ كل شيء مباشرة في ملف Excel.'
  aboutBy='من إعداد ياسين بلعلي'; aboutVer='الإصدار {0}'
  partialOnly='(فقط عند الدفع الجزئي)'; none='(لا شيء)'
  updating='جارٍ تحديث GESTEM إلى الإصدار {0}…'; updated='تم تحديث GESTEM إلى الإصدار {0}.'; upToDate='لديك أحدث إصدار ({0}).'; updateFail='تعذّر البحث عن تحديث (هل يوجد اتصال بالإنترنت؟).'; btnUpdate='البحث عن تحديثات'
 }
}
function L([string]$key) {
    $t = $script:Text[$script:Cfg.Language]
    if ($t -and $t.ContainsKey($key)) { return $t[$key] }
    return $script:Text['es'][$key]
}
function Get-FieldLabel($field) { $l = $field.Label[$script:Cfg.Language]; if (-not $l) { $l = $field.Label['es'] }; return $l }

# =====================================================================
#  AJUSTES DEL PROGRAMA (config.json)
# =====================================================================
function Load-Config {
    $cfg = [ordered]@{ ExcelPath = ''; Language = 'es'; DefaultNationality = 'MARROQUI'; Backup = $true; OpenAfterSave = $false; Tramite = 'canje' }
    if (Test-Path $script:CfgPath) {
        try {
            $j = Get-Content -Raw -Path $script:CfgPath -Encoding UTF8 | ConvertFrom-Json
            foreach ($k in @($cfg.Keys)) { if ($null -ne $j.$k) { $cfg[$k] = $j.$k } }
        } catch { Write-Log "config ilegible: $_" }
    }
    if (-not $script:Langs.Contains($cfg.Language)) { $cfg.Language = 'es' }
    $script:Cfg = $cfg
}
function Save-Config { ($script:Cfg | ConvertTo-Json) | Set-Content -Path $script:CfgPath -Encoding UTF8 }

# =====================================================================
#  EXCEL
# =====================================================================
function Test-ExcelInstalled { return ($null -ne [Type]::GetTypeFromProgID('Excel.Application')) }

function Invoke-Excel([scriptblock]$Action) {
    # Excel rechaza llamadas si está ocupado (por ejemplo, editando una celda): reintentar un poco
    for ($i = 0; $i -lt 6; $i++) {
        try { return & $Action }
        catch {
            $hr = 0
            if ($_.Exception -is [Runtime.InteropServices.COMException]) { $hr = $_.Exception.HResult }
            elseif ($_.Exception.InnerException -is [Runtime.InteropServices.COMException]) { $hr = $_.Exception.InnerException.HResult }
            if (($hr -eq -2147418111 -or $hr -eq -2147417846) -and $i -lt 5) { Start-Sleep -Milliseconds 600; continue }
            if ($hr -eq -2147418111 -or $hr -eq -2147417846) { throw (L 'errBusy') }
            throw
        }
    }
}

function Open-Book([switch]$ForWrite) {
    $path = $script:Cfg.ExcelPath
    if (-not $path -or -not (Test-Path -LiteralPath $path)) { throw (L 'errNoFile') }
    $full = [IO.Path]::GetFullPath($path)
    $fname = [IO.Path]::GetFileName($full)
    $xl = $null; $wb = $null; $mine = $false
    try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { $xl = $null }
    if ($xl) {
        foreach ($w in $xl.Workbooks) {
            $fn = "$($w.FullName)"
            if ($fn -ieq $full -or (($fn -match '^https?:') -and $w.Name -ieq $fname)) { $wb = $w; break }
        }
    }
    if (-not $wb) {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $xl.ScreenUpdating = $false
        $wb = $xl.Workbooks.Open($full)
        $mine = $true
        if ($ForWrite -and $wb.ReadOnly) {
            $wb.Close($false); $xl.Quit()
            Release-Com $wb; Release-Com $xl
            throw (L 'errLocked')
        }
    }
    return [pscustomobject]@{ Xl = $xl; Wb = $wb; Mine = $mine }
}

function Release-Com($o) { if ($o) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) } catch {} } }

function Close-Book($h, [bool]$Save) {
    if (-not $h) { return }
    try {
        if ($Save) { $h.Wb.Save() }
        if ($h.Mine) { $h.Wb.Close($false); $h.Xl.Quit() }
    } finally {
        if ($h.Mine) { Release-Com $h.Wb; Release-Com $h.Xl }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    }
}

function Get-Sheet($h) {
    $T = $script:T
    try { return $h.Wb.Worksheets.Item($T.Sheet) } catch { throw ((L 'errNoSheet') -f $T.Sheet) }
}

function ColNum([string]$c) { $n = 0; foreach ($ch in $c.ToUpper().ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }; return $n }

function Read-ValidationList($ws, [string]$col) {
    # lee la lista desplegable que tiene el Excel en esa columna (rango de AJUSTES o lista escrita)
    $T = $script:T
    try { $f = "$($ws.Range("$col$($T.FirstRow)").Validation.Formula1)" } catch { return @() }
    if (-not $f) { return @() }
    $f = $f.TrimStart('=')
    if ($f -match '^(.+)!(\$?[A-Z]+\$?\d+(:\$?[A-Z]+\$?\d+)?)$') {
        $sheetName = $Matches[1].Trim("'")
        $addr = $Matches[2]
        $rng = $ws.Parent.Worksheets.Item($sheetName).Range($addr)
        $out = New-Object System.Collections.Generic.List[string]
        foreach ($c in $rng.Cells) { $v = "$($c.Value2)".Trim(); if ($v) { $out.Add($v) } }
        return ,$out.ToArray()
    }
    if ($f -match '^\$?[A-Z]+\$?\d+') { return @() }
    return ,@($f.Split(@(',', ';')) | ForEach-Object { $_.Trim().Trim('"') } | Where-Object { $_ })
}

function Read-Tariff($ws, [string]$col) {
    # precio por tipo de permiso: columnas siguientes a la lista de tipos en AJUSTES
    $T = $script:T
    $map = @{}
    try { $f = "$($ws.Range("$col$($T.FirstRow)").Validation.Formula1)".TrimStart('=') } catch { return $map }
    if ($f -match '^(.+)!(\$?[A-Z]+\$?\d+(:\$?[A-Z]+\$?\d+)?)$') {
        $rng = $ws.Parent.Worksheets.Item($Matches[1].Trim("'")).Range($Matches[2])
        foreach ($c in $rng.Cells) {
            $k = "$($c.Value2)".Trim()
            $p = $c.Offset(0, 1).Value2
            if ($k -and ($p -is [double])) { $map[$k] = [double]$p }
        }
    }
    return $map
}

function Load-Data {
    $T = $script:T
    Set-Status (L 'loading')
    $h = $null
    try {
        $h = Invoke-Excel { Open-Book }
        $ws = Get-Sheet $h
        $script:Lists = @{}
        foreach ($f in $T.Fields) {
            if ("$($f.List)" -like 'excel:*') {
                $c = $f.List.Substring(6)
                $script:Lists[$f.List] = Invoke-Excel { Read-ValidationList $ws $c }
            }
        }
        $script:Tariff = Invoke-Excel { Read-Tariff $ws $T.TipoCol }
        $addr = "A$($T.FirstRow):$($T.LastCol)$($T.LastRow)"
        $vals = Invoke-Excel { ,($ws.Range($addr).Value2) }
        $n = $T.LastRow - $T.FirstRow + 1
        $rows = New-Object System.Collections.Generic.List[object]
        $firstEmpty = $null; $nextId = $null
        $payExtra = New-Object System.Collections.Generic.HashSet[string]
        for ($i = 1; $i -le $n; $i++) {
            $name = "$($vals[$i, (ColNum $T.NameCol)])".Trim()
            if (-not $name) {
                if (-not $firstEmpty) { $firstEmpty = $T.FirstRow + $i - 1; $nextId = "$($vals[$i, (ColNum $T.IdCol)])" }
                continue
            }
            $rec = @{ Row = $T.FirstRow + $i - 1 }
            foreach ($f in $T.Fields) {
                $v = $vals[$i, (ColNum $f.Col)]
                if ($f.Type -eq 'Date' -and $v -is [double]) { $v = [DateTime]::FromOADate($v) }
                $rec[$f.Key] = $v
            }
            if ($rec['forma']) { [void]$payExtra.Add("$($rec['forma'])".Trim()) }
            $status = "$($rec['status'])"
            $pend = $vals[$i, (ColNum $T.PendCol)]
            $days = $vals[$i, (ColNum $T.DaysCol)]
            $color = '#FFFFFF'
            if ($T.StatusColors.ContainsKey($status)) { $color = $T.StatusColors[$status] }
            $rows.Add([pscustomobject]@{
                Row = $rec.Row; Id = "$($vals[$i, (ColNum $T.IdCol)])"; Name = $name; Doc = "$($rec['doc'])"
                Tipo = "$($rec['tipo'])"; Status = $status; Color = $color; Pay = "$($rec['paystate'])"
                Pending = $(if ($pend -is [double]) { $pend } else { 0 })
                PendingText = $(if ($pend -is [double] -and $pend -gt 0) { Format-Money $pend } else { '' })
                Days = $(if ($days -is [double]) { [int]$days } else { '' })
                Loc = "$($rec['loc'])"; Data = $rec
            })
        }
        $script:Rows = $rows
        $script:FirstEmpty = $firstEmpty
        $script:NextId = $nextId
        $all = New-Object System.Collections.Generic.List[string]
        foreach ($p in $script:FormasPago) { $all.Add($p) }
        foreach ($p in $payExtra) { if (-not $all.Contains($p) -and $p.Length -le 30) { $all.Add($p) } }
        $script:Lists['payments'] = $all.ToArray()
        $script:Lists['nationalities'] = $script:Nacionalidades
        Close-Book $h $false; $h = $null
        Refresh-ClientGrid
        Update-StatusLine
        return $true
    } catch {
        if ($h) { try { Close-Book $h $false } catch {} }
        Write-Log "Load-Data: $($_.Exception.Message)"
        Set-Status ($_.Exception.Message)
        Show-Error $_.Exception.Message
        return $false
    }
}

function Restore-Formula($ws, [string]$col, [int]$row) {
    # vuelve a poner la fórmula automática de esa columna (copiándola de una fila vacía)
    $T = $script:T
    for ($r = $T.LastRow; $r -ge $T.FirstRow; $r--) {
        if ($r -eq $row) { continue }
        $src = $ws.Range("$col$r")
        if ($src.HasFormula) { $ws.Range("$col$row").FormulaR1C1 = $src.FormulaR1C1; return }
    }
}

function Set-CellValue($ws, [string]$addr, $value, [string]$type) {
    $rng = $ws.Range($addr)
    if ($null -eq $value -or ("$value" -eq '')) { $rng.ClearContents() | Out-Null; return }
    switch ($type) {
        'Date'  { $rng.Value2 = ([DateTime]$value).ToOADate() }
        'Money' { $rng.Value2 = [double]$value }
        default { $rng.Value2 = "'" + "$value" }     # el apóstrofo obliga a guardarlo como texto
    }
}

function Save-Client([hashtable]$vals, $row) {
    $T = $script:T
    $h = $null
    try {
        $h = Invoke-Excel { Open-Book -ForWrite }
        $ws = Get-Sheet $h
        if ($script:Cfg.Backup) { Backup-Excel }
        if (-not $row) {
            $col = Invoke-Excel { ,($ws.Range("$($T.NameCol)$($T.FirstRow):$($T.NameCol)$($T.LastRow)").Value2) }
            $n = $T.LastRow - $T.FirstRow + 1
            for ($i = 1; $i -le $n; $i++) { if (-not "$($col[$i, 1])".Trim()) { $row = $T.FirstRow + $i - 1; break } }
            if (-not $row) { throw (L 'errFull') }
        }
        $tipo = "$($vals['tipo'])"
        Invoke-Excel {
            foreach ($f in $T.Fields) {
                $addr = "$($f.Col)$row"
                $v = $vals[$f.Key]
                switch ($f.Role) {
                    'nationality' {
                        if (-not "$v" -or "$v" -ieq $T.DefaultNationality) { Restore-Formula $ws $f.Col $row }
                        else { Set-CellValue $ws $addr "$v" 'Text' }
                    }
                    'price' {
                        $tar = $null
                        if ($tipo -and $script:Tariff.ContainsKey($tipo)) { $tar = $script:Tariff[$tipo] }
                        if ($null -eq $v -or ($null -ne $tar -and [double]$v -eq $tar)) { Restore-Formula $ws $f.Col $row }
                        else { Set-CellValue $ws $addr $v 'Money' }
                    }
                    'paid' {
                        if ("$($vals['paystate'])" -eq 'Parcial' -and $null -ne $v) { Set-CellValue $ws $addr $v 'Money' }
                        else { Restore-Formula $ws $f.Col $row }
                    }
                    default { Set-CellValue $ws $addr $v $f.Type }
                }
            }
        }
        $id = "$($ws.Range("$($T.IdCol)$row").Value2)"
        Invoke-Excel { Close-Book $h $true }; $h = $null
        Write-Log "guardado fila $row ($id)"
        return $id
    } catch {
        if ($h) { try { Close-Book $h $false } catch {} }
        throw
    }
}

function Backup-Excel {
    try {
        $src = $script:Cfg.ExcelPath
        $name = [IO.Path]::GetFileNameWithoutExtension($src)
        $dst = Join-Path $script:BackupDir ("{0}_{1:yyyyMMdd_HHmmss}.xlsx" -f $name, (Get-Date))
        Copy-Item -LiteralPath $src -Destination $dst -Force
        Get-ChildItem -Path $script:BackupDir -Filter '*.xlsx' | Sort-Object LastWriteTime -Descending | Select-Object -Skip 40 | Remove-Item -Force
    } catch { Write-Log "copia de seguridad fallida: $_" }
}

function Open-ExcelWindow([int]$row = 0) {
    $path = $script:Cfg.ExcelPath
    if (-not $path -or -not (Test-Path -LiteralPath $path)) { Show-Error (L 'errNoFile'); return }
    try {
        $xl = $null
        try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch {}
        if ($xl) {
            foreach ($w in $xl.Workbooks) {
                if ("$($w.FullName)" -ieq [IO.Path]::GetFullPath($path)) {
                    $xl.Visible = $true; $w.Activate()
                    if ($row -gt 0) { $w.Worksheets.Item($script:T.Sheet).Activate(); $xl.Goto($w.Worksheets.Item($script:T.Sheet).Range("B$row")) }
                    return
                }
            }
        }
    } catch {}
    Start-Process -FilePath $path
}

# =====================================================================
#  ACTUALIZACIÓN AUTOMÁTICA (desde GitHub)
# =====================================================================
function Check-Update([switch]$Manual) {
    $me = $MyInvocation.ScriptName
    if (-not $me) { $me = Join-Path $script:AppDir 'GESTEM.ps1' }
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $stamp = [DateTime]::UtcNow.Ticks
        $remote = (Invoke-WebRequest -UseBasicParsing -Uri "$($script:UpdateBase)/version.txt?t=$stamp" -TimeoutSec 5).Content
        $remote = "$remote".Trim()
        if (-not ($remote -match '^\d+(\.\d+){1,3}$')) { throw "version.txt no válido: $remote" }
        if ([version]$remote -le [version]$script:AppVersion) {
            if ($Manual) { Show-Info ((L 'upToDate') -f $script:AppVersion) }
            return
        }
        Set-Status ((L 'updating') -f $remote)
        $tmp = Join-Path $env:TEMP 'GESTEM_nuevo.ps1'
        Invoke-WebRequest -UseBasicParsing -Uri "$($script:UpdateBase)/GESTEM.ps1?t=$stamp" -OutFile $tmp -TimeoutSec 30
        # comprobar que el archivo descargado está completo y es correcto antes de reemplazar nada
        $code = Get-Content -Raw -LiteralPath $tmp -Encoding UTF8
        $errs = $null
        [void][System.Management.Automation.PSParser]::Tokenize($code, [ref]$errs)
        if ($errs.Count -gt 0 -or $code.Length -lt 20000 -or -not $code.Contains("AppVersion = '$remote'")) { throw 'la descarga no es válida' }
        Copy-Item -LiteralPath $me -Destination "$me.anterior" -Force
        Copy-Item -LiteralPath $tmp -Destination $me -Force
        Unblock-File -LiteralPath $me
        Set-Content -Path (Join-Path $script:DataDir 'actualizado.txt') -Value $remote -Encoding UTF8
        Write-Log "actualizado de $($script:AppVersion) a $remote"
        $ps = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
        Start-Process -FilePath $ps -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-WindowStyle', 'Hidden', '-File', "`"$me`"") -WindowStyle Hidden
        $script:Restarting = $true
        $script:Win.Close()
    } catch {
        Write-Log "actualización: $($_.Exception.Message)"
        if ($Manual) { Show-Error ((L 'updateFail') + [Environment]::NewLine + $_.Exception.Message) }
    }
}
# =====================================================================
#  UTILIDADES
# =====================================================================
function Format-Money($v) { if ($null -eq $v -or "$v" -eq '') { return '-' }; return ([double]$v).ToString('N2', $script:ES) + ' €' }
function Parse-Money([string]$s) {
    $s = "$s".Replace('€', '').Replace(' ', '').Trim()
    if (-not $s) { return $null }
    if ($s.Contains('.') -and $s.Contains(',')) { $s = $s.Replace('.', '').Replace(',', '.') } else { $s = $s.Replace(',', '.') }
    $d = 0.0
    if ([double]::TryParse($s, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
    return $null
}
function Get-List([string]$name) {
    # devuelve la lista sin huecos (nunca un elemento vacío)
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($x in @($script:Lists[$name])) { if ($null -ne $x -and "$x".Trim()) { $out.Add("$x") } }
    return ,$out.ToArray()
}
function Brush([string]$hex) { return (New-Object Windows.Media.BrushConverter).ConvertFromString($hex) }
function Show-Error([string]$msg) { [void][Windows.MessageBox]::Show($script:Win, $msg, $script:AppName, 'OK', 'Warning') }
function Show-Info([string]$msg) { [void][Windows.MessageBox]::Show($script:Win, $msg, $script:AppName, 'OK', 'Information') }
function Set-Status([string]$msg) {
    $script:UI.StatusText.Text = $msg
    $script:Win.Dispatcher.Invoke([Action]{}, [Windows.Threading.DispatcherPriority]::Background)
}
function Update-StatusLine {
    if (-not $script:Rows) { return }
    $debt = @($script:Rows | Where-Object { $_.Pending -gt 0 })
    $sum = ($debt | Measure-Object -Property Pending -Sum).Sum
    Set-Status ((L 'ready') -f $script:Rows.Count, $debt.Count, (Format-Money $sum))
}

# =====================================================================
#  VENTANA
# =====================================================================
$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="GESTEM MANAGER" Width="1260" Height="820" MinWidth="1020" MinHeight="660"
        WindowStartupLocation="CenterScreen" Background="#F2F5F9" FontFamily="Segoe UI" FontSize="13" Foreground="#14233A"
        UseLayoutRounding="True">
  <Window.Resources>
    <Style TargetType="TextBox">
      <Setter Property="Height" Value="34"/>
      <Setter Property="Padding" Value="8,0"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="BorderBrush" Value="#C5D0DE"/>
      <Setter Property="Background" Value="White"/>
    </Style>
    <Style TargetType="ComboBox">
      <Setter Property="Height" Value="34"/>
      <Setter Property="Padding" Value="8,0"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style TargetType="DatePicker">
      <Setter Property="Height" Value="34"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="BorderBrush" Value="#C5D0DE"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Margin" Value="0,6"/>
    </Style>
    <Style x:Key="FieldLabel" TargetType="TextBlock">
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Foreground" Value="#4A5B72"/>
      <Setter Property="Margin" Value="0,0,0,5"/>
    </Style>
    <Style x:Key="SectionTitle" TargetType="TextBlock">
      <Setter Property="FontFamily" Value="Bahnschrift, Segoe UI"/>
      <Setter Property="FontSize" Value="15"/>
      <Setter Property="Foreground" Value="#0F4C9A"/>
      <Setter Property="Margin" Value="0,10,0,10"/>
    </Style>
    <Style x:Key="Nav" TargetType="RadioButton">
      <Setter Property="Foreground" Value="#D8E6F7"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Margin" Value="3,0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RadioButton">
            <Border x:Name="b" Padding="14,8" CornerRadius="6" Background="Transparent">
              <ContentPresenter VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="b" Property="Background" Value="#1B5BAA"/>
              </Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="b" Property="Background" Value="#2C6CBE"/>
                <Setter Property="Foreground" Value="White"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Pill" TargetType="RadioButton">
      <Setter Property="Margin" Value="0,0,6,6"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Background" Value="#F4F6F9"/>
      <Setter Property="Foreground" Value="#14233A"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RadioButton">
            <Border x:Name="b" CornerRadius="6" Padding="12,6" Background="{TemplateBinding Background}" BorderBrush="#C5D0DE" BorderThickness="1">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="b" Property="BorderBrush" Value="#0F2C55"/>
                <Setter TargetName="b" Property="BorderThickness" Value="2"/>
                <Setter Property="FontWeight" Value="SemiBold"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="b" Property="BorderBrush" Value="#0F4C9A"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Primary" TargetType="Button">
      <Setter Property="Background" Value="#0F4C9A"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontSize" Value="15"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Height" Value="48"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" Background="{TemplateBinding Background}" CornerRadius="8">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="b" Property="Background" Value="#0B3D7E"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="b" Property="Background" Value="#0B3D7E"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="b" Property="Opacity" Value="0.55"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Btn" TargetType="Button">
      <Setter Property="Background" Value="White"/>
      <Setter Property="Foreground" Value="#14233A"/>
      <Setter Property="Height" Value="36"/>
      <Setter Property="Padding" Value="14,0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" Background="{TemplateBinding Background}" CornerRadius="7" BorderBrush="#C5D0DE" BorderThickness="1" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="b" Property="Background" Value="#EEF3F9"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="b" Property="BorderBrush" Value="#0F4C9A"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <DockPanel>
    <Border DockPanel.Dock="Top" Background="#0F4C9A" Height="62" Padding="20,0">
      <DockPanel>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <Border Width="36" Height="36" CornerRadius="8" Background="White">
            <TextBlock Text="G" FontFamily="Bahnschrift, Segoe UI" FontSize="24" FontWeight="Bold" Foreground="#0F4C9A" HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <StackPanel Margin="10,0,0,0" VerticalAlignment="Center">
            <TextBlock Text="GESTEM" FontFamily="Bahnschrift, Segoe UI" FontSize="20" FontWeight="SemiBold" Foreground="White"/>
            <TextBlock Text="MANAGER" FontFamily="Bahnschrift, Segoe UI" FontSize="10" Foreground="#B9D2F0" Margin="1,-3,0,0"/>
          </StackPanel>
          <TextBlock x:Name="TramiteLabel" Foreground="#B9D2F0" VerticalAlignment="Center" Margin="28,0,8,0"/>
          <ComboBox x:Name="TramiteBox" Width="250" VerticalAlignment="Center"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
          <RadioButton x:Name="NavNew" Style="{StaticResource Nav}" GroupName="nav" IsChecked="True"/>
          <RadioButton x:Name="NavClients" Style="{StaticResource Nav}" GroupName="nav"/>
          <RadioButton x:Name="NavSettings" Style="{StaticResource Nav}" GroupName="nav"/>
          <RadioButton x:Name="NavAbout" Style="{StaticResource Nav}" GroupName="nav"/>
        </StackPanel>
      </DockPanel>
    </Border>

    <Border DockPanel.Dock="Bottom" Background="#E3E9F1" Padding="20,7">
      <DockPanel>
        <TextBlock x:Name="FileText" DockPanel.Dock="Right" Foreground="#4A5B72"/>
        <TextBlock x:Name="StatusText" Foreground="#2B3B52"/>
      </DockPanel>
    </Border>

    <Grid Margin="20,18,20,18">

      <!-- NUEVO / EDITAR CLIENTE -->
      <Grid x:Name="PanelForm">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="330"/>
        </Grid.ColumnDefinitions>
        <Border Background="White" CornerRadius="10" BorderBrush="#DCE3EC" BorderThickness="1">
          <DockPanel>
            <TextBlock x:Name="FormTitle" DockPanel.Dock="Top" FontFamily="Bahnschrift, Segoe UI" FontSize="21" Margin="26,18,26,0"/>
            <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="0,0,6,0">
              <Grid x:Name="FormGrid" Margin="26,6,10,22">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
              </Grid>
            </ScrollViewer>
          </DockPanel>
        </Border>
        <StackPanel Grid.Column="1" Margin="16,0,0,0">
          <Border Background="White" CornerRadius="10" BorderBrush="#DCE3EC" BorderThickness="1">
            <StackPanel>
              <Border Background="#0F4C9A" CornerRadius="10,10,0,0" Padding="16,9">
                <DockPanel>
                  <TextBlock x:Name="CardTag" DockPanel.Dock="Right" Foreground="#B9D2F0" FontSize="12"/>
                  <TextBlock x:Name="CardId" Foreground="White" FontFamily="Bahnschrift, Segoe UI" FontSize="15"/>
                </DockPanel>
              </Border>
              <StackPanel Margin="16,12,16,16">
                <TextBlock x:Name="CardName" FontFamily="Bahnschrift, Segoe UI" FontSize="19" TextWrapping="Wrap"/>
                <TextBlock x:Name="CardSub" Foreground="#4A5B72" Margin="0,4,0,0" TextWrapping="Wrap"/>
                <TextBlock x:Name="CardTipo" Foreground="#4A5B72" Margin="0,2,0,0"/>
                <Border x:Name="CardStatusB" CornerRadius="6" Padding="10,4" Margin="0,10,0,0" HorizontalAlignment="Left">
                  <TextBlock x:Name="CardStatus" FontWeight="SemiBold" Foreground="#2A2A2A"/>
                </Border>
              </StackPanel>
            </StackPanel>
          </Border>
          <Border Background="White" CornerRadius="10" BorderBrush="#DCE3EC" BorderThickness="1" Margin="0,12,0,0" Padding="16,12">
            <Grid>
              <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
              <Grid.RowDefinitions><RowDefinition/><RowDefinition/><RowDefinition/></Grid.RowDefinitions>
              <TextBlock x:Name="SumPriceL" Foreground="#4A5B72" Margin="0,3"/>
              <TextBlock x:Name="SumPrice" Grid.Column="1" FontFamily="Bahnschrift, Segoe UI" FontSize="15" Margin="0,3"/>
              <TextBlock x:Name="SumPaidL" Grid.Row="1" Foreground="#4A5B72" Margin="0,3"/>
              <TextBlock x:Name="SumPaid" Grid.Row="1" Grid.Column="1" FontFamily="Bahnschrift, Segoe UI" FontSize="15" Margin="0,3"/>
              <TextBlock x:Name="SumPendL" Grid.Row="2" FontWeight="SemiBold" Margin="0,3"/>
              <TextBlock x:Name="SumPend" Grid.Row="2" Grid.Column="1" FontFamily="Bahnschrift, Segoe UI" FontSize="17" FontWeight="SemiBold" Margin="0,3"/>
            </Grid>
          </Border>
          <TextBlock x:Name="DupWarn" Foreground="#B42318" TextWrapping="Wrap" Margin="2,10,2,0" Visibility="Collapsed"/>
          <Button x:Name="SaveBtn" Style="{StaticResource Primary}" Margin="0,14,0,0"/>
          <Button x:Name="ClearBtn" Style="{StaticResource Btn}" Margin="0,8,0,0"/>
        </StackPanel>
      </Grid>

      <!-- CLIENTES -->
      <Border x:Name="PanelClients" Visibility="Collapsed" Background="White" CornerRadius="10" BorderBrush="#DCE3EC" BorderThickness="1" Padding="18">
        <DockPanel>
          <DockPanel DockPanel.Dock="Top" Margin="0,0,0,12">
            <StackPanel Orientation="Horizontal" DockPanel.Dock="Right">
              <Button x:Name="RefreshBtn" Style="{StaticResource Btn}" Margin="8,0,0,0"/>
              <Button x:Name="OpenExcelBtn" Style="{StaticResource Btn}" Margin="8,0,0,0"/>
              <Button x:Name="EditBtn" Style="{StaticResource Btn}" Margin="8,0,0,0" Foreground="#0F4C9A" FontWeight="SemiBold"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal">
              <Grid Width="330">
                <TextBox x:Name="SearchBox"/>
                <TextBlock x:Name="SearchHint" IsHitTestVisible="False" Foreground="#8A97A8" Margin="10,0,0,0" VerticalAlignment="Center"/>
              </Grid>
              <ComboBox x:Name="FilterBox" Width="200" Margin="10,0,0,0"/>
              <CheckBox x:Name="OnlyDebt" Margin="14,0,0,0"/>
            </StackPanel>
          </DockPanel>
          <DataGrid x:Name="ClientsGrid" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single"
                    HeadersVisibility="Column" GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="#E8EDF3"
                    RowHeight="34" Background="White" BorderBrush="#DCE3EC" CanUserAddRows="False" CanUserResizeRows="False">
            <DataGrid.ColumnHeaderStyle>
              <Style TargetType="DataGridColumnHeader">
                <Setter Property="Background" Value="#F2F5F9"/>
                <Setter Property="Padding" Value="8,8"/>
                <Setter Property="FontWeight" Value="SemiBold"/>
                <Setter Property="BorderBrush" Value="#DCE3EC"/>
                <Setter Property="BorderThickness" Value="0,0,0,1"/>
              </Style>
            </DataGrid.ColumnHeaderStyle>
            <DataGrid.CellStyle>
              <Style TargetType="DataGridCell">
                <Setter Property="BorderThickness" Value="0"/>
                <Setter Property="Padding" Value="6,0"/>
                <Setter Property="VerticalContentAlignment" Value="Center"/>
              </Style>
            </DataGrid.CellStyle>
            <DataGrid.Columns>
              <DataGridTextColumn Binding="{Binding Id}" Width="120"/>
              <DataGridTextColumn Binding="{Binding Name}" Width="2*"/>
              <DataGridTextColumn Binding="{Binding Doc}" Width="110"/>
              <DataGridTextColumn Binding="{Binding Tipo}" Width="130"/>
              <DataGridTemplateColumn Width="170">
                <DataGridTemplateColumn.CellTemplate>
                  <DataTemplate>
                    <Border Background="{Binding Color}" CornerRadius="5" Padding="8,2" Margin="2,4" HorizontalAlignment="Left">
                      <TextBlock Text="{Binding Status}" Foreground="#2A2A2A"/>
                    </Border>
                  </DataTemplate>
                </DataGridTemplateColumn.CellTemplate>
              </DataGridTemplateColumn>
              <DataGridTextColumn Binding="{Binding Pay}" Width="95"/>
              <DataGridTextColumn Binding="{Binding PendingText}" Width="105">
                <DataGridTextColumn.ElementStyle>
                  <Style TargetType="TextBlock">
                    <Setter Property="Foreground" Value="#B42318"/>
                    <Setter Property="FontWeight" Value="SemiBold"/>
                    <Setter Property="VerticalAlignment" Value="Center"/>
                  </Style>
                </DataGridTextColumn.ElementStyle>
              </DataGridTextColumn>
              <DataGridTextColumn Binding="{Binding Days}" Width="60"/>
            </DataGrid.Columns>
          </DataGrid>
        </DockPanel>
      </Border>

      <!-- AJUSTES -->
      <Border x:Name="PanelSettings" Visibility="Collapsed" Background="White" CornerRadius="10" BorderBrush="#DCE3EC" BorderThickness="1" Padding="28,22" HorizontalAlignment="Left" Width="720" VerticalAlignment="Top">
        <StackPanel>
          <TextBlock x:Name="SetTitle" FontFamily="Bahnschrift, Segoe UI" FontSize="21" Margin="0,0,0,14"/>
          <TextBlock x:Name="WelcomeText" Foreground="#0F4C9A" FontWeight="SemiBold" Margin="0,0,0,12" TextWrapping="Wrap" Visibility="Collapsed"/>
          <TextBlock x:Name="SetFileL" Style="{StaticResource FieldLabel}"/>
          <DockPanel>
            <Button x:Name="ChangeFileBtn" Style="{StaticResource Btn}" DockPanel.Dock="Right" Margin="8,0,0,0"/>
            <TextBox x:Name="FilePathBox" IsReadOnly="True" Background="#F7F9FB"/>
          </DockPanel>
          <Grid Margin="0,16,0,0">
            <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition Width="16"/><ColumnDefinition/></Grid.ColumnDefinitions>
            <StackPanel>
              <TextBlock x:Name="SetLangL" Style="{StaticResource FieldLabel}"/>
              <ComboBox x:Name="LangBox"/>
            </StackPanel>
            <StackPanel Grid.Column="2">
              <TextBlock x:Name="SetNatL" Style="{StaticResource FieldLabel}"/>
              <ComboBox x:Name="NatBox" IsEditable="True"/>
            </StackPanel>
          </Grid>
          <CheckBox x:Name="BackupChk" Margin="0,18,0,4"/>
          <CheckBox x:Name="OpenAfterChk"/>
          <TextBlock x:Name="SetNote" Foreground="#4A5B72" TextWrapping="Wrap" Margin="0,14,0,0"/>
          <StackPanel Orientation="Horizontal" Margin="0,20,0,0">
            <Button x:Name="SaveSettingsBtn" Style="{StaticResource Primary}" Height="40" Padding="22,0" FontSize="14"/>
            <Button x:Name="BackupsBtn" Style="{StaticResource Btn}" Margin="10,0,0,0" Height="40"/>
            <Button x:Name="UpdateBtn" Style="{StaticResource Btn}" Margin="10,0,0,0" Height="40"/>
          </StackPanel>
        </StackPanel>
      </Border>

      <!-- ACERCA DE -->
      <Border x:Name="PanelAbout" Visibility="Collapsed" Background="White" CornerRadius="10" BorderBrush="#DCE3EC" BorderThickness="1" Padding="36,30" HorizontalAlignment="Left" Width="620" VerticalAlignment="Top">
        <StackPanel>
          <StackPanel Orientation="Horizontal">
            <Border x:Name="AboutLogo" Width="64" Height="64" CornerRadius="12" Background="#0F4C9A" Cursor="Hand">
              <TextBlock Text="G" FontFamily="Bahnschrift, Segoe UI" FontSize="40" FontWeight="Bold" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Margin="16,0,0,0" VerticalAlignment="Center">
              <TextBlock Text="GESTEM MANAGER" FontFamily="Bahnschrift, Segoe UI" FontSize="26" FontWeight="SemiBold" Foreground="#0F4C9A"/>
              <TextBlock x:Name="AboutVer" Foreground="#4A5B72"/>
            </StackPanel>
          </StackPanel>
          <TextBlock x:Name="AboutText" TextWrapping="Wrap" Margin="0,22,0,0" FontSize="14" LineHeight="21"/>
          <TextBlock x:Name="AboutBy" FontFamily="Bahnschrift, Segoe UI" FontSize="16" Margin="0,18,0,0"/>
          <TextBlock Text="© 2026" Foreground="#8A97A8" Margin="0,4,0,0"/>
          <TextBlock x:Name="EasterEgg" Text="Por Safae de Yan ♥" FontFamily="Bahnschrift, Segoe UI" FontSize="20" Foreground="#C2185B" Margin="0,22,0,0" Visibility="Collapsed"/>
        </StackPanel>
      </Border>

    </Grid>
  </DockPanel>
</Window>
'@

# =====================================================================
#  CONSTRUIR EL FORMULARIO (a partir de la definición del trámite)
# =====================================================================
function New-FieldBox($field, $control) {
    $sp = New-Object Windows.Controls.StackPanel
    $sp.Margin = '0,0,16,14'
    $lbl = New-Object Windows.Controls.TextBlock
    $lbl.Text = (Get-FieldLabel $field)
    $lbl.Style = $script:Win.FindResource('FieldLabel')
    [void]$sp.Children.Add($lbl)
    [void]$sp.Children.Add($control)
    return $sp
}

function Add-ToGrid($el, [int]$col, [int]$span) {
    $g = $script:UI.FormGrid
    if ($col -eq 0) { $rd = New-Object Windows.Controls.RowDefinition; $rd.Height = [Windows.GridLength]::Auto; $g.RowDefinitions.Add($rd) }
    $r = [Math]::Max(0, $g.RowDefinitions.Count - 1)
    [Windows.Controls.Grid]::SetRow($el, $r)
    [Windows.Controls.Grid]::SetColumn($el, $col)
    [Windows.Controls.Grid]::SetColumnSpan($el, $span)
    [void]$g.Children.Add($el)
}

function Build-Form {
    $T = $script:T
    $g = $script:UI.FormGrid
    $g.Children.Clear(); $g.RowDefinitions.Clear()
    $script:Ctl = @{}
    $script:PillGroups = @{}
    $section = ''; $col = 0
    foreach ($f in $T.Fields) {
        if ($f.Section -ne $section) {
            $section = $f.Section
            $secTitle = New-Object Windows.Controls.TextBlock
            $secTitle.Text = (L ("sec" + $section))
            $secTitle.Style = $script:Win.FindResource('SectionTitle')
            Add-ToGrid $secTitle 0 2; $col = 0
        }
        $span = 1; if ($f.Span -eq 2) { $span = 2 }
        switch ($f.Type) {
            'Combo' {
                $c = New-Object Windows.Controls.ComboBox
                [void]$c.Items.Add('')
                foreach ($it in (Get-List $f.List)) { [void]$c.Items.Add($it) }
                $c.Add_SelectionChanged({ On-FieldChanged $this })
            }
            'ComboEdit' {
                $c = New-Object Windows.Controls.ComboBox
                $c.IsEditable = $true
                $c.IsTextSearchEnabled = $true
                foreach ($it in (Get-List $f.List)) { [void]$c.Items.Add($it) }
                $c.Add_SelectionChanged({ On-FieldChanged $this })
                $c.AddHandler([Windows.Controls.Primitives.TextBoxBase]::TextChangedEvent, [Windows.Controls.TextChangedEventHandler]{ On-FieldChanged $args[0] })
            }
            'Date' {
                $c = New-Object Windows.Controls.DatePicker
                $c.SelectedDateFormat = 'Short'
                $c.Language = [Windows.Markup.XmlLanguage]::GetLanguage('es-ES')
                $c.Add_SelectedDateChanged({ On-FieldChanged $this })
            }
            'Pills' {
                $c = New-Object Windows.Controls.WrapPanel
                $grp = "g_" + $f.Key
                $list = Get-List $f.List
                if ($list.Count -eq 0 -and $f.Key -eq 'paystate') { $list = @('Pagado', 'Parcial', 'Pendiente') }
                foreach ($it in $list) {
                    $rb = New-Object Windows.Controls.RadioButton
                    $rb.Style = $script:Win.FindResource('Pill')
                    $rb.GroupName = $grp
                    $rb.Content = $it
                    $rb.Tag = $f.Key
                    if ($f.Colored -and $T.StatusColors.ContainsKey($it)) { $rb.Background = Brush $T.StatusColors[$it] }
                    $rb.Add_Checked({ On-FieldChanged $this })
                    [void]$c.Children.Add($rb)
                }
                $script:PillGroups[$f.Key] = $c
            }
            'Multiline' {
                $c = New-Object Windows.Controls.TextBox
                $c.Height = 80; $c.AcceptsReturn = $true; $c.TextWrapping = 'Wrap'
                $c.VerticalContentAlignment = 'Top'; $c.Padding = '8,6'
                $c.Add_TextChanged({ On-FieldChanged $this })
            }
            default {
                $c = New-Object Windows.Controls.TextBox
                if ($f.Type -eq 'Upper') { $c.CharacterCasing = 'Upper' }
                $c.Add_TextChanged({ On-FieldChanged $this })
                if ($f.Key -eq 'doc') { $c.Add_LostFocus({ Check-Duplicate }) }
            }
        }
        $c.Tag = $f.Key
        $script:Ctl[$f.Key] = $c
        $box = New-FieldBox $f $c
        if ($span -eq 2) { if ($col -eq 1) { $col = 0 }; Add-ToGrid $box 0 2; $col = 0 }
        else { Add-ToGrid $box $col 1; $col = 1 - $col }
    }
}

function Get-FieldValue($f) {
    $c = $script:Ctl[$f.Key]
    switch ($f.Type) {
        'Combo'     { return "$($c.SelectedItem)".Trim() }
        'ComboEdit' { return "$($c.Text)".Trim().ToUpper() }
        'Date'      { if ($c.SelectedDate) { return [DateTime]$c.SelectedDate } else { return $null } }
        'Money'     { return (Parse-Money $c.Text) }
        'Pills'     { foreach ($rb in $script:PillGroups[$f.Key].Children) { if ($rb.IsChecked) { return "$($rb.Content)" } }; return '' }
        default     { return "$($c.Text)".Trim() }
    }
}

function Set-FieldValue($f, $v) {
    $c = $script:Ctl[$f.Key]
    switch ($f.Type) {
        'Combo' {
            $s = "$v".Trim()
            if ($s -and -not $c.Items.Contains($s)) { [void]$c.Items.Add($s) }
            $c.SelectedItem = $s
        }
        'ComboEdit' { $c.Text = "$v".Trim() }
        'Date' {
            if ($v -is [DateTime]) { $c.SelectedDate = $v }
            elseif ($v -is [double]) { $c.SelectedDate = [DateTime]::FromOADate($v) }
            else { $c.SelectedDate = $null }
        }
        'Money' { if ($v -is [double]) { $c.Text = ([double]$v).ToString('0.##', $script:ES) } else { $c.Text = '' } }
        'Pills' { foreach ($rb in $script:PillGroups[$f.Key].Children) { $rb.IsChecked = ("$($rb.Content)" -eq "$v") } }
        default { $c.Text = "$v".Trim() }
    }
}

function Get-FormValues {
    $vals = @{}
    foreach ($f in $script:T.Fields) { $vals[$f.Key] = Get-FieldValue $f }
    return $vals
}

function Reset-Form {
    $script:Filling = $true
    $script:EditRow = $null; $script:PriceTouched = $false
    foreach ($f in $script:T.Fields) {
        $v = $null
        if ($f.Role -eq 'nationality') { $v = $script:Cfg.DefaultNationality }
        elseif ($f.DefaultToday) { $v = (Get-Date).Date }
        elseif ($f.Default) { $v = $f.Default }
        Set-FieldValue $f $v
    }
    $script:Filling = $false
    $script:UI.FormTitle.Text = (L 'newTitle')
    $script:UI.SaveBtn.Content = (L 'btnSave')
    $script:UI.ClearBtn.Content = (L 'btnClear')
    $script:UI.DupWarn.Visibility = 'Collapsed'
    Update-Preview
    $script:Ctl['name'].Focus() | Out-Null
}

function Fill-FormForEdit($rowObj) {
    $script:Filling = $true
    $script:EditRow = $rowObj
    $script:PriceTouched = $true
    foreach ($f in $script:T.Fields) { Set-FieldValue $f $rowObj.Data[$f.Key] }
    $script:Filling = $false
    $script:UI.FormTitle.Text = ((L 'editTitle') -f $rowObj.Id)
    $script:UI.SaveBtn.Content = (L 'btnSaveEdit')
    $script:UI.ClearBtn.Content = (L 'btnCancel')
    $script:UI.DupWarn.Visibility = 'Collapsed'
    Update-Preview
}

function On-FieldChanged($sender) {
    if ($script:Filling) { return }
    $key = "$($sender.Tag)"
    if ($key -eq 'price') { $script:PriceTouched = $true }
    if ($key -eq 'tipo' -and -not $script:PriceTouched) {
        $tipo = Get-FieldValue ($script:T.Fields | Where-Object { $_.Key -eq 'tipo' })
        $script:Filling = $true
        if ($tipo -and $script:Tariff.ContainsKey($tipo)) { $script:Ctl['price'].Text = ([double]$script:Tariff[$tipo]).ToString('0.##', $script:ES) }
        elseif (-not $tipo) { $script:Ctl['price'].Text = '' }
        $script:Filling = $false
    }
    Update-Preview
}

function Update-Preview {
    $T = $script:T
    $v = Get-FormValues
    $ui = $script:UI
    if ($script:EditRow) { $ui.CardId.Text = $script:EditRow.Id; $ui.CardTag.Text = '' }
    else { $ui.CardId.Text = $(if ($script:NextId) { $script:NextId } else { '—' }); $ui.CardTag.Text = (L 'cardNew') }
    $ui.CardName.Text = $(if ($v.name) { $v.name.ToUpper() } else { L 'cardEmpty' })
    $ui.CardName.Foreground = $(if ($v.name) { Brush '#14233A' } else { Brush '#9AA7B8' })
    $ui.CardSub.Text = (@($v.doc, $v.nat) | Where-Object { $_ }) -join '  ·  '
    $ui.CardTipo.Text = "$($v.tipo)"
    if ($v.status) {
        $ui.CardStatus.Text = $v.status
        $ui.CardStatusB.Background = $(if ($T.StatusColors.ContainsKey($v.status)) { Brush $T.StatusColors[$v.status] } else { Brush '#E8EDF3' })
        $ui.CardStatusB.Visibility = 'Visible'
    } else { $ui.CardStatusB.Visibility = 'Collapsed' }
    # pago
    $price = $v.price
    $paidBox = $script:Ctl['paid']
    $paidBox.IsEnabled = ($v.paystate -eq 'Parcial')
    switch ($v.paystate) {
        'Pagado'  { $paid = $price }
        'Parcial' { $paid = $v.paid }
        default   { $paid = 0 }
    }
    $pend = $null
    if ($null -ne $price) { $pend = [Math]::Max(0, [double]$price - [double]$(if ($null -ne $paid) { $paid } else { 0 })) }
    $ui.SumPrice.Text = Format-Money $price
    $ui.SumPaid.Text = Format-Money $(if ($null -ne $price) { $(if ($null -ne $paid) { $paid } else { 0 }) } else { $null })
    $ui.SumPend.Text = Format-Money $pend
    $ui.SumPend.Foreground = $(if ($pend -gt 0) { Brush '#B42318' } else { Brush '#1E7B34' })
    $ui.SumPendL.Foreground = $ui.SumPend.Foreground
}

function Find-DuplicateDoc([string]$doc) {
    if (-not $doc -or -not $script:Rows) { return $null }
    foreach ($r in $script:Rows) {
        if ($script:EditRow -and $r.Row -eq $script:EditRow.Row) { continue }
        if ("$($r.Doc)".Trim() -ieq $doc) { return $r }
    }
    return $null
}

function Check-Duplicate {
    $d = Find-DuplicateDoc ("$($script:Ctl['doc'].Text)".Trim())
    if ($d) { $script:UI.DupWarn.Text = ((L 'dupWarn') -f $d.Id, $d.Name); $script:UI.DupWarn.Visibility = 'Visible' }
    else { $script:UI.DupWarn.Visibility = 'Collapsed' }
}

function Save-FromForm {
    Write-Log 'Save-FromForm inicio'
    $vals = Get-FormValues
    if (-not $vals.name) { Show-Error (L 'reqName'); $script:Ctl['name'].Focus() | Out-Null; return }
    $dup = Find-DuplicateDoc "$($vals.doc)"
    if ($dup) {
        $nl = [Environment]::NewLine + [Environment]::NewLine
        $ans = [Windows.MessageBox]::Show($script:Win, ((L 'dupAsk') -f $dup.Id, $dup.Name, $nl), $script:AppName, 'YesNo', 'Question')
        if ("$ans" -ne 'Yes') { return }
    }
    $ui = $script:UI
    $ui.SaveBtn.IsEnabled = $false
    $script:Win.Cursor = [Windows.Input.Cursors]::Wait
    Set-Status (L 'savingMsg')
    try {
        $row = $null; if ($script:EditRow) { $row = $script:EditRow.Row }
        $id = Save-Client $vals $row
        $wasEdit = [bool]$script:EditRow
        [void](Load-Data)
        Set-Status ($(if ($wasEdit) { (L 'savedEdit') -f $id } else { (L 'saved') -f $id }))
        Reset-Form
        if ($wasEdit) { $ui.NavClients.IsChecked = $true }
        if ($script:Cfg.OpenAfterSave) { Open-ExcelWindow }
    } catch {
        Write-Log "Save-FromForm: $($_.Exception.Message)"
        Show-Error ((L 'errGeneric') + [Environment]::NewLine + $_.Exception.Message)
        Set-Status ''
    } finally {
        $ui.SaveBtn.IsEnabled = $true
        $script:Win.Cursor = $null
    }
}

# =====================================================================
#  LISTA DE CLIENTES
# =====================================================================
function Refresh-ClientGrid {
    $ui = $script:UI
    if (-not $ui -or -not $script:Rows) { return }
    $q = "$($ui.SearchBox.Text)".Trim().ToLower()
    $st = "$($ui.FilterBox.SelectedItem)"
    $onlyDebt = [bool]$ui.OnlyDebt.IsChecked
    $list = foreach ($r in $script:Rows) {
        if ($q -and -not ("$($r.Name) $($r.Doc) $($r.Id) $($r.Loc)".ToLower().Contains($q))) { continue }
        if ($ui.FilterBox.SelectedIndex -gt 0 -and $r.Status -ne $st) { continue }
        if ($onlyDebt -and $r.Pending -le 0) { continue }
        $r
    }
    $ui.ClientsGrid.ItemsSource = @($list)
}

function Edit-Selected {
    $sel = $script:UI.ClientsGrid.SelectedItem
    if (-not $sel) { Show-Info (L 'selectFirst'); return }
    $script:UI.NavNew.IsChecked = $true
    Fill-FormForEdit $sel
}

# =====================================================================
#  IDIOMA Y NAVEGACIÓN
# =====================================================================
function Apply-Language {
    $ui = $script:UI
    $script:Win.FlowDirection = $(if ($script:Cfg.Language -eq 'ar') { 'RightToLeft' } else { 'LeftToRight' })
    $ui.NavNew.Content = L 'navNew'; $ui.NavClients.Content = L 'navClients'
    $ui.NavSettings.Content = L 'navSettings'; $ui.NavAbout.Content = L 'navAbout'
    $ui.TramiteLabel.Text = L 'tramite'
    $ui.SumPriceL.Text = L 'sumPrice'; $ui.SumPaidL.Text = L 'sumPaid'; $ui.SumPendL.Text = L 'sumPend'
    $ui.SearchHint.Text = L 'search'; $ui.OnlyDebt.Content = L 'onlyDebt'
    $ui.RefreshBtn.Content = L 'btnRefresh'; $ui.OpenExcelBtn.Content = L 'btnOpenExcel'; $ui.EditBtn.Content = L 'btnEdit'
    $cols = $ui.ClientsGrid.Columns
    $heads = @('colId', 'colName', 'colDoc', 'colTipo', 'colStatus', 'colPay', 'colPend', 'colDays')
    for ($i = 0; $i -lt $heads.Count; $i++) { $cols[$i].Header = L $heads[$i] }
    $ui.SetTitle.Text = L 'setTitle'; $ui.SetFileL.Text = L 'setFile'; $ui.ChangeFileBtn.Content = L 'btnChange'
    $ui.SetLangL.Text = L 'setLang'; $ui.SetNatL.Text = L 'setNat'; $ui.BackupChk.Content = L 'setBackup'
    $ui.OpenAfterChk.Content = L 'setOpen'; $ui.SetNote.Text = L 'setNote'
    $ui.SaveSettingsBtn.Content = L 'btnSaveSettings'; $ui.BackupsBtn.Content = L 'btnBackups'; $ui.UpdateBtn.Content = L 'btnUpdate'
    $ui.WelcomeText.Text = L 'welcome'
    $ui.AboutText.Text = L 'aboutText'; $ui.AboutBy.Text = L 'aboutBy'; $ui.AboutVer.Text = ((L 'aboutVer') -f $script:AppVersion)
    # filtros de estado
    $ui.FilterBox.Items.Clear()
    [void]$ui.FilterBox.Items.Add((L 'filterAll'))
    foreach ($s in (Get-List 'excel:N')) { [void]$ui.FilterBox.Items.Add($s) }
    $ui.FilterBox.SelectedIndex = 0
    # trámites
    $ui.TramiteBox.Items.Clear()
    foreach ($t in $script:Tramites) {
        $title = $t.Title[$script:Cfg.Language]; if (-not $title) { $title = $t.Title['es'] }
        [void]$ui.TramiteBox.Items.Add($title)
    }
    $ui.TramiteBox.SelectedIndex = [Math]::Max(0, [array]::IndexOf(@($script:Tramites | ForEach-Object { $_.Id }), $script:T.Id))
    $ui.FileText.Text = $(if ($script:Cfg.ExcelPath) { [IO.Path]::GetFileName($script:Cfg.ExcelPath) } else { '' })
    Build-Form
    Reset-Form
}

function Show-Panel([string]$name) {
    foreach ($p in @('PanelForm', 'PanelClients', 'PanelSettings', 'PanelAbout')) {
        $script:UI[$p].Visibility = $(if ($p -eq $name) { 'Visible' } else { 'Collapsed' })
    }
}

function Load-SettingsPanel {
    $ui = $script:UI
    $ui.FilePathBox.Text = $script:Cfg.ExcelPath
    $ui.LangBox.Items.Clear()
    foreach ($k in $script:Langs.Keys) { [void]$ui.LangBox.Items.Add($script:Langs[$k]) }
    $ui.LangBox.SelectedIndex = [array]::IndexOf(@($script:Langs.Keys), $script:Cfg.Language)
    $ui.NatBox.Items.Clear()
    foreach ($n in $script:Nacionalidades) { [void]$ui.NatBox.Items.Add($n) }
    $ui.NatBox.Text = $script:Cfg.DefaultNationality
    $ui.BackupChk.IsChecked = [bool]$script:Cfg.Backup
    $ui.OpenAfterChk.IsChecked = [bool]$script:Cfg.OpenAfterSave
    $ui.WelcomeText.Visibility = $(if (-not $script:Cfg.ExcelPath) { 'Visible' } else { 'Collapsed' })
}

function Choose-ExcelFile {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Title = L 'chooseFile'
    $dlg.Filter = 'Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm'
    if ($script:Cfg.ExcelPath) { try { $dlg.InitialDirectory = Split-Path -Parent $script:Cfg.ExcelPath } catch {} }
    else { $dlg.InitialDirectory = [Environment]::GetFolderPath('MyDocuments') }
    if ($dlg.ShowDialog($script:Win)) { $script:UI.FilePathBox.Text = $dlg.FileName; return $true }
    return $false
}

function Save-Settings {
    $ui = $script:UI
    $oldPath = $script:Cfg.ExcelPath; $oldLang = $script:Cfg.Language
    $script:Cfg.ExcelPath = $ui.FilePathBox.Text
    if ($ui.LangBox.SelectedIndex -ge 0) { $script:Cfg.Language = @($script:Langs.Keys)[$ui.LangBox.SelectedIndex] }
    $nat = "$($ui.NatBox.Text)".Trim().ToUpper(); if ($nat) { $script:Cfg.DefaultNationality = $nat }
    $script:Cfg.Backup = [bool]$ui.BackupChk.IsChecked
    $script:Cfg.OpenAfterSave = [bool]$ui.OpenAfterChk.IsChecked
    Save-Config
    if ($script:Cfg.ExcelPath -ne $oldPath) { [void](Load-Data) }
    Apply-Language
    Load-SettingsPanel
    Set-Status (L 'settingsSaved')
}

# =====================================================================
#  INICIO
# =====================================================================
Load-Config
$script:T = $script:Tramites | Where-Object { $_.Id -eq $script:Cfg.Tramite } | Select-Object -First 1
if (-not $script:T) { $script:T = $script:Tramites[0] }
$script:Lists = @{ 'nationalities' = $script:Nacionalidades; 'payments' = $script:FormasPago }
$script:Tariff = @{}
$script:Rows = $null
$script:Filling = $false

$script:Win = [Windows.Markup.XamlReader]::Parse($xaml)
$iconPath = Join-Path $script:AppDir 'gestem.ico'
if (Test-Path $iconPath) {
    try {   # se lee en memoria para no bloquear el archivo
        $bmp = New-Object Windows.Media.Imaging.BitmapImage
        $bmp.BeginInit(); $bmp.CacheOption = 'OnLoad'; $bmp.UriSource = [Uri]$iconPath; $bmp.EndInit()
        $script:Win.Icon = $bmp
    } catch {}
}

$script:UI = @{}
foreach ($n in @('TramiteLabel','TramiteBox','NavNew','NavClients','NavSettings','NavAbout','StatusText','FileText',
                 'PanelForm','FormTitle','FormGrid','CardTag','CardId','CardName','CardSub','CardTipo','CardStatusB','CardStatus',
                 'SumPriceL','SumPrice','SumPaidL','SumPaid','SumPendL','SumPend','DupWarn','SaveBtn','ClearBtn',
                 'PanelClients','RefreshBtn','OpenExcelBtn','EditBtn','SearchBox','SearchHint','FilterBox','OnlyDebt','ClientsGrid',
                 'PanelSettings','SetTitle','WelcomeText','SetFileL','ChangeFileBtn','FilePathBox','SetLangL','LangBox','SetNatL','NatBox',
                 'BackupChk','OpenAfterChk','SetNote','SaveSettingsBtn','BackupsBtn','UpdateBtn',
                 'PanelAbout','AboutLogo','AboutVer','AboutText','AboutBy','EasterEgg')) {
    $script:UI[$n] = $script:Win.FindName($n)
}
$ui = $script:UI

# navegación
$ui.NavNew.Add_Checked({ Show-Panel 'PanelForm' })
$ui.NavClients.Add_Checked({ Show-Panel 'PanelClients'; Refresh-ClientGrid; $script:UI.SearchBox.Focus() | Out-Null })
$ui.NavSettings.Add_Checked({ Load-SettingsPanel; Show-Panel 'PanelSettings' })
$ui.NavAbout.Add_Checked({ Show-Panel 'PanelAbout' })

# formulario
$ui.SaveBtn.Add_Click({
    Write-Log 'clic Guardar'
    try { Save-FromForm } catch { Write-Log "Guardar ERROR: $($_.Exception.Message) $($_.InvocationInfo.PositionMessage)"; Show-Error "$($_.Exception.Message)" }
})
$ui.ClearBtn.Add_Click({ $wasEdit = [bool]$script:EditRow; Reset-Form; if ($wasEdit) { $script:UI.NavClients.IsChecked = $true } })

# clientes
$ui.SearchBox.Add_TextChanged({
    $script:UI.SearchHint.Visibility = $(if ($script:UI.SearchBox.Text) { 'Collapsed' } else { 'Visible' })
    Refresh-ClientGrid
})
$ui.FilterBox.Add_SelectionChanged({ Refresh-ClientGrid })
$ui.OnlyDebt.Add_Click({ Refresh-ClientGrid })
$ui.ClientsGrid.Add_MouseDoubleClick({ if ($script:UI.ClientsGrid.SelectedItem) { Edit-Selected } })
$ui.EditBtn.Add_Click({ Edit-Selected })
$ui.RefreshBtn.Add_Click({ $script:Win.Cursor = [Windows.Input.Cursors]::Wait; [void](Load-Data); $script:Win.Cursor = $null })
$ui.OpenExcelBtn.Add_Click({
    $sel = $script:UI.ClientsGrid.SelectedItem
    if ($sel) { Open-ExcelWindow $sel.Row } else { Open-ExcelWindow }
})

# ajustes
$ui.ChangeFileBtn.Add_Click({ [void](Choose-ExcelFile) })
$ui.SaveSettingsBtn.Add_Click({ Save-Settings })
$ui.BackupsBtn.Add_Click({ Start-Process explorer.exe $script:BackupDir })
$ui.UpdateBtn.Add_Click({ $script:Win.Cursor = [Windows.Input.Cursors]::Wait; Check-Update -Manual; $script:Win.Cursor = $null })

# trámite
$ui.TramiteBox.Add_SelectionChanged({
    $i = $script:UI.TramiteBox.SelectedIndex
    if ($i -ge 0 -and $script:Tramites[$i].Id -ne $script:T.Id) {
        $script:T = $script:Tramites[$i]; $script:Cfg.Tramite = $script:T.Id; Save-Config
        [void](Load-Data); Build-Form; Reset-Form
    }
})

# huevo de pascua: 5 clics en el logo de «Acerca de»
$script:LogoClicks = 0
$ui.AboutLogo.Add_MouseLeftButtonUp({
    $script:LogoClicks++
    if ($script:LogoClicks -ge 5) { $script:UI.EasterEgg.Visibility = 'Visible'; $script:LogoClicks = 0 }
})

# atajo: Ctrl+S guarda el cliente
$script:Win.Add_KeyDown({
    if ($_.Key -eq 'S' -and [Windows.Input.Keyboard]::Modifiers -eq 'Control' -and $script:UI.PanelForm.Visibility -eq 'Visible') { Save-FromForm; $_.Handled = $true }
})

$script:Win.Dispatcher.Add_UnhandledException({
    $m = "$($_.Exception.Message)`r`n$($_.Exception.InnerException)"
    Write-Log "UI ERROR: $($_.Exception.ToString())"
    [void][Windows.MessageBox]::Show($m, 'GESTEM MANAGER - error', 'OK', 'Error')
    $_.Handled = $true
})

Apply-Language

$script:Win.Add_ContentRendered({
    Check-Update
    if ($script:Restarting) { return }
    $flag = Join-Path $script:DataDir 'actualizado.txt'
    if (Test-Path $flag) { $v = (Get-Content -Raw $flag).Trim(); Remove-Item $flag -Force; Show-Info ((L 'updated') -f $v) }
    if (-not (Test-ExcelInstalled)) { Show-Error (L 'errNoExcel') }
    if (-not $script:Cfg.ExcelPath -or -not (Test-Path -LiteralPath $script:Cfg.ExcelPath)) {
        Set-Status (L 'noFile')
        $script:UI.NavSettings.IsChecked = $true
        return
    }
    $script:Win.Cursor = [Windows.Input.Cursors]::Wait
    if (Load-Data) { Apply-Language }
    $script:Win.Cursor = $null
})

try {
    [void]$script:Win.ShowDialog()
} catch {
    Write-Log "fatal: $($_.Exception.ToString())"
    [void][Windows.MessageBox]::Show("$($_.Exception.Message)", $script:AppName, 'OK', 'Error')
}
