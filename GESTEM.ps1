#  GESTEM MANAGER  -  Gestión de clientes para gestoría
#  Creado por Yano  ·  2026
#
#  © 2026 Yano. Todos los derechos reservados.
#  Programa con licencia de uso para la empresa titular (Gestoría El Mahi).
#  Prohibida su copia, modificación, distribución o reventa sin permiso escrito del autor.
#  Escribe directamente en el Excel de la gestoría usando Microsoft Excel.
#  Requiere: Windows 10/11, Windows PowerShell 5.1 y Microsoft Excel instalado.

param([switch]$Oculto)
# sin ventana de consola: si se abrió de otra forma (Terminal de Windows), se vuelve a abrir oculto
if (-not $Oculto -and $PSCommandPath) {
    try {
        $psi = New-Object Diagnostics.ProcessStartInfo
        $psi.FileName = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -STA -File "' + $PSCommandPath + '" -Oculto'
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        [void][Diagnostics.Process]::Start($psi)
        exit
    } catch { }
}
$ErrorActionPreference = 'Stop'
trap {
    $msg = "$($_.Exception.Message)`r`n$($_.InvocationInfo.PositionMessage)"
    try { $d = Join-Path $env:APPDATA 'GESTEM'; New-Item -ItemType Directory -Force -Path $d | Out-Null
          Add-Content -Path (Join-Path $d 'gestem.log') -Value ("{0:yyyy-MM-dd HH:mm:ss}  FATAL  {1}" -f (Get-Date), $msg) -Encoding UTF8 } catch {}
    try { Add-Type -AssemblyName PresentationFramework; [void][Windows.MessageBox]::Show($msg, 'GESTEM MANAGER - error', 'OK', 'Error') } catch {}
    exit 1
}
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Security, System.Drawing
# barra de tareas: GESTEM con su propio icono (si no, Windows lo agrupa como PowerShell y enseña el icono de PowerShell)
try {
    Add-Type -Namespace GestemNative -Name Shell -MemberDefinition '[DllImport("shell32.dll")] public static extern int SetCurrentProcessExplicitAppUserModelID([MarshalAs(UnmanagedType.LPWStr)] string AppID);'
    [void][GestemNative.Shell]::SetCurrentProcessExplicitAppUserModelID('Yano.GESTEM.Manager')
} catch { }

$script:AppName    = 'GESTEM MANAGER'
$script:AppVersion = '1.5.10'
$script:UpdateBase = 'https://raw.githubusercontent.com/ghoulproton-dev/gestem-manager/main'
$script:AppDir     = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:DataDir    = Join-Path $env:APPDATA 'GESTEM'
$script:CfgPath    = Join-Path $script:DataDir 'config.json'
$script:BackupDir  = Join-Path $script:DataDir 'copias'
$script:LogPath    = Join-Path $script:DataDir 'gestem.log'
$script:ES         = [Globalization.CultureInfo]::GetCultureInfo('es-ES')
$script:User       = ''
$script:StatsScopeId = $null
$script:PayColors  = @{ 'Pagado' = '#A9D08E'; 'Parcial' = '#FFE699'; 'Pendiente' = '#FF9B9B' }
$script:Cache      = @{}
$script:CacheStamp = 0
New-Item -ItemType Directory -Force -Path $script:DataDir, $script:BackupDir | Out-Null

function Write-Log([string]$msg) {
    # en modo administrador todo se firma como ADMIN (así se distingue de lo que hace el equipo)
    # en el PC del creador (o en modo admin) se firma ADMIN; los avisos automáticos (sin usuario) siguen siendo del sistema
    $who = $(if ($script:IsAdmin -or ($script:IsOwnerPC -and $script:User)) { 'ADMIN' } else { $script:User })
    try { Add-Content -Path $script:LogPath -Value ("{0:yyyy-MM-dd HH:mm:ss}  [{1}] {2}" -f (Get-Date), $who, $msg) -Encoding UTF8 } catch {}
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
        LastCol   = 'U'
        NetCol    = 'I'
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
            @{ Key='resp'; Col='U'; Type='ComboEdit'; Section='client'; Role='user'; List='colleagues'
               Label=@{ es='Responsable'; en='Handled by'; fr='Responsable'; ar='المسؤول' } }
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

# ---- los demás trámites (misma estructura en el Excel: columnas A..T) ----
$script:GenStatusColors = @{ 'Documentación' = '#F4B183'; 'Presentado' = '#FFE699'; 'En Trámite' = '#FFF2CC';
                             'Requerimiento' = '#FF9B9B'; 'Favorable' = '#C6E0B4'; 'Desfavorable' = '#BFBFBF'; 'Finalizado' = '#A9D08E' }
function New-Label([string]$es, [string]$en, [string]$fr, [string]$ar) { return @{ es = $es; en = $en; fr = $fr; ar = $ar } }
function New-GenTramite([string]$Id, [string]$Sheet, [hashtable]$Title, [array]$Extras) {
    $list = New-Object System.Collections.Generic.List[object]
    $list.Add(@{ Key='name'; Col='B'; Type='Upper'; Section='client'; Span=2; Required=$true; Label=(New-Label 'Nombre y apellidos' 'Full name' 'Nom et prénom' 'الاسم الكامل') })
    $list.Add(@{ Key='doc'; Col='D'; Type='Upper'; Section='client'; Label=(New-Label 'Documento (NIE / Pasaporte)' 'ID (NIE / Passport)' 'Pièce (NIE / Passeport)' 'الوثيقة (NIE / جواز السفر)') })
    $list.Add(@{ Key='nat'; Col='C'; Type='ComboEdit'; Section='client'; Role='nationality'; List='nationalities'; Label=(New-Label 'Nacionalidad' 'Nationality' 'Nationalité' 'الجنسية') })
    $list.Add(@{ Key='tel'; Col='E'; Type='Text'; Section='client'; Label=(New-Label 'Teléfono' 'Phone' 'Téléphone' 'الهاتف') })
    $list.Add(@{ Key='resp'; Col='U'; Type='ComboEdit'; Section='client'; Role='user'; List='colleagues'; Label=(New-Label 'Responsable' 'Handled by' 'Responsable' 'المسؤول') })
    $letters = @('F', 'G', 'H', 'I')
    for ($k = 0; $k -lt $Extras.Count; $k++) {
        $ex = $Extras[$k]
        $fx = @{ Key = ('x' + $letters[$k]); Col = $letters[$k]; Section = 'proc'; KeepCase = $true; Label = @{ es = $ex[0] } }
        if ($ex[1]) { $fx.Type = 'ComboEdit'; $fx.List = ('fixed:' + $ex[1]) } else { $fx.Type = 'Text' }
        $list.Add($fx)
    }
    $list.Add(@{ Key='status'; Col='N'; Type='Pills'; Section='proc'; Span=2; List='excel:N'; Colored=$true; Label=(New-Label 'Estado del trámite' 'Procedure status' 'État de la démarche' 'حالة الإجراء') })
    $list.Add(@{ Key='fsol'; Col='O'; Type='Date'; Section='proc'; DefaultToday=$true; Label=(New-Label 'Fecha de solicitud' 'Application date' 'Date de demande' 'تاريخ الطلب') })
    $list.Add(@{ Key='fres'; Col='P'; Type='Date'; Section='proc'; Label=(New-Label 'Fecha de resolución' 'Resolution date' 'Date de résolution' 'تاريخ القرار') })
    $list.Add(@{ Key='tasa'; Col='Q'; Type='Combo'; Section='proc'; List='fixed:Sí|No|No aplica'; Label=(New-Label 'Tasa pagada' 'Fee paid' 'Taxe payée' 'الرسوم مدفوعة') })
    $list.Add(@{ Key='price'; Col='J'; Type='Money'; Section='pay'; Role='gprice'; Label=(New-Label 'Precio total (€)' 'Total price (€)' 'Prix total (€)' 'السعر الإجمالي (€)') })
    $list.Add(@{ Key='paystate'; Col='M'; Type='Pills'; Section='pay'; List='excel:M'; Role='paystate'; Default='Pendiente'; Label=(New-Label 'Estado de pago' 'Payment status' 'État du paiement' 'حالة الدفع') })
    $list.Add(@{ Key='paid'; Col='K'; Type='Money'; Section='pay'; Role='paid'; Label=(New-Label 'Ha pagado (€) — solo si es parcial' 'Amount paid (€) — partial only' 'Montant payé (€) — si partiel' 'المبلغ المدفوع (€) — عند الدفع الجزئي') })
    $list.Add(@{ Key='gastos'; Col='V'; Type='Money'; Section='pay'; Label=(New-Label 'Gastos: tasa, traducción… (€)' 'Costs: fees, translation… (€)' 'Frais : taxe, traduction… (€)' 'المصاريف: رسوم، ترجمة… (€)') })
    $list.Add(@{ Key='forma'; Col='S'; Type='ComboEdit'; Section='pay'; List='payments'; Label=(New-Label 'Forma de pago' 'Payment method' 'Mode de paiement' 'طريقة الدفع') })
    $list.Add(@{ Key='obs'; Col='R'; Type='Multiline'; Section='notes'; Span=2; Label=(New-Label 'Observaciones' 'Notes' 'Observations' 'ملاحظات') })
    return @{ Id = $Id; Title = $Title; Sheet = $Sheet; FirstRow = 11; LastRow = 1009; LastCol = 'W'; NetCol = 'W'; IdCol = 'A'; NameCol = 'B'
              DocCol = 'D'; StatusCol = 'N'; PayCol = 'M'; PendCol = 'L'; DaysCol = 'T'; TipoCol = $null; PriceKey = $Title['es']
              DefaultNationality = 'MARROQUI'; StatusColors = $script:GenStatusColors; Fields = $list.ToArray() }
}
$script:Tramites += New-GenTramite 'nacionalidad' 'NACIONALIDAD' (New-Label 'Nacionalidad española' 'Spanish nationality' 'Nationalité espagnole' 'الجنسية الإسبانية') @(
    @('Tipo', 'Residencia|Matrimonio|Opción|Sefardí|Otro'), @('Examen CCSE / DELE', 'Aprobado|Pendiente|Dispensa'), @('Nº expediente', ''), @('Contraseña / Cl@ve', ''))
$script:Tramites += New-GenTramite 'homologacion' 'HOMOLOGACIÓN' (New-Label 'Homologación' 'Degree recognition' 'Homologation de diplôme' 'معادلة الشهادات') @(
    @('Título', 'Bachillerato|ESO|Universitario|FP|Otro'), @('Relevé de notes', 'Tiene|Falta'), @('Contraseña', ''), @('Nº expediente / volante', ''))
$script:Tramites += New-GenTramite 'autonomo' 'AUTÓNOMO' (New-Label 'Autónomo' 'Self-employed' 'Auto-entrepreneur' 'العمل الحر (أوتونومو)') @(
    @('Actividad', ''), @('Provincia', ''), @('Certificado digital', 'Tiene|No tiene'), @('Contraseña', ''))
$script:Tramites += New-GenTramite 'extranjeria' 'EXTRANJERÍA' (New-Label 'Extranjería' 'Immigration' 'Immigration (Extranjería)' 'شؤون الأجانب') @(
    @('Tipo de gestión', 'Regularización extraordinaria|Arraigo|Renovación|Modificación cta. ajena|Modificación cta. propia|Reagrupación familiar|Tarjeta familiar UE|Otro'),
    @('Provincia', ''), @('Nº expediente', ''), @('Contraseña', ''))
$script:Tramites += New-GenTramite 'traduccion' 'TRADUCCIÓN' (New-Label 'Traducción' 'Translation' 'Traduction' 'الترجمة') @(
    @('Documento', 'Acta de nacimiento|Acta de matrimonio|Relevé de notes|Diploma Bac|Antecedentes penales|Otro'), @('Idioma', 'Árabe → Español|Francés → Español|Otro'),
    @('Nº de páginas', ''), @('Enviado por', ''))
$script:Tramites += New-GenTramite 'matriculacion' 'MATRICULACIÓN' (New-Label 'Matriculación' 'Vehicle registration' 'Immatriculation' 'تسجيل المركبات') @(
    @('Tipo', 'Matriculación|Cambio de titular|Importación|Otro'), @('Matrícula', ''), @('Bastidor', ''), @('Vendedor / comprador', ''))
$script:Tramites += New-GenTramite 'sis' 'SIS' (New-Label 'SIS (Schengen)' 'SIS (Schengen)' 'SIS (Schengen)' 'نظام شنغن (SIS)') @(
    @('Tipo de solicitud', 'Acceso|Rectificación|Supresión'), @('País de la alerta', ''), @('Motivo', 'Expulsión|Prohibición de entrada|Otro'), @('Nº expediente', ''))
$script:Tramites += New-GenTramite 'certificado' 'CERTIFICADO DIGITAL' (New-Label 'Certificado digital' 'Digital certificate' 'Certificat numérique' 'الشهادة الرقمية') @(
    @('Tipo', 'Persona física|Representante'), @('Usuario / inicio de sesión', ''), @('Contraseña', ''), @('Caduca', ''))
$script:Tramites += New-GenTramite 'dispensa' 'DISPENSA EXAMEN' (New-Label 'Dispensa examen nacionalidad' 'Nationality exam exemption' "Dispense d'examen (nationalité)" 'الإعفاء من امتحان الجنسية') @(
    @('Examen', 'CCSE|DELE A2|Ambos'), @('Motivo', 'Edad|Estudios en España|Discapacidad|Otro'), @('Nº expediente', ''), @('Contraseña', ''))
$script:Tramites += New-GenTramite 'visado' 'VISADO ESTUDIOS' (New-Label 'Visado de estudios' 'Student visa' 'Visa étudiant' 'تأشيرة الدراسة') @(
    @('Centro de estudios', ''), @('Consulado', ''), @('Fecha de cita', ''), @('Nº expediente', ''))

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
  colId='ID'; colName='Cliente'; colDoc='NIE / DNI'; colTipo='Tipo'; colStatus='Estado'; colPay='Pago'; colPend='Pendiente'; colDays='Días'
  secExtT='Copia externa'; secExtD='Una copia completa del Excel fuera del programa: en un USB, OneDrive, Google Drive o una carpeta de red. Si algo va mal en este PC, los datos siguen a salvo.'; extDirL='Carpeta de la copia externa'; extChooseBtn='Elegir carpeta…'; extOneDrive='Usar la nube de la empresa'; extNoCloud='No se encontró OneDrive ni Google Drive en este PC. Instala la app de la nube de la empresa e inicia sesión, o elige la carpeta a mano.'; extNow='Hacer copia ahora'; extOpen='Abrir carpeta'; extLast='Última copia: {0}'; extNever='Todavía no se ha hecho ninguna copia'; extNone='Sin carpeta: elige una para proteger los datos'; extNote='GESTEM copia el Excel al abrirse y al cerrarse, y cuando pulsas «Hacer copia ahora». Guarda un archivo por día (60 días). Cada copia es un Excel normal que se abre sin GESTEM. Los ajustes del programa se guardan en «ajustes GESTEM».'; extChoose='Elige dónde guardar la copia externa del Excel'; extFail='No se pudo hacer la copia externa.'; extDone='Copia externa hecha.'
  stAll='Ver todos los trámites'
  scopeAll='Todos los trámites'; stStatus='Estado'; statsTableSt='Por estado · {0}'; statsSubOne='{0} · {1} expedientes · actualizado a las {2}'; stNoStatus='Sin estado'; stClickHint='Haz clic en un trámite para ver solo sus datos'
  setSub='Configura GESTEM para tu oficina. Los cambios se aplican al pulsar «Guardar ajustes».'; secFileT='Archivo de clientes'; secFileD='El Excel donde GESTEM escribe y lee los clientes.'; secCompT='Empresa'; secCompD='Nombre y logo que salen en el menú y en la pantalla de inicio.'; secTeamT='Equipo'; secTeamD='Las personas que usan GESTEM. Salen al iniciar sesión y como responsable.'; secPrefT='Preferencias'; secPrefD='Idioma del programa, valores por defecto y copias de seguridad.'; secHelpT='Ayuda y versión'; secHelpD='Manual de uso, actualizaciones y copias guardadas.'
  chDay='Día'; chWeek='Semana'; chMonth='Mes'; chTitleD='Actividad por día'; chTitleW='Actividad por semana'; chTitleM='Actividad por mes'; chSubD='Últimos 14 días · {0} expedientes · {1}'; chSubW='Últimas 12 semanas · {0} expedientes · {1}'; chSubM='Últimos 12 meses · {0} expedientes · {1}'; chPaid='Sube'; chPend='Baja'; chMoney='facturado (debajo)'; chPrevD='ayer'; chPrevW='la semana anterior'; chPrevM='el mes anterior'; chDeltaTxt='{0} expedientes vs {1}'; chTip='{0}: {1} expedientes · facturado {2} · cobrado {3}'; chWeekL='S{0}'
  btnHelp='Manual de uso'
  navStats='Resumen'; statsTitle='Resumen económico'; statsSub='{0} trámites · {1} expedientes · actualizado a las {2}'; statsTable='Por trámite'; kpiBill='Facturado'; kpiPaid='Cobrado'; kpiPend='Pendiente de cobro'; kpiNet='Beneficio neto'; kpiBillF='{0} expedientes'; kpiPaidF='{0} % de lo facturado'; kpiPendF='{0} clientes con pago pendiente'; kpiNetF='Margen {0} %'; stTramite='Trámite'; stExp='Exped.'; stBill='Facturado'; stPaid='Cobrado'; stPend='Pendiente'; stNet='Neto'; stMargin='Margen'; stShare='Peso en la facturación'; stTotal='TOTAL'; whoAll='Todo el equipo'; perAll='Todo el tiempo'; perMonth='Este mes'; perPrev='Mes anterior'; perYear='Este año'
  setTitle='Ajustes'; setFile='Archivo Excel de clientes'; btnChange='Cambiar…'; setLang='Idioma'
  setNat='Nacionalidad por defecto en clientes nuevos'; setBackup='Hacer una copia de seguridad del Excel antes de cada guardado'
  setOpen='Abrir Excel después de guardar un cliente'; btnBackups='Abrir carpeta de copias'; btnSaveSettings='Guardar ajustes'
  setNote='Los precios y las listas (tipos de permiso, pasos) se cambian en la pestaña AJUSTES del Excel.'
  settingsSaved='Ajustes guardados.'; chooseFile='Elige el archivo Excel de clientes'
  welcome='Para empezar, elige tu archivo Excel de clientes.'
  loading='Leyendo el Excel…'; savingMsg='Guardando en el Excel…'; ready='{0} clientes · {1} con pago pendiente ({2})'
  noFile='Elige tu archivo Excel en Ajustes.'; tramite='Trámite'
  errNoFile='No encuentro el archivo Excel. Elígelo en Ajustes.'
  errNoSheet='Este Excel no tiene la pestaña «{0}». Elige el archivo correcto en Ajustes.'; errTabsFound='Pestañas de este Excel: {0}'; handReadOnly='Este cliente está en la pestaña «{0}» del Excel. Cámbialo allí: GESTEM lo vuelve a leer solo cada vez que se guarda el Excel.'
  errLocked='El Excel está abierto en otro ordenador o en otro programa. Ciérralo allí y vuelve a intentarlo.'
  errBusy='Excel está ocupado: si estás escribiendo en una celda, pulsa Enter en Excel y vuelve a intentarlo.'
  errFull='No quedan filas libres en el Excel.'; errNoExcel='GESTEM MANAGER necesita Microsoft Excel instalado en este ordenador.'
  errGeneric='No se pudo completar la operación:'; selectFirst='Selecciona un cliente de la lista.'
  aboutText='Programa de gestión de clientes para gestorías: canjes de permiso de conducir y otros trámites. Guarda todo directamente en tu Excel.'
  aboutBy='Creado por Yano'; aboutVer='Versión {0}'
  partialOnly='(solo si el pago es Parcial)'; none='(ninguno)'
  updating='Actualizando GESTEM a la versión {0}…'; updated='GESTEM actualizado a la versión {0}.'; upToDate='Ya tienes la última versión ({0}).'; updateFail='No se pudo buscar la actualización (¿hay internet?).'; btnUpdate='Buscar actualizaciones'
  sessTitle='¿Quién está usando GESTEM?'; sessHint='Elige tu nombre. GESTEM lo recuerda y lo pone como responsable en los clientes que guardes.'; sessNew='Otro nombre…'; sessGo='Entrar'; userTip='Usuario: {0}'; onlyMine='Solo mis clientes'; colResp='Responsable'; setColleagues='Compañeras/os (separados por comas)'; setCompany='Nombre de la empresa'; setLogo='Logo de la empresa'; btnLogo='Elegir imagen…'; btnLogoClear='Quitar'; chooseLogo='Elige el logo de la empresa'
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
  colId='ID'; colName='Client'; colDoc='NIE / DNI'; colTipo='Type'; colStatus='Status'; colPay='Payment'; colPend='Outstanding'; colDays='Days'
  secExtT='External backup'; secExtD='A full copy of the Excel file outside the program: on a USB drive, OneDrive, Google Drive or a network folder. If something goes wrong on this PC, the data is still safe.'; extDirL='External backup folder'; extChooseBtn='Choose folder…'; extOneDrive='Use the company cloud'; extNoCloud='No OneDrive or Google Drive found on this PC. Install the company''s cloud app and sign in, or choose the folder manually.'; extNow='Back up now'; extOpen='Open folder'; extLast='Last backup: {0}'; extNever='No backup made yet'; extNone='No folder: choose one to protect the data'; extNote='GESTEM copies the Excel file when it opens and when it closes, and when you press «Back up now». It keeps one file per day (60 days). Each copy is a normal Excel file that opens without GESTEM. Program settings are saved in «ajustes GESTEM».'; extChoose='Choose where to keep the external backup'; extFail='The external backup could not be made.'; extDone='External backup done.'
  stAll='See all procedures'
  scopeAll='All procedures'; stStatus='Status'; statsTableSt='By status · {0}'; statsSubOne='{0} · {1} cases · updated at {2}'; stNoStatus='No status'; stClickHint='Click a procedure to see only its data'
  setSub='Set GESTEM up for your office. Changes apply when you press «Save settings».'; secFileT='Client file'; secFileD='The Excel file GESTEM reads and writes clients to.'; secCompT='Company'; secCompD='Name and logo shown in the menu and on the start screen.'; secTeamT='Team'; secTeamD='The people who use GESTEM. Shown at sign-in and as the person in charge.'; secPrefT='Preferences'; secPrefD='Language, defaults and backups.'; secHelpT='Help and version'; secHelpD='User manual, updates and saved backups.'
  chDay='Day'; chWeek='Week'; chMonth='Month'; chTitleD='Activity per day'; chTitleW='Activity per week'; chTitleM='Activity per month'; chSubD='Last 14 days · {0} cases · {1}'; chSubW='Last 12 weeks · {0} cases · {1}'; chSubM='Last 12 months · {0} cases · {1}'; chPaid='Up'; chPend='Down'; chMoney='billed (below)'; chPrevD='yesterday'; chPrevW='last week'; chPrevM='last month'; chDeltaTxt='{0} cases vs {1}'; chTip='{0}: {1} cases · billed {2} · collected {3}'; chWeekL='W{0}'
  btnHelp='User manual'
  navStats='Summary'; statsTitle='Financial summary'; statsSub='{0} procedures · {1} cases · updated at {2}'; statsTable='By procedure'; kpiBill='Billed'; kpiPaid='Collected'; kpiPend='Outstanding'; kpiNet='Net profit'; kpiBillF='{0} cases'; kpiPaidF='{0} % of billed'; kpiPendF='{0} clients with payment due'; kpiNetF='Margin {0} %'; stTramite='Procedure'; stExp='Cases'; stBill='Billed'; stPaid='Collected'; stPend='Outstanding'; stNet='Net'; stMargin='Margin'; stShare='Share of billing'; stTotal='TOTAL'; whoAll='Whole team'; perAll='All time'; perMonth='This month'; perPrev='Last month'; perYear='This year'
  setTitle='Settings'; setFile='Client Excel file'; btnChange='Change…'; setLang='Language'
  setNat='Default nationality for new clients'; setBackup='Back up the Excel file before every save'
  setOpen='Open Excel after saving a client'; btnBackups='Open backups folder'; btnSaveSettings='Save settings'
  setNote='Prices and lists (licence types, steps) are changed in the AJUSTES tab of the Excel file.'
  settingsSaved='Settings saved.'; chooseFile='Choose the client Excel file'
  welcome='To start, choose your client Excel file.'
  loading='Reading the Excel file…'; savingMsg='Saving to Excel…'; ready='{0} clients · {1} with outstanding payment ({2})'
  noFile='Choose your Excel file in Settings.'; tramite='Procedure'
  errNoFile="Can't find the Excel file. Choose it in Settings."
  errNoSheet='This Excel file has no «{0}» tab. Choose the right file in Settings.'; errTabsFound='Tabs in this Excel file: {0}'; handReadOnly='This client is in the «{0}» tab of the Excel file. Change it there: GESTEM reads it again by itself every time the Excel file is saved.'
  errLocked='The Excel file is open on another computer or program. Close it there and try again.'
  errBusy="Excel is busy: if you're typing in a cell, press Enter in Excel and try again."
  errFull='There are no free rows left in the Excel file.'; errNoExcel='GESTEM MANAGER needs Microsoft Excel installed on this computer.'
  errGeneric="The operation couldn't be completed:"; selectFirst='Select a client in the list.'
  aboutText='Client management for administrative offices: driving licence exchanges and other procedures. Everything is saved straight into your Excel file.'
  aboutBy='Created by Yano'; aboutVer='Version {0}'
  partialOnly='(only if payment is Partial)'; none='(none)'
  updating='Updating GESTEM to version {0}…'; updated='GESTEM updated to version {0}.'; upToDate='You already have the latest version ({0}).'; updateFail="Couldn't check for updates (is there internet?)."; btnUpdate='Check for updates'
  sessTitle='Who is using GESTEM?'; sessHint='Choose your name. GESTEM remembers it and records you as the person in charge of the clients you save.'; sessNew='Another name…'; sessGo='Start'; userTip='User: {0}'; onlyMine='My clients only'; colResp='Handled by'; setColleagues='Colleagues (separated by commas)'; setCompany='Company name'; setLogo='Company logo'; btnLogo='Choose image…'; btnLogoClear='Remove'; chooseLogo='Choose the company logo'
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
  colId='ID'; colName='Client'; colDoc='NIE / DNI'; colTipo='Type'; colStatus='État'; colPay='Paiement'; colPend='Reste'; colDays='Jours'
  secExtT='Copie externe'; secExtD='Une copie complète de l''Excel hors du programme : sur une clé USB, OneDrive, Google Drive ou un dossier réseau. Si quelque chose arrive à ce PC, les données restent en sécurité.'; extDirL='Dossier de la copie externe'; extChooseBtn='Choisir un dossier…'; extOneDrive='Utiliser le cloud de l''entreprise'; extNoCloud='Aucun OneDrive ni Google Drive trouvé sur ce PC. Installez l''application cloud de l''entreprise et connectez-vous, ou choisissez le dossier à la main.'; extNow='Faire une copie maintenant'; extOpen='Ouvrir le dossier'; extLast='Dernière copie : {0}'; extNever='Aucune copie pour l''instant'; extNone='Aucun dossier : choisissez-en un pour protéger les données'; extNote='GESTEM copie l''Excel à l''ouverture et à la fermeture, et avec « Faire une copie maintenant ». Un fichier par jour (60 jours). Chaque copie est un Excel normal qui s''ouvre sans GESTEM.'; extChoose='Choisissez où garder la copie externe'; extFail='La copie externe n''a pas pu être faite.'; extDone='Copie externe faite.'
  stAll='Voir toutes les démarches'
  scopeAll='Toutes les démarches'; stStatus='État'; statsTableSt='Par état · {0}'; statsSubOne='{0} · {1} dossiers · mis à jour à {2}'; stNoStatus='Sans état'; stClickHint='Cliquez sur une démarche pour voir seulement ses données'
  setSub='Configurez GESTEM pour votre bureau. Les changements s''appliquent avec « Enregistrer les paramètres ».'; secFileT='Fichier des clients'; secFileD='Le fichier Excel où GESTEM lit et écrit les clients.'; secCompT='Entreprise'; secCompD='Nom et logo affichés dans le menu et sur l''écran d''accueil.'; secTeamT='Équipe'; secTeamD='Les personnes qui utilisent GESTEM. Affichées à la connexion et comme responsable.'; secPrefT='Préférences'; secPrefD='Langue, valeurs par défaut et copies de sécurité.'; secHelpT='Aide et version'; secHelpD='Mode d''emploi, mises à jour et copies enregistrées.'
  chDay='Jour'; chWeek='Semaine'; chMonth='Mois'; chTitleD='Activité par jour'; chTitleW='Activité par semaine'; chTitleM='Activité par mois'; chSubD='14 derniers jours · {0} dossiers · {1}'; chSubW='12 dernières semaines · {0} dossiers · {1}'; chSubM='12 derniers mois · {0} dossiers · {1}'; chPaid='Hausse'; chPend='Baisse'; chMoney='facturé (dessous)'; chPrevD='hier'; chPrevW='la semaine dernière'; chPrevM='le mois dernier'; chDeltaTxt='{0} dossiers vs {1}'; chTip='{0} : {1} dossiers · facturé {2} · encaissé {3}'; chWeekL='S{0}'
  btnHelp='Mode d''emploi'
  navStats='Résumé'; statsTitle='Résumé financier'; statsSub='{0} démarches · {1} dossiers · mis à jour à {2}'; statsTable='Par démarche'; kpiBill='Facturé'; kpiPaid='Encaissé'; kpiPend='Reste à encaisser'; kpiNet='Bénéfice net'; kpiBillF='{0} dossiers'; kpiPaidF='{0} % du facturé'; kpiPendF='{0} clients avec paiement en attente'; kpiNetF='Marge {0} %'; stTramite='Démarche'; stExp='Dossiers'; stBill='Facturé'; stPaid='Encaissé'; stPend='Reste'; stNet='Net'; stMargin='Marge'; stShare='Part de la facturation'; stTotal='TOTAL'; whoAll='Toute l''équipe'; perAll='Depuis le début'; perMonth='Ce mois-ci'; perPrev='Mois dernier'; perYear='Cette année'
  setTitle='Paramètres'; setFile='Fichier Excel des clients'; btnChange='Changer…'; setLang='Langue'
  setNat='Nationalité par défaut des nouveaux clients'; setBackup="Faire une copie de sécurité de l'Excel avant chaque enregistrement"
  setOpen="Ouvrir Excel après l'enregistrement d'un client"; btnBackups='Ouvrir le dossier des copies'; btnSaveSettings='Enregistrer les paramètres'
  setNote="Les prix et les listes (types de permis, étapes) se modifient dans l'onglet AJUSTES de l'Excel."
  settingsSaved='Paramètres enregistrés.'; chooseFile='Choisissez le fichier Excel des clients'
  welcome='Pour commencer, choisissez votre fichier Excel des clients.'
  loading="Lecture de l'Excel…"; savingMsg="Enregistrement dans l'Excel…"; ready='{0} clients · {1} avec paiement en attente ({2})'
  noFile='Choisissez votre fichier Excel dans Paramètres.'; tramite='Démarche'
  errNoFile='Fichier Excel introuvable. Choisissez-le dans Paramètres.'
  errNoSheet="Cet Excel n'a pas d'onglet « {0} ». Choisissez le bon fichier dans Paramètres."; errTabsFound='Onglets de cet Excel : {0}'; handReadOnly='Ce client se trouve dans l''onglet « {0} » de l''Excel. Modifiez-le là-bas : GESTEM le relit tout seul à chaque enregistrement de l''Excel.'
  errLocked="L'Excel est ouvert sur un autre ordinateur ou programme. Fermez-le puis réessayez."
  errBusy="Excel est occupé : si vous écrivez dans une cellule, appuyez sur Entrée dans Excel puis réessayez."
  errFull="Il n'y a plus de lignes libres dans l'Excel."; errNoExcel='GESTEM MANAGER a besoin de Microsoft Excel installé sur cet ordinateur.'
  errGeneric="L'opération n'a pas pu aboutir :"; selectFirst='Sélectionnez un client dans la liste.'
  aboutText="Gestion des clients pour gestorías : échanges de permis de conduire et autres démarches. Tout est enregistré directement dans votre Excel."
  aboutBy='Créé par Yano'; aboutVer='Version {0}'
  partialOnly='(seulement si le paiement est Partiel)'; none='(aucun)'
  updating='Mise à jour de GESTEM vers la version {0}…'; updated='GESTEM mis à jour vers la version {0}.'; upToDate='Vous avez déjà la dernière version ({0}).'; updateFail="Impossible de chercher la mise à jour (internet ?)."; btnUpdate='Rechercher des mises à jour'
  sessTitle='Qui utilise GESTEM ?'; sessHint="Choisissez votre nom. GESTEM s'en souvient et vous indique comme responsable des clients que vous enregistrez."; sessNew='Autre nom…'; sessGo='Entrer'; userTip='Utilisateur : {0}'; onlyMine='Mes clients seulement'; colResp='Responsable'; setColleagues='Collègues (séparés par des virgules)'; setCompany="Nom de l'entreprise"; setLogo="Logo de l'entreprise"; btnLogo='Choisir une image…'; btnLogoClear='Retirer'; chooseLogo="Choisissez le logo de l'entreprise"
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
  colId='الرقم'; colName='العميل'; colDoc='NIE / DNI'; colTipo='النوع'; colStatus='الحالة'; colPay='الدفع'; colPend='المتبقي'; colDays='الأيام'
  secExtT='نسخة خارجية'; secExtD='نسخة كاملة من ملف Excel خارج البرنامج: على USB أو OneDrive أو Google Drive أو مجلد شبكة. إذا حدث خطأ في هذا الحاسوب تبقى البيانات آمنة.'; extDirL='مجلد النسخة الخارجية'; extChooseBtn='اختيار مجلد…'; extOneDrive='استخدام سحابة الشركة'; extNoCloud='لم يُعثر على OneDrive أو Google Drive في هذا الحاسوب. ثبّت تطبيق سحابة الشركة وسجّل الدخول، أو اختر المجلد يدويًا.'; extNow='نسخ الآن'; extOpen='فتح المجلد'; extLast='آخر نسخة: {0}'; extNever='لم تُعمل أي نسخة بعد'; extNone='لا يوجد مجلد: اختر واحدًا لحماية البيانات'; extNote='ينسخ GESTEM ملف Excel عند الفتح وعند الإغلاق وعند الضغط على «نسخ الآن». ملف لكل يوم (60 يومًا). كل نسخة ملف Excel عادي يُفتح بدون GESTEM.'; extChoose='اختر مكان حفظ النسخة الخارجية'; extFail='تعذّر عمل النسخة الخارجية.'; extDone='تمت النسخة الخارجية.'
  stAll='عرض كل الإجراءات'
  scopeAll='كل الإجراءات'; stStatus='الحالة'; statsTableSt='حسب الحالة · {0}'; statsSubOne='{0} · {1} ملفات · آخر تحديث {2}'; stNoStatus='بدون حالة'; stClickHint='اضغط على إجراء لعرض بياناته فقط'
  setSub='اضبط GESTEM لمكتبك. تُطبَّق التغييرات عند الضغط على «حفظ الإعدادات».'; secFileT='ملف العملاء'; secFileD='ملف Excel الذي يقرأ منه GESTEM ويكتب فيه العملاء.'; secCompT='الشركة'; secCompD='الاسم والشعار في القائمة وشاشة البداية.'; secTeamT='الفريق'; secTeamD='الأشخاص الذين يستخدمون GESTEM. يظهرون عند الدخول وكمسؤولين.'; secPrefT='التفضيلات'; secPrefD='اللغة والقيم الافتراضية والنسخ الاحتياطية.'; secHelpT='المساعدة والإصدار'; secHelpD='دليل الاستخدام والتحديثات والنسخ المحفوظة.'
  chDay='يوم'; chWeek='أسبوع'; chMonth='شهر'; chTitleD='النشاط حسب اليوم'; chTitleW='النشاط حسب الأسبوع'; chTitleM='النشاط حسب الشهر'; chSubD='آخر 14 يومًا · {0} ملفات · {1}'; chSubW='آخر 12 أسبوعًا · {0} ملفات · {1}'; chSubM='آخر 12 شهرًا · {0} ملفات · {1}'; chPaid='صعود'; chPend='هبوط'; chMoney='المفوتر (بالأسفل)'; chPrevD='أمس'; chPrevW='الأسبوع الماضي'; chPrevM='الشهر الماضي'; chDeltaTxt='{0} ملفات مقارنة بـ {1}'; chTip='{0}: {1} ملفات · مفوتر {2} · محصّل {3}'; chWeekL='أ{0}'
  btnHelp='دليل الاستخدام'
  navStats='الملخص'; statsTitle='الملخص المالي'; statsSub='{0} إجراءات · {1} ملفات · آخر تحديث {2}'; statsTable='حسب الإجراء'; kpiBill='المفوتر'; kpiPaid='المحصّل'; kpiPend='المتبقي للتحصيل'; kpiNet='صافي الربح'; kpiBillF='{0} ملفات'; kpiPaidF='{0} % من المفوتر'; kpiPendF='{0} عملاء عليهم دفع'; kpiNetF='الهامش {0} %'; stTramite='الإجراء'; stExp='الملفات'; stBill='المفوتر'; stPaid='المحصّل'; stPend='المتبقي'; stNet='الصافي'; stMargin='الهامش'; stShare='الحصة من الفوترة'; stTotal='المجموع'; whoAll='كل الفريق'; perAll='كل الفترات'; perMonth='هذا الشهر'; perPrev='الشهر الماضي'; perYear='هذه السنة'
  setTitle='الإعدادات'; setFile='ملف Excel للعملاء'; btnChange='تغيير…'; setLang='اللغة'
  setNat='الجنسية الافتراضية للعملاء الجدد'; setBackup='عمل نسخة احتياطية من ملف Excel قبل كل حفظ'
  setOpen='فتح Excel بعد حفظ العميل'; btnBackups='فتح مجلد النسخ'; btnSaveSettings='حفظ الإعدادات'
  setNote='تُغيَّر الأسعار والقوائم (أنواع الرخص، الخطوات) في تبويب AJUSTES داخل ملف Excel.'
  settingsSaved='تم حفظ الإعدادات.'; chooseFile='اختر ملف Excel للعملاء'
  welcome='للبدء، اختر ملف Excel الخاص بالعملاء.'
  loading='جارٍ قراءة ملف Excel…'; savingMsg='جارٍ الحفظ في Excel…'; ready='{0} عميل · {1} بدفع متبقٍ ({2})'
  noFile='اختر ملف Excel من الإعدادات.'; tramite='الإجراء'
  errNoFile='لم أجد ملف Excel. اختره من الإعدادات.'
  errNoSheet='ملف Excel هذا لا يحتوي على تبويب «{0}». اختر الملف الصحيح من الإعدادات.'; errTabsFound='تبويبات ملف Excel هذا: {0}'; handReadOnly='هذا العميل موجود في تبويب «{0}» في ملف Excel. عدّله هناك: يعيد GESTEM قراءته تلقائياً كلما حُفظ ملف Excel.'
  errLocked='ملف Excel مفتوح على جهاز أو برنامج آخر. أغلقه هناك ثم حاول مرة أخرى.'
  errBusy='برنامج Excel مشغول: إذا كنت تكتب في خلية، اضغط Enter في Excel ثم حاول مرة أخرى.'
  errFull='لا توجد صفوف فارغة في ملف Excel.'; errNoExcel='يحتاج GESTEM MANAGER إلى برنامج Microsoft Excel مثبتًا على هذا الحاسوب.'
  errGeneric='تعذّر إتمام العملية:'; selectFirst='اختر عميلًا من القائمة.'
  aboutText='برنامج لإدارة عملاء مكاتب الخدمات الإدارية: استبدال رخص القيادة وإجراءات أخرى. يُحفظ كل شيء مباشرة في ملف Excel.'
  aboutBy='من إعداد Yano'; aboutVer='الإصدار {0}'
  partialOnly='(فقط عند الدفع الجزئي)'; none='(لا شيء)'
  updating='جارٍ تحديث GESTEM إلى الإصدار {0}…'; updated='تم تحديث GESTEM إلى الإصدار {0}.'; upToDate='لديك أحدث إصدار ({0}).'; updateFail='تعذّر البحث عن تحديث (هل يوجد اتصال بالإنترنت؟).'; btnUpdate='البحث عن تحديثات'
  sessTitle='من يستخدم GESTEM؟'; sessHint='اختر اسمك. سيتذكره GESTEM ويسجّلك كمسؤول عن العملاء الذين تحفظهم.'; sessNew='اسم آخر…'; sessGo='دخول'; userTip='المستخدم: {0}'; onlyMine='عملائي فقط'; colResp='المسؤول'; setColleagues='الزملاء (مفصولون بفواصل)'; setCompany='اسم الشركة'; setLogo='شعار الشركة'; btnLogo='اختيار صورة…'; btnLogoClear='إزالة'; chooseLogo='اختر شعار الشركة'
 }
}
# textos de documentos, recibos y exportar
$script:TextMore = @{
 es = @{
  conflictAsk = '{0} cambió este cliente a las {1}, mientras lo tenías abierto. ¿Guardar igualmente tus cambios encima de los suyos?'
  someoneElse = 'Otra persona'
  liveUpdated = '↻ Lista actualizada con cambios de otro PC'
  loading = 'Cargando los datos…'
  savingMsg = 'Guardando…'
  errStoreBusy = 'Otro PC está guardando en este momento. Vuelve a intentarlo en unos segundos.'
  errNoStore = 'No se encuentran los datos de GESTEM. Pulsa «Actualizar» en Clientes.'
  errPendingSync = 'Hay clientes guardados en GESTEM que todavía no están en el Excel. Abre el Excel en este PC y pulsa «Actualizar» antes de volver a importar.'
  importAsk = 'Esto vuelve a leer TODOS los clientes desde el Excel y reemplaza los datos de GESTEM por los del Excel. Úsalo solo si el Excel se ha cambiado a mano. ¿Continuar?'
  importDone = '✓ Datos importados de nuevo desde el Excel'
  savedNoExcel = '✓ Guardado en GESTEM (exp. {0}). El Excel se actualizará en cuanto se pueda abrir.'
  btnImport = 'Volver a importar del Excel'
  dataNote = 'GESTEM guarda los clientes en su propia carpeta «GESTEM datos», junto al Excel, y actualiza el Excel en cada guardado.'
  calToday = 'Hoy'
  btnDocs = 'Documentos'
  btnReceipt = 'Recibo PDF'
  btnExport = 'Exportar'
  expPdf = 'Lista en PDF'
  expCsv = 'Lista en CSV (para Excel)'
  colDocs = 'Docs'
  docsTitle = 'Documentos'
  docsSub = 'Expediente {0}  ·  {1}'
  docsAdd = 'Añadir archivos…'
  docsOpen = 'Abrir'
  docsRemove = 'Quitar'
  docsFolder = 'Abrir carpeta'
  docsClose = 'Cerrar'
  docsEmpty = 'Todavía no hay documentos de este cliente. Arrastra aquí los escaneos (PDF, fotos…) o pulsa «Añadir archivos».'
  docsDrop = 'Suelta los archivos para guardarlos'
  docsCount = '{0} documento(s)'
  docsAdded = '✓ {0} documento(s) añadido(s)'
  docsPick = 'Elige primero un documento de la lista.'
  docsRemoveAsk = '¿Quitar «{0}»? Irá a la papelera de Windows.'
  docsFilter = 'Documentos y fotos'
  docsAll = 'Todos los archivos'
  noPdfPrinter = 'Falta la impresora «Microsoft Print to PDF» de Windows. Actívala en: Panel de control → Programas → Activar o desactivar las características de Windows.'
  pdfFail = 'Windows no ha creado el PDF.'
  rcTitle = 'RECIBO'
  rcSubtitle = 'Recibo de pago'
  rcNo = 'Nº {0}'
  rcDate = 'Fecha: {0}'
  rcClient = 'Cliente'
  rcProc = 'Trámite'
  rcDoc = 'Documento'
  rcCase = 'Expediente'
  rcResp = 'Atendido por'
  rcType = 'Tipo'
  rcApplied = 'Solicitud'
  rcStatusP = 'Estado'
  rcPrice = 'Importe del trámite'
  rcPaid = 'Pagado'
  rcPend = 'Pendiente'
  rcMethod = 'Forma de pago'
  rcPaidFull = 'Pagado'
  rcPartial = 'Pago parcial'
  rcSign = 'Firma y sello'
  rcNote = 'Este documento justifica el pago recibido por el trámite indicado. No es una factura.'
  rcFooter = 'Generado con GESTEM MANAGER'
  rcDone = '✓ Recibo {0} guardado en los documentos del cliente'
  exTitle = 'Clientes  ·  {0}'
  exFilters = 'Filtros'
  exNoFilter = 'Todos los clientes'
  exTotal = '{0} clientes  ·  {1} con pago pendiente'
  exPage = 'Página {0} de {1}'
  exEmpty = 'No hay clientes en la lista para exportar.'
  exDone = '✓ Exportado: {0}'
 }
 en = @{
  conflictAsk = '{0} changed this client at {1} while you had it open. Save your changes over theirs anyway?'
  someoneElse = 'Someone else'
  liveUpdated = '↻ List updated with changes from another PC'
  loading = 'Loading data…'
  savingMsg = 'Saving…'
  errStoreBusy = 'Another PC is saving right now. Try again in a few seconds.'
  errNoStore = 'GESTEM data not found. Press «Refresh» in Clients.'
  errPendingSync = 'Some clients saved in GESTEM are not in the Excel file yet. Open Excel on this PC and press «Refresh» before importing again.'
  importAsk = 'This reads ALL clients from the Excel file again and replaces GESTEM''s data with the Excel''s. Only use it if the Excel file was changed by hand. Continue?'
  importDone = '✓ Data imported again from the Excel file'
  savedNoExcel = '✓ Saved in GESTEM (case {0}). The Excel file will be updated as soon as it can be opened.'
  btnImport = 'Import again from Excel'
  dataNote = 'GESTEM keeps clients in its own «GESTEM datos» folder next to the Excel file, and updates the Excel on every save.'
  calToday = 'Today'
  btnDocs = 'Documents'
  btnReceipt = 'PDF receipt'
  btnExport = 'Export'
  expPdf = 'List as PDF'
  expCsv = 'List as CSV (for Excel)'
  colDocs = 'Docs'
  docsTitle = 'Documents'
  docsSub = 'Case {0}  ·  {1}'
  docsAdd = 'Add files…'
  docsOpen = 'Open'
  docsRemove = 'Remove'
  docsFolder = 'Open folder'
  docsClose = 'Close'
  docsEmpty = 'No documents for this client yet. Drag scans here (PDF, photos…) or press «Add files».'
  docsDrop = 'Drop the files to save them'
  docsCount = '{0} document(s)'
  docsAdded = '✓ {0} document(s) added'
  docsPick = 'Pick a document from the list first.'
  docsRemoveAsk = 'Remove «{0}»? It goes to the Windows Recycle Bin.'
  docsFilter = 'Documents and photos'
  docsAll = 'All files'
  noPdfPrinter = 'The Windows «Microsoft Print to PDF» printer is missing. Turn it on in: Control Panel → Programs → Turn Windows features on or off.'
  pdfFail = 'Windows did not create the PDF.'
  rcTitle = 'RECEIPT'
  rcSubtitle = 'Payment receipt'
  rcNo = 'No. {0}'
  rcDate = 'Date: {0}'
  rcClient = 'Client'
  rcProc = 'Procedure'
  rcDoc = 'ID document'
  rcCase = 'Case'
  rcResp = 'Handled by'
  rcType = 'Type'
  rcApplied = 'Applied'
  rcStatusP = 'Status'
  rcPrice = 'Procedure amount'
  rcPaid = 'Paid'
  rcPend = 'Outstanding'
  rcMethod = 'Payment method'
  rcPaidFull = 'Paid'
  rcPartial = 'Partial payment'
  rcSign = 'Signature and stamp'
  rcNote = 'This document confirms the payment received for the procedure shown. It is not an invoice.'
  rcFooter = 'Created with GESTEM MANAGER'
  rcDone = '✓ Receipt {0} saved in the client''s documents'
  exTitle = 'Clients  ·  {0}'
  exFilters = 'Filters'
  exNoFilter = 'All clients'
  exTotal = '{0} clients  ·  {1} with payment due'
  exPage = 'Page {0} of {1}'
  exEmpty = 'There are no clients in the list to export.'
  exDone = '✓ Exported: {0}'
 }
 fr = @{
  conflictAsk = '{0} a modifié ce client à {1} pendant que vous l''aviez ouvert. Enregistrer quand même vos modifications par-dessus les siennes ?'
  someoneElse = 'Quelqu''un d''autre'
  liveUpdated = '↻ Liste mise à jour avec des changements d''un autre PC'
  loading = 'Chargement des données…'
  savingMsg = 'Enregistrement…'
  errStoreBusy = 'Un autre PC enregistre en ce moment. Réessayez dans quelques secondes.'
  errNoStore = 'Données de GESTEM introuvables. Cliquez sur « Actualiser » dans Clients.'
  errPendingSync = 'Des clients enregistrés dans GESTEM ne sont pas encore dans l''Excel. Ouvrez Excel sur ce PC et cliquez sur « Actualiser » avant de réimporter.'
  importAsk = 'Ceci relit TOUS les clients depuis l''Excel et remplace les données de GESTEM par celles de l''Excel. À utiliser seulement si l''Excel a été modifié à la main. Continuer ?'
  importDone = '✓ Données réimportées depuis l''Excel'
  savedNoExcel = '✓ Enregistré dans GESTEM (dossier {0}). L''Excel sera mis à jour dès qu''il pourra être ouvert.'
  btnImport = 'Réimporter depuis l''Excel'
  dataNote = 'GESTEM garde les clients dans son propre dossier « GESTEM datos », à côté de l''Excel, et met l''Excel à jour à chaque enregistrement.'
  calToday = 'Aujourd''hui'
  btnDocs = 'Documents'
  btnReceipt = 'Reçu PDF'
  btnExport = 'Exporter'
  expPdf = 'Liste en PDF'
  expCsv = 'Liste en CSV (pour Excel)'
  colDocs = 'Docs'
  docsTitle = 'Documents'
  docsSub = 'Dossier {0}  ·  {1}'
  docsAdd = 'Ajouter des fichiers…'
  docsOpen = 'Ouvrir'
  docsRemove = 'Retirer'
  docsFolder = 'Ouvrir le dossier'
  docsClose = 'Fermer'
  docsEmpty = 'Aucun document pour ce client. Glissez ici les scans (PDF, photos…) ou cliquez sur « Ajouter des fichiers ».'
  docsDrop = 'Déposez les fichiers pour les enregistrer'
  docsCount = '{0} document(s)'
  docsAdded = '✓ {0} document(s) ajouté(s)'
  docsPick = 'Choisissez d''abord un document dans la liste.'
  docsRemoveAsk = 'Retirer « {0} » ? Il ira dans la corbeille de Windows.'
  docsFilter = 'Documents et photos'
  docsAll = 'Tous les fichiers'
  noPdfPrinter = 'L''imprimante Windows « Microsoft Print to PDF » est absente. Activez-la dans : Panneau de configuration → Programmes → Activer ou désactiver des fonctionnalités Windows.'
  pdfFail = 'Windows n''a pas créé le PDF.'
  rcTitle = 'REÇU'
  rcSubtitle = 'Reçu de paiement'
  rcNo = 'N° {0}'
  rcDate = 'Date : {0}'
  rcClient = 'Client'
  rcProc = 'Démarche'
  rcDoc = 'Pièce d''identité'
  rcCase = 'Dossier'
  rcResp = 'Traité par'
  rcType = 'Type'
  rcApplied = 'Demande'
  rcStatusP = 'État'
  rcPrice = 'Montant de la démarche'
  rcPaid = 'Payé'
  rcPend = 'Reste à payer'
  rcMethod = 'Mode de paiement'
  rcPaidFull = 'Payé'
  rcPartial = 'Paiement partiel'
  rcSign = 'Signature et cachet'
  rcNote = 'Ce document justifie le paiement reçu pour la démarche indiquée. Ce n''est pas une facture.'
  rcFooter = 'Créé avec GESTEM MANAGER'
  rcDone = '✓ Reçu {0} enregistré dans les documents du client'
  exTitle = 'Clients  ·  {0}'
  exFilters = 'Filtres'
  exNoFilter = 'Tous les clients'
  exTotal = '{0} clients  ·  {1} avec paiement en attente'
  exPage = 'Page {0} sur {1}'
  exEmpty = 'Aucun client dans la liste à exporter.'
  exDone = '✓ Exporté : {0}'
 }
 ar = @{
  conflictAsk = '{0} عدّل هذا العميل في {1} بينما كان مفتوحًا لديك. هل تريد حفظ تغييراتك فوق تغييراته على أي حال؟'
  someoneElse = 'شخص آخر'
  liveUpdated = '↻ تم تحديث القائمة بتغييرات من حاسوب آخر'
  loading = 'جارٍ تحميل البيانات…'
  savingMsg = 'جارٍ الحفظ…'
  errStoreBusy = 'حاسوب آخر يحفظ الآن. حاول مرة أخرى بعد ثوانٍ.'
  errNoStore = 'لم يُعثر على بيانات GESTEM. اضغط «تحديث» في العملاء.'
  errPendingSync = 'هناك عملاء محفوظون في GESTEM لم يصلوا بعد إلى ملف Excel. افتح Excel على هذا الحاسوب واضغط «تحديث» قبل إعادة الاستيراد.'
  importAsk = 'سيعيد هذا قراءة كل العملاء من ملف Excel ويستبدل بيانات GESTEM ببيانات Excel. استخدمه فقط إذا عُدّل ملف Excel يدويًا. هل تريد المتابعة؟'
  importDone = '✓ أُعيد استيراد البيانات من ملف Excel'
  savedNoExcel = '✓ حُفظ في GESTEM (الملف {0}). سيُحدَّث ملف Excel فور إمكان فتحه.'
  btnImport = 'إعادة الاستيراد من Excel'
  dataNote = 'يحفظ GESTEM العملاء في مجلده الخاص «GESTEM datos» بجانب ملف Excel، ويحدّث Excel عند كل حفظ.'
  calToday = 'اليوم'
  btnDocs = 'المستندات'
  btnReceipt = 'إيصال PDF'
  btnExport = 'تصدير'
  expPdf = 'القائمة بصيغة PDF'
  expCsv = 'القائمة بصيغة CSV (لـ Excel)'
  colDocs = 'مستندات'
  docsTitle = 'المستندات'
  docsSub = 'الملف {0}  ·  {1}'
  docsAdd = 'إضافة ملفات…'
  docsOpen = 'فتح'
  docsRemove = 'إزالة'
  docsFolder = 'فتح المجلد'
  docsClose = 'إغلاق'
  docsEmpty = 'لا توجد مستندات لهذا العميل بعد. اسحب الملفات الممسوحة هنا (PDF، صور…) أو اضغط «إضافة ملفات».'
  docsDrop = 'أفلت الملفات لحفظها'
  docsCount = '{0} مستند'
  docsAdded = '✓ تمت إضافة {0} مستند'
  docsPick = 'اختر مستندًا من القائمة أولًا.'
  docsRemoveAsk = 'إزالة «{0}»؟ سينتقل إلى سلة محذوفات Windows.'
  docsFilter = 'مستندات وصور'
  docsAll = 'كل الملفات'
  noPdfPrinter = 'طابعة Windows «Microsoft Print to PDF» غير موجودة. فعّلها من: لوحة التحكم ← البرامج ← تشغيل ميزات Windows أو إيقاف تشغيلها.'
  pdfFail = 'لم ينشئ Windows ملف PDF.'
  rcTitle = 'إيصال'
  rcSubtitle = 'إيصال دفع'
  rcNo = 'رقم {0}'
  rcDate = 'التاريخ: {0}'
  rcClient = 'العميل'
  rcProc = 'الإجراء'
  rcDoc = 'الوثيقة'
  rcCase = 'الملف'
  rcResp = 'المسؤول'
  rcType = 'النوع'
  rcApplied = 'الطلب'
  rcStatusP = 'الحالة'
  rcPrice = 'مبلغ الإجراء'
  rcPaid = 'المدفوع'
  rcPend = 'المتبقي'
  rcMethod = 'طريقة الدفع'
  rcPaidFull = 'مدفوع'
  rcPartial = 'دفع جزئي'
  rcSign = 'التوقيع والختم'
  rcNote = 'تثبت هذه الوثيقة الدفعة المستلمة عن الإجراء المذكور. وهي ليست فاتورة.'
  rcFooter = 'أُنشئ بواسطة GESTEM MANAGER'
  rcDone = '✓ تم حفظ الإيصال {0} في مستندات العميل'
  exTitle = 'العملاء  ·  {0}'
  exFilters = 'المرشحات'
  exNoFilter = 'كل العملاء'
  exTotal = '{0} عملاء  ·  {1} عليهم دفع'
  exPage = 'الصفحة {0} من {1}'
  exEmpty = 'لا يوجد عملاء في القائمة للتصدير.'
  exDone = '✓ تم التصدير: {0}'
 }
}
foreach ($lg in $script:TextMore.Keys) { foreach ($k in $script:TextMore[$lg].Keys) { $script:Text[$lg][$k] = $script:TextMore[$lg][$k] } }
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
    $cfg = [ordered]@{ ExcelPath = ''; Language = 'es'; DefaultNationality = 'MARROQUI'; Backup = $true; OpenAfterSave = $false; Tramite = 'canje'; TabsChecked = @(); Colleagues = @('SAFAE', 'CHAYMAE'); LastUser = ''; CompanyName = ''; CompanyLogo = ''; ExtBackupDir = ''; LastExtBackup = '' }
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

function ConvertTo-SheetKey([string]$s) {
    # nombre de pestaña sin tildes, mayúsculas ni espacios de más («Homologacion » = «HOMOLOGACIÓN»)
    $d = "$s".Trim().ToUpperInvariant().Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object Text.StringBuilder
    foreach ($ch in $d.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($ch) }
    }
    return ($sb.ToString() -replace '\s+', ' ')
}
function Test-ExcelBusyError($e) {
    $x = $e; while ($x) { if ($x.HResult -eq -2147418111 -or $x.HResult -eq -2147417846) { return $true }; $x = $x.InnerException }
    return $false
}
# otros nombres con los que la gestoría tiene las pestañas en su Excel
$script:SheetAliases = @{
    nacionalidad = @('NACIONALIDAD ESPAÑOLA'); autonomo = @('AUTÓNOMOS'); certificado = @('CERT DIG', 'CERTIFICADO')
    dispensa = @('DISPENSA'); visado = @('VISADO'); extranjeria = @('EXTRANJERÍA'); homologacion = @('HOMOLOGACION')
}
function Test-GestemLayout($ws, $t) {
    # ¿la pestaña tiene el formato de GESTEM? (cabecera «Cliente»/«Nombre…» en la columna del nombre, fila 10)
    # las pestañas hechas a mano (nombres en la columna A desde la fila 2…) NO se leen ni se escriben nunca
    if (-not (Test-GenTramite $t)) { return $true }
    $hdr = ''
    try { $hdr = ConvertTo-SheetKey "$($ws.Range("$($t.NameCol)$($t.FirstRow - 1)").Value2)" } catch { if (Test-ExcelBusyError $_.Exception) { throw }; return $false }
    return ($hdr.StartsWith('CLIENTE') -or $hdr.StartsWith('NOMBRE'))
}
function Find-TramiteSheet($wb, $t) {
    # candidatas: nombre de GESTEM, sus alias y pestañas que empiecen igual («AUTÓNOMOS», «SIS GESTEM»…)
    # solo vale la primera que tenga el formato de GESTEM
    $cands = New-Object System.Collections.Generic.List[object]
    $ws = Find-Sheet $wb $t.Sheet; if ($ws) { $cands.Add($ws) }
    if ($script:SheetAliases -and $script:SheetAliases.ContainsKey($t.Id)) {
        foreach ($a in @($script:SheetAliases[$t.Id])) { if ($a) { $ws = Find-Sheet $wb $a; if ($ws) { $cands.Add($ws) } } }
    }
    $key = ConvertTo-SheetKey $t.Sheet
    foreach ($w in $wb.Worksheets) {
        $k = ConvertTo-SheetKey "$($w.Name)"
        if ($k.StartsWith($key)) { $rest = $k.Substring($key.Length); if ($rest -eq 'S' -or $rest -eq 'ES' -or $rest.StartsWith(' ')) { $cands.Add($w) } }
    }
    foreach ($c in $cands) { if (Test-GestemLayout $c $t) { return $c } }
    return $null
}
function Find-Sheet($wb, [string]$name) {
    # busca la pestaña aunque tenga otra tilde, mayúsculas o espacios; $null si de verdad no está
    # (si Excel está ocupado se lanza el error de «ocupado», nunca «no tiene la pestaña»)
    try { return $wb.Worksheets.Item($name) } catch { if (Test-ExcelBusyError $_.Exception) { throw } }
    $key = ConvertTo-SheetKey $name
    foreach ($w in $wb.Worksheets) { if ((ConvertTo-SheetKey "$($w.Name)") -eq $key) { return $w } }
    return $null
}
function Test-GenTramite($t) { return ($t.Id -ne 'canje') }
function Get-Sheet($h, [switch]$Create) {
    # con -Create (solo al escribir un cliente): si al Excel le falta la pestaña del trámite, GESTEM la crea
    $T = $script:T
    $ws = Invoke-Excel { Find-TramiteSheet $h.Wb $T }
    if ($ws) { return $ws }
    if ($Create -and (Test-GenTramite $T)) { return (Invoke-Excel { New-TramiteSheet $h $T }) }
    $names = @(); try { $names = @($h.Wb.Worksheets | ForEach-Object { "«$($_.Name)»" }) } catch { }
    $msg = (L 'errNoSheet') -f $T.Sheet
    if ($names.Count) { $msg += "`r`n`r`n" + ((L 'errTabsFound') -f ($names -join ', ')) }
    $script:MissingTabs[$T.Id] = (Get-ExcelStamp)
    throw $msg
}
$script:GenStatusList = @('Documentación', 'Presentado', 'En Trámite', 'Requerimiento', 'Favorable', 'Desfavorable', 'Finalizado')
function Get-DefaultIdPrefix($t) {
    $k = (ConvertTo-SheetKey $t.Sheet) -replace '[^A-Z]', ''
    if (-not $k) { return '' }
    return ($k.Substring(0, [Math]::Min(3, $k.Length)) + '-')
}
function New-EmptyStore($t, $h) {
    # trámite sin pestaña en el Excel: GESTEM empieza con la lista vacía (el Excel es solo una copia)
    $price = $null
    if ($h -and $t.PriceKey) { try { $price = Read-TramitePrice $h.Wb $t.PriceKey } catch { } }
    $meta = @{ Lists = @{ 'excel:N' = $script:GenStatusList; 'excel:M' = @('Pagado', 'Parcial', 'Pendiente') }
               Tariff = @{}; TariffNet = @{}; DefaultPrice = $price; IdPrefix = (Get-DefaultIdPrefix $t); IdWidth = 3 }
    return @{ Meta = $meta; Clients = (New-Object System.Collections.Generic.List[object]) }
}
function New-TramiteSheet($h, $t) {
    # crea en el Excel la pestaña del trámite (cabeceras, nº de expediente y fórmulas básicas)
    $wb = $h.Wb
    # si ya hay una pestaña suya con ese nombre (hecha a mano), la de GESTEM se llama «… GESTEM» y la suya no se toca
    $name = $t.Sheet
    $taken = @(foreach ($w in $wb.Worksheets) { ConvertTo-SheetKey "$($w.Name)" })
    if ($taken -contains (ConvertTo-SheetKey $name)) { $name = "$($t.Sheet) GESTEM" }
    $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
    $ws.Name = $name
    $ws.Range('A1').Value2 = "$($t.Title['es'])".ToUpper()
    $ws.Range('A1').Font.Bold = $true; $ws.Range('A1').Font.Size = 14
    $ws.Range('A2').Value2 = 'Pestaña creada por GESTEM MANAGER'
    $hr = $t.FirstRow - 1
    $heads = @{ A = 'ID Expediente'; L = 'Pendiente (€)'; T = 'Días'; W = 'Neto (€)' }
    foreach ($f in $t.Fields) { $heads[$f.Col] = "$($f.Label['es'])" }
    foreach ($k in $heads.Keys) { $ws.Range("$k$hr").Value2 = $heads[$k] }
    $hdr = $ws.Range("A$($hr):$($t.LastCol)$hr"); $hdr.Font.Bold = $true; $hdr.Interior.Color = 14277081
    $f1 = $t.FirstRow; $l1 = $t.LastRow
    $meta = $null; try { $s = Read-Store $t; if ($s) { $meta = $s.Meta } } catch { }
    if (-not $meta) { $meta = (New-EmptyStore $t $null).Meta }
    $n = $l1 - $f1 + 1
    $ids = New-Object 'object[,]' $n, 1
    for ($i = 0; $i -lt $n; $i++) { $ids[$i, 0] = Get-RowId $t $meta ($f1 + $i) }
    $ws.Range("A$($f1):A$l1").Value2 = $ids
    $ws.Range("L$($f1):L$l1").Formula = "=IF(J$f1=`"`",`"`",IF(M$f1=`"Pagado`",0,MAX(0,J$f1-N(K$f1))))"
    $ws.Range("T$($f1):T$l1").Formula = "=IF(OR(O$f1=`"`",N$f1=`"Finalizado`",N$f1=`"Desfavorable`"),`"`",TODAY()-O$f1)"
    $ws.Range("W$($f1):W$l1").Formula = "=IF(J$f1=`"`",`"`",J$f1-N(V$f1))"
    foreach ($c in 'O', 'P') { $ws.Range("$c$($f1):$c$l1").NumberFormat = 'dd/mm/yyyy' }
    foreach ($c in 'J', 'K', 'L', 'V', 'W') { $ws.Range("$c$($f1):$c$l1").NumberFormat = '#,##0.00 "€"' }
    $ws.Range("T$($f1):T$l1").NumberFormat = '0'
    $ws.Columns("A:$($t.LastCol)").ColumnWidth = 16
    $ws.Columns('B').ColumnWidth = 32; $ws.Columns('R').ColumnWidth = 40
    Write-Log "pestaña creada en el Excel: $($t.Sheet)"
    return $ws
}
# =====================================================================
#  PESTAÑAS HECHAS A MANO  · © Yano
#  Si el Excel de la gestoría tiene un trámite en una pestaña propia (nombres en la columna A, pagos escritos
#  como «PAGADO 70€»…), GESTEM la lee tal cual, directamente del archivo (sin abrir Excel), y la vuelve a leer
#  cada vez que se guarda el Excel. Esos clientes se ven en GESTEM pero se cambian en el Excel:
#  GESTEM nunca escribe en una pestaña hecha a mano.
# =====================================================================
$script:HandStamp   = $null
$script:HandTabs    = @{}       # trámite -> pestaña hecha a mano que se está leyendo
$script:HandRowBase = 100000    # clientes de una pestaña hecha a mano: fila = 100000 + fila del Excel (solo lectura)

function Read-XlsxSheets([string]$path, [scriptblock]$Pick) {
    # lee el .xlsx directamente (aunque esté abierto en Excel). $Pick recibe todos los nombres y devuelve los que hay que leer.
    # devuelve @{ '__names' = nombres; '<pestaña>' = @{ fila = @{ columna = valor } } }
    try { Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop } catch { }
    $out = @{}
    $fs = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try {
        $zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Read)
        $rd = {
            param($n)
            $e = $zip.GetEntry($n); if (-not $e) { return $null }
            $sr = New-Object IO.StreamReader($e.Open())
            try { $x = New-Object Xml.XmlDocument; $x.LoadXml($sr.ReadToEnd()); return $x } finally { $sr.Dispose() }
        }
        $shared = New-Object System.Collections.Generic.List[string]
        $ss = & $rd 'xl/sharedStrings.xml'
        if ($ss) { foreach ($si in $ss.SelectNodes("//*[local-name()='si']")) { $shared.Add((@($si.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.InnerText }) -join '')) } }
        $wbx = & $rd 'xl/workbook.xml'
        $rels = & $rd 'xl/_rels/workbook.xml.rels'
        $target = @{}
        foreach ($r in $rels.SelectNodes("//*[local-name()='Relationship']")) { $target[$r.GetAttribute('Id')] = $r.GetAttribute('Target') }
        $sheets = @()
        foreach ($sh in $wbx.SelectNodes("//*[local-name()='sheet']")) {
            $rid = ($sh.Attributes | Where-Object { $_.LocalName -eq 'id' } | Select-Object -First 1).Value
            $sheets += [pscustomobject]@{ Name = $sh.GetAttribute('name'); Path = "$($target[$rid])" }
        }
        $all = @($sheets | ForEach-Object { $_.Name })
        $out['__names'] = $all
        $want = @(& $Pick $all)
        foreach ($s in $sheets) {
            if ($want -notcontains $s.Name) { continue }
            $p = $s.Path.TrimStart('/'); if (-not $p.StartsWith('xl/')) { $p = 'xl/' + $p }
            $e = $zip.GetEntry($p); if (-not $e) { continue }
            # lectura en flujo: las pestañas tienen miles de celdas vacías con formato; solo se miran las que tienen valor
            $cells = @{}
            $xr = [Xml.XmlReader]::Create($e.Open())
            try {
                $ref = $null; $ct = $null; $inIs = $false; $buf = $null; $more = $xr.Read()
                while ($more) {
                    $next = $true    # ReadElementContentAsString ya avanza: entonces no se vuelve a leer
                    if ($xr.NodeType -eq [Xml.XmlNodeType]::Element) {
                        switch ($xr.LocalName) {
                            'c' { $ref = $xr.GetAttribute('r'); $ct = $xr.GetAttribute('t'); $buf = $null; if ($xr.IsEmptyElement) { $ref = $null } }
                            'v' { if ($ref) { $buf = $xr.ReadElementContentAsString(); $next = $false } }
                            'is' { $inIs = $true; $buf = '' }
                            't' { if ($inIs -and $ref) { $buf += $xr.ReadElementContentAsString(); $next = $false } }
                        }
                    } elseif ($xr.NodeType -eq [Xml.XmlNodeType]::EndElement) {
                        if ($xr.LocalName -eq 'is') { $inIs = $false }
                        elseif ($xr.LocalName -eq 'c' -and $ref) {
                            if ($null -ne $buf -and $ref -match '^([A-Z]+)(\d+)$') {
                                $col = ColNum $Matches[1]; $row = [int]$Matches[2]; $v = $null
                                if ($row -gt 10 -and $cells.ContainsKey(10) -and $cells[10].ContainsKey(2)) {
                                    # pestaña con el formato de GESTEM (cabecera en B10): no hace falta leer más
                                    $h10 = ConvertTo-SheetKey "$($cells[10][2])"
                                    if ($h10.StartsWith('CLIENTE') -or $h10.StartsWith('NOMBRE')) { break }
                                }
                                if ($ct -eq 's') { $v = $shared[[int]$buf] }
                                elseif ($ct -eq 'str' -or $ct -eq 'inlineStr') { $v = $buf }
                                elseif ($ct -eq 'b' -or $ct -eq 'e') { $v = $null }
                                else { $d = 0.0; if ([double]::TryParse($buf, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { $v = $d } else { $v = $buf } }
                                if ($null -ne $v -and "$v".Trim() -ne '') {
                                    if (-not $cells.ContainsKey($row)) { $cells[$row] = @{} }
                                    $cells[$row][$col] = $v
                                }
                            }
                            $ref = $null
                        }
                    }
                    if ($next) { $more = $xr.Read() } else { $more = -not $xr.EOF }
                }
            } finally { $xr.Dispose() }
            $out[$s.Name] = $cells
        }
        $zip.Dispose()
    } finally { $fs.Dispose() }
    return $out
}

function Get-TabCandidates([string[]]$names, $t) {
    # pestañas que pueden ser de este trámite: su nombre, sus alias y las que empiezan igual
    $res = New-Object System.Collections.Generic.List[string]
    $keys = @{}; foreach ($n in $names) { $keys[$n] = ConvertTo-SheetKey $n }
    $wanted = @($t.Sheet)
    if ($script:SheetAliases -and $script:SheetAliases.ContainsKey($t.Id)) { $wanted += @($script:SheetAliases[$t.Id]) }
    foreach ($w in $wanted) { $k = ConvertTo-SheetKey $w; foreach ($n in $names) { if ($keys[$n] -eq $k -and -not $res.Contains($n)) { $res.Add($n) } } }
    $key = ConvertTo-SheetKey $t.Sheet
    foreach ($n in $names) {
        $k = $keys[$n]
        if ($k.StartsWith($key)) { $rest = $k.Substring($key.Length); if (($rest -eq 'S' -or $rest -eq 'ES' -or $rest.StartsWith(' ')) -and -not $res.Contains($n)) { $res.Add($n) } }
    }
    return $res.ToArray()
}
function Test-CellsGestem($cells, $t) {
    # misma prueba que Test-GestemLayout pero sobre las celdas leídas del archivo
    $r = $cells[[int]($t.FirstRow - 1)]; if (-not $r) { return $false }
    $h = ConvertTo-SheetKey "$($r[(ColNum $t.NameCol)])"
    return ($h.StartsWith('CLIENTE') -or $h.StartsWith('NOMBRE'))
}
function ConvertTo-HandKey([string]$s) { return ((ConvertTo-SheetKey $s) -replace '[^A-Z0-9 ]', ' ' -replace '\s+', ' ').Trim() }
function Get-HandRole([string]$k, $t) {
    # qué es cada columna según su cabecera (sin tildes ni mayúsculas)
    if ($k -match '^(NOMBRE|CLIENTE)') { return 'name' }
    if ($k -match '\b(NIE|DNI|PASAPORTE|DOCUMENTO)\b') { return 'doc' }
    if ($k -match 'TELEFONO|MOVIL|\bTLF\b') { return 'tel' }
    if ($k -match 'NACIONALIDAD') { return 'nat' }
    if ($k -match '^(NO PAGADO|PENDIENTE|PENDIETE)') { return 'pend' }
    if ($k -match '^(PAGADO|HA PAGADO|COBRADO|A CUENTA)$') { return 'paid' }
    if ($k -match '^PAGOS?$|FORMA DE PAGO') { return 'paytext' }
    if ($k -match 'PRECIO|IMPORTE|HONORARIOS|^TOTAL') { return 'price' }
    if ($k -match 'OBSERVACION|NOTAS|COMENTARIO') { return 'obs' }
    if ($k -match 'ESTADO|SITUACION') { return 'status' }
    if ($k -match 'RESOLUCION') { return 'fres' }
    if ($k -match 'FECHA|PRESENTACION|SOLICITUD') { return 'fsol' }
    # campos propios del trámite (Contraseña, Tipo de gestión, Relevé de notes…), también por sus opciones (BACHILLERATO -> Título)
    $kw = @($k.Split(' ') | Where-Object { $_.Length -ge 5 })
    foreach ($f in $t.Fields) {
        if ($f.Key -notlike 'x*') { continue }
        $lb = ConvertTo-HandKey "$($f.Label['es'])"
        if ($k -eq $lb -or $lb.Contains($k) -or $k.Contains($lb)) { return $f.Key }
        foreach ($w in $kw) { if (@($lb.Split(' ')) -contains $w) { return $f.Key } }
        if ("$($f.List)" -like 'fixed:*') { foreach ($o in $f.List.Substring(6).Split('|')) { if ((ConvertTo-HandKey $o) -eq $k) { return $f.Key } } }
    }
    return $null
}
function Get-HandLayout($cells, $t) {
    # fila de cabecera: una de las 5 primeras con «NOMBRE…» o «CLIENTE»; sin cabecera, el nombre es la columna A
    $hdr = 0
    for ($r = 1; $r -le 5 -and -not $hdr; $r++) {
        if (-not $cells.ContainsKey($r)) { continue }
        foreach ($c in $cells[$r].Keys) { if ((ConvertTo-HandKey "$($cells[$r][$c])") -match '^(NOMBRE|CLIENTE)') { $hdr = $r; break } }
    }
    $map = @{}; $pres = $false
    if ($hdr) {
        foreach ($c in ($cells[$hdr].Keys | Sort-Object)) {
            $k = ConvertTo-HandKey "$($cells[$hdr][$c])"
            if (-not $k) { continue }
            $role = Get-HandRole $k $t
            if ($role -and @($map.Values) -contains $role -and $role -notin 'obs', 'paytext', 'status') { $role = $null }    # la primera columna gana
            $map[[int]$c] = $role
            if ($role -eq 'fsol' -and $k -match 'PRESENTACION') { $pres = $true }
            if (-not $role) { $map[[int]$c] = '#' + "$($cells[$hdr][$c])".Trim() }    # cabecera desconocida: va a observaciones con su nombre
        }
    }
    $nameCol = 1
    foreach ($e in $map.GetEnumerator()) { if ($e.Value -eq 'name') { $nameCol = [int]$e.Key; break } }
    $map[$nameCol] = 'name'
    return @{ HeaderRow = $hdr; Map = $map; NameCol = $nameCol; PresDate = $pres }
}
function Format-HandText($v) {
    if ($v -is [double]) { return $(if ($v -eq [Math]::Floor($v)) { ([long]$v).ToString() } else { $v.ToString('0.##', [Globalization.CultureInfo]::InvariantCulture) }) }
    return ("$v" -replace '\s+', ' ').Trim()
}
function ConvertFrom-HandDate($v) {
    # fecha de Excel (número) o «dd/mm/aaaa» dentro de un texto
    if ($v -is [double]) { if ($v -ge 20000 -and $v -le 80000) { return [DateTime]::FromOADate($v).Date }; return $null }
    if ("$v" -match '(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})') {
        try { return (New-Object DateTime ([int]$Matches[3]), ([int]$Matches[2]), ([int]$Matches[1])) } catch { }
    }
    return $null
}
function Get-HandNumber($v) {
    if ($v -is [double]) { return [double]$v }
    if ("$v" -match '(\d+(?:[.,]\d{1,2})?)') { return (Parse-Money $Matches[1]) }
    return $null
}
function Convert-HandStatus([string[]]$texts, [bool]$hasPresDate) {
    # estado de GESTEM a partir de lo que escriben (los estados «fuertes» primero)
    $all = ConvertTo-HandKey ($texts -join ' ')
    if (-not $all -and -not $hasPresDate) { return $null }
    if ($all -match 'FINALIZ|ENTREGAD|TERMINAD') { return 'Finalizado' }
    if ($all -match 'DESFAVORABLE|DENEGAD') { return 'Desfavorable' }
    if ($all -match 'FAVORABLE|CONCEDID|APROBAD') { return 'Favorable' }
    if ($all -match 'REQUERIM') { return 'Requerimiento' }
    if ($all -match 'PRESENTAD') { return 'Presentado' }
    if ($all -match 'TRAMIT|EN GESTION') { return 'En Trámite' }
    if ($hasPresDate) { return 'Presentado' }
    if ($all -match 'ESPERA|PREPARAD|DOCUMENTAC|FALTA|PENDIENTE PARA') { return 'Documentación' }
    return $null
}
function ConvertFrom-HandRow($row, $lay, $t) {
    $FV = @{}; foreach ($f in $t.Fields) { $FV[$f.Key] = $null }
    $notes = New-Object System.Collections.Generic.List[string]
    $stTexts = New-Object System.Collections.Generic.List[string]
    $payTexts = New-Object System.Collections.Generic.List[string]
    $price = $null; $paid = $null; $pend = $null; $presDate = $false
    foreach ($c in ($row.Keys | Sort-Object)) {
        $v = $row[$c]
        $role = $lay.Map[[int]$c]
        if ($role -eq 'name') { continue }
        $txt = Format-HandText $v
        if (-not $txt) { continue }
        switch -Wildcard ($role) {
            'doc'     { $FV['doc'] = $txt.ToUpper() }
            'tel'     { $FV['tel'] = $txt }
            'nat'     { $FV['nat'] = $txt.ToUpper() }
            'x*'      { $FV[$role] = $txt }
            'price'   { $n = Get-HandNumber $v; if ($null -ne $n) { $price = $n } else { $notes.Add($txt) } }
            'paid'    { $n = Get-HandNumber $v; if ($null -ne $n) { $paid = $n } else { $payTexts.Add($txt); $notes.Add($txt) } }
            'pend'    { $n = Get-HandNumber $v; if ($null -ne $n) { $pend = $n } else { $notes.Add($txt) } }
            'paytext' { $payTexts.Add($txt); $notes.Add("Pago: $txt") }
            'status'  { $stTexts.Add($txt); $notes.Add($txt) }
            'obs'     { $notes.Add($txt) }
            { $_ -eq 'fsol' -or $_ -eq 'fres' } {
                $d = ConvertFrom-HandDate $v
                if ($d) { $FV[$role] = $d; if ($role -eq 'fsol' -and $lay.PresDate) { $presDate = $true } }
                $rest = ("$txt" -replace '\d{1,2}[/\-.]\d{1,2}[/\-.]\d{4}', '').Trim(' ', '-')
                if (-not $d -or $rest) { $stTexts.Add($txt); $notes.Add($txt) }
            }
            '#*'      { $notes.Add("$($role.Substring(1)): $txt"); $stTexts.Add($txt) }
            default {
                # columna sin cabecera: fecha, estado o nota
                $d = ConvertFrom-HandDate $v
                if ($d -and -not $FV['fsol'] -and ($v -is [double] -or $txt -match '^\d{1,2}[/\-.]\d{1,2}[/\-.]\d{4}$')) { $FV['fsol'] = $d }
                else { $stTexts.Add($txt); $notes.Add($txt) }
            }
        }
    }
    # pago
    foreach ($p in $payTexts) {
        $k = ConvertTo-HandKey $p; $n = Get-HandNumber $p
        if ($k -match 'NO PAGAD|PENDIENTE') { if ($null -eq $price -and $null -ne $n) { $price = $n }; if ($null -eq $paid) { $paid = 0.0 } }
        elseif ($k -match 'PAGAD') { if ($null -eq $price -and $null -ne $n) { $price = $n }; if ($null -eq $paid -and $null -ne $price) { $paid = $price } }
        elseif ($null -eq $price -and $null -ne $n) { $price = $n }
    }
    if ($null -eq $price -and $null -eq $paid -and $null -eq $pend) {
        # «255 EUROS PAGADO …» escrito en otra columna
        foreach ($s in $notes) { if ((ConvertTo-HandKey $s) -match '(\d+(?:[.,]\d{1,2})?)\s*(EUROS?|E)?\s*PAGAD') { $price = Parse-Money $Matches[1]; $paid = $price; break } }
    }
    if ($null -eq $price -and $null -ne $pend) { $price = $(if ($null -ne $paid) { $paid } else { 0.0 }) + $pend }
    $FV['price'] = $price
    if ($null -ne $price -and $null -ne $paid) {
        if ($paid -ge $price) { $FV['paystate'] = 'Pagado'; $FV['paid'] = $price }
        elseif ($paid -gt 0) { $FV['paystate'] = 'Parcial'; $FV['paid'] = $paid }
        else { $FV['paystate'] = 'Pendiente'; $FV['paid'] = 0.0 }
    } elseif ($null -ne $price) { $FV['paystate'] = 'Pendiente' }
    elseif ($null -ne $paid -and $paid -gt 0) { $FV['paystate'] = 'Parcial'; $FV['paid'] = $paid }    # pagó algo, precio desconocido
    $FV['status'] = Convert-HandStatus $stTexts.ToArray() $presDate
    if ($notes.Count) { $FV['obs'] = ($notes | Select-Object -Unique) -join ' · ' }
    return $FV
}
function ConvertFrom-HandSheet($cells, $t) {
    $lay = Get-HandLayout $cells $t
    $prefix = Get-DefaultIdPrefix $t
    $list = New-Object System.Collections.Generic.List[object]
    if (-not $cells.Count) { return ,$list }
    $maxRow = ($cells.Keys | Measure-Object -Maximum).Maximum
    $tname = ConvertTo-HandKey $t.Sheet
    $gap = 0
    for ($r = $lay.HeaderRow + 1; $r -le $maxRow; $r++) {
        $row = $cells[$r]
        $name = ''; if ($row) { $name = Format-HandText $row[$lay.NameCol] }
        if (-not $name) { $gap++; if ($list.Count -gt 0 -and $gap -ge 8) { break }; continue }
        $gap = 0
        $nk = ConvertTo-HandKey $name
        if ($nk -match '^(PRECIOS?|TOTAL(ES)?|TARIFAS?)\b') { break }        # tabla de precios debajo de los clientes
        if ($nk -eq $tname -or ($tname -and $nk.StartsWith($tname))) { continue }   # título («MATRICULACION»)
        $FV = ConvertFrom-HandRow $row $lay $t
        $FV['name'] = $name.ToUpper()
        $list.Add(@{ Row = $script:HandRowBase + $r; Id = $prefix + ([string]($r - $lay.HeaderRow)).PadLeft(3, '0'); F = $FV; Sync = $true; New = $false; Updated = ''; By = 'Excel' })
    }
    return ,$list
}
function Sync-HandmadeTabs([switch]$Force) {
    # lee las pestañas hechas a mano si el Excel ha cambiado desde la última vez; devuelve $true si cambió algo
    $p = "$($script:Cfg.ExcelPath)"
    if (-not $p -or -not (Test-Path -LiteralPath $p)) { return $false }
    $stamp = (Get-Item -LiteralPath $p).LastWriteTimeUtc.Ticks
    if (-not $Force -and $script:HandStamp -eq $stamp) { return $false }
    $gen = @($script:Tramites | Where-Object { Test-GenTramite $_ })
    $book = $null
    try {
        $book = Read-XlsxSheets $p {
            param($all)
            $w = @(); foreach ($t in $gen) { $w += @(Get-TabCandidates $all $t) }
            return $w
        }
    } catch { Write-Log "pestañas propias: no se pudo leer el Excel: $($_.Exception.Message)"; return $false }
    $changed = $false
    foreach ($t in $gen) {
        $src = $null
        foreach ($n in (Get-TabCandidates $book['__names'] $t)) { if ($book.ContainsKey($n) -and -not (Test-CellsGestem $book[$n] $t)) { $src = $n; break } }
        $recs = New-Object System.Collections.Generic.List[object]
        if ($src) { $recs = ConvertFrom-HandSheet $book[$src] $t; $script:HandTabs[$t.Id] = $src } else { [void]$script:HandTabs.Remove($t.Id) }
        $lk = Enter-StoreLock
        try {
            $s = Read-Store $t
            if (-not $s -and $recs.Count -eq 0) { continue }
            if (-not $s) { $s = New-EmptyStore $t $null }
            $sig = {
                param($list)
                (@($list | Sort-Object { [int]$_.Row } | ForEach-Object {
                    $c = $_; "$($c.Row)|$($c.Id)|" + ((@($t.Fields | ForEach-Object { "$(ConvertTo-StoreValue $_ $c.F[$_.Key])" })) -join '|')
                })) -join "`n"
            }
            $old = @($s.Clients | Where-Object { [int]$_.Row -ge $script:HandRowBase })
            if ((& $sig $old) -ne (& $sig $recs)) {
                $keep = New-Object System.Collections.Generic.List[object]
                foreach ($c in $s.Clients) { if ([int]$c.Row -lt $script:HandRowBase) { $keep.Add($c) } }
                foreach ($c in $recs) { $keep.Add($c) }
                $s.Clients = $keep
                Write-Store $t $s
                $changed = $true
                if ($src) { Write-Log ("leído del Excel (pestaña «{0}»): {1} clientes" -f $src, $recs.Count) }
            }
        } finally { Exit-StoreLock $lk }
    }
    $script:HandStamp = $stamp
    return $changed
}

# pestañas que no se encontraron: no se vuelve a abrir el Excel para buscarlas hasta que el Excel cambie
$script:MissingTabs = @{}
$script:MissingWarned = @{}
function Get-ExcelStamp {
    try { return (Get-Item -LiteralPath $script:Cfg.ExcelPath).LastWriteTimeUtc.Ticks } catch { return 0 }
}
function Test-TabKnownMissing($t) {
    if (-not $script:MissingTabs.ContainsKey($t.Id)) { return $false }
    if ($script:MissingTabs[$t.Id] -eq (Get-ExcelStamp)) { return $true }
    $script:MissingTabs.Remove($t.Id); $script:MissingWarned.Remove($t.Id)    # el Excel cambió: se vuelve a mirar
    return $false
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

function Read-TramitePrice($wb, [string]$key) {
    # precio del trámite en la tabla «PRECIOS POR TRÁMITE» de AJUSTES (vacío = sin precio)
    try { $aj = $wb.Worksheets.Item('AJUSTES') } catch { return $null }
    $cells = $aj.Range('B1:C200').Value2
    for ($r = 1; $r -le 200; $r++) {
        if ("$($cells[$r, 1])".Trim() -ieq $key -and ($cells[$r, 2] -is [double])) { return [double]$cells[$r, 2] }
    }
    return $null
}

function Read-TramiteData($h) {
    # lee una pestaña entera (la del trámite $script:T) y devuelve todo lo que necesita la ventana
    $T = $script:T
    $ws = Get-Sheet $h
    $lists = @{}
    foreach ($f in $T.Fields) {
        if ("$($f.List)" -like 'excel:*') {
            $c = $f.List.Substring(6)
            $lists[$f.List] = Invoke-Excel { Read-ValidationList $ws $c }
        }
    }
    $tariff = @{}
    if ($T.TipoCol) { $tariff = Invoke-Excel { Read-Tariff $ws $T.TipoCol } }
    $defPrice = $null
    if ($T.PriceKey) { $defPrice = Invoke-Excel { Read-TramitePrice $h.Wb $T.PriceKey } }
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
            Tipo = $(if ($rec.ContainsKey('tipo')) { "$($rec['tipo'])" } else { "$($rec['xF'])" }); Status = $status; Color = $color; Pay = "$($rec['paystate'])"
            Pending = $(if ($pend -is [double]) { $pend } else { 0 })
            PendingText = $(if ($pend -is [double] -and $pend -gt 0) { Format-Money $pend } else { '' })
            Days = $(if ($days -is [double]) { [int]$days } else { '' })
            Loc = "$($rec['loc'])"; Resp = "$($rec['resp'])"; Data = $rec
            PriceV = $(if ($vals[$i, 10] -is [double]) { $vals[$i, 10] } else { 0 })
            PaidV = $(if ($vals[$i, 11] -is [double]) { $vals[$i, 11] } else { 0 })
            NetV = $(if ($T.NetCol -and $vals[$i, (ColNum $T.NetCol)] -is [double]) { $vals[$i, (ColNum $T.NetCol)] } else { 0 })
            Fecha = $(if ($rec['fsol'] -is [DateTime]) { $rec['fsol'] } else { $null })
        })
    }
    $all = New-Object System.Collections.Generic.List[string]
    foreach ($p in $script:FormasPago) { $all.Add($p) }
    foreach ($p in $payExtra) { if (-not $all.Contains($p) -and $p.Length -le 30) { $all.Add($p) } }
    $lists['payments'] = $all.ToArray()
    $lists['nationalities'] = $script:Nacionalidades
    $col = New-Object System.Collections.Generic.List[string]
    foreach ($c in @($script:Cfg.Colleagues)) { if ("$c".Trim() -and -not $col.Contains("$c")) { $col.Add("$c") } }
    foreach ($r in $rows) { $x = "$($r.Resp)".Trim().ToUpper(); if ($x -and -not $col.Contains($x)) { $col.Add($x) } }
    $lists['colleagues'] = $col.ToArray()
    return @{ Lists = $lists; Tariff = $tariff; DefaultPrice = $defPrice; Rows = $rows; FirstEmpty = $firstEmpty; NextId = $nextId }
}
function Use-TramiteData($d) {
    $script:Lists = $d.Lists
    $script:Tariff = $d.Tariff
    $script:DefaultPrice = $d.DefaultPrice
    $script:Rows = $d.Rows
    $script:FirstEmpty = $d.FirstEmpty
    $script:NextId = $d.NextId
    Update-DocsMap
    Refresh-ClientGrid
    Update-StatusLine
}
function Get-CachedData {
    # cambiar de trámite: instantáneo si el Excel no ha cambiado desde la última lectura
    if ($script:Cache -and $script:Cache.ContainsKey($script:T.Id) -and $script:CacheStamp -eq (Get-FileStamp)) {
        Use-TramiteData $script:Cache[$script:T.Id]
        return $true
    }
    return (Load-Data -All)
}

function Restore-Formula($ws, [string]$col, [int]$row) {
    # vuelve a poner la fórmula automática de esa columna (copiándola de otra fila)
    # una sola llamada a Excel para encontrar las fórmulas (antes se miraban las 1000 filas una a una)
    $T = $script:T
    $fx = $null
    # devuelve $false si esa columna no tiene fórmulas (pestaña creada por GESTEM): entonces se escribe el valor
    try { $fx = $ws.Range("$col$($T.FirstRow):$col$($T.LastRow)").SpecialCells(-4123) } catch { if (Test-ExcelBusyError $_.Exception) { throw }; return $false }   # -4123 = celdas con fórmula
    foreach ($src in $fx.Cells) {
        if ($src.Row -ne $row) { $ws.Range("$col$row").FormulaR1C1 = $src.FormulaR1C1; return $true }
    }
    return [bool]$ws.Range("$col$row").HasFormula    # la única fórmula es la de esa fila: se deja
}

function ConvertTo-CellNumber($v) {
    if ($null -eq $v -or "$v" -eq '') { return $null }
    if ($v -is [double] -or $v -is [int] -or $v -is [long] -or $v -is [decimal] -or $v -is [single]) { return [double]$v }
    return (Parse-Money "$v")
}
function ConvertTo-CellDate($v) {
    # devuelve la fecha como número de Excel (o $null)
    if ($null -eq $v -or "$v" -eq '') { return $null }
    if ($v -is [DateTime]) { return $v.ToOADate() }
    if ($v -is [double] -or $v -is [int] -or $v -is [decimal]) { return [double]$v }
    $d = [DateTime]::MinValue
    foreach ($fmt in 'yyyy-MM-dd', 'dd/MM/yyyy', 'yyyy-MM-ddTHH:mm:ss') {
        if ([DateTime]::TryParseExact("$v", $fmt, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$d)) { return $d.ToOADate() }
    }
    return $null
}
function Set-CellValue($ws, [string]$addr, $value, [string]$type) {
    $rng = $ws.Range($addr)
    if ($null -eq $value -or ("$value" -eq '')) { $rng.ClearContents() | Out-Null; return }
    switch ($type) {
        'Date'  { $rng.Value2 = [double]$value }     # ya convertida a número de Excel
        'Money' { $rng.Value2 = [double]$value }
        default { $rng.Value2 = "'" + "$value" }     # el apóstrofo obliga a guardarlo como texto
    }
}

function Save-Client([hashtable]$vals, $row, [string]$ExpectName, [switch]$IsNew, $Handle) {
    # escribe un cliente en su fila del Excel (con -Handle usa un Excel ya abierto)
    $T = $script:T
    $h = $Handle; $own = -not $Handle
    try {
        if ($own) { $h = Invoke-Excel { Open-Book -ForWrite } }
        $ws = Get-Sheet $h -Create
        if ($own -and $script:Cfg.Backup) { Backup-Excel }
        if ($row -and $IsNew) {
            # si esa fila del Excel ya tiene otro cliente (escrito a mano), se usa la primera libre
            $curName = "$($ws.Range("$($T.NameCol)$row").Value2)".Trim()
            if ($curName -and $curName -ne "$ExpectName".Trim()) { $row = $null }
        }
        if (-not $row) {
            $col = Invoke-Excel { ,($ws.Range("$($T.NameCol)$($T.FirstRow):$($T.NameCol)$($T.LastRow)").Value2) }
            $n = $T.LastRow - $T.FirstRow + 1
            for ($i = 1; $i -le $n; $i++) { if (-not "$($col[$i, 1])".Trim()) { $row = $T.FirstRow + $i - 1; break } }
            if (-not $row) { throw (L 'errFull') }
        }
        $tipo = "$($vals['tipo'])"
        # 1) se preparan y comprueban TODOS los valores antes de tocar el Excel (nunca queda una fila a medias)
        $ops = New-Object System.Collections.Generic.List[object]
        foreach ($f in $T.Fields) {
            $v = $vals[$f.Key]
            $op = @{ Col = $f.Col; Kind = 'value'; Value = $null; Type = 'Text' }
            switch ($f.Role) {
                'nationality' {
                    if (-not "$v" -or "$v" -ieq $T.DefaultNationality) { $op.Kind = 'formula'; $op.Value = $T.DefaultNationality } else { $op.Value = "$v" }
                }
                'price' {
                    $tar = $null; $pv = ConvertTo-CellNumber $v
                    if ($tipo -and $script:Tariff.ContainsKey($tipo)) { $tar = [double]$script:Tariff[$tipo] }
                    if ($null -eq $pv -or ($null -ne $tar -and $pv -eq $tar)) { $op.Kind = 'formula'; $op.Value = $(if ($null -ne $pv) { $pv } else { $tar }); $op.Type = 'Money' } else { $op.Value = $pv; $op.Type = 'Money' }
                }
                'paid' {
                    $pv = ConvertTo-CellNumber $v
                    if ("$($vals['paystate'])" -eq 'Parcial' -and $null -ne $pv) { $op.Value = $pv; $op.Type = 'Money' } else { $op.Kind = 'formula'; $op.Value = $pv; $op.Type = 'Money' }
                }
                default {
                    if ($f.Type -eq 'Date') { $op.Value = ConvertTo-CellDate $v; $op.Type = 'Date' }
                    elseif ($f.Type -eq 'Money') { $op.Value = ConvertTo-CellNumber $v; $op.Type = 'Money' }
                    elseif ($null -ne $v -and "$v" -ne '') { $op.Value = "$v" }
                }
            }
            $ops.Add($op)
        }
        # 2) se escribe la fila
        Invoke-Excel {
            foreach ($op in $ops) {
                # sin fórmula en esa columna (pestaña creada por GESTEM): se escribe el valor calculado
                if ($op.Kind -eq 'formula' -and (Restore-Formula $ws $op.Col $row)) { continue }
                Set-CellValue $ws "$($op.Col)$row" $op.Value $op.Type
            }
        }
        $id = "$($ws.Range("$($T.IdCol)$row").Value2)"
        if ($own) { Invoke-Excel { Close-Book $h $true } }; $h = $null
        Write-Log "Excel: fila $row ($id)"
        return @{ Id = $id; Row = [int]$row }
    } catch {
        if ($own -and $h) { try { Close-Book $h $false } catch {} }
        throw
    }
}

# =====================================================================
#  DATOS PROPIOS DE GESTEM (la copia principal) · el Excel se actualiza en cada guardado · © Yano
# =====================================================================
function Get-StoreDir {
    $x = "$($script:Cfg.ExcelPath)"
    if (-not $x) { return $null }
    return (Join-Path (Split-Path -Parent $x) 'GESTEM datos')
}
function Get-StorePath($t) { $d = Get-StoreDir; if (-not $d) { return $null }; return (Join-Path $d ($t.Id + '.json')) }
function Get-FileStamp {
    # marca de cambios de los datos de GESTEM (si otro PC guarda, se vuelve a leer)
    $d = Get-StoreDir
    try {
        if ($d -and (Test-Path -LiteralPath $d)) {
            $m = Get-ChildItem -LiteralPath $d -Filter '*.json' -File | Measure-Object -Property LastWriteTimeUtc -Maximum
            if ($m.Maximum) { return ([DateTime]$m.Maximum).Ticks }
        }
    } catch { }
    return 0
}
function Enter-StoreLock {
    $d = Get-StoreDir
    if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Force -Path $d | Out-Null }
    $lk = Join-Path $d 'bloqueo.lock'
    for ($i = 0; $i -lt 50; $i++) {
        try {
            $fs = New-Object IO.FileStream($lk, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
            $b = [Text.Encoding]::UTF8.GetBytes("$($env:COMPUTERNAME) $($script:User) $(Get-Date -Format s)")
            $fs.Write($b, 0, $b.Length); $fs.Dispose()
            return $lk
        } catch {
            # un bloqueo de más de 1 minuto es de un programa que se cerró mal
            try { if (((Get-Date) - (Get-Item -LiteralPath $lk).LastWriteTime).TotalSeconds -gt 60) { Remove-Item -LiteralPath $lk -Force } } catch { }
            Start-Sleep -Milliseconds 200
        }
    }
    throw (L 'errStoreBusy')
}
function Exit-StoreLock($lk) { if ($lk) { try { Remove-Item -LiteralPath $lk -Force } catch { } } }

function ConvertTo-StoreValue($f, $v) {
    if ($null -eq $v -or "$v" -eq '') { return $null }
    if ($f.Type -eq 'Date') { if ($v -is [DateTime]) { return $v.ToString('yyyy-MM-dd') }; return "$v" }
    if ($f.Type -eq 'Money') { if ($v -is [double] -or $v -is [int] -or $v -is [decimal]) { return [double]$v }; $d = Parse-Money "$v"; return $d }
    if ($v -is [double]) { return $v }
    return "$v"
}
function ConvertFrom-StoreValue($f, $v) {
    if ($null -eq $v -or "$v" -eq '') { return $null }
    if ($f.Type -eq 'Date') {
        if ($v -is [DateTime]) { return $v }
        $d = [DateTime]::MinValue
        if ([DateTime]::TryParseExact("$v", 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
        return $null
    }
    if ($f.Type -eq 'Money') { return (ConvertTo-CellNumber $v) }    # texto en una columna de importes («CNTA …») = sin importe
    return $v
}
function Read-Store($t) {
    $p = Get-StorePath $t
    if (-not $p -or -not (Test-Path -LiteralPath $p)) { return $null }
    $j = [IO.File]::ReadAllText($p, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $meta = @{ Lists = @{}; Tariff = @{}; TariffNet = @{}; DefaultPrice = $null; IdPrefix = "$($j.meta.idPrefix)"; IdWidth = 3 }
    if ($j.meta.idWidth) { $meta.IdWidth = [int]$j.meta.idWidth }
    if ($j.meta.lists) { foreach ($pr in $j.meta.lists.PSObject.Properties) { $meta.Lists[$pr.Name] = @($pr.Value | ForEach-Object { "$_" }) } }
    if ($j.meta.tariff) { foreach ($pr in $j.meta.tariff.PSObject.Properties) { $meta.Tariff[$pr.Name] = [double]$pr.Value } }
    if ($j.meta.tariffNet) { foreach ($pr in $j.meta.tariffNet.PSObject.Properties) { $meta.TariffNet[$pr.Name] = [double]$pr.Value } }
    if ($null -ne $j.meta.defaultPrice) { $meta.DefaultPrice = [double]$j.meta.defaultPrice }
    $clients = New-Object System.Collections.Generic.List[object]
    foreach ($c in @($j.clients)) {
        if (-not $c) { continue }
        $FV = @{}
        foreach ($f in $t.Fields) { $FV[$f.Key] = ConvertFrom-StoreValue $f $c.f.($f.Key) }
        $clients.Add(@{ Row = [int]$c.row; Id = "$($c.id)"; F = $FV; Sync = [bool]$c.sync; New = [bool]$c.nuevo; Updated = "$($c.updated)"; By = "$($c.by)" })
    }
    return @{ Meta = $meta; Clients = $clients }
}
function Write-Store($t, $store) {
    $p = Get-StorePath $t
    $fields = @{}
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($c in ($store.Clients | Sort-Object { $_.Row })) {
        $o = [ordered]@{}
        foreach ($f in $t.Fields) { $o[$f.Key] = ConvertTo-StoreValue $f $c.F[$f.Key] }
        $list.Add([ordered]@{ row = $c.Row; id = $c.Id; sync = [bool]$c.Sync; nuevo = [bool]$c.New; updated = $c.Updated; by = $c.By; f = $o })
    }
    $m = $store.Meta
    $obj = [ordered]@{
        app = 'GESTEM MANAGER'; formato = 1; tramite = $t.Id; hoja = $t.Sheet; guardado = (Get-Date).ToString('s')
        meta = [ordered]@{ idPrefix = $m.IdPrefix; idWidth = $m.IdWidth; defaultPrice = $m.DefaultPrice; tariff = $m.Tariff; tariffNet = $m.TariffNet; lists = $m.Lists }
        clients = $list.ToArray()
    }
    $json = $obj | ConvertTo-Json -Depth 8
    $tmp = "$p.tmp"
    [IO.File]::WriteAllText($tmp, $json, (New-Object Text.UTF8Encoding($true)))
    if (Test-Path -LiteralPath $p) {
        # copia diaria del archivo de datos antes de cambiarlo
        try {
            $bk = Join-Path $script:BackupDir ("datos_{0}_{1:yyyy-MM-dd}.json" -f $t.Id, (Get-Date))
            if (-not (Test-Path -LiteralPath $bk)) { Copy-Item -LiteralPath $p -Destination $bk -Force }
        } catch { }
        try { [IO.File]::Replace($tmp, $p, [NullString]::Value) }
        catch { Copy-Item -LiteralPath $tmp -Destination $p -Force; Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    } else {
        Move-Item -LiteralPath $tmp -Destination $p -Force
    }
}

# --- valores que en el Excel calculan sus fórmulas ---
function Resolve-Record($t, [hashtable]$FV, $meta) {
    $natF = $t.Fields | Where-Object { $_.Role -eq 'nationality' } | Select-Object -First 1
    if ($natF -and -not "$($FV[$natF.Key])".Trim() -and "$($FV['name'])".Trim()) { $FV[$natF.Key] = $t.DefaultNationality }
    $tipo = "$($FV['tipo'])"
    foreach ($k in 'price', 'paid', 'gastos') { if ($FV.ContainsKey($k)) { $FV[$k] = ConvertTo-CellNumber $FV[$k] } }
    if ($t.TipoCol -and $null -eq $FV['price'] -and $tipo -and $meta.Tariff.ContainsKey($tipo)) { $FV['price'] = $meta.Tariff[$tipo] }
    $price = $FV['price']
    $ps = "$($FV['paystate'])"
    if ($null -eq $price) { if ($ps -ne 'Parcial') { $FV['paid'] = $null } }
    elseif ($ps -eq 'Pagado') { $FV['paid'] = [double]$price }
    elseif ($ps -eq 'Pendiente') { $FV['paid'] = [double]0 }
    elseif ($ps -ne 'Parcial') { $FV['paid'] = $null }
    return $FV
}
function New-RowObject($t, $c, $meta) {
    $FV = $c.F
    $rec = @{ Row = $c.Row }
    foreach ($k in $FV.Keys) { $rec[$k] = $FV[$k] }
    $price = ConvertTo-CellNumber $FV['price']; $paid = ConvertTo-CellNumber $FV['paid']; $ps = "$($FV['paystate'])"; $status = "$($FV['status'])"
    $pend = 0.0
    if ($null -ne $price) {
        if ($ps -ne 'Pagado') { $pv = 0.0; if ($null -ne $paid) { $pv = [double]$paid }; $pend = [math]::Max([double]0, [double]$price - $pv) }
    }
    $days = ''
    $closed = @('Finalizado'); if (-not $t.TipoCol) { $closed += 'Desfavorable' }
    if ($FV['fsol'] -is [DateTime] -and $status -notin $closed) { $days = [int]([DateTime]::Today - $FV['fsol'].Date).TotalDays }
    $net = 0.0
    if ($null -ne $price) {
        if ($t.NetCol -eq 'W') { $g = 0.0; $gv = ConvertTo-CellNumber $FV['gastos']; if ($null -ne $gv) { $g = $gv }; $net = [double]$price - $g }
        elseif ($t.TipoCol -and $meta.TariffNet.ContainsKey("$($FV['tipo'])")) { $net = $meta.TariffNet["$($FV['tipo'])"] }
    }
    $color = '#FFFFFF'
    if ($t.StatusColors.ContainsKey($status)) { $color = $t.StatusColors[$status] }
    return [pscustomobject]@{
        Row = $c.Row; Id = $c.Id; Name = "$($FV['name'])"; Doc = "$($FV['doc'])"
        Tipo = $(if ($FV.ContainsKey('tipo')) { "$($FV['tipo'])" } else { "$($FV['xF'])" }); Status = $status; Color = $color; Pay = $ps
        Pending = $pend
        PendingText = $(if ($pend -gt 0) { Format-Money $pend } else { '' })
        Days = $days
        Loc = "$($FV['loc'])"; Resp = "$($FV['resp'])"; Data = $rec
        PriceV = $(if ($null -ne $price) { [double]$price } else { 0 })
        PaidV = $(if ($null -ne $paid) { [double]$paid } else { 0 })
        NetV = $net
        Fecha = $(if ($FV['fsol'] -is [DateTime]) { $FV['fsol'] } else { $null })
        Sync = $c.Sync; Updated = "$($c.Updated)"; By = "$($c.By)"
    }
}
function Get-RowId($t, $meta, [int]$row) {
    if (-not $meta.IdPrefix) { return '' }
    return $meta.IdPrefix + ([string]($row - $t.FirstRow + 1)).PadLeft($meta.IdWidth, '0')
}
function Build-TramiteData($t, $store) {
    $meta = $store.Meta
    $rows = New-Object System.Collections.Generic.List[object]
    $used = New-Object System.Collections.Generic.HashSet[int]
    $payExtra = New-Object System.Collections.Generic.HashSet[string]
    foreach ($c in ($store.Clients | Sort-Object { $_.Row })) {
        $rows.Add((New-RowObject $t $c $meta)); [void]$used.Add([int]$c.Row)
        if ("$($c.F['forma'])".Trim()) { [void]$payExtra.Add("$($c.F['forma'])".Trim()) }
    }
    $firstEmpty = $null
    for ($r = $t.FirstRow; $r -le $t.LastRow; $r++) { if (-not $used.Contains($r)) { $firstEmpty = $r; break } }
    $lists = @{}
    foreach ($k in $meta.Lists.Keys) { $lists[$k] = $meta.Lists[$k] }
    $all = New-Object System.Collections.Generic.List[string]
    foreach ($p in $script:FormasPago) { $all.Add($p) }
    foreach ($p in $payExtra) { if (-not $all.Contains($p) -and $p.Length -le 30) { $all.Add($p) } }
    $lists['payments'] = $all.ToArray()
    $lists['nationalities'] = $script:Nacionalidades
    $col = New-Object System.Collections.Generic.List[string]
    foreach ($c in @($script:Cfg.Colleagues)) { if ("$c".Trim() -and -not $col.Contains("$c")) { $col.Add("$c") } }
    foreach ($r in $rows) { $x = "$($r.Resp)".Trim().ToUpper(); if ($x -and -not $col.Contains($x)) { $col.Add($x) } }
    $lists['colleagues'] = $col.ToArray()
    return @{ Lists = $lists; Tariff = $meta.Tariff; DefaultPrice = $meta.DefaultPrice; Rows = $rows; FirstEmpty = $firstEmpty
              NextId = $(if ($firstEmpty) { Get-RowId $t $meta $firstEmpty } else { '' }); Store = $store }
}

# --- importar del Excel (la primera vez, o cuando se pide) ---
function Read-TariffNet($ws, [string]$col) {
    $T = $script:T
    $map = @{}
    try { $f = "$($ws.Range("$col$($T.FirstRow)").Validation.Formula1)".TrimStart('=') } catch { return $map }
    if ($f -match '^(.+)!(\$?[A-Z]+\$?\d+(:\$?[A-Z]+\$?\d+)?)$') {
        $rng = $ws.Parent.Worksheets.Item($Matches[1].Trim("'")).Range($Matches[2])
        foreach ($c in $rng.Cells) {
            $k = "$($c.Value2)".Trim()
            $p = $c.Offset(0, 2).Value2
            if ($k -and ($p -is [double])) { $map[$k] = [double]$p }
        }
    }
    return $map
}
function Read-ExcelMeta($h, $t, $d) {
    # listas y precios de AJUSTES + formato de los nº de expediente
    $ws = Get-Sheet $h
    $lists = @{}
    foreach ($k in $d.Lists.Keys) { if ($k -like 'excel:*') { $lists[$k] = @($d.Lists[$k]) } }
    if (Test-GenTramite $t) {
        # pestaña sin listas desplegables (p. ej. creada por GESTEM): listas de siempre
        if (-not @($lists['excel:N'] | Where-Object { "$_".Trim() }).Count) { $lists['excel:N'] = $script:GenStatusList }
        if (-not @($lists['excel:M'] | Where-Object { "$_".Trim() }).Count) { $lists['excel:M'] = @('Pagado', 'Parcial', 'Pendiente') }
    }
    $net = @{}
    if ($t.TipoCol) { $net = Invoke-Excel { Read-TariffNet $ws $t.TipoCol } }
    $prefix = ''; $width = 3
    $first = "$($ws.Range("$($t.IdCol)$($t.FirstRow)").Value2)"
    if ($first -match '^(.*?)(\d+)$') { $prefix = $Matches[1]; $width = $Matches[2].Length }
    elseif (-not $first.Trim() -and (Test-GenTramite $t)) { $prefix = Get-DefaultIdPrefix $t }
    return @{ Lists = $lists; Tariff = $d.Tariff; TariffNet = $net; DefaultPrice = $d.DefaultPrice; IdPrefix = $prefix; IdWidth = $width }
}
function Import-TramiteFromExcel($h, $t) {
    if ((Test-GenTramite $t) -and -not (Invoke-Excel { Find-TramiteSheet $h.Wb $t })) {
        # el Excel no tiene esta pestaña: se empieza vacío (sin aviso) y la pestaña se crea al guardar el primer cliente
        $store = New-EmptyStore $t $h
        Write-Store $t $store
        Write-Log ("sin pestaña en el Excel, trámite empezado vacío: {0}" -f $t.Sheet)
        return (Read-Store $t)
    }
    $d = Read-TramiteData $h
    $meta = Read-ExcelMeta $h $t $d
    $clients = New-Object System.Collections.Generic.List[object]
    foreach ($r in $d.Rows) {
        $FV = @{}
        foreach ($f in $t.Fields) { $FV[$f.Key] = $r.Data[$f.Key] }
        $clients.Add(@{ Row = [int]$r.Row; Id = "$($r.Id)"; F = $FV; Sync = $true; Updated = (Get-Date).ToString('s'); By = 'Excel' })
    }
    $store = @{ Meta = $meta; Clients = $clients }
    Write-Store $t $store
    Write-Log ("importó del Excel: {0} ({1} clientes)" -f $t.Sheet, $clients.Count)
    return (Read-Store $t)    # mismos valores que se guardaron (fechas e importes ya convertidos)
}
function Update-StoreMeta($h, $t) {
    # vuelve a leer del Excel solo las listas y los precios (los clientes no se tocan)
    $d = Read-TramiteData $h
    $meta = Read-ExcelMeta $h $t $d
    $lk = Enter-StoreLock
    try {
        $s = Read-Store $t
        if (-not $s) { return }
        $s.Meta = $meta
        Write-Store $t $s
    } finally { Exit-StoreLock $lk }
}

# --- guardar: primero en GESTEM, luego en el Excel ---
function Save-ClientStore([hashtable]$vals, $editRow) {
    $t = $script:T
    $lk = Enter-StoreLock
    $rec = $null; $store = $null
    try {
        $store = Read-Store $t
        if (-not $store) { throw (L 'errNoStore') }
        $FV = @{}
        foreach ($f in $t.Fields) { $FV[$f.Key] = $vals[$f.Key] }
        $FV = Resolve-Record $t $FV $store.Meta
        if ($editRow) {
            $rec = $store.Clients | Where-Object { $_.Row -eq [int]$editRow } | Select-Object -First 1
        }
        if ($rec) {
            $rec.F = $FV
        } else {
            $used = New-Object System.Collections.Generic.HashSet[int]
            foreach ($c in $store.Clients) { [void]$used.Add([int]$c.Row) }
            $row = $null
            for ($r = $t.FirstRow; $r -le $t.LastRow; $r++) { if (-not $used.Contains($r)) { $row = $r; break } }
            if (-not $row) { throw (L 'errFull') }
            $rec = @{ Row = $row; Id = (Get-RowId $t $store.Meta $row); F = $FV; Sync = $false; New = $true; Updated = ''; By = '' }
            $store.Clients.Add($rec)
        }
        $rec.Sync = $false
        $rec.Updated = (Get-Date).ToString('s')
        $rec.By = $(if ($script:User) { $script:User } else { 'ADMIN' })
        Write-Store $t $store
    } finally { Exit-StoreLock $lk }
    # el Excel se actualiza después; si no se puede ahora, se hará más tarde
    $synced = Sync-RecordToExcel $t $rec
    return @{ Id = $rec.Id; Synced = $synced }
}
function Sync-RecordToExcel($t, $rec, $h) {
    $prevT = $script:T; $prevTar = $script:Tariff
    try {
        $script:T = $t
        $s0 = Read-Store $t
        if ($s0) { $script:Tariff = $s0.Meta.Tariff }
        $vals = @{}
        foreach ($k in $rec.F.Keys) { $vals[$k] = $rec.F[$k] }
        # campos con fórmula en el Excel: se dejan a la fórmula cuando coinciden
        $r = Save-Client $vals $rec.Row -ExpectName "$($rec.F['name'])" -IsNew:([bool]$rec.New) -Handle $h
        $lk = Enter-StoreLock
        try {
            $s = Read-Store $t
            $c = $s.Clients | Where-Object { $_.Row -eq [int]$rec.Row } | Select-Object -First 1
            if ($c) {
                if ($r.Row -ne $rec.Row) { $c.Row = $r.Row; Write-Log ("expediente {0} pasa a la fila {1} del Excel" -f $rec.Id, $r.Row) }
                if ($r.Id) { $c.Id = $r.Id }
                $c.Sync = $true; $c.New = $false
                Write-Store $t $s
            }
            $rec.Row = $r.Row; if ($r.Id) { $rec.Id = $r.Id }
        } finally { Exit-StoreLock $lk }
        return $true
    } catch {
        Write-Log ("Excel no actualizado ({0}): {1} (línea {2})" -f $rec.Id, $_.Exception.Message, $_.InvocationInfo.ScriptLineNumber)
        return $false
    } finally { $script:T = $prevT; $script:Tariff = $prevTar }
}
function Sync-PendingToExcel {
    # clientes guardados en GESTEM que todavía no están en el Excel
    $todo = @()
    foreach ($t in $script:Tramites) {
        $s = $null; try { $s = Read-Store $t } catch { }
        if ($s) { foreach ($c in $s.Clients) { if (-not $c.Sync) { $todo += ,@($t, $c) } } }
    }
    if (-not $todo.Count) { return 0 }
    $h = $null; $n = 0
    try {
        $h = Invoke-Excel { Open-Book -ForWrite }
        if ($script:Cfg.Backup) { Backup-Excel }
        foreach ($x in $todo) { if (Sync-RecordToExcel $x[0] $x[1] -h $h) { $n++ } }
        Invoke-Excel { Close-Book $h $true }; $h = $null
    } catch {
        Write-Log "sincronizar con el Excel: $($_.Exception.Message)"
        if ($h) { try { Close-Book $h $false } catch { } }
    }
    if ($n) { Write-Log ("Excel actualizado con {0} cliente(s) pendiente(s)" -f $n) }
    return $n
}
function Get-PendingSyncCount {
    $n = 0
    foreach ($t in $script:Tramites) { if ($script:Cache.ContainsKey($t.Id)) { foreach ($r in $script:Cache[$t.Id].Rows) { if (-not $r.Sync) { $n++ } } } }
    return $n
}
function Import-AllFromExcel {
    # vuelve a crear los datos de GESTEM a partir del Excel (pide confirmación)
    [void](Sync-PendingToExcel)
    [void](Load-Data -All)
    if ((Get-PendingSyncCount) -gt 0) { Show-Error (L 'errPendingSync'); return }
    $ans = [Windows.MessageBox]::Show($script:Win, (L 'importAsk'), $script:AppName, 'YesNo', 'Warning')
    if ("$ans" -ne 'Yes') { return }
    $script:Win.Cursor = [Windows.Input.Cursors]::Wait
    $h = $null; $cur = $script:T
    try {
        $h = Invoke-Excel { Open-Book }
        $lk = Enter-StoreLock
        try {
            foreach ($t in $script:Tramites) {
                $script:T = $t
                try { [void](Import-TramiteFromExcel $h $t) } catch { Write-Log "importar ($($t.Sheet)): $($_.Exception.Message)" }
            }
        } finally { Exit-StoreLock $lk }
        $script:T = $cur
        Close-Book $h $false; $h = $null
        # los trámites que solo tienen una pestaña hecha a mano se acaban de vaciar: se vuelven a leer ahora
        try { [void](Sync-HandmadeTabs -Force) } catch { Write-Log "pestañas propias: $($_.Exception.Message)" }
        [void](Load-Data -All)
        Set-Status (L 'importDone')
    } catch {
        $script:T = $cur
        if ($h) { try { Close-Book $h $false } catch { } }
        Show-Error $_.Exception.Message
    } finally { $script:Win.Cursor = $null }
}

function Refresh-FromExcel {
    # «Actualizar»: listas y precios desde AJUSTES del Excel + envía al Excel lo pendiente
    $h = $null; $cur = $script:T
    try {
        if (Test-ExcelInstalled) {
            $h = Invoke-Excel { Open-Book }
            foreach ($t in $script:Tramites) {
                $script:T = $t
                try { Update-StoreMeta $h $t } catch { Write-Log "listas ($($t.Sheet)): $($_.Exception.Message)" }
            }
            $script:T = $cur
            Close-Book $h $false; $h = $null
        }
    } catch {
        $script:T = $cur
        if ($h) { try { Close-Book $h $false } catch { } }
        Write-Log "actualizar desde el Excel: $($_.Exception.Message)"
    }
    try { [void](Sync-HandmadeTabs -Force) } catch { Write-Log "pestañas propias: $($_.Exception.Message)" }
    [void](Sync-PendingToExcel)
    return (Load-Data -All)
}
function Load-Data([switch]$All, [switch]$Quiet) {
    # lee los datos de GESTEM (rápido, sin Excel); la primera vez los importa del Excel
    if (-not $Quiet) { Set-Status (L 'loading') }
    $h = $null
    $cur = $script:T
    try {
        if (-not $script:Cfg.ExcelPath -or -not (Test-Path -LiteralPath $script:Cfg.ExcelPath)) { throw (L 'errNoFile') }
        $newCache = $script:Cache
        if ($All -or $null -eq $newCache) { $newCache = @{} }
        $list = @($cur)
        if ($All) { $list = @($cur) + @($script:Tramites | Where-Object { $_.Id -ne $cur.Id }) }
        foreach ($t in $list) {
            $script:T = $t
            try {
                $s = Read-Store $t
                if ($s -and -not $Quiet -and (Test-GenTramite $t) -and (@($script:Cfg.TabsChecked) -notcontains $t.Id)) {
                    # una sola vez por trámite: ¿el Excel tiene una pestaña con el formato de GESTEM?
                    #  - sí y GESTEM está vacío -> se importa
                    #  - no -> se quitan los clientes que versiones anteriores leyeron de una pestaña hecha a mano (datos sin sentido)
                    $done = $false
                    try {
                        if (-not $h) { $h = Invoke-Excel { Open-Book } }
                        $tab = Invoke-Excel { Find-TramiteSheet $h.Wb $t }
                        $lk = Enter-StoreLock
                        try {
                            $s2 = Read-Store $t
                            if ($s2 -and $tab -and $s2.Clients.Count -eq 0) { $s = Import-TramiteFromExcel $h $t }
                            elseif ($s2 -and -not $tab) {
                                $keep = New-Object System.Collections.Generic.List[object]
                                foreach ($c in $s2.Clients) { if ("$($c.By)" -ne 'Excel' -or [int]$c.Row -ge $script:HandRowBase) { $keep.Add($c) } }
                                if ($keep.Count -ne $s2.Clients.Count) {
                                    Write-Log ("{0}: {1} cliente(s) leídos de una pestaña que no es de GESTEM, quitados" -f $t.Sheet, ($s2.Clients.Count - $keep.Count))
                                    $s2.Clients = $keep; $s2.Meta = (New-EmptyStore $t $h).Meta
                                    Write-Store $t $s2; $s = Read-Store $t
                                }
                            }
                        } finally { Exit-StoreLock $lk }
                        $done = $true
                    } catch { Write-Log "pestaña no leída ($($t.Sheet)): $($_.Exception.Message)" }
                    finally { if ($done) { $script:Cfg.TabsChecked = @(@($script:Cfg.TabsChecked) + $t.Id | Where-Object { $_ } | Select-Object -Unique); Save-Config } }
                }
                if (-not $s -and $Quiet) { if ($script:Cache -and $script:Cache.ContainsKey($t.Id)) { $newCache[$t.Id] = $script:Cache[$t.Id] }; continue }
                if (-not $s -and (Test-TabKnownMissing $t)) {
                    if ($t.Id -ne $cur.Id) { continue }
                    throw ((L 'errNoSheet') -f $t.Sheet)
                }
                if (-not $s) {
                    if (-not $h) { $h = Invoke-Excel { Open-Book } }
                    $lk = Enter-StoreLock
                    try { $s = Read-Store $t; if (-not $s) { $s = Import-TramiteFromExcel $h $t } } finally { Exit-StoreLock $lk }
                }
                $newCache[$t.Id] = Build-TramiteData $t $s
            } catch {
                if ($t.Id -eq $cur.Id) { throw }
                Write-Log "pestaña no leída ($($t.Sheet)): $($_.Exception.Message)"
            }
        }
        $script:T = $cur
        if ($h) { Close-Book $h $false; $h = $null }
        if (-not $newCache.ContainsKey($cur.Id)) { throw (L 'errNoStore') }
        $script:Cache = $newCache
        $script:CacheStamp = Get-FileStamp
        Use-TramiteData $script:Cache[$cur.Id]
        return $true
    } catch {
        $script:T = $cur
        if ($h) { try { Close-Book $h $false } catch {} }
        Write-Log "Load-Data: $($_.Exception.Message)"
        if (-not $Quiet) {
            Set-Status ($_.Exception.Message)
            # el aviso de pestaña que falta sale una sola vez (no en cada cambio de trámite)
            if (-not ($script:MissingTabs.ContainsKey($cur.Id) -and $script:MissingWarned.ContainsKey($cur.Id))) {
                if ($script:MissingTabs.ContainsKey($cur.Id)) { $script:MissingWarned[$cur.Id] = $true }
                Show-Error $_.Exception.Message
            }
        }
        return $false
    }
}
function Sync-DataBackup([string]$dir) {
    # copia de los datos de GESTEM en la carpeta externa
    $root = Get-StoreDir
    if (-not $root -or -not (Test-Path -LiteralPath $root)) { return }
    $dst = Join-Path $dir 'datos GESTEM'
    if (-not (Test-Path -LiteralPath $dst)) { New-Item -ItemType Directory -Force -Path $dst | Out-Null }
    foreach ($f in (Get-ChildItem -LiteralPath $root -Filter '*.json' -File -ErrorAction SilentlyContinue)) {
        try { Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $dst $f.Name) -Force } catch { }
    }
}

# ---------- varios PCs a la vez ----------
function Get-EditConflict {
    # ¿alguien cambió este cliente desde que se abrió para editar?
    if (-not $script:EditRow) { return $null }
    try {
        $s = Read-Store $script:T
        if (-not $s) { return $null }
        $c = $s.Clients | Where-Object { $_.Row -eq [int]$script:EditRow.Row } | Select-Object -First 1
        if ($c -and "$($c.Updated)" -and "$($c.Updated)" -ne "$($script:EditRow.Updated)") { return $c }
    } catch { Write-Log "DEBUG conflicto: $($_.Exception.Message)" }
    return $null
}
function Format-WhoWhen($c) {
    $who = "$($c.By)"
    if (-not $who -or $who -in 'Excel', 'ADMIN') { $who = L 'someoneElse' }
    $when = ''
    $d = [DateTime]::MinValue
    if ([DateTime]::TryParse("$($c.Updated)", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$d)) {
        $when = $(if ($d.Date -eq [DateTime]::Today) { $d.ToString('HH:mm') } else { $d.ToString('dd/MM HH:mm') })
    }
    return @($who, $when)
}
function Test-ExcelFree {
    # el Excel está libre si nadie lo tiene abierto (o si lo tiene abierto este mismo PC)
    $p = "$($script:Cfg.ExcelPath)"
    if (-not $p -or -not (Test-Path -LiteralPath $p)) { return $false }
    try { $fs = [IO.File]::Open($p, 'Open', 'ReadWrite', 'None'); $fs.Dispose(); return $true }
    catch { return [bool](Get-Process -Name EXCEL -ErrorAction SilentlyContinue) }
}
function Invoke-LiveRefresh {
    # cada 20 s: si otro PC guardó, se actualiza la lista; cada 3 min se reintenta lo pendiente del Excel
    if ($script:Busy -or -not $script:User) { return }
    $script:Busy = $true
    try {
        $script:LiveTick++
        $own = $false
        try { [void](Sync-HandmadeTabs) } catch { Write-Log "pestañas propias: $($_.Exception.Message)" }    # ¿se guardó el Excel? se vuelven a leer sus pestañas
        if ($script:LiveTick % 9 -eq 0 -and (Get-PendingSyncCount) -gt 0 -and (Test-ExcelFree)) {
            if ((Sync-PendingToExcel) -gt 0) { $own = $true; $script:CacheStamp = -1 }
        }
        if ((Get-FileStamp) -ne $script:CacheStamp) {
            $selRow = $null
            if ($script:UI.ClientsGrid.SelectedItem) { $selRow = $script:UI.ClientsGrid.SelectedItem.Row }
            if (Load-Data -All -Quiet) {
                if ($selRow) {
                    foreach ($it in $script:UI.ClientsGrid.Items) { if ($it.Row -eq $selRow) { $script:UI.ClientsGrid.SelectedItem = $it; $script:UI.ClientsGrid.ScrollIntoView($it); break } }
                }
                if ($script:UI.PanelStats.Visibility -eq 'Visible') { Update-Stats }
                if (-not $own) { Set-Status (L 'liveUpdated') }
            }
        }
    } catch { Write-Log "DEBUG actualización automática: $($_.Exception.Message)" }
    finally { $script:Busy = $false }
}
$script:Busy = $false
$script:Switching = $false
$script:LiveTick = 0
$script:LiveTimer = New-Object Windows.Threading.DispatcherTimer
$script:LiveTimer.Interval = [TimeSpan]::FromSeconds(20)
$script:LiveTimer.Add_Tick({ Invoke-LiveRefresh })

# nube de la empresa: primero OneDrive de empresa (Microsoft 365), luego Google Drive, luego OneDrive personal
function Find-CompanyCloud {
    $cands = @()
    if ($env:OneDriveCommercial) { $cands += $env:OneDriveCommercial }
    foreach ($d in [IO.DriveInfo]::GetDrives()) {
        try {
            if ($d.IsReady) { foreach ($n in 'Mi unidad', 'My Drive', 'Mon Drive') { $cands += (Join-Path $d.RootDirectory.FullName $n) } }
        } catch { }
    }
    $cands += (Join-Path $env:USERPROFILE 'Google Drive\Mi unidad'); $cands += (Join-Path $env:USERPROFILE 'Google Drive')
    if ($env:OneDrive) { $cands += $env:OneDrive }
    if ($env:OneDriveConsumer) { $cands += $env:OneDriveConsumer }
    foreach ($c in $cands) { if ($c -and (Test-Path -LiteralPath $c)) { return (Join-Path $c 'GESTEM copias') } }
    return $null
}
# registro de actividad legible en la carpeta de la copia externa (nube de la empresa):
# un archivo por PC, lo más reciente arriba, SIN nombres de clientes (solo nº de expediente y trámite)
function Export-ActivityLog([string]$dir) {
    try {
        if (-not (Test-Path -LiteralPath $script:LogPath)) { return }
        $lines = @(Get-Content -LiteralPath $script:LogPath -Encoding UTF8 -Tail 4000)
        [array]::Reverse($lines)
        $out = New-Object System.Collections.Generic.List[string]
        $keys = @($script:AdmTr.Keys | Sort-Object { $_.Length } -Descending)
        foreach ($ln in $lines) {
            if ($ln -notmatch '^(\d{4})-(\d\d)-(\d\d) (\d\d:\d\d):\d\d  \[([^\]]*)\] (.+)$') { continue }
            $d = "$($Matches[3])/$($Matches[2])/$($Matches[1]) $($Matches[4])"; $who = $Matches[5]; $txt = $Matches[6]
            if ($txt -match '^(Load-Data|Save-FromForm|clic |DEBUG|pestaña no leída|copia externa: |icono:)') { continue }
            # quita el nombre del cliente
            if ($txt -match '^(añadió|editó) cliente (\S+) · .* · (.+)$') { $txt = "$($Matches[1]) cliente $($Matches[2]) · $($Matches[3])" }
            foreach ($k in $keys) {
                if ($k.Length -ge 8 -and $txt.StartsWith($k)) { $v = $script:AdmTr[$k]['en']; if ($v) { $txt = $v + $txt.Substring($k.Length) }; break }
            }
            $label = $(if ($who -eq 'ADMIN') { 'ADMIN' } elseif ($who) { $who } else { 'system' })
            $out.Add(("{0}   {1,-12} {2}" -f $d, $label, $txt))
            if ($out.Count -ge 2000) { break }
        }
        $head = @("GESTEM MANAGER - activity on PC $($env:COMPUTERNAME)", "Updated: $((Get-Date).ToString('dd/MM/yyyy HH:mm'))  ·  newest first  ·  no client names", '')
        $file = Join-Path $dir ("actividad GESTEM - {0}.txt" -f $env:COMPUTERNAME)
        [IO.File]::WriteAllLines($file, [string[]]($head + $out), (New-Object Text.UTF8Encoding($true)))
    } catch { }
}
function Backup-External([switch]$Quiet) {
    # copia completa del Excel en una carpeta externa (USB, OneDrive, Google Drive, red): un archivo por día, 60 días
    $dir = "$($script:Cfg.ExtBackupDir)".Trim()
    if (-not $dir -or -not $script:Cfg.ExcelPath) { return $false }
    try {
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        $src = $script:Cfg.ExcelPath
        $name = [IO.Path]::GetFileNameWithoutExtension($src)
        $ext = [IO.Path]::GetExtension($src)
        $dst = Join-Path $dir ("{0}_{1:yyyy-MM-dd}{2}" -f $name, (Get-Date), $ext)
        # se copia por flujo de lectura compartida: funciona aunque el Excel esté abierto
        $in = New-Object IO.FileStream($src, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        try {
            $out = New-Object IO.FileStream(($dst + '.tmp'), [IO.FileMode]::Create, [IO.FileAccess]::Write)
            try { $in.CopyTo($out) } finally { $out.Dispose() }
        } finally { $in.Dispose() }
        Move-Item -LiteralPath ($dst + '.tmp') -Destination $dst -Force
        # también los ajustes del programa (empresa, logo, compañeras, idioma…) para poder restaurarlo todo en otro PC
        $setDir = Join-Path $dir 'ajustes GESTEM'
        if (-not (Test-Path -LiteralPath $setDir)) { New-Item -ItemType Directory -Force -Path $setDir | Out-Null }
        Get-ChildItem -LiteralPath $script:DataDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq 'config.json' -or $_.Name -like 'empresa.*' } | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $setDir -Force }
        Get-ChildItem -LiteralPath $dir -Filter ($name + '_????-??-??' + $ext) | Sort-Object Name -Descending | Select-Object -Skip 60 | Remove-Item -Force -ErrorAction SilentlyContinue
        $script:Cfg.LastExtBackup = (Get-Date).ToString('yyyy-MM-dd HH:mm')
        Save-Config
        Write-Log "copia externa: $dst"
        Sync-DocsBackup $dir
        Sync-DataBackup $dir
        Export-ActivityLog $dir
        return $true
    } catch {
        Write-Log "copia externa fallida: $($_.Exception.Message)"
        if (-not $Quiet) { Show-Error ((L 'extFail') + [Environment]::NewLine + $_.Exception.Message) }
        return $false
    }
}
function Update-ExtBackupInfo {
    $ui = $script:UI
    $d = "$($ui.ExtDirBox.Text)".Trim()
    $when = "$($script:Cfg.LastExtBackup)"
    try { $when = [datetime]::ParseExact($when, 'yyyy-MM-dd HH:mm', $null).ToString('dd/MM/yyyy HH:mm') } catch {}
    $ui.ExtLastTxt.Text = $(if ($script:Cfg.LastExtBackup -and $d -and $d -eq "$($script:Cfg.ExtBackupDir)") { (L 'extLast') -f $when } elseif ($d) { L 'extNever' } else { L 'extNone' })
    $ui.ExtLastTxt.Foreground = Brush $(if ($script:Cfg.LastExtBackup -and $d) { '#1B7F45' } else { '#B42318' })
    $ui.ExtNowBtn.IsEnabled = [bool]$d
    $ui.ExtOpenBtn.IsEnabled = [bool]$d
}
function Choose-ExtFolder {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = L 'extChoose'
    $dlg.ShowNewFolderButton = $true
    if ($script:UI.ExtDirBox.Text) { $dlg.SelectedPath = $script:UI.ExtDirBox.Text }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:UI.ExtDirBox.Text = Join-Path $dlg.SelectedPath 'GESTEM copias'
        Update-ExtBackupInfo
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
                    $tab = Find-TramiteSheet $w $script:T
                    if ($row -ge $script:HandRowBase -and $script:HandTabs[$script:T.Id]) {
                        $tab = Find-Sheet $w $script:HandTabs[$script:T.Id]; $row = $row - $script:HandRowBase
                        if ($tab) { $tab.Activate(); $xl.Goto($tab.Range("A$row")) }
                        return
                    }
                    if ($row -gt 0 -and $tab) { $tab.Activate(); $xl.Goto($tab.Range("B$row")) }
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
        $psi = New-Object Diagnostics.ProcessStartInfo
        $psi.FileName = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -STA -File "' + $me + '" -Oculto'
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        [void][Diagnostics.Process]::Start($psi)
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
    if ($name -like 'fixed:*') { return ,@($name.Substring(6).Split('|') | Where-Object { $_ }) }
    foreach ($x in @($script:Lists[$name])) { if ($null -ne $x -and "$x".Trim()) { $out.Add("$x") } }
    return ,$out.ToArray()
}
function Brush([string]$hex) { return (New-Object Windows.Media.BrushConverter).ConvertFromString($hex) }
function Show-Error([string]$msg) {
    try { Write-Log ("aviso mostrado: " + ($msg -replace '\s+', ' ')) } catch { }    # lo que ve la usuaria queda en «Errores» del admin
    [void][Windows.MessageBox]::Show($script:Win, $msg, $script:AppName, 'OK', 'Warning')
}
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
        Title="GESTEM MANAGER" Width="1320" Height="840" MinWidth="1100" MinHeight="660"
        WindowStartupLocation="CenterScreen" WindowState="Maximized" Background="#F7F5F5" FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="13" Foreground="#17202E"
        UseLayoutRounding="True">
  <Window.Resources>
    <!-- GESTEM · paleta roja y blanca: rojo #C4161C, rojo oscuro #A3171F, fondo #F7F5F5 -->
    <Style x:Key="ThumbS" TargetType="Thumb">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Thumb">
            <Border x:Name="t" CornerRadius="4" Background="#C9CFD8"/>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="t" Property="Background" Value="#9AA4B2"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="TrackBtn" TargetType="RepeatButton">
      <Setter Property="Focusable" Value="False"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RepeatButton"><Border Background="Transparent"/></ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ScrollBar">
      <Setter Property="Width" Value="12"/>
      <Setter Property="MinWidth" Value="12"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ScrollBar">
            <Grid Background="Transparent">
              <Track x:Name="PART_Track" IsDirectionReversed="True" Margin="3,2">
                <Track.DecreaseRepeatButton><RepeatButton Style="{StaticResource TrackBtn}" Command="ScrollBar.PageUpCommand"/></Track.DecreaseRepeatButton>
                <Track.IncreaseRepeatButton><RepeatButton Style="{StaticResource TrackBtn}" Command="ScrollBar.PageDownCommand"/></Track.IncreaseRepeatButton>
                <Track.Thumb><Thumb Style="{StaticResource ThumbS}"/></Track.Thumb>
              </Track>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="Orientation" Value="Horizontal">
          <Setter Property="Width" Value="Auto"/>
          <Setter Property="MinWidth" Value="0"/>
          <Setter Property="Height" Value="12"/>
          <Setter Property="MinHeight" Value="12"/>
          <Setter Property="Template">
            <Setter.Value>
              <ControlTemplate TargetType="ScrollBar">
                <Grid Background="Transparent">
                  <Track x:Name="PART_Track" IsDirectionReversed="False" Margin="2,3">
                    <Track.DecreaseRepeatButton><RepeatButton Style="{StaticResource TrackBtn}" Command="ScrollBar.PageLeftCommand"/></Track.DecreaseRepeatButton>
                    <Track.IncreaseRepeatButton><RepeatButton Style="{StaticResource TrackBtn}" Command="ScrollBar.PageRightCommand"/></Track.IncreaseRepeatButton>
                    <Track.Thumb><Thumb Style="{StaticResource ThumbS}"/></Track.Thumb>
                  </Track>
                </Grid>
              </ControlTemplate>
            </Setter.Value>
          </Setter>
        </Trigger>
      </Style.Triggers>
    </Style>

    <Style TargetType="TextBox">
      <Setter Property="Height" Value="36"/>
      <Setter Property="Padding" Value="10,0"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="BorderBrush" Value="#D3D8E0"/>
      <Setter Property="Background" Value="White"/>
      <Setter Property="Foreground" Value="#17202E"/>
      <Setter Property="CaretBrush" Value="#C4161C"/>
      <Setter Property="SelectionBrush" Value="#F4A7AB"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="8">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" VerticalAlignment="{TemplateBinding VerticalContentAlignment}"
                            Focusable="False" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="{TemplateBinding VerticalScrollBarVisibility}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="#9AA4B2"/></Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="#C4161C"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="bd" Property="Opacity" Value="0.55"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <ControlTemplate x:Key="ComboToggle" TargetType="ToggleButton">
      <Border x:Name="bd" Background="White" BorderBrush="#D3D8E0" BorderThickness="1" CornerRadius="8">
        <Path x:Name="arrow" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,13,0"
              Data="M0,0 L4.5,4.5 L9,0" Stroke="#667085" StrokeThickness="1.6"/>
      </Border>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="#9AA4B2"/></Trigger>
        <Trigger Property="IsChecked" Value="True">
          <Setter TargetName="bd" Property="BorderBrush" Value="#C4161C"/>
          <Setter TargetName="arrow" Property="Stroke" Value="#C4161C"/>
        </Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
    <Style TargetType="ComboBox">
      <Setter Property="Height" Value="36"/>
      <Setter Property="Foreground" Value="#17202E"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="MaxDropDownHeight" Value="320"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid>
              <ToggleButton x:Name="tg" Template="{StaticResource ComboToggle}" Focusable="False" ClickMode="Press"
                            IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"/>
              <ContentPresenter x:Name="cp" IsHitTestVisible="False" Margin="11,0,32,0" VerticalAlignment="Center" HorizontalAlignment="Left"
                                Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"/>
              <TextBox x:Name="PART_EditableTextBox" Style="{x:Null}" Visibility="Hidden" Margin="4,1,32,1" Padding="6,0"
                       Background="Transparent" BorderThickness="0" VerticalContentAlignment="Center" Foreground="#17202E"
                       CaretBrush="#C4161C" IsReadOnly="{TemplateBinding IsReadOnly}"/>
              <Popup x:Name="PART_Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False" PopupAnimation="Fade">
                <Border Margin="0,4,10,12" MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}" MaxHeight="{TemplateBinding MaxDropDownHeight}"
                        Background="White" BorderBrush="#E1E5EB" BorderThickness="1" CornerRadius="10">
                  <Border.Effect><DropShadowEffect BlurRadius="18" ShadowDepth="4" Direction="270" Opacity="0.16" Color="#3A0A0D"/></Border.Effect>
                  <ScrollViewer Margin="4" SnapsToDevicePixels="True">
                    <ItemsPresenter KeyboardNavigation.DirectionalNavigation="Contained"/>
                  </ScrollViewer>
                </Border>
              </Popup>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsEditable" Value="True">
                <Setter TargetName="PART_EditableTextBox" Property="Visibility" Value="Visible"/>
                <Setter TargetName="cp" Property="Visibility" Value="Hidden"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.5"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ComboBoxItem">
      <Setter Property="Foreground" Value="#17202E"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBoxItem">
            <Border x:Name="bd" CornerRadius="6" Padding="10,7" MinHeight="32" Background="Transparent">
              <ContentPresenter VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsHighlighted" Value="True"><Setter TargetName="bd" Property="Background" Value="#F4F1F1"/></Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter Property="FontWeight" Value="SemiBold"/>
                <Setter Property="Foreground" Value="#C4161C"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="CalDay" TargetType="CalendarDayButton">
      <Setter Property="MinWidth" Value="40"/>
      <Setter Property="MinHeight" Value="38"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Foreground" Value="#2B2526"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CalendarDayButton">
            <Grid Background="Transparent">
              <Ellipse x:Name="hov" Width="34" Height="34" Fill="#FBEDEE" Visibility="Collapsed"/>
              <Ellipse x:Name="today" Width="34" Height="34" Stroke="#C4161C" StrokeThickness="1.4" Visibility="Collapsed"/>
              <Ellipse x:Name="sel" Width="34" Height="34" Visibility="Collapsed">
                <Ellipse.Fill>
                  <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                    <GradientStop Color="#D42A31" Offset="0"/>
                    <GradientStop Color="#9E1820" Offset="1"/>
                  </LinearGradientBrush>
                </Ellipse.Fill>
                <Ellipse.Effect><DropShadowEffect BlurRadius="8" ShadowDepth="2" Direction="270" Opacity="0.3" Color="#9E1820"/></Ellipse.Effect>
              </Ellipse>
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="hov" Property="Visibility" Value="Visible"/>
                <Setter Property="Foreground" Value="#9E1820"/>
              </Trigger>
              <Trigger Property="IsInactive" Value="True"><Setter Property="Foreground" Value="#C9C1C3"/></Trigger>
              <Trigger Property="IsToday" Value="True">
                <Setter TargetName="today" Property="Visibility" Value="Visible"/>
                <Setter Property="FontWeight" Value="SemiBold"/>
                <Setter Property="Foreground" Value="#9E1820"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="sel" Property="Visibility" Value="Visible"/>
                <Setter TargetName="today" Property="Visibility" Value="Collapsed"/>
                <Setter Property="Foreground" Value="White"/>
                <Setter Property="FontWeight" Value="SemiBold"/>
              </Trigger>
              <Trigger Property="IsBlackedOut" Value="True"><Setter Property="Opacity" Value="0.3"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="CalBtn" TargetType="CalendarButton">
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Foreground" Value="#2B2526"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CalendarButton">
            <Border x:Name="b" Width="62" Height="44" Margin="4" CornerRadius="10" Background="Transparent">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="b" Property="Background" Value="#FBEDEE"/>
                <Setter Property="Foreground" Value="#9E1820"/>
              </Trigger>
              <Trigger Property="HasSelectedDays" Value="True">
                <Setter TargetName="b" Property="Background" Value="#C4161C"/>
                <Setter Property="Foreground" Value="White"/>
                <Setter Property="FontWeight" Value="SemiBold"/>
              </Trigger>
              <Trigger Property="IsInactive" Value="True"><Setter Property="Foreground" Value="#C9C1C3"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="CalNav" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Focusable" Value="False"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" Width="30" Height="30" CornerRadius="15" Background="#1FFFFFFF">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Background" Value="#40FFFFFF"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.35"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="CalHead" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontSize" Value="15"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Focusable" Value="False"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" CornerRadius="8" Padding="10,4" Background="Transparent" HorizontalAlignment="Center">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Background" Value="#26FFFFFF"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="CalItem" TargetType="CalendarItem">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CalendarItem">
            <ControlTemplate.Resources>
              <DataTemplate x:Key="{x:Static CalendarItem.DayTitleTemplateResourceKey}">
                <TextBlock FontSize="11" FontWeight="SemiBold" Foreground="#A69C9E" HorizontalAlignment="Center" Margin="0,10,0,4">
                  <TextBlock.Style>
                    <Style TargetType="TextBlock">
                      <Setter Property="Text" Value="{Binding}"/>
                      <Style.Triggers>
                        <DataTrigger Binding="{Binding}" Value="X"><Setter Property="Text" Value="Mi"/></DataTrigger>
                      </Style.Triggers>
                    </Style>
                  </TextBlock.Style>
                </TextBlock>
              </DataTemplate>
            </ControlTemplate.Resources>
            <Grid x:Name="PART_Root">
              <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
              <Border CornerRadius="11,11,0,0" Padding="10,10">
                <Border.Background>
                  <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                    <GradientStop Color="#9E1820" Offset="0"/>
                    <GradientStop Color="#6E0E14" Offset="1"/>
                  </LinearGradientBrush>
                </Border.Background>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <Button x:Name="PART_PreviousButton" Style="{StaticResource CalNav}">
                    <TextBlock Text="&#xE76B;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="11" Foreground="White"/>
                  </Button>
                  <Button x:Name="PART_HeaderButton" Grid.Column="1" Style="{StaticResource CalHead}"/>
                  <Button x:Name="PART_NextButton" Grid.Column="2" Style="{StaticResource CalNav}">
                    <TextBlock Text="&#xE76C;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="11" Foreground="White"/>
                  </Button>
                </Grid>
              </Border>
              <Grid x:Name="PART_MonthView" Grid.Row="1" Margin="10,0,10,6" Visibility="Visible">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
              </Grid>
              <Grid x:Name="PART_YearView" Grid.Row="1" Margin="10,10,10,10" Visibility="Hidden" HorizontalAlignment="Center">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
              </Grid>
              <Rectangle x:Name="PART_DisabledVisual" Grid.RowSpan="2" Fill="#80FFFFFF" Visibility="Collapsed" RadiusX="12" RadiusY="12"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="PART_DisabledVisual" Property="Visibility" Value="Visible"/></Trigger>
              <DataTrigger Binding="{Binding DisplayMode, RelativeSource={RelativeSource FindAncestor, AncestorType={x:Type Calendar}}}" Value="Year">
                <Setter TargetName="PART_MonthView" Property="Visibility" Value="Hidden"/>
                <Setter TargetName="PART_YearView" Property="Visibility" Value="Visible"/>
              </DataTrigger>
              <DataTrigger Binding="{Binding DisplayMode, RelativeSource={RelativeSource FindAncestor, AncestorType={x:Type Calendar}}}" Value="Decade">
                <Setter TargetName="PART_MonthView" Property="Visibility" Value="Hidden"/>
                <Setter TargetName="PART_YearView" Property="Visibility" Value="Visible"/>
              </DataTrigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="CalToday" TargetType="Button">
      <Setter Property="Foreground" Value="#9E1820"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Focusable" Value="False"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" CornerRadius="16" Padding="16,6" Background="#FBEDEE" HorizontalAlignment="Center">
              <StackPanel Orientation="Horizontal">
                <TextBlock Text="&#xE8D1;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="12" Foreground="{TemplateBinding Foreground}" VerticalAlignment="Center" Margin="0,0,7,0"/>
                <ContentPresenter VerticalAlignment="Center"/>
              </StackPanel>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="b" Property="Background" Value="#C4161C"/>
                <Setter Property="Foreground" Value="White"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Cal" TargetType="Calendar">
      <Setter Property="CalendarDayButtonStyle" Value="{StaticResource CalDay}"/>
      <Setter Property="CalendarButtonStyle" Value="{StaticResource CalBtn}"/>
      <Setter Property="CalendarItemStyle" Value="{StaticResource CalItem}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Calendar">
            <Border Margin="0,6,16,18" Background="White" BorderBrush="#E4DFE0" BorderThickness="1" CornerRadius="12">
              <Border.Effect><DropShadowEffect BlurRadius="26" ShadowDepth="6" Direction="270" Opacity="0.22" Color="#3A0A0D"/></Border.Effect>
              <StackPanel>
                <CalendarItem x:Name="PART_CalendarItem" Style="{TemplateBinding CalendarItemStyle}" Background="White" BorderBrush="White" BorderThickness="0"/>
                <Border BorderBrush="#F1EDED" BorderThickness="0,1,0,0" Padding="0,10,0,12">
                  <Button x:Name="TodayBtn" Style="{StaticResource CalToday}" Content="Hoy"/>
                </Border>
              </StackPanel>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="DateText" TargetType="DatePickerTextBox">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="#17202E"/>
      <Setter Property="CaretBrush" Value="#C4161C"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DatePickerTextBox">
            <ScrollViewer x:Name="PART_ContentHost" Margin="10,0,4,0" VerticalAlignment="Center" Focusable="False"
                          HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden"/>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="DateBtn" TargetType="Button">
      <Setter Property="Foreground" Value="#8A7F81"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" Width="36" Background="Transparent" CornerRadius="0,8,8,0">
              <TextBlock Text="&#xE787;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="15" Foreground="{TemplateBinding Foreground}"
                         HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Foreground" Value="#C4161C"/>
                <Setter TargetName="b" Property="Background" Value="#F7F3F3"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="DatePicker">
      <Setter Property="Height" Value="36"/>
      <Setter Property="IsTodayHighlighted" Value="True"/>
      <Setter Property="CalendarStyle" Value="{StaticResource Cal}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DatePicker">
            <Border x:Name="bd" Background="White" BorderBrush="#D9D4D5" BorderThickness="1" CornerRadius="8">
              <Grid x:Name="PART_Root">
                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <DatePickerTextBox x:Name="PART_TextBox" Style="{StaticResource DateText}"/>
                <TextBlock x:Name="ph" Text="dd/mm/aaaa" Foreground="#B0A8AA" Margin="11,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False" Visibility="Collapsed"/>
                <Button x:Name="PART_Button" Grid.Column="1" Style="{StaticResource DateBtn}" Focusable="False" Cursor="Hand"/>
                <Popup x:Name="PART_Popup" PlacementTarget="{Binding ElementName=bd}" Placement="Bottom" StaysOpen="False" AllowsTransparency="True" PopupAnimation="Fade"/>
              </Grid>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="#A69C9E"/></Trigger>
              <Trigger Property="IsKeyboardFocusWithin" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="#C4161C"/></Trigger>
              <MultiTrigger>
                <MultiTrigger.Conditions>
                  <Condition Property="SelectedDate" Value="{x:Null}"/>
                  <Condition Property="IsKeyboardFocusWithin" Value="False"/>
                </MultiTrigger.Conditions>
                <Setter TargetName="ph" Property="Visibility" Value="Visible"/>
              </MultiTrigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="bd" Property="Opacity" Value="0.55"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Foreground" Value="#344054"/>
      <Setter Property="Margin" Value="0,6"/>
      <Setter Property="Cursor" Value="Hand"/>
    </Style>
    <Style x:Key="FieldLabel" TargetType="TextBlock">
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Foreground" Value="#4B5565"/>
      <Setter Property="Margin" Value="1,0,0,6"/>
    </Style>
    <Style x:Key="SectionTitle" TargetType="TextBlock">
      <Setter Property="FontFamily" Value="Segoe UI Variable Display, Segoe UI"/>
      <Setter Property="FontSize" Value="15"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="#1E1A1B"/>
      <Setter Property="Margin" Value="0,0,0,0"/>
    </Style>
    <Style x:Key="Heading" TargetType="TextBlock">
      <Setter Property="FontFamily" Value="Segoe UI Variable Display, Segoe UI"/>
      <Setter Property="FontSize" Value="23"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="#1E1A1B"/>
    </Style>
    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="White"/>
      <Setter Property="CornerRadius" Value="14"/>
      <Setter Property="BorderBrush" Value="#E4E7EC"/>
      <Setter Property="BorderThickness" Value="1"/>
    </Style>

    <Style x:Key="Nav" TargetType="RadioButton">
      <Setter Property="Foreground" Value="#F3D6D8"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Margin" Value="0,2"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RadioButton">
            <Grid>
              <Border x:Name="b" CornerRadius="9" Padding="14,11" Background="Transparent">
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="30"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <TextBlock x:Name="ic" Text="{TemplateBinding Tag}" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="16"
                             VerticalAlignment="Center" Foreground="#E3A9AE"/>
                  <ContentPresenter Grid.Column="1" VerticalAlignment="Center"/>
                </Grid>
              </Border>
              <Border x:Name="bar" Width="3" CornerRadius="2" Margin="0,10" HorizontalAlignment="Left" Background="White" Visibility="Hidden"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="b" Property="Background" Value="#14FFFFFF"/>
                <Setter Property="Foreground" Value="White"/>
              </Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="b" Property="Background" Value="#26FFFFFF"/>
                <Setter TargetName="bar" Property="Visibility" Value="Visible"/>
                <Setter TargetName="ic" Property="Foreground" Value="White"/>
                <Setter Property="Foreground" Value="White"/>
                <Setter Property="FontWeight" Value="SemiBold"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="UserChip" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" Background="#33000000" CornerRadius="12" Padding="10,10" BorderBrush="#24FFFFFF" BorderThickness="1">
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <Border Width="32" Height="32" CornerRadius="16" Background="White">
                  <TextBlock Text="&#xE77B;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="15" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
                <ContentPresenter Grid.Column="1" Margin="10,0,6,0" VerticalAlignment="Center"/>
                <TextBlock Grid.Column="2" Text="&#xE8AB;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="12" Foreground="#E3A9AE" VerticalAlignment="Center"/>
              </Grid>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Background" Value="#4D000000"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Pill" TargetType="RadioButton">
      <Setter Property="Margin" Value="0,0,8,8"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Background" Value="White"/>
      <Setter Property="Foreground" Value="#2B2526"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RadioButton">
            <Border x:Name="b" CornerRadius="17" Padding="12,7,15,7" Background="White" BorderBrush="#DCD6D7" BorderThickness="1">
              <StackPanel Orientation="Horizontal">
                <Ellipse x:Name="dot" Width="9" Height="9" Margin="0,0,8,0" VerticalAlignment="Center" Fill="{TemplateBinding Background}" Stroke="#33000000" StrokeThickness="0.6"/>
                <TextBlock x:Name="chk" Text="&#xE73E;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="11" Margin="0,1,7,0" VerticalAlignment="Center" Visibility="Collapsed"/>
                <ContentPresenter VerticalAlignment="Center"/>
              </StackPanel>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="Uid" Value="plain"><Setter TargetName="dot" Property="Visibility" Value="Collapsed"/></Trigger>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="#A69C9E"/></Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="#C4161C"/></Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="b" Property="Background" Value="{Binding Background, RelativeSource={RelativeSource TemplatedParent}}"/>
                <Setter TargetName="b" Property="BorderBrush" Value="#2B2526"/>
                <Setter TargetName="b" Property="BorderThickness" Value="1.5"/>
                <Setter TargetName="dot" Property="Visibility" Value="Collapsed"/>
                <Setter TargetName="chk" Property="Visibility" Value="Visible"/>
                <Setter TargetName="b" Property="Effect">
                  <Setter.Value><DropShadowEffect BlurRadius="10" ShadowDepth="2" Direction="270" Opacity="0.18" Color="#000000"/></Setter.Value>
                </Setter>
                <Setter Property="FontWeight" Value="SemiBold"/>
              </Trigger>
              <MultiTrigger>
                <MultiTrigger.Conditions>
                  <Condition Property="IsChecked" Value="True"/>
                  <Condition Property="Uid" Value="plain"/>
                </MultiTrigger.Conditions>
                <Setter TargetName="b" Property="Background" Value="#221E1F"/>
                <Setter TargetName="b" Property="BorderBrush" Value="#221E1F"/>
                <Setter Property="Foreground" Value="White"/>
                <Setter TargetName="chk" Property="Foreground" Value="#FF5A62"/>
              </MultiTrigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Seg" TargetType="RadioButton">
      <Setter Property="Foreground" Value="#6E6466"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RadioButton">
            <Border x:Name="b" CornerRadius="8" Padding="16,6" Background="Transparent">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter Property="Foreground" Value="#1E1A1B"/></Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="b" Property="Background" Value="White"/>
                <Setter TargetName="b" Property="Effect">
                  <Setter.Value><DropShadowEffect BlurRadius="8" ShadowDepth="1" Direction="270" Opacity="0.15" Color="#000000"/></Setter.Value>
                </Setter>
                <Setter Property="Foreground" Value="#9E1820"/>
                <Setter Property="FontWeight" Value="SemiBold"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Primary" TargetType="Button">
      <Setter Property="Background">
        <Setter.Value>
          <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
            <GradientStop Color="#CC1B21" Offset="0"/>
            <GradientStop Color="#9A1218" Offset="1"/>
          </LinearGradientBrush>
        </Setter.Value>
      </Setter>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontSize" Value="15"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Height" Value="50"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" Background="{TemplateBinding Background}" CornerRadius="10" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Background" Value="#A3171F"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="b" Property="Background" Value="#7A0F15"/></Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="b" Property="Background" Value="#A3171F"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="b" Property="Opacity" Value="0.55"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Btn" TargetType="Button">
      <Setter Property="Background" Value="White"/>
      <Setter Property="Foreground" Value="#17202E"/>
      <Setter Property="Height" Value="38"/>
      <Setter Property="Padding" Value="16,0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" Background="{TemplateBinding Background}" CornerRadius="9" BorderBrush="#D3D8E0" BorderThickness="1" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="b" Property="Background" Value="#F3F5F9"/>
                <Setter TargetName="b" Property="BorderBrush" Value="#9AA4B2"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="#C4161C"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Grid>
  <DockPanel>
    <!-- MENÚ LATERAL -->
    <Border DockPanel.Dock="Left" Width="240">
      <Border.Background>
        <LinearGradientBrush StartPoint="0,0" EndPoint="0.3,1">
          <GradientStop Color="#9E1820" Offset="0"/>
          <GradientStop Color="#7E1118" Offset="0.45"/>
          <GradientStop Color="#560A0F" Offset="1"/>
        </LinearGradientBrush>
      </Border.Background>
      <DockPanel Margin="16,26,16,18">
        <StackPanel DockPanel.Dock="Top" Margin="0,0,0,26">
          <Grid Width="92" Height="92" HorizontalAlignment="Center">
            <Grid.Effect><DropShadowEffect BlurRadius="26" ShadowDepth="5" Direction="270" Opacity="0.5" Color="#2A0306"/></Grid.Effect>
            <Ellipse Fill="White"/>
            <Ellipse x:Name="SideLogo" Margin="4"/>
          </Grid>
          <TextBlock x:Name="SideCompany" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="15" FontWeight="Bold" Foreground="White"
                     HorizontalAlignment="Center" TextAlignment="Center" TextWrapping="Wrap" Margin="0,16,0,0"/>
          <Border Height="1" Width="40" Background="#80FFFFFF" Margin="0,12,0,10"/>
          <TextBlock HorizontalAlignment="Center" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="12" Foreground="#E9B9BD"><Run Text="GESTEM" FontWeight="SemiBold" Foreground="White"/><Run Text=" Manager"/></TextBlock>
        </StackPanel>
        <Button x:Name="UserBtn" DockPanel.Dock="Bottom" Style="{StaticResource UserChip}"/>
        <TextBlock x:Name="FileText" DockPanel.Dock="Bottom" Foreground="#D99FA4" FontSize="11" Margin="8,0,8,10" TextTrimming="CharacterEllipsis"/>
        <StackPanel>
          <RadioButton x:Name="NavNew" Style="{StaticResource Nav}" GroupName="nav" IsChecked="True" Tag="&#xE8FA;"/>
          <RadioButton x:Name="NavClients" Style="{StaticResource Nav}" GroupName="nav" Tag="&#xE716;"/>
          <RadioButton x:Name="NavStats" Style="{StaticResource Nav}" GroupName="nav" Tag="&#xE9D2;"/>
          <RadioButton x:Name="NavSettings" Style="{StaticResource Nav}" GroupName="nav" Tag="&#xE713;"/>
          <RadioButton x:Name="NavAbout" Style="{StaticResource Nav}" GroupName="nav" Tag="&#xE946;"/>
        </StackPanel>
      </DockPanel>
    </Border>

    <!-- CONTENIDO: ancho máximo para que no se separe todo en pantalla completa -->
    <Grid MaxWidth="1440">
    <DockPanel>
    <!-- BARRA SUPERIOR: trámite -->
    <Border x:Name="TopBar" DockPanel.Dock="Top" Padding="28,20,28,0">
      <DockPanel>
        <TextBlock x:Name="StatusText" DockPanel.Dock="Right" Foreground="#667085" FontSize="12" VerticalAlignment="Center" Margin="16,0,0,0"/>
        <StackPanel Orientation="Horizontal">
          <TextBlock x:Name="TramiteLabel" Foreground="#667085" VerticalAlignment="Center" Margin="0,0,12,0"/>
          <ComboBox x:Name="TramiteBox" Width="300" Height="40" FontSize="14" FontWeight="SemiBold"/>
          <ComboBox x:Name="StatsScope" Width="300" Height="40" FontSize="14" FontWeight="SemiBold" Visibility="Collapsed"/>
        </StackPanel>
      </DockPanel>
    </Border>

    <Grid Margin="28,16,28,22">

      <!-- NUEVO / EDITAR CLIENTE -->
      <Grid x:Name="PanelForm">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="320"/>
        </Grid.ColumnDefinitions>
        <Border Style="{StaticResource Card}">
          <DockPanel>
            <TextBlock x:Name="FormTitle" DockPanel.Dock="Top" Style="{StaticResource Heading}" Margin="28,22,28,0"/>
            <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="0,0,6,0" Margin="0,0,0,8">
              <Grid x:Name="FormGrid" Margin="28,4,12,22">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
              </Grid>
            </ScrollViewer>
          </DockPanel>
        </Border>
        <StackPanel Grid.Column="1" Margin="18,0,0,0">
          <Border CornerRadius="14" Background="White" BorderBrush="#E4E7EC" BorderThickness="1">
            <Border.Effect><DropShadowEffect BlurRadius="22" ShadowDepth="3" Direction="270" Opacity="0.07" Color="#3A0A0D"/></Border.Effect>
            <StackPanel>
              <Border CornerRadius="13,13,0,0" Padding="18,12,18,10" BorderBrush="#D8202A" BorderThickness="0,0,0,3">
                <Border.Background>
                  <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
                    <GradientStop Color="#221E1F" Offset="0"/>
                    <GradientStop Color="#363031" Offset="1"/>
                  </LinearGradientBrush>
                </Border.Background>
                <DockPanel>
                  <TextBlock x:Name="CardTag" DockPanel.Dock="Right" Foreground="#A69C9E" FontSize="12" VerticalAlignment="Center"/>
                  <TextBlock x:Name="CardId" Foreground="White" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="15" FontWeight="SemiBold"/>
                </DockPanel>
              </Border>
              <StackPanel Margin="18,14,18,18">
                <TextBlock x:Name="CardName" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="20" FontWeight="SemiBold" TextWrapping="Wrap"/>
                <TextBlock x:Name="CardSub" Foreground="#667085" Margin="0,5,0,0" TextWrapping="Wrap"/>
                <TextBlock x:Name="CardTipo" Foreground="#667085" Margin="0,2,0,0"/>
                <Border x:Name="CardStatusB" CornerRadius="12" Padding="12,4" Margin="0,12,0,0" HorizontalAlignment="Left">
                  <TextBlock x:Name="CardStatus" FontWeight="SemiBold" Foreground="#2A2A2A"/>
                </Border>
              </StackPanel>
              <Border BorderBrush="#EEF0F3" BorderThickness="0,1,0,0" Padding="18,12,18,16">
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <Grid.RowDefinitions><RowDefinition/><RowDefinition/><RowDefinition/></Grid.RowDefinitions>
                  <TextBlock x:Name="SumPriceL" Foreground="#667085" Margin="0,4"/>
                  <TextBlock x:Name="SumPrice" Grid.Column="1" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="15" Margin="0,4"/>
                  <TextBlock x:Name="SumPaidL" Grid.Row="1" Foreground="#667085" Margin="0,4"/>
                  <TextBlock x:Name="SumPaid" Grid.Row="1" Grid.Column="1" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="15" Margin="0,4"/>
                  <TextBlock x:Name="SumPendL" Grid.Row="2" FontWeight="SemiBold" Margin="0,8,0,0" VerticalAlignment="Center"/>
                  <TextBlock x:Name="SumPend" Grid.Row="2" Grid.Column="1" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="20" FontWeight="SemiBold" Margin="0,8,0,0"/>
                </Grid>
              </Border>
            </StackPanel>
          </Border>
          <TextBlock x:Name="DupWarn" Foreground="#B42318" TextWrapping="Wrap" Margin="2,12,2,0" Visibility="Collapsed"/>
          <Button x:Name="SaveBtn" Style="{StaticResource Primary}" Margin="0,16,0,0"/>
          <Button x:Name="ClearBtn" Style="{StaticResource Btn}" Margin="0,8,0,0"/>
        </StackPanel>
      </Grid>

      <!-- CLIENTES -->
      <Border x:Name="PanelClients" Visibility="Collapsed" Style="{StaticResource Card}" Padding="20,18">
        <DockPanel>
          <DockPanel DockPanel.Dock="Top" Margin="0,0,0,14">
            <StackPanel Orientation="Horizontal" DockPanel.Dock="Right">
              <Button x:Name="DocsBtn" Style="{StaticResource Btn}" Margin="8,0,0,0"/>
              <Button x:Name="ReceiptBtn" Style="{StaticResource Btn}" Margin="8,0,0,0"/>
              <Button x:Name="ExportBtn" Style="{StaticResource Btn}" Margin="8,0,0,0">
                <Button.ContextMenu>
                  <ContextMenu>
                    <MenuItem x:Name="ExpPdf"/>
                    <MenuItem x:Name="ExpCsv"/>
                  </ContextMenu>
                </Button.ContextMenu>
              </Button>
              <Button x:Name="RefreshBtn" Style="{StaticResource Btn}" Margin="8,0,0,0"/>
              <Button x:Name="OpenExcelBtn" Style="{StaticResource Btn}" Margin="8,0,0,0"/>
              <Button x:Name="EditBtn" Style="{StaticResource Primary}" Margin="8,0,0,0" Height="38" FontSize="13" Padding="18,0"/>
            </StackPanel>
            <WrapPanel>
              <Grid Width="250">
                <TextBox x:Name="SearchBox" Padding="32,0,10,0"/>
                <TextBlock Text="&#xE721;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="13" Foreground="#98A2B3" IsHitTestVisible="False" Margin="12,0,0,0" VerticalAlignment="Center"/>
                <TextBlock x:Name="SearchHint" IsHitTestVisible="False" Foreground="#98A2B3" Margin="34,0,0,0" VerticalAlignment="Center"/>
              </Grid>
              <ComboBox x:Name="FilterBox" Width="170" Margin="10,0,0,0"/>
              <CheckBox x:Name="OnlyDebt" Margin="16,0,0,0"/>
              <CheckBox x:Name="OnlyMine" Margin="16,0,0,0"/>
            </WrapPanel>
          </DockPanel>
          <DataGrid x:Name="ClientsGrid" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single"
                    HeadersVisibility="Column" GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="#EEF0F3"
                    RowHeight="40" Background="White" BorderThickness="0" CanUserAddRows="False" CanUserResizeRows="False"
                    AlternatingRowBackground="#FAFBFC" Foreground="#17202E">
            <DataGrid.ContextMenu>
              <ContextMenu>
                <MenuItem x:Name="CmEdit"/>
                <MenuItem x:Name="CmReceipt"/>
                <MenuItem x:Name="CmDocs"/>
              </ContextMenu>
            </DataGrid.ContextMenu>
            <DataGrid.ColumnHeaderStyle>
              <Style TargetType="DataGridColumnHeader">
                <Setter Property="Background" Value="White"/>
                <Setter Property="Foreground" Value="#667085"/>
                <Setter Property="FontSize" Value="12"/>
                <Setter Property="FontWeight" Value="SemiBold"/>
                <Setter Property="Padding" Value="8,10"/>
                <Setter Property="BorderBrush" Value="#E4E7EC"/>
                <Setter Property="BorderThickness" Value="0,0,0,1"/>
              </Style>
            </DataGrid.ColumnHeaderStyle>
            <DataGrid.RowStyle>
              <Style TargetType="DataGridRow">
                <Setter Property="Cursor" Value="Hand"/>
                <Style.Triggers>
                  <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#F7F5F5"/></Trigger>
                </Style.Triggers>
              </Style>
            </DataGrid.RowStyle>
            <DataGrid.CellStyle>
              <Style TargetType="DataGridCell">
                <Setter Property="BorderThickness" Value="0"/>
                <Setter Property="Padding" Value="6,0"/>
                <Setter Property="VerticalContentAlignment" Value="Center"/>
                <Setter Property="Template">
                  <Setter.Value>
                    <ControlTemplate TargetType="DataGridCell">
                      <Border Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                        <ContentPresenter VerticalAlignment="Center"/>
                      </Border>
                    </ControlTemplate>
                  </Setter.Value>
                </Setter>
                <Style.Triggers>
                  <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#FBE3E4"/>
                    <Setter Property="Foreground" Value="#7A0F15"/>
                  </Trigger>
                </Style.Triggers>
              </Style>
            </DataGrid.CellStyle>
            <DataGrid.Columns>
              <DataGridTextColumn Binding="{Binding Id}" Width="112"/>
              <DataGridTextColumn Binding="{Binding Name}" Width="2*" MinWidth="200"/>
              <DataGridTextColumn Binding="{Binding Doc}" Width="*" MinWidth="100"/>
              <DataGridTextColumn Binding="{Binding Tipo}" Width="*" MinWidth="100"/>
              <DataGridTemplateColumn Width="150">
                <DataGridTemplateColumn.CellTemplate>
                  <DataTemplate>
                    <Border Background="{Binding Color}" CornerRadius="11" Padding="10,2" Margin="0,6" HorizontalAlignment="Left">
                      <TextBlock Text="{Binding Status}" Foreground="#2A2A2A" FontSize="12"/>
                    </Border>
                  </DataTemplate>
                </DataGridTemplateColumn.CellTemplate>
              </DataGridTemplateColumn>
              <DataGridTextColumn Binding="{Binding Pay}" Width="85"/>
              <DataGridTextColumn Binding="{Binding PendingText}" Width="95">
                <DataGridTextColumn.ElementStyle>
                  <Style TargetType="TextBlock">
                    <Setter Property="Foreground" Value="#B42318"/>
                    <Setter Property="FontWeight" Value="SemiBold"/>
                    <Setter Property="VerticalAlignment" Value="Center"/>
                  </Style>
                </DataGridTextColumn.ElementStyle>
              </DataGridTextColumn>
              <DataGridTextColumn Binding="{Binding Days}" Width="50"/>
              <DataGridTextColumn Binding="{Binding Resp}" Width="100"/>
              <DataGridTemplateColumn x:Name="DocsCol" Width="64">
                <DataGridTemplateColumn.CellTemplate>
                  <DataTemplate>
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Visibility="{Binding DocsVis}">
                      <TextBlock Text="&#xE723;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="12" Foreground="#9E1820" VerticalAlignment="Center"/>
                      <TextBlock Text="{Binding DocsN}" Margin="5,0,0,0" FontWeight="SemiBold" Foreground="#4B4345" VerticalAlignment="Center"/>
                    </StackPanel>
                  </DataTemplate>
                </DataGridTemplateColumn.CellTemplate>
              </DataGridTemplateColumn>
            </DataGrid.Columns>
          </DataGrid>
        </DockPanel>
      </Border>

      <!-- RESUMEN ECONÓMICO -->
      <ScrollViewer x:Name="PanelStats" Visibility="Collapsed" VerticalScrollBarVisibility="Auto" Padding="0,0,8,0">
      <StackPanel>
        <DockPanel Margin="0,0,0,18">
          <StackPanel Orientation="Horizontal" DockPanel.Dock="Right" VerticalAlignment="Center">
            <ComboBox x:Name="StatsWho" Width="180"/>
            <ComboBox x:Name="StatsPeriod" Width="170" Margin="10,0,0,0"/>
            <Button x:Name="StatsRefresh" Style="{StaticResource Btn}" Margin="10,0,0,0" Padding="12,0" ToolTip="Actualizar">
              <TextBlock Text="&#xE72C;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="14"/>
            </Button>
          </StackPanel>
          <StackPanel>
            <StackPanel Orientation="Horizontal">
              <TextBlock x:Name="StatsTitle" Style="{StaticResource Heading}" VerticalAlignment="Center"/>
              <Button x:Name="StatsAllBtn" Margin="14,2,0,0" VerticalAlignment="Center" Cursor="Hand" Visibility="Collapsed">
                <Button.Template>
                  <ControlTemplate TargetType="Button">
                    <Border x:Name="b" CornerRadius="15" Padding="12,5" Background="#FBEDEE" BorderBrush="#F2C4C7" BorderThickness="1">
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#xE72B;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="11" Foreground="#9E1820" VerticalAlignment="Center" Margin="0,1,7,0"/>
                        <ContentPresenter VerticalAlignment="Center"/>
                      </StackPanel>
                    </Border>
                    <ControlTemplate.Triggers>
                      <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Background" Value="#F7DADC"/></Trigger>
                    </ControlTemplate.Triggers>
                  </ControlTemplate>
                </Button.Template>
                <TextBlock x:Name="StatsAllTxt" Foreground="#9E1820" FontWeight="SemiBold" FontSize="13"/>
              </Button>
            </StackPanel>
            <TextBlock x:Name="StatsSub" Foreground="#8A7F81" Margin="0,4,0,0"/>
          </StackPanel>
        </DockPanel>
        <UniformGrid Columns="4" Margin="-8,0,-8,18">
          <Border Style="{StaticResource Card}" Margin="8,0" Padding="22,18">
            <StackPanel>
              <TextBlock x:Name="KpiBillL" Foreground="#6E6466" FontSize="14" FontWeight="SemiBold"/>
              <TextBlock x:Name="KpiBill" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="34" FontWeight="Bold" Foreground="#1E1A1B" Margin="0,8,0,0"/>
              <TextBlock x:Name="KpiBillF" Foreground="#6E6466" FontSize="13" Margin="0,6,0,0"/>
            </StackPanel>
          </Border>
          <Border Style="{StaticResource Card}" Margin="8,0" Padding="22,18">
            <StackPanel>
              <TextBlock x:Name="KpiPaidL" Foreground="#6E6466" FontSize="14" FontWeight="SemiBold"/>
              <TextBlock x:Name="KpiPaid" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="34" FontWeight="Bold" Foreground="#1E1A1B" Margin="0,8,0,0"/>
              <Grid x:Name="KpiPaidTrack" Height="6" Margin="0,10,0,0">
                <Grid.ColumnDefinitions><ColumnDefinition Width="0*"/><ColumnDefinition Width="1*"/></Grid.ColumnDefinitions>
                <Border Grid.ColumnSpan="2" Background="#F1EDED" CornerRadius="3"/>
                <Border Background="#3E8E41" CornerRadius="3"/>
              </Grid>
              <TextBlock x:Name="KpiPaidF" Foreground="#6E6466" FontSize="13" Margin="0,6,0,0"/>
            </StackPanel>
          </Border>
          <Border Style="{StaticResource Card}" Margin="8,0" Padding="22,18">
            <StackPanel>
              <TextBlock x:Name="KpiPendL" Foreground="#6E6466" FontSize="14" FontWeight="SemiBold"/>
              <TextBlock x:Name="KpiPend" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="34" FontWeight="Bold" Foreground="#B42318" Margin="0,8,0,0"/>
              <TextBlock x:Name="KpiPendF" Foreground="#6E6466" FontSize="13" Margin="0,6,0,0"/>
            </StackPanel>
          </Border>
          <Border Margin="8,0" Padding="22,18" CornerRadius="14">
            <Border.Background>
              <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                <GradientStop Color="#9E1820" Offset="0"/>
                <GradientStop Color="#560A0F" Offset="1"/>
              </LinearGradientBrush>
            </Border.Background>
            <Border.Effect><DropShadowEffect BlurRadius="22" ShadowDepth="5" Direction="270" Opacity="0.22" Color="#560A0F"/></Border.Effect>
            <StackPanel>
              <TextBlock x:Name="KpiNetL" Foreground="#F3D6D8" FontSize="14" FontWeight="SemiBold"/>
              <TextBlock x:Name="KpiNet" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="34" FontWeight="Bold" Foreground="White" Margin="0,8,0,0"/>
              <TextBlock x:Name="KpiNetF" Foreground="#F3D6D8" FontSize="13" Margin="0,6,0,0"/>
            </StackPanel>
          </Border>
        </UniformGrid>
        <!-- ANALÍTICA: día / semana / mes -->
        <Border Style="{StaticResource Card}" Padding="24,20" Margin="0,0,0,18">
          <StackPanel>
            <DockPanel Margin="0,0,0,6">
              <Border DockPanel.Dock="Right" Background="#F3EFEF" CornerRadius="10" Padding="3" VerticalAlignment="Center">
                <StackPanel Orientation="Horizontal">
                  <RadioButton x:Name="ChartDay" Style="{StaticResource Seg}" GroupName="chart"/>
                  <RadioButton x:Name="ChartWeek" Style="{StaticResource Seg}" GroupName="chart"/>
                  <RadioButton x:Name="ChartMonth" Style="{StaticResource Seg}" GroupName="chart" IsChecked="True"/>
                </StackPanel>
              </Border>
              <StackPanel>
                <StackPanel Orientation="Horizontal">
                  <TextBlock x:Name="ChartTitle" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="17" FontWeight="SemiBold" Foreground="#1E1A1B" VerticalAlignment="Center"/>
                  <Border x:Name="ChartDeltaB" CornerRadius="11" Padding="9,3" Margin="12,0,0,0" VerticalAlignment="Center" Visibility="Collapsed">
                    <TextBlock x:Name="ChartDelta" FontSize="12.5" FontWeight="SemiBold"/>
                  </Border>
                </StackPanel>
                <TextBlock x:Name="ChartSub" Foreground="#6E6466" FontSize="13" Margin="0,3,0,0"/>
              </StackPanel>
            </DockPanel>
            <StackPanel Orientation="Horizontal" Margin="0,8,0,10">
              <Border Width="18" Height="3" CornerRadius="1.5" Background="#1F9D55" VerticalAlignment="Center"/>
              <TextBlock x:Name="ChartLegPaid" Foreground="#3B3536" Margin="7,0,18,0" VerticalAlignment="Center"/>
              <Border Width="18" Height="3" CornerRadius="1.5" Background="#D8202A" VerticalAlignment="Center"/>
              <TextBlock x:Name="ChartLegPend" Foreground="#3B3536" Margin="7,0,18,0" VerticalAlignment="Center"/>
              <TextBlock Text="€" Foreground="#8A7F81" FontWeight="SemiBold" VerticalAlignment="Center"/>
              <TextBlock x:Name="ChartLegMoney" Foreground="#3B3536" Margin="6,0,0,0" VerticalAlignment="Center"/>
            </StackPanel>
            <Grid x:Name="ChartGrid">
              <Grid.RowDefinitions><RowDefinition Height="210"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
              <Canvas x:Name="ChartCanvas" ClipToBounds="False" Background="Transparent"/>
              <Grid x:Name="ChartLabels" Grid.Row="1" Margin="0,6,0,0"/>
            </Grid>
          </StackPanel>
        </Border>
        <Border Style="{StaticResource Card}" Padding="24,20">
          <DockPanel>
            <TextBlock x:Name="StatsTable" DockPanel.Dock="Top" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="17" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="0,0,0,12"/>
            <Grid x:Name="StatsGrid"/>
          </DockPanel>
        </Border>
      </StackPanel>
      </ScrollViewer>

      <!-- AJUSTES -->
      <ScrollViewer x:Name="PanelSettings" Visibility="Collapsed" VerticalScrollBarVisibility="Auto" Padding="0,10,8,0">
        <StackPanel MaxWidth="1040" HorizontalAlignment="Stretch">
          <DockPanel Margin="0,0,0,18">
            <Button x:Name="SaveSettingsBtn" Style="{StaticResource Primary}" DockPanel.Dock="Right" Height="42" Padding="24,0" FontSize="14" VerticalAlignment="Center"/>
            <StackPanel>
              <TextBlock x:Name="SetTitle" Style="{StaticResource Heading}"/>
              <TextBlock x:Name="SetSub" Foreground="#7A7072" Margin="0,4,0,0"/>
            </StackPanel>
          </DockPanel>
          <TextBlock x:Name="WelcomeText" Foreground="#C4161C" FontWeight="SemiBold" Margin="0,0,0,12" TextWrapping="Wrap" Visibility="Collapsed"/>
          <Border Style="{StaticResource Card}">
            <StackPanel>
            <Border x:Name="SecFileBox" Padding="28,22" BorderBrush="#F1EDED" BorderThickness="0,0,0,1">
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="270"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                <StackPanel Margin="0,0,28,0">
                  <StackPanel Orientation="Horizontal">
                    <Border Width="30" Height="30" CornerRadius="15" Background="#FBEDEE">
                      <TextBlock Text="&#xE8A5;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="13" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <TextBlock x:Name="SecFileT" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="15.5" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="10,0,0,0" VerticalAlignment="Center"/>
                  </StackPanel>
                  <TextBlock x:Name="SecFileD" Foreground="#7A7072" FontSize="12.5" TextWrapping="Wrap" LineHeight="19" Margin="40,6,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1">
                  <TextBlock x:Name="SetFileL" Style="{StaticResource FieldLabel}"/>
                  <DockPanel>
                    <Button x:Name="ChangeFileBtn" Style="{StaticResource Btn}" DockPanel.Dock="Right" Margin="8,0,0,0"/>
                    <TextBox x:Name="FilePathBox" IsReadOnly="True" Background="#F8F6F6"/>
                  </DockPanel>
                  <DockPanel Margin="0,10,0,0">
                    <Button x:Name="ImportXlBtn" Style="{StaticResource Btn}" DockPanel.Dock="Right" Margin="8,0,0,0"/>
                    <TextBlock x:Name="DataNote" Foreground="#7A7072" FontSize="12.5" TextWrapping="Wrap" VerticalAlignment="Center"/>
                  </DockPanel>
                  <TextBlock x:Name="SetNote" Foreground="#7A7072" FontSize="12.5" TextWrapping="Wrap" Margin="1,10,0,0"/>
                </StackPanel>
              </Grid>
            </Border>
            <Border x:Name="SecExtBox" Padding="28,22" BorderBrush="#F1EDED" BorderThickness="0,0,0,1">
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="270"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                <StackPanel Margin="0,0,28,0">
                  <StackPanel Orientation="Horizontal">
                    <Border Width="30" Height="30" CornerRadius="15" Background="#FBEDEE">
                      <TextBlock Text="&#xE753;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="13" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <TextBlock x:Name="SecExtT" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="15.5" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="10,0,0,0" VerticalAlignment="Center"/>
                  </StackPanel>
                  <TextBlock x:Name="SecExtD" Foreground="#7A7072" FontSize="12.5" TextWrapping="Wrap" LineHeight="19" Margin="40,6,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1">
                  <TextBlock x:Name="ExtDirL" Style="{StaticResource FieldLabel}"/>
                  <DockPanel>
                    <Button x:Name="ExtClearBtn" Style="{StaticResource Btn}" DockPanel.Dock="Right" Margin="8,0,0,0" Padding="12,0"/>
                    <Button x:Name="ExtDirBtn" Style="{StaticResource Btn}" DockPanel.Dock="Right" Margin="8,0,0,0"/>
                    <TextBox x:Name="ExtDirBox" IsReadOnly="True" Background="#F8F6F6"/>
                  </DockPanel>
                  <WrapPanel Margin="0,8,0,0">
                    <Button x:Name="ExtOneDriveBtn" Style="{StaticResource Btn}" Margin="0,0,8,8" Visibility="Collapsed"/>
                    <Button x:Name="ExtNowBtn" Style="{StaticResource Btn}" Margin="0,0,8,8"/>
                    <Button x:Name="ExtOpenBtn" Style="{StaticResource Btn}" Margin="0,0,8,8"/>
                  </WrapPanel>
                  <StackPanel Orientation="Horizontal" Margin="1,2,0,0">
                    <TextBlock Text="&#xE73E;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="12" Foreground="#1B7F45" VerticalAlignment="Center" Margin="0,0,6,0"/>
                    <TextBlock x:Name="ExtLastTxt" FontSize="12.5" FontWeight="SemiBold" VerticalAlignment="Center"/>
                  </StackPanel>
                  <TextBlock x:Name="ExtNote" Foreground="#7A7072" FontSize="12.5" TextWrapping="Wrap" Margin="1,10,0,0"/>
                </StackPanel>
              </Grid>
            </Border>
            <Border x:Name="SecCompBox" Padding="28,22" BorderBrush="#F1EDED" BorderThickness="0,0,0,1">
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="270"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                <StackPanel Margin="0,0,28,0">
                  <StackPanel Orientation="Horizontal">
                    <Border Width="30" Height="30" CornerRadius="15" Background="#FBEDEE">
                      <TextBlock Text="&#xE80F;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="13" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <TextBlock x:Name="SecCompT" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="15.5" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="10,0,0,0" VerticalAlignment="Center"/>
                  </StackPanel>
                  <TextBlock x:Name="SecCompD" Foreground="#7A7072" FontSize="12.5" TextWrapping="Wrap" LineHeight="19" Margin="40,6,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1">
                  <TextBlock x:Name="SetCompanyL" Style="{StaticResource FieldLabel}"/>
                  <TextBox x:Name="CompanyBox"/>
                  <TextBlock x:Name="SetLogoL" Style="{StaticResource FieldLabel}" Margin="1,16,0,6"/>
                  <StackPanel Orientation="Horizontal">
                    <Grid Width="52" Height="52">
                      <Ellipse Fill="White" Stroke="#E4DFE0" StrokeThickness="1"/>
                      <Ellipse x:Name="LogoPreview" Margin="3"/>
                    </Grid>
                    <Button x:Name="LogoBtn" Style="{StaticResource Btn}" Margin="14,0,0,0" VerticalAlignment="Center"/>
                    <Button x:Name="LogoClearBtn" Style="{StaticResource Btn}" Margin="8,0,0,0" Padding="12,0" VerticalAlignment="Center"/>
                  </StackPanel>
                </StackPanel>
              </Grid>
            </Border>
            <Border x:Name="SecTeamBox" Padding="28,22" BorderBrush="#F1EDED" BorderThickness="0,0,0,1">
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="270"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                <StackPanel Margin="0,0,28,0">
                  <StackPanel Orientation="Horizontal">
                    <Border Width="30" Height="30" CornerRadius="15" Background="#FBEDEE">
                      <TextBlock Text="&#xE716;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="13" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <TextBlock x:Name="SecTeamT" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="15.5" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="10,0,0,0" VerticalAlignment="Center"/>
                  </StackPanel>
                  <TextBlock x:Name="SecTeamD" Foreground="#7A7072" FontSize="12.5" TextWrapping="Wrap" LineHeight="19" Margin="40,6,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1">
                  <TextBlock x:Name="SetCollL" Style="{StaticResource FieldLabel}"/>
                  <TextBox x:Name="ColleaguesBox"/>
                </StackPanel>
              </Grid>
            </Border>
            <Border x:Name="SecPrefBox" Padding="28,22" BorderBrush="#F1EDED" BorderThickness="0,0,0,1">
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="270"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                <StackPanel Margin="0,0,28,0">
                  <StackPanel Orientation="Horizontal">
                    <Border Width="30" Height="30" CornerRadius="15" Background="#FBEDEE">
                      <TextBlock Text="&#xE713;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="13" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <TextBlock x:Name="SecPrefT" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="15.5" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="10,0,0,0" VerticalAlignment="Center"/>
                  </StackPanel>
                  <TextBlock x:Name="SecPrefD" Foreground="#7A7072" FontSize="12.5" TextWrapping="Wrap" LineHeight="19" Margin="40,6,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1">
                  <Grid>
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
                </StackPanel>
              </Grid>
            </Border>
            <Border x:Name="SecHelpBox" Padding="28,22">
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="270"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                <StackPanel Margin="0,0,28,0">
                  <StackPanel Orientation="Horizontal">
                    <Border Width="30" Height="30" CornerRadius="15" Background="#FBEDEE">
                      <TextBlock Text="&#xE897;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="13" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <TextBlock x:Name="SecHelpT" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="15.5" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="10,0,0,0" VerticalAlignment="Center"/>
                  </StackPanel>
                  <TextBlock x:Name="SecHelpD" Foreground="#7A7072" FontSize="12.5" TextWrapping="Wrap" LineHeight="19" Margin="40,6,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1">
                  <WrapPanel>
                    <Button x:Name="HelpBtn" Style="{StaticResource Btn}" Margin="0,0,8,8" Height="40"/>
                    <Button x:Name="UpdateBtn" Style="{StaticResource Btn}" Margin="0,0,8,8" Height="40"/>
                    <Button x:Name="BackupsBtn" Style="{StaticResource Btn}" Margin="0,0,8,8" Height="40"/>
                  </WrapPanel>
                </StackPanel>
              </Grid>
            </Border>
            </StackPanel>
          </Border>
        </StackPanel>
      </ScrollViewer>

      <!-- ACERCA DE -->
      <Border x:Name="PanelAbout" Visibility="Collapsed" Style="{StaticResource Card}" Padding="38,32" HorizontalAlignment="Center" Width="620" VerticalAlignment="Top" Margin="0,30,0,0">
        <StackPanel>
          <StackPanel Orientation="Horizontal">
            <Border x:Name="AboutLogo" Width="72" Height="72" Background="Transparent" Cursor="Hand">
              <Image x:Name="AboutImg" RenderOptions.BitmapScalingMode="HighQuality"/>
            </Border>
            <StackPanel Margin="18,0,0,0" VerticalAlignment="Center">
              <TextBlock Text="GESTEM MANAGER" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="26" FontWeight="SemiBold" Foreground="#1E1A1B"/>
              <TextBlock x:Name="AboutVer" Foreground="#667085"/>
              <TextBlock x:Name="AboutCompany" Foreground="#C4161C" FontWeight="SemiBold" Margin="0,2,0,0"/>
            </StackPanel>
          </StackPanel>
          <TextBlock x:Name="AboutText" TextWrapping="Wrap" Margin="0,24,0,0" FontSize="14" LineHeight="22" Foreground="#344054"/>
          <TextBlock x:Name="AboutBy" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="16" FontWeight="SemiBold" Foreground="#C4161C" Margin="0,20,0,0"/>
          <TextBlock Text="© 2026 Yano · Todos los derechos reservados" Foreground="#6E6466" FontWeight="SemiBold" Margin="0,4,0,0"/>
          <TextBlock Text="Programa con licencia de uso para la empresa titular. Prohibida su copia, modificación, distribución o reventa sin permiso escrito del autor." Foreground="#98A2B3" FontSize="12" TextWrapping="Wrap" Margin="0,4,0,0"/>
          <TextBlock x:Name="EasterEgg" Text="Por Safae de Yan ♥" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="20" Foreground="#C2185B" Margin="0,22,0,0" Visibility="Collapsed"/>
        </StackPanel>
      </Border>

    </Grid>
    </DockPanel>
    </Grid>
  </DockPanel>
  <Grid x:Name="PanelSession" Background="#E0141112" Visibility="Collapsed">
    <Border Background="White" CornerRadius="18" Padding="36,30" Width="540" HorizontalAlignment="Center" VerticalAlignment="Center">
      <Border.Effect><DropShadowEffect BlurRadius="40" ShadowDepth="8" Direction="270" Opacity="0.35" Color="#000000"/></Border.Effect>
      <StackPanel>
        <StackPanel Orientation="Horizontal">
          <Ellipse x:Name="SessLogo" Width="58" Height="58"/>
          <StackPanel Margin="14,0,0,0" VerticalAlignment="Center">
            <TextBlock x:Name="SessCompany" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="17" FontWeight="Bold" Foreground="#1E1A1B"/>
            <TextBlock FontSize="12" Foreground="#667085" Margin="0,2,0,0"><Run Text="GESTEM" FontWeight="SemiBold" Foreground="#C4161C"/><Run Text=" Manager"/></TextBlock>
          </StackPanel>
        </StackPanel>
        <TextBlock x:Name="SessTitle" Style="{StaticResource Heading}" FontSize="22" Margin="0,26,0,6"/>
        <TextBlock x:Name="SessHint" Foreground="#667085" TextWrapping="Wrap"/>
        <WrapPanel x:Name="SessNames" Margin="0,18,0,6"/>
        <DockPanel Margin="0,8,0,0">
          <Button x:Name="SessGo" Style="{StaticResource Primary}" DockPanel.Dock="Right" Height="38" MinWidth="100" FontSize="13" Margin="8,0,0,0"/>
          <Grid>
            <TextBox x:Name="SessNew" CharacterCasing="Upper" Height="38"/>
            <TextBlock x:Name="SessNewHint" IsHitTestVisible="False" Foreground="#98A2B3" Margin="11,0,0,0" VerticalAlignment="Center"/>
          </Grid>
        </DockPanel>
      </StackPanel>
    </Border>
  </Grid>
  <Grid x:Name="PanelHelp" Background="#C0141112" Visibility="Collapsed">
    <Border Background="White" CornerRadius="18" Width="780" Margin="0,40" HorizontalAlignment="Center" VerticalAlignment="Stretch">
      <Border.Effect><DropShadowEffect BlurRadius="40" ShadowDepth="8" Direction="270" Opacity="0.35" Color="#000000"/></Border.Effect>
      <DockPanel>
        <Border DockPanel.Dock="Top" CornerRadius="18,18,0,0" Padding="34,24,24,20">
          <Border.Background>
            <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
              <GradientStop Color="#9E1820" Offset="0"/>
              <GradientStop Color="#560A0F" Offset="1"/>
            </LinearGradientBrush>
          </Border.Background>
          <DockPanel>
            <Button x:Name="HelpClose" DockPanel.Dock="Right" VerticalAlignment="Top" Cursor="Hand" ToolTip="Cerrar">
              <Button.Template>
                <ControlTemplate TargetType="Button">
                  <Border x:Name="b" Width="34" Height="34" CornerRadius="17" Background="#26FFFFFF">
                    <TextBlock Text="&#xE711;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="12" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                  </Border>
                  <ControlTemplate.Triggers>
                    <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Background" Value="#4DFFFFFF"/></Trigger>
                  </ControlTemplate.Triggers>
                </ControlTemplate>
              </Button.Template>
            </Button>
            <StackPanel>
              <TextBlock Text="Manual de uso" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="24" FontWeight="SemiBold" Foreground="White"/>
              <TextBlock Text="GESTEM Manager · todo lo que necesitas para trabajar con el programa" Foreground="#F3D6D8" Margin="0,4,0,0"/>
            </StackPanel>
          </DockPanel>
        </Border>
        <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="0,0,6,0">
          <StackPanel Margin="34,26,30,10">
          <Grid Margin="0,0,0,22">
            <Grid.ColumnDefinitions><ColumnDefinition Width="52"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Border Width="38" Height="38" CornerRadius="19" Background="#FBEDEE" VerticalAlignment="Top">
              <TextBlock Text="&#xE77B;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="16" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Grid.Column="1">
              <TextBlock Text="1. Empezar" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="16" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="0,8,0,8"/>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Al abrir GESTEM, elige tu nombre en «¿Quién está usando GESTEM?». Si no estás en la lista, escríbelo en «Otro nombre…» y pulsa «Entrar»." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Tu nombre sale abajo a la izquierda. Haz clic en él para cambiar de compañera sin cerrar el programa." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Arriba, en «Trámite», elige con qué trabajas: Canje DGT, Homologación, Extranjería, Autónomo… Cada trámite tiene su pestaña en el Excel." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
            </StackPanel>
          </Grid>
          <Grid Margin="0,0,0,22">
            <Grid.ColumnDefinitions><ColumnDefinition Width="52"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Border Width="38" Height="38" CornerRadius="19" Background="#FBEDEE" VerticalAlignment="Top">
              <TextBlock Text="&#xE8FA;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="16" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Grid.Column="1">
              <TextBlock Text="2. Añadir un cliente nuevo" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="16" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="0,8,0,8"/>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Menú «Nuevo cliente». Rellena los datos de arriba abajo: nombre, documento, teléfono, datos del trámite, estado y pago." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="La nacionalidad (MARROQUI), la fecha de hoy y el responsable (tu nombre) se ponen solos. Puedes cambiarlos." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Si el NIE o pasaporte ya existe, sale un aviso en rojo para no repetir clientes." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Pulsa «Guardar cliente» (o Ctrl + S). El cliente se escribe en el Excel al momento, con su número de expediente." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
            </StackPanel>
          </Grid>
          <Grid Margin="0,0,0,22">
            <Grid.ColumnDefinitions><ColumnDefinition Width="52"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Border Width="38" Height="38" CornerRadius="19" Background="#FBEDEE" VerticalAlignment="Top">
              <TextBlock Text="&#xE8C7;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="16" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Grid.Column="1">
              <TextBlock Text="3. Pagos" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="16" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="0,8,0,8"/>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Precio total: lo que paga el cliente. En el canje DGT se pone solo según el tipo de permiso." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Estado de pago: «Pagado» (lo ha pagado todo), «Parcial» (escribe cuánto ha pagado en «Ha pagado») o «Pendiente» (no ha pagado nada)." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Gastos: lo que pagáis vosotras (tasa, traducción…). El beneficio NETO = precio − gastos se calcula solo en el Excel." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Lo que falta por cobrar se calcula solo y sale en rojo." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
            </StackPanel>
          </Grid>
          <Grid Margin="0,0,0,22">
            <Grid.ColumnDefinitions><ColumnDefinition Width="52"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Border Width="38" Height="38" CornerRadius="19" Background="#FBEDEE" VerticalAlignment="Top">
              <TextBlock Text="&#xE716;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="16" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Grid.Column="1">
              <TextBlock Text="4. Buscar y cambiar clientes" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="16" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="0,8,0,8"/>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Menú «Clientes». Busca por nombre, NIE, número de expediente o localizador." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Filtra por estado, marca «Solo con pago pendiente» para ver quién debe dinero, o «Solo mis clientes» para ver solo los tuyos." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Doble clic en un cliente (o «Editar cliente») para cambiar su estado, su pago o sus datos. Pulsa «Guardar cliente» al terminar." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="«Abrir Excel» abre el archivo; «Actualizar» vuelve a leerlo si alguien lo ha cambiado a mano." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
            </StackPanel>
          </Grid>
          <Grid Margin="0,0,0,22">
            <Grid.ColumnDefinitions><ColumnDefinition Width="52"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Border Width="38" Height="38" CornerRadius="19" Background="#FBEDEE" VerticalAlignment="Top">
              <TextBlock Text="&#xE723;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="16" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Grid.Column="1">
              <TextBlock Text="5. Documentos y recibos" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="16" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="0,8,0,8"/>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="En «Clientes», elige un cliente y pulsa «Documentos» (o clic derecho sobre él). Arrastra los escaneos (NIE, permiso, formularios, fotos…) o pulsa «Añadir archivos»." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Los documentos se guardan en la carpeta «GESTEM documentos», junto al Excel, una carpeta por expediente. El clip de la lista indica cuántos tiene cada cliente." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="«Recibo PDF» crea un recibo de pago con el logo de la empresa, lo guarda en los documentos del cliente y lo abre para imprimirlo o enviarlo." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="El recibo justifica lo pagado; no es una factura." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
            </StackPanel>
          </Grid>
          <Grid Margin="0,0,0,22">
            <Grid.ColumnDefinitions><ColumnDefinition Width="52"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Border Width="38" Height="38" CornerRadius="19" Background="#FBEDEE" VerticalAlignment="Top">
              <TextBlock Text="&#xEDE1;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="16" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Grid.Column="1">
              <TextBlock Text="6. Exportar la lista" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="16" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="0,8,0,8"/>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="En «Clientes», pulsa «Exportar»: «Lista en PDF» (informe para imprimir, con totales) o «Lista en CSV» (se abre en Excel con todas las columnas)." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Se exporta lo que ves: si has buscado o filtrado, solo salen esos clientes." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
            </StackPanel>
          </Grid>
          <Grid Margin="0,0,0,22">
            <Grid.ColumnDefinitions><ColumnDefinition Width="52"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Border Width="38" Height="38" CornerRadius="19" Background="#FBEDEE" VerticalAlignment="Top">
              <TextBlock Text="&#xE9D2;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="16" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Grid.Column="1">
              <TextBlock Text="7. Resumen económico" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="16" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="0,8,0,8"/>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Menú «Resumen». Arriba ves lo facturado, lo cobrado, lo pendiente de cobro y el beneficio neto." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Arriba puedes elegir «Todos los trámites» o uno solo; las cifras y la gráfica cambian con él." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="La gráfica muestra la evolución por día, semana o mes: verde cuando sube, rojo cuando baja." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Filtra por compañera y por periodo: este mes, el mes anterior o este año (según la fecha de solicitud)." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
            </StackPanel>
          </Grid>
          <Grid Margin="0,0,0,22">
            <Grid.ColumnDefinitions><ColumnDefinition Width="52"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Border Width="38" Height="38" CornerRadius="19" Background="#FBEDEE" VerticalAlignment="Top">
              <TextBlock Text="&#xE8A5;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="16" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Grid.Column="1">
              <TextBlock Text="8. El Excel y las copias" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="16" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="0,8,0,8"/>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="GESTEM guarda los clientes en su propia carpeta «GESTEM datos», junto al Excel, y actualiza vuestro Excel en cada guardado. Si el Excel no se puede abrir en ese momento, se actualiza solo más tarde." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Antes de cada guardado se hace una copia de seguridad del Excel (Ajustes → «Abrir carpeta de copias»)." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Copia externa: al abrir y al cerrar GESTEM se guarda una copia completa (Excel, ajustes y documentos) en la nube de la empresa (OneDrive o Google Drive) o en la carpeta elegida en Ajustes." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
            </StackPanel>
          </Grid>
          <Grid Margin="0,0,0,22">
            <Grid.ColumnDefinitions><ColumnDefinition Width="52"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Border Width="38" Height="38" CornerRadius="19" Background="#FBEDEE" VerticalAlignment="Top">
              <TextBlock Text="&#xE713;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="16" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Grid.Column="1">
              <TextBlock Text="9. Ajustes" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="16" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="0,8,0,8"/>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Archivo de clientes: «Cambiar…» para elegir otro Excel." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Copia externa: «Usar la nube de la empresa» la configura sola; también puedes elegir un USB o una carpeta." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Empresa: nombre y logo. Equipo: los nombres de las compañeras. Preferencias: idioma (español, inglés, francés o árabe) y nacionalidad por defecto." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Pulsa «Guardar ajustes» para aplicar los cambios." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
            </StackPanel>
          </Grid>
          <Grid Margin="0,0,0,22">
            <Grid.ColumnDefinitions><ColumnDefinition Width="52"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Border Width="38" Height="38" CornerRadius="19" Background="#FBEDEE" VerticalAlignment="Top">
              <TextBlock Text="&#xE72E;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="16" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Grid.Column="1">
              <TextBlock Text="10. Privacidad y registro" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="16" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="0,8,0,8"/>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Por seguridad, GESTEM anota quién añade o cambia cada cliente, quién crea recibos y cuándo se abre y se cierra el programa." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Una copia de ese registro, sin nombres de clientes, se guarda junto a la copia externa. Los datos de los clientes nunca salen de vuestro PC ni de la nube de la empresa." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
            </StackPanel>
          </Grid>
          <Grid Margin="0,0,0,22">
            <Grid.ColumnDefinitions><ColumnDefinition Width="52"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Border Width="38" Height="38" CornerRadius="19" Background="#FBEDEE" VerticalAlignment="Top">
              <TextBlock Text="&#xE946;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="16" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <StackPanel Grid.Column="1">
              <TextBlock Text="11. Actualizaciones y ayuda" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="16" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="0,8,0,8"/>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="GESTEM se actualiza solo al abrirse cuando hay una versión nueva. También puedes pulsar «Buscar actualizaciones»." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="Si algo no funciona, cierra GESTEM y vuelve a abrirlo. Si sigue igual, avisa a Yano." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
              <DockPanel Margin="0,0,0,6"><Ellipse DockPanel.Dock="Left" Width="5" Height="5" Fill="#C4161C" Margin="2,8,10,0" VerticalAlignment="Top"/><TextBlock Text="GESTEM MANAGER © 2026 Yano. Todos los derechos reservados. Uso bajo licencia (ver LICENCIA.txt)." TextWrapping="Wrap" Foreground="#3B3536" FontSize="13.5" LineHeight="21"/></DockPanel>
            </StackPanel>
          </Grid>
          </StackPanel>
        </ScrollViewer>
      </DockPanel>
    </Border>
  </Grid>
  <Grid x:Name="PanelDocs" Background="#C0141112" Visibility="Collapsed" AllowDrop="True">
    <Border x:Name="DocsCard" Background="White" CornerRadius="18" Width="720" Margin="0,50" HorizontalAlignment="Center" VerticalAlignment="Stretch" AllowDrop="True">
      <Border.Effect><DropShadowEffect BlurRadius="40" ShadowDepth="8" Direction="270" Opacity="0.35" Color="#000000"/></Border.Effect>
      <DockPanel>
        <Border DockPanel.Dock="Top" CornerRadius="18,18,0,0" Padding="30,22,22,18">
          <Border.Background>
            <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
              <GradientStop Color="#9E1820" Offset="0"/>
              <GradientStop Color="#560A0F" Offset="1"/>
            </LinearGradientBrush>
          </Border.Background>
          <DockPanel>
            <Button x:Name="DocsClose" DockPanel.Dock="Right" VerticalAlignment="Top" Cursor="Hand" ToolTip="Cerrar">
              <Button.Template>
                <ControlTemplate TargetType="Button">
                  <Border x:Name="b" Width="34" Height="34" CornerRadius="17" Background="#26FFFFFF">
                    <TextBlock Text="&#xE711;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="12" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                  </Border>
                  <ControlTemplate.Triggers>
                    <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Background" Value="#4DFFFFFF"/></Trigger>
                  </ControlTemplate.Triggers>
                </ControlTemplate>
              </Button.Template>
            </Button>
            <StackPanel>
              <StackPanel Orientation="Horizontal">
                <TextBlock Text="&#xE723;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="18" Foreground="White" VerticalAlignment="Center" Margin="0,0,10,0"/>
                <TextBlock x:Name="DocsTitle" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="22" FontWeight="SemiBold" Foreground="White"/>
              </StackPanel>
              <TextBlock x:Name="DocsSub" Foreground="#F3D5D7" Margin="0,4,0,0" TextTrimming="CharacterEllipsis"/>
            </StackPanel>
          </DockPanel>
        </Border>
        <Border DockPanel.Dock="Bottom" BorderBrush="#EEE9EA" BorderThickness="0,1,0,0" Padding="24,14">
          <DockPanel>
            <StackPanel Orientation="Horizontal" DockPanel.Dock="Right">
              <Button x:Name="DocsOpenBtn" Style="{StaticResource Btn}" Margin="8,0,0,0"/>
              <Button x:Name="DocsRemoveBtn" Style="{StaticResource Btn}" Margin="8,0,0,0"/>
              <Button x:Name="DocsFolderBtn" Style="{StaticResource Btn}" Margin="8,0,0,0"/>
              <Button x:Name="DocsAddBtn" Style="{StaticResource Primary}" Height="38" Padding="18,0" FontSize="13" Margin="8,0,0,0"/>
            </StackPanel>
            <StackPanel VerticalAlignment="Center">
              <TextBlock x:Name="DocsCount" Foreground="#6E6466" FontWeight="SemiBold"/>
              <TextBlock x:Name="DocsMsg" Foreground="#1B7F45" FontWeight="SemiBold"/>
            </StackPanel>
          </DockPanel>
        </Border>
        <Grid Margin="24,16,24,10">
          <ListBox x:Name="DocsList" BorderThickness="0" Background="Transparent" ScrollViewer.HorizontalScrollBarVisibility="Disabled">
            <ListBox.ItemContainerStyle>
              <Style TargetType="ListBoxItem">
                <Setter Property="Padding" Value="0"/>
                <Setter Property="Margin" Value="0,0,0,6"/>
                <Setter Property="Cursor" Value="Hand"/>
                <Setter Property="Template">
                  <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                      <Border x:Name="bd" Background="#FAF8F8" BorderBrush="#EEE9EA" BorderThickness="1" CornerRadius="10" Padding="14,10">
                        <ContentPresenter/>
                      </Border>
                      <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#F5EFF0"/></Trigger>
                        <Trigger Property="IsSelected" Value="True"><Setter TargetName="bd" Property="Background" Value="#FBEDEE"/><Setter TargetName="bd" Property="BorderBrush" Value="#D8A7AB"/></Trigger>
                      </ControlTemplate.Triggers>
                    </ControlTemplate>
                  </Setter.Value>
                </Setter>
              </Style>
            </ListBox.ItemContainerStyle>
            <ListBox.ItemTemplate>
              <DataTemplate>
                <DockPanel>
                  <Border Width="36" Height="36" CornerRadius="8" Background="#FBEDEE" Margin="0,0,12,0">
                    <TextBlock Text="{Binding Glyph}" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="16" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                  </Border>
                  <StackPanel VerticalAlignment="Center">
                    <TextBlock Text="{Binding Name}" FontWeight="SemiBold" Foreground="#1E1A1B" TextTrimming="CharacterEllipsis"/>
                    <TextBlock Text="{Binding Info}" Foreground="#8A7F81" FontSize="12" Margin="0,2,0,0"/>
                  </StackPanel>
                </DockPanel>
              </DataTemplate>
            </ListBox.ItemTemplate>
          </ListBox>
          <StackPanel x:Name="DocsEmpty" VerticalAlignment="Center" HorizontalAlignment="Center" IsHitTestVisible="False">
            <Border Width="64" Height="64" CornerRadius="32" Background="#FBEDEE" HorizontalAlignment="Center">
              <TextBlock Text="&#xE898;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="24" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <TextBlock x:Name="DocsEmptyT" Foreground="#6E6466" TextWrapping="Wrap" TextAlignment="Center" MaxWidth="420" Margin="0,14,0,0" LineHeight="20"/>
          </StackPanel>
          <Border x:Name="DocsDrop" CornerRadius="12" BorderBrush="#9E1820" BorderThickness="2" Background="#E6FFFFFF" Visibility="Collapsed" IsHitTestVisible="False">
            <TextBlock x:Name="DocsDropT" Foreground="#9E1820" FontSize="16" FontWeight="SemiBold" HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
        </Grid>
      </DockPanel>
    </Border>
  </Grid>
  <Grid x:Name="PanelPaused" Background="#F2141112" Visibility="Collapsed">
    <Border Background="White" CornerRadius="18" Padding="40,34" Width="520" HorizontalAlignment="Center" VerticalAlignment="Center">
      <StackPanel>
        <Border Width="56" Height="56" CornerRadius="28" Background="#FBEDEE" HorizontalAlignment="Left">
          <TextBlock Text="&#xE769;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="22" Foreground="#9E1820" HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>
        <TextBlock Text="GESTEM está en pausa" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="22" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="0,18,0,8"/>
        <TextBlock x:Name="PausedText" Foreground="#4B4345" FontSize="14" TextWrapping="Wrap" LineHeight="21"/>
        <TextBlock Text="Contacta con Yano." Foreground="#8A7F81" Margin="0,14,0,0"/>
      </StackPanel>
    </Border>
  </Grid>
  <Grid x:Name="PanelAdminLogin" Background="#E6141112" Visibility="Collapsed">
    <Border Background="White" CornerRadius="18" Padding="36,30" Width="440" HorizontalAlignment="Center" VerticalAlignment="Center">
      <Border.Effect><DropShadowEffect BlurRadius="40" ShadowDepth="8" Direction="270" Opacity="0.35" Color="#000000"/></Border.Effect>
      <StackPanel>
        <StackPanel Orientation="Horizontal">
          <Border Width="40" Height="40" CornerRadius="20" Background="#221E1F">
            <TextBlock Text="&#xE72E;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="16" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <TextBlock Text="Administrador" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="20" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="12,0,0,0" VerticalAlignment="Center"/>
        </StackPanel>
        <TextBlock x:Name="AdmLoginHint" Foreground="#6E6466" TextWrapping="Wrap" Margin="0,16,0,12"/>
        <PasswordBox x:Name="AdmPw" Height="38" Padding="10,0" VerticalContentAlignment="Center" BorderBrush="#D9D4D5"/>
        <PasswordBox x:Name="AdmPw2" Height="38" Padding="10,0" VerticalContentAlignment="Center" BorderBrush="#D9D4D5" Margin="0,8,0,0" Visibility="Collapsed"/>
        <TextBlock x:Name="AdmKeyHint" Foreground="#8A7F81" FontSize="12" Margin="0,6,0,0" TextWrapping="Wrap"/>
        <TextBlock x:Name="AdmLoginMsg" Foreground="#B42318" FontWeight="SemiBold" Margin="0,8,0,0" TextWrapping="Wrap"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
          <Button x:Name="AdmLoginCancel" Style="{StaticResource Btn}" Content="Cancelar"/>
          <Button x:Name="AdmLoginBtn" Style="{StaticResource Primary}" Height="38" Padding="22,0" FontSize="13" Margin="8,0,0,0" Content="Entrar"/>
        </StackPanel>
      </StackPanel>
    </Border>
  </Grid>
  <Grid x:Name="PanelAdmin" Background="#C0141112" Visibility="Collapsed">
    <Border Background="#F7F5F5" CornerRadius="18" Width="900" Margin="0,40" HorizontalAlignment="Center" VerticalAlignment="Stretch">
      <Border.Effect><DropShadowEffect BlurRadius="40" ShadowDepth="8" Direction="270" Opacity="0.35" Color="#000000"/></Border.Effect>
      <DockPanel>
        <Border DockPanel.Dock="Top" CornerRadius="18,18,0,0" Padding="30,22,24,18">
          <Border.Background>
            <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
              <GradientStop Color="#2A2526" Offset="0"/>
              <GradientStop Color="#141112" Offset="1"/>
            </LinearGradientBrush>
          </Border.Background>
          <DockPanel>
            <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Top">
              <Button x:Name="AdmLogout" Style="{StaticResource Btn}" Content="Salir del modo admin" Height="34" Margin="0,0,8,0"/>
              <Button x:Name="AdmClose" Cursor="Hand" ToolTip="Cerrar">
                <Button.Template>
                  <ControlTemplate TargetType="Button">
                    <Border x:Name="b" Width="34" Height="34" CornerRadius="17" Background="#26FFFFFF">
                      <TextBlock Text="&#xE711;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="12" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <ControlTemplate.Triggers>
                      <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Background" Value="#4DFFFFFF"/></Trigger>
                    </ControlTemplate.Triggers>
                  </ControlTemplate>
                </Button.Template>
              </Button>
            </StackPanel>
            <StackPanel>
              <StackPanel Orientation="Horizontal">
                <TextBlock Text="&#xE72E;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="18" Foreground="#E0262E" VerticalAlignment="Center" Margin="0,0,10,0"/>
                <TextBlock Text="Modo administrador" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="22" FontWeight="SemiBold" Foreground="White"/>
              </StackPanel>
              <TextBlock Text="Solo visible para Yano · © 2026 todos los derechos reservados" Foreground="#A69C9E" Margin="0,4,0,0"/>
            </StackPanel>
          </DockPanel>
        </Border>
        <Border DockPanel.Dock="Top" Padding="30,16,30,0">
          <Border Background="#ECE7E8" CornerRadius="10" Padding="3" HorizontalAlignment="Left">
            <StackPanel Orientation="Horizontal">
              <RadioButton x:Name="AdmTabLocks" Style="{StaticResource Seg}" GroupName="adm" IsChecked="True" Content="Bloqueos"/>
              <RadioButton x:Name="AdmTabLog" Style="{StaticResource Seg}" GroupName="adm" Content="Actividad"/>
              <RadioButton x:Name="AdmTabTools" Style="{StaticResource Seg}" GroupName="adm" Content="Herramientas"/>
            </StackPanel>
          </Border>
        </Border>
        <Grid Margin="30,16,30,26">
          <!-- BLOQUEOS -->
          <Border x:Name="AdmPageLocks" Style="{StaticResource Card}" Padding="26,22" VerticalAlignment="Top">
            <StackPanel>
              <TextBlock Text="Oculta partes de Ajustes a las compañeras" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="16" FontWeight="SemiBold" Foreground="#1E1A1B"/>
              <TextBlock Text="Lo que bloquees desaparece de Ajustes para ellas (no lo ven ni lo pueden cambiar). Tú lo sigues viendo mientras estés en modo admin." Foreground="#6E6466" TextWrapping="Wrap" Margin="0,4,0,14"/>
              <CheckBox x:Name="AdmLockFile" Content="Archivo de clientes (Excel)"/>
              <CheckBox x:Name="AdmLockExt" Content="Copia externa"/>
              <CheckBox x:Name="AdmLockComp" Content="Empresa (nombre y logo)"/>
              <CheckBox x:Name="AdmLockTeam" Content="Equipo (compañeras)"/>
              <CheckBox x:Name="AdmLockPref" Content="Preferencias (idioma, nacionalidad, copias)"/>
              <CheckBox x:Name="AdmLockHelp" Content="Ayuda y versión (actualizaciones)"/>
              <Border Height="1" Background="#EEE9EA" Margin="0,10"/>
              <CheckBox x:Name="AdmLockAll" Content="Ocultar Ajustes por completo"/>
              <CheckBox x:Name="AdmIgnorePause" Content="Ignorar la pausa remota en este PC"/>
              <CheckBox x:Name="AdmOwnerPC" Content="Este PC es mío: todo lo que se haga aquí cuenta como admin"/>
              <StackPanel Orientation="Horizontal" Margin="0,16,0,0">
                <Button x:Name="AdmSaveLocks" Style="{StaticResource Primary}" Height="40" Padding="22,0" FontSize="14" Content="Guardar bloqueos"/>
                <TextBlock x:Name="AdmLocksMsg" Foreground="#1B7F45" FontWeight="SemiBold" VerticalAlignment="Center" Margin="12,0,0,0"/>
              </StackPanel>
            </StackPanel>
          </Border>
          <!-- ACTIVIDAD -->
          <Border x:Name="AdmPageLog" Style="{StaticResource Card}" Padding="22,18" Visibility="Collapsed">
            <DockPanel>
              <DockPanel DockPanel.Dock="Top" Margin="0,0,0,10">
                <Button x:Name="AdmLogRefresh" Style="{StaticResource Btn}" DockPanel.Dock="Right" Content="Actualizar" Height="34"/>
                <Border DockPanel.Dock="Right" Background="#ECE7E8" CornerRadius="10" Padding="3" Margin="0,0,10,0" VerticalAlignment="Center">
                  <StackPanel Orientation="Horizontal">
                    <RadioButton x:Name="AdmLogAll" Style="{StaticResource Seg}" GroupName="admlog" IsChecked="True" Content="Todo"/>
                    <RadioButton x:Name="AdmLogTeam" Style="{StaticResource Seg}" GroupName="admlog" Content="Equipo"/>
                    <RadioButton x:Name="AdmLogAdmin" Style="{StaticResource Seg}" GroupName="admlog" Content="Admin"/>
                    <RadioButton x:Name="AdmLogErr" Style="{StaticResource Seg}" GroupName="admlog" Content="Errores"/>
                  </StackPanel>
                </Border>
                <TextBlock x:Name="AdmLogTitle" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="16" FontWeight="SemiBold" Foreground="#1E1A1B" VerticalAlignment="Center"/>
              </DockPanel>
              <ScrollViewer VerticalScrollBarVisibility="Auto">
                <ItemsControl x:Name="AdmLog">
                  <ItemsControl.ItemTemplate>
                    <DataTemplate>
                      <Border BorderBrush="#F1EDED" BorderThickness="0,0,0,1" Padding="4,9">
                        <Grid>
                          <Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="120"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                          <TextBlock Text="{Binding Time}" Foreground="#8A7F81" FontSize="12.5" VerticalAlignment="Center"/>
                          <Border Grid.Column="1" Background="{Binding UserBg}" CornerRadius="10" Padding="9,2" HorizontalAlignment="Left" VerticalAlignment="Center">
                            <TextBlock Text="{Binding User}" FontSize="12" FontWeight="SemiBold" Foreground="{Binding UserFg}"/>
                          </Border>
                          <TextBlock Grid.Column="2" Text="{Binding Text}" Foreground="{Binding Fg}" FontSize="13" TextWrapping="Wrap" VerticalAlignment="Center"/>
                        </Grid>
                      </Border>
                    </DataTemplate>
                  </ItemsControl.ItemTemplate>
                </ItemsControl>
              </ScrollViewer>
            </DockPanel>
          </Border>
          <!-- HERRAMIENTAS -->
          <ScrollViewer x:Name="AdmPageTools" Visibility="Collapsed" VerticalScrollBarVisibility="Auto">
            <StackPanel>
              <Border Style="{StaticResource Card}" Padding="26,22" Margin="0,0,0,14">
                <StackPanel>
                  <TextBlock Text="Control remoto (GitHub)" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="16" FontWeight="SemiBold" Foreground="#1E1A1B"/>
                  <TextBlock x:Name="AdmRemoteTxt" Foreground="#4B4345" TextWrapping="Wrap" Margin="0,6,0,0" LineHeight="20"/>
                </StackPanel>
              </Border>
              <Border Style="{StaticResource Card}" Padding="26,22" Margin="0,0,0,14">
                <StackPanel>
                  <TextBlock Text="Herramientas" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="16" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="0,0,0,12"/>
                  <WrapPanel>
                    <Button x:Name="AdmUpdate" Style="{StaticResource Btn}" Content="Buscar actualización ahora" Margin="0,0,8,8"/>
                    <Button x:Name="AdmOpenData" Style="{StaticResource Btn}" Content="Abrir carpeta de datos" Margin="0,0,8,8"/>
                    <Button x:Name="AdmOpenApp" Style="{StaticResource Btn}" Content="Abrir carpeta del programa" Margin="0,0,8,8"/>
                    <Button x:Name="AdmShowAll" Style="{StaticResource Btn}" Content="Ver Ajustes completos" Margin="0,0,8,8"/>
                  </WrapPanel>
                  <TextBlock x:Name="AdmInfo" Foreground="#8A7F81" FontSize="12.5" TextWrapping="Wrap" Margin="0,6,0,0"/>
                </StackPanel>
              </Border>
              <Border Style="{StaticResource Card}" Padding="26,22">
                <StackPanel>
                  <TextBlock Text="Cambiar contraseña de administrador" FontFamily="Segoe UI Variable Display, Segoe UI" FontSize="16" FontWeight="SemiBold" Foreground="#1E1A1B" Margin="0,0,0,12"/>
                  <Grid>
                    <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition Width="10"/><ColumnDefinition/><ColumnDefinition Width="10"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                    <PasswordBox x:Name="AdmNewPw" Height="38" Padding="10,0" VerticalContentAlignment="Center" BorderBrush="#D9D4D5"/>
                    <PasswordBox x:Name="AdmNewPw2" Grid.Column="2" Height="38" Padding="10,0" VerticalContentAlignment="Center" BorderBrush="#D9D4D5"/>
                    <Button x:Name="AdmChangePw" Grid.Column="4" Style="{StaticResource Btn}" Content="Cambiar" Height="38"/>
                  </Grid>
                  <TextBlock x:Name="AdmPwMsg" FontWeight="SemiBold" Margin="0,8,0,0"/>
                </StackPanel>
              </Border>
            </StackPanel>
          </ScrollViewer>
        </Grid>
      </DockPanel>
    </Border>
  </Grid>
  </Grid>
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

# botón «Hoy» del calendario desplegable
function Connect-CalendarToday($dp) {
    try {
        $pop = $dp.Template.FindName('PART_Popup', $dp)
        if (-not $pop -or -not $pop.Child) { return }
        $cal = $pop.Child
        [void]$cal.ApplyTemplate()
        $b = $cal.Template.FindName('TodayBtn', $cal)
        if (-not $b) { return }
        $b.Content = L 'calToday'
        $b.Tag = $dp
        if ($b.Uid -ne 'w') {
            $b.Uid = 'w'
            $b.Add_Click({ $d = $this.Tag; $d.SelectedDate = [DateTime]::Today; $d.DisplayDate = [DateTime]::Today; $d.IsDropDownOpen = $false })
        }
    } catch { Write-Log "calendario: $($_.Exception.Message)" }
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
            $secText = New-Object Windows.Controls.TextBlock
            $secText.Text = (L ("sec" + $section))
            $secText.Style = $script:Win.FindResource('SectionTitle')
            $secMark = New-Object Windows.Controls.Border
            $secMark.Width = 4; $secMark.Height = 16; $secMark.CornerRadius = 2; $secMark.Margin = '0,0,10,0'
            $secMark.Background = Brush '#D8202A'; $secMark.VerticalAlignment = 'Center'
            $secTitle = New-Object Windows.Controls.StackPanel
            $secTitle.Orientation = 'Horizontal'; $secTitle.Margin = '0,18,0,12'
            [void]$secTitle.Children.Add($secMark); [void]$secTitle.Children.Add($secText)
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
                # calendario en el idioma del programa (todos con fecha dd/mm/aaaa y semana desde el lunes)
                $cal = @{ es = 'es-ES'; en = 'en-GB'; fr = 'fr-FR'; ar = 'ar-MA' }["$($script:Cfg.Language)"]
                if (-not $cal) { $cal = 'es-ES' }
                $c.Language = [Windows.Markup.XmlLanguage]::GetLanguage($cal)
                $c.Add_SelectedDateChanged({ On-FieldChanged $this })
                $c.Add_CalendarOpened({ Connect-CalendarToday $this })
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
                    elseif ($f.Key -eq 'paystate' -and $script:PayColors.ContainsKey($it)) { $rb.Background = Brush $script:PayColors[$it] }
                    else { $rb.Uid = 'plain' }
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
        'ComboEdit' { if ($f.KeepCase) { return "$($c.Text)".Trim() }; return "$($c.Text)".Trim().ToUpper() }
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
        elseif ($f.Role -eq 'user') { $v = $script:User }
        elseif ($f.DefaultToday) { $v = (Get-Date).Date }
        elseif ($f.Role -eq 'gprice' -and $null -ne $script:DefaultPrice) { $v = $script:DefaultPrice }
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
    $ui.CardTipo.Text = $(if ($v.tipo) { "$($v.tipo)" } else { "$($v.xF)" })
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
    $script:Busy = $true
    $conf = Get-EditConflict
    $script:Busy = $false
    if ($conf) {
        $ww = Format-WhoWhen $conf
        $ans = [Windows.MessageBox]::Show($script:Win, ((L 'conflictAsk') -f $ww[0], $ww[1]), $script:AppName, 'YesNo', 'Warning')
        if ("$ans" -ne 'Yes') { return }
    }
    $ui = $script:UI
    $ui.SaveBtn.IsEnabled = $false
    $script:Busy = $true
    $script:Win.Cursor = [Windows.Input.Cursors]::Wait
    Set-Status (L 'savingMsg')
    try {
        $row = $null; if ($script:EditRow) { $row = $script:EditRow.Row }
        $res = Save-ClientStore $vals $row
        $id = $res.Id
        $wasEdit = [bool]$script:EditRow
        $tTitle = $script:T.Title['es']
        Write-Log ("{0} cliente {1} · {2} · {3}" -f $(if ($wasEdit) { 'editó' } else { 'añadió' }), $id, "$($vals.name)", $tTitle)
        [void](Load-Data)
        Set-Status ($(if ($wasEdit) { (L 'savedEdit') -f $id } else { (L 'saved') -f $id }))
        if (-not $res.Synced) { Set-Status ((L 'savedNoExcel') -f $id) }
        Reset-Form
        if ($wasEdit) { $ui.NavClients.IsChecked = $true }
        if ($script:Cfg.OpenAfterSave) { Open-ExcelWindow }
    } catch {
        Write-Log "Save-FromForm: $($_.Exception.Message)"
        Show-Error ((L 'errGeneric') + [Environment]::NewLine + $_.Exception.Message)
        Set-Status ''
    } finally {
        $ui.SaveBtn.IsEnabled = $true
        $script:Busy = $false
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
        if ($ui.OnlyMine.IsChecked -and $script:User -and "$($r.Resp)" -ine $script:User) { continue }
        $nd = Get-DocsCount $r
        $r | Add-Member -NotePropertyName DocsN -NotePropertyValue $nd -Force
        $r | Add-Member -NotePropertyName DocsVis -NotePropertyValue $(if ($nd -gt 0) { 'Visible' } else { 'Collapsed' }) -Force
        $r
    }
    $ui.ClientsGrid.ItemsSource = @($list)
}

function Edit-Selected {
    try {
        $sel = $script:UI.ClientsGrid.SelectedItem
        if (-not $sel) { Show-Info (L 'selectFirst'); return }
        if ([int]$sel.Row -ge $script:HandRowBase) { Show-Info ((L 'handReadOnly') -f $script:HandTabs[$script:T.Id]); return }
        $script:UI.NavNew.IsChecked = $true
        Fill-FormForEdit $sel
    } catch {
        Write-Log "DEBUG editar: $($_.Exception.Message) (línea $($_.InvocationInfo.ScriptLineNumber))"
        Show-Error $_.Exception.Message
    }
}

# =====================================================================
#  DOCUMENTOS, RECIBOS PDF Y EXPORTAR  · © Yano
# =====================================================================
$script:PdfPrinter = 'Microsoft Print to PDF'
$script:DocsMap = @{}
$script:DocsRow = $null

function Get-SafeName([string]$s) { return (("$s" -replace '[\\/:*?"<>|]', ' ') -replace '\s+', ' ').Trim().TrimEnd('.') }
function Get-DocsRoot {
    $x = "$($script:Cfg.ExcelPath)"
    if (-not $x) { return $null }
    return (Join-Path (Split-Path -Parent $x) 'GESTEM documentos')
}
function Get-TramiteDocsDir { $r = Get-DocsRoot; if (-not $r) { return $null }; return (Join-Path $r (Get-SafeName $script:T.Sheet)) }
function Get-ClientDocsDir($r, [switch]$Create) {
    $tdir = Get-TramiteDocsDir
    if (-not $tdir) { return $null }
    $id = Get-SafeName "$($r.Id)"
    if (-not $id) { $id = "fila $($r.Row)" }
    if (Test-Path -LiteralPath $tdir) {
        $hit = Get-ChildItem -LiteralPath $tdir -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $id -or $_.Name.StartsWith("$id - ") } | Select-Object -First 1
        if ($hit) { return $hit.FullName }
    }
    $p = Join-Path $tdir ("$id - " + (Get-SafeName "$($r.Name)"))
    if ($Create) { New-Item -ItemType Directory -Force -Path $p | Out-Null }
    return $p
}
function Update-DocsMap {
    # nº de documentos por expediente del trámite actual (para la columna de la lista)
    $script:DocsMap = @{}
    try {
        $tdir = Get-TramiteDocsDir
        if (-not $tdir -or -not (Test-Path -LiteralPath $tdir)) { return }
        foreach ($d in (Get-ChildItem -LiteralPath $tdir -Directory -ErrorAction SilentlyContinue)) {
            $id = $d.Name; $i = $id.IndexOf(' - '); if ($i -gt 0) { $id = $id.Substring(0, $i) }
            $script:DocsMap[$id] = @(Get-ChildItem -LiteralPath $d.FullName -File -ErrorAction SilentlyContinue).Count
        }
    } catch { }
}
function Get-DocsCount($r) {
    $id = Get-SafeName "$($r.Id)"; if (-not $id) { $id = "fila $($r.Row)" }
    if ($script:DocsMap.ContainsKey($id)) { return [int]$script:DocsMap[$id] }
    return 0
}
function Get-SelectedClient {
    $sel = $script:UI.ClientsGrid.SelectedItem
    if (-not $sel) { Show-Info (L 'selectFirst') }
    return $sel
}

# ---------- panel de documentos ----------
function Open-DocsPanel($r) {
    if (-not $r) { return }
    if (-not (Get-DocsRoot)) { Show-Error (L 'noFile'); return }
    $script:DocsRow = $r
    $ui = $script:UI
    $ui.DocsTitle.Text = L 'docsTitle'
    $ui.DocsSub.Text = ((L 'docsSub') -f $r.Id, $r.Name)
    $ui.DocsMsg.Text = ''
    Fill-DocsList
    $ui.PanelDocs.Visibility = 'Visible'
}
function Fill-DocsList {
    $ui = $script:UI
    $items = New-Object System.Collections.Generic.List[object]
    $dir = Get-ClientDocsDir $script:DocsRow
    if ($dir -and (Test-Path -LiteralPath $dir)) {
        foreach ($f in (Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)) {
            $ext = $f.Extension.ToLower()
            $glyph = [string][char]0xE8A5
            if ($ext -in '.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tif', '.tiff', '.heic', '.webp') { $glyph = [string][char]0xEB9F }
            elseif ($ext -eq '.pdf') { $glyph = [string][char]0xEA90 }
            $kb = [math]::Max([double]1, [math]::Round($f.Length / 1KB))
            $size = $(if ($kb -ge 1024) { ([double]$kb / 1024).ToString('0.0', $script:ES) + ' MB' } else { "$kb KB" })
            $items.Add([pscustomobject]@{ Name = $f.Name; Info = "$size  ·  $($f.LastWriteTime.ToString('dd/MM/yyyy HH:mm'))"; Glyph = $glyph; Path = $f.FullName })
        }
    }
    $ui.DocsList.ItemsSource = $items
    $ui.DocsEmpty.Visibility = $(if ($items.Count) { 'Collapsed' } else { 'Visible' })
    $ui.DocsCount.Text = $(if ($items.Count) { (L 'docsCount') -f $items.Count } else { '' })
}
function Add-ClientDocs($paths) {
    $r = $script:DocsRow
    if (-not $r) { return }
    $files = @($paths | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) })
    if (-not $files.Count) { return }
    try {
        $dir = Get-ClientDocsDir $r -Create
        $n = 0
        foreach ($p in $files) {
            $name = [IO.Path]::GetFileNameWithoutExtension($p); $ext = [IO.Path]::GetExtension($p)
            $dst = Join-Path $dir ($name + $ext); $k = 2
            while (Test-Path -LiteralPath $dst) { $dst = Join-Path $dir ("$name ($k)$ext"); $k++ }
            Copy-Item -LiteralPath $p -Destination $dst
            $n++
        }
        Write-Log ("añadió {0} documento(s) al expediente {1}" -f $n, $r.Id)
        $script:UI.DocsMsg.Text = ((L 'docsAdded') -f $n)
    } catch {
        Show-Error ((L 'errGeneric') + [Environment]::NewLine + $_.Exception.Message)
    }
    Fill-DocsList
    Update-DocsMap; Refresh-ClientGrid
}
function Add-DocsFromDialog {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Multiselect = $true
    $dlg.Title = L 'docsAdd'
    $dlg.Filter = (L 'docsFilter') + '|*.pdf;*.jpg;*.jpeg;*.png;*.tif;*.tiff;*.heic;*.webp;*.doc;*.docx;*.xls;*.xlsx|' + (L 'docsAll') + '|*.*'
    if ($dlg.ShowDialog($script:Win)) { Add-ClientDocs $dlg.FileNames }
}
function Open-SelectedDoc {
    $it = $script:UI.DocsList.SelectedItem
    if (-not $it) { return }
    try { Start-Process -FilePath $it.Path } catch { Show-Error $_.Exception.Message }
}
function Remove-SelectedDoc {
    $it = $script:UI.DocsList.SelectedItem
    if (-not $it) { Show-Info (L 'docsPick'); return }
    $ans = [Windows.MessageBox]::Show($script:Win, ((L 'docsRemoveAsk') -f $it.Name), $script:AppName, 'YesNo', 'Question')
    if ("$ans" -ne 'Yes') { return }
    try {
        Add-Type -AssemblyName Microsoft.VisualBasic
        [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($it.Path, 'OnlyErrorDialogs', 'SendToRecycleBin')
        Write-Log ("quitó un documento del expediente {0}" -f $script:DocsRow.Id)
    } catch { Show-Error $_.Exception.Message }
    Fill-DocsList
    Update-DocsMap; Refresh-ClientGrid
}
function Open-DocsFolder {
    $dir = Get-ClientDocsDir $script:DocsRow -Create
    Start-Process explorer.exe -ArgumentList "`"$dir`""
}
function Sync-DocsBackup([string]$dir) {
    # copia de los documentos escaneados en la carpeta externa (solo lo nuevo o cambiado)
    $root = Get-DocsRoot
    if (-not $root -or -not (Test-Path -LiteralPath $root)) { return }
    $dstRoot = Join-Path $dir 'documentos'
    foreach ($f in (Get-ChildItem -LiteralPath $root -File -Recurse -ErrorAction SilentlyContinue)) {
        try {
            $rel = $f.FullName.Substring($root.Length).TrimStart('\')
            $dst = Join-Path $dstRoot $rel
            $di = Split-Path -Parent $dst
            if (-not (Test-Path -LiteralPath $di)) { New-Item -ItemType Directory -Force -Path $di | Out-Null }
            if (Test-Path -LiteralPath $dst) {
                $o = Get-Item -LiteralPath $dst
                if ($o.Length -eq $f.Length -and $o.LastWriteTimeUtc -ge $f.LastWriteTimeUtc) { continue }
            }
            Copy-Item -LiteralPath $f.FullName -Destination $dst -Force
        } catch { }
    }
}

# ---------- PDF (impresora «Microsoft Print to PDF» de Windows, sin diálogos) ----------
function Test-PdfPrinter {
    try { return (@([Drawing.Printing.PrinterSettings]::InstalledPrinters) -contains $script:PdfPrinter) } catch { return $false }
}
function New-PdfFont([double]$size, [string]$weight = 'Regular') {
    $fam = 'Segoe UI'; $st = [Drawing.FontStyle]::Regular
    if ($weight -eq 'Semibold') { $fam = 'Segoe UI Semibold' }
    elseif ($weight -eq 'Bold') { $st = [Drawing.FontStyle]::Bold }
    elseif ($weight -eq 'Light') { $fam = 'Segoe UI Light' }
    return (New-Object Drawing.Font($fam, [single]$size, $st, [Drawing.GraphicsUnit]::Point))
}
function PdfColor([string]$hex) { return [Drawing.ColorTranslator]::FromHtml($hex) }
function PdfText($g, $text, [double]$x, [double]$y, [double]$w, [double]$size, [string]$weight = 'Regular', [string]$color = '#1E1A1B', [string]$align = 'Near') {
    $font = New-PdfFont $size $weight
    $br = New-Object Drawing.SolidBrush((PdfColor $color))
    $sf = New-Object Drawing.StringFormat
    $sf.Alignment = [Drawing.StringAlignment]$align
    $sf.Trimming = [Drawing.StringTrimming]::EllipsisCharacter
    $sf.FormatFlags = [Drawing.StringFormatFlags]::NoWrap
    $h = $size * 0.55 + 1.5
    $g.DrawString("$text", $font, $br, (New-Object Drawing.RectangleF([single]$x, [single]$y, [single]$w, [single]$h)), $sf)
    $sf.Dispose(); $br.Dispose(); $font.Dispose()
}
function PdfWrap($g, $text, [double]$x, [double]$y, [double]$w, [double]$h, [double]$size, [string]$color = '#6E6466') {
    $font = New-PdfFont $size
    $br = New-Object Drawing.SolidBrush((PdfColor $color))
    $g.DrawString("$text", $font, $br, (New-Object Drawing.RectangleF([single]$x, [single]$y, [single]$w, [single]$h)))
    $br.Dispose(); $font.Dispose()
}
function PdfRect($g, [double]$x, [double]$y, [double]$w, [double]$h, [string]$fill, [double]$radius = 0) {
    $br = New-Object Drawing.SolidBrush((PdfColor $fill))
    if ($radius -le 0) { $g.FillRectangle($br, [single]$x, [single]$y, [single]$w, [single]$h) }
    else {
        $d = [single]($radius * 2)
        $gp = New-Object Drawing.Drawing2D.GraphicsPath
        $gp.AddArc([single]$x, [single]$y, $d, $d, 180, 90)
        $gp.AddArc([single]($x + $w - $d), [single]$y, $d, $d, 270, 90)
        $gp.AddArc([single]($x + $w - $d), [single]($y + $h - $d), $d, $d, 0, 90)
        $gp.AddArc([single]$x, [single]($y + $h - $d), $d, $d, 90, 90)
        $gp.CloseFigure()
        $g.FillPath($br, $gp); $gp.Dispose()
    }
    $br.Dispose()
}
function PdfLine($g, [double]$x1, [double]$y1, [double]$x2, [double]$y2, [string]$color = '#E6E1E2', [double]$width = 0.3) {
    $pen = New-Object Drawing.Pen((PdfColor $color), [single]$width)
    $g.DrawLine($pen, [single]$x1, [single]$y1, [single]$x2, [single]$y2)
    $pen.Dispose()
}
function PdfLogo($g, [double]$x, [double]$y, [double]$size) {
    $p = "$($script:Cfg.CompanyLogo)"
    if (-not $p -or -not (Test-Path -LiteralPath $p)) { return $false }
    try {
        $img = [Drawing.Image]::FromFile($p)
        $g.DrawImage($img, (New-Object Drawing.RectangleF([single]$x, [single]$y, [single]$size, [single]$size)))
        $img.Dispose()
        return $true
    } catch { return $false }
}
function Invoke-PdfPrint([string]$path, [bool]$landscape, [scriptblock]$draw) {
    # $draw recibe ($g, nº de página) y devuelve $true si hay más páginas; los datos van en $script:PdfData
    $doc = New-Object Drawing.Printing.PrintDocument
    $doc.PrinterSettings.PrinterName = $script:PdfPrinter
    $doc.PrinterSettings.PrintToFile = $true
    $doc.PrinterSettings.PrintFileName = $path
    $doc.DocumentName = [IO.Path]::GetFileNameWithoutExtension($path)
    $doc.PrintController = New-Object Drawing.Printing.StandardPrintController
    $a4 = $doc.PrinterSettings.PaperSizes | Where-Object { $_.Kind -eq [Drawing.Printing.PaperKind]::A4 } | Select-Object -First 1
    if ($a4) { $doc.DefaultPageSettings.PaperSize = $a4 }
    $doc.DefaultPageSettings.Landscape = $landscape
    $doc.DefaultPageSettings.Margins = New-Object Drawing.Printing.Margins(0, 0, 0, 0)
    $script:PdfPage = 0
    $script:PdfDraw = $draw
    $script:PdfErr = $null
    $doc.Add_PrintPage({
        param($s, $e)
        try {
            $g = $e.Graphics
            $g.PageUnit = [Drawing.GraphicsUnit]::Millimeter
            $g.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $g.TextRenderingHint = [Drawing.Text.TextRenderingHint]::AntiAliasGridFit
            $g.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $g.TranslateTransform([single](-$e.PageSettings.HardMarginX * 0.254), [single](-$e.PageSettings.HardMarginY * 0.254))
            $more = & $script:PdfDraw $g $script:PdfPage
            $script:PdfPage++
            $e.HasMorePages = [bool]$more
        } catch { $script:PdfErr = $_.Exception.Message; $e.HasMorePages = $false }
    })
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    try { $doc.Print() } finally { $doc.Dispose() }
    if ($script:PdfErr) { throw $script:PdfErr }
    # la impresora escribe el archivo en segundo plano: esperar a que exista
    for ($i = 0; $i -lt 60 -and -not (Test-Path -LiteralPath $path); $i++) { Start-Sleep -Milliseconds 150 }
    if (-not (Test-Path -LiteralPath $path)) { throw (L 'pdfFail') }
    # esperar a que el tamaño deje de cambiar
    $last = -1
    for ($i = 0; $i -lt 40; $i++) { $len = (Get-Item -LiteralPath $path).Length; if ($len -gt 0 -and $len -eq $last) { break }; $last = $len; Start-Sleep -Milliseconds 200 }
}

# ---------- recibo ----------
function Draw-Receipt($g, $page) {
    $d = $script:PdfData
    $LE = 18; $RE = 192; $W = $RE - $LE
    # cabecera
    PdfRect $g 0 0 210 4 '#9E1820'
    $hasLogo = PdfLogo $g $LE 14 20
    $tx = $(if ($hasLogo) { $LE + 25 } else { $LE })
    PdfText $g $d.Company $tx 16.5 95 15 'Semibold' '#1E1A1B'
    PdfText $g (L 'rcSubtitle') $tx 25 95 9 'Regular' '#8A7F81'
    PdfText $g (L 'rcTitle') 110 14 $($RE - 110) 24 'Light' '#9E1820' 'Far'
    PdfText $g ((L 'rcNo') -f $d.No) 110 27 $($RE - 110) 10 'Semibold' '#1E1A1B' 'Far'
    PdfText $g ((L 'rcDate') -f $d.Date) 110 32.5 $($RE - 110) 9 'Regular' '#6E6466' 'Far'
    PdfLine $g $LE 44 $RE 44 '#E6E1E2' 0.35
    # cliente y trámite
    $y = 52
    PdfText $g (L 'rcClient').ToUpper() $LE $y 80 7.5 'Semibold' '#9E1820'
    PdfText $g (L 'rcProc').ToUpper() 110 $y 80 7.5 'Semibold' '#9E1820'
    PdfText $g $d.Name $LE ($y + 5) 88 13 'Semibold' '#1E1A1B'
    PdfText $g $d.Tramite 110 ($y + 5) 82 13 'Semibold' '#1E1A1B'
    $ly = $y + 14
    $left = @(@((L 'rcDoc'), $d.Doc), @((L 'rcCase'), $d.Id), @((L 'rcResp'), $d.Resp))
    $right = @(@((L 'rcType'), $d.Tipo), @((L 'rcApplied'), $d.Fsol), @((L 'rcStatusP'), $d.Status))
    for ($i = 0; $i -lt 3; $i++) {
        $yy = $ly + $i * 7
        PdfText $g $left[$i][0] $LE $yy 30 8.5 'Regular' '#8A7F81'
        PdfText $g $(if ("$($left[$i][1])".Trim()) { $left[$i][1] } else { '—' }) ($LE + 30) $yy 58 9 'Semibold' '#2B2526'
        PdfText $g $right[$i][0] 110 $yy 32 8.5 'Regular' '#8A7F81'
        PdfText $g $(if ("$($right[$i][1])".Trim()) { $right[$i][1] } else { '—' }) 142 $yy 50 9 'Semibold' '#2B2526'
    }
    # importes
    $by = 100
    PdfRect $g $LE $by $W 52 '#F7F5F5' 3
    PdfText $g (L 'rcPrice') ($LE + 8) ($by + 8) 100 10 'Regular' '#4B4345'
    PdfText $g (Format-Money $d.Price) 110 ($by + 8) ($RE - 118) 10 'Semibold' '#1E1A1B' 'Far'
    PdfText $g (L 'rcPaid') ($LE + 8) ($by + 17) 100 10 'Regular' '#4B4345'
    PdfText $g (Format-Money $d.Paid) 110 ($by + 17) ($RE - 118) 10 'Semibold' '#1F7A45' 'Far'
    PdfLine $g ($LE + 8) ($by + 27) ($RE - 8) ($by + 27) '#DDD6D7' 0.3
    PdfText $g (L 'rcPend') ($LE + 8) ($by + 32) 100 12 'Semibold' '#1E1A1B'
    PdfText $g (Format-Money $d.Pend) 100 ($by + 31) ($RE - 108) 16 'Semibold' $(if ($d.Pend -gt 0) { '#B42318' } else { '#1F7A45' }) 'Far'
    if ("$($d.Forma)".Trim()) { PdfText $g ((L 'rcMethod') + ': ' + $d.Forma) ($LE + 8) ($by + 43) 120 8.5 'Regular' '#6E6466' }
    # sello de estado
    $stamp = $(if ($d.Pend -le 0 -and $d.Price -gt 0) { (L 'rcPaidFull') } elseif ($d.Paid -gt 0) { (L 'rcPartial') } else { '' })
    if ($stamp) {
        $sc = $(if ($d.Pend -le 0) { '#1F7A45' } else { '#B7791F' })
        PdfRect $g ($RE - 52) 160 52 10 $(if ($d.Pend -le 0) { '#E7F4EC' } else { '#FDF3E1' }) 5
        PdfText $g $stamp.ToUpper() ($RE - 52) 162.3 52 9 'Semibold' $sc 'Center'
    }
    # firma
    PdfLine $g $LE 215 ($LE + 70) 215 '#B9B0B2' 0.3
    PdfText $g (L 'rcSign') $LE 217 70 8 'Regular' '#8A7F81'
    PdfWrap $g (L 'rcNote') $LE 232 $W 12 8 '#8A7F81'
    # pie
    PdfLine $g $LE 280 $RE 280 '#EEE9EA' 0.3
    PdfText $g $d.Company $LE 282.5 100 7.5 'Regular' '#A69C9E'
    PdfText $g (L 'rcFooter') 100 282.5 ($RE - 100) 7.5 'Regular' '#A69C9E' 'Far'
    return $false
}
function New-ClientReceipt($r) {
    if (-not $r) { return }
    if (-not (Test-PdfPrinter)) { Show-Error (L 'noPdfPrinter'); return }
    $script:Win.Cursor = [Windows.Input.Cursors]::Wait
    try {
        $dir = Get-ClientDocsDir $r -Create
        $n = @(Get-ChildItem -LiteralPath $dir -Filter 'Recibo *.pdf' -ErrorAction SilentlyContinue).Count + 1
        $id = "$($r.Id)".Trim(); if (-not $id) { $id = "F$($r.Row)" }
        $no = "$id-$n"
        $price = [double]$r.PriceV; $pend = [double]$r.Pending
        $paid = [math]::Max([double]0, $price - $pend)
        $title = $script:T.Title[$script:Cfg.Language]; if (-not $title) { $title = $script:T.Title['es'] }
        $fs = $r.Data['fsol']
        $script:PdfData = @{
            Company = $(if ("$($script:Cfg.CompanyName)".Trim()) { "$($script:Cfg.CompanyName)" } else { 'GESTEM MANAGER' })
            No = $no; Date = (Get-Date).ToString('dd/MM/yyyy'); Name = "$($r.Name)"; Doc = "$($r.Doc)"; Id = $id
            Tramite = $title; Tipo = "$($r.Tipo)"; Status = "$($r.Status)"; Resp = "$($r.Resp)"
            Fsol = $(if ($fs -is [DateTime]) { $fs.ToString('dd/MM/yyyy') } else { '' })
            Price = $price; Paid = $paid; Pend = $pend; Forma = "$($r.Data['forma'])"
        }
        $file = Join-Path $dir ("Recibo {0} {1:yyyy-MM-dd}.pdf" -f (Get-SafeName $no), (Get-Date))
        Invoke-PdfPrint $file $false ${function:Draw-Receipt}
        Write-Log ("creó el recibo {0}" -f $no)
        Update-DocsMap; Refresh-ClientGrid
        if (Test-Path -LiteralPath $file) { Start-Process -FilePath $file }
        Set-Status ((L 'rcDone') -f $no)
    } catch {
        Write-Log "recibo: $($_.Exception.Message)"
        Show-Error ((L 'errGeneric') + [Environment]::NewLine + $_.Exception.Message)
    } finally { $script:Win.Cursor = $null }
}

# ---------- exportar la lista ----------
function Get-ExportFile([string]$ext, [string]$filter) {
    $title = $script:T.Title[$script:Cfg.Language]; if (-not $title) { $title = $script:T.Title['es'] }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.FileName = Get-SafeName ("{0} {1} {2:yyyy-MM-dd}" -f (L 'navClients'), $title, (Get-Date))
    $dlg.DefaultExt = $ext
    $xd = Split-Path -Parent "$($script:Cfg.ExcelPath)"
    if ($xd -and (Test-Path -LiteralPath $xd)) { $dlg.InitialDirectory = $xd }
    $dlg.Filter = $filter
    if ($dlg.ShowDialog($script:Win)) { return $dlg.FileName }
    return $null
}
function Export-ClientsCsv {
    $rows = @($script:UI.ClientsGrid.ItemsSource)
    if (-not $rows.Count) { Show-Info (L 'exEmpty'); return }
    $file = Get-ExportFile '.csv' 'CSV (Excel)|*.csv'
    if (-not $file) { return }
    try {
        $q = { param($v) $s = "$v"; if ($s -match '[;"\r\n]') { '"' + $s.Replace('"', '""') + '"' } else { $s } }
        $fields = @($script:T.Fields)
        $head = @((L 'colId')) + @($fields | ForEach-Object { Get-FieldLabel $_ }) + @((L 'colPend'))
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add((($head | ForEach-Object { & $q $_ }) -join ';'))
        foreach ($r in $rows) {
            $cells = @(& $q $r.Id)
            foreach ($f in $fields) {
                $v = $r.Data[$f.Key]
                if ($v -is [DateTime]) { $v = $v.ToString('dd/MM/yyyy') }
                elseif ($v -is [double]) { $v = $v.ToString('0.##', $script:ES) }
                $cells += (& $q $v)
            }
            $cells += (& $q ([double]$r.Pending).ToString('0.00', $script:ES))
            $lines.Add(($cells -join ';'))
        }
        [IO.File]::WriteAllLines($file, [string[]]$lines, (New-Object Text.UTF8Encoding($true)))
        Write-Log ("exportó la lista a CSV ({0} clientes)" -f $rows.Count)
        Set-Status ((L 'exDone') -f [IO.Path]::GetFileName($file))
        Start-Process explorer.exe -ArgumentList "/select,`"$file`""
    } catch { Show-Error ((L 'errGeneric') + [Environment]::NewLine + $_.Exception.Message) }
}
function Draw-ClientList($g, $page) {
    $d = $script:PdfData
    $LE = 12; $RE = 285
    $cols = @(@('colId', 27), @('colName', 60), @('colDoc', 27), @('colTipo', 33), @('colStatus', 40), @('colPay', 22), @('colPend', 24), @('colDays', 12), @('colResp', 28))
    # cabecera
    PdfRect $g 0 0 297 3 '#9E1820'
    $hasLogo = PdfLogo $g $LE 9 13
    $tx = $(if ($hasLogo) { $LE + 16 } else { $LE })
    PdfText $g $d.Company $tx 9.5 150 12 'Semibold' '#1E1A1B'
    PdfText $g $d.Title $tx 16 170 9 'Regular' '#6E6466'
    PdfText $g $d.Sub 150 10.5 ($RE - 150) 8 'Regular' '#8A7F81' 'Far'
    PdfText $g $d.Sub2 150 15.5 ($RE - 150) 8 'Regular' '#8A7F81' 'Far'
    # cabecera de tabla
    $y = 27
    PdfRect $g $LE $y ($RE - $LE) 7 '#F3EFEF' 1.5
    $x = $LE + 2
    foreach ($c in $cols) {
        $al = $(if ($c[0] -in 'colPend', 'colDays') { 'Far' } else { 'Near' })
        PdfText $g (L $c[0]) $x ($y + 1.6) ($c[1] - 3) 7.5 'Semibold' '#6E6466' $al
        $x += $c[1]
    }
    $y += 8.5
    $start = $page * $d.PerPage
    $end = [math]::Min($d.Rows.Count, $start + $d.PerPage)
    for ($i = $start; $i -lt $end; $i++) {
        $r = $d.Rows[$i]
        $x = $LE + 2
        $vals = @("$($r.Id)", "$($r.Name)", "$($r.Doc)", "$($r.Tipo)", "$($r.Status)", "$($r.Pay)", "$($r.PendingText)", "$($r.Days)", "$($r.Resp)")
        for ($c = 0; $c -lt $cols.Count; $c++) {
            $w = $cols[$c][1]
            if ($c -eq 4 -and "$($r.Color)" -and "$($r.Color)" -ne '#FFFFFF') {
                PdfRect $g $x ($y + 1.6) 2.6 2.6 "$($r.Color)" 1.3
                PdfText $g $vals[$c] ($x + 4) ($y + 0.8) ($w - 7) 8 'Regular' '#2B2526'
            } else {
                $al = $(if ($c -in 6, 7) { 'Far' } else { 'Near' })
                $col = $(if ($c -eq 6) { '#B42318' } else { '#2B2526' })
                $wt = $(if ($c -in 1, 6) { 'Semibold' } else { 'Regular' })
                PdfText $g $vals[$c] $x ($y + 0.8) ($w - 3) 8 $wt $col $al
            }
            $x += $w
        }
        PdfLine $g $LE ($y + 6.2) $RE ($y + 6.2) '#F1EDED' 0.25
        $y += 6.4
    }
    $last = ($end -ge $d.Rows.Count)
    if ($last) {
        $y += 3
        PdfRect $g $LE $y ($RE - $LE) 9 '#F7F5F5' 1.5
        PdfText $g $d.Total ($LE + 3) ($y + 2.2) 200 9 'Semibold' '#1E1A1B'
        PdfText $g $d.TotalPend 180 ($y + 2.2) ($RE - 183) 9 'Semibold' '#B42318' 'Far'
    }
    PdfText $g ((L 'exPage') -f ($page + 1), $d.Pages) $LE 201 100 7.5 'Regular' '#A69C9E'
    PdfText $g 'GESTEM MANAGER' 180 201 ($RE - 180) 7.5 'Regular' '#A69C9E' 'Far'
    return (-not $last)
}
function Export-ClientsPdf {
    $rows = @($script:UI.ClientsGrid.ItemsSource)
    if (-not $rows.Count) { Show-Info (L 'exEmpty'); return }
    if (-not (Test-PdfPrinter)) { Show-Error (L 'noPdfPrinter'); return }
    $file = Get-ExportFile '.pdf' 'PDF|*.pdf'
    if (-not $file) { return }
    $script:Win.Cursor = [Windows.Input.Cursors]::Wait
    try {
        $title = $script:T.Title[$script:Cfg.Language]; if (-not $title) { $title = $script:T.Title['es'] }
        $per = 25
        $sumPend = 0.0; $nPend = 0
        foreach ($r in $rows) { if ([double]$r.Pending -gt 0) { $sumPend += [double]$r.Pending; $nPend++ } }
        $filt = @()
        if ("$($script:UI.SearchBox.Text)".Trim()) { $filt += '«' + "$($script:UI.SearchBox.Text)".Trim() + '»' }
        if ($script:UI.FilterBox.SelectedIndex -gt 0) { $filt += "$($script:UI.FilterBox.SelectedItem)" }
        if ($script:UI.OnlyDebt.IsChecked) { $filt += (L 'onlyDebt') }
        if ($script:UI.OnlyMine.IsChecked) { $filt += (L 'onlyMine') }
        $script:PdfData = @{
            Company = $(if ("$($script:Cfg.CompanyName)".Trim()) { "$($script:Cfg.CompanyName)" } else { 'GESTEM MANAGER' })
            Title = ((L 'exTitle') -f $title)
            Sub = (Get-Date).ToString('dd/MM/yyyy HH:mm') + $(if ($script:User) { "  ·  $($script:User)" } else { '' })
            Sub2 = $(if ($filt.Count) { (L 'exFilters') + ': ' + ($filt -join ', ') } else { (L 'exNoFilter') })
            Rows = $rows; PerPage = $per; Pages = [int][math]::Max([double]1, [math]::Ceiling($rows.Count / $per))
            Total = ((L 'exTotal') -f $rows.Count, $nPend); TotalPend = (L 'colPend') + ': ' + (Format-Money $sumPend)
        }
        Invoke-PdfPrint $file $true ${function:Draw-ClientList}
        Write-Log ("exportó la lista a PDF ({0} clientes)" -f $rows.Count)
        Set-Status ((L 'exDone') -f [IO.Path]::GetFileName($file))
        if (Test-Path -LiteralPath $file) { Start-Process -FilePath $file }
    } catch {
        Write-Log "exportar PDF: $($_.Exception.Message)"
        Show-Error ((L 'errGeneric') + [Environment]::NewLine + $_.Exception.Message)
    } finally { $script:Win.Cursor = $null }
}

# =====================================================================
#  IDIOMA Y NAVEGACIÓN
# =====================================================================
function Apply-Language {
    $ui = $script:UI
    $script:Win.FlowDirection = $(if ($script:Cfg.Language -eq 'ar') { 'RightToLeft' } else { 'LeftToRight' })
    $ui.NavNew.Content = L 'navNew'; $ui.NavClients.Content = L 'navClients'; $ui.NavStats.Content = L 'navStats'
    $ui.HelpBtn.Content = L 'btnHelp'; $ui.SetSub.Text = L 'setSub'
    $ui.ExtDirL.Text = L 'extDirL'; $ui.ExtDirBtn.Content = L 'extChooseBtn'; $ui.ExtClearBtn.Content = L 'btnLogoClear'
    $ui.ExtOneDriveBtn.Content = L 'extOneDrive'; $ui.ExtNowBtn.Content = L 'extNow'; $ui.ExtOpenBtn.Content = L 'extOpen'; $ui.ExtNote.Text = L 'extNote'
    foreach ($k in 'File', 'Ext', 'Comp', 'Team', 'Pref', 'Help') { $ui["Sec$($k)T"].Text = L "sec$($k)T"; $ui["Sec$($k)D"].Text = L "sec$($k)D" }
    $ui.ChartDay.Content = L 'chDay'; $ui.ChartWeek.Content = L 'chWeek'; $ui.ChartMonth.Content = L 'chMonth'
    $ui.ChartLegPaid.Text = L 'chPaid'; $ui.ChartLegPend.Text = L 'chPend'; $ui.ChartLegMoney.Text = L 'chMoney'
    $ui.StatsTitle.Text = L 'statsTitle'; $ui.StatsTable.Text = L 'statsTable'
    $ui.KpiBillL.Text = L 'kpiBill'; $ui.KpiPaidL.Text = L 'kpiPaid'; $ui.KpiPendL.Text = L 'kpiPend'; $ui.KpiNetL.Text = L 'kpiNet'
    Fill-StatsFilters
    $ui.NavSettings.Content = L 'navSettings'; $ui.NavAbout.Content = L 'navAbout'
    $ui.TramiteLabel.Text = L 'tramite'
    $ui.SumPriceL.Text = L 'sumPrice'; $ui.SumPaidL.Text = L 'sumPaid'; $ui.SumPendL.Text = L 'sumPend'
    $ui.SearchHint.Text = L 'search'; $ui.OnlyDebt.Content = L 'onlyDebt'
    $ui.RefreshBtn.Content = L 'btnRefresh'; $ui.OpenExcelBtn.Content = L 'btnOpenExcel'; $ui.EditBtn.Content = L 'btnEdit'
    $cols = $ui.ClientsGrid.Columns
    $heads = @('colId', 'colName', 'colDoc', 'colTipo', 'colStatus', 'colPay', 'colPend', 'colDays', 'colResp')
    $ui.OnlyMine.Content = L 'onlyMine'; $ui.SetCollL.Text = L 'setColleagues'
    $ui.SetCompanyL.Text = L 'setCompany'; $ui.SetLogoL.Text = L 'setLogo'; $ui.LogoBtn.Content = L 'btnLogo'; $ui.LogoClearBtn.Content = L 'btnLogoClear'
    $ui.UserBtn.Content = $(if ($script:User) { (L 'userTip') -f $script:User } else { L 'sessGo' })
    for ($i = 0; $i -lt $heads.Count; $i++) { $cols[$i].Header = L $heads[$i] }
    $ui.DocsCol.Header = L 'colDocs'
    $ui.DocsBtn.Content = L 'btnDocs'; $ui.ReceiptBtn.Content = L 'btnReceipt'; $ui.ExportBtn.Content = (L 'btnExport') + '  ' + [char]0x25BE
    $ui.ExpPdf.Header = L 'expPdf'; $ui.ExpCsv.Header = L 'expCsv'
    $ui.CmEdit.Header = L 'btnEdit'; $ui.CmReceipt.Header = L 'btnReceipt'; $ui.CmDocs.Header = L 'btnDocs'
    $ui.DocsAddBtn.Content = L 'docsAdd'; $ui.DocsOpenBtn.Content = L 'docsOpen'; $ui.DocsRemoveBtn.Content = L 'docsRemove'; $ui.DocsFolderBtn.Content = L 'docsFolder'
    $ui.DocsEmptyT.Text = L 'docsEmpty'; $ui.DocsDropT.Text = L 'docsDrop'; $ui.DocsClose.ToolTip = L 'docsClose'
    $ui.SetTitle.Text = L 'setTitle'; $ui.SetFileL.Text = L 'setFile'; $ui.ChangeFileBtn.Content = L 'btnChange'
    $ui.ImportXlBtn.Content = L 'btnImport'; $ui.DataNote.Text = L 'dataNote'
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
    Apply-AdminLanguage
    Build-Form
    Reset-Form
}

function Show-Panel([string]$name) {
    foreach ($p in @('PanelForm', 'PanelClients', 'PanelStats', 'PanelSettings', 'PanelAbout')) {
        $script:UI[$p].Visibility = $(if ($p -eq $name) { 'Visible' } else { 'Collapsed' })
    }
    # en Resumen el selector de arriba incluye «Todos los trámites»
    $onStats = ($name -eq 'PanelStats')
    # la barra de trámite solo tiene sentido en Nuevo cliente, Clientes y Resumen
    $script:UI.TopBar.Visibility = $(if ($name -in 'PanelForm', 'PanelClients', 'PanelStats') { 'Visible' } else { 'Collapsed' })
    $script:UI.TramiteBox.Visibility = $(if ($onStats) { 'Collapsed' } else { 'Visible' })
    $script:UI.StatsScope.Visibility = $(if ($onStats) { 'Visible' } else { 'Collapsed' })
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
    $ui.ColleaguesBox.Text = (@($script:Cfg.Colleagues) -join ', ')
    $ui.CompanyBox.Text = "$($script:Cfg.CompanyName)"
    $script:PendingLogo = "$($script:Cfg.CompanyLogo)"
    $ui.LogoPreview.Fill = (Get-LogoBrush $script:PendingLogo)
    $ui.ExtDirBox.Text = "$($script:Cfg.ExtBackupDir)"
    $od = "$env:OneDrive"; if (-not $od) { $od = "$env:OneDriveConsumer" }
    $ui.ExtOneDriveBtn.Visibility = $(if ($od -and (Test-Path -LiteralPath $od)) { 'Visible' } else { 'Collapsed' })
    Update-ExtBackupInfo
    $ui.BackupChk.IsChecked = [bool]$script:Cfg.Backup
    $ui.OpenAfterChk.IsChecked = [bool]$script:Cfg.OpenAfterSave
    $ui.WelcomeText.Visibility = $(if (-not $script:Cfg.ExcelPath) { 'Visible' } else { 'Collapsed' })
}

# =====================================================================
#  EMPRESA (nombre y logo: cada empresa pone los suyos en Ajustes)
# =====================================================================
function Load-ImageFile([string]$path) {
    if (-not $path -or -not (Test-Path -LiteralPath $path)) { return $null }
    try {
        $bmp = New-Object Windows.Media.Imaging.BitmapImage
        $bmp.BeginInit(); $bmp.CacheOption = 'OnLoad'; $bmp.CreateOptions = 'IgnoreImageCache'
        $bmp.UriSource = New-Object Uri($path); $bmp.DecodePixelWidth = 256; $bmp.EndInit(); $bmp.Freeze()
        return $bmp
    } catch { Write-Log "logo no válido: $($_.Exception.Message)"; return $null }
}
function Get-LogoBrush([string]$path) {
    $src = Load-ImageFile $path
    if (-not $src) { $src = $script:AppLogo }
    if (-not $src) { return $null }
    $b = New-Object Windows.Media.ImageBrush($src)
    $b.Stretch = 'Uniform'
    return $b
}
function Save-CompanyLogo([string]$path) {
    # copia la imagen elegida dentro de la carpeta de datos (así no depende del archivo original)
    if (-not $path) { return '' }
    if (-not (Test-Path -LiteralPath $path)) { return '' }
    $ext = [IO.Path]::GetExtension($path).ToLower()
    $dest = Join-Path $script:DataDir ("empresa" + $ext)
    if ([IO.Path]::GetFullPath($path) -ieq [IO.Path]::GetFullPath($dest)) { return $dest }
    try {
        Get-ChildItem -LiteralPath $script:DataDir -Filter 'empresa.*' -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in '.png','.jpg','.jpeg','.bmp','.gif' } | Remove-Item -Force -ErrorAction SilentlyContinue
        Copy-Item -LiteralPath $path -Destination $dest -Force
        return $dest
    } catch { Write-Log "logo: $($_.Exception.Message)"; return $path }
}
function Choose-Logo {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Title = L 'chooseLogo'
    $dlg.Filter = 'Imagen (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp'
    $dlg.InitialDirectory = [Environment]::GetFolderPath('MyPictures')
    if ($dlg.ShowDialog($script:Win)) {
        if (-not (Load-ImageFile $dlg.FileName)) { Show-Error (L 'chooseLogo'); return }
        $script:PendingLogo = $dlg.FileName
        $script:UI.LogoPreview.Fill = (Get-LogoBrush $script:PendingLogo)
    }
}
function Apply-Branding {
    $ui = $script:UI
    $name = "$($script:Cfg.CompanyName)".Trim()
    $brush = Get-LogoBrush "$($script:Cfg.CompanyLogo)"
    $ui.SideLogo.Fill = $brush
    $ui.SessLogo.Fill = $brush
    $ui.SideCompany.Text = $name.ToUpper()
    $ui.SideCompany.Visibility = $(if ($name) { 'Visible' } else { 'Collapsed' })
    $ui.SessCompany.Text = $(if ($name) { $name } else { 'GESTEM' })
    $ui.AboutCompany.Text = $name
    $script:Win.Title = $(if ($name) { "GESTEM MANAGER · $name" } else { 'GESTEM MANAGER' })
}
# accesos directos sin ventana: abren GESTEM con conhost --headless (nada de Terminal de Windows)
function Repair-Shortcuts {
    try {
        $ps = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $ch = Join-Path $env:WINDIR 'System32\conhost.exe'
        if (-not (Test-Path -LiteralPath $ch)) { return }
        $me = Join-Path $script:AppDir 'GESTEM.ps1'
        $want = '--headless "' + $ps + '" -NoProfile -ExecutionPolicy Bypass -STA -File "' + $me + '" -Oculto'
        $ico = Join-Path $script:AppDir 'gestem.ico'
        $wantIco = $(if (Test-Path -LiteralPath $ico) { "$ico,0" } else { '' })
        $sh = New-Object -ComObject WScript.Shell
        foreach ($d in @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('Programs'))) {
            $f = Join-Path $d 'GESTEM MANAGER.lnk'
            if (-not (Test-Path -LiteralPath $f)) { continue }
            $l = $sh.CreateShortcut($f)
            if ($l.TargetPath -like '*conhost.exe' -and $l.Arguments -eq $want -and (-not $wantIco -or $l.IconLocation -ieq $wantIco)) { continue }
            if ($l.TargetPath -notlike '*powershell.exe' -and $l.TargetPath -notlike '*conhost.exe') { continue }
            $l.TargetPath = $ch
            $l.Arguments = $want
            $l.WorkingDirectory = $script:AppDir
            $l.WindowStyle = 7
            if ($wantIco) { $l.IconLocation = $wantIco }    # icono de GESTEM, no el de PowerShell ni el de la consola
            $l.Save()
            Write-Log "acceso directo corregido: $f"
        }
    } catch { Write-Log "accesos directos: $($_.Exception.Message)" }
}
function Import-DefaultBranding {
    # primera vez: si el instalador trae empresa.png / empresa.txt junto al programa, se usan solos
    $changed = $false
    if (-not $script:Cfg.CompanyLogo) {
        foreach ($e in '.png', '.jpg', '.jpeg') {
            $f = Join-Path $script:AppDir ("empresa" + $e)
            if (Test-Path -LiteralPath $f) { $script:Cfg.CompanyLogo = Save-CompanyLogo $f; $changed = $true; break }
        }
    }
    if (-not $script:Cfg.CompanyName) {
        $t = Join-Path $script:AppDir 'empresa.txt'
        if (Test-Path -LiteralPath $t) { $script:Cfg.CompanyName = (Get-Content -LiteralPath $t -Encoding UTF8 -TotalCount 1).Trim(); $changed = $true }
    }
    if ($changed) { Save-Config }
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
    $script:Cfg.CompanyName = "$($ui.CompanyBox.Text)".Trim()
    $script:Cfg.CompanyLogo = Save-CompanyLogo $script:PendingLogo
    Apply-Branding
    $script:Cfg.Colleagues = @("$($ui.ColleaguesBox.Text)".Split(',') | ForEach-Object { $_.Trim().ToUpper() } | Where-Object { $_ })
    $script:Lists['colleagues'] = $script:Cfg.Colleagues
    $newExt = "$($ui.ExtDirBox.Text)".Trim()
    $extChanged = ($newExt -ne "$($script:Cfg.ExtBackupDir)")
    $script:Cfg.ExtBackupDir = $newExt
    if ($extChanged) { $script:Cfg.LastExtBackup = ''; if ($newExt) { [void](Backup-External) } }
    $script:Cfg.Backup = [bool]$ui.BackupChk.IsChecked
    $script:Cfg.OpenAfterSave = [bool]$ui.OpenAfterChk.IsChecked
    Save-Config
    if ($script:Cfg.ExcelPath -ne $oldPath) { [void](Load-Data -All) }
    Apply-Language
    Load-SettingsPanel
    Write-Log 'cambió los ajustes'
    Set-Status (L 'settingsSaved')
}

# =====================================================================
#  RESUMEN ECONÓMICO (todas las pestañas del Excel)
# =====================================================================
function Format-Euro0($v) { return ([double]$v).ToString('N0', $script:ES) + ' €' }
function Fill-StatsFilters {
    $ui = $script:UI
    $script:StatsFilling = $true
    $whoSel = $ui.StatsWho.SelectedIndex; $perSel = $ui.StatsPeriod.SelectedIndex
    $ui.StatsWho.Items.Clear()
    [void]$ui.StatsWho.Items.Add((L 'whoAll'))
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($c in @($script:Cfg.Colleagues)) { if ("$c".Trim() -and -not $names.Contains("$c")) { $names.Add("$c") } }
    if ($script:Cache) {
        foreach ($d in $script:Cache.Values) { foreach ($r in $d.Rows) { $x = "$($r.Resp)".Trim().ToUpper(); if ($x -and -not $names.Contains($x)) { $names.Add($x) } } }
    }
    foreach ($n in $names) { [void]$ui.StatsWho.Items.Add($n) }
    $ui.StatsScope.Items.Clear()
    [void]$ui.StatsScope.Items.Add((L 'scopeAll'))
    $si = 0; $k = 0
    foreach ($t in $script:Tramites) {
        $k++
        $tt = $t.Title[$script:Cfg.Language]; if (-not $tt) { $tt = $t.Title['es'] }
        [void]$ui.StatsScope.Items.Add($tt)
        if ($script:StatsScopeId -and $t.Id -eq $script:StatsScopeId) { $si = $k }
    }
    $ui.StatsScope.SelectedIndex = $si
    $ui.StatsPeriod.Items.Clear()
    foreach ($k in 'perAll', 'perMonth', 'perPrev', 'perYear') { [void]$ui.StatsPeriod.Items.Add((L $k)) }
    $ui.StatsWho.SelectedIndex = $(if ($whoSel -ge 0 -and $whoSel -lt $ui.StatsWho.Items.Count) { $whoSel } else { 0 })
    $ui.StatsPeriod.SelectedIndex = $(if ($perSel -ge 0) { $perSel } else { 0 })
    $script:StatsFilling = $false
}
function New-StatCell([string]$text, [int]$col, [string]$align = 'Right', [string]$color = '#2B2526', [bool]$bold = $false, [double]$size = 14.5) {
    $t = New-Object Windows.Controls.TextBlock
    $t.Text = $text; $t.FontSize = $size; $t.Foreground = Brush $color
    $t.VerticalAlignment = 'Center'; $t.TextAlignment = $align; $t.Margin = '8,0'
    if ($bold) { $t.FontWeight = 'SemiBold' }
    [Windows.Controls.Grid]::SetColumn($t, $col)
    return $t
}
function Draw-ChartLine {
    $cv = $script:UI.ChartCanvas
    $cv.Children.Clear()
    $bk = $script:ChartBuckets
    if (-not $bk -or $bk.Count -lt 2) { return }
    $W = [double]$cv.ActualWidth; $H = [double]$cv.ActualHeight
    if ($W -le 0 -or $H -le 0) { return }
    $n = $bk.Count
    $max = [double]1; foreach ($b in $bk) { if ($b.N -gt $max) { $max = [double]$b.N } }
    $top = [double]30; $bottom = [double]12; $plotH = $H - $top - $bottom
    $step = $W / $n
    # líneas guía
    foreach ($f in 0, 0.5, 1) {
        $y = $top + $plotH * $f
        $gl = New-Object Windows.Shapes.Line
        $gl.X1 = 0; $gl.X2 = $W; $gl.Y1 = $y; $gl.Y2 = $y
        $gl.Stroke = Brush $(if ($f -eq 1) { '#E4DFE0' } else { '#F3EFEF' }); $gl.StrokeThickness = 1
        if ($f -lt 1) { $gl.StrokeDashArray = New-Object Windows.Media.DoubleCollection(,[double[]]@(4, 4)) }
        [void]$cv.Children.Add($gl)
    }
    $pts = New-Object System.Collections.Generic.List[Windows.Point]
    for ($i = 0; $i -lt $n; $i++) {
        $x = $step * ($i + 0.5)
        $y = $top + $plotH * (1 - ([double]$bk[$i].N / $max))
        $pts.Add((New-Object Windows.Point($x, $y)))
    }
    # línea recta entre puntos y área sombreada debajo
    $base = $top + $plotH
    $af = New-Object Windows.Media.PathFigure
    $af.StartPoint = New-Object Windows.Point($pts[0].X, $base)
    $af.IsClosed = $true
    $af.Segments.Add((New-Object Windows.Media.LineSegment($pts[0], $false)))
    for ($i = 1; $i -lt $n; $i++) {
        $p0 = $pts[$i - 1]; $p1 = $pts[$i]; $dx = ($p1.X - $p0.X) / 2
        $af.Segments.Add((New-Object Windows.Media.LineSegment($p1, $false)))
    }
    $af.Segments.Add((New-Object Windows.Media.LineSegment((New-Object Windows.Point($pts[$n - 1].X, $base)), $false)))
    $ageo = New-Object Windows.Media.PathGeometry; $ageo.Figures.Add($af)
    $area = New-Object Windows.Shapes.Path; $area.Data = $ageo
    $ag = New-Object Windows.Media.LinearGradientBrush; $ag.StartPoint = New-Object Windows.Point(0, 0); $ag.EndPoint = New-Object Windows.Point(0, 1)
    $ag.GradientStops.Add((New-Object Windows.Media.GradientStop(([Windows.Media.ColorConverter]::ConvertFromString('#2A9E1820')), 0)))
    $ag.GradientStops.Add((New-Object Windows.Media.GradientStop(([Windows.Media.ColorConverter]::ConvertFromString('#009E1820')), 1)))
    $area.Fill = $ag
    [void]$cv.Children.Add($area)
    # periodo actual: línea vertical fina
    $vx = $pts[$n - 1].X
    $vl = New-Object Windows.Shapes.Line
    $vl.X1 = $vx; $vl.X2 = $vx; $vl.Y1 = $pts[$n - 1].Y; $vl.Y2 = $base
    $vl.Stroke = Brush '#559E1820'; $vl.StrokeThickness = 1; $vl.StrokeDashArray = New-Object Windows.Media.DoubleCollection(,[double[]]@(3, 3))
    [void]$cv.Children.Add($vl)
    # tramos: verde si sube, rojo si baja, gris si igual
    $colors = New-Object System.Collections.Generic.List[string]
    $colors.Add('#A69C9E')
    for ($i = 1; $i -lt $n; $i++) {
        $c = $(if ($bk[$i].N -gt $bk[$i - 1].N) { '#1F9D55' } elseif ($bk[$i].N -lt $bk[$i - 1].N) { '#D8202A' } else { '#B8B0B2' })
        $colors.Add($c)
        $p0 = $pts[$i - 1]; $p1 = $pts[$i]; $dx = ($p1.X - $p0.X) / 2
        $fig = New-Object Windows.Media.PathFigure; $fig.StartPoint = $p0
        $fig.Segments.Add((New-Object Windows.Media.LineSegment($p1, $true)))
        $geo = New-Object Windows.Media.PathGeometry; $geo.Figures.Add($fig)
        $seg = New-Object Windows.Shapes.Path; $seg.Data = $geo
        $seg.Stroke = Brush $c; $seg.StrokeThickness = 2.6
        $seg.StrokeStartLineCap = 'Round'; $seg.StrokeEndLineCap = 'Round'; $seg.StrokeLineJoin = 'Round'
        [void]$cv.Children.Add($seg)
    }
    # puntos y cifras
    for ($i = 0; $i -lt $n; $i++) {
        $isNow = ($i -eq $n - 1)
        $d = $(if ($isNow) { [double]13 } else { [double]8 })
        if ($isNow) {
            $halo = New-Object Windows.Shapes.Ellipse; $halo.Width = 26; $halo.Height = 26
            $halo.Fill = Brush $(if ($colors[$i] -eq '#1F9D55') { '#261F9D55' } elseif ($colors[$i] -eq '#D8202A') { '#26D8202A' } else { '#26A69C9E' })
            [Windows.Controls.Canvas]::SetLeft($halo, $pts[$i].X - 13); [Windows.Controls.Canvas]::SetTop($halo, $pts[$i].Y - 13)
            [void]$cv.Children.Add($halo)
        }
        $dot = New-Object Windows.Shapes.Ellipse
        $dot.Width = $d; $dot.Height = $d
        $dot.Fill = $(if ($isNow) { Brush $colors[$i] } else { Brush '#FFFFFF' })
        $dot.Stroke = $(if ($isNow) { Brush '#FFFFFF' } else { Brush $colors[$i] }); $dot.StrokeThickness = $(if ($isNow) { 2.5 } else { 2 })
        $dot.ToolTip = (L 'chTip') -f $bk[$i].Long, $bk[$i].N, (Format-Euro0 $bk[$i].Bill), (Format-Euro0 $bk[$i].Paid)
        $dot.Cursor = [Windows.Input.Cursors]::Hand
        [Windows.Controls.Canvas]::SetLeft($dot, $pts[$i].X - $d / 2); [Windows.Controls.Canvas]::SetTop($dot, $pts[$i].Y - $d / 2)
        [void]$cv.Children.Add($dot)
        if ($bk[$i].N -gt 0 -or $isNow) {
            $tx = New-Object Windows.Controls.TextBlock
            $tx.Text = "$($bk[$i].N)"; $tx.Width = 50; $tx.TextAlignment = 'Center'
            $tx.FontFamily = New-Object Windows.Media.FontFamily('Segoe UI Variable Display, Segoe UI')
            $tx.FontSize = $(if ($isNow) { 16 } else { 12.5 }); $tx.FontWeight = $(if ($isNow) { 'Bold' } else { 'SemiBold' })
            $tx.Foreground = Brush $(if ($isNow) { '#1E1A1B' } else { '#6E6466' })
            [Windows.Controls.Canvas]::SetLeft($tx, $pts[$i].X - 25); [Windows.Controls.Canvas]::SetTop($tx, $pts[$i].Y - $(if ($isNow) { 34 } else { 26 }))
            [void]$cv.Children.Add($tx)
        }
    }
    # animación: la línea se dibuja de izquierda a derecha
    if ($script:AnimateChart) {
        $script:AnimateChart = $false
        $r0 = New-Object Windows.Rect(0, -50, 0, ($H + 100))
        $r1 = New-Object Windows.Rect(0, -50, ($W + 30), ($H + 100))
        $clip = New-Object Windows.Media.RectangleGeometry($r0)
        $cv.Clip = $clip
        $ra = New-Object Windows.Media.Animation.RectAnimation($r0, $r1, (New-Object Windows.Duration([TimeSpan]::FromMilliseconds(950))))
        $ce = New-Object Windows.Media.Animation.CubicEase; $ce.EasingMode = 'EaseInOut'; $ra.EasingFunction = $ce
        $clip.BeginAnimation([Windows.Media.RectangleGeometry]::RectProperty, $ra)
    } else { $cv.Clip = $null }
}
function Update-Chart {
    $ui = $script:UI
    $who = $null; if ($ui.StatsWho.SelectedIndex -gt 0) { $who = "$($ui.StatsWho.SelectedItem)" }
    $mode = $(if ($ui.ChartDay.IsChecked) { 'D' } elseif ($ui.ChartWeek.IsChecked) { 'W' } else { 'M' })
    $today = (Get-Date).Date
    $es = $script:ES
    $buckets = New-Object System.Collections.Generic.List[object]
    if ($mode -eq 'D') {
        for ($k = 13; $k -ge 0; $k--) { $d = $today.AddDays(-$k); $buckets.Add(@{ From = $d; To = $d.AddDays(1); Label = $d.ToString('ddd d', $es); Long = $d.ToString('dddd d MMMM', $es) }) }
    } elseif ($mode -eq 'W') {
        $off = (([int]$today.DayOfWeek) + 6) % 7
        $mon = $today.AddDays(-$off)
        for ($k = 11; $k -ge 0; $k--) {
            $d = $mon.AddDays(-7 * $k)
            $wk = $es.Calendar.GetWeekOfYear($d, [Globalization.CalendarWeekRule]::FirstFourDayWeek, [DayOfWeek]::Monday)
            $buckets.Add(@{ From = $d; To = $d.AddDays(7); Label = ((L 'chWeekL') -f $wk); Long = ($d.ToString('d MMM', $es) + ' – ' + $d.AddDays(6).ToString('d MMM', $es)) })
        }
    } else {
        $m0 = Get-Date -Year $today.Year -Month $today.Month -Day 1
        $m0 = $m0.Date
        for ($k = 11; $k -ge 0; $k--) { $d = $m0.AddMonths(-$k); $buckets.Add(@{ From = $d; To = $d.AddMonths(1); Label = $d.ToString('MMM', $es); Long = $d.ToString('MMMM yyyy', $es) }) }
    }
    foreach ($b in $buckets) { $b.N = 0; $b.Bill = [double]0; $b.Paid = [double]0 }
    $start = $buckets[0].From; $end = $buckets[$buckets.Count - 1].To
    if ($script:Cache) {
        foreach ($cid in @($script:Cache.Keys)) {
            if ($script:StatsScopeId -and $cid -ne $script:StatsScopeId) { continue }
            $d = $script:Cache[$cid]
            foreach ($r in $d.Rows) {
                if (-not $r.Fecha -or $r.Fecha -lt $start -or $r.Fecha -ge $end) { continue }
                if ($who -and "$($r.Resp)".Trim() -ine $who) { continue }
                foreach ($b in $buckets) {
                    if ($r.Fecha -ge $b.From -and $r.Fecha -lt $b.To) { $b.N++; $b.Bill += $r.PriceV; $b.Paid += $r.PaidV; break }
                }
            }
        }
    }
    $sumN = 0; $sumBill = [double]0
    foreach ($b in $buckets) { $sumN += $b.N; $sumBill += $b.Bill }
    $script:ChartBuckets = $buckets
    # etiquetas: periodo y dinero
    $lg = $ui.ChartLabels
    $lg.Children.Clear(); $lg.ColumnDefinitions.Clear()
    for ($i = 0; $i -lt $buckets.Count; $i++) {
        $b = $buckets[$i]
        $cd = New-Object Windows.Controls.ColumnDefinition; $cd.Width = New-Object Windows.GridLength([double]1, [Windows.GridUnitType]::Star); $lg.ColumnDefinitions.Add($cd)
        $isNow = ($i -eq $buckets.Count - 1)
        $foot = New-Object Windows.Controls.StackPanel; $foot.HorizontalAlignment = 'Center'
        $lab = New-Object Windows.Controls.TextBlock
        $lab.Text = $b.Label; $lab.FontSize = 12.5; $lab.HorizontalAlignment = 'Center'; $lab.FontWeight = 'SemiBold'
        $lab.Foreground = Brush $(if ($isNow) { '#9E1820' } else { '#3B3536' })
        $mon = New-Object Windows.Controls.TextBlock
        $mon.Text = $(if ($b.Bill -gt 0) { Format-Euro0 $b.Bill } else { '–' }); $mon.FontSize = 11.5; $mon.HorizontalAlignment = 'Center'
        $mon.Foreground = Brush '#8A7F81'; $mon.Margin = '0,2,0,0'
        [void]$foot.Children.Add($lab); [void]$foot.Children.Add($mon)
        [Windows.Controls.Grid]::SetColumn($foot, $i); [void]$lg.Children.Add($foot)
    }
    # comparación con el periodo anterior
    $n = $buckets.Count
    $diff = $buckets[$n - 1].N - $buckets[$n - 2].N
    $sign = $(if ($diff -gt 0) { '▲ +' + $diff } elseif ($diff -lt 0) { '▼ ' + $diff } else { '= 0' })
    $ui.ChartDelta.Text = (L 'chDeltaTxt') -f $sign, (L ('chPrev' + $mode))
    $ui.ChartDelta.Foreground = Brush $(if ($diff -gt 0) { '#1B7F45' } elseif ($diff -lt 0) { '#B42318' } else { '#6E6466' })
    $ui.ChartDeltaB.Background = Brush $(if ($diff -gt 0) { '#E6F4EC' } elseif ($diff -lt 0) { '#FDECEC' } else { '#F1EDED' })
    $ui.ChartDeltaB.Visibility = 'Visible'
    $script:AnimateChart = $true
    Draw-ChartLine
    $ui.ChartTitle.Text = L ('chTitle' + $mode)
    $ui.ChartSub.Text = (L ('chSub' + $mode)) -f $sumN, (Format-Euro0 $sumBill)
}
function Start-CountUp($items) {
    # las cifras suben desde 0 hasta su valor (0,7 s)
    if ($script:CountTimer) { $script:CountTimer.Stop() }
    $script:CountItems = $items
    $script:CountStart = [DateTime]::Now
    foreach ($it in $items) { $it.Tb.Text = Format-Euro0 0 }
    $tm = New-Object Windows.Threading.DispatcherTimer
    $tm.Interval = [TimeSpan]::FromMilliseconds(16)
    $tm.Add_Tick({
        $p = ([DateTime]::Now - $script:CountStart).TotalMilliseconds / 700
        if ($p -ge 1) { $p = 1 }
        $e = 1 - [math]::Pow(1 - $p, 3)
        foreach ($it in $script:CountItems) { $it.Tb.Text = Format-Euro0 ($it.To * $e) }
        if ($p -ge 1) { $script:CountTimer.Stop() }
    })
    $script:CountTimer = $tm
    $tm.Start()
}
function Update-Stats {
    $ui = $script:UI
    if ($script:StatsFilling) { return }
    $g = $ui.StatsGrid
    $g.Children.Clear(); $g.RowDefinitions.Clear(); $g.ColumnDefinitions.Clear()
    foreach ($w in @(@(2.2, 'Star'), @(70, 'Pixel'), @(1, 'Star'), @(1, 'Star'), @(1, 'Star'), @(1, 'Star'), @(70, 'Pixel'), @(1.4, 'Star'))) {
        $cd = New-Object Windows.Controls.ColumnDefinition
        $cd.Width = New-Object Windows.GridLength([double]$w[0], [Windows.GridUnitType]$w[1])
        $g.ColumnDefinitions.Add($cd)
    }
    $who = $null; if ($ui.StatsWho.SelectedIndex -gt 0) { $who = "$($ui.StatsWho.SelectedItem)" }
    $today = (Get-Date).Date; $from = $null; $to = $null
    switch ($ui.StatsPeriod.SelectedIndex) {
        1 { $from = Get-Date -Year $today.Year -Month $today.Month -Day 1; $to = $from.AddMonths(1) }
        2 { $to = Get-Date -Year $today.Year -Month $today.Month -Day 1; $from = $to.AddMonths(-1) }
        3 { $from = Get-Date -Year $today.Year -Month 1 -Day 1; $to = $from.AddYears(1) }
    }
    if ($from) { $from = $from.Date; $to = $to.Date }
    $scope = $script:StatsScopeId
    $scopeTitle = ''
    $lines = New-Object System.Collections.Generic.List[object]
    $tot = @{ N = 0; Bill = 0.0; Paid = 0.0; Pend = 0.0; Net = 0.0; Debt = 0 }
    $nTram = 0
    foreach ($t in $script:Tramites) {
        if (-not $script:Cache -or -not $script:Cache.ContainsKey($t.Id)) { continue }
        if ($scope -and $t.Id -ne $scope) { continue }
        $nTram++
        $title = $t.Title[$script:Cfg.Language]; if (-not $title) { $title = $t.Title['es'] }
        if ($scope) { $scopeTitle = $title }
        $groups = [ordered]@{}
        foreach ($r in $script:Cache[$t.Id].Rows) {
            if ($who -and "$($r.Resp)".Trim() -ine $who) { continue }
            if ($from -and (-not $r.Fecha -or $r.Fecha -lt $from -or $r.Fecha -ge $to)) { continue }
            $key = $(if ($scope) { $(if ("$($r.Status)".Trim()) { "$($r.Status)".Trim() } else { L 'stNoStatus' }) } else { $t.Id })
            if (-not $groups.Contains($key)) { $groups[$key] = @{ N = 0; Bill = 0.0; Paid = 0.0; Pend = 0.0; Net = 0.0; Debt = 0 } }
            $a = $groups[$key]
            $a.N++; $a.Bill += $r.PriceV; $a.Paid += $r.PaidV; $a.Pend += $r.Pending; $a.Net += $r.NetV
            if ($r.Pending -gt 0) { $a.Debt++ }
        }
        if ($scope) {
            foreach ($st in @($t.StatusColors.Keys)) { if (-not $groups.Contains($st)) { $groups[$st] = @{ N = 0; Bill = 0.0; Paid = 0.0; Pend = 0.0; Net = 0.0; Debt = 0 } } }
            foreach ($st in @($groups.Keys)) {
                $a = $groups[$st]
                foreach ($k in @($tot.Keys)) { $tot[$k] += $a[$k] }
                $col = $(if ($t.StatusColors.ContainsKey($st)) { $t.StatusColors[$st] } else { '#C9C1C3' })
                $lines.Add([pscustomobject]@{ Title = $st; A = $a; Color = $col; Id = $null })
            }
        } else {
            $a = $(if ($groups.Contains($t.Id)) { $groups[$t.Id] } else { @{ N = 0; Bill = 0.0; Paid = 0.0; Pend = 0.0; Net = 0.0; Debt = 0 } })
            foreach ($k in @($tot.Keys)) { $tot[$k] += $a[$k] }
            $lines.Add([pscustomobject]@{ Title = $title; A = $a; Color = '#C4161C'; Id = $t.Id })
        }
    }
    $sorted = @($lines | Sort-Object { $_.A.Bill }, { $_.A.N } -Descending)
    $max = 0.0; foreach ($l in $sorted) { if ($l.A.Bill -gt $max) { $max = $l.A.Bill } }
    # cabecera
    $rowIdx = 0
    $rd = New-Object Windows.Controls.RowDefinition; $rd.Height = 34; $g.RowDefinitions.Add($rd)
    $heads = @(@($(if ($scope) { 'stStatus' } else { 'stTramite' }), 'Left'), @('stExp', 'Right'), @('stBill', 'Right'), @('stPaid', 'Right'), @('stPend', 'Right'), @('stNet', 'Right'), @('stMargin', 'Right'), @('stShare', 'Left'))
    for ($c = 0; $c -lt $heads.Count; $c++) {
        $h = New-StatCell (L $heads[$c][0]) $c $heads[$c][1] '#6E6466' $true 12.5
        [Windows.Controls.Grid]::SetRow($h, 0); [void]$g.Children.Add($h)
    }
    $line = New-Object Windows.Controls.Border; $line.BorderBrush = Brush '#E9E4E5'; $line.BorderThickness = '0,0,0,1'
    [Windows.Controls.Grid]::SetColumnSpan($line, 8); [void]$g.Children.Add($line)
    foreach ($l in $sorted) {
        $rowIdx++
        $a = $l.A
        $rd = New-Object Windows.Controls.RowDefinition; $rd.Height = 54; $g.RowDefinitions.Add($rd)
        $muted = ($a.N -eq 0)
        $bg = New-Object Windows.Controls.Border; $bg.BorderBrush = Brush '#F1EDED'; $bg.BorderThickness = '0,0,0,1'
        [Windows.Controls.Grid]::SetRow($bg, $rowIdx); [Windows.Controls.Grid]::SetColumnSpan($bg, 8); [void]$g.Children.Add($bg)
        $name = New-Object Windows.Controls.StackPanel; $name.Orientation = 'Horizontal'; $name.Margin = '8,0'; $name.VerticalAlignment = 'Center'
        $dot = New-Object Windows.Shapes.Ellipse; $dot.Width = 8; $dot.Height = 8; $dot.Margin = '0,0,10,0'
        $dot.Fill = Brush $(if ($muted) { '#D9D3D4' } else { $l.Color })
        if ($scope) { $dot.Width = 10; $dot.Height = 10; $dot.Stroke = Brush '#33000000'; $dot.StrokeThickness = 0.6 }
        $nt = New-Object Windows.Controls.TextBlock; $nt.Text = $l.Title; $nt.FontSize = 15
        $nt.Foreground = Brush $(if ($muted) { '#A69C9E' } else { '#1E1A1B' }); if (-not $muted) { $nt.FontWeight = 'SemiBold' }
        [void]$name.Children.Add($dot); [void]$name.Children.Add($nt)
        [Windows.Controls.Grid]::SetRow($name, $rowIdx); [void]$g.Children.Add($name)
        $fg = $(if ($muted) { '#B8B0B2' } else { '#2B2526' })
        $margin = $(if ($a.Bill -gt 0) { [math]::Round(100 * $a.Net / $a.Bill) } else { $null })
        $cells = @(
            @("$($a.N)", 1, $fg),
            @((Format-Euro0 $a.Bill), 2, $fg),
            @((Format-Euro0 $a.Paid), 3, $(if ($muted) { $fg } else { '#2E7D32' })),
            @((Format-Euro0 $a.Pend), 4, $(if ($a.Pend -gt 0) { '#B42318' } else { $fg })),
            @((Format-Euro0 $a.Net), 5, $fg),
            @($(if ($null -ne $margin) { "$margin %" } else { '–' }), 6, $fg)
        )
        foreach ($c in $cells) {
            $tb = New-StatCell $c[0] $c[1] 'Right' $c[2] (-not $muted -and $c[1] -in 2, 5)
            [Windows.Controls.Grid]::SetRow($tb, $rowIdx); [void]$g.Children.Add($tb)
        }
        # barra de peso
        $bar = New-Object Windows.Controls.Grid; $bar.Height = 10; $bar.Margin = '12,0,8,0'; $bar.VerticalAlignment = 'Center'
        $pct = $(if ($max -gt 0) { [double]$a.Bill / [double]$max } else { [double]0 })
        $c1 = New-Object Windows.Controls.ColumnDefinition; $c1.Width = New-Object Windows.GridLength([math]::Max($pct, 0.0001), 'Star')
        $c2 = New-Object Windows.Controls.ColumnDefinition; $c2.Width = New-Object Windows.GridLength([math]::Max(1 - $pct, 0.0001), 'Star')
        $bar.ColumnDefinitions.Add($c1); $bar.ColumnDefinitions.Add($c2)
        $track = New-Object Windows.Controls.Border; $track.Background = Brush '#F3EFEF'; $track.CornerRadius = 4
        [Windows.Controls.Grid]::SetColumnSpan($track, 2); [void]$bar.Children.Add($track)
        if ($pct -gt 0) {
            $fill = New-Object Windows.Controls.Border; $fill.CornerRadius = 4
            $lg = New-Object Windows.Media.LinearGradientBrush; $lg.StartPoint = New-Object Windows.Point(0, 0); $lg.EndPoint = New-Object Windows.Point(1, 0)
            $lg.GradientStops.Add((New-Object Windows.Media.GradientStop(([Windows.Media.ColorConverter]::ConvertFromString('#D8202A')), 0)))
            $lg.GradientStops.Add((New-Object Windows.Media.GradientStop(([Windows.Media.ColorConverter]::ConvertFromString('#7E1118')), 1)))
            $fill.Background = $lg
            [void]$bar.Children.Add($fill)
        }
        [Windows.Controls.Grid]::SetRow($bar, $rowIdx); [Windows.Controls.Grid]::SetColumn($bar, 7); [void]$g.Children.Add($bar)
        if ($l.Id) {
            $hit = New-Object Windows.Controls.Border
            $hit.Background = [Windows.Media.Brushes]::Transparent; $hit.Cursor = [Windows.Input.Cursors]::Hand; $hit.Tag = $l.Id
            $hit.ToolTip = L 'stClickHint'; $hit.CornerRadius = 6
            $hit.Add_MouseEnter({ $this.Background = Brush '#0A9E1820' })
            $hit.Add_MouseLeave({ $this.Background = [Windows.Media.Brushes]::Transparent })
            $hit.Add_MouseLeftButtonUp({ $script:StatsScopeId = "$($this.Tag)"; Fill-StatsFilters; Update-Stats })
            [Windows.Controls.Grid]::SetRow($hit, $rowIdx); [Windows.Controls.Grid]::SetColumnSpan($hit, 8); [void]$g.Children.Add($hit)
        }
    }
    # total
    $rowIdx++
    $rd = New-Object Windows.Controls.RowDefinition; $rd.Height = 60; $g.RowDefinitions.Add($rd)
    $tb = New-Object Windows.Controls.Border; $tb.Background = Brush '#FAF7F7'; $tb.BorderBrush = Brush '#1E1A1B'; $tb.BorderThickness = '0,1.5,0,0'; $tb.CornerRadius = '0,0,8,8'
    [Windows.Controls.Grid]::SetRow($tb, $rowIdx); [Windows.Controls.Grid]::SetColumnSpan($tb, 8); [void]$g.Children.Add($tb)
    $tmargin = $(if ($tot.Bill -gt 0) { [math]::Round(100 * $tot.Net / $tot.Bill) } else { $null })
    $cells = @(
        @((L 'stTotal'), 0, '#1E1A1B', 'Left'), @("$($tot.N)", 1, '#1E1A1B', 'Right'),
        @((Format-Euro0 $tot.Bill), 2, '#1E1A1B', 'Right'), @((Format-Euro0 $tot.Paid), 3, '#2E7D32', 'Right'),
        @((Format-Euro0 $tot.Pend), 4, $(if ($tot.Pend -gt 0) { '#B42318' } else { '#1E1A1B' }), 'Right'),
        @((Format-Euro0 $tot.Net), 5, '#9E1820', 'Right'), @($(if ($null -ne $tmargin) { "$tmargin %" } else { '–' }), 6, '#1E1A1B', 'Right')
    )
    foreach ($c in $cells) {
        $t2 = New-StatCell $c[0] $c[1] $c[3] $c[2] $true 16
        [Windows.Controls.Grid]::SetRow($t2, $rowIdx); [void]$g.Children.Add($t2)
    }
    # títulos según el alcance
    $ui.StatsTable.Text = $(if ($scope) { (L 'statsTableSt') -f $scopeTitle } else { L 'statsTable' })
    $ui.StatsAllTxt.Text = L 'stAll'
    $ui.StatsAllBtn.Visibility = $(if ($scope) { 'Visible' } else { 'Collapsed' })
    # tabla: aparece con un fundido suave
    $dur = New-Object Windows.Duration([TimeSpan]::FromMilliseconds(450))
    $ease = New-Object Windows.Media.Animation.CubicEase; $ease.EasingMode = 'EaseOut'
    $tt = New-Object Windows.Media.TranslateTransform(0, 12); $g.RenderTransform = $tt
    $fa = New-Object Windows.Media.Animation.DoubleAnimation([double]0, [double]1, $dur); $fa.EasingFunction = $ease
    $ya = New-Object Windows.Media.Animation.DoubleAnimation([double]12, [double]0, $dur); $ya.EasingFunction = $ease
    $g.BeginAnimation([Windows.UIElement]::OpacityProperty, $fa)
    $tt.BeginAnimation([Windows.Media.TranslateTransform]::YProperty, $ya)
    # tarjetas
    $ui.KpiBill.Text = Format-Euro0 $tot.Bill
    $ui.KpiBillF.Text = (L 'kpiBillF') -f $tot.N
    $ui.KpiPaid.Text = Format-Euro0 $tot.Paid
    $pp = $(if ($tot.Bill -gt 0) { [math]::Min([double]1, [double]$tot.Paid / [double]$tot.Bill) } else { [double]0 })
    $ui.KpiPaidTrack.ColumnDefinitions[0].Width = New-Object Windows.GridLength([math]::Max($pp, 0.0001), 'Star')
    $ui.KpiPaidTrack.ColumnDefinitions[1].Width = New-Object Windows.GridLength([math]::Max(1 - $pp, 0.0001), 'Star')
    $ui.KpiPaidF.Text = (L 'kpiPaidF') -f [math]::Round(100 * $pp)
    $ui.KpiPend.Text = Format-Euro0 $tot.Pend
    $ui.KpiPendF.Text = (L 'kpiPendF') -f $tot.Debt
    $ui.KpiNet.Text = Format-Euro0 $tot.Net
    $ui.KpiNetF.Text = (L 'kpiNetF') -f $(if ($null -ne $tmargin) { $tmargin } else { 0 })
    $ui.StatsSub.Text = $(if ($scope) { (L 'statsSubOne') -f $scopeTitle, $tot.N, (Get-Date -Format 'HH:mm') } else { (L 'statsSub') -f $nTram, $tot.N, (Get-Date -Format 'HH:mm') })
    Start-CountUp @(@{ Tb = $ui.KpiBill; To = [double]$tot.Bill }, @{ Tb = $ui.KpiPaid; To = [double]$tot.Paid }, @{ Tb = $ui.KpiPend; To = [double]$tot.Pend }, @{ Tb = $ui.KpiNet; To = [double]$tot.Net })
    Update-Chart
}

# =====================================================================
#  ADMINISTRADOR (oculto) · CONTROL REMOTO · © Yano
# =====================================================================
$script:AdminPath = Join-Path $script:DataDir 'admin.dat'
$script:Admin = @{ Salt = ''; Hash = ''; Locks = @{}; IgnorePause = $false; OwnerPC = $false }
# PC del creador: se reconoce solo (huella cifrada de nombre del PC + usuario de Windows, no se guarda el nombre)
$script:OwnerPrints = @('799CEE89435398F2A3BAA342547E3E364B49C27853D8398CD2B837E49B83271D')
function Test-OwnerPC {
    if ($script:Admin.OwnerPC) { return $true }
    try {
        $sha = [Security.Cryptography.SHA256]::Create()
        $b = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(('GESTEM|' + $env:COMPUTERNAME + '|' + $env:USERNAME).ToUpperInvariant()))
        return ($script:OwnerPrints -contains ([BitConverter]::ToString($b) -replace '-', ''))
    } catch { return $false }
}
$script:IsOwnerPC = $false
$script:IsAdmin = $false
$script:AdmLogFilter = 'all'
$script:LockKeys = @('File', 'Ext', 'Comp', 'Team', 'Pref', 'Help', 'All')
$script:AdmTr = @{
 'GESTEM está en pausa' = @{ en = 'GESTEM is paused'; fr = 'GESTEM est en pause'; ar = 'GESTEM متوقف مؤقتًا' }
 'Contacta con Yano.' = @{ en = 'Contact Yano.'; fr = 'Contactez Yano.'; ar = 'تواصل مع Yano.' }
 'Administrador' = @{ en = 'Administrator'; fr = 'Administrateur'; ar = 'المسؤول' }
 'Cancelar' = @{ en = 'Cancel'; fr = 'Annuler'; ar = 'إلغاء' }
 'Entrar' = @{ en = 'Sign in'; fr = 'Entrer'; ar = 'دخول' }
 'Salir del modo admin' = @{ en = 'Leave admin mode'; fr = 'Quitter le mode admin'; ar = 'الخروج من وضع المسؤول' }
 'Cerrar' = @{ en = 'Close'; fr = 'Fermer'; ar = 'إغلاق' }
 'Modo administrador' = @{ en = 'Admin mode'; fr = 'Mode administrateur'; ar = 'وضع المسؤول' }
 'Solo visible para Yano · © 2026 todos los derechos reservados' = @{ en = 'Only visible to Yano · © 2026 all rights reserved'; fr = 'Visible uniquement par Yano · © 2026 tous droits réservés'; ar = 'مرئي فقط لـ Yano · © 2026 جميع الحقوق محفوظة' }
 'Bloqueos' = @{ en = 'Locks'; fr = 'Verrous'; ar = 'القيود' }
 'Actividad' = @{ en = 'Activity'; fr = 'Activité'; ar = 'النشاط' }
 'Herramientas' = @{ en = 'Tools'; fr = 'Outils'; ar = 'الأدوات' }
 'Oculta partes de Ajustes a las compañeras' = @{ en = 'Hide parts of Settings from your colleagues'; fr = 'Masquer des parties des Paramètres aux collègues'; ar = 'إخفاء أجزاء من الإعدادات عن الزملاء' }
 'Lo que bloquees desaparece de Ajustes para ellas (no lo ven ni lo pueden cambiar). Tú lo sigues viendo mientras estés en modo admin.' = @{ en = 'Whatever you lock disappears from their Settings (they can''t see or change it). You still see it while you are in admin mode.'; fr = 'Ce que vous verrouillez disparaît de leurs Paramètres (elles ne le voient pas et ne peuvent pas le modifier). Vous le voyez toujours en mode admin.'; ar = 'ما تقفله يختفي من إعداداتهم (لا يرونه ولا يمكنهم تغييره). تبقى تراه أثناء وضع المسؤول.' }
 'Archivo de clientes (Excel)' = @{ en = 'Client file (Excel)'; fr = 'Fichier des clients (Excel)'; ar = 'ملف العملاء (Excel)' }
 'Copia externa' = @{ en = 'External backup'; fr = 'Copie externe'; ar = 'النسخة الخارجية' }
 'Empresa (nombre y logo)' = @{ en = 'Company (name and logo)'; fr = 'Entreprise (nom et logo)'; ar = 'الشركة (الاسم والشعار)' }
 'Equipo (compañeras)' = @{ en = 'Team (colleagues)'; fr = 'Équipe (collègues)'; ar = 'الفريق (الزملاء)' }
 'Preferencias (idioma, nacionalidad, copias)' = @{ en = 'Preferences (language, nationality, backups)'; fr = 'Préférences (langue, nationalité, copies)'; ar = 'التفضيلات (اللغة، الجنسية، النسخ)' }
 'Ayuda y versión (actualizaciones)' = @{ en = 'Help and version (updates)'; fr = 'Aide et version (mises à jour)'; ar = 'المساعدة والإصدار (التحديثات)' }
 'Ocultar Ajustes por completo' = @{ en = 'Hide Settings completely'; fr = 'Masquer complètement les Paramètres'; ar = 'إخفاء الإعدادات بالكامل' }
 'Ignorar la pausa remota en este PC' = @{ en = 'Ignore the remote pause on this PC'; fr = 'Ignorer la pause à distance sur ce PC'; ar = 'تجاهل الإيقاف عن بُعد على هذا الحاسوب' }
 'Guardar bloqueos' = @{ en = 'Save locks'; fr = 'Enregistrer les verrous'; ar = 'حفظ القيود' }
 'Actualizar' = @{ en = 'Refresh'; fr = 'Actualiser'; ar = 'تحديث' }
 'Control remoto (GitHub)' = @{ en = 'Remote control (GitHub)'; fr = 'Contrôle à distance (GitHub)'; ar = 'التحكم عن بُعد (GitHub)' }
 'Buscar actualización ahora' = @{ en = 'Check for updates now'; fr = 'Rechercher une mise à jour'; ar = 'البحث عن تحديث الآن' }
 'Abrir carpeta de datos' = @{ en = 'Open data folder'; fr = 'Ouvrir le dossier des données'; ar = 'فتح مجلد البيانات' }
 'Abrir carpeta del programa' = @{ en = 'Open program folder'; fr = 'Ouvrir le dossier du programme'; ar = 'فتح مجلد البرنامج' }
 'Ver Ajustes completos' = @{ en = 'Show full Settings'; fr = 'Voir tous les Paramètres'; ar = 'عرض الإعدادات كاملة' }
 'Cambiar contraseña de administrador' = @{ en = 'Change admin password'; fr = 'Changer le mot de passe administrateur'; ar = 'تغيير كلمة مرور المسؤول' }
 'Cambiar' = @{ en = 'Change'; fr = 'Changer'; ar = 'تغيير' }
 'Primera vez: crea tu contraseña de administrador (mínimo 6 caracteres) y escríbela dos veces. Guárdala bien: no se puede recuperar.' = @{ en = 'First time: create your admin password (at least 6 characters) and type it twice. Keep it safe: it can''t be recovered.'; fr = 'Première fois : créez votre mot de passe administrateur (6 caractères minimum) et tapez-le deux fois. Gardez-le bien : il ne peut pas être récupéré.'; ar = 'المرة الأولى: أنشئ كلمة مرور المسؤول (6 أحرف على الأقل) واكتبها مرتين. احتفظ بها جيدًا: لا يمكن استرجاعها.' }
 'Escribe tu contraseña de administrador.' = @{ en = 'Type your admin password.'; fr = 'Tapez votre mot de passe administrateur.'; ar = 'اكتب كلمة مرور المسؤول.' }
 'Mínimo 6 caracteres.' = @{ en = 'At least 6 characters.'; fr = '6 caractères minimum.'; ar = '6 أحرف على الأقل.' }
 'Las dos contraseñas no son iguales.' = @{ en = 'The two passwords don''t match.'; fr = 'Les deux mots de passe ne sont pas identiques.'; ar = 'كلمتا المرور غير متطابقتين.' }
 'Contraseña incorrecta.' = @{ en = 'Wrong password.'; fr = 'Mot de passe incorrect.'; ar = 'كلمة المرور غير صحيحة.' }
 'Bloq Mayús está activado.' = @{ en = 'Caps Lock is on.'; fr = 'Verr. Maj est activé.'; ar = 'مفتاح الأحرف الكبيرة مفعّل.' }
 'Teclado: {0}' = @{ en = 'Keyboard: {0}'; fr = 'Clavier : {0}'; ar = 'لوحة المفاتيح: {0}' }
 'Versión {0} · datos: {1} · programa: {2}' = @{ en = 'Version {0} · data: {1} · program: {2}'; fr = 'Version {0} · données : {1} · programme : {2}'; ar = 'الإصدار {0} · البيانات: {1} · البرنامج: {2}' }
 'Actividad reciente · {0} registros' = @{ en = 'Recent activity · {0} entries'; fr = 'Activité récente · {0} entrées'; ar = 'النشاط الأخير · {0} سجلات' }
 '✓ Guardado' = @{ en = '✓ Saved'; fr = '✓ Enregistré'; ar = '✓ تم الحفظ' }
 '✓ Contraseña cambiada' = @{ en = '✓ Password changed'; fr = '✓ Mot de passe changé'; ar = '✓ تم تغيير كلمة المرور' }
 'El programa está en pausa temporalmente.' = @{ en = 'The program is temporarily paused.'; fr = 'Le programme est temporairement en pause.'; ar = 'البرنامج متوقف مؤقتًا.' }
 'GESTEM lee el archivo control.json de tu repositorio de GitHub al abrirse (nunca envía datos). Campos: «mensaje» (texto que verán una vez), «pausa» (true/false) y «motivo».' = @{ en = 'When it opens, GESTEM reads control.json from your GitHub repository (it never sends data). Fields: «mensaje» (text they see once), «pausa» (true/false) and «motivo» (reason).'; fr = 'À l''ouverture, GESTEM lit control.json dans votre dépôt GitHub (il n''envoie jamais de données). Champs : « mensaje » (texte vu une fois), « pausa » (true/false) et « motivo » (raison).'; ar = 'عند الفتح يقرأ GESTEM ملف control.json من مستودعك على GitHub (ولا يرسل أي بيانات). الحقول: «mensaje» (نص يُعرض مرة واحدة)، «pausa» (true/false) و«motivo» (السبب).' }
 'Estado: no hay control.json o no se pudo leer → todo normal.' = @{ en = 'Status: no control.json, or it couldn''t be read → everything normal.'; fr = 'État : pas de control.json ou illisible → tout est normal.'; ar = 'الحالة: لا يوجد control.json أو تعذّرت قراءته ← كل شيء طبيعي.' }
 'Estado actual → mensaje: {0}   ·   pausa: {1}' = @{ en = 'Current status → message: {0}   ·   pause: {1}'; fr = 'État actuel → message : {0}   ·   pause : {1}'; ar = 'الحالة الحالية ← الرسالة: {0}   ·   الإيقاف: {1}' }
 '(vacío)' = @{ en = '(empty)'; fr = '(vide)'; ar = '(فارغ)' }
 'Este PC es mío: todo lo que se haga aquí cuenta como admin' = @{ en = 'This PC is mine: everything done here counts as admin'; fr = 'Ce PC est à moi : tout ce qui est fait ici compte comme admin'; ar = 'هذا الحاسوب لي: كل ما يُفعل هنا يُحسب للمسؤول' }
 'Todo' = @{ en = 'All'; fr = 'Tout'; ar = 'الكل' }
 'Errores' = @{ en = 'Errors'; fr = 'Erreurs'; ar = 'الأخطاء' }
 'Errores · {0} en este PC (los más recientes arriba)' = @{ en = 'Errors · {0} on this PC (most recent first)'; fr = 'Erreurs · {0} sur ce PC (les plus récentes en haut)'; ar = 'الأخطاء · {0} على هذا الجهاز (الأحدث أولاً)' }
 'Equipo' = @{ en = 'Team'; fr = 'Équipe'; ar = 'الفريق' }
 'Admin' = @{ en = 'Admin'; fr = 'Admin'; ar = 'المسؤول' }
 'Sistema' = @{ en = 'System'; fr = 'Système'; ar = 'النظام' }
 'Actividad · {0} del equipo · {1} del admin' = @{ en = 'Activity · {0} by the team · {1} by the admin'; fr = 'Activité · {0} par l''équipe · {1} par l''admin'; ar = 'النشاط · {0} من الفريق · {1} من المسؤول' }
 'SÍ' = @{ en = 'YES'; fr = 'OUI'; ar = 'نعم' }
 'no' = @{ en = 'no'; fr = 'non'; ar = 'لا' }
 'añadió cliente' = @{ en = 'added client'; fr = 'a ajouté le client'; ar = 'أضاف العميل' }
 'editó cliente' = @{ en = 'edited client'; fr = 'a modifié le client'; ar = 'عدّل العميل' }
 'cambió los ajustes' = @{ en = 'changed the settings'; fr = 'a modifié les paramètres'; ar = 'غيّر الإعدادات' }
 'cerró GESTEM' = @{ en = 'closed GESTEM'; fr = 'a fermé GESTEM'; ar = 'أغلق GESTEM' }
 'sesión iniciada' = @{ en = 'signed in'; fr = 'session ouverte'; ar = 'بدأ الجلسة' }
 'copia externa:' = @{ en = 'external backup:'; fr = 'copie externe :'; ar = 'نسخة خارجية:' }
 'admin: acceso' = @{ en = 'admin: signed in'; fr = 'admin : connexion'; ar = 'المسؤول: دخول' }
 'admin: salida' = @{ en = 'admin: left'; fr = 'admin : sortie'; ar = 'المسؤول: خروج' }
 'admin: intento de acceso fallido' = @{ en = 'admin: failed sign-in attempt'; fr = 'admin : tentative de connexion échouée'; ar = 'المسؤول: محاولة دخول فاشلة' }
 'admin: contraseña cambiada' = @{ en = 'admin: password changed'; fr = 'admin : mot de passe changé'; ar = 'المسؤول: تم تغيير كلمة المرور' }
 'admin: bloqueos guardados' = @{ en = 'admin: locks saved'; fr = 'admin : verrous enregistrés'; ar = 'المسؤول: تم حفظ القيود' }
 'pausa remota activa' = @{ en = 'remote pause active'; fr = 'pause à distance active'; ar = 'الإيقاف عن بُعد مفعّل' }
}
# manual en otros idiomas (las claves son el texto en español)
$script:HelpTr = @{
 '1. Empezar' = @{ en = '1. Getting started'; fr = '1. Démarrer'; ar = '1. البداية' }
 'Al abrir GESTEM, elige tu nombre en «¿Quién está usando GESTEM?». Si no estás en la lista, escríbelo en «Otro nombre…» y pulsa «Entrar».' = @{ en = 'When GESTEM opens, pick your name in «Who is using GESTEM?». If you are not on the list, type it in «Another name…» and press «Start».'; fr = 'À l''ouverture, choisissez votre nom dans « Qui utilise GESTEM ? ». Si vous n''êtes pas dans la liste, tapez-le dans « Autre nom… » et cliquez sur « Entrer ».'; ar = 'عند فتح GESTEM اختر اسمك في «من يستخدم GESTEM؟». إذا لم يكن اسمك في القائمة فاكتبه في «اسم آخر…» واضغط «دخول».' }
 'Tu nombre sale abajo a la izquierda. Haz clic en él para cambiar de compañera sin cerrar el programa.' = @{ en = 'Your name appears at the bottom left. Click it to switch colleague without closing the program.'; fr = 'Votre nom apparaît en bas à gauche. Cliquez dessus pour changer de collègue sans fermer le programme.'; ar = 'يظهر اسمك أسفل اليسار. اضغط عليه لتغيير الزميل دون إغلاق البرنامج.' }
 'Arriba, en «Trámite», elige con qué trabajas: Canje DGT, Homologación, Extranjería, Autónomo… Cada trámite tiene su pestaña en el Excel.' = @{ en = 'At the top, in «Procedure», choose what you work on: DGT licence exchange, Homologation, Immigration, Self-employed… Each procedure has its own tab in the Excel file.'; fr = 'En haut, dans « Démarche », choisissez sur quoi vous travaillez : échange de permis DGT, homologation, étrangers, autoentrepreneur… Chaque démarche a son onglet dans l''Excel.'; ar = 'في الأعلى، في «الإجراء»، اختر ما تعمل عليه: استبدال رخصة DGT، المعادلة، الأجانب، العمل الحر… لكل إجراء ورقته في ملف Excel.' }
 '2. Añadir un cliente nuevo' = @{ en = '2. Adding a new client'; fr = '2. Ajouter un nouveau client'; ar = '2. إضافة عميل جديد' }
 'Menú «Nuevo cliente». Rellena los datos de arriba abajo: nombre, documento, teléfono, datos del trámite, estado y pago.' = @{ en = 'Menu «New client». Fill in the details from top to bottom: name, ID document, phone, procedure details, status and payment.'; fr = 'Menu « Nouveau client ». Remplissez de haut en bas : nom, pièce d''identité, téléphone, démarche, état et paiement.'; ar = 'قائمة «عميل جديد». املأ البيانات من الأعلى إلى الأسفل: الاسم، الوثيقة، الهاتف، بيانات الإجراء، الحالة والدفع.' }
 'La nacionalidad (MARROQUI), la fecha de hoy y el responsable (tu nombre) se ponen solos. Puedes cambiarlos.' = @{ en = 'Nationality (MARROQUI), today''s date and the person in charge (your name) are filled in automatically. You can change them.'; fr = 'La nationalité (MARROQUI), la date du jour et le responsable (votre nom) se remplissent seuls. Vous pouvez les changer.'; ar = 'الجنسية (MARROQUI) وتاريخ اليوم والمسؤول (اسمك) تُملأ تلقائيًا. يمكنك تغييرها.' }
 'Si el NIE o pasaporte ya existe, sale un aviso en rojo para no repetir clientes.' = @{ en = 'If the NIE or passport already exists, a red warning appears so clients are not entered twice.'; fr = 'Si le NIE ou le passeport existe déjà, un avertissement rouge s''affiche pour éviter les doublons.'; ar = 'إذا كان NIE أو جواز السفر موجودًا مسبقًا يظهر تنبيه أحمر لتجنّب تكرار العميل.' }
 'Pulsa «Guardar cliente» (o Ctrl + S). El cliente se escribe en el Excel al momento, con su número de expediente.' = @{ en = 'Press «Save client» (or Ctrl + S). The client is written to the Excel file straight away, with its case number.'; fr = 'Cliquez sur « Enregistrer le client » (ou Ctrl + S). Le client est écrit dans l''Excel immédiatement, avec son numéro de dossier.'; ar = 'اضغط «حفظ العميل» (أو Ctrl + S). يُكتب العميل في Excel فورًا مع رقم ملفه.' }
 '3. Pagos' = @{ en = '3. Payments'; fr = '3. Paiements'; ar = '3. المدفوعات' }
 'Precio total: lo que paga el cliente. En el canje DGT se pone solo según el tipo de permiso.' = @{ en = 'Total price: what the client pays. For the DGT exchange it is set automatically from the licence type.'; fr = 'Prix total : ce que paie le client. Pour l''échange DGT, il se remplit selon le type de permis.'; ar = 'السعر الإجمالي: ما يدفعه العميل. في استبدال DGT يُحدَّد تلقائيًا حسب نوع الرخصة.' }
 'Estado de pago: «Pagado» (lo ha pagado todo), «Parcial» (escribe cuánto ha pagado en «Ha pagado») o «Pendiente» (no ha pagado nada).' = @{ en = 'Payment status: «Pagado» (paid in full), «Parcial» (type how much was paid in «Amount paid») or «Pendiente» (nothing paid yet).'; fr = 'État du paiement : « Pagado » (tout payé), « Parcial » (indiquez le montant dans « Montant payé ») ou « Pendiente » (rien payé).'; ar = 'حالة الدفع: «Pagado» (دفع كل شيء)، «Parcial» (اكتب المبلغ في «المبلغ المدفوع») أو «Pendiente» (لم يدفع شيئًا).' }
 'Gastos: lo que pagáis vosotras (tasa, traducción…). El beneficio NETO = precio − gastos se calcula solo en el Excel.' = @{ en = 'Costs: what the office pays (fees, translation…). NET profit = price − costs is calculated automatically in the Excel file.'; fr = 'Frais : ce que paie le bureau (taxe, traduction…). Le bénéfice NET = prix − frais se calcule seul dans l''Excel.'; ar = 'المصاريف: ما يدفعه المكتب (رسوم، ترجمة…). الربح الصافي = السعر − المصاريف يُحسب تلقائيًا في Excel.' }
 'Lo que falta por cobrar se calcula solo y sale en rojo.' = @{ en = 'The amount still to collect is calculated automatically and shown in red.'; fr = 'Le reste à encaisser se calcule seul et s''affiche en rouge.'; ar = 'المبلغ المتبقي يُحسب تلقائيًا ويظهر بالأحمر.' }
 '4. Buscar y cambiar clientes' = @{ en = '4. Finding and editing clients'; fr = '4. Chercher et modifier des clients'; ar = '4. البحث عن العملاء وتعديلهم' }
 'Menú «Clientes». Busca por nombre, NIE, número de expediente o localizador.' = @{ en = 'Menu «Clients». Search by name, NIE, case number or reference code.'; fr = 'Menu « Clients ». Cherchez par nom, NIE, numéro de dossier ou code de référence.'; ar = 'قائمة «العملاء». ابحث بالاسم أو NIE أو رقم الملف أو الرمز المرجعي.' }
 'Filtra por estado, marca «Solo con pago pendiente» para ver quién debe dinero, o «Solo mis clientes» para ver solo los tuyos.' = @{ en = 'Filter by status, tick «Outstanding payment only» to see who owes money, or «My clients only» to see just yours.'; fr = 'Filtrez par état, cochez « Paiement en attente seulement » pour voir qui doit de l''argent, ou « Mes clients seulement ».'; ar = 'صفِّ حسب الحالة، أو فعّل «الدفع المعلّق فقط» لترى من عليه مال، أو «عملائي فقط».' }
 'Doble clic en un cliente (o «Editar cliente») para cambiar su estado, su pago o sus datos. Pulsa «Guardar cliente» al terminar.' = @{ en = 'Double-click a client (or «Edit client») to change their status, payment or details. Press «Save client» when done.'; fr = 'Double-cliquez sur un client (ou « Modifier ») pour changer son état, son paiement ou ses données. Cliquez sur « Enregistrer » à la fin.'; ar = 'انقر مرتين على العميل (أو «تعديل العميل») لتغيير حالته أو دفعه أو بياناته. اضغط «حفظ العميل» عند الانتهاء.' }
 '«Abrir Excel» abre el archivo; «Actualizar» vuelve a leerlo si alguien lo ha cambiado a mano.' = @{ en = '«Open Excel» opens the file; «Refresh» reads it again if someone changed it by hand.'; fr = '« Ouvrir Excel » ouvre le fichier ; « Actualiser » le relit si quelqu''un l''a modifié à la main.'; ar = '«فتح Excel» يفتح الملف؛ «تحديث» يعيد قراءته إذا عدّله أحد يدويًا.' }
 '5. Documentos y recibos' = @{ en = '5. Documents and receipts'; fr = '5. Documents et reçus'; ar = '5. المستندات والإيصالات' }
 'En «Clientes», elige un cliente y pulsa «Documentos» (o clic derecho sobre él). Arrastra los escaneos (NIE, permiso, formularios, fotos…) o pulsa «Añadir archivos».' = @{ en = 'In «Clients», select a client and press «Documents» (or right-click it). Drag in the scans (NIE, licence, forms, photos…) or press «Add files».'; fr = 'Dans « Clients », choisissez un client et cliquez sur « Documents » (ou clic droit). Glissez les scans (NIE, permis, formulaires, photos…) ou cliquez sur « Ajouter des fichiers ».'; ar = 'في «العملاء» اختر عميلًا واضغط «المستندات» (أو انقر بالزر الأيمن). اسحب الملفات الممسوحة (NIE، الرخصة، الاستمارات، الصور…) أو اضغط «إضافة ملفات».' }
 'Los documentos se guardan en la carpeta «GESTEM documentos», junto al Excel, una carpeta por expediente. El clip de la lista indica cuántos tiene cada cliente.' = @{ en = 'Documents are kept in the «GESTEM documentos» folder next to the Excel file, one folder per case. The paperclip in the list shows how many each client has.'; fr = 'Les documents sont rangés dans le dossier « GESTEM documentos », à côté de l''Excel, un dossier par dossier client. Le trombone de la liste indique combien chaque client en a.'; ar = 'تُحفظ المستندات في مجلد «GESTEM documentos» بجانب ملف Excel، مجلد لكل ملف. يدل مشبك الورق في القائمة على عددها لكل عميل.' }
 '«Recibo PDF» crea un recibo de pago con el logo de la empresa, lo guarda en los documentos del cliente y lo abre para imprimirlo o enviarlo.' = @{ en = '«PDF receipt» creates a payment receipt with the company logo, saves it in the client''s documents and opens it to print or send.'; fr = '« Reçu PDF » crée un reçu de paiement avec le logo de l''entreprise, l''enregistre dans les documents du client et l''ouvre pour l''imprimer ou l''envoyer.'; ar = '«إيصال PDF» ينشئ إيصال دفع بشعار الشركة، ويحفظه في مستندات العميل ويفتحه للطباعة أو الإرسال.' }
 'El recibo justifica lo pagado; no es una factura.' = @{ en = 'The receipt confirms what was paid; it is not an invoice.'; fr = 'Le reçu justifie ce qui a été payé ; ce n''est pas une facture.'; ar = 'الإيصال يثبت ما تم دفعه؛ وهو ليس فاتورة.' }
 '6. Exportar la lista' = @{ en = '6. Exporting the list'; fr = '6. Exporter la liste'; ar = '6. تصدير القائمة' }
 'En «Clientes», pulsa «Exportar»: «Lista en PDF» (informe para imprimir, con totales) o «Lista en CSV» (se abre en Excel con todas las columnas).' = @{ en = 'In «Clients», press «Export»: «List as PDF» (printable report with totals) or «List as CSV» (opens in Excel with every column).'; fr = 'Dans « Clients », cliquez sur « Exporter » : « Liste en PDF » (rapport imprimable avec totaux) ou « Liste en CSV » (s''ouvre dans Excel avec toutes les colonnes).'; ar = 'في «العملاء» اضغط «تصدير»: «القائمة بصيغة PDF» (تقرير للطباعة مع المجاميع) أو «القائمة بصيغة CSV» (يُفتح في Excel بكل الأعمدة).' }
 'Se exporta lo que ves: si has buscado o filtrado, solo salen esos clientes.' = @{ en = 'What you see is what gets exported: if you searched or filtered, only those clients are included.'; fr = 'On exporte ce que vous voyez : si vous avez cherché ou filtré, seuls ces clients sont inclus.'; ar = 'يُصدَّر ما تراه: إذا بحثت أو صفّيت فلن يظهر إلا هؤلاء العملاء.' }
 '7. Resumen económico' = @{ en = '7. Financial summary'; fr = '7. Résumé financier'; ar = '7. الملخص المالي' }
 'Menú «Resumen». Arriba ves lo facturado, lo cobrado, lo pendiente de cobro y el beneficio neto.' = @{ en = 'Menu «Summary». At the top you see what was billed, collected, still outstanding and the net profit.'; fr = 'Menu « Résumé ». En haut : le facturé, l''encaissé, le reste à encaisser et le bénéfice net.'; ar = 'قائمة «الملخص». في الأعلى ترى المفوتر والمحصّل والمتبقي وصافي الربح.' }
 'Arriba puedes elegir «Todos los trámites» o uno solo; las cifras y la gráfica cambian con él.' = @{ en = 'At the top you can choose «All procedures» or a single one; the figures and the chart follow your choice.'; fr = 'En haut, choisissez « Toutes les démarches » ou une seule ; les chiffres et le graphique suivent votre choix.'; ar = 'في الأعلى يمكنك اختيار «كل الإجراءات» أو إجراء واحد؛ تتغير الأرقام والرسم حسب اختيارك.' }
 'La gráfica muestra la evolución por día, semana o mes: verde cuando sube, rojo cuando baja.' = @{ en = 'The chart shows the trend by day, week or month: green when it goes up, red when it goes down.'; fr = 'Le graphique montre l''évolution par jour, semaine ou mois : vert quand ça monte, rouge quand ça baisse.'; ar = 'يعرض الرسم التطور حسب اليوم أو الأسبوع أو الشهر: أخضر عند الصعود وأحمر عند الهبوط.' }
 'Filtra por compañera y por periodo: este mes, el mes anterior o este año (según la fecha de solicitud).' = @{ en = 'Filter by colleague and by period: this month, last month or this year (by application date).'; fr = 'Filtrez par collègue et par période : ce mois-ci, le mois dernier ou cette année (selon la date de demande).'; ar = 'صفِّ حسب الزميل والفترة: هذا الشهر أو الشهر الماضي أو هذه السنة (حسب تاريخ الطلب).' }
 '8. El Excel y las copias' = @{ en = '8. The Excel file and backups'; fr = '8. L''Excel et les copies'; ar = '8. ملف Excel والنسخ الاحتياطية' }
 'GESTEM guarda los clientes en su propia carpeta «GESTEM datos», junto al Excel, y actualiza vuestro Excel en cada guardado. Si el Excel no se puede abrir en ese momento, se actualiza solo más tarde.' = @{ en = 'GESTEM keeps clients in its own «GESTEM datos» folder next to the Excel file, and updates your Excel on every save. If the Excel can''t be opened at that moment, it is updated automatically later.'; fr = 'GESTEM garde les clients dans son propre dossier « GESTEM datos », à côté de l''Excel, et met votre Excel à jour à chaque enregistrement. Si l''Excel ne peut pas être ouvert à ce moment-là, il est mis à jour plus tard tout seul.'; ar = 'يحفظ GESTEM العملاء في مجلده الخاص «GESTEM datos» بجانب ملف Excel، ويحدّث ملف Excel عند كل حفظ. وإذا تعذّر فتح Excel في تلك اللحظة يُحدَّث تلقائيًا لاحقًا.' }
 'Antes de cada guardado se hace una copia de seguridad del Excel (Ajustes → «Abrir carpeta de copias»).' = @{ en = 'A backup of the Excel file is made before every save (Settings → «Open backups folder»).'; fr = 'Une copie de l''Excel est faite avant chaque enregistrement (Paramètres → « Ouvrir le dossier des copies »).'; ar = 'تُنشأ نسخة احتياطية من Excel قبل كل حفظ (الإعدادات ← «فتح مجلد النسخ»).' }
 'Copia externa: al abrir y al cerrar GESTEM se guarda una copia completa (Excel, ajustes y documentos) en la nube de la empresa (OneDrive o Google Drive) o en la carpeta elegida en Ajustes.' = @{ en = 'External backup: when GESTEM opens and closes, a full copy (Excel, settings and documents) is saved to the company cloud (OneDrive or Google Drive) or to the folder chosen in Settings.'; fr = 'Copie externe : à l''ouverture et à la fermeture, une copie complète (Excel, paramètres et documents) est enregistrée dans le cloud de l''entreprise (OneDrive ou Google Drive) ou dans le dossier choisi.'; ar = 'النسخة الخارجية: عند فتح GESTEM وإغلاقه تُحفظ نسخة كاملة (Excel والإعدادات والمستندات) في سحابة الشركة (OneDrive أو Google Drive) أو في المجلد المختار.' }
 '9. Ajustes' = @{ en = '9. Settings'; fr = '9. Paramètres'; ar = '9. الإعدادات' }
 'Archivo de clientes: «Cambiar…» para elegir otro Excel.' = @{ en = 'Client file: «Change…» to choose another Excel file.'; fr = 'Fichier des clients : « Changer… » pour choisir un autre Excel.'; ar = 'ملف العملاء: «تغيير…» لاختيار ملف Excel آخر.' }
 'Copia externa: «Usar la nube de la empresa» la configura sola; también puedes elegir un USB o una carpeta.' = @{ en = 'External backup: «Use the company cloud» sets it up by itself; you can also choose a USB drive or a folder.'; fr = 'Copie externe : « Utiliser le cloud de l''entreprise » la configure seule ; vous pouvez aussi choisir une clé USB ou un dossier.'; ar = 'النسخة الخارجية: «استخدام سحابة الشركة» يضبطها تلقائيًا؛ ويمكنك أيضًا اختيار USB أو مجلد.' }
 'Empresa: nombre y logo. Equipo: los nombres de las compañeras. Preferencias: idioma (español, inglés, francés o árabe) y nacionalidad por defecto.' = @{ en = 'Company: name and logo. Team: colleagues'' names. Preferences: language (Spanish, English, French or Arabic) and default nationality.'; fr = 'Entreprise : nom et logo. Équipe : noms des collègues. Préférences : langue (espagnol, anglais, français ou arabe) et nationalité par défaut.'; ar = 'الشركة: الاسم والشعار. الفريق: أسماء الزملاء. التفضيلات: اللغة (الإسبانية أو الإنجليزية أو الفرنسية أو العربية) والجنسية الافتراضية.' }
 'Pulsa «Guardar ajustes» para aplicar los cambios.' = @{ en = 'Press «Save settings» to apply the changes.'; fr = 'Cliquez sur « Enregistrer les paramètres » pour appliquer les changements.'; ar = 'اضغط «حفظ الإعدادات» لتطبيق التغييرات.' }
 '10. Privacidad y registro' = @{ en = '10. Privacy and activity log'; fr = '10. Confidentialité et journal'; ar = '10. الخصوصية وسجل النشاط' }
 'Por seguridad, GESTEM anota quién añade o cambia cada cliente, quién crea recibos y cuándo se abre y se cierra el programa.' = @{ en = 'For security, GESTEM records who adds or changes each client, who creates receipts, and when the program is opened and closed.'; fr = 'Par sécurité, GESTEM note qui ajoute ou modifie chaque client, qui crée des reçus et quand le programme est ouvert et fermé.'; ar = 'لأسباب أمنية يسجّل GESTEM من يضيف أو يعدّل كل عميل، ومن ينشئ الإيصالات، ومتى يُفتح البرنامج ويُغلق.' }
 'Una copia de ese registro, sin nombres de clientes, se guarda junto a la copia externa. Los datos de los clientes nunca salen de vuestro PC ni de la nube de la empresa.' = @{ en = 'A copy of that log, without client names, is kept next to the external backup. Client data never leaves your PC or the company cloud.'; fr = 'Une copie de ce journal, sans noms de clients, est gardée avec la copie externe. Les données des clients ne quittent jamais votre PC ni le cloud de l''entreprise.'; ar = 'تُحفظ نسخة من هذا السجل، دون أسماء العملاء، بجانب النسخة الخارجية. بيانات العملاء لا تغادر حاسوبكم ولا سحابة الشركة أبدًا.' }
 '11. Actualizaciones y ayuda' = @{ en = '11. Updates and help'; fr = '11. Mises à jour et aide'; ar = '11. التحديثات والمساعدة' }
 'GESTEM se actualiza solo al abrirse cuando hay una versión nueva. También puedes pulsar «Buscar actualizaciones».' = @{ en = 'GESTEM updates itself when it opens if there is a new version. You can also press «Check for updates».'; fr = 'GESTEM se met à jour tout seul à l''ouverture s''il y a une nouvelle version. Vous pouvez aussi cliquer sur « Rechercher des mises à jour ».'; ar = 'يتحدّث GESTEM تلقائيًا عند الفتح إذا توفّر إصدار جديد. ويمكنك أيضًا الضغط على «البحث عن تحديثات».' }
 'Si algo no funciona, cierra GESTEM y vuelve a abrirlo. Si sigue igual, avisa a Yano.' = @{ en = 'If something does not work, close GESTEM and open it again. If it still fails, contact Yano.'; fr = 'Si quelque chose ne marche pas, fermez GESTEM et rouvrez-le. Si le problème continue, prévenez Yano.'; ar = 'إذا لم يعمل شيء فأغلق GESTEM وافتحه من جديد. وإذا استمرت المشكلة فتواصل مع Yano.' }
 'GESTEM MANAGER © 2026 Yano. Todos los derechos reservados. Uso bajo licencia (ver LICENCIA.txt).' = @{ en = 'GESTEM MANAGER © 2026 Yano. All rights reserved. Used under licence (see LICENCIA.txt).'; fr = 'GESTEM MANAGER © 2026 Yano. Tous droits réservés. Utilisé sous licence (voir LICENCIA.txt).'; ar = 'GESTEM MANAGER © 2026 Yano. جميع الحقوق محفوظة. يُستخدم بموجب ترخيص (انظر LICENCIA.txt).' }
 'Manual de uso' = @{ en = 'User manual'; fr = 'Mode d''emploi'; ar = 'دليل الاستخدام' }
 'GESTEM Manager · todo lo que necesitas para trabajar con el programa' = @{ en = 'GESTEM Manager · everything you need to work with the program'; fr = 'GESTEM Manager · tout ce qu''il faut pour travailler avec le programme'; ar = 'GESTEM Manager · كل ما تحتاجه للعمل بالبرنامج' }
}
foreach ($k in $script:HelpTr.Keys) { $script:AdmTr[$k] = $script:HelpTr[$k] }
# pantalla «Acerca de» en otros idiomas
$script:AboutTr = @{
 '© 2026 Yano · Todos los derechos reservados' = @{ en = '© 2026 Yano · All rights reserved'; fr = '© 2026 Yano · Tous droits réservés'; ar = '© 2026 Yano · جميع الحقوق محفوظة' }
 'Programa con licencia de uso para la empresa titular. Prohibida su copia, modificación, distribución o reventa sin permiso escrito del autor.' = @{ en = 'Licensed for use by the company that holds it. Copying, modifying, distributing or reselling it without the author''s written permission is prohibited.'; fr = 'Programme sous licence d''utilisation pour l''entreprise titulaire. Copie, modification, distribution ou revente interdites sans l''autorisation écrite de l''auteur.'; ar = 'برنامج مرخّص للاستخدام من قِبل الشركة صاحبة الترخيص. يُمنع نسخه أو تعديله أو توزيعه أو إعادة بيعه دون إذن مكتوب من المؤلف.' }
}
foreach ($k in $script:AboutTr.Keys) { $script:AdmTr[$k] = $script:AboutTr[$k] }
# texto del área de administrador en el idioma elegido (las claves son el texto en español)
function AdmT([string]$es) {
    $lang = "$($script:Cfg.Language)"
    if ($lang -and $lang -ne 'es' -and $script:AdmTr.ContainsKey($es)) { $v = $script:AdmTr[$es][$lang]; if ($v) { return $v } }
    return $es
}
$script:TrText = @{}; $script:TrTip = @{}
function Set-AdmTree($node) {
    if (-not ($node -is [Windows.DependencyObject])) { return }
    if ($node -is [Windows.Controls.TextBlock]) {
        if (-not $script:TrText.ContainsKey($node) -and $script:AdmTr.ContainsKey("$($node.Text)")) { $script:TrText[$node] = "$($node.Text)" }
        if ($script:TrText.ContainsKey($node)) { $node.Text = AdmT $script:TrText[$node] }
    } elseif ($node -is [Windows.Controls.ContentControl] -and $node.Content -is [string]) {
        if (-not $script:TrText.ContainsKey($node) -and $script:AdmTr.ContainsKey("$($node.Content)")) { $script:TrText[$node] = "$($node.Content)" }
        if ($script:TrText.ContainsKey($node)) { $node.Content = AdmT $script:TrText[$node] }
    }
    if ($node -is [Windows.FrameworkElement] -and $node.ToolTip -is [string]) {
        if (-not $script:TrTip.ContainsKey($node) -and $script:AdmTr.ContainsKey("$($node.ToolTip)")) { $script:TrTip[$node] = "$($node.ToolTip)" }
        if ($script:TrTip.ContainsKey($node)) { $node.ToolTip = AdmT $script:TrTip[$node] }
    }
    foreach ($c in [Windows.LogicalTreeHelper]::GetChildren($node)) { Set-AdmTree $c }
}
function Apply-AdminLanguage {
    foreach ($p in 'PanelPaused', 'PanelAdminLogin', 'PanelAdmin', 'PanelHelp', 'PanelAbout') { Set-AdmTree $script:UI[$p] }
}
# registro de actividad: traduce solo el principio conocido de cada línea
function Convert-LogText([string]$txt) {
    if ("$($script:Cfg.Language)" -eq 'es') { return $txt }
    foreach ($k in @($script:AdmTr.Keys | Sort-Object { $_.Length } -Descending)) {
        if ($k.Length -ge 8 -and $txt.StartsWith($k)) { return (AdmT $k) + $txt.Substring($k.Length) }
    }
    return $txt
}

function Load-Admin {
    if (-not (Test-Path -LiteralPath $script:AdminPath)) { return }
    try {
        $enc = [IO.File]::ReadAllBytes($script:AdminPath)
        $raw = [Text.Encoding]::UTF8.GetString([Security.Cryptography.ProtectedData]::Unprotect($enc, $null, 'CurrentUser'))
        $j = $raw | ConvertFrom-Json
        $script:Admin.Salt = "$($j.Salt)"; $script:Admin.Hash = "$($j.Hash)"; $script:Admin.IgnorePause = [bool]$j.IgnorePause; $script:Admin.OwnerPC = [bool]$j.OwnerPC
        $script:Admin.Locks = @{}
        if ($j.Locks) { foreach ($pr in $j.Locks.PSObject.Properties) { $script:Admin.Locks[$pr.Name] = [bool]$pr.Value } }
    } catch { Write-Log "admin: no se pudo leer ($($_.Exception.Message))" }
}
function Save-Admin {
    $json = $script:Admin | ConvertTo-Json -Depth 4
    $enc = [Security.Cryptography.ProtectedData]::Protect([Text.Encoding]::UTF8.GetBytes($json), $null, 'CurrentUser')
    if (Test-Path -LiteralPath $script:AdminPath) { (Get-Item -LiteralPath $script:AdminPath -Force).Attributes = 'Normal' }
    [IO.File]::WriteAllBytes($script:AdminPath, $enc)
    (Get-Item -LiteralPath $script:AdminPath -Force).Attributes = 'Hidden'
}
function Get-PwHash([string]$pw, [string]$salt) {
    $sha = [Security.Cryptography.SHA256]::Create()
    $b = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes("$salt|GESTEM|$pw"))
    return ([BitConverter]::ToString($b) -replace '-', '')
}
function Get-LockSet([string]$k) { return ($script:Admin.Locks.ContainsKey($k) -and $script:Admin.Locks[$k]) }
function Test-Locked([string]$k) { if ($script:IsOwnerPC) { return $false }; return (Get-LockSet $k) }
function Apply-Locks {
    $ui = $script:UI
    foreach ($k in 'File', 'Ext', 'Comp', 'Team', 'Pref', 'Help') {
        $ui["Sec$($k)Box"].Visibility = $(if ((Test-Locked $k) -and -not $script:IsAdmin) { 'Collapsed' } else { 'Visible' })
    }
    $hideAll = (Test-Locked 'All') -and -not $script:IsAdmin
    $ui.NavSettings.Visibility = $(if ($hideAll) { 'Collapsed' } else { 'Visible' })
    if ($hideAll -and $ui.PanelSettings.Visibility -eq 'Visible') { $ui.NavNew.IsChecked = $true }
}
function Show-AdminLogin {
    $ui = $script:UI
    $first = -not $script:Admin.Hash
    $ui.AdmLoginHint.Text = $(if ($first) { AdmT 'Primera vez: crea tu contraseña de administrador (mínimo 6 caracteres) y escríbela dos veces. Guárdala bien: no se puede recuperar.' } else { AdmT 'Escribe tu contraseña de administrador.' })
    $ui.AdmPw2.Visibility = $(if ($first) { 'Visible' } else { 'Collapsed' })
    $ui.AdmPw.Password = ''; $ui.AdmPw2.Password = ''; $ui.AdmLoginMsg.Text = ''
    Update-KeyHint
    $ui.PanelAdminLogin.Visibility = 'Visible'
    [void]$ui.AdmPw.Focus()
}
# aviso de Bloq Mayús y del teclado activo (AZERTY/QWERTY cambian la contraseña)
function Update-KeyHint {
    $ui = $script:UI
    $layout = ''
    try { $layout = [Windows.Forms.InputLanguage]::CurrentInputLanguage.LayoutName } catch { }
    if (-not $layout) { try { $layout = [Windows.Input.InputLanguageManager]::Current.CurrentInputLanguage.DisplayName } catch { } }
    $txt = $(if ($layout) { (AdmT 'Teclado: {0}') -f $layout } else { '' })
    $caps = $false
    try { $caps = [Windows.Input.Keyboard]::IsKeyToggled([Windows.Input.Key]::CapsLock) } catch { }
    if ($caps) {
        $ui.AdmKeyHint.Text = (AdmT 'Bloq Mayús está activado.') + $(if ($txt) { "   ·   $txt" } else { '' })
        $ui.AdmKeyHint.Foreground = Brush '#B42318'
    } else {
        $ui.AdmKeyHint.Text = $txt
        $ui.AdmKeyHint.Foreground = Brush '#8A7F81'
    }
}
function Invoke-AdminLogin {
    $ui = $script:UI
    $pw = $ui.AdmPw.Password
    if (-not $script:Admin.Hash) {
        if ($pw.Length -lt 6) { $ui.AdmLoginMsg.Text = AdmT 'Mínimo 6 caracteres.'; return }
        if ($pw -ne $ui.AdmPw2.Password) { $ui.AdmLoginMsg.Text = AdmT 'Las dos contraseñas no son iguales.'; return }
        $script:Admin.Salt = [guid]::NewGuid().ToString('N')
        $script:Admin.Hash = Get-PwHash $pw $script:Admin.Salt
        Save-Admin
    } elseif ((Get-PwHash $pw $script:Admin.Salt) -ne $script:Admin.Hash) {
        $ui.AdmLoginMsg.Text = AdmT 'Contraseña incorrecta.'
        Write-Log 'admin: intento de acceso fallido'
        return
    }
    $ui.AdmPw.Password = ''; $ui.AdmPw2.Password = ''
    $ui.PanelAdminLogin.Visibility = 'Collapsed'
    $script:IsAdmin = $true
    Write-Log 'admin: acceso'
    Apply-Locks
    Open-AdminPanel
}
function Open-AdminPanel {
    $ui = $script:UI
    foreach ($k in $script:LockKeys) { $ui["AdmLock$k"].IsChecked = (Get-LockSet $k) }
    $ui.AdmIgnorePause.IsChecked = [bool]$script:Admin.IgnorePause
    $ui.AdmOwnerPC.IsChecked = [bool]$script:IsOwnerPC
    $ui.AdmLocksMsg.Text = ''; $ui.AdmPwMsg.Text = ''
    $ui.AdmInfo.Text = (AdmT 'Versión {0} · datos: {1} · programa: {2}') -f $script:AppVersion, $script:DataDir, $script:AppDir
    Update-RemoteText
    Fill-AdminLog
    $ui.PanelAdmin.Visibility = 'Visible'
}
function Fill-AdminLog {
    $ui = $script:UI
    $items = New-Object System.Collections.Generic.List[object]
    $nAdm = 0; $nTeam = 0; $nErr = 0
    $onlyErr = ($script:AdmLogFilter -eq 'errors')
    if (Test-Path -LiteralPath $script:LogPath) {
        $lines = @(Get-Content -LiteralPath $script:LogPath -Encoding UTF8 -Tail 3000)
        [array]::Reverse($lines)
        foreach ($ln in $lines) {
            # los cierres inesperados se guardan con otro formato: «fecha  FATAL  mensaje»
            if ($ln -match '^(\d{4})-(\d\d)-(\d\d) (\d\d:\d\d):\d\d  FATAL  (.+)$') { $ln = "$($Matches[1])-$($Matches[2])-$($Matches[3]) $($Matches[4]):00  [] fatal: $($Matches[5])" }
            if ($ln -match '^(\d{4})-(\d\d)-(\d\d) (\d\d:\d\d):\d\d  \[([^\]]*)\] (.+)$') {
                $txt = $Matches[6]; $user = $Matches[5]
                $time = "$($Matches[3])/$($Matches[2])/$($Matches[1]) $($Matches[4])"    # se guarda ya: los -match de abajo cambian $Matches
                $isErr = Test-ErrorLogLine $txt
                if ($isErr) { $nErr++ }
                if ($onlyErr) {
                    if (-not $isErr -or $items.Count -ge 400) { continue }
                    $items.Add([pscustomobject]@{
                        Time = $time
                        User = $(if ($user -eq 'ADMIN') { AdmT 'Admin' } elseif ($user) { $user } else { AdmT 'Sistema' })
                        UserBg = '#FDECEC'; UserFg = '#B42318'
                        Text = $txt; Fg = '#B42318'
                    })
                    continue
                }
                if ($txt -match '^(Load-Data|Save-FromForm|clic |DEBUG|pestaña no leída|aviso mostrado: |fatal: )') { continue }
                $isAdm = ($user -eq 'ADMIN')
                $kind = $(if ($isAdm) { 'admin' } elseif ($user) { 'team' } else { 'system' })
                if ($kind -eq 'admin') { $nAdm++ } elseif ($kind -eq 'team') { $nTeam++ }
                if ($script:AdmLogFilter -eq 'team' -and $kind -ne 'team') { continue }
                if ($script:AdmLogFilter -eq 'admin' -and $kind -ne 'admin') { continue }
                $fg = '#2B2526'
                if ($txt -match 'fallid|error|no se pudo|incorrect') { $fg = '#B42318' }
                elseif ($txt -match '^admin') { $fg = '#9E1820' }
                elseif ($txt -match '^(añadió|editó)') { $fg = '#1E1A1B' }
                elseif ($txt -match '^(copia externa|sesión)') { $fg = '#6E6466' }
                $items.Add([pscustomobject]@{
                    Time = $time
                    User = $(switch ($kind) { 'admin' { AdmT 'Admin' } 'team' { $user } default { AdmT 'Sistema' } })
                    UserBg = $(switch ($kind) { 'admin' { '#9E1820' } 'team' { '#F3EFEF' } default { '#FFFFFF' } })
                    UserFg = $(switch ($kind) { 'admin' { '#FFFFFF' } 'team' { '#2B2526' } default { '#A69C9E' } })
                    Text = (Convert-LogText $txt); Fg = $fg
                })
                if ($items.Count -ge 400) { continue }
            }
        }
    }
    $ui.AdmLog.ItemsSource = $items
    if ($onlyErr) { $ui.AdmLogTitle.Text = (AdmT 'Errores · {0} en este PC (los más recientes arriba)') -f $nErr }
    else { $ui.AdmLogTitle.Text = (AdmT 'Actividad · {0} del equipo · {1} del admin') -f $nTeam, $nAdm }
    if ($ui.AdmLogErr) { $ui.AdmLogErr.Content = $(if ($nErr) { (AdmT 'Errores') + " ($nErr)" } else { AdmT 'Errores' }) }
}
function Test-ErrorLogLine([string]$txt) {
    # líneas del registro que son fallos (también los que no enseñan ningún aviso)
    if ($txt -match '^(fatal: |UI ERROR|aviso mostrado: |Load-Data: |Save-FromForm: |Guardar ERROR|Excel no actualizado|sincronizar con el Excel: |actualizar desde el Excel: |actualización: |pestaña no leída|listas \(|importar \(|config ilegible|accesos directos: |icono: |logo: |logo no válido|recibo: |exportar PDF: |calendario: |copia externa fallida|copia de seguridad fallida|admin: no se pudo|DEBUG (error|editar|conflicto|actualización automática): )') { return $true }
    return $false
}
function Show-AdminTab([string]$name) {
    $ui = $script:UI
    foreach ($p in 'AdmPageLocks', 'AdmPageLog', 'AdmPageTools') { $ui[$p].Visibility = $(if ($p -eq $name) { 'Visible' } else { 'Collapsed' }) }
}
function Save-AdminLocks {
    $ui = $script:UI
    foreach ($k in $script:LockKeys) { $script:Admin.Locks[$k] = [bool]$ui["AdmLock$k"].IsChecked }
    $script:Admin.IgnorePause = [bool]$ui.AdmIgnorePause.IsChecked
    $script:Admin.OwnerPC = [bool]$ui.AdmOwnerPC.IsChecked
    $script:IsOwnerPC = Test-OwnerPC
    Save-Admin
    Apply-Locks
    Write-Log ('admin: bloqueos guardados (' + (@($script:LockKeys | Where-Object { Get-LockSet $_ }) -join ', ') + ')')
    $ui.AdmLocksMsg.Text = AdmT '✓ Guardado'
}
function Change-AdminPw {
    $ui = $script:UI
    $a = $ui.AdmNewPw.Password; $b = $ui.AdmNewPw2.Password
    if ($a.Length -lt 6) { $ui.AdmPwMsg.Foreground = Brush '#B42318'; $ui.AdmPwMsg.Text = AdmT 'Mínimo 6 caracteres.'; return }
    if ($a -ne $b) { $ui.AdmPwMsg.Foreground = Brush '#B42318'; $ui.AdmPwMsg.Text = AdmT 'Las dos contraseñas no son iguales.'; return }
    $script:Admin.Salt = [guid]::NewGuid().ToString('N')
    $script:Admin.Hash = Get-PwHash $a $script:Admin.Salt
    Save-Admin
    $ui.AdmNewPw.Password = ''; $ui.AdmNewPw2.Password = ''
    $ui.AdmPwMsg.Foreground = Brush '#1B7F45'; $ui.AdmPwMsg.Text = AdmT '✓ Contraseña cambiada'
    Write-Log 'admin: contraseña cambiada'
}
function Close-Admin {
    Write-Log 'admin: salida'
    $script:IsAdmin = $false
    $script:UI.PanelAdmin.Visibility = 'Collapsed'
    Apply-Locks
}
# --- control remoto: GESTEM solo LEE control.json de GitHub; nunca envía datos
$script:RemotePath = Join-Path $script:DataDir 'control.cache'
$script:Remote = $null
function Check-Remote {
    $j = $null
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $r = Invoke-WebRequest -UseBasicParsing -Uri "$($script:UpdateBase)/control.json?t=$([DateTime]::UtcNow.Ticks)" -TimeoutSec 5
        $j = $r.Content | ConvertFrom-Json
        Set-Content -LiteralPath $script:RemotePath -Value $r.Content -Encoding UTF8
    } catch {
        $code = $null; try { $code = [int]$_.Exception.Response.StatusCode } catch {}
        if ($code -eq 404) { if (Test-Path -LiteralPath $script:RemotePath) { Remove-Item -LiteralPath $script:RemotePath -Force } }
        elseif (Test-Path -LiteralPath $script:RemotePath) { try { $j = Get-Content -Raw -LiteralPath $script:RemotePath -Encoding UTF8 | ConvertFrom-Json } catch {} }
    }
    $script:Remote = $j
    if (-not $j) { return }
    $msg = "$($j.mensaje)".Trim()
    if ($msg) {
        $seenPath = Join-Path $script:DataDir 'control.visto'
        $seen = ''; if (Test-Path -LiteralPath $seenPath) { $seen = (Get-Content -Raw -LiteralPath $seenPath -Encoding UTF8).Trim() }
        if ($seen -ne $msg) { Show-Info $msg; Set-Content -LiteralPath $seenPath -Value $msg -Encoding UTF8 }
    }
    Apply-Pause
}
function Apply-Pause {
    $ui = $script:UI
    $j = $script:Remote
    $paused = ($j -and [bool]$j.pausa -and -not $script:Admin.IgnorePause -and -not $script:IsAdmin -and -not $script:IsOwnerPC)
    if ($paused) {
        $ui.PausedText.Text = $(if ("$($j.motivo)".Trim()) { "$($j.motivo)" } else { AdmT 'El programa está en pausa temporalmente.' })
        $ui.PanelPaused.Visibility = 'Visible'
        Write-Log 'pausa remota activa'
    } else { $ui.PanelPaused.Visibility = 'Collapsed' }
}
function Update-RemoteText {
    $j = $script:Remote
    $base = AdmT 'GESTEM lee el archivo control.json de tu repositorio de GitHub al abrirse (nunca envía datos). Campos: «mensaje» (texto que verán una vez), «pausa» (true/false) y «motivo».'
    if (-not $j) { $script:UI.AdmRemoteTxt.Text = "$base`n`n" + (AdmT 'Estado: no hay control.json o no se pudo leer → todo normal.') ; return }
    $m = $(if ("$($j.mensaje)".Trim()) { "«$($j.mensaje)»" } else { AdmT '(vacío)' })
    $pz = $(if ([bool]$j.pausa) { AdmT 'SÍ' } else { AdmT 'no' })
    $script:UI.AdmRemoteTxt.Text = "$base`n`n" + ((AdmT 'Estado actual → mensaje: {0}   ·   pausa: {1}') -f $m, $pz)
}

# =====================================================================
#  SESIÓN (quién está usando el programa)
# =====================================================================
function Show-Session {
    $ui = $script:UI
    $ui.SessTitle.Text = L 'sessTitle'; $ui.SessHint.Text = L 'sessHint'
    $ui.SessGo.Content = L 'sessGo'; $ui.SessNewHint.Text = L 'sessNew'
    $ui.SessNames.Children.Clear()
    $people = @($script:Cfg.Colleagues | Where-Object { "$_".Trim() })
    if ($script:Cfg.LastUser) { $people = @($script:Cfg.LastUser) + @($people | Where-Object { $_ -ne $script:Cfg.LastUser }) }
    foreach ($p in $people) {
        $b = New-Object Windows.Controls.Button
        $b.Style = $script:Win.FindResource('Btn')
        $b.Content = $p; $b.Tag = $p
        $b.Height = 42; $b.FontSize = 15; $b.Padding = '20,0'; $b.Margin = '0,0,8,8'
        if ($p -eq $script:Cfg.LastUser) { $b.FontWeight = [Windows.FontWeights]::SemiBold }
        $b.Add_Click({ Start-Session "$($this.Tag)" })
        [void]$ui.SessNames.Children.Add($b)
    }
    $ui.SessNew.Text = ''
    $ui.SessNewHint.Visibility = 'Visible'
    $ui.PanelSession.Visibility = 'Visible'
}

function Start-Session([string]$who) {
    $who = "$who".Trim().ToUpper()
    if (-not $who) { $who = "$($script:Cfg.LastUser)".Trim() }
    if (-not $who) { return }
    $script:User = $who
    $script:Cfg.LastUser = $who
    $people = @($script:Cfg.Colleagues | Where-Object { "$_".Trim() })
    if ($people -notcontains $who) { $people += $who }
    $script:Cfg.Colleagues = $people
    Save-Config
    $known = @($script:Lists['colleagues'])
    if ($known -notcontains $who) { $script:Lists['colleagues'] = @($known + $who | Where-Object { $_ }) }
    $script:UI.PanelSession.Visibility = 'Collapsed'
    $script:UI.UserBtn.Content = ((L 'userTip') -f $who)
    Write-Log "sesión iniciada"
    if ("$($script:Cfg.ExtBackupDir)".Trim() -and (Test-Path -LiteralPath $script:Cfg.ExtBackupDir)) { Export-ActivityLog $script:Cfg.ExtBackupDir }   # así se ve enseguida quién lo abrió
    if (-not $script:EditRow) { Build-Form; Reset-Form }
    Refresh-ClientGrid
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
# icono de GESTEM (rojo y blanco) incrustado: ventana, «Acerca de», logo por defecto y acceso directo
$script:AppIconB64 = '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'
$script:AppLogo = $null
try {
    $icoBytes = [Convert]::FromBase64String($script:AppIconB64)
    $dec = [Windows.Media.Imaging.BitmapDecoder]::Create((New-Object IO.MemoryStream(,$icoBytes)), [Windows.Media.Imaging.BitmapCreateOptions]::None, [Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
    $script:AppLogo = $dec.Frames | Sort-Object { $_.PixelWidth } -Descending | Select-Object -First 1
    $script:Win.Icon = $script:AppLogo
    $iconPath = Join-Path $script:AppDir 'gestem.ico'   # el acceso directo usa este archivo: se renueva si es el antiguo
    $old = $null; if (Test-Path -LiteralPath $iconPath) { $old = [IO.File]::ReadAllBytes($iconPath) }
    if (-not $old -or $old.Length -ne $icoBytes.Length) { [IO.File]::WriteAllBytes($iconPath, $icoBytes) }
} catch { Write-Log "icono: $($_.Exception.Message)" }

$script:UI = @{}
foreach ($n in @('TramiteLabel','TramiteBox','NavNew','NavClients','NavSettings','NavAbout','StatusText','FileText',
                 'PanelForm','FormTitle','FormGrid','CardTag','CardId','CardName','CardSub','CardTipo','CardStatusB','CardStatus',
                 'SumPriceL','SumPrice','SumPaidL','SumPaid','SumPendL','SumPend','DupWarn','SaveBtn','ClearBtn',
                 'PanelClients','RefreshBtn','OpenExcelBtn','EditBtn','SearchBox','SearchHint','FilterBox','OnlyDebt','ClientsGrid',
                 'PanelSettings','SetTitle','WelcomeText','SetFileL','ChangeFileBtn','FilePathBox','SetLangL','LangBox','SetNatL','NatBox',
                 'BackupChk','OpenAfterChk','SetNote','SaveSettingsBtn','BackupsBtn','UpdateBtn',
                 'PanelAbout','AboutLogo','AboutVer','AboutText','AboutBy','EasterEgg',
                 'UserBtn','OnlyMine','SetCollL','ColleaguesBox','PanelSession','SessTitle','SessHint','SessNames','SessGo','SessNew','SessNewHint',
                 'HelpBtn','HelpClose','PanelHelp','SetSub',
                 'SecExtT','SecExtD','ImportXlBtn','DataNote','ExtDirL','ExtDirBox','ExtDirBtn','ExtClearBtn','ExtOneDriveBtn','ExtNowBtn','ExtOpenBtn','ExtLastTxt','ExtNote',
                 'SecFileT','SecFileD','SecCompT','SecCompD','SecTeamT','SecTeamD','SecPrefT','SecPrefD','SecHelpT','SecHelpD',
                 'ChartDay','ChartWeek','ChartMonth','ChartTitle','ChartSub','ChartLegPaid','ChartLegPend','ChartGrid',
                 'ChartCanvas','ChartLabels','ChartDelta','ChartDeltaB','ChartLegMoney',
                 'DocsBtn','ReceiptBtn','ExportBtn','ExpPdf','ExpCsv','CmEdit','CmReceipt','CmDocs','DocsCol','PanelDocs','DocsCard','DocsClose','DocsTitle','DocsSub','DocsOpenBtn','DocsRemoveBtn','DocsFolderBtn','DocsAddBtn','DocsCount','DocsMsg','DocsList','DocsEmpty','DocsEmptyT','DocsDrop','DocsDropT',
                 'PanelPaused','PausedText','PanelAdminLogin','AdmLoginHint','AdmKeyHint','AdmPw','AdmPw2','AdmLoginMsg','AdmLoginCancel','AdmLoginBtn','PanelAdmin','AdmLogout','AdmClose','AdmTabLocks','AdmTabLog','AdmLogErr','AdmTabTools','AdmPageLocks','AdmPageLog','AdmPageTools','AdmLockFile','AdmLockExt','AdmLockComp','AdmLockTeam','AdmLockPref','AdmLockHelp','AdmLockAll','AdmIgnorePause','AdmOwnerPC','AdmSaveLocks','AdmLocksMsg','AdmLogRefresh','AdmLogAll','AdmLogTeam','AdmLogAdmin','AdmLogTitle','AdmLog','AdmRemoteTxt','AdmUpdate','AdmOpenData','AdmOpenApp','AdmShowAll','AdmInfo','AdmNewPw','AdmNewPw2','AdmChangePw','AdmPwMsg','SecFileBox','SecExtBox','SecCompBox','SecTeamBox','SecPrefBox','SecHelpBox',
                 'TopBar','StatsScope','StatsAllBtn','StatsAllTxt','NavStats','PanelStats','StatsWho','StatsPeriod','StatsRefresh','StatsTitle','StatsSub','StatsTable','StatsGrid',
                 'KpiBillL','KpiBill','KpiBillF','KpiPaidL','KpiPaid','KpiPaidTrack','KpiPaidF','KpiPendL','KpiPend','KpiPendF','KpiNetL','KpiNet','KpiNetF',
                 'SideLogo','SideCompany','SessLogo','SessCompany','AboutImg','AboutCompany','SetCompanyL','CompanyBox','SetLogoL','LogoPreview','LogoBtn','LogoClearBtn')) {
    $script:UI[$n] = $script:Win.FindName($n)
}
# menús y columnas no siempre se encuentran por nombre: se toman directamente
if (-not $script:UI.ExpPdf) { $script:UI.ExpPdf = $script:UI.ExportBtn.ContextMenu.Items[0]; $script:UI.ExpCsv = $script:UI.ExportBtn.ContextMenu.Items[1] }
if (-not $script:UI.CmEdit) { $cm = $script:UI.ClientsGrid.ContextMenu; $script:UI.CmEdit = $cm.Items[0]; $script:UI.CmReceipt = $cm.Items[1]; $script:UI.CmDocs = $cm.Items[2] }
if (-not $script:UI.DocsCol) { $script:UI.DocsCol = $script:UI.ClientsGrid.Columns[$script:UI.ClientsGrid.Columns.Count - 1] }
$script:UI.AboutImg.Source = $script:AppLogo
$script:PendingLogo = ''
Import-DefaultBranding
Repair-Shortcuts
Apply-Branding

$ui = $script:UI

# navegación
$ui.NavNew.Add_Checked({ Show-Panel 'PanelForm' })
$ui.NavClients.Add_Checked({ Show-Panel 'PanelClients'; Refresh-ClientGrid; $script:UI.SearchBox.Focus() | Out-Null })
$ui.HelpBtn.Add_Click({ $script:UI.PanelHelp.Visibility = 'Visible' })
$ui.HelpClose.Add_Click({ $script:UI.PanelHelp.Visibility = 'Collapsed' })
$ui.PanelHelp.Add_MouseLeftButtonDown({ if ($_.OriginalSource -eq $script:UI.PanelHelp) { $script:UI.PanelHelp.Visibility = 'Collapsed' } })
$ui.ChartCanvas.Add_SizeChanged({ Draw-ChartLine })
foreach ($cb in @($ui.ChartDay, $ui.ChartWeek, $ui.ChartMonth)) { $cb.Add_Checked({ if ($script:UI.PanelStats.Visibility -eq 'Visible') { Update-Chart } }) }
$ui.NavStats.Add_Checked({ Show-Panel 'PanelStats'; Fill-StatsFilters; Update-Stats })
$ui.StatsScope.Add_SelectionChanged({
    if ($script:StatsFilling) { return }
    $ix = $script:UI.StatsScope.SelectedIndex
    $script:StatsScopeId = $(if ($ix -gt 0) { $script:Tramites[$ix - 1].Id } else { $null })
    if ($script:UI.PanelStats.Visibility -eq 'Visible') { Update-Stats }
})
$ui.ExtDirBtn.Add_Click({ Choose-ExtFolder })
$ui.ExtClearBtn.Add_Click({ $script:UI.ExtDirBox.Text = ''; Update-ExtBackupInfo })
$ui.ExtOneDriveBtn.Add_Click({ $c = Find-CompanyCloud; if ($c) { $script:UI.ExtDirBox.Text = $c; Update-ExtBackupInfo } else { Show-Error (L 'extNoCloud') } })
$ui.ExtNowBtn.Add_Click({
    $old = $script:Cfg.ExtBackupDir; $script:Cfg.ExtBackupDir = "$($script:UI.ExtDirBox.Text)".Trim()
    $script:Win.Cursor = [Windows.Input.Cursors]::Wait
    $ok = Backup-External
    $script:Win.Cursor = $null
    if (-not $ok) { $script:Cfg.ExtBackupDir = $old } else { Save-Config; Set-Status (L 'extDone') }
    Update-ExtBackupInfo
})
$ui.ExtOpenBtn.Add_Click({ $d = "$($script:UI.ExtDirBox.Text)".Trim(); if ($d) { if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Force -Path $d | Out-Null }; Start-Process explorer.exe -ArgumentList "`"$d`"" } })
# --- administrador
$ui.AdmLoginBtn.Add_Click({ Invoke-AdminLogin })
$ui.AdmPw.Add_KeyDown({ if ($_.Key -eq 'Return') { Invoke-AdminLogin } })
$ui.AdmPw2.Add_KeyDown({ if ($_.Key -eq 'Return') { Invoke-AdminLogin } })
foreach ($pb in $ui.AdmPw, $ui.AdmPw2, $ui.AdmNewPw, $ui.AdmNewPw2) { $pb.Add_KeyUp({ Update-KeyHint }); $pb.Add_GotKeyboardFocus({ Update-KeyHint }) }
$ui.AdmLoginCancel.Add_Click({ $script:UI.PanelAdminLogin.Visibility = 'Collapsed' })
$ui.AdmClose.Add_Click({ $script:UI.PanelAdmin.Visibility = 'Collapsed' })
$ui.AdmLogout.Add_Click({ Close-Admin; Apply-Pause })
$ui.AdmTabLocks.Add_Checked({ Show-AdminTab 'AdmPageLocks' })
$ui.AdmTabLog.Add_Checked({ Fill-AdminLog; Show-AdminTab 'AdmPageLog' })
$ui.AdmTabTools.Add_Checked({ Update-RemoteText; Show-AdminTab 'AdmPageTools' })
$ui.AdmSaveLocks.Add_Click({ Save-AdminLocks })
$ui.AdmLogRefresh.Add_Click({ Fill-AdminLog })
$ui.AdmLogAll.Add_Checked({ $script:AdmLogFilter = 'all'; Fill-AdminLog })
$ui.AdmLogTeam.Add_Checked({ $script:AdmLogFilter = 'team'; Fill-AdminLog })
$ui.AdmLogAdmin.Add_Checked({ $script:AdmLogFilter = 'admin'; Fill-AdminLog })
$ui.AdmLogErr.Add_Checked({ $script:AdmLogFilter = 'errors'; Fill-AdminLog })
$ui.AdmChangePw.Add_Click({ Change-AdminPw })
$ui.AdmUpdate.Add_Click({ $script:UI.PanelAdmin.Visibility = 'Collapsed'; Check-Update -Manual })
$ui.AdmOpenData.Add_Click({ Start-Process explorer.exe -ArgumentList "`"$($script:DataDir)`"" })
$ui.AdmOpenApp.Add_Click({ Start-Process explorer.exe -ArgumentList "`"$($script:AppDir)`"" })
$ui.AdmShowAll.Add_Click({ $script:UI.PanelAdmin.Visibility = 'Collapsed'; Load-SettingsPanel; $script:UI.NavSettings.IsChecked = $true; Show-Panel 'PanelSettings' })
Load-Admin
$script:IsOwnerPC = Test-OwnerPC
Apply-Locks
$ui.StatsAllBtn.Add_Click({ $script:StatsScopeId = $null; Fill-StatsFilters; Update-Stats })
$ui.StatsWho.Add_SelectionChanged({ if ($script:UI.PanelStats.Visibility -eq 'Visible') { Update-Stats } })
$ui.StatsPeriod.Add_SelectionChanged({ if ($script:UI.PanelStats.Visibility -eq 'Visible') { Update-Stats } })
$ui.StatsRefresh.Add_Click({ $script:Win.Cursor = [Windows.Input.Cursors]::Wait; [void](Load-Data -All); Fill-StatsFilters; Update-Stats; $script:Win.Cursor = $null })
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
$ui.DocsBtn.Add_Click({ $r = Get-SelectedClient; if ($r) { Open-DocsPanel $r } })
$ui.ReceiptBtn.Add_Click({ $r = Get-SelectedClient; if ($r) { New-ClientReceipt $r } })
$ui.ExportBtn.Add_Click({ $m = $this.ContextMenu; $m.PlacementTarget = $this; $m.Placement = 'Bottom'; $m.IsOpen = $true })
$ui.ExpPdf.Add_Click({ Export-ClientsPdf })
$ui.ExpCsv.Add_Click({ Export-ClientsCsv })
$ui.CmEdit.Add_Click({ Edit-Selected })
$ui.CmReceipt.Add_Click({ $r = Get-SelectedClient; if ($r) { New-ClientReceipt $r } })
$ui.CmDocs.Add_Click({ $r = Get-SelectedClient; if ($r) { Open-DocsPanel $r } })
# clic derecho: selecciona la fila que hay debajo antes de abrir el menú
$ui.ClientsGrid.Add_PreviewMouseRightButtonDown({
    $o = $_.OriginalSource
    while ($o -and -not ($o -is [Windows.Controls.DataGridRow])) { $o = [Windows.Media.VisualTreeHelper]::GetParent($o) }
    if ($o) { $o.IsSelected = $true }
})
$ui.DocsClose.Add_Click({ $script:UI.PanelDocs.Visibility = 'Collapsed' })
$ui.DocsAddBtn.Add_Click({ Add-DocsFromDialog })
$ui.DocsOpenBtn.Add_Click({ if ($script:UI.DocsList.SelectedItem) { Open-SelectedDoc } else { Show-Info (L 'docsPick') } })
$ui.DocsRemoveBtn.Add_Click({ Remove-SelectedDoc })
$ui.DocsFolderBtn.Add_Click({ Open-DocsFolder })
$ui.DocsList.Add_MouseDoubleClick({ Open-SelectedDoc })
foreach ($dz in $ui.PanelDocs, $ui.DocsCard) {
    $dz.Add_DragEnter({ if ($_.Data.GetDataPresent([Windows.DataFormats]::FileDrop)) { $_.Effects = 'Copy'; $script:UI.DocsDrop.Visibility = 'Visible' } else { $_.Effects = 'None' }; $_.Handled = $true })
    $dz.Add_DragOver({ $_.Effects = $(if ($_.Data.GetDataPresent([Windows.DataFormats]::FileDrop)) { 'Copy' } else { 'None' }); $_.Handled = $true })
    $dz.Add_DragLeave({ $script:UI.DocsDrop.Visibility = 'Collapsed' })
    $dz.Add_Drop({
        $script:UI.DocsDrop.Visibility = 'Collapsed'
        if ($_.Data.GetDataPresent([Windows.DataFormats]::FileDrop)) { Add-ClientDocs @($_.Data.GetData([Windows.DataFormats]::FileDrop)) }
        $_.Handled = $true
    })
}
$ui.RefreshBtn.Add_Click({ $script:Win.Cursor = [Windows.Input.Cursors]::Wait; [void](Refresh-FromExcel); $script:Win.Cursor = $null })
$ui.ImportXlBtn.Add_Click({ Import-AllFromExcel })
$ui.OpenExcelBtn.Add_Click({
    $sel = $script:UI.ClientsGrid.SelectedItem
    if ($sel) { Open-ExcelWindow $sel.Row } else { Open-ExcelWindow }
})

# ajustes
$ui.ChangeFileBtn.Add_Click({ [void](Choose-ExcelFile) })
$ui.LogoBtn.Add_Click({ Choose-Logo })
$ui.LogoClearBtn.Add_Click({ $script:PendingLogo = ''; $script:UI.LogoPreview.Fill = (Get-LogoBrush '') })
$ui.SaveSettingsBtn.Add_Click({ Save-Settings })
$ui.BackupsBtn.Add_Click({ Start-Process explorer.exe $script:BackupDir })
$ui.UpdateBtn.Add_Click({ $script:Win.Cursor = [Windows.Input.Cursors]::Wait; Check-Update -Manual; $script:Win.Cursor = $null })

# sesión
$ui.UserBtn.Add_Click({ Show-Session })
$ui.SessGo.Add_Click({ Start-Session $script:UI.SessNew.Text })
$ui.SessNew.Add_KeyDown({ if ($_.Key -eq 'Return') { Start-Session $script:UI.SessNew.Text } })
$ui.SessNew.Add_TextChanged({ $script:UI.SessNewHint.Visibility = $(if ($script:UI.SessNew.Text) { 'Collapsed' } else { 'Visible' }) })
$ui.OnlyMine.Add_Click({ Refresh-ClientGrid })

# trámite
$ui.TramiteBox.Add_SelectionChanged({
    $i = $script:UI.TramiteBox.SelectedIndex
    if ($i -ge 0 -and $script:Tramites[$i].Id -ne $script:T.Id) {
        # mientras se carga un trámite no se puede cambiar a otro (si no, la ficha y el trámite se mezclan)
        if ($script:Switching) {
            $script:UI.TramiteBox.SelectedIndex = [array]::IndexOf(@($script:Tramites | ForEach-Object { $_.Id }), $script:T.Id)
            return
        }
        $script:Switching = $true; $wasBusy = $script:Busy; $script:Busy = $true
        $script:Win.Cursor = [Windows.Input.Cursors]::Wait
        try {
            $script:T = $script:Tramites[$i]; $script:Cfg.Tramite = $script:T.Id; Save-Config
            $script:Rows = $null; $script:UI.ClientsGrid.ItemsSource = $null
            if (Get-CachedData) { Apply-Language } else { Build-Form; Reset-Form }
            if ($script:UI.PanelStats.Visibility -eq 'Visible') { $script:StatsScopeId = $script:T.Id; Fill-StatsFilters; Update-Stats }
        } finally { $script:Switching = $false; $script:Busy = $wasBusy; $script:Win.Cursor = $null }
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
    $key = $(if ($_.Key -eq 'System') { $_.SystemKey } else { $_.Key })
    $mods = [int][Windows.Input.Keyboard]::Modifiers
    $want = [int][Windows.Input.ModifierKeys]::Control -bor [int][Windows.Input.ModifierKeys]::Shift -bor [int][Windows.Input.ModifierKeys]::Alt
    if ($key -eq 'A' -and ($mods -band $want) -eq $want) {
        if ($script:IsAdmin) { Open-AdminPanel } else { Show-AdminLogin }
        $_.Handled = $true; return
    }
    if ($_.Key -eq 'S' -and [Windows.Input.Keyboard]::Modifiers -eq 'Control' -and $script:UI.PanelForm.Visibility -eq 'Visible') { Save-FromForm; $_.Handled = $true }
})

$script:Win.Dispatcher.Add_UnhandledException({
    $m = "$($_.Exception.Message)`r`n$($_.Exception.InnerException)"
    Write-Log "UI ERROR: $($_.Exception.ToString())"
    [void][Windows.MessageBox]::Show($m, 'GESTEM MANAGER - error', 'OK', 'Error')
    $_.Handled = $true
})

Apply-Language

# copia al cerrar
$script:Win.Add_Closing({
    if ($script:Restarting) { return }
    # errores silenciosos de la sesión (botones, eventos) al registro, para poder revisarlos
    try { foreach ($er in @($Error | Where-Object { "$($_.Exception.Message)" -notmatch 'GetActiveObject|MK_E_UNAVAILABLE' } | Select-Object -First 15)) { $pos = ''; if ($er.InvocationInfo) { $pos = " (línea $($er.InvocationInfo.ScriptLineNumber))" }; Write-Log ("DEBUG error: $($er.Exception.Message)$pos") } } catch { }
    if ($script:User) { Write-Log 'cerró GESTEM' }
    if ($script:Cfg.ExtBackupDir -and $script:Cfg.ExcelPath) { try { [void](Backup-External -Quiet) } catch {} }
})

$script:Win.Add_ContentRendered({
    Check-Update
    if ($script:Restarting) { return }
    Check-Remote
    $flag = Join-Path $script:DataDir 'actualizado.txt'
    if (Test-Path $flag) { $v = (Get-Content -Raw $flag).Trim(); Remove-Item $flag -Force; Show-Info ((L 'updated') -f $v) }
    if (-not (Test-ExcelInstalled)) { Show-Error (L 'errNoExcel') }
    if (-not $script:Cfg.ExcelPath -or -not (Test-Path -LiteralPath $script:Cfg.ExcelPath)) {
        Set-Status (L 'noFile')
        $script:UI.NavSettings.IsChecked = $true
        return
    }
    $script:Win.Cursor = [Windows.Input.Cursors]::Wait
    try { [void](Sync-HandmadeTabs -Force) } catch { Write-Log "pestañas propias: $($_.Exception.Message)" }
    if (Load-Data -All) {
        Apply-Language
        if ((Get-PendingSyncCount) -gt 0) { if ((Sync-PendingToExcel) -gt 0) { [void](Load-Data -All) } }
        $script:LiveTimer.Start()
    }
    $script:Win.Cursor = $null
    # primera vez sin carpeta externa: se usa sola la nube de la empresa si este PC la tiene
    if (-not $script:Cfg.ExtBackupDir) {
        $c = Find-CompanyCloud
        if ($c) { $script:Cfg.ExtBackupDir = $c; Save-Config; Write-Log "copia externa configurada automáticamente: $c" }
    }
    if ($script:Cfg.ExtBackupDir) { [void](Backup-External -Quiet) }    # copia al abrir
    Show-Session
})

try {
    [void]$script:Win.ShowDialog()
} catch {
    Write-Log "fatal: $($_.Exception.ToString())"
    [void][Windows.MessageBox]::Show("$($_.Exception.Message)", $script:AppName, 'OK', 'Error')
}
