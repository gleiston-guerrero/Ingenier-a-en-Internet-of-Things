// Genera el borrador del proyecto de carrera en formato Word.
// Uso: NODE_PATH=$(npm root -g) node src/generar_borrador.js
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType, LevelFormat,
  PageNumber, PageBreak,
} = require('docx');
const M = require('./malla.js');

const FONT = 'Calibri';
const COLOR = '1F4E79';
const W = 9026; // ancho útil A4 con márgenes de 2,54 cm

// ---------- helpers ----------
// Texto con **negrita** y [[campo por completar]] resaltado en amarillo.
function runs(text, base = {}) {
  const out = [];
  const re = /(\*\*.+?\*\*|\[\[.+?\]\])/g;
  let last = 0; let m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), font: FONT, ...base }));
    const tok = m[0];
    if (tok.startsWith('**')) out.push(new TextRun({ text: tok.slice(2, -2), bold: true, font: FONT, ...base }));
    else out.push(new TextRun({ text: '[' + tok.slice(2, -2) + ']', font: FONT, highlight: 'yellow', ...base }));
    last = m.index + tok.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), font: FONT, ...base }));
  return out;
}
const p = (t, o = {}) => new Paragraph({ children: runs(t, o.run), spacing: { after: 120, line: 276 }, alignment: o.align || AlignmentType.JUSTIFIED, ...o.para });
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t, font: FONT })], pageBreakBefore: true });
const h1n = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t, font: FONT })] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t, font: FONT })] });
const h3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun({ text: t, font: FONT })] });
const bl = (t) => new Paragraph({ numbering: { reference: 'bul', level: 0 }, children: runs(t), spacing: { after: 60, line: 276 } });
const note = (t) => new Paragraph({
  children: runs(t, { italics: true, size: 20 }),
  spacing: { after: 120 },
  border: { left: { style: BorderStyle.SINGLE, size: 12, color: 'BF9000', space: 8 } },
  indent: { left: 200 },
});

const border = { style: BorderStyle.SINGLE, size: 4, color: 'A6A6A6' };
const borders = { top: border, bottom: border, left: border, right: border };
function cell(text, width, o = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    columnSpan: o.span,
    children: [new Paragraph({
      alignment: o.align || AlignmentType.LEFT,
      children: runs(String(text), { size: o.size || 19, bold: o.bold, color: o.color }),
    })],
  });
}
function table(headers, rows, widths, o = {}) {
  const head = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => cell(h, widths[i], { fill: COLOR, bold: true, color: 'FFFFFF', align: o.aligns ? o.aligns[i] : undefined })),
  });
  const body = rows.map((r) => new TableRow({
    cantSplit: true,
    children: r.map((c, i) => cell(c, widths[i], {
      align: o.aligns ? o.aligns[i] : undefined,
      fill: o.boldRow && o.boldRow(r) ? 'DEEAF6' : undefined,
      bold: o.boldRow && o.boldRow(r),
    })),
  }));
  return new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows: [head, ...body] });
}
const gap = () => new Paragraph({ children: [], spacing: { after: 120 } });
const fmt = (n) => n.toLocaleString('es-EC');

// ---------- cálculos ----------
const cursos = M.build();
const T = M.totales(cursos);
const porPao = {};
cursos.forEach((c) => { porPao[c.pao] = (porPao[c.pao] || 0) + c.cr; });
const porArea = {};
cursos.forEach((c) => { porArea[c.area] = (porArea[c.area] || 0) + c.cr; });
const practicas = cursos.filter((c) => c.tipo === 'P');
const hPract = practicas.reduce((s, c) => s + c.H, 0);
const hServ = cursos.find((c) => c.codigo === '4.5').H;
const hLab = hPract - hServ;
const noPract = cursos.filter((c) => c.tipo !== 'P');
const interaccion = noPract.reduce((s, c) => s + c.c + c.p, 0) / T.H * 100;
const minContacto = Math.min(...noPract.map((c) => c.c / c.cr));
const semanas45 = (T.H / 8) / 45;
const pct = (x, t) => (x / t * 100).toFixed(1).replace('.', ',') + ' %';

// ---------- contenido ----------
const C = [];

// Portada
C.push(new Paragraph({ children: [], spacing: { before: 1800 } }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'UNIVERSIDAD TÉCNICA ESTATAL DE QUEVEDO', font: FONT, size: 28, bold: true, color: COLOR })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 }, children: [new TextRun({ text: 'Facultad de Ciencias de la Computación [confirmar unidad proponente]', font: FONT, size: 22, highlight: 'yellow' })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: 'PROYECTO DE CREACIÓN DE CARRERA', font: FONT, size: 40, bold: true, color: COLOR })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 }, children: [new TextRun({ text: 'Ingeniería en Internet de las Cosas', font: FONT, size: 36 })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Tercer nivel de grado · Modalidad presencial', font: FONT, size: 24 })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 900 }, children: [new TextRun({ text: 'BORRADOR v0.2 · 26 de septiembre de 2026', font: FONT, size: 24, bold: true })] }));
C.push(note('**Estado del documento.** Borrador de trabajo para discusión interna. No está aprobado por ningún órgano de la UTEQ ni presentado al CES. Todo dato resaltado en amarillo, como [[este ejemplo]], falta por completar o confirmar. Las cifras de demanda ocupacional y de recursos institucionales no se han levantado: no se inventaron y quedan marcadas como pendientes.'));
C.push(note('**Base normativa.** Reglamento de Régimen Académico del CES (2022, reforma de 11-IV-2023), Guía Metodológica para la presentación de carreras (codificada en 2023) y Reglamento de Armonización de la Nomenclatura RPC-SE-05-No.014-2023. No se revisó la reforma 2024 del RRA ni el estatuto UTEQ posterior a 2019.'));

// Contenido
C.push(h1('Contenido'));
[
  '1. Estructura del expediente y correspondencia con este borrador',
  '2. Datos generales de la carrera',
  '3. Informe académico',
  '4. Estudio de pertinencia',
  '5. Justificación epistemológica',
  '6. Malla curricular',
  '7. Recursos, cohorte y convenios',
  '8. Pendientes y hoja de ruta',
  'Anexo A. Fuentes',
].forEach((t) => C.push(new Paragraph({ children: runs(t), spacing: { after: 80 } })));

// 1. Estructura
C.push(h1('1. Estructura del expediente y correspondencia con este borrador'));
C.push(p('El CES no recibe un documento libre: el proyecto se carga en la plataforma PPCP con la información y documentación que fija la Guía Metodológica (numerales 1.1 a 1.4 y Anexo 1 o 2). Esta tabla relaciona cada requisito con la sección de este borrador.'));
C.push(table(
  ['Requisito de la Guía Metodológica', 'Dónde está en este borrador', 'Estado'],
  [
    ['1.1 Información institucional: datos de la IES, aval CACES si no está acreditada, datos del rector y del responsable', 'Portada y §2', 'Pendiente: datos del rector y del responsable; estatus CACES de la UTEQ'],
    ['1.2 Datos generales: nivel, modalidad, campos RANT, denominación, titulación, itinerarios, perfiles, segunda lengua, lugar, cohorte, periodos, créditos, prácticas', '§2 (tabla completa), §3.2 a §3.3, §6', 'Propuesto, con campos por confirmar'],
    ['1.2 Justificación si la denominación o el campo detallado son nuevos (Anexo III o IV del RANT)', '§5', 'Redactada; falta transponer al formato del Anexo III'],
    ['1.2 Resolución del órgano colegiado superior (Consejo Universitario)', '§8', 'Pendiente'],
    ['1.2 Informe académico aprobado por el órgano colegiado superior (Anexo 1)', '§3', 'Estructura lista; falta la matriz CACES'],
    ['Anexo 2: informe de par académico (solo si la UTEQ no es acreditada ni tiene aval CACES)', '§8', 'Depende del estatus CACES'],
    ['1.3 Convenios legalizados', '§7.3', 'Pendiente'],
    ['1.4 Declaración de recursos académicos, equipamiento, infraestructura y modelo educativo', '§7.1 y §7.2', 'Plantilla lista; faltan datos reales'],
    ['Estudio de pertinencia (sustento del informe académico y del principio del Art. 107 de la LOES)', '§4', 'Marco y datos nacionales listos; falta el trabajo de campo'],
  ],
  [3900, 2626, 2500],
));
C.push(gap());
C.push(p('**Cómo se organiza el informe académico (Anexo 1).** Contiene información de la IES y de la carrera, un resumen de la propuesta (objetivo general, perfil de egreso, líneas de investigación, vinculación con la sociedad y prácticas preprofesionales), la descripción del cumplimiento de los criterios y estándares básicos de calidad del CACES, y una conclusión firmada por un responsable afín a la carrera. El §3 sigue ese orden.'));

// 2. Datos generales
C.push(h1('2. Datos generales de la carrera'));
C.push(table(
  ['Campo (Guía 1.2)', 'Propuesta', 'Base o estado'],
  [
    ['Denominación', 'Ingeniería en Internet de las Cosas', 'Ver §5.5; verificar Anexo II vigente del RANT'],
    ['Titulación', 'Ingeniero/a en Internet de las Cosas', 'Art. 7 RANT'],
    ['Nivel de formación', 'Tercer nivel de grado', 'Art. 11 RRA'],
    ['Modalidad', 'Presencial', 'Art. 55 RRA. Carrera con carga de laboratorio. [[Confirmar]]'],
    ['¿Proyecto experimental o innovador?', 'No', 'Solo IES acreditadas pueden presentar proyectos experimentales (Art. 96 RRA)'],
    ['¿Proyecto en red?', 'No', 'Art. 96 RRA'],
    ['Campo amplio, específico y detallado', '06 Tecnologías de la información y la comunicación (TIC) / 1 Tecnologías de la información y la comunicación / 1 Ciencias computacionales (código 0611). [[Confirmar con el CES]]', 'Anexo I 2023 del RANT. Sustento y alternativa en §5.3'],
    ['Itinerarios', '2: A) IoT para agroindustria y ambiente; B) IoT industrial y ciudades inteligentes', 'Art. 16 RRA, máximo 3'],
    ['Perfil de ingreso', 'Título de bachiller, con acceso por el sistema de nivelación y admisión', 'Art. 13 RRA. Detalle en §3.2'],
    ['Perfil de egreso', 'Ocho resultados de aprendizaje', '§3.3'],
    ['Segunda lengua', 'Suficiencia de inglés como requisito de titulación, nivel B1 del Marco Común Europeo o equivalente', 'Art. 64 RRA: el nivel mínimo para el tercer nivel de grado es B1'],
    ['Lugar de ejecución', 'Campus Central, Quevedo, Los Ríos', 'Dirección institucional según memorando UTEQ-VICACAD-2026-1787-M'],
    ['Estudiantes por cohorte', '40 estudiantes en 1 paralelo', 'Art. 96 RRA. Debe respaldarse con la capacidad de los laboratorios (§7.1)'],
    ['Número de periodos académicos', '8', `${semanas45.toFixed(1).replace('.', ',')} semanas por período a 45 h/semana; ver §6.1`],
    ['Total de créditos', `${T.cr} créditos (${fmt(T.H)} horas)`, 'Rango del Art. 15 RRA: 120 a 150 créditos'],
    ['Créditos de aprendizaje en contacto con el docente', `${(T.c / 48).toFixed(1).replace('.', ',')} (${fmt(T.c)} h)`, 'Cálculo de la malla propuesta'],
    ['Créditos de aprendizaje práctico-experimental', `${(T.p / 48).toFixed(1).replace('.', ',')} (${fmt(T.p)} h)`, 'Incluye las prácticas preprofesionales'],
    ['Créditos de aprendizaje autónomo', `${(T.a / 48).toFixed(1).replace('.', ',')} (${fmt(T.a)} h)`, 'Cálculo de la malla propuesta'],
    ['Prácticas preprofesionales', `${hPract} h en total: ${hLab} h laborales y ${hServ} h de servicio comunitario`, `Mínimos del Art. 43 RRA: 240 h y 60 h. Equivalen a ${pct(hPract, T.H)} de la carrera; el tope es 10 %`],
    ['Resolución del órgano colegiado superior', '[[Número y fecha, cuando exista]]', 'Consejo Universitario de la UTEQ'],
  ],
  [2600, 3626, 2800],
));

// 3. Informe académico
C.push(h1('3. Informe académico'));
C.push(h2('3.1 Objetivos de la carrera'));
C.push(p('**Objetivo general.** Formar ingenieros e ingenieras capaces de diseñar, implementar, operar y asegurar soluciones de Internet de las Cosas que integren electrónica embebida, comunicaciones, plataformas en la nube y analítica de datos, con criterios de ética, sostenibilidad y atención a las necesidades productivas del país y de la provincia de Los Ríos.'));
C.push(p('**Objetivos específicos.**'));
[
  'Desarrollar dispositivos conectados con sensores y actuadores, desde el diseño del circuito hasta el firmware.',
  'Implementar redes y protocolos de comunicación adecuados a cada contexto de despliegue, urbano o rural.',
  'Construir plataformas que recojan, procesen y analicen datos de dispositivos, incluida la inteligencia artificial en el borde.',
  'Proteger los sistemas y los datos que manejan, conforme a la normativa ecuatoriana de protección de datos personales.',
  'Impulsar la investigación aplicada y el emprendimiento tecnológico vinculados al sector productivo de la zona.',
].forEach((t) => C.push(bl(t)));

C.push(h2('3.2 Perfil de ingreso'));
C.push(p('Título de bachiller o su equivalente y cumplimiento de los requisitos del sistema de acceso, nivelación y admisión de las instituciones públicas (Art. 13 del RRA). Se valoran, sin ser requisito, el razonamiento lógico-matemático, el interés por la tecnología y el gusto por el trabajo experimental.'));
C.push(p('**Proceso de ingreso en la UTEQ.** Según la página de admisión de la UTEQ para 2026, el aspirante sigue siete fases: registro nacional en la plataforma del Ministerio de Educación (MINEDEC); registro en la UTEQ y elección de carrera; simulador de la evaluación con varios intentos; evaluación de competencias generales y específicas; publicación de la nota de postulación; confirmación o cambio de carrera; y aceptación del cupo en la plataforma del MINEDEC. La nota de postulación combina la nota de bachillerato (50 %) y la de la evaluación (50 %), más un puntaje adicional de acción afirmativa cuando corresponde, con un máximo de 1.000 puntos. La UTEQ evalúa según el Art. 36 del Reglamento de Nivelación y Admisión, y las áreas dependen de la carrera: entre las que publica figuran matemáticas, razonamiento lógico, lengua e inglés. Aplica una política de cupos de acción afirmativa del 5 % en todas las carreras.'));
C.push(p('**Enfoque de derechos.** El Art. 5 del RRA obliga a las IES a concretar acciones afirmativas hacia los grupos de atención prioritaria (mujeres, pueblos y nacionalidades, personas con discapacidad, entre otros) y a incorporarlas en el plan institucional de igualdad. Para esta carrera se propone fijar una meta de participación de mujeres y prever adecuaciones de laboratorio para estudiantes con discapacidad. [[Sustentar la meta con la matrícula por sexo de las carreras TIC de la UTEQ y definirla con la Unidad de Admisión y el plan de igualdad]]'));
C.push(h3('Tipos de bachillerato de los aspirantes'));
C.push(p('Los aspirantes llegan de trayectorias muy distintas. Según el Ministerio de Educación, el Bachillerato General Unificado tiene tres años y se ofrece en dos opciones principales, más dos bachilleratos complementarios. A ello se suman vías que la página del Ministerio menciona solo por enlace y que hay que considerar por su peso en una universidad pública de la costa.'));
C.push(table(
  ['Tipo de bachillerato', 'Rasgos', 'Carga semanal en matemática y física*', 'Brecha probable frente a los prerrequisitos de IoT'],
  [
    ['Bachillerato en Ciencias (BGU)', 'Formación científico-humanística. 40 períodos semanales; en tercer año, 10 horas de asignaturas optativas', 'Matemática 5, 5 y 4 h (1.º, 2.º y 3.º año). Física 3, 3 y 2 h', 'Programación, electrónica y trabajo de laboratorio casi ausentes. La competencia digital es desigual'],
    ['Bachillerato Técnico (BGU), figuras de la familia Tecnologías: desarrollo de software, redes y telecomunicaciones, seguridad informática, soporte informático, ciencias de datos', '45 períodos semanales; módulos técnicos de 12 h en 1.º y 2.º año y 21 h en 3.º', 'Matemática 3, 3 y 2 h. Física 2 h por año', 'Buen punto de partida en programación y redes. Déficit en matemática y física: unas 8 horas de matemática en los tres años, frente a 14 en Ciencias'],
    ['Bachillerato Técnico, figuras de la familia Industrial: electrónica, mecatrónica, electromecánica industrial', 'Igual estructura; formación práctica en taller', 'Igual que el anterior', 'Fortaleza en circuitos y taller. Déficit en matemática, física y programación'],
    ['Bachillerato Técnico, otras familias: agropecuaria, administrativa y financiera, turismo, diseño, construcción', 'Igual estructura', 'Igual que el anterior', 'Déficit en matemática y física, y sin base técnica cercana a IoT'],
    ['Bachillerato Técnico Productivo y Complementario en Artes', 'Bachilleratos complementarios que se cursan después del general', 'Depende del bachillerato de origen', 'Igual que el bachillerato de base, con menos práctica científica en el caso de Artes'],
    ['Otras vías: intercultural bilingüe, escolaridad inconclusa (jóvenes y adultos), currículos internacionales, títulos del extranjero', 'Trayectorias heterogéneas; el Art. 13 del RRA obliga a aceptar títulos extranjeros reconocidos o equiparados por el Ministerio de Educación', 'Variable', 'Muy dispersa: en adultos, tiempo sin estudiar; en intercultural bilingüe, posible brecha de lengua. [[Verificar en el Ministerio si el Bachillerato Internacional aplica a la zona]]'],
  ],
  [2300, 2300, 2100, 2326],
));
C.push(note('* Malla 2024-2025 de la región Sierra-Amazonía publicada por Primicias a partir del Ministerio de Educación. El Acuerdo MINEDUC-2024-00065-A reformó el Bachillerato Técnico desde el año lectivo 2026-2027 de la Costa, con 21 horas técnicas y 19 de tronco común; no se pudo confirmar cuántas horas de matemática y física deja. [[Pedir al Ministerio de Educación la malla vigente en la Costa y los graduados de Los Ríos por tipo de bachillerato y figura profesional (AMIE)]]'));

C.push(h3('Conocimientos mínimos que el primer período da por sabidos'));
C.push(p('Las asignaturas del primer período (cálculo, álgebra, fundamentos de programación, física, comunicación y una introducción a la carrera) suponen:'));
[
  '**Matemática:** aritmética y álgebra, ecuaciones e inecuaciones, funciones, trigonometría y geometría analítica básica.',
  '**Física:** magnitudes y unidades, cinemática, dinámica, energía y electricidad elemental.',
  '**Pensamiento lógico y algorítmico:** secuenciar pasos, condiciones y repeticiones.',
  '**Lectura y escritura académicas y competencia digital:** leer textos técnicos, redactar y usar herramientas digitales de estudio.',
  '**Inglés técnico elemental:** la documentación de sensores y plataformas está en inglés.',
].forEach((t) => C.push(bl(t)));

C.push(h3('Propuesta de nivelación específica para la carrera'));
C.push(p('El curso de nivelación de la UTEQ se dicta en modalidad virtual y busca fortalecer los conocimientos y habilidades básicas de los aspirantes; el Art. 14 del RRA permite a las IES diseñar estrategias de nivelación. Como referencia histórica, la nivelación de la SENESCYT para el área de ingenierías dedicaba 200 horas a matemáticas, 100 a física, 100 a química y 140 a un tronco común. Para IoT se propone reemplazar química, que la carrera no usa, por pensamiento computacional e introducción a la electrónica:'));
C.push(table(
  ['Asignatura de nivelación', 'Horas', 'Contenidos', 'Quiénes la necesitan más'],
  [
    ['N1 Matemática para ingeniería', '120', 'Álgebra, ecuaciones, funciones, trigonometría, geometría analítica y vectores en el plano', 'Bachilleratos Técnicos, Artes, vías especiales'],
    ['N2 Física básica', '80', 'Magnitudes, cinemática, dinámica, energía y electricidad elemental (ley de Ohm)', 'Técnicos no industriales, Artes, vías especiales'],
    ['N3 Pensamiento computacional y lógica', '80', 'Algoritmos, pseudocódigo, estructuras de control y un primer lenguaje (Python)', 'Ciencias, Técnicos no informáticos'],
    ['N4 Comunicación académica y competencia digital', '60', 'Lectura técnica, redacción, herramientas digitales y vocabulario técnico en inglés', 'Todos; en especial adultos e intercultural bilingüe'],
    ['N5 Introducción a la electrónica y al IoT', '60', 'Circuitos simples, sensores y un primer proyecto con microcontrolador', 'Ciencias, Técnicos no industriales'],
    ['Total', '400', '16 semanas a 25 horas semanales, en modalidad virtual con sesiones sincrónicas', ''],
  ],
  [2400, 800, 3626, 2200],
  { aligns: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.LEFT, AlignmentType.LEFT], boldRow: (r) => r[0] === 'Total' },
));
C.push(gap());
C.push(p('**Rutas según un diagnóstico de entrada.** En la primera semana se aplica una prueba diagnóstica de matemática, física y lógica. El aspirante con nota suficiente en un área queda exonerado de esa asignatura y usa las horas en tutorías. Como hipótesis inicial:'));
[
  '**Ruta A, Bachillerato Técnico de Tecnologías e Industrial:** refuerzo de N1 y N2 con tutorías adicionales; exoneración parcial de N3 y N5.',
  '**Ruta B, Bachillerato en Ciencias:** exoneración parcial de N1 y N2; refuerzo de N3 y N5.',
  '**Ruta C, otros bachilleratos y vías especiales:** las cinco asignaturas, con tutoría personalizada.',
].forEach((t) => C.push(bl(t)));
C.push(note('Esta propuesta es de la carrera. Su aplicación depende de la Unidad de Admisión y Nivelación y del Reglamento de Nivelación y Admisión, que la UTEQ cita en su evaluación de ingreso; los créditos de nivelación no forman parte de los 136 de la malla. La UTEQ no publica la duración ni las asignaturas de su nivelación actual. [[Confirmar con la Unidad de Admisión y Nivelación si puede haber una nivelación por carrera y con qué duración]]'));

C.push(h2('3.3 Perfil de egreso'));
C.push(p('El egresado o egresada logra los siguientes resultados de aprendizaje. La columna de la derecha indica las asignaturas de la malla (§6) que los desarrollan.'));
C.push(table(
  ['Código', 'Resultado de aprendizaje', 'Asignaturas principales'],
  M.RA.map(([c, t, as]) => [c, t, as.join(', ')]),
  [900, 5926, 2200],
));
C.push(gap());
C.push(p('**Campos de desempeño profesional.** Ingeniero o ingeniera de soluciones IoT; arquitecto o arquitecta de soluciones conectadas; desarrollador o desarrolladora de firmware y sistemas embebidos; especialista en redes de sensores y conectividad; analista de datos de dispositivos; consultor o consultora en agricultura de precisión y agroindustria; emprendedor o emprendedora de base tecnológica. [[Contrastar con el estudio de demanda ocupacional del §4.9]]'));

C.push(h2('3.4 Líneas de investigación'));
C.push(p('Se proponen cinco líneas, que deben alinearse con las líneas y dominios académicos vigentes en la UTEQ. [[Verificar contra el plan de investigación institucional]]'));
[
  'Sistemas ciberfísicos y agricultura de precisión.',
  'Redes de sensores y comunicaciones de bajo consumo.',
  'Inteligencia artificial en el borde y analítica de datos de sensores.',
  'Ciberseguridad y privacidad en sistemas conectados.',
  'IoT para la sostenibilidad ambiental, el agua y las ciudades inteligentes.',
].forEach((t) => C.push(bl(t)));
C.push(p('La investigación formativa se desarrolla en las asignaturas de proyecto de cada período y en la metodología de la investigación (6.5); la titulación integra los resultados (Arts. 31 y 32 del RRA).'));

C.push(h2('3.5 Vinculación con la sociedad'));
C.push(p('Se plantean tres frentes, dentro de las líneas operativas del Art. 41 del RRA:'));
[
  '**Educación continua:** cursos cortos de sensórica y monitoreo para productores agropecuarios y agroindustriales de la zona.',
  '**Proyectos y servicios especializados:** pilotos de monitoreo ambiental, de riego y de cadena de frío con gremios, cooperativas y gobiernos locales.',
  '**Prácticas preprofesionales y servicio comunitario:** ver §3.6.',
].forEach((t) => C.push(bl(t)));
C.push(p('[[Identificar los actores concretos y las cartas de intención en el §7.3]]'));

C.push(h2('3.6 Prácticas preprofesionales'));
C.push(table(
  ['Componente', 'Asignatura', 'Período', 'Horas', 'Mínimo Art. 43'],
  [
    ['Servicio comunitario', '4.5 Prácticas de servicio comunitario', '4', String(hServ), '60 h'],
    ['Prácticas laborales', '7.5 Prácticas preprofesionales laborales I', '7', String(cursos.find((c) => c.codigo === '7.5').H), '240 h en total'],
    ['Prácticas laborales', '8.2 Prácticas preprofesionales laborales II', '8', String(cursos.find((c) => c.codigo === '8.2').H), 'con la anterior'],
    ['Total', '', '', String(hPract), `Tope: 10 % = ${Math.floor(T.H * 0.1)} h`],
  ],
  [1700, 3226, 900, 1000, 2200],
  { aligns: [AlignmentType.LEFT, AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.LEFT], boldRow: (r) => r[0] === 'Total' },
));
C.push(gap());
C.push(p('Las prácticas se realizan en entornos empresariales, institucionales o comunitarios, con un responsable académico y un informe de la entidad receptora (Art. 44 del RRA). La UTEQ puede reconocer ayudantías de cátedra o de investigación como práctica (Art. 45).'));

C.push(h2('3.7 Itinerarios académicos'));
M.ITINERARIOS.forEach((it) => {
  C.push(p(`**${it.nombre}.** Tres optativas, una por cada uno de los períodos 7 y 8:`));
  it.optativas.forEach((o, i) => C.push(bl(`Optativa ${['I', 'II', 'III'][i]}: ${o}`)));
});
C.push(p('El itinerario cursado puede constar en el título (Art. 8 del RANT y Art. 16 del RRA).'));

C.push(h2('3.8 Metodología, evaluación y titulación'));
C.push(p('La formación se organiza en torno a proyectos: cada período cierra con un entregable integrador que combina los contenidos de sus laboratorios. Las asignaturas de laboratorio dedican cerca de un tercio de sus horas al componente práctico-experimental. [[Alinear con el modelo educativo y el sistema de evaluación de la UTEQ]]'));
C.push(p('La unidad de titulación se compone del diseño del proyecto (7.6) y del trabajo de integración curricular (8.3), cuyos créditos forman parte de los de la carrera (Art. 26 del RRA). [[Definir las opciones de titulación según el reglamento UTEQ]]'));

C.push(h2('3.9 Cumplimiento de los criterios y estándares del CACES'));
C.push(p('El informe debe describir, para cada criterio y estándar del modelo de evaluación de carreras del CACES, cómo lo cumple el proyecto. El sitio del CACES describe un modelo genérico de evaluación del entorno de aprendizaje con cinco criterios; se usan aquí como estructura, y los estándares de cada uno deben tomarse del modelo vigente. [[Confirmar criterios y estándares en el modelo vigente del CACES]]'));
C.push(table(
  ['Criterio (modelo genérico CACES)', 'Dónde lo sustenta este borrador', 'Estado'],
  [
    ['Pertinencia: planificación y vinculación con la sociedad', '§4 (estudio de pertinencia), §3.5 (vinculación), §3.4 (líneas de investigación)', 'Marco y datos listos; falta trabajo de campo'],
    ['Organización y recursos: gestión académica', '§7.1 (infraestructura), §8.2 (ruta de aprobación)', 'Falta inventario real'],
    ['Profesores', '§7.2 (planta docente)', 'Falta levantar docentes'],
    ['Currículo', '§3.3 (perfil de egreso), §6 (malla y verificación normativa)', 'Faltan sílabos y prerrequisitos'],
    ['Estudiantes', '§3.2 (perfil de ingreso), §3.6 (prácticas), acción afirmativa', 'Faltan apoyos y tutorías'],
  ],
  [3200, 3826, 2000],
));

// 4. Pertinencia
C.push(h1('4. Estudio de pertinencia'));
C.push(note('Este estudio combina estadísticas oficiales verificadas (4.2 a 4.8), un ejercicio de dimensionamiento de la demanda con supuestos de trabajo (4.9 a 4.11) y el plan del trabajo de campo (4.12), que todavía no se hizo. Las estadísticas dan contexto y muestran brechas; la demanda de ingenieros IoT depende de supuestos que solo las encuestas del 4.12 pueden confirmar.'));
C.push(h2('4.1 Marco'));
C.push(p('El Art. 107 de la LOES define la pertinencia como la respuesta de la educación superior a las necesidades de la sociedad, a la planificación nacional y a la prospectiva de desarrollo científico y tecnológico, y exige articular la oferta con las necesidades locales y regionales, las tendencias del mercado ocupacional y la estructura productiva de la provincia y la región. La Guía Metodológica pide sustentar la demanda estudiantil y la ocupacional con fuentes primarias y secundarias oficiales.'));

C.push(h2('4.2 Transformación digital y talento'));
C.push(p('La Política Pública para la Transformación Digital del Ecuador 2025-2030, del Ministerio de Telecomunicaciones y de la Sociedad de la Información (MINTEL), incluye el Internet de las Cosas entre las tendencias de transformación digital (p. 46) y señala que las tecnologías emergentes, entre ellas el IoT, pueden aplicarse al desarrollo sostenible en áreas como la salud, la energía y la agricultura (p. 82).'));
C.push(p('El mismo documento afirma que Ecuador sufre un déficit de talento digital y que muchos graduados en áreas TIC carecen de experiencia y de actualización en nuevas herramientas (pp. 51 y 52). Recoge además que solo el 15,27 % de la población tenía en 2019 habilidades avanzadas para instalar y configurar software, según el INEC (p. 55).'));

C.push(h2('4.3 Oferta nacional de carreras TIC'));
C.push(p('Según los datos de la SENESCYT recogidos en esa política (Tabla 7, pp. 52 y 53), en 2024 había 197 carreras TIC en el país. La oferta se concentra en software y es muy escasa en IoT:'));
C.push(table(
  ['Carrera TIC (SENESCYT, 2024)', 'Carreras ofertadas por universidades'],
  [
    ['Desarrollo de software', '62'],
    ['Tecnología superior en desarrollo de software', '36'],
    ['Tecnología superior en redes y telecomunicaciones', '14'],
    ['Ciberseguridad', '9'],
    ['Internet de las cosas', '2'],
    ['Técnico superior en internet de las cosas', '1'],
  ],
  [6026, 3000],
  { aligns: [AlignmentType.LEFT, AlignmentType.CENTER] },
));
C.push(gap());
C.push(p('El mismo informe cita 7.476 graduados en carreras TIC en 2021, frente a 5.079 en 2020 (texto de la p. 53); su Tabla 8 consigna 7.485 para 2021, una discrepancia menor en la fuente. Para dimensionar el mercado estudiantil conviene descargar del portal de datos abiertos de la SENESCYT la matrícula y los graduados TIC por provincia. [[Descargar la serie de matrícula TIC de Los Ríos, Guayas, Bolívar y Santo Domingo]]'));
C.push(note('Lectura provisional: la oferta IoT es marginal frente al software, lo que apoya la pertinencia de una carrera especializada. Es un argumento de oferta, no de demanda.'));

C.push(h2('4.4 Sector productivo de Los Ríos'));
C.push(p('La provincia de Los Ríos, donde funciona el Campus Central de la UTEQ, concentra una parte central de la producción agrícola del país. Según la Encuesta de Superficie y Producción Agropecuaria Continua (ESPAC) 2024 del INEC, publicada en abril de 2025:'));
C.push(table(
  ['Cultivo (ESPAC 2024)', 'Participación de Los Ríos en la producción nacional'],
  [
    ['Banano', '39,8 %'],
    ['Maíz duro seco', '43,8 %'],
    ['Palma africana', '28,6 %'],
    ['Arroz en cáscara', '27,3 % (segunda provincia, tras Guayas)'],
  ],
  [4600, 4426],
  { aligns: [AlignmentType.LEFT, AlignmentType.LEFT] },
));
C.push(gap());
C.push(p('El Módulo de Información Ambiental y Tecnificación Agropecuaria (MIATA) de la ESPAC 2024 muestra, a nivel nacional, cuánto margen hay para incorporar medición y datos al campo:'));
[
  'Solo el 28,2 % de la superficie cultivada del país se regó en 2024 (1,30 millones de hectáreas), y el 40,9 % de la superficie de cultivos permanentes está bajo riego.',
  'El 88,9 % de las unidades de producción nunca ha hecho un análisis de suelo; el 43,0 % dice que no percibe la necesidad y el 35,1 % que no conoce sus beneficios.',
  'Solo el 12,5 % de las unidades de producción recibió capacitación o asistencia técnica agrícola en 2024, y quienes no la recibieron citan, entre otras razones, que las sedes están lejos (19,5 %) y que los técnicos no están disponibles (15,3 %).',
].forEach((t) => C.push(bl(t)));
C.push(note('Lectura provisional: la sensórica de suelo, agua y clima, y la trazabilidad, son brechas medibles en el sector que más pesa en la provincia. Falta desagregar el MIATA para Los Ríos, si el INEC publica las bases, y confirmar cuántas empresas agroindustriales de la zona ya usan monitoreo.'));

C.push(h2('4.5 Mercado laboral'));
C.push(p('La Encuesta Nacional de Empleo, Desempleo y Subempleo (ENEMDU) anual 2025 del INEC ofrece este contexto para Los Ríos frente al país:'));
C.push(table(
  ['Indicador (ENEMDU anual 2025)', 'Los Ríos', 'Nacional'],
  [
    ['Tasa de empleo adecuado o pleno', '32,5 % (36,3 % en 2024)', '37,1 %'],
    ['Tasa de subempleo', '28,4 % (24,0 % en 2024)', '[[Completar]]'],
    ['Tasa de desempleo', '1,7 %', '3,6 %'],
  ],
  [3800, 3000, 2226],
  { aligns: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER] },
));
C.push(gap());
C.push(p('En Los Ríos casi no hay desempleo abierto, pero el empleo adecuado está por debajo del promedio nacional y el subempleo subió en 2025. Es un contexto de baja productividad del trabajo, en el que la tecnificación puede aportar; no prueba demanda de ingenieros IoT. [[Pedir al INEC la ENEMDU por rama de actividad para Los Ríos y por nivel de instrucción superior]]'));

C.push(h2('4.6 Conectividad'));
C.push(p('Un despliegue IoT depende de la conectividad. Según el INEC (ENEMDU, julio de 2025), el 71,3 % de los hogares del país tiene acceso a internet: 76,6 % en el área urbana y 58,7 % en la rural, frente a 38,0 % en la rural en 2022. La ARCOTEL informa, para el cuarto trimestre de 2024, una penetración de internet fijo de 17,48 % y de internet móvil de 65,6 %, y ubica a Los Ríos entre las cinco provincias con más radiobases instaladas (Guayas, Pichincha, Manabí, Azuay y Los Ríos).'));
C.push(p('Dos lecturas se derivan: la conectividad rural mejora, lo que amplía el mercado de soluciones IoT; y una brecha persiste, lo que exige formar profesionales capaces de diseñar con redes de bajo consumo y conectividad intermitente. [[Agregar cobertura móvil y de fibra por parroquia rural de Los Ríos desde los mapas de cobertura de la ARCOTEL y las metas del MINTEL]]'));

C.push(h2('4.7 Oferta de la UTEQ y análisis de traslape'));
C.push(p('Según el PEDI 2021-2025, la UTEQ tenía 29 carreras de grado aprobadas al cierre de 2020, entre ellas Software (RPC-SO-03-No.022-2018), Telemática (RPC-SO-03-No.049-2020) y Electricidad (RPC-SO-04-No.077-2020), y 9.653 estudiantes matriculados. La carrera más cercana a IoT es Telemática. Hay que demostrar en qué se diferencia.'));
C.push(table(
  ['Aspecto', 'Telemática (vigente)', 'Ingeniería en IoT (propuesta)'],
  [
    ['Objeto de estudio', '[[Completar desde el proyecto aprobado]]', 'Sistemas de dispositivos conectados con percepción y actuación sobre el entorno físico'],
    ['Perfil de egreso', '[[Completar]]', 'Ver §3.3'],
    ['Núcleo de la malla', '[[Completar]]', 'Electrónica embebida, comunicaciones, plataformas y analítica de datos de sensores'],
    ['Campo ocupacional', '[[Completar]]', 'Ver §3.3, con énfasis agroindustrial e industrial'],
    ['Asignaturas compartidas o convalidables', '[[Comparar mallas]]', '[[Comparar mallas]]'],
  ],
  [2100, 3000, 3926],
));
C.push(gap());
C.push(p('Si la comparación muestra un traslape mayor al de una carrera nueva, la alternativa es un itinerario IoT dentro de una carrera vigente mediante un ajuste curricular sustantivo (Art. 110 del RRA). Fuera de la UTEQ, la UEES Online publica una carrera de "Ingeniero en Internet de las Cosas" con la resolución RPC-SO-26-No.428-2023 [[verificar en el CES]].'));

C.push(h2('4.8 Empresas de la zona'));
C.push(p('El Registro Estadístico de Empresas 2025 del INEC, construido con registros del SRI y del IESS, contabiliza en Los Ríos 32.993 empresas, la undécima provincia del país. De ellas, 1.968 registran ventas y empleo en el IESS (octava provincia) y el empleo registrado equivalente suma 66.742 puestos (séptima provincia).'));
C.push(p('El catastro de contribuyentes del SRI, actualizado el 1 de septiembre de 2026, permite ver los sectores. Se cuentan 276.712 RUC registrados en Los Ríos, de los cuales 112.164 están activos. Como el catastro no trae el tamaño, se toma como empresa formal a la sociedad o a la persona natural obligada a llevar contabilidad: son 6.957 (5.332 sociedades y 1.625 personas naturales). Por actividad económica:'));
C.push(table(
  ['Grupo de actividad (CIIU Rev. 4.1)', 'Sociedades', 'Personas naturales obligadas', 'Empresas formales'],
  [
    ['G1 Agricultura, ganadería y silvicultura', '851', '546', '1.397'],
    ['G2 Pesca y acuicultura', '13', '1', '14'],
    ['G3 Alimentos y bebidas (agroindustria)', '76', '17', '93'],
    ['G4 Otras manufacturas', '133', '30', '163'],
    ['G5 Energía, agua y saneamiento', '97', '0', '97'],
    ['G6 Transporte y almacenamiento', '461', '46', '507'],
    ['G7 Comercio al por mayor', '661', '406', '1.067'],
    ['G8 Administración pública (GAD y entidades)', '96', '0', '96'],
    ['G9 Proveedores tecnológicos (TIC, ingeniería, I+D)', '229', '28', '257'],
    ['G10 Construcción', '570', '17', '587'],
    ['G11 Salud', '186', '15', '201'],
    ['Otros sectores', '1.959', '519', '2.478'],
    ['Total', '5.332', '1.625', '6.957']
  ],
  [4626, 1400, 1700, 1300],
  { aligns: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER], boldRow: (r) => r[0] === 'Total' },
));
C.push(gap());
C.push(p('Los grupos G1 a G9 suman 3.691 empresas formales y son el universo donde puede haber demanda de perfiles IoT. Se concentran en los cantones de Quevedo (845), Babahoyo (724), Ventanas (353), Buena Fe (316), Vinces (247). Además, las altas de sociedades que siguen activas pasaron de 263 en 2019 a 536 en 2025, lo que indica un tejido empresarial que se amplía. Los datos de la sección 4.4 muestran que el sector agropecuario es el mayor grupo, con 1.397 empresas formales.'));

C.push(h2('4.9 Estimación de la demanda a 4 y 5 años'));
C.push(note('Ninguna fuente oficial publica cuántas empresas ecuatorianas adoptan IoT ni cuántos profesionales de ese perfil contratan. Por eso los porcentajes de adopción de esta sección son supuestos de trabajo, no datos. El ejercicio sirve para dimensionar y para ver qué supuestos pesan más; la encuesta del 4.12 debe reemplazarlos. El modelo completo, con todas las fórmulas, está en 06_modelo_demanda/Modelo_demanda_oferta_IoT_LosRios.xlsx.'));
C.push(p('**Método.** Para cada grupo de actividad, los puestos de perfil IoT en 2030 y 2031 son: empresas formales del SRI, por un factor de empresas operativas, por el porcentaje que adopta, por los profesionales que contrata cada empresa adoptante.'));
[
  '**Factor de empresas operativas: 30 %.** Es el cociente entre las 1.968 empresas del INEC con ventas y empleo en el IESS y las 6.541 empresas formales del SRI sin administración pública, enseñanza y salud. Se aplica a todos los grupos salvo el sector público, donde los gobiernos autónomos descentralizados se toman todos como operativos.',
  '**Adopción.** Porcentaje de empresas operativas que emplea al menos un perfil IoT. Se usan tres escenarios: conservador (por ejemplo, 1 % en 2030 y 1,5 % en 2031 en agricultura), base (2,5 % y 3,5 %) y optimista (5 % y 7 %). Los proveedores tecnológicos y el sector público tienen porcentajes mayores.',
  '**Profesionales por empresa adoptante:** 1 en el escenario conservador, 1,5 en el base y 2 en el optimista.',
  '**Profesionales que ya trabajan en las empresas:** 0, porque no hay carrera IoT en la zona. [[Verificar en la encuesta]]',
].forEach((t) => C.push(bl(t)));
C.push(table(
  ['Escenario', 'Puestos a 4 años (2030)', 'Puestos a 5 años (2031)', 'Puestos nuevos por año, 2027-2031'],
  [
    ['Conservador', '35', '50', '9,9'],
    ['Base', '120', '165', '32,9'],
    ['Optimista', '283', '374', '74,8'],
  ],
  [2400, 2200, 2200, 2226],
  { aligns: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER] },
));
C.push(gap());
C.push(p('**Dónde está la demanda en el escenario base.** La agricultura aporta pocos puestos aunque sea el sector más grande, porque la adopción supuesta es baja. Los dos grupos que más pesan son el sector público y los proveedores tecnológicos:'));
C.push(table(
  ['Grupo', 'Empresas formales', 'Empresas operativas', 'Puestos base 2031', '% del total'],
  [
    ['G1 Agricultura, ganadería y silvicultura', '1.397', '419,1', '22,0', '13 %'],
    ['G2 Pesca y acuicultura', '14', '4,2', '0,4', '0 %'],
    ['G3 Alimentos y bebidas (agroindustria)', '93', '27,9', '5,9', '4 %'],
    ['G4 Otras manufacturas', '163', '48,9', '8,1', '5 %'],
    ['G5 Energía, agua y saneamiento', '97', '29,1', '6,1', '4 %'],
    ['G6 Transporte y almacenamiento', '507', '152,1', '16,0', '10 %'],
    ['G7 Comercio al por mayor', '1.067', '320,1', '9,6', '6 %'],
    ['G8 Administración pública (GAD y entidades)', '96', '96,0', '50,4', '31 %'],
    ['G9 Proveedores tecnológicos (TIC, ingeniería, I+D)', '257', '77,1', '46,3', '28 %']
  ],
  [3626, 1300, 1400, 1400, 1300],
  { aligns: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER] },
));
C.push(gap());
C.push(note('Consecuencia para el proyecto: la tesis agroindustrial de la carrera no se sostiene con estos supuestos por sí sola. El resultado depende de que los gobiernos locales y los proveedores tecnológicos contraten perfiles IoT, o de que la adopción agrícola supere lo supuesto. La encuesta debe medir ambos frentes.'));

C.push(h2('4.10 Oferta de graduados TIC'));
C.push(p('**Nacional.** Según la SENESCYT, recogida en la política del MINTEL, se graduaron 5.079 personas en carreras TIC en 2020 y 7.476 en 2021.'));
C.push(p('**UTEQ.** El informe de rendición de cuentas 2024 reporta 1.620 graduados de grado en el año (669 en el segundo período 2023-2024 y 951 en el primero de 2024-2025), con una tasa de titulación del 60,96 % para la cohorte 2018-2019. Para las carreras TIC de la Facultad de Ciencias de la Ingeniería, esa cohorte tiene estos graduados:'));
C.push(table(
  ['Carrera (cohorte 2018-2019)', 'Graduados', 'Estudiantes de la cohorte', 'Tasa de titulación'],
  [
    ['Ingeniería en Telemática', '16', '47', '34,0 %'],
    ['Telemática (rediseño)', '19', '40', '47,5 %'],
    ['Software (rediseño)', '10', '81', '12,4 %'],
    ['Fila sin nombre de carrera en el informe, bajo la Facultad de Ciencias de la Ingeniería', '24', '77', '31,2 %'],
  ],
  [4626, 1300, 1700, 1400],
  { aligns: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER] },
));
C.push(gap());
C.push(p('Las tres carreras identificadas suman 45 graduados y, con la fila sin nombre, 69. Además, el 67,32 % de los graduados de la UTEQ que respondió el seguimiento trabaja en un puesto relacionado con su formación, y el 32,68 % no tiene empleo.'));
C.push(note('Límite: la SENESCYT publica una base de matrícula de universidades y escuelas politécnicas 2015-2023 con provincia de la sede y campo del conocimiento, pero el servidor donde está alojada no respondió al consultarla, y no se encontró una serie de graduados por provincia. La oferta de la zona se estima entonces con dos referencias y un intervalo. Cuando se descargue la base, se filtra el campo amplio 06 y la provincia Los Ríos, y se reemplazan los supuestos de la hoja Supuestos del modelo. [[Descargar la base de matrícula TIC de Los Ríos y de las IES de la zona (Universidad Técnica de Babahoyo y otras)]]'));
C.push(table(
  ['Oferta de graduados TIC por año en IES de Los Ríos', 'Valor', 'Base del cálculo'],
  [
    ['Baja', '69', 'Máximo identificado en las carreras TIC de la UTEQ, cohorte 2018-2019'],
    ['Base', '187', '2,5 % de los 7.476 graduados TIC del país en 2021'],
    ['Alta', '397', 'Participación de la población de Los Ríos en el país (5,3 %, Censo 2022)'],
  ],
  [3300, 1000, 4726],
  { aligns: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.LEFT] },
));
C.push(gap());
C.push(p('No todos esos graduados tienen competencias IoT: la mayoría viene de software y sistemas. Se supone que tienen perfil útil para IoT el 2 % en la oferta baja, el 5 % en la base y el 10 % en la alta. Con eso, los graduados con perfil IoT por año son 1,4, 9,3 y 39,7.'));

C.push(h2('4.11 Brecha entre oferta y demanda'));
C.push(p('La primera promoción de la carrera se gradúa en 2031, si el primer ingreso es en 2027 y dura ocho períodos. Con 40 estudiantes por cohorte y la tasa de titulación de la UTEQ (60,96 %), cada promoción sería de unos 24 graduados. La brecha es la demanda menos los graduados con perfil IoT disponibles; un valor positivo indica falta de profesionales y uno negativo, exceso.'));
C.push(table(
  ['Brecha en puestos, con oferta base', 'A 4 años (2030), sin la carrera', 'A 5 años (2031), sin la carrera', 'A 5 años (2031), con la carrera'],
  [
    ['Demanda conservadora', '−3', '3', '−22'],
    ['Demanda base', '82', '118', '94'],
    ['Demanda optimista', '245', '328', '303']
  ],
  [3000, 2000, 2000, 2026],
  { aligns: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER] },
));
C.push(gap());
C.push(p('En régimen, a partir de 2032, la comparación anual es más útil:'));
C.push(table(
  ['Escenario', 'Puestos nuevos por año', 'Graduados IoT sin la carrera', 'Graduados de la carrera', 'Cobertura sin la carrera', 'Cobertura con la carrera'],
  [
    ['Conservador', '9,9', '9,3', '24,4', '94 %', '340 %'],
    ['Base', '32,9', '9,3', '24,4', '28 %', '102 %'],
    ['Optimista', '74,8', '9,3', '24,4', '12 %', '45 %']
  ],
  [1500, 1500, 1600, 1500, 1500, 1426],
  { aligns: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER] },
));
C.push(gap());
[
  '**Sin la carrera**, la oferta con perfil IoT cubre el 28 % de los puestos nuevos del escenario base y el 12 % del optimista.',
  '**Con la carrera**, una cohorte de 40 estudiantes cubre casi exactamente la demanda del escenario base. Es el tamaño adecuado si la adopción se acerca a los supuestos base.',
  'En el **escenario conservador** habría sobreoferta: la carrera y la oferta actual triplicarían los puestos nuevos. En el **optimista** faltarían profesionales aun con la carrera; se podría abrir un segundo paralelo más adelante.',
  'La demanda es muy sensible a la adopción: si la adopción base se duplica y cada empresa contrata dos profesionales, los puestos a 5 años pasan de unos 165 a unos 440 (hoja Sensibilidad del modelo).',
].forEach((t) => C.push(bl(t)));
C.push(note('Lectura provisional: la carrera se justifica en el escenario base y se arriesga en el conservador. El tamaño de la cohorte y la decisión de abrir un segundo paralelo deben tomarse con los resultados de la encuesta, no con este ejercicio.'));

C.push(h2('4.12 Trabajo de campo pendiente'));
C.push(p('Lo que las estadísticas oficiales no dan, y debe levantarse para reemplazar los supuestos:'));
[
  'Encuesta a empleadores de los grupos G1 a G9, con énfasis en agroindustria, gobiernos locales y proveedores tecnológicos: ¿tiene o piensa tener perfiles IoT?, ¿cuántos y en qué plazo?, competencias más demandadas y salarios de referencia.',
  'Encuesta a estudiantes de último año de bachillerato de la zona: interés en la carrera y tipo de bachillerato. La UTEQ puede sumar el dato de postulaciones por carrera de su Unidad de Admisión.',
  'Entrevistas a expertos y a graduados de Software, Telemática y Electricidad de la UTEQ.',
  'Revisión de ofertas de empleo publicadas para perfiles IoT, con fecha y fuente.',
].forEach((t) => C.push(bl(t)));
C.push(p('**Muestra.** Para poblaciones finitas se propone n = N·z²·p·q / (e²·(N−1) + z²·p·q), con z = 1,96, p = q = 0,5 y error e entre 5 % y 8 %. Con las 3.691 empresas formales de G1 a G9 como población, el tamaño de muestra estaría entre unas 145 y 350 empresas. [[Definir el marco muestral con el SRI y estratificar por grupo y cantón]]'));

C.push(h2('4.13 Conclusión provisional'));
C.push(p('Con la evidencia disponible, la carrera es coherente con la política de transformación digital, llena un vacío en la oferta nacional y responde a brechas medibles de tecnificación en el sector que más pesa en Los Ríos. El dimensionamiento sugiere que una cohorte de 40 estudiantes equilibra la oferta y la demanda en el escenario base, con riesgo de sobreoferta si la adopción es baja. Su aprobación depende de tres cosas que este borrador no resuelve: la demanda ocupacional medida en el territorio, la oferta real de graduados TIC de la zona y la diferenciación frente a Telemática. [[Reescribir esta sección cuando exista el trabajo de campo]]'));

// 5. Justificación epistemológica
C.push(h1('5. Justificación epistemológica'));
C.push(note('El Art. 25 del RANT exige esta justificación solo si la denominación o la titulación no constan en el Anexo II vigente, y debe presentarse en el formato del Anexo III 2023, que no se pudo consultar. El contenido de abajo está listo para transponerse a ese formato.'));
C.push(h2('5.1 Objeto de estudio'));
C.push(p('El objeto de estudio es el diseño, la operación y el aseguramiento de sistemas donde objetos físicos, dotados de sensores, actuadores y capacidad de comunicación, se conectan entre sí y con plataformas de cómputo para observar y modificar el entorno. La Unión Internacional de Telecomunicaciones lo define como una infraestructura global que interconecta objetos físicos y virtuales mediante tecnologías de información y comunicación interoperables (Recomendación UIT-T Y.2060, 2012).'));
C.push(h2('5.2 Fundamentos teóricos y metodológicos'));
[
  '**Cibernética y teoría de sistemas.** El control y la comunicación en máquinas y organismos (Wiener, 1948) y la teoría general de sistemas (von Bertalanffy, 1968) explican los lazos de medición, decisión y actuación que gobiernan un sistema IoT.',
  '**Computación ubicua.** La visión de una computación integrada en el entorno cotidiano (Weiser, 1991) es el antecedente conceptual directo del IoT. El término "Internet of Things" se atribuye a Kevin Ashton, en 1999 (Ashton, 2009).',
  '**Sistemas ciberfísicos.** La integración de cómputo, redes y procesos físicos, con sus problemas de tiempo real y fiabilidad (Lee, 2008), aporta la base de ingeniería del diseño de dispositivos y control.',
  '**Ingeniería de sistemas y de requisitos.** El diseño por capas (percepción, red, plataforma, aplicación) y la ingeniería de requisitos ordenan el trabajo de un proyecto IoT completo.',
  '**Ciencia de datos.** El tratamiento de series temporales y el aprendizaje automático, incluso en dispositivos de recursos limitados, convierten los datos de sensores en decisiones.',
].forEach((t) => C.push(bl(t)));
C.push(h2('5.3 Ubicación en los campos del conocimiento'));
C.push(p('El RANT toma como referencia la Clasificación Internacional Normalizada de la Educación (CINE 2013). Su Anexo I 2023 define, entre otros, estos campos detallados relevantes para IoT:'));
C.push(table(
  ['Campo amplio', 'Campo específico', 'Campo detallado (código)'],
  [
    ['06 Tecnologías de la información y la comunicación (TIC)', '1 Tecnologías de la información y la comunicación', '1 Ciencias computacionales (0611); 2 Diseño y administración de redes y bases de datos (0612); 3 Desarrollo y análisis de software y aplicaciones (0613); 81 Sistemas de información; 82 Auditoría de tecnologías'],
    ['07 Ingeniería, industria y construcción', '1 Ingeniería y profesiones afines', '4 Electrónica, automatización y sonido (0714); 82 Mecatrónica (0782); 84 Telecomunicaciones (0784)'],
  ],
  [2600, 2600, 3826],
));
C.push(gap());
C.push(p('El Art. 17 del RANT dice que en una carrera interdisciplinaria el campo amplio lo determina el tema principal, medido por créditos. El Art. 20 dice que si una carrera cubre dos o más campos detallados se clasifica en el de mayor número de créditos. Con la malla propuesta, los créditos disciplinares se reparten así:'));
const campoTot = {};
cursos.forEach((c) => { if (c.campo) campoTot[c.campo] = (campoTot[c.campo] || 0) + c.cr; });
const totDisc = Object.values(campoTot).reduce((a, b) => a + b, 0);
const campoRows = ['0611', '0612', '0613', '0714'].map((k) => [k, M.CAMPOS[k], String(campoTot[k]), pct(campoTot[k], totDisc)]);
campoRows.push(['06', 'Subtotal TIC (0611 + 0612 + 0613)', String(campoTot['0611'] + campoTot['0612'] + campoTot['0613']), pct(campoTot['0611'] + campoTot['0612'] + campoTot['0613'], totDisc)]);
C.push(table(['Código', 'Campo detallado', 'Créditos', '% de los créditos disciplinares'], campoRows, [1000, 4626, 1400, 2000], {
  aligns: [AlignmentType.CENTER, AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER], boldRow: (r) => r[0] === '06',
}));
C.push(gap());
C.push(p(`Se cuentan solo las asignaturas con contenido disciplinar (${totDisc} de ${T.cr} créditos); quedan fuera las ciencias básicas, la gestión, las optativas, las prácticas y la titulación.`));
C.push(p(`**Propuesta.** Clasificar la carrera en el campo amplio 06 (TIC), campo específico 1 y campo detallado 1, Ciencias computacionales (0611). Tres razones: el subtotal TIC (${campoTot['0611'] + campoTot['0612'] + campoTot['0613']} créditos) supera a la electrónica (${campoTot['0714']}), por lo que la regla del tema principal del Art. 17 lleva al campo amplio 06; dentro de TIC, 0611 es el detallado con más créditos (${campoTot['0611']}); y el Anexo II del RANT que se pudo consultar ya ubica el título de Técnico/a Superior en Internet de las Cosas en el campo detallado de Ciencias computacionales.`));
C.push(note('**Riesgo.** Si el CES lee el Art. 20 solo, sin pasar antes por el Art. 17, el detallado con más créditos de la malla es Electrónica, automatización y sonido (0714, ' + campoTot['0714'] + ' créditos), que cae en el campo amplio 07. Hay tres salidas: consultar por escrito a la Coordinación de Planificación Académica del CES antes de presentar; reequilibrar la malla para que 0611 iguale o supere a 0714; o clasificar en 0714 y aceptar el campo amplio de ingeniería. [[Decidir con el equipo de diseño y el CES]]'));
C.push(h2('5.4 Diferenciación frente a carreras cercanas'));
C.push(p('Frente a Software, IoT añade el hardware, la conectividad y la operación sobre el mundo físico. Frente a Telemática, añade el diseño de dispositivos, los sistemas embebidos y el ciclo completo del dato. Frente a Electricidad, se centra en la comunicación y el cómputo, no en la energía. [[Sustentar con la comparación de mallas del §4.7]]'));
C.push(h2('5.5 Denominación y título'));
C.push(p('Se propone "Ingeniería en Internet de las Cosas" con el título de "Ingeniero/a en Internet de las Cosas" (Art. 7 del RANT). Se evita la sigla inglesa IoT en la denominación oficial. El RANT permite anglicismos reconocidos en castellano (Disposición General Cuarta), pero la forma en castellano no exige argumentar ese punto.'));
C.push(p('**Estado en el Anexo II.** La copia del Anexo II 2023 que se pudo consultar, actualizada hasta el 15 de marzo de 2023, incluye "Técnico/a Superior en Internet de las Cosas" pero no un título de ingeniería de grado con ese nombre. La UEES Online cita una aprobación posterior (RPC-SO-26-No.428-2023) para "Ingeniero en Internet de las Cosas", lo que sugiere que el título se incorporó después de esa fecha, ya que el CES actualiza el anexo cuando aprueba nuevas denominaciones (Disposición General Quinta del RANT). Se debe verificar el Anexo II vigente:'));
[
  'Si el título ya consta, se adopta con esa redacción exacta y no hace falta justificación epistemológica.',
  'Si no consta, se presenta esta justificación (5.1 a 5.4) en el formato del Anexo III 2023.',
].forEach((t) => C.push(bl(t)));
C.push(h2('5.6 Referencias de esta sección'));
[
  'Ashton, K. (2009). That "Internet of Things" thing. RFID Journal.',
  'Lee, E. A. (2008). Cyber physical systems: design challenges. 11th IEEE International Symposium on Object and Component-Oriented Real-Time Distributed Computing.',
  'UIT-T (2012). Recomendación Y.2060: Visión general de la Internet de las cosas.',
  'von Bertalanffy, L. (1968). General System Theory. George Braziller.',
  'Weiser, M. (1991). The computer for the 21st century. Scientific American, 265(3).',
  'Wiener, N. (1948). Cybernetics: or Control and Communication in the Animal and the Machine. MIT Press.',
].forEach((t) => C.push(bl(t)));
C.push(note('Estas referencias no se verificaron en línea. [[Comprobar cada una antes de la presentación]]'));

// 6. Malla
C.push(h1('6. Malla curricular'));
C.push(h2('6.1 Verificación de parámetros normativos'));
C.push(table(
  ['Parámetro', 'Exigencia', 'Malla propuesta', 'Cumple'],
  [
    ['Créditos totales (Art. 15 RRA)', '120 a 150', `${T.cr}`, 'Sí'],
    ['Equivalencia (Art. 9)', '1 crédito = 48 h', `${fmt(T.H)} h en total`, 'Sí'],
    ['Períodos por año (Art. 10)', 'Al menos 2', '2 (8 períodos en 4 años)', 'Sí'],
    ['Dedicación semanal (Art. 10)', 'Promedio de 45 h', `${fmt(T.H / 8)} h por período = ${semanas45.toFixed(1).replace('.', ',')} semanas a 45 h`, 'Sí, si el período dura 18 semanas [[confirmar calendario UTEQ]]'],
    ['Prácticas laborales (Art. 43)', 'Mínimo 240 h', `${hLab} h`, 'Sí'],
    ['Servicio comunitario (Art. 43)', 'Mínimo 60 h', `${hServ} h`, 'Sí'],
    ['Tope de prácticas (Art. 43)', 'Máximo 10 % de las horas', pct(hPract, T.H), 'Sí'],
    ['Contacto con el docente (Arts. 55 y 56)', 'Al menos 16 h por crédito', `Mínimo ${minContacto.toFixed(1).replace('.', ',')} h por crédito`, 'Sí'],
    ['Interacción directa en modalidad presencial (Art. 55)', 'Al menos 51 % de los créditos', `${interaccion.toFixed(1).replace('.', ',')} % (contacto más práctico-experimental, sin prácticas preprofesionales)`, 'Sí'],
    ['Créditos de titulación (Art. 26)', 'Incluidos en el total', `${cursos.find((c) => c.codigo === '7.6').cr + cursos.find((c) => c.codigo === '8.3').cr} créditos`, 'Sí'],
  ],
  [2700, 2000, 2626, 1700],
));
C.push(gap());
C.push(note('La distribución de horas por componente es una hipótesis de diseño: 40 % contacto y 15 % práctico en asignaturas teóricas; 34 % y 32 % en las de laboratorio; 100 % práctico en las prácticas preprofesionales. Debe ajustarse asignatura por asignatura con el cuerpo docente.'));

C.push(h2('6.2 Malla por período'));
for (let i = 1; i <= 8; i++) {
  const lista = cursos.filter((c) => c.pao === i);
  const tot = lista.reduce((a, c) => ({ cr: a.cr + c.cr, H: a.H + c.H, c: a.c + c.c, p: a.p + c.p, a: a.a + c.a }), { cr: 0, H: 0, c: 0, p: 0, a: 0 });
  C.push(h3(`Período ${i}`));
  const rows = lista.map((c) => [c.codigo, c.nombre, String(c.cr), String(c.H), String(c.c), String(c.p), String(c.a)]);
  rows.push(['', 'Total del período', String(tot.cr), String(tot.H), String(tot.c), String(tot.p), String(tot.a)]);
  C.push(table(
    ['Cód.', 'Asignatura', 'Créd.', 'Horas', 'Contacto', 'Práctico', 'Autónomo'],
    rows,
    [600, 3826, 700, 800, 1000, 1000, 1100],
    { aligns: [AlignmentType.LEFT, AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER], boldRow: (r) => r[1] === 'Total del período' },
  ));
  C.push(gap());
}
C.push(h3('Total de la carrera'));
C.push(table(
  ['', 'Créditos', 'Horas'],
  [
    ['Aprendizaje en contacto con el docente', (T.c / 48).toFixed(1).replace('.', ','), fmt(T.c)],
    ['Aprendizaje práctico-experimental', (T.p / 48).toFixed(1).replace('.', ','), fmt(T.p)],
    ['Aprendizaje autónomo', (T.a / 48).toFixed(1).replace('.', ','), fmt(T.a)],
    ['Total', String(T.cr), fmt(T.H)],
  ],
  [5026, 2000, 2000],
  { aligns: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER], boldRow: (r) => r[0] === 'Total' },
));
C.push(gap());
C.push(note('Faltan por definir: los prerrequisitos y correquisitos, los códigos institucionales, los sílabos y los resultados de aprendizaje por asignatura. Son parte del expediente curricular, no de los campos de la plataforma.'));

C.push(h2('6.3 Optativas por itinerario'));
M.ITINERARIOS.forEach((it) => {
  C.push(p(`**${it.nombre}.**`));
  C.push(table(['Optativa', 'Período', 'Nombre'], it.optativas.map((o, i) => [['I', 'II', 'III'][i], ['7', '7', '8'][i], o]), [1500, 1200, 6326], { aligns: [AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.LEFT] }));
  C.push(gap());
});

// 7. Recursos
C.push(h1('7. Recursos, cohorte y convenios'));
C.push(h2('7.1 Infraestructura y equipamiento'));
C.push(p('La Guía exige una declaración firmada de que la UTEQ cuenta con recursos académicos, equipamiento, infraestructura y un modelo educativo acordes con la carrera. Esta lista define lo que la carrera necesita; el inventario real debe levantarlo la Facultad.'));
C.push(table(
  ['Espacio o recurso', 'Uso principal', 'Disponible hoy', 'Brecha'],
  [
    ['Laboratorio de electrónica y sistemas embebidos', 'Asignaturas 2.3, 3.2, 4.1, 4.2, 6.3', '[[Completar]]', '[[Completar]]'],
    ['Laboratorio de redes y comunicaciones inalámbricas', 'Asignaturas 3.3, 4.4, 5.1', '[[Completar]]', '[[Completar]]'],
    ['Laboratorio de cómputo, nube y ciencia de datos', 'Asignaturas 1.3 a 3.1, 4.3, 5.2, 5.4, 6.1', '[[Completar]]', '[[Completar]]'],
    ['Instrumentación: osciloscopios, fuentes, analizadores', 'Laboratorios de electrónica', '[[Completar]]', '[[Completar]]'],
    ['Kits de sensores, actuadores y placas de desarrollo', 'Todas las asignaturas de laboratorio', '[[Completar]]', '[[Completar]]'],
    ['Créditos de nube y licencias de software', 'Asignaturas 5.2, 5.4, 7.1', '[[Completar]]', '[[Completar]]'],
    ['Parcela o finca experimental con sensórica', 'Itinerario A; proyectos de vinculación', '[[Completar: aprovechar los campus agropecuarios]]', '[[Completar]]'],
    ['Biblioteca y bases de datos científicas', 'Toda la carrera', '[[Completar]]', '[[Completar]]'],
  ],
  [3000, 2600, 1800, 1626],
));
C.push(h2('7.2 Planta docente'));
C.push(p('El expediente debe declarar la planta docente que atenderá la carrera. Los perfiles requeridos son: electrónica y sistemas embebidos, redes y comunicaciones, ingeniería de software, ciencia de datos e inteligencia artificial, ciberseguridad, y ciencias básicas. [[Levantar docentes titulares y a contrato, con título y área de conocimiento, y comparar con los requisitos del modelo CACES]]'));
C.push(h2('7.3 Cohorte y convenios'));
C.push(p('La cohorte inicial se fija según la capacidad de los laboratorios (Art. 96 del RRA). [[Definir estudiantes por cohorte y por paralelo]]'));
C.push(table(
  ['Tipo de convenio', 'Entidades por contactar', 'Estado'],
  [
    ['Prácticas laborales', 'Empresas agroindustriales, de telecomunicaciones y de servicios de la zona', '[[Sin gestionar]]'],
    ['Servicio comunitario', 'Gobiernos autónomos descentralizados, cooperativas, comunidades rurales', '[[Sin gestionar]]'],
    ['Cooperación académica', 'Universidades y centros de investigación con IoT', '[[Sin gestionar]]'],
  ],
  [2400, 4626, 2000],
));

// 8. Pendientes
C.push(h1('8. Pendientes y hoja de ruta'));
C.push(h2('8.1 Pendientes de verificación'));
C.push(table(
  ['Pendiente', 'Por qué importa'],
  [
    ['Estatus CACES de la UTEQ', 'Define si el informe académico usa el Anexo 1 o si exige informe de par académico (Anexo 2)'],
    ['Anexo II vigente del RANT', 'Determina si el título "Ingeniero/a en Internet de las Cosas" ya existe. La copia consultada llega hasta el 15-III-2023 y solo trae el título de Técnico/a Superior en IoT'],
    ['Consulta al CES sobre el campo detallado', 'La clasificación en 0611 debe confirmarse con la Coordinación de Planificación Académica, por el peso de la electrónica en la malla (§5.3)'],
    ['Reforma 2024 del RRA', 'No se pudo descargar; verificar que no cambió los Arts. 15, 43 ni 96 a 99'],
    ['Estatuto y normativa interna UTEQ vigentes', 'Se leyó la versión 2019; confirmar órganos, nombres y trámite interno de diseño curricular'],
    ['Modelo CACES de evaluación de carreras', 'Necesario para la matriz de cumplimiento del §3.9'],
    ['Calendario académico UTEQ', 'Confirmar que el período de 18 semanas cuadra con las 45 h semanales'],
    ['Formato del Anexo III del RANT', 'La justificación epistemológica debe presentarse en ese formato'],
    ['Base de matrícula TIC de la SENESCYT', 'Su servidor no respondió; hace falta para reemplazar los proxies de la oferta de graduados (4.10)'],
    ['Graduados de Los Ríos por tipo de bachillerato (AMIE)', 'Dimensiona los grupos de la nivelación (3.2)'],
    ['Encuesta a empleadores', 'Reemplaza los supuestos de adopción del modelo de demanda (4.9)'],
    ['Nivelación por carrera en la UTEQ', 'Confirmar con Admisión y Nivelación que la propuesta de 400 horas es viable'],
  ],
  [3200, 5826],
));
C.push(h2('8.2 Hoja de ruta'));
C.push(p('Ruta interna de la UTEQ según el Estatuto 2019, seguida del trámite ante el CES (Arts. 96 a 101 del RRA):'));
[
  'Equipo de diseño curricular: completar este borrador, el estudio de campo y la comparación con Telemática.',
  'Consejo Directivo de la Facultad: propone el proyecto de diseño de carrera al Consejo Académico.',
  'Consejo Académico: analiza y sugiere al Consejo Universitario que solicite la creación al CES.',
  'Consejo Universitario: emite la resolución del órgano colegiado superior.',
  'Rectorado: carga el expediente en la plataforma PPCP del CES.',
  'CES: verificación de requisitos en 15 días, más 10 para subsanar; acuerdo de la Comisión en 7 días; resolución del Pleno en 5 días.',
  'Notificación, registro en el SNIESE y apertura de la oferta.',
].forEach((t, i) => C.push(new Paragraph({ numbering: { reference: 'num', level: 0 }, children: runs(t), spacing: { after: 60, line: 276 } })));

// Anexo
C.push(h1('Anexo A. Fuentes'));
[
  'Consejo de Educación Superior. Reglamento de Régimen Académico (2022, reforma RPC-SE-03-No.008-2023).',
  'Consejo de Educación Superior. Guía Metodológica para la presentación de carreras y programas, ajustes curriculares sustantivos y no sustantivos (RPC-SE-14-No.038-2022, codificada en 2023).',
  'Consejo de Educación Superior. Reglamento de Armonización de la Nomenclatura de Títulos Profesionales y Grados Académicos (RPC-SE-05-No.014-2023).',
  'Ministerio de Telecomunicaciones y de la Sociedad de la Información (2025). Política Pública para la Transformación Digital del Ecuador 2025-2030. Datos de SENESCYT (2024) e INEC (2024) citados en ella.',
  'Universidad Técnica Estatal de Quevedo. Estatuto (2019). Plan Estratégico de Desarrollo Institucional 2021-2025, actualización de diciembre de 2024.',
  'Universidad Técnica Estatal de Quevedo. Memorando UTEQ-VICACAD-2026-1787-M, de 02 de septiembre de 2026.',
  'UEES Online. Carrera de Internet de las Cosas. Consultada el 26-IX-2026.',
  'Consejo de Educación Superior. Anexo I 2023 y Anexo II 2023 (actualizado hasta el 15-III-2023) del RANT.',
  'INEC. Encuesta de Superficie y Producción Agropecuaria Continua (ESPAC) 2024, abril de 2025, y Módulo de Información Ambiental y Tecnificación Agropecuaria (MIATA) 2024.',
  'INEC. Encuesta Nacional de Empleo, Desempleo y Subempleo (ENEMDU) anual 2025, y tabulados de Tecnologías de la Información y Comunicación, julio de 2025.',
  'ARCOTEL. Boletín estadístico, cierre de año 2024 (Boletín 2025-01).',
  'UTEQ. Unidad de Admisión y Nivelación: proceso de admisión 2026 y evaluación por competencias generales y específicas (uteq.edu.ec/admision).',
  'CACES. Modelo genérico de evaluación del entorno de aprendizaje de carreras (criterios según el sitio institucional).',
  'SRI. Registro Único de Contribuyentes, Los Ríos (datos abiertos, actualizado el 1-IX-2026). INEC. Registro Estadístico de Empresas 2025. INEC. Censo de Población y Vivienda 2022.',
  'Ministerio de Educación, Deporte y Cultura. Bachillerato General, Bachillerato en Ciencias y Bachillerato Técnico (educacion.gob.ec); Acuerdo MINEDUC-2024-00065-A. Primicias (2024): materias y horas de clase del año lectivo 2024-2025.',
  'UTEQ. Informe de rendición de cuentas 2024. Redalyc: Nivelación propuesta por la SENESCYT, vivencias en la Universidad Central del Ecuador.',
].forEach((t) => C.push(bl(t)));

// ---------- documento ----------
const doc = new Document({
  creator: 'Borrador de trabajo',
  title: 'Proyecto de creación de carrera: Ingeniería en Internet de las Cosas (borrador v0.2)',
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 32, bold: true, font: FONT, color: COLOR }, paragraph: { spacing: { before: 240, after: 200 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 26, bold: true, font: FONT, color: COLOR }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 23, bold: true, font: FONT, color: '404040' }, paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      { reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
      { reference: 'num', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 360 } } } }] },
    ],
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1300, left: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'Borrador v0.2 · Ingeniería en Internet de las Cosas · UTEQ', font: FONT, size: 17, color: '7F7F7F' })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ['Página ', PageNumber.CURRENT, ' de ', PageNumber.TOTAL_PAGES], font: FONT, size: 17, color: '7F7F7F' })] })] }) },
    children: C,
  }],
});

const out = path.join(__dirname, '..', 'Proyecto_Carrera_IoT_borrador_v0.2.docx');
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log('OK', out, buf.length); });
