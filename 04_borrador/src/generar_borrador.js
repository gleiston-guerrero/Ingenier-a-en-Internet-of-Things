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
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 900 }, children: [new TextRun({ text: 'BORRADOR v0.1 · 26 de septiembre de 2026', font: FONT, size: 24, bold: true })] }));
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
    ['Campo amplio, específico y detallado', '[[Definir tras revisar el Anexo I 2023 del RANT]]. Hipótesis: campo amplio 06, Tecnologías de la información y la comunicación', 'Ver §5.3, con el conteo de créditos por área'],
    ['Itinerarios', '2: A) IoT para agroindustria y ambiente; B) IoT industrial y ciudades inteligentes', 'Art. 16 RRA, máximo 3'],
    ['Perfil de ingreso', 'Título de bachiller, con acceso por el sistema de nivelación y admisión', 'Art. 13 RRA. Detalle en §3.2'],
    ['Perfil de egreso', 'Ocho resultados de aprendizaje', '§3.3'],
    ['Segunda lengua', '[[Definir: p. ej. suficiencia de inglés como requisito de titulación]]', 'Art. 64 RRA. Depende de la normativa UTEQ'],
    ['Lugar de ejecución', 'Campus Central, Quevedo, Los Ríos [[confirmar campus]]', 'Dirección institucional según memorando UTEQ-VICACAD-2026-1787-M'],
    ['Estudiantes por cohorte', '[[Definir según capacidad de laboratorios; sugerencia inicial: 1 paralelo]]', 'Art. 96 RRA'],
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
C.push(p('Bachiller o equivalente, que haya cumplido los requisitos del Sistema de Nivelación y Admisión (Art. 13 del RRA). Se valoran, sin ser requisito, el razonamiento lógico-matemático, el interés por la tecnología y el trabajo experimental. [[Alinear con el perfil de ingreso y el proceso de nivelación vigentes en la UTEQ]]'));

C.push(h2('3.3 Perfil de egreso'));
C.push(p('El egresado o egresada logra los siguientes resultados de aprendizaje. La columna de la derecha indica las asignaturas de la malla (§6) que los desarrollan.'));
C.push(table(
  ['Código', 'Resultado de aprendizaje', 'Asignaturas principales'],
  M.RA.map(([c, t, as]) => [c, t, as.join(', ')]),
  [900, 5926, 2200],
));
C.push(gap());
C.push(p('**Campos de desempeño profesional.** Ingeniero o ingeniera de soluciones IoT; arquitecto o arquitecta de soluciones conectadas; desarrollador o desarrolladora de firmware y sistemas embebidos; especialista en redes de sensores y conectividad; analista de datos de dispositivos; consultor o consultora en agricultura de precisión y agroindustria; emprendedor o emprendedora de base tecnológica. [[Contrastar con el estudio de demanda ocupacional del §4.5]]'));

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
C.push(p('El informe debe describir, para cada criterio y estándar del modelo de evaluación de carreras del CACES, cómo lo cumple el proyecto. Este borrador no incorpora esa matriz porque no se pudo consultar el modelo vigente.'));
C.push(table(
  ['Criterio o estándar CACES', 'Descripción de cumplimiento', 'Evidencia'],
  [
    ['[[Tomar del modelo CACES vigente]]', '[[Completar]]', '[[Completar]]'],
    ['[[Tomar del modelo CACES vigente]]', '[[Completar]]', '[[Completar]]'],
    ['[[Tomar del modelo CACES vigente]]', '[[Completar]]', '[[Completar]]'],
  ],
  [3000, 3826, 2200],
));

// 4. Pertinencia
C.push(h1('4. Estudio de pertinencia'));
C.push(note('Este estudio tiene dos partes. Las secciones 4.1 a 4.4 usan datos oficiales verificados. Las secciones 4.5 a 4.6 son el plan del trabajo de campo, que todavía no se hizo: sin él no hay demanda ocupacional demostrada.'));
C.push(h2('4.1 Marco'));
C.push(p('El Art. 107 de la LOES define la pertinencia como la respuesta de la educación superior a las necesidades de la sociedad, a la planificación nacional y a la prospectiva de desarrollo científico y tecnológico, y exige articular la oferta con las necesidades locales y regionales, las tendencias del mercado ocupacional y la estructura productiva de la provincia y la región. La Guía Metodológica pide sustentar la demanda estudiantil y la ocupacional con fuentes primarias y secundarias oficiales.'));

C.push(h2('4.2 Contexto nacional de transformación digital'));
C.push(p('La Política Pública para la Transformación Digital del Ecuador 2025-2030, del Ministerio de Telecomunicaciones y de la Sociedad de la Información, incluye el Internet de las Cosas entre las tendencias de transformación digital (p. 46) y señala que las tecnologías emergentes, entre ellas el IoT, pueden aplicarse al desarrollo sostenible en áreas como la salud, la energía y la agricultura (p. 82).'));
C.push(p('El mismo documento afirma que Ecuador sufre un déficit de talento digital, y que muchos graduados en áreas TIC carecen de experiencia y de actualización en nuevas herramientas (pp. 51 y 52). Recoge además que solo el 15,27 % de la población tenía habilidades avanzadas para instalar y configurar software en 2019, según el INEC (p. 55).'));

C.push(h2('4.3 Oferta nacional de carreras TIC'));
C.push(p('Según los datos de la SENESCYT recogidos en esa política (Tabla 7, p. 52 y 53), en 2024 había 197 carreras TIC en el país. La distribución muestra una oferta concentrada en software y muy escasa en IoT:'));
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
C.push(p('El mismo informe cita 7.476 graduados en carreras TIC en 2021, frente a 5.079 en 2020 (texto de la p. 53); su Tabla 8 consigna 7.485 para 2021, una discrepancia menor en la fuente. [[Verificar contra las cifras originales de la SENESCYT antes de citar]]'));
C.push(note('Lectura provisional: la oferta IoT es marginal frente al software, lo que apoya la pertinencia de una carrera especializada. Pero es un argumento de oferta, no de demanda: falta demostrar que los empleadores contratarán a estos profesionales.'));

C.push(h2('4.4 Oferta de la UTEQ y análisis de traslape'));
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

C.push(h2('4.5 Demanda estudiantil y ocupacional: plan de trabajo de campo'));
C.push(p('**Fuentes secundarias oficiales por consultar:**'));
[
  'SENESCYT: oferta, matrícula y graduados en TIC y en el campo detallado que resulte.',
  'INEC: empleo, subempleo y estructura de ocupaciones (ENEMDU), y uso de TIC en empresas.',
  'MAG e INEC: encuesta de superficie y producción agropecuaria (ESPAC) para dimensionar el sector agropecuario de Los Ríos.',
  'MINTEL y ARCOTEL: conectividad y transformación digital, incluida la rural.',
  'Superintendencia de Compañías: empresas por actividad económica en la zona.',
  'Plan Nacional de Desarrollo vigente y plan de desarrollo y ordenamiento territorial de Los Ríos.',
].forEach((t) => C.push(bl(t)));
C.push(p('**Fuentes primarias por levantar:**'));
[
  'Encuesta a empleadores de los sectores agropecuario, agroindustrial, industrial y de servicios públicos de la zona: necesidad de perfiles IoT, vacantes previstas, competencias más demandadas y salarios de referencia.',
  'Encuesta a estudiantes de último año de bachillerato de la zona: interés en la carrera.',
  'Entrevistas a expertos y a graduados de Software, Telemática y Electricidad de la UTEQ.',
  'Revisión de ofertas de empleo publicadas para perfiles IoT, con fecha y fuente.',
].forEach((t) => C.push(bl(t)));
C.push(p('**Muestra.** Para poblaciones finitas se propone n = N·z²·p·q / (e²·(N−1) + z²·p·q), con z = 1,96, p = q = 0,5 y error e entre 5 % y 8 %. [[Definir N para cada población]]'));
C.push(table(
  ['Indicador que debe salir del estudio', 'Valor', 'Fuente'],
  [
    ['Empresas e instituciones de la zona con necesidad de perfiles IoT', '[[Completar]]', 'Encuesta a empleadores'],
    ['Vacantes previstas a 3 y 5 años', '[[Completar]]', 'Encuesta a empleadores'],
    ['Bachilleres interesados en la carrera', '[[Completar]]', 'Encuesta a bachilleres'],
    ['Graduados TIC de la zona por año', '[[Completar]]', 'SENESCYT'],
    ['Brecha entre demanda y oferta', '[[Completar]]', 'Cálculo propio'],
  ],
  [4800, 1500, 2726],
));

C.push(h2('4.6 Conclusión provisional'));
C.push(p('Con la evidencia disponible, la carrera es coherente con la política de transformación digital y llena un vacío en la oferta nacional. Su aprobación depende de dos cosas que este borrador no resuelve: la demanda ocupacional medida en el territorio y la diferenciación frente a Telemática. [[Reescribir esta sección cuando exista el trabajo de campo]]'));

// 5. Justificación epistemológica
C.push(h1('5. Justificación epistemológica'));
C.push(note('El Art. 25 del RANT exige esta justificación solo si la denominación o la titulación no constan en el anexo vigente, y debe presentarse en el formato del Anexo III 2023, que no se pudo consultar. El contenido de abajo está listo para transponerse a ese formato.'));
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
C.push(p('El RANT toma como referencia la Clasificación Internacional Normalizada de la Educación (CINE). Cuando una carrera cubre varios campos detallados, se clasifica en el de mayor número de créditos (Art. 20). La malla propuesta reparte así sus créditos:'));
const areaRows = ['SW', 'ELE', 'RED', 'BAS', 'GES', 'OPT', 'INT'].map((a) => [M.AREAS[a], String(porArea[a]), pct(porArea[a], T.cr)]);
C.push(table(['Área de la malla', 'Créditos', '% del total'], areaRows.concat([['Total', String(T.cr), '100,0 %']]), [5026, 2000, 2000], {
  aligns: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER], boldRow: (r) => r[0] === 'Total',
}));
C.push(gap());
C.push(p(`Software y datos (${porArea.SW} créditos) más redes y comunicaciones (${porArea.RED}) suman ${porArea.SW + porArea.RED} créditos de contenido TIC, frente a ${porArea.ELE} de electrónica y automatización. Esto sostiene el campo amplio de Tecnologías de la información y la comunicación. Pero por sí sola, la regla del Art. 20 llevaría el campo detallado hacia software, lo que no describe bien una carrera de IoT. El equipo debe decidir entre tres caminos: rebalancear la malla, elegir el campo detallado que mejor represente la carrera y justificarlo, o, si ningún campo detallado vigente sirve, solicitar uno nuevo con el Anexo IV (Art. 26 del RANT). [[Decidir tras revisar el Anexo I 2023]]`));
C.push(h2('5.4 Diferenciación frente a carreras cercanas'));
C.push(p('Frente a Software, IoT añade el hardware, la conectividad y la operación sobre el mundo físico. Frente a Telemática, añade el diseño de dispositivos, los sistemas embebidos y el ciclo completo del dato. Frente a Electricidad, se centra en la comunicación y el cómputo, no en la energía. [[Sustentar con la comparación de mallas del §4.4]]'));
C.push(h2('5.5 Denominación y título'));
C.push(p('Se propone "Ingeniería en Internet de las Cosas" con el título de "Ingeniero/a en Internet de las Cosas". Se evita la sigla inglesa IoT en la denominación oficial. El RANT permite anglicismos reconocidos en castellano (Disposición General Cuarta), pero la forma en castellano no exige argumentar ese punto. Si el CES ya registra esta denominación en el Anexo II, esta justificación no es necesaria. [[Verificar en el Anexo II vigente]]'));
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
    ['Anexo I y Anexo II vigentes del RANT', 'Determinan el campo detallado y si la denominación ya existe'],
    ['Reforma 2024 del RRA', 'No se pudo descargar; verificar que no cambió los Arts. 15, 43 ni 96 a 99'],
    ['Estatuto y normativa interna UTEQ vigentes', 'Se leyó la versión 2019; confirmar órganos, nombres y trámite interno de diseño curricular'],
    ['Modelo CACES de evaluación de carreras', 'Necesario para la matriz de cumplimiento del §3.9'],
    ['Calendario académico UTEQ', 'Confirmar que el período de 18 semanas cuadra con las 45 h semanales'],
    ['Formato del Anexo III del RANT', 'La justificación epistemológica debe presentarse en ese formato'],
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
].forEach((t) => C.push(bl(t)));

// ---------- documento ----------
const doc = new Document({
  creator: 'Borrador de trabajo',
  title: 'Proyecto de creación de carrera: Ingeniería en Internet de las Cosas (borrador v0.1)',
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
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'Borrador v0.1 · Ingeniería en Internet de las Cosas · UTEQ', font: FONT, size: 17, color: '7F7F7F' })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ['Página ', PageNumber.CURRENT, ' de ', PageNumber.TOTAL_PAGES], font: FONT, size: 17, color: '7F7F7F' })] })] }) },
    children: C,
  }],
});

const out = path.join(__dirname, '..', 'Proyecto_Carrera_IoT_borrador_v0.1.docx');
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log('OK', out, buf.length); });
