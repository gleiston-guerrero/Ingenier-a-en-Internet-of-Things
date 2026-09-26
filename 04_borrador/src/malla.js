// Datos de la malla propuesta. Todo total del documento se calcula desde aquí.
// Tipos de distribución de horas (proporciones contacto / práctico-experimental / autónomo):
//   T = teoría, L = laboratorio, P = práctica preprofesional, X = titulación
// Áreas (para el conteo por campo, Art. 20 RANT): BAS ciencias básicas, SW software y datos,
//   RED redes y comunicaciones, ELE electrónica y automatización, GES gestión/ética/humanística,
//   INT integración (prácticas, proyectos, titulación), OPT optativas de itinerario.
const RATIOS = {
  T: { c: 0.40, p: 0.15 },
  L: { c: 0.34, p: 0.32 },
  P: { c: 0.0, p: 1.0 },
  X: { c: 0.34, p: 0.16 },
};

const AREAS = {
  BAS: 'Ciencias básicas',
  SW: 'Software y datos',
  RED: 'Redes y comunicaciones',
  ELE: 'Electrónica y automatización',
  GES: 'Gestión, ética y comunicación',
  INT: 'Integración (prácticas, proyectos, titulación)',
  OPT: 'Optativas de itinerario',
};

// [código, nombre, créditos, tipo, área]
const PAO = [
  [
    ['1.1', 'Cálculo diferencial e integral', 4, 'T', 'BAS'],
    ['1.2', 'Álgebra lineal y matemática discreta', 3, 'T', 'BAS'],
    ['1.3', 'Fundamentos de programación', 4, 'L', 'SW'],
    ['1.4', 'Física para ingeniería', 3, 'T', 'BAS'],
    ['1.5', 'Comunicación académica y técnicas de estudio', 2, 'T', 'GES'],
    ['1.6', 'Introducción a la ingeniería del Internet de las Cosas', 1, 'T', 'INT'],
  ],
  [
    ['2.1', 'Cálculo multivariable y ecuaciones diferenciales', 3, 'T', 'BAS'],
    ['2.2', 'Programación orientada a objetos', 4, 'L', 'SW'],
    ['2.3', 'Circuitos eléctricos y electrónica analógica', 4, 'L', 'ELE'],
    ['2.4', 'Probabilidad y estadística', 3, 'T', 'BAS'],
    ['2.5', 'Sistemas digitales', 3, 'L', 'ELE'],
  ],
  [
    ['3.1', 'Estructuras de datos y algoritmos', 4, 'L', 'SW'],
    ['3.2', 'Electrónica digital y microcontroladores', 4, 'L', 'ELE'],
    ['3.3', 'Redes de computadoras', 4, 'L', 'RED'],
    ['3.4', 'Señales y sistemas', 3, 'T', 'ELE'],
    ['3.5', 'Ética profesional y responsabilidad social', 2, 'T', 'GES'],
  ],
  [
    ['4.1', 'Sistemas embebidos y programación de firmware', 4, 'L', 'ELE'],
    ['4.2', 'Sensores, actuadores e instrumentación', 4, 'L', 'ELE'],
    ['4.3', 'Bases de datos', 3, 'L', 'SW'],
    ['4.4', 'Protocolos y arquitecturas de IoT', 4, 'L', 'RED'],
    ['4.5', 'Prácticas de servicio comunitario', 2, 'P', 'INT'],
  ],
  [
    ['5.1', 'Comunicaciones inalámbricas y redes de sensores', 4, 'L', 'RED'],
    ['5.2', 'Plataformas en la nube y computación en el borde', 4, 'L', 'SW'],
    ['5.3', 'Ciberseguridad y privacidad en IoT', 3, 'L', 'RED'],
    ['5.4', 'Ciencia de datos y series temporales', 3, 'L', 'SW'],
    ['5.5', 'Ingeniería de software y de requisitos', 3, 'T', 'SW'],
  ],
  [
    ['6.1', 'Inteligencia artificial en el borde (TinyML)', 4, 'L', 'SW'],
    ['6.2', 'Sistemas ciberfísicos, control y automatización', 4, 'L', 'ELE'],
    ['6.3', 'Diseño de hardware y placas de circuito impreso', 4, 'L', 'ELE'],
    ['6.4', 'Gestión de proyectos y emprendimiento tecnológico', 3, 'T', 'GES'],
    ['6.5', 'Metodología de la investigación', 2, 'T', 'GES'],
  ],
  [
    ['7.1', 'Arquitectura y despliegue de soluciones IoT a escala', 3, 'L', 'SW'],
    ['7.2', 'Gemelos digitales y simulación', 3, 'L', 'SW'],
    ['7.3', 'Optativa I (itinerario)', 3, 'L', 'OPT'],
    ['7.4', 'Optativa II (itinerario)', 3, 'L', 'OPT'],
    ['7.5', 'Prácticas preprofesionales laborales I', 3, 'P', 'INT'],
    ['7.6', 'Diseño del proyecto de titulación', 2, 'X', 'INT'],
  ],
  [
    ['8.1', 'Optativa III (itinerario)', 3, 'L', 'OPT'],
    ['8.2', 'Prácticas preprofesionales laborales II', 3, 'P', 'INT'],
    ['8.3', 'Trabajo de integración curricular', 8, 'X', 'INT'],
    ['8.4', 'Legislación TIC y protección de datos personales', 3, 'T', 'GES'],
  ],
];

const ITINERARIOS = [
  {
    nombre: 'Itinerario A: IoT para agroindustria y ambiente',
    optativas: [
      'Agricultura de precisión y teledetección',
      'Monitoreo ambiental y gestión del agua',
      'Trazabilidad y cadena de frío agroindustrial',
    ],
  },
  {
    nombre: 'Itinerario B: IoT industrial y ciudades inteligentes',
    optativas: [
      'IoT industrial y mantenimiento predictivo',
      'Energía, edificios y ciudades inteligentes',
      'Redes privadas 5G y sistemas de misión crítica',
    ],
  },
];

// Resultados de aprendizaje del perfil de egreso y asignaturas que los desarrollan (códigos)
const RA = [
  ['RA1', 'Diseña arquitecturas IoT de extremo a extremo (percepción, red, plataforma y aplicación) a partir de requisitos del usuario y de normas técnicas.', ['1.6', '4.4', '5.5', '7.1']],
  ['RA2', 'Desarrolla hardware y firmware para dispositivos embebidos con sensores y actuadores, optimizando consumo y fiabilidad.', ['2.3', '3.2', '4.1', '4.2', '6.3']],
  ['RA3', 'Implementa la conectividad de los dispositivos con protocolos y tecnologías de comunicación adecuados al contexto de despliegue.', ['3.3', '4.4', '5.1']],
  ['RA4', 'Integra plataformas en la nube y en el borde para almacenar, procesar y visualizar datos de dispositivos conectados.', ['4.3', '5.2', '5.4', '7.1']],
  ['RA5', 'Aplica modelos de analítica e inteligencia artificial a datos de sensores, incluida su ejecución en dispositivos de recursos limitados.', ['2.4', '5.4', '6.1', '7.2']],
  ['RA6', 'Protege sistemas IoT aplicando criterios de ciberseguridad y de privacidad, y la normativa ecuatoriana de protección de datos personales.', ['5.3', '8.4']],
  ['RA7', 'Gestiona proyectos de base tecnológica con criterios de sostenibilidad, ética profesional y emprendimiento.', ['3.5', '6.4', '7.5', '8.2']],
  ['RA8', 'Investiga e innova con método científico y comunica sus resultados de forma clara, en equipos multidisciplinarios.', ['1.5', '6.5', '7.6', '8.3']],
];

function build() {
  const cursos = [];
  PAO.forEach((lista, i) => {
    lista.forEach(([codigo, nombre, cr, tipo, area]) => {
      const H = 48 * cr;
      const c = Math.round(H * RATIOS[tipo].c);
      const p = Math.round(H * RATIOS[tipo].p);
      cursos.push({ pao: i + 1, codigo, nombre, cr, tipo, area, H, c, p, a: H - c - p });
    });
  });
  return cursos;
}

function totales(cursos) {
  const t = { cr: 0, H: 0, c: 0, p: 0, a: 0 };
  cursos.forEach((x) => { t.cr += x.cr; t.H += x.H; t.c += x.c; t.p += x.p; t.a += x.a; });
  return t;
}

module.exports = { PAO, AREAS, ITINERARIOS, RA, build, totales };
