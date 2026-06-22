// Genera el CV en espanol con el MISMO formato que el CV en ingles:
// fuente Arial, sin tablas, sin cuadros, parrafos limpios (compatible con ATS).
const { Document, Packer, Paragraph, TextRun, AlignmentType } = require('docx');
const fs = require('fs');

const FONT = 'Arial';

// Tamanos en half-points (docx usa half-points): 12pt = 24, 10.5pt = 21
const SZ_HEAD = 24;   // 12pt
const SZ_BODY = 21;   // 10.5pt

// --- helpers ---------------------------------------------------------------

function blank() {
  return new Paragraph({ children: [new TextRun({ text: '', font: FONT, size: SZ_BODY })] });
}

// Linea de contacto: etiqueta en negrita + valor normal
function contact(label, value) {
  return new Paragraph({
    spacing: { after: 0 },
    children: [
      new TextRun({ text: label + ' ', bold: true, size: SZ_HEAD, font: FONT }),
      new TextRun({ text: value, size: SZ_HEAD, font: FONT }),
    ]
  });
}

// Titulo de seccion (negrita 12pt)
function section(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text: text, bold: true, size: SZ_HEAD, font: FONT })]
  });
}

// Linea en negrita 10.5pt (cargo, empresa, fechas)
function boldLine(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 0 },
    children: [new TextRun({ text: text, bold: true, size: SZ_BODY, font: FONT })]
  });
}

// Parrafo de descripcion (10.5pt, justificado)
function body(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 60 },
    children: [new TextRun({ text: text, size: SZ_BODY, font: FONT })]
  });
}

// --- contenido -------------------------------------------------------------

const children = [];

// Encabezado
children.push(new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: 'Faiber Andrés Montes Gómez', bold: true, size: SZ_HEAD, font: FONT })] }));
children.push(new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: 'Ingeniero Financiero', bold: true, size: SZ_HEAD, font: FONT })] }));
children.push(contact('Teléfono:', '+57 3504747211'));
children.push(contact('Correo:', 'andresmono1986@hotmail.com'));
children.push(contact('LinkedIn:', 'https://www.linkedin.com/in/faiber-andres-montes'));
children.push(contact('Github:', 'https://github.com/faiber1986'));
children.push(contact('Ubicación:', 'Medellín, Antioquia, Colombia'));
children.push(blank());

// Titular profesional
children.push(new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { after: 0 },
  children: [new TextRun({ text: 'Especialista Senior en Datos | Especialista en IA & LLM | Ingeniero Financiero | RAG Pipelines, MLOps, Snowflake & AWS | Construyendo Arquitecturas de Datos Escalables para Fintech Global', bold: true, size: SZ_HEAD, font: FONT })]
}));
children.push(blank());

// Perfil profesional
children.push(section('PERFIL PROFESIONAL'));
children.push(body('Lidero el "Production Premium" escalando modelos de IA en APIs de alta disponibilidad con Python, FastAPI, Docker y Snowflake. Catalicé la habilitación de IA en CESDE capacitando a más de 1.000 líderes, impulsando ganancias de productividad del 25%. Optimicé modelos de riesgo en N-iX (reducción del 40% de latencia) y generé ahorros de más de $200K en Colanta. Experto en RAG y MLOps, conectando el riesgo financiero de alto impacto con IA lista para producción.'));
children.push(blank());

// Experiencia laboral
children.push(section('EXPERIENCIA LABORAL'));

children.push(boldLine('Especialista en Datos e Ingeniero de IA'));
children.push(boldLine('Freelance. Medellín.'));
children.push(boldLine('Marzo 2026 - Actualidad'));
children.push(body('Diseñé agentes automatizados de investigación y generación de contenido con LangChain y las APIs de OpenAI/Claude, reduciendo los flujos administrativos manuales en 15 horas por semana.'));
children.push(body('Implementé protocolos de gobernanza y seguridad en IA, educando a los stakeholders en mitigación de alucinaciones, detección de sesgos y privacidad ética de datos en despliegues de LLM.'));
children.push(body('Construí y desplegué chatbots RAG personalizados y agentes de análisis financiero con FastAPI y bases de datos vectoriales (Pinecone, Chroma), automatizando el 100% del reporte semanal para clientes corporativos.'));
children.push(blank());

children.push(boldLine('Instructor de IA.'));
children.push(boldLine('CESDE. Medellín.'));
children.push(boldLine('Octubre - Diciembre 2025'));
children.push(body('Lideré la estrategia de habilitación en IA para más de 500 profesionales no técnicos, enseñando Ingeniería de Prompts avanzada y flujos RAG (Generación Aumentada por Recuperación) para aumentar la productividad operativa en un 25%.'));
children.push(body('Desarrollé y dicté un currículo sobre IA Responsable, enfocado en mitigación de sesgos, privacidad de datos e integración ética de LLMs (ChatGPT, Claude, Gemini) para el liderazgo corporativo.'));
children.push(body('Automaticé flujos administrativos y de investigación para los estudiantes con herramientas de IA no-code/low-code, reduciendo el tiempo de tareas manuales en promedio 15 horas por semana por estudiante.'));
children.push(blank());

children.push(boldLine('Ingeniero de Datos y Desarrollador Python.'));
children.push(boldLine('N-iX. Medellín.'));
children.push(boldLine('Diciembre 2024 - Agosto 2025'));
children.push(body('Orquesté la migración de modelos complejos de riesgo crediticio de R legado a Python de grado productivo (FastAPI), reduciendo la latencia de ejecución en un 40% y habilitando el scoring en tiempo real vía API.'));
children.push(body('Arquitecté pipelines ETL/ELT automatizados en AWS (Glue, S3, Redshift) para procesar datasets logísticos de múltiples terabytes, garantizando un 99.9% de integridad de datos para reportes de cadena de suministro global.'));
children.push(body('Diseñé una suite de pruebas robusta con Pytest, alcanzando un 90% de cobertura de código y eliminando bugs críticos en catálogos de datos financieros de alto impacto.'));
children.push(body('Optimicé el rendimiento de consultas SQL y el modelado de datos (Star Schema), logrando una reducción del 25% en costos de cómputo en la nube y dashboards de BI más rápidos.'));
children.push(blank());

children.push(boldLine('Especialista en Datos e Ingeniero Financiero.'));
children.push(boldLine('Freelance. Medellín.'));
children.push(boldLine('Febrero 2023 - Diciembre 2024'));
children.push(body('Diseñé e implementé dashboards financieros personalizados y REST APIs (FastAPI) para clientes corporativos, automatizando el 100% de los ciclos de reporte financiero semanal.'));
children.push(body('Construí modelos predictivos de series de tiempo para el pronóstico de variables económicas, superando los benchmarks manuales previos en un 20% de precisión predictiva.'));
children.push(body('Administré entornos de Snowflake & AWS para empresas medianas, implementando mejores prácticas de DataOps que redujeron el tiempo de inactividad de datos en un 65%.'));
children.push(blank());

children.push(boldLine('Científico de Datos y Analista de Project Finance.'));
children.push(boldLine('Colanta. Medellín.'));
children.push(boldLine('Agosto 2021 - Enero 2023'));
children.push(body('Desarrollé y desplegué modelos de predicción de demanda con PyTorch y Scikit-learn, aumentando la precisión en un 15% y generando ahorros estimados de $200K en desperdicios logísticos anuales.'));
children.push(body('Utilicé PySpark en AWS para procesar y analizar datasets contables de gran escala, identificando más de $50K en irregularidades fiscales mediante ingeniería de características avanzada.'));
children.push(body('Visualicé KPIs financieros críticos para stakeholders ejecutivos con Power BI y Qlik Sense, reduciendo el tiempo de obtención de insights en evaluaciones de riesgo trimestrales de días a minutos.'));
children.push(blank());

// Formacion academica
children.push(section('FORMACIÓN ACADÉMICA'));
children.push(body('Posgrado en Gestión de Riesgo Financiero. Universidad de Medellín. Medellín. 2021'));
children.push(body('(Modelado de Machine Learning, Modelado de Valor en Riesgo (VaR) y Análisis de Riesgo Estocástico)'));
children.push(blank());
children.push(body('Posgrado en Legislación Tributaria. UNAULA. Medellín. 2020'));
children.push(body('(Algoritmos de Cumplimiento Tributario, Reporte Regulatorio Automatizado e Integridad de Datos Fiscales)'));
children.push(blank());
children.push(body('Ingeniero Financiero. Universidad de Medellín. Medellín. 2018.'));
children.push(blank());

// Otra formacion academica
children.push(section('OTRA FORMACIÓN ACADÉMICA'));
children.push(body('Front End Developer. ONE-Oracle. 2024'));
children.push(body('Data Science 4 All. Correlation One. 2022'));
children.push(body('Data Scientist. Digital House. 2022'));
children.push(body('Data Scientist. Universidad EAFIT. 2021'));
children.push(body('Front End Developer. Comfenalco Antioquia. 2021'));
children.push(body('Certificados - Coursera, LinkedIn Learning. 2020 - 2023'));

// --- documento -------------------------------------------------------------

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: SZ_BODY } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    children: children
  }]
});

Packer.toBuffer(doc).then(function(buffer) {
  fs.writeFileSync('assets/Faiber-Andres-Montes-CV-ES.docx', buffer);
  console.log('OK: CV en espanol generado (formato identico al ingles).');
}).catch(function(err) {
  console.error('ERROR:', err.message);
  process.exit(1);
});
