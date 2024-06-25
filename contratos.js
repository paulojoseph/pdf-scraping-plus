const fs = require('fs');
const path = require('path');
const pdfParse = require('pdf-parse');
const excel = require('excel4node');

async function extractData(pdfBuffer) {
  return new Promise((resolve, reject) => {
    pdfParse(pdfBuffer).then(result => {
      const rawText = result.text;
      const jsonData = {
        contratantes: extractContratantes(rawText),
        contratada: extractContratada(rawText),
        cnpjContratado: extractCNPJContratado(rawText),
        objetoContrato: extractObjetoContrato(rawText),
        valorGlobalContrato: extractValorGlobalContrato(rawText),
        vigencia: extractVigencia(rawText),
        gestorContrato: extractGestorContrato(rawText)
      };
      resolve(jsonData);
    }).catch(reject);
  });
}

function cleanText(text) {
  return text.replace(/\s+/g, ' ').trim(); // Remove múltiplos espaços e os substitui por um único espaço
}

function extractContratantes(text) {
  // Limpa o texto para remover espaços extras
  text = cleanText(text);

  // Primeira tentativa: CLÁUSULA 1. DAS PARTES
  const regexClausula = /CLÁUSULA 1\. \s?DAS PARTES:\s*([^]*?)(?=I\.II\s*–\s*CONTRATADA:)/;
  const matchClausula = text.match(regexClausula);
  if (matchClausula) {
    const contratanteClausula = matchClausula[1].trim().split(',')[0].trim();
    return [contratanteClausula];
  }
  
  // Segunda tentativa: I.I – CONTRATANTE
  const regexContratantes = /(?:I\.I\s*–\s*CONTRATANTE:|I\.I\s*CONTRATANTE:|I\.I\s*–\s*CONTRATANTES:|CONTRATANTES?:)\s*([^]+?)(?=\s*(?:I\.II\s*–\s*CONTRATADA:|representadas\s*na\s*forma\s*de\s*seus\s*Estatutos\s*Sociais|denominada\s*simplesmente\s*CONTRATANTE))/g;
  const matchContratantes = text.match(regexContratantes);
  if (matchContratantes) {
    const contratantes = matchContratantes[0]
      .split(';')
      .map(item => item.trim().split(',')[0].trim())
      .filter(item => item !== '')
      .map(item => item.replace(/^(?:I\.I\s*–\s*CONTRATANTE:|I\.I\s*–\s*CONTRATANTES:|I\.I\s*CONTRATANTE:|CONTRATANTES:|CONTRATANTE:)\s*/, ''));
      
    return contratantes;
  }
  
  // Terceira tentativa: Caso não haja tópico I.I – CONTRATANTE
  const regexDirectClausula = /CLÁUSULA 1\. DAS PARTES:\s*([^,]+),/;
  const matchDirectClausula = text.match(regexDirectClausula);
  if (matchDirectClausula) {
    const contratanteDirectClausula = matchDirectClausula[1].trim();
    return [contratanteDirectClausula];
  }

  // Verificação final se nenhum contratante foi encontrado
  const regexFinalAttempt = /CLÁUSULA 1\. DAS PARTES:\s*([^]*?)(?=I\.II\s*–\s*CONTRATADA:)/;
  const matchFinalAttempt = text.match(regexFinalAttempt);
  if (matchFinalAttempt) {
    const contratanteFinalAttempt = matchFinalAttempt[1].split('\n').map(line => line.trim().split(',')[0]).filter(line => line.length > 0);
    return contratanteFinalAttempt;
  }

  // Última tentativa: Extrair texto após "CLÁUSULA 1. DAS PARTES:" e antes da primeira vírgula
  const regexFallback = /CLÁUSULA 1\. DAS PARTES:\s*([^,]+),/;
  const matchFallback = text.match(regexFallback);
  if (matchFallback) {
    const contratanteFallback = matchFallback[1].trim();
    return [contratanteFallback];
  }

  return [];
}

function extractContratada(text) {
  const regexContratada = /I\.II\s*–\s*CONTRATADA:\s*([^,]+),/;
  const matchContratada = text.match(regexContratada);
  return matchContratada ? matchContratada[1].trim() : '';
}

function extractCNPJContratado(text) {
  const sectionRegex = /I\.II\s*–\s*CONTRATADA:[^]*?(?=CLÁUSULA\s*2[ºª]?\.\s*DO\s*OBJETO)/g;
  const sectionMatch = text.match(sectionRegex);
  if (sectionMatch) {
    const cnpjRegex = /\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}/;
    const cnpjMatch = sectionMatch[0].match(cnpjRegex);
    return cnpjMatch ? cnpjMatch[0] : '';
  }
  return '';
}

function extractObjetoContrato(text) {
  const regex = /(Objeto|OBJETO):\s*([^]*?)\s*(Valor|Preço):?/;
  const match = text.match(regex);
  return match ? match[2].replace(/\s+/g, ' ').trim() : '';
}

function extractValorGlobalContrato(text) {
  // Regex para o formato original de valor
  const regexValor1 = /Valor(?:\s+Total\s+Estimado)?:?\s*R\$\s*([\d,.]+)/i;
  let match = text.match(regexValor1);

  if (match) {
    return match[1].replace(/\s+/g, ' ').trim();
  }

  // Regex para o novo formato de valor
  const regexValor2 = /Preço:\s*Conforme\s+tabela\s+LPU/i;
  match = text.match(regexValor2);

  if (match) {
    return 'Conforme tabela LPU';
  }

  return '';
}

function extractVigencia(text) {
  // Regex para o formato original de vigência
  const regexVigencia1 = /Da\s*[Vv]igência:\s*I\s*–\s*Termo\s*[Ii]nicial:\s*([\d\w\s]+);?\s*II\s*–\s*Termo\s*[Ff]inal:\s*([\d\w\s]+)\.?/i;
  let match = text.match(regexVigencia1);

  if (!match) {
    // Regex para o novo formato de vigência
    const regexVigencia2 = /CLÁUSULA\s*\d+\.\s*DA\s*VIGÊNCIA[^]*?termos?\s*(?:de\s*)?vigência[^]*?\s*I\.Termo\s*[Ii]nicial:\s*([\d\w\s]+)\.\s*II\.Termo\s*[Ff]inal:\s*([\d\w\s]+)\./i;
    match = text.match(regexVigencia2);
  }

  if (!match) {
    // Regex para o formato "Da vigência:"
    const regexVigencia3 = /Da\s*[Vv]igência:\s*I\.Termo\s*[Ii]nicial:\s*([\d\w\s]+)\.\s*II\.Termo\s*[Ff]inal:\s*([\d\w\s]+)\./i;
    match = text.match(regexVigencia3);
  }

  if (match) {
    const inicio = formatarData(match[1].trim());
    const fim = formatarData(match[2].trim());
    return { inicio, fim };
  }
  return {};
}

// Função auxiliar para converter data no formato desejado
function formatarData(data) {
  const meses = {
    'janeiro': '01',
    'fevereiro': '02',
    'março': '03',
    'abril': '04',
    'maio': '05',
    'junho': '06',
    'julho': '07',
    'agosto': '08',
    'setembro': '09',
    'outubro': '10',
    'novembro': '11',
    'dezembro': '12'
  };

  const partes = data.toLowerCase().split(' ');
  const dia = partes[0].padStart(2, '0');
  const mes = meses[partes[2]];
  const ano = partes[4];
  return `${dia}/${mes}/${ano}`;
}

function extractGestorContrato(text) {
  const regex = /(ESPECIALISTA|ANALISTA)\s*RESPONSÁVEL:\s*([^\n]+)/;
  const match = text.match(regex);
  return match ? match[2].trim() : '';
}

// Função para processar todos os PDFs em uma pasta e criar a tabela
async function processPdfsInFolder(folderPath) {
  const startTime = new Date(); // Marca o tempo de início
  const files = fs.readdirSync(folderPath);
  const pdfFiles = files.filter(file => path.extname(file).toLowerCase() === '.pdf');
  const allData = [];

  for (const pdfFile of pdfFiles) {
    const pdfPath = path.join(folderPath, pdfFile);
    const pdfBuffer = fs.readFileSync(pdfPath);
    const data = await extractData(pdfBuffer);
    data.nomeArquivo = pdfFile;
    allData.push(data);
  }

  createExcelFile(allData, startTime);
}

// Função para criar um arquivo Excel a partir dos dados extraídos
function createExcelFile(data, startTime) {
  const workbook = new excel.Workbook();
  const worksheet = workbook.addWorksheet('Contratos');

  const columns = [
    { header: 'Nome do Arquivo', key: 'nomeArquivo' },
    { header: 'Contratantes', key: 'contratantes' },
    { header: 'Contratada', key: 'contratada' },
    { header: 'CNPJ Contratado', key: 'cnpjContratado' },
    { header: 'Objeto do Contrato', key: 'objetoContrato' },
    { header: 'Valor Global do Contrato', key: 'valorGlobalContrato' },
    { header: 'Vigência (Início)', key: 'vigencia_inicio' },
    { header: 'Vigência (Fim)', key: 'vigencia_fim' },
    { header: 'Gestor do Contrato', key: 'gestorContrato' }
  ];

  // Adicionar cabeçalhos
  columns.forEach((col, index) => {
    worksheet.cell(1, index + 1).string(col.header);
  });

  // Adicionar dados
  data.forEach((item, rowIndex) => {
    worksheet.cell(rowIndex + 2, 1).string(item.nomeArquivo);
    worksheet.cell(rowIndex + 2, 2).string(item.contratantes.join('; '));
    worksheet.cell(rowIndex + 2, 3).string(item.contratada);
    worksheet.cell(rowIndex + 2, 4).string(item.cnpjContratado);
    worksheet.cell(rowIndex + 2, 5).string(item.objetoContrato);
    worksheet.cell(rowIndex + 2, 6).string(item.valorGlobalContrato);
    worksheet.cell(rowIndex + 2, 7).string(item.vigencia ? item.vigencia.inicio : '');
    worksheet.cell(rowIndex + 2, 8).string(item.vigencia ? item.vigencia.fim : '');
    worksheet.cell(rowIndex + 2, 9).string(item.gestorContrato);
  });

  const endTime = new Date(); // Marca o tempo de fim
  const duration = (endTime - startTime) / 1000; // Duração em segundos

  // Adicionar a duração total no final do arquivo
  const durationMinutes = Math.floor(duration / 60);
  const durationSeconds = duration % 60;

  worksheet.cell(data.length + 4, 1).string(`Duração total: ${durationMinutes} minutos e ${durationSeconds.toFixed(2)} segundos`);

  workbook.write('Contratos.xlsx');
}

// Função principal para iniciar o processamento
async function main(folderPath) {
  try {
    await processPdfsInFolder(folderPath);
    console.log('Processamento concluído. O arquivo Contratos.xlsx foi criado.');
  } catch (error) {
    console.error('Erro ao processar os arquivos PDF:', error);
  }
}

// Executa o script com o caminho da pasta como argumento
const folderPath = process.argv[2];
if (!folderPath) {
  console.error('Por favor, forneça o caminho para a pasta com os arquivos PDF.');
  process.exit(1);
}
main(folderPath);
