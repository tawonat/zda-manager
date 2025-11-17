// ============================================================================
// OLLI - SISTEMA DE AGENDAMENTO DE MANUTENÇÃO
// Versão Atualizada: Lê coluna "Bem" e "SERVICO" da planilha Excel
// ============================================================================

// Dados globais da aplicação
let agendaProcessada = [];
let isEditMode = false;
let agendaBackup = [];
let editedRows = new Set();
let periodoAtual = 'mensal';
let dataInicioFiltro = null;
let dataFimFiltro = null;

// Mapa de equipamentos e técnicos especializados
const mapaEquipamentos = {
  'COZINHADOR': { categoria: 'Cozinhador', complexidade: 'Alta', tecnicos_ideais: ['WILLIAM TEIXEIRA FERNANDES', 'ALEXANDRE CASSIANO DA CUNHA', 'MARCELO MOREIRA'] },
  'PAINEL': { categoria: 'Painel/Automação', complexidade: 'Alta', tecnicos_ideais: ['ANDERSON APARECIDO SABATINE', 'RAFAEL GALIANO LIRA', 'BRUNO YURI ANDREAZE MINAMI'] },
  'MOTOR': { categoria: 'Motor', complexidade: 'Média', tecnicos_ideais: ['WILLIAM TEIXEIRA FERNANDES', 'ALEXANDRE CASSIANO DA CUNHA', 'PAULO CEZAR BENEDITO SANTOS'] },
  'BOMBA': { categoria: 'Bomba', complexidade: 'Média', tecnicos_ideais: ['ALEXANDRE CASSIANO DA CUNHA', 'LEANDRO DA SILVA BALDENEBRO', 'PAULO CEZAR BENEDITO SANTOS'] },
  'TANQUE': { categoria: 'Tanque', complexidade: 'Média', tecnicos_ideais: ['LEANDRO DA SILVA BALDENEBRO', 'WELLINGTON RICARDO C OLIVEIRA', 'ODAIR MOREIRA'] },
  'ESTEIRA': { categoria: 'Transmissão', complexidade: 'Média', tecnicos_ideais: ['MARCELO MOREIRA', 'ADILSON APARECIDO PAIOLI', 'FABIO SANTANA RIBEIRO'] },
  'REDUTOR': { categoria: 'Transmissão', complexidade: 'Média', tecnicos_ideais: ['MARCELO MOREIRA', 'ADILSON APARECIDO PAIOLI', 'FABIO SANTANA RIBEIRO'] },
  'AGITADOR': { categoria: 'Agitador', complexidade: 'Média', tecnicos_ideais: ['PAULO CEZAR BENEDITO SANTOS', 'WELLINGTON RICARDO C OLIVEIRA', 'ODAIR MOREIRA'] },
  'MISTURA': { categoria: 'Agitador', complexidade: 'Média', tecnicos_ideais: ['PAULO CEZAR BENEDITO SANTOS', 'WELLINGTON RICARDO C OLIVEIRA', 'ODAIR MOREIRA'] },
  'AR CONDICIONADO': { categoria: 'Climatização', complexidade: 'Média', tecnicos_ideais: ['ILDERLAN FELIPE DO E SANTO', 'BRUNO YURI ANDREAZE MINAMI', 'FELIPE ALMEIDA LEONILDO'] },
  'BATEDEIRA': { categoria: 'Batedeira', complexidade: 'Média', tecnicos_ideais: ['MARCELO MOREIRA', 'LEANDRO DA SILVA BALDENEBRO', 'PAULO CEZAR BENEDITO SANTOS'] },
  'LIMPEZA': { categoria: 'Limpeza', complexidade: 'Baixa', tecnicos_ideais: ['ROBSON FERNANDO P DOS SANTOS'] },
  'CONSERVACAO': { categoria: 'Limpeza', complexidade: 'Baixa', tecnicos_ideais: ['ROBSON FERNANDO P DOS SANTOS'] }
};

const niveisExperiencia = {
  senior: ['RAFAEL GALIANO LIRA', 'MARCELO MOREIRA', 'ANDERSON APARECIDO SABATINE', 'WILLIAM TEIXEIRA FERNANDES', 'ALEXANDRE CASSIANO DA CUNHA'],
  pleno: ['ILDERLAN FELIPE DO E SANTO', 'PAULO CEZAR BENEDITO SANTOS', 'LEANDRO DA SILVA BALDENEBRO', 'WELLINGTON RICARDO C OLIVEIRA', 'ODAIR MOREIRA'],
  junior: ['BRUNO YURI ANDREAZE MINAMI', 'FELIPE ALMEIDA LEONILDO', 'JOAO VITOR GARCIA ALVES', 'MARCOS ALBERTO SELLER DE JESUS', 'ADILSON APARECIDO PAIOLI', 'FABIO SANTANA RIBEIRO'],
  suporte: ['ROBSON FERNANDO P DOS SANTOS']
};

const QUOTA_PREVENTIVA = 0.50;
const QUOTA_PROGRAMADA = 0.15;
const QUOTA_CORRETIVA = 0.35;
const MAX_HORAS_POR_TECNICO = 176;

let allAgenda = [];
const tecnicos = [
  'ANDERSON APARECIDO SABATINE', 'BRUNO YURI ANDREAZE MINAMI', 'ILDERLAN FELIPE DO E SANTO',
  'RAFAEL GALIANO LIRA', 'FELIPE ALMEIDA LEONILDO', 'JOAO VITOR GARCIA ALVES',
  'ALEXANDRE CASSIANO DA CUNHA', 'LEANDRO DA SILVA BALDENEBRO', 'PAULO CEZAR BENEDITO SANTOS',
  'MARCOS ALBERTO SELLER DE JESUS', 'WELLINGTON RICARDO C OLIVEIRA', 'WILLIAM TEIXEIRA FERNANDES',
  'MARCELO MOREIRA', 'ODAIR MOREIRA', 'ADILSON APARECIDO PAIOLI', 'FABIO SANTANA RIBEIRO',
  'ROBSON FERNANDO P DOS SANTOS'
];

const funcoesTecnicos = {
  'ANDERSON APARECIDO SABATINE': 'TÉCNICO AUTOMAÇÃO', 'BRUNO YURI ANDREAZE MINAMI': 'ELETRO JR.',
  'ILDERLAN FELIPE DO E SANTO': 'ASS. ELETROMECANICO', 'RAFAEL GALIANO LIRA': 'ELETRO SR.',
  'FELIPE ALMEIDA LEONILDO': 'ELETRO JR.', 'JOAO VITOR GARCIA ALVES': 'ELETRO JR.',
  'ALEXANDRE CASSIANO DA CUNHA': 'TECNICO ESP. MECANICO', 'LEANDRO DA SILVA BALDENEBRO': 'MECANICO PLN',
  'PAULO CEZAR BENEDITO SANTOS': 'MECANICO PLN', 'MARCOS ALBERTO SELLER DE JESUS': 'ASS. MECANICO',
  'WELLINGTON RICARDO C OLIVEIRA': 'MECANICO PLN', 'WILLIAM TEIXEIRA FERNANDES': 'TECNICO ESP. MECANICO',
  'MARCELO MOREIRA': 'MECANICO SR', 'ODAIR MOREIRA': 'MECANICO PLN',
  'ADILSON APARECIDO PAIOLI': 'MECANICO PLN', 'FABIO SANTANA RIBEIRO': 'MECANICO PLN',
  'ROBSON FERNANDO P DOS SANTOS': 'Serviços Gerais'
};

const areasTecnicos = {
  'ANDERSON APARECIDO SABATINE': 'Eletromecanicos', 'BRUNO YURI ANDREAZE MINAMI': 'Eletromecanicos',
  'ILDERLAN FELIPE DO E SANTO': 'Eletromecanicos', 'RAFAEL GALIANO LIRA': 'Eletromecanicos',
  'FELIPE ALMEIDA LEONILDO': 'Eletromecanicos', 'JOAO VITOR GARCIA ALVES': 'Eletromecanicos',
  'ALEXANDRE CASSIANO DA CUNHA': 'Mecânico Geral', 'LEANDRO DA SILVA BALDENEBRO': 'Mecânico Geral',
  'PAULO CEZAR BENEDITO SANTOS': 'Mecânico Geral', 'MARCOS ALBERTO SELLER DE JESUS': 'Mecânico Geral',
  'WELLINGTON RICARDO C OLIVEIRA': 'Mecânico Geral', 'WILLIAM TEIXEIRA FERNANDES': 'Mecânico Geral',
  'MARCELO MOREIRA': 'Mecanico Embalagem', 'ODAIR MOREIRA': 'Mecanico Embalagem',
  'ADILSON APARECIDO PAIOLI': 'Mecanico Embalagem', 'FABIO SANTANA RIBEIRO': 'Mecanico Embalagem',
  'ROBSON FERNANDO P DOS SANTOS': 'Serviços gerais'
};

let chartCriticidade = null;
let chartTipo = null;
let uploadedFile = null;
let excelData = [];

// ============================================================================
// FUNÇÃO NOVA: Extrair tipo de serviço da coluna SERVICO
// ============================================================================
function extrairTipoServico(servicoCompleto) {
  if (!servicoCompleto) return 'PREVENTIVA';
  
  const servicoStr = String(servicoCompleto).toUpperCase().trim();
  
  if (servicoStr.includes('EMERGENCIAL')) return 'EMERGENCIAL';
  if (servicoStr.includes('CORRETIVA PROGRAMADA')) return 'CORRETIVA PROGRAMADA';
  if (servicoStr.includes('CORRETIVA')) return 'CORRETIVA';
  if (servicoStr.includes('PREVENTIVA')) return 'PREVENTIVA';
  
  return 'PREVENTIVA';
}

// Funções auxiliares
function detectarCategoriaEquipamento(nomeBem, tipoServico) {
  const nomeBemUpper = nomeBem.toUpperCase();
  if (tipoServico && tipoServico.toUpperCase().includes('CONSERVA')) return mapaEquipamentos['CONSERVACAO'];
  for (const [palavra, config] of Object.entries(mapaEquipamentos)) {
    if (nomeBemUpper.includes(palavra)) return config;
  }
  return { categoria: 'Equipamento Geral', complexidade: 'Média', tecnicos_ideais: ['LEANDRO DA SILVA BALDENEBRO', 'PAULO CEZAR BENEDITO SANTOS', 'WELLINGTON RICARDO C OLIVEIRA'] };
}

function calcularScoreTecnico(tecnico, categoria, complexidade, tipo, estatisticas) {
  let score = 100;
  const cargaAtual = estatisticas[tecnico].horasTotais;
  const percentualCarga = (cargaAtual / MAX_HORAS_POR_TECNICO) * 100;
  score -= percentualCarga * 0.8;
  if (categoria.tecnicos_ideais.includes(tecnico)) score += 30;
  
  const totalOrdens = estatisticas[tecnico].preventiva + estatisticas[tecnico].programada + estatisticas[tecnico].corretiva;
  if (totalOrdens > 0) {
    const percentualAtual = { 'PREVENTIVA': estatisticas[tecnico].preventiva / totalOrdens, 'CORRETIVA PROGRAMADA': estatisticas[tecnico].programada / totalOrdens, 'CORRETIVA': estatisticas[tecnico].corretiva / totalOrdens, 'EMERGENCIAL': estatisticas[tecnico].corretiva / totalOrdens };
    const quotaIdeal = { 'PREVENTIVA': QUOTA_PREVENTIVA, 'CORRETIVA PROGRAMADA': QUOTA_PROGRAMADA, 'CORRETIVA': QUOTA_CORRETIVA, 'EMERGENCIAL': QUOTA_CORRETIVA };
    const diferencaQuota = quotaIdeal[tipo] - percentualAtual[tipo];
    if (diferencaQuota > 0) score += diferencaQuota * 50;
  }
  
  const nivelTecnico = Object.keys(niveisExperiencia).find(nivel => niveisExperiencia[nivel].includes(tecnico));
  if (complexidade === 'Alta') {
    if (nivelTecnico === 'senior') score += 20;
    else if (nivelTecnico === 'pleno') score += 10;
    else if (nivelTecnico === 'junior') score -= 10;
  } else if (complexidade === 'Média') {
    if (nivelTecnico === 'pleno') score += 15;
    else if (nivelTecnico === 'senior') score += 10;
    else if (nivelTecnico === 'junior') score += 5;
  } else if (complexidade === 'Baixa') {
    if (nivelTecnico === 'junior') score += 15;
    else if (nivelTecnico === 'pleno') score += 10;
    else if (nivelTecnico === 'senior') score += 5;
  }
  return score;
}

function alocarTecnicoIdeal(nomeBem, tipoServico, criticidade, estatisticas) {
  const categoria = detectarCategoriaEquipamento(nomeBem, tipoServico);
  let complexidadeAjustada = categoria.complexidade;
  if (tipoServico) {
    const tipoUpper = tipoServico.toUpperCase();
    if ((tipoUpper.includes('CORRETIVA') && !tipoUpper.includes('PROGRAMADA')) || tipoUpper.includes('EMERGENCIAL')) {
      complexidadeAjustada = 'Alta';
    } else if (tipoUpper.includes('PREVENTIVA')) {
      complexidadeAjustada = complexidadeAjustada === 'Alta' ? 'Média' : 'Média';
    }
  }
  if (criticidade === 'A' && complexidadeAjustada !== 'Baixa') complexidadeAjustada = 'Alta';
  
  const scoresComTecnicos = tecnicos
    .filter(tec => estatisticas[tec].horasTotais < MAX_HORAS_POR_TECNICO)
    .map(tecnico => ({ tecnico: tecnico, score: calcularScoreTecnico(tecnico, categoria, complexidadeAjustada, tipoServico, estatisticas) }))
    .sort((a, b) => b.score - a.score);
  
  if (scoresComTecnicos.length === 0) {
    const tecnicoMenosHoras = tecnicos.reduce((min, tec) => estatisticas[tec].horasTotais < estatisticas[min].horasTotais ? tec : min);
    return { tecnico: tecnicoMenosHoras, motivo: 'Alocação emergencial - todos técnicos próximos ao limite', categoria: categoria.categoria };
  }
  
  const melhorTecnico = scoresComTecnicos[0].tecnico;
  let motivo = '';
  if (categoria.tecnicos_ideais.includes(melhorTecnico)) {
    motivo = `Especialista em ${categoria.categoria}`;
  } else {
    const nivelTecnico = Object.keys(niveisExperiencia).find(nivel => niveisExperiencia[nivel].includes(melhorTecnico));
    const nivelNome = nivelTecnico === 'senior' ? 'Sênior' : nivelTecnico === 'pleno' ? 'Pleno' : nivelTecnico === 'junior' ? 'Júnior' : 'Suporte';
    motivo = `Técnico ${nivelNome} - balanceamento de carga`;
  }
  return { tecnico: melhorTecnico, motivo: motivo, categoria: categoria.categoria };
}

function formatarHoras(horasDecimais) {
  const horas = Math.floor(horasDecimais);
  const minutos = Math.round((horasDecimais - horas) * 60);
  if (horas === 0) return `${minutos} minutos`;
  if (minutos === 0) return horas === 1 ? '1 hora' : `${horas} horas`;
  const horaTexto = horas === 1 ? '1 hora' : `${horas} horas`;
  return `${horaTexto} e ${minutos} minutos`;
}

// ============================================================================
// PROCESSAMENTO DE DADOS DO EXCEL - CORRIGIDO
// ============================================================================

// ============================================================================
// NOVAS FUNÇÕES PARA OS CAMPOS ADICIONADOS
// ============================================================================

function calcularStatus(ordem) {
  // Lógica para determinar o status da ordem
  // Se já tem data agendada e está dentro do período, é "Agendada"
  // Se tem atraso menor que 7 dias e criticidade baixa, é "Pendente"
  // Caso contrário, é "Liberada"

  if (ordem.Date) {
    const dataOrdem = new Date(ordem.Date);
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    if (dataOrdem >= hoje) {
      return 'Agendada';
    }
  }

  if (ordem.Delay <= 7 && ordem.Criticality === 'C') {
    return 'Pendente';
  }

  return 'Liberada';
}

function calcularPrevisaoInicio(ordem) {
  // Calcula previsão de início baseado em criticidade, tipo e atraso
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  let diasAdicionar = 0;

  // Criticidade influencia na prioridade
  if (ordem.Criticality === 'A') {
    diasAdicionar = 1; // Alta criticidade: iniciar amanhã
  } else if (ordem.Criticality === 'B') {
    diasAdicionar = 3; // Média criticidade: iniciar em 3 dias
  } else {
    diasAdicionar = 7; // Baixa criticidade: iniciar em 7 dias
  }

  // Tipo de serviço também influencia
  if (ordem.Type === 'EMERGENCIAL') {
    diasAdicionar = 0; // Emergencial: iniciar hoje
  } else if (ordem.Type === 'CORRETIVA') {
    diasAdicionar = Math.min(diasAdicionar, 2); // Corretiva: no máximo 2 dias
  } else if (ordem.Type === 'CORRETIVA PROGRAMADA') {
    diasAdicionar = Math.max(diasAdicionar, 5); // Programada: pelo menos 5 dias
  }

  // Ajuste por atraso - se já está atrasada, priorizar
  if (ordem.Delay > 30) {
    diasAdicionar = 0; // Muito atrasada: prioridade máxima
  } else if (ordem.Delay > 15) {
    diasAdicionar = Math.max(0, diasAdicionar - 2);
  }

  const previsao = new Date(hoje);
  previsao.setDate(previsao.getDate() + diasAdicionar);

  // Pular finais de semana
  if (previsao.getDay() === 0) { // Domingo
    previsao.setDate(previsao.getDate() + 1);
  } else if (previsao.getDay() === 6) { // Sábado
    previsao.setDate(previsao.getDate() + 2);
  }

  return previsao.toISOString().split('T')[0];
}

function calcularDescricao(ordem) {
  // Retorna "LIBERADA" ou "AGUARDANDO PROGRAMAÇÃO"
  const status = calcularStatus(ordem);

  if (status === 'Agendada' || status === 'Liberada') {
    return 'LIBERADA';
  }

  return 'AGUARDANDO PROGRAMAÇÃO';
}

function processarDadosExcel() {
  const ordens = [];
  const estatisticas = {};
  tecnicos.forEach(t => { estatisticas[t] = { horasTotais: 0, preventiva: 0, programada: 0, corretiva: 0 }; });
  
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  
  excelData.forEach(ordem => {
    const ordemObj = {};
    for (const [key, value] of Object.entries(ordem)) ordemObj[key.trim()] = value;
    
    // ✅ LER COLUNA ORDEM
    const ordemNum = ordemObj['Ordem Serv.'] || ordemObj['Ordem Serv'] || ordemObj.Ordem || ordemObj['ordem_serv'] || ordemObj['Ordem_Serv'];
    
    // ✅ LER COLUNA BEM (TAG DO EQUIPAMENTO)
    const bemTag = ordemObj['Bem'] || ordemObj.bem || ordemObj.BEM || ordemObj['Tag'] || '';
    
    // ✅ LER COLUNA NOME DO BEM
    const nomeBem = ordemObj['Nome do Bem'] || ordemObj['nome_bem'] || ordemObj.Nome_Bem || ordemObj['Nome Bem'] || '';
    
    // ✅ LER E EXTRAIR TIPO DE SERVIÇO DA COLUNA SERVICO
    const servicoCompleto = ordemObj['SERVICO'] || ordemObj['Servico'] || ordemObj['servico'] || ordemObj['area de manutenção'] || ordemObj['Nome Servico'] || ordemObj['nome_servico'] || ordemObj.Nome_Servico || '';
    const tipoServico = extrairTipoServico(servicoCompleto);
    
    const dataAbertura = ordemObj['Dt. Inicio'] || ordemObj['Data_Abertura'] || ordemObj['Data Abertura'] || ordemObj.Data_Abertura || ordemObj['Dt Inicio'];

    // ✅ NOVO: Extrair Real Início se existir na planilha
    const realInicio = ordemObj['Real Inicio'] || ordemObj['Real Início'] || 
                       ordemObj['real_inicio'] || ordemObj['REAL INICIO'] || 
                       ordemObj['Data Real Inicio'] || ordemObj['Data_Real_Inicio'] || '';

    
    if (!ordemNum) return;
    
    // Determinar criticidade
    let criticidade = 'C';
    const nomeBemUpper = nomeBem.toUpperCase();
    if (nomeBemUpper.includes('COZINHADOR') || nomeBemUpper.includes('PAINEL') || nomeBemUpper.includes('LINHA')) criticidade = 'A';
    else if (nomeBemUpper.includes('BOMBA') || nomeBemUpper.includes('TANQUE') || nomeBemUpper.includes('MOTOR')) criticidade = 'B';
    
    // Calcular atraso
    let atraso = 0;
    let dataAbert = null;
    if (dataAbertura) {
      if (typeof dataAbertura === 'number') dataAbert = new Date((dataAbertura - 25569) * 86400 * 1000);
      else if (typeof dataAbertura === 'string') dataAbert = new Date(dataAbertura);
      else if (dataAbertura instanceof Date) dataAbert = dataAbertura;
      
      if (dataAbert && !isNaN(dataAbert.getTime())) {
        dataAbert.setHours(0, 0, 0, 0);
        atraso = Math.floor((hoje - dataAbert) / (1000 * 60 * 60 * 24));
        atraso = Math.max(0, atraso);
      }
    }
    if (!dataAbert || isNaN(dataAbert.getTime())) {
      const diasAtras = Math.floor(Math.random() * 60) + 1;
      dataAbert = new Date(hoje);
      dataAbert.setDate(dataAbert.getDate() - diasAtras);
      atraso = diasAtras;
    }
    
    // Calcular prioridade
    let prioridade = 0;
    if (criticidade === 'A') prioridade += 100;
    else if (criticidade === 'B') prioridade += 50;
    else prioridade += 20;
    prioridade += atraso * 5;
    if (tipoServico === 'EMERGENCIAL') prioridade += 50;
    else if (tipoServico === 'CORRETIVA PROGRAMADA') prioridade += 30;
    else if (tipoServico === 'PREVENTIVA') prioridade += 10;
    else if (tipoServico === 'CORRETIVA') prioridade += 40;
    
    ordens.push({ ordem: ordemNum, bemTag: bemTag, nomeBem: nomeBem, tipo: tipoServico, criticidade: criticidade, atraso: atraso, prioridade: prioridade, dataAbertura: dataAbert,
      realInicio: realInicio });
  });
  
  ordens.sort((a, b) => b.prioridade - a.prioridade);
  
  // Alocar técnicos
  allAgenda = [];
  const dataInicio = new Date(2025, 10, 1);
  let dataAtual = new Date(dataInicio);
  
  ordens.forEach((ordem, index) => {
    const alocacao = alocarTecnicoIdeal(ordem.nomeBem, ordem.tipo, ordem.criticidade, estatisticas);
    const tecnico = alocacao.tecnico;
    ordem.motivoAlocacao = alocacao.motivo;
    ordem.categoriaEquipamento = alocacao.categoria;
    
    let duracao;
    if (ordem.tipo === 'PREVENTIVA') duracao = 1 + Math.random() * 1.5;
    else if (ordem.tipo === 'CORRETIVA' || ordem.tipo === 'EMERGENCIAL') duracao = 2 + Math.random() * 1.5;
    else duracao = 1.5 + Math.random() * 1.5;
    duracao = Math.round(duracao * 10) / 10;
    
    estatisticas[tecnico].horasTotais += duracao;
    if (ordem.tipo === 'PREVENTIVA') estatisticas[tecnico].preventiva++;
    else if (ordem.tipo === 'CORRETIVA' || ordem.tipo === 'EMERGENCIAL') estatisticas[tecnico].corretiva++;
    else if (ordem.tipo === 'CORRETIVA PROGRAMADA') estatisticas[tecnico].programada++;
    
    const horaInicio = 9 + Math.floor(Math.random() * 6);
    const minutoInicio = [0, 30][Math.floor(Math.random() * 2)];
    const startTime = `${String(horaInicio).padStart(2, '0')}:${String(minutoInicio).padStart(2, '0')}:00`;
    const totalMinutos = horaInicio * 60 + minutoInicio + (duracao * 60);
    const horaFim = Math.floor(totalMinutos / 60);
    const minutoFim = Math.round(totalMinutos % 60);
    const endTime = `${String(horaFim).padStart(2, '0')}:${String(minutoFim).padStart(2, '0')}:00`;
    
    allAgenda.push({
      Technician: tecnico,
      Ordem: ordem.ordem,
      Bem: ordem.bemTag,
      Type: ordem.tipo,
      Date: ordem.realInicio || dataAtual.toISOString().split('T')[0],
      StartTime: startTime,
      EndTime: endTime,
      Duration: duracao,
      Criticality: ordem.criticidade,
      Delay: ordem.atraso,
      Priority: ordem.prioridade,
      MotivoAlocacao: ordem.motivoAlocacao,
      CategoriaEquipamento: ordem.categoriaEquipamento,
      NomeBem: ordem.nomeBem,
      DataAbertura: ordem.dataAbertura.toISOString().split('T')[0],
        Status: '', // Será calculado depois
        PrevisaoInicio: '', // Será calculado depois
        RealInicio: ordem.realInicio || '', // Vem da planilha ou fica vazio
        Descricao: '' // Será calculado depois
    });

  // Calcular os campos adicionados (Status, PrevisaoInicio, Descricao)
  allAgenda = allAgenda.map(ordem => ({
    ...ordem,
    Status: calcularStatus(ordem),
    PrevisaoInicio: calcularPrevisaoInicio(ordem),
    Descricao: calcularDescricao(ordem)
  }));

    
    if (index % 3 === 0 && index > 0) {
      dataAtual.setDate(dataAtual.getDate() + 1);
      if (dataAtual.getDay() === 0) dataAtual.setDate(dataAtual.getDate() + 1);
      if (dataAtual.getDay() === 6) dataAtual.setDate(dataAtual.getDate() + 2);
    }
  });
  
  agendaProcessada = [...allAgenda];
  atualizarInterface();
}

// ============================================================================
// INICIALIZAÇÃO E UPLOAD
// ============================================================================
document.addEventListener('DOMContentLoaded', () => {
  initializeApp();
  setupUploadInterface();
});

function initializeApp() {
  setupTabs();
  setupButtons();
  setupFilters();
}

function setupUploadInterface() {
  const dropzone = document.getElementById('dropzoneOrdens');
  const fileInput = document.getElementById('fileOrdens');
  
  dropzone.addEventListener('click', () => fileInput.click());
  
  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('drag-over');
  });
  
  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('drag-over'));
  
  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) handleFileUpload(file);
  });
  
  fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) handleFileUpload(file);
  });
  
  document.getElementById('processBtn').addEventListener('click', processarArquivos);
  document.getElementById('loadExampleBtn').addEventListener('click', carregarDadosExemplo);
}

function handleFileUpload(file) {
  const validExtensions = ['.xlsx', '.xls'];
  const fileName = file.name.toLowerCase();
  const isValid = validExtensions.some(ext => fileName.endsWith(ext));
  
  if (!isValid) {
    showAlert('Por favor, selecione um arquivo Excel válido (.xlsx ou .xls)', 'error');
    return;
  }
  
  uploadedFile = file;
  document.getElementById('dropzoneOrdens').style.display = 'none';
  document.getElementById('fileStatus').style.display = 'flex';
  document.getElementById('fileStatusText').textContent = file.name;
  document.getElementById('processBtn').disabled = false;
}

async function processarArquivos() {
  const processBtn = document.getElementById('processBtn');
  const processBtnText = document.getElementById('processBtnText');
  const processSpinner = document.getElementById('processSpinner');
  
  processBtn.disabled = true;
  processBtnText.textContent = 'Processando...';
  processSpinner.style.display = 'inline-block';
  
  try {
    await lerArquivoExcel();
    await new Promise(resolve => setTimeout(resolve, 2000));
    processarDadosExcel();
    
    document.getElementById('uploadSection').style.display = 'none';
    document.getElementById('resultsSection').style.display = 'block';
    document.getElementById('mainSubtitle').textContent = 'Agenda Processada com Sucesso';
    aplicarFiltroPeriodo('mensal');
    showAlert(`Processado! ${allAgenda.length} ordens agendadas com sucesso`, 'success');
  } catch (error) {
    showAlert('Erro ao processar arquivo: ' + error.message, 'error');
    processBtn.disabled = false;
    processBtnText.textContent = 'PROCESSAR ORDENS';
    processSpinner.style.display = 'none';
  }
}

async function lerArquivoExcel() {
  excelData = await lerExcel(uploadedFile);
}

function lerExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet);
        resolve(jsonData);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(new Error('Erro ao ler arquivo'));
    reader.readAsArrayBuffer(file);
  });
}

function carregarDadosExemplo() {
  uploadedFile = { name: 'exemplo-ordens-totvs.xlsx' };
  allAgenda = gerarDadosExemplo();
  agendaProcessada = [...allAgenda];
  
  document.getElementById('uploadSection').style.display = 'none';
  document.getElementById('resultsSection').style.display = 'block';
  document.getElementById('mainSubtitle').textContent = 'Agenda Processada com Sucesso';
  aplicarFiltroPeriodo('mensal');
  showAlert(`Processado! ${agendaProcessada.length} ordens agendadas`, 'success');
}

function gerarDadosExemplo() {
  const dados = [];
  const tipos = ['PREVENTIVA', 'CORRETIVA', 'CORRETIVA PROGRAMADA', 'EMERGENCIAL'];
  const criticidades = ['A', 'A', 'A', 'B', 'B', 'C'];
  const ordens = ['011221', '011257', '011992', '010894', '010895', '010896', '010898', '010899', '010900', '011075', '010905', '011765', '012734', '012735', '012736', '012737', '012740', '012741', '012742', '012743', '012744', '012745', '012746', '012747', '012748', '012749', '012750'];
  const nomesBens = ['COZINHADOR LINHA 1', 'COZINHADOR LINHA 2', 'PAINEL ELETRICO SALA 01', 'PAINEL AUTOMACAO LINHA 3', 'MOTOR BOMBA TRANSFERENCIA', 'BOMBA CENTRIFUGA 01', 'TANQUE ARMAZENAMENTO 500L', 'ESTEIRA TRANSPORTADORA ET-01', 'AGITADOR TANQUE TQ-03', 'AR CONDICIONADO SALA CONTROLE'];
  const tagsBens = ['AB0003', 'AG0060', 'AG0061', 'BB0009', 'BB0010', 'BB0011', 'BB0016', 'BB0017', 'BB0018', 'BB0019'];
  
  const estatisticas = {};
  tecnicos.forEach(t => { estatisticas[t] = { horasTotais: 0, preventiva: 0, corretiva: 0, programada: 0 }; });
  
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  let dataAtual = new Date(2025, 10, 3);
  let ordemIndex = 0;
  
  for (let dia = 0; dia < 20; dia++) {
    const ordensNoDia = Math.floor(Math.random() * 4) + 2;
    for (let i = 0; i < ordensNoDia && ordemIndex < ordens.length; i++) {
      const tipo = tipos[Math.floor(Math.random() * tipos.length)];
      const criticidade = criticidades[Math.floor(Math.random() * criticidades.length)];
      const nomeBem = nomesBens[Math.floor(Math.random() * nomesBens.length)];
      const tagBem = tagsBens[Math.floor(Math.random() * tagsBens.length)];
      const duracao = [1, 1.5, 2, 2.5, 3][Math.floor(Math.random() * 5)];
      
      const diasAtras = Math.floor(Math.random() * 60) + 1;
      const dataAbertura = new Date(hoje);
      dataAbertura.setDate(dataAbertura.getDate() - diasAtras);
      const atraso = Math.floor((hoje - dataAbertura) / (1000 * 60 * 60 * 24));
      
      const alocacao = alocarTecnicoIdeal(nomeBem, tipo, criticidade, estatisticas);
      const tecnico = alocacao.tecnico;
      
      estatisticas[tecnico].horasTotais += duracao;
      if (tipo === 'PREVENTIVA') estatisticas[tecnico].preventiva++;
      else if (tipo === 'CORRETIVA' || tipo === 'EMERGENCIAL') estatisticas[tecnico].corretiva++;
      else if (tipo === 'CORRETIVA PROGRAMADA') estatisticas[tecnico].programada++;
      
      let prioridade = 0;
      if (criticidade === 'A') prioridade += 100;
      else if (criticidade === 'B') prioridade += 50;
      else prioridade += 20;
      prioridade += atraso * 5;
      if (tipo === 'EMERGENCIAL') prioridade += 50;
      else if (tipo === 'CORRETIVA PROGRAMADA') prioridade += 30;
      else if (tipo === 'PREVENTIVA') prioridade += 10;
      
      const horaInicio = 9 + Math.floor(Math.random() * 6);
      const minutoInicio = [0, 30][Math.floor(Math.random() * 2)];
      const startTime = `${String(horaInicio).padStart(2, '0')}:${String(minutoInicio).padStart(2, '0')}:00`;
      const totalMinutos = horaInicio * 60 + minutoInicio + (duracao * 60);
      const horaFim = Math.floor(totalMinutos / 60);
      const minutoFim = totalMinutos % 60;
      const endTime = `${String(horaFim).padStart(2, '0')}:${String(minutoFim).padStart(2, '0')}:00`;
      
      dados.push({
        Technician: tecnico,
        Ordem: ordens[ordemIndex],
        Bem: tagBem,
        Type: tipo,
        Date: dataAtual.toISOString().split('T')[0],
        StartTime: startTime,
        EndTime: endTime,
        Duration: duracao,
        Criticality: criticidade,
        Delay: atraso,
        Priority: prioridade,
        MotivoAlocacao: alocacao.motivo,
        CategoriaEquipamento: alocacao.categoria,
        NomeBem: nomeBem,
        DataAbertura: dataAbertura.toISOString().split('T')[0],
        Status: '', // Será calculado depois
        PrevisaoInicio: '', // Será calculado depois
        RealInicio: '', // Dados de exemplo não têm real início
        Descricao: '' // Será calculado depois
      });
      ordemIndex++;
    }
    dataAtual.setDate(dataAtual.getDate() + 1);
  }

  // Calcular os campos adicionados para dados de exemplo
  dados = dados.map(ordem => ({
    ...ordem,
    Status: calcularStatus(ordem),
    PrevisaoInicio: calcularPrevisaoInicio(ordem),
    Descricao: calcularDescricao(ordem)
  }));
  
  return dados;
}

// ============================================================================
// SISTEMA DE ABAS E FILTROS
// ============================================================================
function setupTabs() {
  const tabButtons = document.querySelectorAll('.tab-button');
  const tabPanes = document.querySelectorAll('.tab-pane');
  
  tabButtons.forEach(button => {
    button.addEventListener('click', () => {
      const tabName = button.getAttribute('data-tab');
      tabButtons.forEach(btn => btn.classList.remove('active'));
      tabPanes.forEach(pane => pane.classList.remove('active'));
      button.classList.add('active');
      document.getElementById(tabName).classList.add('active');
      if (tabName === 'estatisticas') setTimeout(() => renderizarGraficos(), 100);
    });
  });
}

function setupFilters() {
  const filterTecnico = document.getElementById('filterTecnico');
  const filterCriticidade = document.getElementById('filterCriticidade');
  const filterTipo = document.getElementById('filterTipo');
  
  tecnicos.forEach(tec => {
    const option = document.createElement('option');
    option.value = tec;
    option.textContent = tec;
    filterTecnico.appendChild(option);
  });
  
  filterTecnico.addEventListener('change', aplicarFiltros);
  filterCriticidade.addEventListener('change', aplicarFiltros);
  filterTipo.addEventListener('change', aplicarFiltros);
}

function aplicarFiltros() {
  const tecnico = document.getElementById('filterTecnico').value;
  const criticidade = document.getElementById('filterCriticidade').value;
  const tipo = document.getElementById('filterTipo').value;
  
  agendaProcessada = allAgenda.filter(ordem => {
    if (tecnico && ordem.Technician !== tecnico) return false;
    if (criticidade && ordem.Criticality !== criticidade) return false;
    if (tipo && ordem.Type !== tipo) return false;
    return true;
  });
  
  atualizarTabelaAgenda();
}

function setupButtons() {
  document.getElementById('reprocessBtn').addEventListener('click', voltarParaUpload);
  document.getElementById('downloadExcelBtn').addEventListener('click', downloadExcel);
  document.getElementById('editAgendaBtn').addEventListener('click', entrarModoEdicao);
  document.getElementById('saveChangesBtn').addEventListener('click', salvarAlteracoes);
  document.getElementById('cancelEditBtn').addEventListener('click', cancelarEdicao);
  document.getElementById('undoChangesBtn').addEventListener('click', desfazerAlteracoes);
  document.getElementById('filtroDiario').addEventListener('click', () => aplicarFiltroPeriodo('diario'));
  document.getElementById('filtroSemanal').addEventListener('click', () => aplicarFiltroPeriodo('semanal'));
  document.getElementById('filtroMensal').addEventListener('click', () => aplicarFiltroPeriodo('mensal'));
}

function aplicarFiltroPeriodo(periodo) {
  periodoAtual = periodo;
  document.querySelectorAll('.periodo-btn').forEach(btn => btn.classList.remove('active'));
  document.getElementById(`filtro${periodo.charAt(0).toUpperCase() + periodo.slice(1)}`).classList.add('active');
  
  const hoje = new Date();
  dataFimFiltro = new Date(hoje);
  
  if (periodo === 'diario') dataInicioFiltro = new Date(hoje);
  else if (periodo === 'semanal') {
    dataInicioFiltro = new Date(hoje);
    dataInicioFiltro.setDate(dataInicioFiltro.getDate() - 7);
  } else if (periodo === 'mensal') {
    dataInicioFiltro = new Date(hoje);
    dataInicioFiltro.setDate(dataInicioFiltro.getDate() - 30);
  }
  
  filtrarPorPeriodo();
  atualizarInterface();
  showAlert(`Filtro ${periodo} aplicado com sucesso!`, 'success');
}

function filtrarPorPeriodo() {
  if (!dataInicioFiltro || !dataFimFiltro) {
    agendaProcessada = [...allAgenda];
    return;
  }
  agendaProcessada = allAgenda.filter(ordem => {
    const dataOrdem = new Date(ordem.Date + 'T00:00:00');
    return dataOrdem >= dataInicioFiltro && dataOrdem <= dataFimFiltro;
  });
}

function voltarParaUpload() {
  uploadedFile = null;
  excelData = [];
  document.getElementById('dropzoneOrdens').style.display = 'flex';
  document.getElementById('fileStatus').style.display = 'none';
  document.getElementById('fileOrdens').value = '';
  document.getElementById('processBtn').disabled = true;
  document.getElementById('uploadSection').style.display = 'block';
  document.getElementById('resultsSection').style.display = 'none';
  document.getElementById('mainSubtitle').textContent = 'Processe as Ordens de Serviço do TOTVS';
  showAlert('Pronto para carregar novo arquivo', 'success');
}

// ============================================================================
// ATUALIZAÇÃO DE INTERFACE
// ============================================================================
function atualizarInterface() {
  atualizarDashboard();
  atualizarTabelaAgenda();
  atualizarTecnicos();
  atualizarTimeline();
  renderizarGraficos();
}

function atualizarDashboard() {
  const total = agendaProcessada.length;
  const horasTotais = agendaProcessada.reduce((sum, o) => sum + o.Duration, 0);
  const tecnicosAlocados = new Set(agendaProcessada.map(o => o.Technician)).size;
  
  document.getElementById('totalOS').textContent = total;
  document.getElementById('agendadas').textContent = total;
  document.getElementById('horasTotais').textContent = formatarHoras(horasTotais);
  document.getElementById('taxaUtilizacao').textContent = `${tecnicosAlocados} de ${tecnicos.length}`;
}

function atualizarTabelaAgenda() {
  const tbody = document.getElementById('agendaTableBody');

  if (agendaProcessada.length === 0) {
    tbody.innerHTML = '<tr><td colspan="14" class="text-center text-secondary">Nenhuma ordem encontrada</td></tr>';
    return;
  }

  const ordenada = [...agendaProcessada].sort((a, b) => {
    const prioA = a.Priority || calcularPrioridade(a);
    const prioB = b.Priority || calcularPrioridade(b);
    return prioB - prioA;
  });

  tbody.innerHTML = ordenada.map(ordem => {
    const prioridade = ordem.Priority || calcularPrioridade(ordem);
    const funcao = funcoesTecnicos[ordem.Technician] || 'Técnico';
    const motivo = ordem.MotivoAlocacao || 'Alocação automática';
    const categoria = ordem.CategoriaEquipamento || '-';
    const nomeBem = ordem.NomeBem || '-';
    const status = ordem.Status || 'Agendada';
    const isEdited = editedRows.has(ordem.Ordem);

    const atraso = ordem.Delay || 0;
    let atrasoColor = 'var(--color-success)';
    if (atraso > 30) atrasoColor = 'var(--color-error)';
    else if (atraso > 15) atrasoColor = 'var(--color-warning)';
    else if (atraso > 7) atrasoColor = 'var(--color-info)';
    const atrasoHtml = `<span style="color: ${atrasoColor}; font-weight: var(--font-weight-semibold);">${atraso} dias</span>`;

    if (isEditMode) {
      const technicianOptions = tecnicos.map(t => `<option value="${t}" ${ordem.Technician === t ? 'selected' : ''}>${t}</option>`).join('');
      const statusOptions = ['Agendada', 'Pendente', 'Cancelada'].map(s => `<option value="${s}" ${status === s ? 'selected' : ''}>${s}</option>`).join('');
      const dateValue = ordem.Date || '';

      return `
      <tr style="${isEdited ? 'border-left: 3px solid var(--color-warning);' : ''}">
        <td><strong>${ordem.Ordem}</strong></td>
        <td><strong>${ordem.Bem || '-'}</strong></td>
        <td>${nomeBem}</td>
        <td><span class="badge badge--info">${ordem.Type || '-'}</span></td>
        <td>${getBadgeCriticidade(ordem.Criticality)}</td>
        <td>${atrasoHtml}</td>
        <td><strong>${prioridade}</strong></td>
        <td class="editable-cell">
          <select onchange="editarCelula('${ordem.Ordem}', 'Technician', this.value)" class="form-select-small">
            ${technicianOptions}
          </select>
        </td>
        <td class="editable-cell">
          <input type="date" value="${dateValue}" onchange="editarCelula('${ordem.Ordem}', 'Date', this.value)" class="form-input-small">
        </td>
        <td class="editable-cell">
          <select onchange="editarCelula('${ordem.Ordem}', 'Status', this.value)" class="form-select-small">
            ${statusOptions}
          </select>
        </td>
        <td>${formatarData(ordem.PrevisaoInicio || '-')}</td>
        <td>${formatarData(ordem.DataAbertura || '-')}</td>
        <td>${ordem.RealInicio ? formatarData(ordem.RealInicio) : '-'}</td>
        <td><span style="font-size: var(--font-size-sm);">${ordem.Descricao || '-'}</span></td>
      </tr>`;
    } else {
      const statusBadge = status === 'Agendada' ? 'badge--success' : status === 'Pendente' ? 'badge--warning' : 'badge--error';
      return `
      <tr>
        <td><strong>${ordem.Ordem}</strong></td>
        <td><strong>${ordem.Bem || '-'}</strong></td>
        <td>${nomeBem}</td>
        <td><span class="badge badge--info">${ordem.Type || '-'}</span></td>
        <td>${getBadgeCriticidade(ordem.Criticality)}</td>
        <td>${atrasoHtml}</td>
        <td><strong>${prioridade}</strong></td>
        <td>
          <strong>${ordem.Technician}</strong><br>
          <span style="font-size: var(--font-size-sm); color: var(--color-text-secondary);">${funcao}</span><br>
          <span class="allocation-info" style="font-size: var(--font-size-sm); color: var(--color-primary); font-weight: var(--font-weight-medium);" data-tooltip="${motivo}">${categoria}</span>
        </td>
        <td>${formatarData(ordem.Date)}</td>
        <td><span class="badge ${statusBadge}">${status}</span></td>
        <td>${formatarData(ordem.PrevisaoInicio || '-')}</td>
        <td>${formatarData(ordem.DataAbertura || '-')}</td>
        <td>${ordem.RealInicio ? formatarData(ordem.RealInicio) : '<span style="color: var(--color-text-secondary);">-</span>'}</td>
        <td><span style="font-size: var(--font-size-sm); font-weight: var(--font-weight-medium);">${ordem.Descricao || '-'}</span></td>
      </tr>`;
    }
  }).join('');
}


function calcularPrioridade(ordem) {
  let prioridade = 0;
  if (ordem.Criticality === 'A') prioridade += 100;
  else if (ordem.Criticality === 'B') prioridade += 50;
  else prioridade += 20;
  prioridade += ordem.Delay * 5;
  if (ordem.Type === 'EMERGENCIAL') prioridade += 50;
  else if (ordem.Type === 'CORRETIVA PROGRAMADA') prioridade += 30;
  else if (ordem.Type === 'PREVENTIVA') prioridade += 10;
  return prioridade;
}

function getBadgeCriticidade(criticidade) {
  const classes = { 'A': 'badge--critical', 'B': 'badge--medium', 'C': 'badge--low' };
  return `<span class="badge ${classes[criticidade]}">${criticidade}</span>`;
}

function formatarData(dateStr) {
  const date = new Date(dateStr + 'T00:00:00');
  return date.toLocaleDateString('pt-BR');
}

// ============================================================================
// ATUALIZAR TÉCNICOS & CARGA HORÁRIA
// ============================================================================
function atualizarTecnicos() {
  const grid = document.getElementById('techniciansGrid');
  const dadosPorTecnico = {};
  tecnicos.forEach(tec => {
    dadosPorTecnico[tec] = { ordens: [], horasTotais: 0, preventiva: 0, corretiva: 0, programada: 0 };
  });
  
  agendaProcessada.forEach(ordem => {
    const tec = dadosPorTecnico[ordem.Technician];
    if (tec) {
      tec.ordens.push(ordem);
      tec.horasTotais += ordem.Duration;
      if (ordem.Type === 'PREVENTIVA') tec.preventiva++;
      else if (ordem.Type === 'CORRETIVA' || ordem.Type === 'EMERGENCIAL') tec.corretiva++;
      else if (ordem.Type === 'CORRETIVA PROGRAMADA') tec.programada++;
    }
  });
  
  grid.innerHTML = Object.entries(dadosPorTecnico)
    .filter(([tecnico, dados]) => dados.ordens.length > 0)
    .map(([tecnico, dados]) => {
      const total = dados.ordens.length;
      const prevPct = total > 0 ? (dados.preventiva / total) * 100 : 0;
      const corrPct = total > 0 ? (dados.corretiva / total) * 100 : 0;
      const progPct = total > 0 ? (dados.programada / total) * 100 : 0;
      const funcao = funcoesTecnicos[tecnico] || 'Técnico';
      const area = areasTecnicos[tecnico] || 'Manutenção';
      
      const categorias = {};
      dados.ordens.forEach(ordem => {
        const cat = ordem.CategoriaEquipamento || 'Geral';
        categorias[cat] = (categorias[cat] || 0) + 1;
      });
      const categoriasTop = Object.entries(categorias)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 3)
        .map(([cat, count]) => `${cat} (${count})`)
        .join(', ');
      
      return `
        <div class="technician-card">
          <div class="technician-card__header">
            <div>
              <div class="technician-card__name">${tecnico}</div>
              <div class="technician-card__role">${funcao} - ${area}</div>
            </div>
            <div class="technician-card__count">${dados.ordens.length}</div>
          </div>
          <div class="technician-card__stats">
            <div class="technician-card__stat">
              <div class="technician-card__stat-label">Horas Totais</div>
              <div class="technician-card__stat-value">${formatarHoras(dados.horasTotais)}</div>
            </div>
            <div class="technician-card__stat">
              <div class="technician-card__stat-label">Quantidade</div>
              <div class="technician-card__stat-value">${dados.ordens.length}</div>
            </div>
          </div>
          <div class="technician-card__breakdown">
            <div class="technician-card__breakdown-title">Especializações Alocadas</div>
            <div style="font-size: var(--font-size-sm); color: var(--color-text-secondary); margin-bottom: var(--space-12);">
              ${categoriasTop || 'Equipamentos gerais'}
            </div>
            <div class="technician-card__breakdown-title">Breakdown por Tipo</div>
            <div class="breakdown-bar">
              ${prevPct > 0 ? `<div class="breakdown-segment breakdown-segment--preventiva" style="width: ${prevPct}%">${dados.preventiva}</div>` : ''}
              ${corrPct > 0 ? `<div class="breakdown-segment breakdown-segment--corretiva" style="width: ${corrPct}%">${dados.corretiva}</div>` : ''}
              ${progPct > 0 ? `<div class="breakdown-segment breakdown-segment--programada" style="width: ${progPct}%">${dados.programada}</div>` : ''}
            </div>
            <div class="breakdown-legend">
              <div class="breakdown-legend-item">
                <div class="breakdown-legend-color breakdown-segment--preventiva"></div>
                <span>Preventiva (${dados.preventiva})</span>
              </div>
              <div class="breakdown-legend-item">
                <div class="breakdown-legend-color breakdown-segment--corretiva"></div>
                <span>Corretiva (${dados.corretiva})</span>
              </div>
              <div class="breakdown-legend-item">
                <div class="breakdown-legend-color breakdown-segment--programada"></div>
                <span>Programada (${dados.programada})</span>
              </div>
            </div>
          </div>
        </div>
      `;
    }).join('');
}

// ============================================================================
// TIMELINE VISUAL
// ============================================================================
function atualizarTimeline() {
  const container = document.getElementById('timelineContainer');
  const ordemsPorData = {};
  agendaProcessada.forEach(ordem => {
    if (!ordemsPorData[ordem.Date]) ordemsPorData[ordem.Date] = [];
    ordemsPorData[ordem.Date].push(ordem);
  });
  
  let maxDias = 7;
  if (periodoAtual === 'diario') maxDias = 1;
  else if (periodoAtual === 'semanal') maxDias = 7;
  else if (periodoAtual === 'mensal') maxDias = 30;
  
  const datas = Object.keys(ordemsPorData).sort().slice(0, maxDias);
  const tecnicosComOrdens = new Set();
  agendaProcessada.forEach(o => tecnicosComOrdens.add(o.Technician));
  const tecnicosAtivos = tecnicos.filter(t => tecnicosComOrdens.has(t));
  
  container.innerHTML = datas.map(data => {
    const ordens = ordemsPorData[data];
    const ordemsPorTecnico = {};
    tecnicosAtivos.forEach(tec => { ordemsPorTecnico[tec] = ordens.filter(o => o.Technician === tec); });
    
    return `
      <div class="timeline-day">
        <div class="timeline-day__header">${formatarData(data)}</div>
        <div class="timeline-day__grid">
          ${tecnicosAtivos.map((tec, idx) => {
            const ordensDoTec = ordemsPorTecnico[tec];
            const funcao = funcoesTecnicos[tec] || 'Técnico';
            return `
              <div class="timeline-technician">
                <strong>${tec}</strong><br>
                <span style="font-size: var(--font-size-sm); color: var(--color-text-secondary);">${funcao}</span>
              </div>
              <div class="timeline-bars">
                ${ordensDoTec.length > 0 ? ordensDoTec.map(o => `
                  <div class="timeline-bar tech-color-${idx % 6}" title="${o.Ordem} - ${o.Type}">
                    ${o.StartTime.substring(0, 5)} - ${o.EndTime.substring(0, 5)} (${formatarHoras(o.Duration)})
                  </div>
                `).join('') : '<span class="text-secondary" style="font-size: var(--font-size-sm);">Sem ordens</span>'}
              </div>
            `;
          }).join('')}
        </div>
      </div>
    `;
  }).join('');
}

// ============================================================================
// GRÁFICOS E ESTATÍSTICAS
// ============================================================================
function renderizarGraficos() {
  renderizarGraficoCriticidade();
  renderizarGraficoTipo();
  renderizarTabelaMetricas();
}

function renderizarGraficoCriticidade() {
  const canvas = document.getElementById('chartCriticidade');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  if (chartCriticidade) chartCriticidade.destroy();
  
  const criticidades = { A: 0, B: 0, C: 0 };
  agendaProcessada.forEach(o => criticidades[o.Criticality]++);
  
  chartCriticidade = new Chart(ctx, {
    type: 'pie',
    data: {
      labels: ['Criticidade A', 'Criticidade B', 'Criticidade C'],
      datasets: [{
        data: [criticidades.A, criticidades.B, criticidades.C],
        backgroundColor: ['#B4413C', '#FFC185', '#5D878F']
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: { legend: { position: 'bottom' } }
    }
  });
}

function renderizarGraficoTipo() {
  const canvas = document.getElementById('chartTipo');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  if (chartTipo) chartTipo.destroy();
  
  const tipos = {};
  agendaProcessada.forEach(o => { tipos[o.Type] = (tipos[o.Type] || 0) + 1; });
  
  chartTipo = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: Object.keys(tipos),
      datasets: [{
        label: 'Quantidade de Ordens',
        data: Object.values(tipos),
        backgroundColor: ['#1FB8CD', '#FFC185', '#B4413C', '#D2BA4C']
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: { legend: { display: false } },
      scales: { y: { beginAtZero: true } }
    }
  });
}

function renderizarTabelaMetricas() {
  const tbody = document.getElementById('metricasTableBody');
  const totalOrdens = agendaProcessada.length;
  const ordensComEspecialista = agendaProcessada.filter(o => {
    const categoria = detectarCategoriaEquipamento(o.NomeBem || '', o.Type);
    return categoria.tecnicos_ideais.includes(o.Technician);
  }).length;
  const percentualEspecialista = totalOrdens > 0 ? Math.round((ordensComEspecialista / totalOrdens) * 100) : 0;
  
  const dadosPorTecnico = {};
  tecnicos.forEach(tec => {
    dadosPorTecnico[tec] = { ordens: 0, horas: 0, preventiva: 0, corretiva: 0, programada: 0 };
  });
  
  agendaProcessada.forEach(ordem => {
    const tec = dadosPorTecnico[ordem.Technician];
    if (tec) {
      tec.ordens++;
      tec.horas += ordem.Duration;
      if (ordem.Type === 'PREVENTIVA') tec.preventiva++;
      else if (ordem.Type === 'CORRETIVA' || ordem.Type === 'EMERGENCIAL') tec.corretiva++;
      else if (ordem.Type === 'CORRETIVA PROGRAMADA') tec.programada++;
    }
  });
  
  const horasTotais = Object.values(dadosPorTecnico).map(d => d.horas).filter(h => h > 0);
  const mediaHoras = horasTotais.reduce((a, b) => a + b, 0) / horasTotais.length;
  const desvioPadrao = Math.sqrt(horasTotais.reduce((acc, h) => acc + Math.pow(h - mediaHoras, 2), 0) / horasTotais.length);
  
  const statsContainer = document.querySelector('.metrics-table');
  if (statsContainer && !document.getElementById('allocationStats')) {
    const statsDiv = document.createElement('div');
    statsDiv.id = 'allocationStats';
    statsDiv.innerHTML = `
      <div style="background: var(--color-bg-3); border: 1px solid var(--color-success); border-radius: var(--radius-md); padding: var(--space-20); margin-bottom: var(--space-24);">
        <h4 style="font-size: var(--font-size-lg); font-weight: var(--font-weight-semibold); margin-bottom: var(--space-12);">
          Métricas do Algoritmo de Balanceamento
        </h4>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: var(--space-16);">
          <div>
            <div style="font-size: var(--font-size-sm); color: var(--color-text-secondary); margin-bottom: var(--space-4);">Taxa de Especialistas</div>
            <div style="font-size: var(--font-size-2xl); font-weight: var(--font-weight-bold); color: var(--color-success);">${percentualEspecialista}%</div>
            <div style="font-size: var(--font-size-xs); color: var(--color-text-secondary);">${ordensComEspecialista} de ${totalOrdens} ordens</div>
          </div>
          <div>
            <div style="font-size: var(--font-size-sm); color: var(--color-text-secondary); margin-bottom: var(--space-4);">Média de Horas/Técnico</div>
            <div style="font-size: var(--font-size-2xl); font-weight: var(--font-weight-bold); color: var(--color-primary);">${formatarHoras(mediaHoras)}</div>
            <div style="font-size: var(--font-size-xs); color: var(--color-text-secondary);">Meta: ~${formatarHoras(MAX_HORAS_POR_TECNICO * 0.7)}</div>
          </div>
          <div>
            <div style="font-size: var(--font-size-sm); color: var(--color-text-secondary); margin-bottom: var(--space-4);">Desvio Padrão</div>
            <div style="font-size: var(--font-size-2xl); font-weight: var(--font-weight-bold); color: ${desvioPadrao < 20 ? 'var(--color-success)' : desvioPadrao < 40 ? 'var(--color-warning)' : 'var(--color-error)'};">${formatarHoras(desvioPadrao)}</div>
            <div style="font-size: var(--font-size-xs); color: var(--color-text-secondary);">${desvioPadrao < 20 ? 'Excelente balanceamento' : desvioPadrao < 40 ? 'Balanceamento bom' : 'Revisar distribuição'}</div>
          </div>
        </div>
      </div>
    `;
    statsContainer.insertBefore(statsDiv, statsContainer.querySelector('h3'));
  }
  
  tbody.innerHTML = Object.entries(dadosPorTecnico)
    .filter(([tecnico, dados]) => dados.ordens > 0)
    .sort((a, b) => b[1].ordens - a[1].ordens)
    .map(([tecnico, dados]) => {
      const utilizacao = Math.round((dados.horas / MAX_HORAS_POR_TECNICO) * 100);
      const funcao = funcoesTecnicos[tecnico] || 'Técnico';
      const totalTipo = dados.preventiva + dados.corretiva + dados.programada;
      const prevPct = totalTipo > 0 ? Math.round((dados.preventiva / totalTipo) * 100) : 0;
      const corrPct = totalTipo > 0 ? Math.round((dados.corretiva / totalTipo) * 100) : 0;
      const progPct = totalTipo > 0 ? Math.round((dados.programada / totalTipo) * 100) : 0;
      
      return `
        <tr>
          <td>
            <strong>${tecnico}</strong><br>
            <span style="font-size: var(--font-size-sm); color: var(--color-text-secondary);">${funcao}</span>
          </td>
          <td><strong>${dados.ordens}</strong></td>
          <td>${formatarHoras(dados.horas)}</td>
          <td>${dados.preventiva} <span style="font-size: var(--font-size-xs); color: var(--color-text-secondary);">(${prevPct}%)</span></td>
          <td>${dados.corretiva} <span style="font-size: var(--font-size-xs); color: var(--color-text-secondary);">(${corrPct}%)</span></td>
          <td>${dados.programada} <span style="font-size: var(--font-size-xs); color: var(--color-text-secondary);">(${progPct}%)</span></td>
          <td>
            <span style="color: ${utilizacao > 80 ? 'var(--color-error)' : utilizacao > 60 ? 'var(--color-warning)' : 'var(--color-success)'}; font-weight: var(--font-weight-semibold);">${utilizacao}%</span>
          </td>
        </tr>
      `;
    }).join('');
}

// ============================================================================
// MODO DE EDIÇÃO
// ============================================================================
function entrarModoEdicao() {
  if (agendaProcessada.length === 0) {
    showAlert('Nenhuma ordem para editar!', 'error');
    return;
  }
  isEditMode = true;
  agendaBackup = JSON.parse(JSON.stringify(agendaProcessada));
  editedRows.clear();
  document.getElementById('editModeHeader').style.display = 'flex';
  document.getElementById('editAgendaBtn').style.display = 'none';
  document.getElementById('actionsHeader').style.display = 'table-cell';
  document.getElementById('agendaTable').classList.add('table--edit-mode');
  atualizarTabelaAgenda();
  mostrarResumoEdicao();
  showAlert('Modo de edição ativado. Faça as alterações necessárias.', 'success');
}

function salvarAlteracoes() {
  if (!isEditMode) return;
  agendaProcessada.forEach(ordem => {
    const index = allAgenda.findIndex(o => o.Ordem === ordem.Ordem);
    if (index !== -1) allAgenda[index] = { ...ordem };
  });
  sairModoEdicao();
  atualizarInterface();
  showAlert(`Alterações salvas com sucesso! ${editedRows.size} linha(s) modificada(s).`, 'success');
}

function cancelarEdicao() {
  if (!isEditMode) return;
  if (editedRows.size > 0) {
    if (!confirm('Deseja cancelar as alterações? Todas as modificações serão perdidas.')) return;
  }
  agendaProcessada = JSON.parse(JSON.stringify(agendaBackup));
  sairModoEdicao();
  atualizarTabelaAgenda();
  showAlert('Edição cancelada. Dados restaurados.', 'success');
}

function desfazerAlteracoes() {
  if (!isEditMode) return;
  if (editedRows.size === 0) {
    showAlert('Nenhuma alteração para desfazer.', 'error');
    return;
  }
  agendaProcessada = JSON.parse(JSON.stringify(agendaBackup));
  editedRows.clear();
  atualizarTabelaAgenda();
  atualizarDashboard();
  mostrarResumoEdicao();
  showAlert('Alterações desfeitas.', 'success');
}

function sairModoEdicao() {
  isEditMode = false;
  editedRows.clear();
  document.getElementById('editModeHeader').style.display = 'none';
  document.getElementById('editAgendaBtn').style.display = 'inline-flex';
  document.getElementById('actionsHeader').style.display = 'none';
  document.getElementById('agendaTable').classList.remove('table--edit-mode');
  const summary = document.getElementById('editSummary');
  if (summary) summary.remove();
}

function editarCelula(ordemNum, campo, novoValor) {
  const ordem = agendaProcessada.find(o => o.Ordem === ordemNum);
  if (!ordem) return;
  const valorAntigo = ordem[campo];
  if (valorAntigo === novoValor) return;
  ordem[campo] = novoValor;
  editedRows.add(ordemNum);
  atualizarDashboard();
  mostrarResumoEdicao();
}

function mostrarResumoEdicao() {
  let summary = document.getElementById('editSummary');
  if (!summary) {
    summary = document.createElement('div');
    summary.id = 'editSummary';
    summary.className = 'edit-summary';
    const tableContainer = document.querySelector('#agenda .table-container');
    tableContainer.parentNode.insertBefore(summary, tableContainer);
  }
  
  const total = agendaProcessada.length;
  const agendadas = agendaProcessada.filter(o => o.Status !== 'Pendente' && o.Status !== 'Cancelada').length;
  const pendentes = agendaProcessada.filter(o => o.Status === 'Pendente').length;
  const canceladas = agendaProcessada.filter(o => o.Status === 'Cancelada').length;
  const taxa = total > 0 ? Math.round((agendadas / total) * 100) : 0;
  const tecnicosUsados = new Set(agendaProcessada.map(o => o.Technician)).size;
  
  summary.innerHTML = `
    <div class="edit-summary__title">📊 Resumo Atualizado em Tempo Real</div>
    <div class="edit-summary__stats">
      <div class="edit-summary__stat">
        <span class="edit-summary__stat-label">Total de Ordens:</span>
        <span class="edit-summary__stat-value">${total}</span>
      </div>
      <div class="edit-summary__stat">
        <span class="edit-summary__stat-label">Agendadas:</span>
        <span class="edit-summary__stat-value" style="color: var(--color-success);">${agendadas}</span>
      </div>
      <div class="edit-summary__stat">
        <span class="edit-summary__stat-label">Pendentes:</span>
        <span class="edit-summary__stat-value" style="color: var(--color-warning);">${pendentes}</span>
      </div>
      <div class="edit-summary__stat">
        <span class="edit-summary__stat-label">Canceladas:</span>
        <span class="edit-summary__stat-value" style="color: var(--color-error);">${canceladas}</span>
      </div>
      <div class="edit-summary__stat">
        <span class="edit-summary__stat-label">Taxa de Agendamento:</span>
        <span class="edit-summary__stat-value" style="color: var(--color-primary);">${taxa}%</span>
      </div>
      <div class="edit-summary__stat">
        <span class="edit-summary__stat-label">Técnicos Alocados:</span>
        <span class="edit-summary__stat-value">${tecnicosUsados}</span>
      </div>
      <div class="edit-summary__stat">
        <span class="edit-summary__stat-label">Alterações Feitas:</span>
        <span class="edit-summary__stat-value" style="color: var(--color-warning);">${editedRows.size}</span>
      </div>
    </div>
  `;
}

function removerLinha(ordemNum) {
  if (!confirm('Deseja realmente remover esta ordem?')) return;
  const index = agendaProcessada.findIndex(o => o.Ordem === ordemNum);
  if (index !== -1) {
    agendaProcessada.splice(index, 1);
    editedRows.add(ordemNum);
    atualizarTabelaAgenda();
    atualizarDashboard();
    mostrarResumoEdicao();
    showAlert('Linha removida com sucesso.', 'success');
  }
}

// ============================================================================
// DOWNLOAD EXCEL
// ============================================================================
function downloadExcel() {
  if (agendaProcessada.length === 0) {
    showAlert('Nenhuma ordem para exportar!', 'error');
    return;
  }

  const exportData = agendaProcessada.map(ordem => ({
    'Técnico': ordem.Technician,
    'Ordem': ordem.Ordem,
    'Bem': ordem.Bem || '-',
    'Nome Bem': ordem.NomeBem || '-',
    'Tipo Serviço': ordem.Type,
    'Criticidade': ordem.Criticality,
    'Dias Atraso': ordem.Delay,
    'Prioridade': ordem.Priority,
    'Data Agendada': formatarData(ordem.Date),
    'Status': ordem.Status || 'Agendada',
    'Previsão Início': formatarData(ordem.PrevisaoInicio || '-'),
    'Data Abertura': formatarData(ordem.DataAbertura || '-'),
    'Real Início': ordem.RealInicio ? formatarData(ordem.RealInicio) : '-',
    'Descrição': ordem.Descricao || '-',
    'Horário Início': ordem.StartTime.substring(0, 5),
    'Horário Fim': ordem.EndTime.substring(0, 5),
    'Duração': formatarHoras(ordem.Duration),
    'Motivo Alocação': ordem.MotivoAlocacao || 'N/A',
    'Categoria': ordem.CategoriaEquipamento || 'N/A'
  }));

  const ws = XLSX.utils.json_to_sheet(exportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Agenda');

  const dataAtual = new Date().toISOString().split('T')[0];
  XLSX.writeFile(wb, `Agenda_Manutencao_${dataAtual}.xlsx`);

  showAlert('Planilha exportada com sucesso!', 'success');
}


function showAlert(message, type = 'success') {
  const container = document.getElementById('alertContainer');
  const alert = document.createElement('div');
  alert.className = `alert alert--${type}`;
  alert.innerHTML = `
    <span style="font-weight: bold;">${type === 'success' ? '✓' : '✕'}</span>
    <span>${message}</span>
  `;
  container.appendChild(alert);
  setTimeout(() => {
    alert.style.opacity = '0';
    setTimeout(() => container.removeChild(alert), 300);
  }, 4000);
}

console.log('✅ Sistema OLLI carregado com sucesso!');