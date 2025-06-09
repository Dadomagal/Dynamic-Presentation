// --- CONFIGURAÇÕES GLOBAIS ---
// Estes IDs são placeholders genéricos. Para usar este script em sua própria apresentação,
// você precisará substituí-los pelos IDs reais da sua apresentação e da sua Cloud Function.

// ID da sua apresentação do Google Slides.
// Você pode encontrá-lo na URL da apresentação, geralmente é a parte entre
// "/d/" e "/edit/":
// Exemplo: https://docs.google.com/presentation/d/SEU_ID_DA_APRESENTACAO/edit
const PRESENTATION_ID = 'ID_GENERICO_DA_APRESENTACAO'; 

// URL da sua Cloud Function.
// Certifique-se de que esta URL está correta e que sua Cloud Function
// está implantada e acessível.
const CLOUD_FUNCTION_URL = 'https://sua-cloud-function-url-generica.run.app'; 

// Tempo em minutos para atrasar a execução da análise após a abertura da apresentação.
// Isso dá um tempo para o usuário interagir ou visualizar os slides iniciais.
const ANALYSIS_START_DELAY_MINUTES = 0.5; // 30 segundos

// Mensagens padrão para feedback ao usuário sobre o status da análise.
const LOADING_MESSAGE_TEXT = "Executando análises...";
const ERROR_MESSAGE_PREFIX = "Erro na análise:"; // Prefixo para identificar caixas de erro

// ID do slide onde a mensagem de carregamento e erros serão exibidos.
// Geralmente, é o primeiro slide da apresentação. Para encontrar o ID de um slide,
// navegue até o slide desejado e copie a parte final da URL após '#slide=id.'.
const LOADING_MESSAGE_SLIDE_ID = 'ID_DO_SLIDE_DE_CARREGAMENTO'; 

// IDs dos slides que serão preenchidos com os resultados da análise.
// Estes IDs são específicos da sua apresentação e DEVEM ser verificados e atualizados.
const SLIDE_6_ID = 'ID_SLIDE_PARA_EXPLORATORIA'; 
const SLIDE_7_ID = 'ID_SLIDE_PARA_DISTRIBUICOES';
const SLIDE_8_ID = 'ID_SLIDE_PARA_ASSOCCIACOES';
const SLIDE_9_ID = 'ID_SLIDE_PARA_MODELO';
const SLIDE_10_ID = 'ID_SLIDE_PARA_FATORES_RISCO';
const SLIDE_11_ID = 'ID_SLIDE_PARA_IMPACTO_CALL_CENTER';
const SLIDE_12_ID = 'ID_SLIDE_PARA_INSIGHTS_FINAIS';

// Lista de IDs dos slides que devem ter seu conteúdo limpo ao abrir a apresentação.
// Isso garante que a apresentação comece "vazia" para ser preenchida dinamicamente.
const SLIDES_TO_CLEAR_ON_OPEN = [
  SLIDE_6_ID,
  SLIDE_7_ID,
  SLIDE_8_ID,
  SLIDE_9_ID,
  SLIDE_10_ID,
  SLIDE_11_ID,
  SLIDE_12_ID
];

// --- CONFIGURAÇÕES DE LAYOUT E ESTILO ---
// Estas dimensões são baseadas no layout padrão do Google Slides em pontos.
const SLIDE_WIDTH = 960; 
const SLIDE_HEIGHT = 540; 

// Margens padrão para posicionamento de elementos nos slides.
const MARGIN_LEFT = 50;
const MARGIN_RIGHT = 50;
const MARGIN_TOP = 50;
const MARGIN_BOTTOM = 50;

// Largura e altura disponíveis para conteúdo dentro das margens.
const CONTENT_WIDTH = SLIDE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT; 
const CONTENT_HEIGHT = SLIDE_HEIGHT - MARGIN_TOP - MARGIN_BOTTOM; 

// Estilos de texto e título padrão para manter consistência visual.
const TITLE_FONT_SIZE = 22;
const TITLE_COLOR = '#1f4e79'; 
const TEXT_FONT_SIZE = 12; 
const TEXT_LINE_HEIGHT = 20; 

// Ajustes específicos para o layout e tamanho de fonte de slides individuais.
const SLIDE6_TEXT_FONT_SIZE = 14; 
const SLIDE6_TEXT_INDENT = 20; 

const SLIDE12_TEXT_FONT_SIZE = 9; 
const SLIDE12_TEXT_LINE_HEIGHT = 12; 

// Chaves para persistir o status da análise nas Propriedades do Usuário.
// Isso permite que o script saiba o estado da análise mesmo após o fechamento da apresentação.
const ANALYSIS_STATUS_KEY = 'analysis_status';
const STATUS_RUNNING = 'RUNNING';
const STATUS_COMPLETED = 'COMPLETED';
const STATUS_NOT_STARTED = 'NOT_STARTED';
const STATUS_ERROR = 'ERROR';


// --- FUNÇÕES PRINCIPAIS ---

/**
 * Função padrão do Google Apps Script que é executada automaticamente ao abrir a apresentação.
 * Cria o menu personalizado, limpa os slides e programa a execução da análise.
 */
function onOpen() {
  // Cria um menu personalizado "Análises" na barra de ferramentas do Google Slides.
  SlidesApp.getUi()
    .createMenu('🔧 Análises')
    .addItem('🚀 Executar Análise', 'executarAnaliseMenu')
    .addItem('🧹 Limpar Dados', 'limparDadosMenu')
    .addItem('📊 Status da Análise', 'verificarStatusAnalise')
    .addToUi();

  // Limpa o conteúdo dos slides de destino imediatamente ao abrir a apresentação.
  clearTargetSlidesOnOpen();

  // Exibe uma mensagem de carregamento no slide inicial.
  exibirMensagemCarregamento();

  // Configura um gatilho de tempo para iniciar a análise completa automaticamente após um atraso.
  ScriptApp.newTrigger('executarAnaliseCompleta')
    .timeBased()
    .at(new Date(Date.now() + ANALYSIS_START_DELAY_MINUTES * 60 * 1000))
    .create();

  // Define o status inicial da análise como "Não Iniciada".
  PropertiesService.getUserProperties().setProperty(ANALYSIS_STATUS_KEY, STATUS_NOT_STARTED);

  console.log('onOpen executado. Menu "Análises" criado e gatilho programado.');
}

/**
 * Função acionada pelo item de menu "Executar Análise".
 * Inicia o processo de análise completa.
 */
function executarAnaliseMenu() {
  try {
    console.log('Análise iniciada pelo menu...');

    // Remove mensagens anteriores e exibe carregamento.
    removerMensagemCarregamento();
    exibirMensagemCarregamento();

    // Define o status como "Em Execução" antes de iniciar a análise.
    PropertiesService.getUserProperties().setProperty(ANALYSIS_STATUS_KEY, STATUS_RUNNING);

    // Executa a análise completa.
    executarAnaliseCompleta();

  } catch (e) {
    console.error('Erro ao executar análise pelo menu: ' + e.message);
    // Em caso de erro, define o status como "Erro" e exibe uma mensagem no slide.
    PropertiesService.getUserProperties().setProperty(ANALYSIS_STATUS_KEY, STATUS_ERROR);

    const presentation = SlidesApp.openById(PRESENTATION_ID);
    const slide = presentation.getSlideById(LOADING_MESSAGE_SLIDE_ID);
    if (slide) {
      removerMensagemCarregamento();
      slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH, TEXT_LINE_HEIGHT * 3)
        .getText().setText(ERROR_MESSAGE_PREFIX + ' ' + e.message)
        .getTextStyle().setFontSize(TEXT_FONT_SIZE).setForegroundColor('#FF0000');
    }
  }
}

/**
 * Função acionada pelo item de menu "Limpar Dados".
 * Limpa o conteúdo dos slides de dados e redefine o status da análise.
 * Exibe um modal de confirmação.
 */
function limparDadosMenu() {
  try {
    console.log('Limpando dados dos slides 6-12...');

    const presentation = SlidesApp.openById(PRESENTATION_ID);
    let slidesLimpos = 0;

    // Limpa cada slide da lista.
    SLIDES_TO_CLEAR_ON_OPEN.forEach(slideId => {
      try {
        const slide = presentation.getSlideById(slideId);
        clearSlideContent(slide);
        slidesLimpos++;
        console.log(`Slide ${slideId} limpo com sucesso.`);
      } catch (e) {
        console.warn(`Aviso: Não foi possível limpar o slide ${slideId}. Erro: ${e.message}`);
      }
    });

    // Remove mensagens de carregamento/erro do slide inicial.
    removerMensagemCarregamento();

    // Redefine o status da análise para "Não Iniciada" após a limpeza.
    PropertiesService.getUserProperties().setProperty(ANALYSIS_STATUS_KEY, STATUS_NOT_STARTED);

    // Exibe confirmação de limpeza usando modal HTML.
    exibirModalCustomizado('Limpeza Concluída', `✓ Dados limpos com sucesso!\n${slidesLimpos} slides processados.`, '#008000');

    console.log(`Limpeza concluída. ${slidesLimpos} slides processados.`);

  } catch (e) {
    console.error('Erro ao limpar dados: ' + e.message);

    // Exibe erro usando modal HTML.
    exibirModalCustomizado('Erro na Limpeza', `❌ Erro ao limpar dados: ${e.message}`, '#FF0000');
  }
}

/**
 * Função acionada pelo item de menu "Status da Análise".
 * Verifica e exibe o status atual da análise.
 */
function verificarStatusAnalise() {
  try {
    console.log('Verificando status da análise...');

    const presentation = SlidesApp.openById(PRESENTATION_ID);

    // Remove mensagens anteriores.
    removerMensagemCarregamento();

    // Obtém o status persistente.
    const analysisStatus = PropertiesService.getUserProperties().getProperty(ANALYSIS_STATUS_KEY);

    // Verifica se os slides têm conteúdo (análise já executada).
    let slidesPreenchidos = 0;
    SLIDES_TO_CLEAR_ON_OPEN.forEach(slideId => {
      try {
        const slide = presentation.getSlideById(slideId);
        const elements = slide.getPageElements();
        if (elements.length > 0) {
          slidesPreenchidos++;
        }
      } catch (e) {
        console.warn(`Não foi possível verificar o slide ${slideId}`);
      }
    });

    // Determina o status a ser exibido.
    let statusTitle = '';
    let statusText = '';
    let statusColor = '#555555';

    if (analysisStatus === STATUS_RUNNING) {
      statusTitle = 'Status da Análise: EM EXECUÇÃO';
      statusText = '⏳ A análise está em execução...\nAguarde a conclusão.';
      statusColor = '#FFA500'; // Laranja
    } else if (analysisStatus === STATUS_COMPLETED || slidesPreenchidos > 0) { // Considera COMPLETED ou se já há conteúdo.
      statusTitle = 'Status da Análise: CONCLUÍDA';
      statusText = `✅ Análise concluída com sucesso!\n${slidesPreenchidos} de ${SLIDES_TO_CLEAR_ON_OPEN.length} slides preenchidos.`;
      statusColor = '#008000'; // Verde
    } else if (analysisStatus === STATUS_ERROR) {
      statusTitle = 'Status da Análise: ERRO';
      statusText = '❌ Houve um erro na última execução da análise.';
      statusColor = '#FF0000'; // Vermelho
    } else { // STATUS_NOT_STARTED ou null.
      statusTitle = 'Status da Análise: NÃO EXECUTADA';
      statusText = '⚪ Nenhuma análise executada.\nUse "Executar Análise" para iniciar.';
      statusColor = '#666666'; // Cinza
    }

    // Exibe o status no modal HTML.
    exibirModalCustomizado(statusTitle, statusText, statusColor);

    console.log('Status verificado: ' + statusText.replace('\n', ' '));

  } catch (e) {
    console.error('Erro ao verificar status: ' + e.message);

    // Exibe erro no modal HTML.
    exibirModalCustomizado('Erro ao Verificar Status', `❌ Erro ao verificar status: ${e.message}`, '#FF0000');
  }
}


/**
 * Limpa todo o conteúdo dos slides especificados na lista SLIDES_TO_CLEAR_ON_OPEN.
 * Esta função é chamada ao abrir a apresentação para garantir que os slides de dados estejam vazios.
 * @param {GoogleAppsScript.Slides.Slide} slide O objeto do slide a ser limpo.
 */
function clearTargetSlidesOnOpen() {
  const presentation = SlidesApp.openById(PRESENTATION_ID);
  SLIDES_TO_CLEAR_ON_OPEN.forEach(slideId => {
    try {
      const slide = presentation.getSlideById(slideId);
      clearSlideContent(slide); 
      console.log(`Slide ${slideId} limpo ao abrir a apresentação.`);
    } catch (e) {
      console.warn(`Aviso: Não foi possível limpar o slide ${slideId}. Ele pode não existir ou ter um ID incorreto. Erro: ${e.message}`);
    }
  });
}


/**
 * Exibe uma mensagem de carregamento no slide especificado (LOADING_MESSAGE_SLIDE_ID).
 * Se uma mensagem antiga existir, ela é removida para evitar duplicatas.
 */
function exibirMensagemCarregamento() {
  const presentation = SlidesApp.openById(PRESENTATION_ID);
  const slide = presentation.getSlideById(LOADING_MESSAGE_SLIDE_ID);

  // Remove qualquer caixa de texto de carregamento antiga ou de erro.
  const shapes = slide.getShapes();
  for (let i = 0; i < shapes.length; i++) {
    const textContent = shapes[i].getText().asString();
    if (textContent.includes(LOADING_MESSAGE_TEXT.substring(0, 10)) || textContent.includes(ERROR_MESSAGE_PREFIX)) {
      shapes[i].remove();
      console.log('Mensagem de carregamento/erro antiga removida.');
    }
  }

  // Adiciona a nova caixa de texto com a mensagem de carregamento.
  const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH / 2, TEXT_LINE_HEIGHT * 2);
  textBox.getText().setText(LOADING_MESSAGE_TEXT);
  // Estiliza o texto.
  textBox.getText().getTextStyle().setFontSize(TEXT_FONT_SIZE).setForegroundColor('#555555').setBold(true);
  textBox.setRotation(0); 
  console.log('Mensagem de carregamento exibida.');
}

/**
 * Remove a mensagem de carregamento do slide especificado (LOADING_MESSAGE_SLIDE_ID).
 * É chamada após a conclusão bem-sucedida do preenchimento dos slides.
 */
function removerMensagemCarregamento() {
  try {
    const presentation = SlidesApp.openById(PRESENTATION_ID);
    const slide = presentation.getSlideById(LOADING_MESSAGE_SLIDE_ID);
    if (!slide) { 
      console.warn('Aviso: Slide de mensagem de carregamento não encontrado para remoção.');
      return;
    }
    const shapes = slide.getShapes();
    for (let i = 0; i < shapes.length; i++) {
      // Verifica se o texto da forma contém o início da mensagem de carregamento ou de erro.
      const textContent = shapes[i].getText().asString();
      if (textContent.includes(LOADING_MESSAGE_TEXT.substring(0, 10)) || textContent.includes(ERROR_MESSAGE_PREFIX)) {
        shapes[i].remove();
        console.log('Mensagem de carregamento/erro removida.');
      }
    }
  } catch (e) {
    console.error('Erro ao remover mensagem de carregamento/erro: ' + e.message);
  }
}

/**
 * Função principal para executar a análise completa e preencher os slides.
 * Esta função é acionada pelo gatilho de tempo ou pelo menu manual.
 */
function executarAnaliseCompleta() {
  try {
    console.log('Iniciando execução da análise completa...');
    removerMensagemCarregamento(); 
    exibirMensagemCarregamento(); 

    // Define o status como "Em Execução" ao iniciar a análise.
    PropertiesService.getUserProperties().setProperty(ANALYSIS_STATUS_KEY, STATUS_RUNNING);

    // Faz a chamada à Cloud Function para obter os dados da análise.
    const response = callCloudFunction();
    if (!response) {
      console.error('Falha ao obter resposta da Cloud Function. Abortando preenchimento.');
      PropertiesService.getUserProperties().setProperty(ANALYSIS_STATUS_KEY, STATUS_ERROR); 
      return;
    }

    // Analisa a resposta JSON da Cloud Function.
    const data = JSON.parse(response).data;
    console.log('Dados da Cloud Function recebidos. Preenchendo slides...');

    // Chama as funções para preencher cada slide com os dados correspondentes.
    preencherSlide6(data.analise_exploratoria);
    preencherSlide7(data.distribuicoes);
    preencherSlide8(data.associacoes);
    preencherSlide9(data.modelo);
    preencherSlide10(data.fatores_risco);
    preencherSlide11(data.call_center_impact);
    preencherSlide12(data.insights);

    // Remove a mensagem de carregamento após todos os slides serem preenchidos.
    removerMensagemCarregamento();
    console.log('Análise completa concluída e slides preenchidos.');

    // Define o status como "Concluída" após o sucesso.
    PropertiesService.getUserProperties().setProperty(ANALYSIS_STATUS_KEY, STATUS_COMPLETED);

    // Exibe o modal de conclusão.
    exibirModalConclusao();

  } catch (e) {
    console.error('Erro na execução da análise completa: ' + e.message);
    removerMensagemCarregamento();
    // Adiciona uma caixa de texto de erro no slide para depuração.
    const presentation = SlidesApp.openById(PRESENTATION_ID);
    const slide = presentation.getSlideById(LOADING_MESSAGE_SLIDE_ID);
    if (slide) { 
      slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, MARGIN_TOP + 50, CONTENT_WIDTH, TEXT_LINE_HEIGHT * 3)
        .getText().setText(ERROR_MESSAGE_PREFIX + ' ' + e.message)
        .getTextStyle().setFontSize(TEXT_FONT_SIZE).setForegroundColor('#FF0000');
    }
    // Define o status como "Erro".
    PropertiesService.getUserProperties().setProperty(ANALYSIS_STATUS_KEY, STATUS_ERROR);
  } finally {
    // Apaga os gatilhos de tempo para evitar execuções futuras desnecessárias.
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'executarAnaliseCompleta') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    console.log('Gatilhos de análise limpos.');
  }
}

/**
 * Faz a chamada HTTP POST para a Cloud Function configurada.
 * @returns {string|null} O corpo da resposta da Cloud Function como string JSON, ou null em caso de erro.
 */
function callCloudFunction() {
  try {
    const options = {
      'method': 'post', 
      'contentType': 'application/json', 
      'payload': JSON.stringify({
        action: 'full_analysis'
      }), 
      'muteHttpExceptions': true 
    };
    const response = UrlFetchApp.fetch(CLOUD_FUNCTION_URL, options); 
    const responseCode = response.getResponseCode(); 
    const responseBody = response.getContentText(); 

    if (responseCode >= 200 && responseCode < 300) {
      console.log('Cloud Function chamada com sucesso. Resposta: ' + responseBody);
      return responseBody;
    } else {
      console.error(`Erro ao chamar Cloud Function: Código ${responseCode}, Mensagem: ${responseBody}`);
      return null;
    }
  } catch (e) {
    console.error('Exceção ao chamar Cloud Function: ' + e.message);
    return null;
  }
}

// --- FUNÇÃO AUXILIAR PARA CAPITALIZAÇÃO (SIMILAR AO .title() do Python) ---
/**
 * Converte uma string para "Title Case", capitalizando a primeira letra de cada palavra
 * e realizando substituições específicas para termos comuns em português.
 * @param {string} str A string de entrada.
 * @returns {string} A string formatada em Title Case.
 */
function toTitleCase(str) {
  if (!str) return '';
  let processedStr = String(str).replace(/_/g, ' '); 

  // Aplica substituições específicas (case-insensitive).
  processedStr = processedStr.replace(/sexo /gi, 'Sexo ');
  processedStr = processedStr.replace(/duracao contrato /gi, 'Duração Contrato ');
  processedStr = processedStr.replace(/assinatura /gi, 'Assinatura ');

  // Capitaliza a primeira letra de cada palavra.
  return processedStr.replace(/\b\w/g, char => char.toUpperCase());
}


// --- FUNÇÕES DE PREENCHIMENTO DOS SLIDES ---

/**
 * Limpa todo o conteúdo de um slide específico (formas, imagens, tabelas, etc.).
 * @param {GoogleAppsScript.Slides.Slide} slide O objeto do slide a ser limpo.
 */
function clearSlideContent(slide) {
  if (!slide) { 
    console.error("Erro: Slide é indefinido ou nulo em clearSlideContent.");
    return;
  }
  const elements = slide.getPageElements();
  for (let i = 0; i < elements.length; i++) {
    elements[i].remove(); 
  }
}

/**
 * Adiciona uma imagem a um slide a partir de uma string base64.
 * @param {GoogleAppsScript.Slides.Slide} slide O slide onde a imagem será adicionada.
 * @param {string} base64Data A string base64 da imagem.
 * @param {number} left A posição X (em pontos) da imagem.
 * @param {number} top A posição Y (em pontos) da imagem.
 * @param {number} width A largura (em pontos) da imagem.
 * @param {number} height A altura (em pontos) da imagem.
 */
function addImageFromBase64(slide, base64Data, left, top, width, height) {
  if (!slide || !base64Data) return;

  try {
    const cleanBase64 = base64Data.replace(/^data:image\/[a-z]+;base64,/, '');
    const blob = Utilities.newBlob(
      Utilities.base64Decode(cleanBase64),
      'image/png',
      'chart.png'
    );

    const image = slide.insertImage(blob);
    image.setLeft(left).setTop(top).setWidth(width).setHeight(height);

  } catch (error) {
    console.error('Erro ao adicionar imagem:', error);
    const errorBox = slide.insertTextBox('Erro ao carregar gráfico');
    errorBox.setLeft(left).setTop(top).setWidth(width).setHeight(50);
    errorBox.getText().getTextStyle().setItalic(true).setForegroundColor('#ff0000');
  }
}


/**
 * Preenche o slide 6 com os dados da análise exploratória.
 * @param {object} data Os dados brutos da análise exploratória.
 */
function preencherSlide6(data) {
  const presentation = SlidesApp.openById(PRESENTATION_ID);
  const slide = presentation.getSlideById(SLIDE_6_ID);
  clearSlideContent(slide); 

  if (data.erro) {
    slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH, TEXT_LINE_HEIGHT * 3)
      .getText().setText('Erro na análise exploratória: ' + data.erro)
      .getTextStyle().setFontSize(TEXT_FONT_SIZE).setForegroundColor('#FF0000');
    return;
  }

  let yPos = MARGIN_TOP;
  // Título.
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, yPos, CONTENT_WIDTH, TITLE_FONT_SIZE + 10)
    .getText().setText('Análise Exploratória dos Dados') 
    .getTextStyle().setFontSize(TITLE_FONT_SIZE).setBold(true).setForegroundColor(TITLE_COLOR);
  yPos += TITLE_FONT_SIZE + 20; 

  // Conteúdo de texto.
  yPos += SLIDE6_TEXT_INDENT;

  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, yPos, CONTENT_WIDTH, SLIDE6_TEXT_FONT_SIZE + 5) 
    .getText().setText(`Dataset: ${data.total_registros} registros`) 
    .getTextStyle().setFontSize(SLIDE6_TEXT_FONT_SIZE); 
  yPos += SLIDE6_TEXT_FONT_SIZE + 10; 
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, yPos, CONTENT_WIDTH, SLIDE6_TEXT_FONT_SIZE + 5)
    .getText().setText(`Variáveis: ${data.total_variaveis}`) 
    .getTextStyle().setFontSize(SLIDE6_TEXT_FONT_SIZE); 
  yPos += SLIDE6_TEXT_FONT_SIZE + 10;
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, yPos, CONTENT_WIDTH, SLIDE6_TEXT_FONT_SIZE + 5)
    .getText().setText(`Taxa de cancelamento: ${data.taxa_cancelamento}%`) 
    .getTextStyle().setFontSize(SLIDE6_TEXT_FONT_SIZE); 
  yPos += SLIDE6_TEXT_FONT_SIZE + 10;
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, yPos, CONTENT_WIDTH, SLIDE6_TEXT_FONT_SIZE + 5)
    .getText().setText(`Registros válidos: ${data.registros_validos}`) 
    .getTextStyle().setFontSize(SLIDE6_TEXT_FONT_SIZE); 
  yPos += SLIDE6_TEXT_FONT_SIZE + 10;

  // Imagem no Slide 6 (se houver).
  if (data.imagem_base64) {
    const imgWidth = 300;
    const imgHeight = 280;
    const imgLeft = 420;
    const imgTop = 100;

    addImageFromBase64(slide, data.imagem_base64, imgLeft, imgTop, imgWidth, imgHeight);
  }

  console.log('Slide 6 preenchido.');
}

/**
 * Preenche o slide 7 com os dados das distribuições e o gráfico.
 * @param {object} data Os dados das distribuições, incluindo a imagem em base64.
 */
function preencherSlide7(data) {
  const presentation = SlidesApp.openById(PRESENTATION_ID);
  const slide = presentation.getSlideById(SLIDE_7_ID);
  clearSlideContent(slide);

  if (data.erro) {
    slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH, TEXT_LINE_HEIGHT * 3)
      .getText().setText('Erro ao gerar distribuições: ' + data.erro)
      .getTextStyle().setFontSize(TEXT_FONT_SIZE).setForegroundColor('#FF0000');
    return;
  }

  let yPos = MARGIN_TOP;
  // Título.
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, yPos, CONTENT_WIDTH, TITLE_FONT_SIZE + 10)
    .getText().setText('Distribuição das Variáveis Numéricas')
    .getTextStyle().setFontSize(TITLE_FONT_SIZE).setBold(true).setForegroundColor(TITLE_COLOR);
  yPos += TITLE_FONT_SIZE + 20; 

  // Proporções para texto e imagem.
  const textColumnWidth = CONTENT_WIDTH * 0.35;
  const imageColumnWidth = CONTENT_WIDTH * 0.60;
  const spaceBetweenColumns = CONTENT_WIDTH - textColumnWidth - imageColumnWidth;

  let currentTextYPos = yPos;

  // Conteúdo de texto.
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, TEXT_LINE_HEIGHT)
    .getText().setText('Principais estatísticas:')
    .getTextStyle().setFontSize(TEXT_FONT_SIZE).setBold(true);
  currentTextYPos += TEXT_LINE_HEIGHT + 5;

  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, TEXT_LINE_HEIGHT)
    .getText().setText(`• Idade média: ${data.idade_media} anos`)
    .getTextStyle().setFontSize(TEXT_FONT_SIZE);
  currentTextYPos += TEXT_LINE_HEIGHT;
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, TEXT_LINE_HEIGHT)
    .getText().setText(`• Frequência de uso: ${data.freq_uso_media}`)
    .getTextStyle().setFontSize(TEXT_FONT_SIZE);
  currentTextYPos += TEXT_LINE_HEIGHT;
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, TEXT_LINE_HEIGHT)
    .getText().setText(`• Gasto médio: R$ ${data.gasto_medio}`)
    .getTextStyle().setFontSize(TEXT_FONT_SIZE);
  currentTextYPos += TEXT_LINE_HEIGHT;
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, TEXT_LINE_HEIGHT)
    .getText().setText(`• Ligações call center: ${data.ligacoes_media}`)
    .getTextStyle().setFontSize(TEXT_FONT_SIZE);
  currentTextYPos += TEXT_LINE_HEIGHT + 10;

  if (data.imagem_base64) {
    const imgWidth = 300;
    const imgHeight = 250;
    const imgLeft = 420; 
    const imgTop = 100; 

    addImageFromBase64(slide, data.imagem_base64, imgLeft, imgTop, imgWidth, imgHeight);
  }

  console.log('Slide 7 preenchido.');
}

/**
 * Preenche o slide 8 com a análise de associações e o gráfico.
 * @param {object} data Os dados das associações, incluindo resultados de testes e imagem base64.
 */
function preencherSlide8(data) {
  const presentation = SlidesApp.openById(PRESENTATION_ID);
  const slide = presentation.getSlideById(SLIDE_8_ID);
  clearSlideContent(slide);

  if (data.erro) {
    slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH, TEXT_LINE_HEIGHT * 3)
      .getText().setText('Erro na análise de associações: ' + data.erro)
      .getTextStyle().setFontSize(TEXT_FONT_SIZE).setForegroundColor('#FF0000');
    return;
  }

  let yPos = MARGIN_TOP;
  // Título.
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, yPos, CONTENT_WIDTH, TITLE_FONT_SIZE + 10)
    .getText().setText('Associações com Cancelamento')
    .getTextStyle().setFontSize(TITLE_FONT_SIZE).setBold(true).setForegroundColor(TITLE_COLOR);
  yPos += TITLE_FONT_SIZE + 20; 

  // Proporções para texto e imagem.
  const textColumnWidth = CONTENT_WIDTH * 0.35;
  const imageColumnWidth = CONTENT_WIDTH * 0.60;
  const spaceBetweenColumns = CONTENT_WIDTH - textColumnWidth - imageColumnWidth;

  let currentTextYPos = yPos;

  // Conteúdo de texto.
  if (data.testes && data.testes.length > 0) {
    slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, TEXT_LINE_HEIGHT)
      .getText().setText('Testes qui-quadrado:')
      .getTextStyle().setFontSize(TEXT_FONT_SIZE).setBold(true);
    currentTextYPos += TEXT_LINE_HEIGHT + 5;

    data.testes.forEach(teste => {
      slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, TEXT_LINE_HEIGHT)
        .getText().setText(`• ${toTitleCase(teste.variavel)}: ${teste.resultado}`)
        .getTextStyle().setFontSize(TEXT_FONT_SIZE);
      currentTextYPos += TEXT_LINE_HEIGHT;
      slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, TEXT_LINE_HEIGHT)
        .getText().setText(`   p-valor: ${teste.p_valor}`)
        .getTextStyle().setFontSize(TEXT_FONT_SIZE);
      currentTextYPos += TEXT_LINE_HEIGHT;
    });
    currentTextYPos += 10;
  } else {
    slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, TEXT_LINE_HEIGHT * 2)
      .getText().setText('Nenhum teste de associação significativa encontrado ou dados insuficientes.')
      .getTextStyle().setFontSize(TEXT_FONT_SIZE);
    currentTextYPos += TEXT_LINE_HEIGHT * 2 + 10;
  }

  if (data.imagem_base64) {
    const imgWidth = 300;
    const imgHeight = 280;
    const imgLeft = 420; 
    const imgTop = 100; 

    addImageFromBase64(slide, data.imagem_base64, imgLeft, imgTop, imgWidth, imgHeight);
  }

  console.log('Slide 8 preenchido.');
}

/**
 * Preenche o slide 9 com os resultados do modelo de regressão logística e a matriz de confusão.
 * @param {object} data Os dados do modelo, incluindo métricas e imagem base64 da matriz.
 */
function preencherSlide9(data) {
  const presentation = SlidesApp.openById(PRESENTATION_ID);
  const slide = presentation.getSlideById(SLIDE_9_ID);
  clearSlideContent(slide);

  if (data.erro) {
    slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH, TEXT_LINE_HEIGHT * 3)
      .getText().setText('Erro ao construir modelo: ' + data.erro)
      .getTextStyle().setFontSize(TEXT_FONT_SIZE).setForegroundColor('#FF0000');
    return;
  }

  let yPos = MARGIN_TOP;
  // Título.
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, yPos, CONTENT_WIDTH, TITLE_FONT_SIZE + 10)
    .getText().setText('Performance do Modelo Preditivo')
    .getTextStyle().setFontSize(TITLE_FONT_SIZE).setBold(true).setForegroundColor(TITLE_COLOR);
  yPos += TITLE_FONT_SIZE + 20; 

  // Área para texto à esquerda e imagem à direita.
  const textColumnWidth = CONTENT_WIDTH * 0.40;
  const imageColumnWidth = CONTENT_WIDTH * 0.55;
  const spaceBetweenColumns = CONTENT_WIDTH - textColumnWidth - imageColumnWidth;

  let currentTextYPos = yPos;

  // Texto das métricas.
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, TEXT_LINE_HEIGHT)
    .getText().setText('Métricas de Performance:')
    .getTextStyle().setFontSize(TEXT_FONT_SIZE).setBold(true);
  currentTextYPos += TEXT_LINE_HEIGHT + 5;

  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, TEXT_LINE_HEIGHT)
    .getText().setText(`• Acurácia: ${data.acuracia}`)
    .getTextStyle().setFontSize(TEXT_FONT_SIZE);
  currentTextYPos += TEXT_LINE_HEIGHT;
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, TEXT_LINE_HEIGHT)
    .getText().setText(`• Precisão: ${data.precisao}`)
    .getTextStyle().setFontSize(TEXT_FONT_SIZE);
  currentTextYPos += TEXT_LINE_HEIGHT;
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, TEXT_LINE_HEIGHT)
    .getText().setText(`• Recall: ${data.recall}`)
    .getTextStyle().setFontSize(TEXT_FONT_SIZE);
  currentTextYPos += TEXT_LINE_HEIGHT;
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, TEXT_LINE_HEIGHT)
    .getText().setText(`• F1-Score: ${data.f1_score}`)
    .getTextStyle().setFontSize(TEXT_FONT_SIZE);
  currentTextYPos += TEXT_LINE_HEIGHT + 10;

  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, TEXT_LINE_HEIGHT * 2)
    .getText().setText(`Modelo prevê cancelamentos com ${data.acuracia} de acurácia.`)
    .getTextStyle().setFontSize(TEXT_FONT_SIZE);
  currentTextYPos += TEXT_LINE_HEIGHT * 2;


  // Imagem da Matriz de Confusão.
  if (data.matriz_confusao_base64) {
    const imgWidth = 300;
    const imgHeight = 280;
    const imgLeft = 420; 
    const imgTop = 100; 

    addImageFromBase64(slide, data.matriz_confusao_base64, imgLeft, imgTop, imgWidth, imgHeight);
  }

  console.log('Slide 9 preenchido.');
}

/**
 * Preenche o slide 10 com a análise dos fatores de risco.
 * @param {object} data Os dados dos fatores de risco, incluindo top fatores e gráfico.
 */
function preencherSlide10(data) {
  const presentation = SlidesApp.openById(PRESENTATION_ID);
  const slide = presentation.getSlideById(SLIDE_10_ID);
  clearSlideContent(slide);

  if (data.erro) {
    slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH, TEXT_LINE_HEIGHT * 3)
      .getText().setText('Erro ao analisar fatores de risco: ' + data.erro)
      .getTextStyle().setFontSize(TEXT_FONT_SIZE).setForegroundColor('#FF0000');
    return;
  }

  let yPos = MARGIN_TOP;
  // Título.
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, yPos, CONTENT_WIDTH, TITLE_FONT_SIZE + 10)
    .getText().setText('Principais Fatores de Risco')
    .getTextStyle().setFontSize(TITLE_FONT_SIZE).setBold(true).setForegroundColor(TITLE_COLOR);
  yPos += TITLE_FONT_SIZE + 20; 

  const textColumnWidth = CONTENT_WIDTH * 0.40;
  const imageColumnWidth = CONTENT_WIDTH * 0.55;
  const spaceBetweenColumns = CONTENT_WIDTH - textColumnWidth - imageColumnWidth;

  let currentTextYPos = yPos;

  // Conteúdo de texto.
  if (data.top_fatores && data.top_fatores.length > 0) {
    slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, TEXT_LINE_HEIGHT)
      .getText().setText('Top 5 fatores de influência:')
      .getTextStyle().setFontSize(TEXT_FONT_SIZE).setBold(true);
    currentTextYPos += TEXT_LINE_HEIGHT + 5;
    data.top_fatores.forEach((fator, index) => {
      const impacto = fator.coeficiente > 0 ? 'AUMENTA o risco' : 'DIMINUI o risco';
      slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, TEXT_LINE_HEIGHT)
        .getText().setText(`${index + 1}. ${toTitleCase(fator.variavel)}: ${impacto}`)
        .getTextStyle().setFontSize(TEXT_FONT_SIZE);
      currentTextYPos += TEXT_LINE_HEIGHT;
    });
    currentTextYPos += 10;
  } else {
    slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, TEXT_LINE_HEIGHT * 2)
      .getText().setText('Nenhum fator de risco significativo encontrado.')
      .getTextStyle().setFontSize(TEXT_FONT_SIZE);
    currentTextYPos += TEXT_LINE_HEIGHT * 2 + 10;
  }

  if (data.grafico_fatores_base64) {
    const imgWidth = 300;
    const imgHeight = 280;
    const imgLeft = 420; 
    const imgTop = 100; 

    addImageFromBase64(slide, data.grafico_fatores_base64, imgLeft, imgTop, imgWidth, imgHeight);
  }

  console.log('Slide 10 preenchido.');
}

/**
 * Preenche o slide 11 com a análise de impacto do call center.
 * @param {object} data Os dados do impacto do call center, incluindo insights e gráfico.
 */
function preencherSlide11(data) {
  const presentation = SlidesApp.openById(PRESENTATION_ID);
  const slide = presentation.getSlideById(SLIDE_11_ID);
  clearSlideContent(slide);

  if (data.erro) {
    slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH, TEXT_LINE_HEIGHT * 3)
      .getText().setText('Erro na análise de impacto do Call Center: ' + data.erro)
      .getTextStyle().setFontSize(TEXT_FONT_SIZE).setForegroundColor('#FF0000');
    return;
  }

  let yPos = MARGIN_TOP;
  // Título.
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, yPos, CONTENT_WIDTH, TITLE_FONT_SIZE + 10)
    .getText().setText('Análise: Call Center e Cancelamento') 
    .getTextStyle().setFontSize(TITLE_FONT_SIZE).setBold(true).setForegroundColor(TITLE_COLOR);
  yPos += TITLE_FONT_SIZE + 20; 

  const textColumnWidth = CONTENT_WIDTH * 0.45;
  const imageColumnWidth = CONTENT_WIDTH * 0.50;
  const spaceBetweenColumns = CONTENT_WIDTH - textColumnWidth - imageColumnWidth;

  let currentTextYPos = yPos;

  // Insights de texto.
  if (data.insights_text && data.insights_text.length > 0) {
    data.insights_text.forEach(insight => {
      const cleanedInsight = insight.startsWith('• ') ? insight.substring(2) : insight;
      slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, TEXT_LINE_HEIGHT * 3)
        .getText().setText(`• ${cleanedInsight}`)
        .getTextStyle().setFontSize(TEXT_FONT_SIZE);
      currentTextYPos += TEXT_LINE_HEIGHT * 3;
    });
  }
  currentTextYPos += 10;

  if (data.imagem_base64) {
    const imgWidth = 300;
    const imgHeight = 280;
    const imgLeft = 420; 
    const imgTop = 100; 

    addImageFromBase64(slide, data.imagem_base64, imgLeft, imgTop, imgWidth, imgHeight);
  }

  console.log('Slide 11 preenchido.');
}

/**
 * Preenche o slide 12 com os insights de negócio e segmentação de clientes.
 * @param {object} data Os dados dos insights, incluindo contagens de risco, recomendações e gráfico.
 */
function preencherSlide12(data) {
  const presentation = SlidesApp.openById(PRESENTATION_ID);
  const slide = presentation.getSlideById(SLIDE_12_ID);
  clearSlideContent(slide);

  if (data.erro) {
    slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH, TEXT_LINE_HEIGHT * 3)
      .getText().setText('Erro ao gerar insights: ' + data.erro)
      .getTextStyle().setFontSize(TEXT_FONT_SIZE).setForegroundColor('#FF0000');
    return;
  }

  let yPos = MARGIN_TOP;
  // Título.
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, yPos, CONTENT_WIDTH, TITLE_FONT_SIZE + 10)
    .getText().setText('Insights e Resumo Final')
    .getTextStyle().setFontSize(TITLE_FONT_SIZE).setBold(true).setForegroundColor(TITLE_COLOR);
  yPos += TITLE_FONT_SIZE + 20;

  const textColumnWidth = CONTENT_WIDTH * 0.40;
  const imageColumnWidth = CONTENT_WIDTH * 0.55;
  const spaceBetweenColumns = CONTENT_WIDTH - textColumnWidth - imageColumnWidth;

  let currentTextYPos = yPos;

  // Conteúdo de texto - Segmentação de Clientes.
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, SLIDE12_TEXT_LINE_HEIGHT)
    .getText().setText('Segmentação de Clientes:')
    .getTextStyle().setFontSize(SLIDE12_TEXT_FONT_SIZE).setBold(true);
  currentTextYPos += SLIDE12_TEXT_LINE_HEIGHT + 5;

  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, SLIDE12_TEXT_LINE_HEIGHT)
    .getText().setText(`• Alto Risco: ${data.alto_risco} clientes`)
    .getTextStyle().setFontSize(SLIDE12_TEXT_FONT_SIZE);
  currentTextYPos += SLIDE12_TEXT_LINE_HEIGHT;
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, SLIDE12_TEXT_LINE_HEIGHT)
    .getText().setText(`• Médio Risco: ${data.medio_risco} clientes`)
    .getTextStyle().setFontSize(SLIDE12_TEXT_FONT_SIZE);
  currentTextYPos += SLIDE12_TEXT_LINE_HEIGHT;
  slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, SLIDE12_TEXT_LINE_HEIGHT)
    .getText().setText(`• Baixo Risco: ${data.baixo_risco} clientes`)
    .getTextStyle().setFontSize(SLIDE12_TEXT_FONT_SIZE);
  currentTextYPos += SLIDE12_TEXT_LINE_HEIGHT + 10;

  // Recomendações.
  if (data.recomendacoes && data.recomendacoes.length > 0) {
    slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, SLIDE12_TEXT_LINE_HEIGHT)
      .getText().setText('Recomendações:')
      .getTextStyle().setFontSize(SLIDE12_TEXT_FONT_SIZE).setBold(true);
    currentTextYPos += SLIDE12_TEXT_LINE_HEIGHT + 5;
    data.recomendacoes.forEach(rec => {
      const cleanedRec = rec.startsWith('• ') ? rec.substring(2) : rec;
      slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, SLIDE12_TEXT_LINE_HEIGHT * 2)
        .getText().setText(`• ${cleanedRec}`)
        .getTextStyle().setFontSize(SLIDE12_TEXT_FONT_SIZE);
      currentTextYPos += SLIDE12_TEXT_LINE_HEIGHT * 2;
    });
    currentTextYPos += 10;
  }

  // Fatores Críticos.
  if (data.fatores_criticos && data.fatores_criticos.length > 0) {
    slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, SLIDE12_TEXT_LINE_HEIGHT)
      .getText().setText('Fatores Críticos:')
      .getTextStyle().setFontSize(SLIDE12_TEXT_FONT_SIZE).setBold(true);
    currentTextYPos += SLIDE12_TEXT_LINE_HEIGHT + 5;
    data.fatores_criticos.forEach(fator => {
      const cleanedFator = fator.startsWith('• ') ? fator.substring(2) : fator;
      slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, MARGIN_LEFT, currentTextYPos, textColumnWidth, SLIDE12_TEXT_LINE_HEIGHT)
        .getText().setText(`• ${cleanedFator}`)
        .getTextStyle().setFontSize(SLIDE12_TEXT_FONT_SIZE);
      currentTextYPos += SLIDE12_TEXT_LINE_HEIGHT;
    });
  }

  // Imagem.
  if (data.segmentacao_base64) {
    const imgWidth = 300;
    const imgHeight = 280;
    const imgLeft = 420; 
    const imgTop = 100; 

    addImageFromBase64(slide, data.segmentacao_base64, imgLeft, imgTop, imgWidth, imgHeight);
  }

  console.log('Slide 12 preenchido.');
}

/**
 * Exibe um modal customizado com título, mensagem e cor de texto.
 * @param {string} title O título do modal.
 * @param {string} message A mensagem a ser exibida no modal.
 * @param {string} textColor A cor do texto da mensagem (ex: '#008000' para verde, '#FF0000' para vermelho).
 */
function exibirModalCustomizado(title, message, textColor) {
  try {
    const htmlOutput = HtmlService.createHtmlOutput(`
      <div style="text-align: center; font-family: Arial, sans-serif; padding: 20px;">
        <h2 style="color: #1f4e79; margin-bottom: 20px;">${title}</h2>
        <p style="color: ${textColor}; font-weight: bold; margin-bottom: 30px; white-space: pre-wrap;">${message}</p>
        <button onclick="google.script.host.close()"
                       style="background-color: #4285f4; color: white; border: none; padding: 10px 20px;
                              border-radius: 4px; cursor: pointer; font-size: 14px;">
          OK
        </button>
      </div>
    `).setWidth(400).setHeight(250);

    SlidesApp.getUi().showModalDialog(htmlOutput, 'Informação da Análise');

  } catch (e) {
    console.error('Erro ao exibir modal customizado: ' + e.message);
    SlidesApp.getUi().alert(`${title}\n\n${message}`);
  }
}

/**
 * Exibe o modal de conclusão da análise, similar ao da imagem.
 */
function exibirModalConclusao() {
  try {
    const htmlOutput = HtmlService.createHtmlOutput(`
      <div style="text-align: center; font-family: Arial, sans-serif; padding: 20px;">
        <h2 style="color: #1f4e79; margin-bottom: 20px;">Status da Análise: COMPLETED</h2>
        <p style="color: #666; margin-bottom: 20px;">Última execução: ${new Date().toLocaleString('pt-BR')}</p>
        <div style="color: #008000; font-weight: bold; margin-bottom: 30px;">
          ✅ Análise concluída com sucesso!
        </div>
        <button onclick="google.script.host.close()"
                               style="background-color: #4285f4; color: white; border: none; padding: 10px 20px;
                                      border-radius: 4px; cursor: pointer; font-size: 14px;">
          OK
        </button>
      </div>
    `).setWidth(400).setHeight(250);

    SlidesApp.getUi().showModalDialog(htmlOutput, 'Script em execução');

  } catch (e) {
    console.error('Erro ao exibir modal de conclusão: ' + e.message);
  }
}
