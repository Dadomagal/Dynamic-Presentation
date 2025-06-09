# 🚀 Apresentação Dinâmica de Análise Estatística (Google Slides + Cloud Function)

Este repositório demonstra uma **apresentação dinâmica e interativa** construída no Google Slides, capaz de executar análises estatísticas em tempo real e preencher slides com resultados e gráficos gerados por uma Cloud Function. Este projeto foi um desafio pessoal para **consolidar conhecimentos em análise estatística** e explorar a integração de programação com ferramentas de apresentação.

### **Visão Geral do Projeto**

A iniciativa surgiu do desejo de ir além de uma apresentação estática. O objetivo foi criar uma experiência onde a análise de dados não apenas é exibida, mas também pode ser "ativada" e "visualizada" na hora, tornando o processo de apresentação mais engajador e interativo. É uma prova de conceito sobre como a automação e a computação em nuvem podem transformar a forma como compartilhamos insights de dados.

### **Como Funciona**

O projeto é dividido em três componentes principais:

1.  **Google Slides**: A plataforma da apresentação, onde os slides são criados e os resultados da análise são exibidos. Contém um script Google Apps Script embutido.

2.  **Google Apps Script (JavaScript)**: Embutido no Google Slides, este script atua como o "cérebro" da apresentação. Ele:
    * Cria um menu personalizado no Slides para controle da análise.
    * Gerencia o estado da análise (em execução, concluída, erro).
    * Limpa slides para novas execuções.
    * **Chama a Cloud Function** para iniciar a análise de dados.
    * Recebe os resultados (texto e imagens Base64) da Cloud Function.
    * **Preenche os slides dinamicamente** com os insights e gráficos.

3.  **Google Cloud Function (Python)**: Um serviço serverless que hospeda a lógica de análise estatística. Ele:
    * Carrega e pré-processa os dados (ou gera dados de exemplo se o carregamento falhar).
    * Realiza análises exploratórias, testes estatísticos (Qui-Quadrado) e constrói um modelo de Regressão Logística para prever cancelamentos.
    * Gera gráficos (Matriz de Confusão, Fatores de Risco, Segmentação de Clientes) como imagens Base64.
    * Retorna todos os resultados e imagens em um formato JSON para o Google Apps Script.

### **Por Que Construir Isso?**

A escolha de construir esta apresentação dinâmica foi impulsionada por dois pilares:

* **Consolidação de Conhecimento**: Reforçar a compreensão de todo o ciclo de vida de um projeto de análise de dados, desde o pré-processamento até a interpretação e comunicação dos resultados.
* **Desafio de Interatividade e Programação**: Aprofundar-me na integração de diferentes plataformas (Google Slides, Google Apps Script, Google Cloud Function) e no uso de programação para criar uma experiência de usuário inovadora em uma apresentação. É um exemplo de como a programação pode adicionar uma camada extra de dinamismo e profissionalismo.

### **Tecnologias Envolvidas**

* **Google Slides**
* **Google Apps Script (JavaScript)**
* **Google Cloud Functions (Python)**
* **Python Libraries**: Pandas, NumPy, Matplotlib, Seaborn, Scikit-learn, SciPy.
* **Google Cloud Storage**: Para armazenar o dataset (opcional, pode usar dados de exemplo).

### **Como Replicar/Usar o Projeto**

1.  **Configurar a Cloud Function (Python)**:
    * Faça o deploy do código Python (`analise_cancelamentos_cloud_function.py` ou similar) como uma Google Cloud Function.
    * Certifique-se de configurar as permissões corretas para a Cloud Function acessar o Google Cloud Storage (se estiver usando um dataset real).
    * Anote a **URL de invocação** da sua Cloud Function.
2.  **Configurar o Google Slides (Apps Script)**:
    * Abra sua apresentação do Google Slides (ou crie uma nova).
    * Vá em `Extensões > Apps Script`.
    * Copie e cole o código JavaScript (`apresentacao_dinamica_apps_script.js` ou similar) no editor do Apps Script.
    * **Atualize os IDs** nas configurações globais do script JavaScript:
        * `PRESENTATION_ID`: O ID da sua apresentação (na URL do Slides).
        * `CLOUD_FUNCTION_URL`: A URL da Cloud Function que você implantou.
        * `LOADING_MESSAGE_SLIDE_ID`: O ID do slide onde a mensagem de carregamento será exibida (geralmente o primeiro).
        * `SLIDE_X_ID` (todos os slides de destino): Os IDs dos slides onde os resultados serão preenchidos.
    * Salve o projeto Apps Script.
    * Execute a função `onOpen` manualmente uma vez (`Executar > onOpen` no editor do Apps Script) para autorizar o script e criar o menu.
3.  **Interagir com a Apresentação**:
    * Ao abrir a apresentação, o script limpa os slides de destino e exibe uma mensagem de "Carregando".
    * Após um pequeno atraso (configurável), a análise é iniciada automaticamente.
    * Alternativamente, use o menu personalizado `🔧 Análises > 🚀 Executar Análise` para iniciar manualmente.
    * O menu também oferece opções para `🧹 Limpar Dados` e `📊 Status da Análise`.

### **Resultados e Demonstrativo**

A apresentação demonstrará visualmente:

* Estatísticas exploratórias do dataset.
* As associações entre variáveis e o cancelamento.
* A performance do modelo preditivo através da matriz de confusão.
* Os principais fatores de risco para o cancelamento de clientes.
* Insights detalhados sobre o impacto do call center.
* Uma segmentação clara dos clientes por nível de risco, com recomendações de negócio.


---

<br>

---

# 🚀 Dynamic Presentation of Statistical Analysis (Google Slides + Cloud Function)

This repository showcases a **dynamic and interactive presentation** built on Google Slides, capable of executing statistical analyses in real-time and populating slides with results and charts generated by a Cloud Function. This project was a personal challenge to **consolidate knowledge in statistical analysis** and explore the integration of programming with presentation tools.

### **Project Overview**

The initiative stemmed from a desire to move beyond a static presentation. The goal was to create an experience where data analysis is not just displayed, but can also be "activated" and "visualized" on the fly, making the presentation process more engaging and interactive. It's a proof of concept on how automation and cloud computing can transform how we share data insights.

### **How It Works**

The project is divided into three main components:

1.  **Google Slides**: The presentation platform, where slides are created and analysis results are displayed. It contains an embedded Google Apps Script.

2.  **Google Apps Script (JavaScript)**: Embedded within Google Slides, this script acts as the "brain" of the presentation. It:
    * Creates a custom menu in Slides for analysis control.
    * Manages the analysis state (running, completed, error).
    * Clears slides for new executions.
    * **Calls the Cloud Function** to initiate data analysis.
    * Receives results (text and Base64 images) from the Cloud Function.
    * **Dynamically populates the slides** with insights and charts.

3.  **Google Cloud Function (Python)**: A serverless service that hosts the statistical analysis logic. It:
    * Loads and preprocesses data (or generates sample data if loading fails).
    * Performs exploratory analyses, statistical tests (Chi-Square), and builds a Logistic Regression model to predict churn.
    * Generates charts (Confusion Matrix, Risk Factors, Customer Segmentation) as Base64 images.
    * Returns all results and images in a JSON format to the Google Apps Script.

### **Why Build This?**

The decision to build this dynamic presentation was driven by two core pillars:

* **Knowledge Consolidation**: To reinforce the understanding of the entire data analysis project lifecycle, from preprocessing to interpretation and communication of results.
* **Interactivity and Programming Challenge**: To delve deeper into the integration of different platforms (Google Slides, Google Apps Script, Google Cloud Function) and the use of programming to create an innovative user experience in a presentation. It serves as an example of how programming can add an extra layer of dynamism and professionalism.

### **Technologies Involved**

* **Google Slides**
* **Google Apps Script (JavaScript)**
* **Google Cloud Functions (Python)**
* **Python Libraries**: Pandas, NumPy, Matplotlib, Seaborn, Scikit-learn, SciPy.
* **Google Cloud Storage**: For storing the dataset (optional, can use sample data).

### **How to Replicate/Use the Project**

1.  **Configure the Cloud Function (Python)**:
    * Deploy the Python code (`analise_cancelamentos_cloud_function.py` or similar) as a Google Cloud Function.
    * Ensure proper permissions are configured for the Cloud Function to access Google Cloud Storage (if using a real dataset).
    * Note down the **invocation URL** of your Cloud Function.
2.  **Configure Google Slides (Apps Script)**:
    * Open your Google Slides presentation (or create a new one).
    * Go to `Extensions > Apps Script`.
    * Copy and paste the JavaScript code (`apresentacao_dinamica_apps_script.js` or similar) into the Apps Script editor.
    * **Update the IDs** in the global settings of the JavaScript script:
        * `PRESENTATION_ID`: The ID of your presentation (from the Slides URL).
        * `CLOUD_FUNCTION_URL`: The URL of the Cloud Function you deployed.
        * `LOADING_MESSAGE_SLIDE_ID`: The ID of the slide where the loading message will be displayed (usually the first one).
        * `SLIDE_X_ID` (all target slides): The IDs of the slides where results will be populated.
    * Save the Apps Script project.
    * Run the `onOpen` function manually once (`Run > onOpen` in the Apps Script editor) to authorize the script and create the menu.
3.  **Interact with the Presentation**:
    * When opening the presentation, the script clears the target slides and displays a "Loading" message.
    * After a short (configurable) delay, the analysis starts automatically.
    * Alternatively, use the custom menu `🔧 Análises > 🚀 Executar Análise` to start manually.
    * The menu also offers options to `🧹 Limpar Dados` and `📊 Status da Análise`.

### **Results and Demonstrative Output**

The presentation will visually demonstrate:

* Exploratory statistics of the dataset.
* Associations between variables and churn.
* The performance of the predictive model through the confusion matrix.
* Key risk factors for customer churn.
* Detailed insights on call center impact.
* A clear segmentation of customers by risk level, with business recommendations.
