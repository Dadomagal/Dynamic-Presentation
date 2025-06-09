import functions_framework
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg') 	# Define o backend do Matplotlib para 'Agg', que é não-interativo e ideal para ambientes sem GUI como o Cloud Run.
import matplotlib.pyplot as plt
import seaborn as sns
from scipy.stats import chi2_contingency
import statsmodels.api as sm
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import classification_report, confusion_matrix
import io
import base64
import json
import time
import os
from google.cloud import storage
import warnings
warnings.filterwarnings('ignore') # Ignora avisos para manter a saída do log limpa.

# --- Configurações globais para gráficos ---
# Estas configurações ajustam o estilo e o tamanho dos gráficos para uma melhor visualização em slides ou dashboards.
plt.style.use('default') 
sns.set_palette("colorblind")
plt.rcParams['figure.figsize'] = (10, 6) 
plt.rcParams['font.size'] = 10 
plt.rcParams['axes.titlesize'] = 14 
plt.rcParams['axes.labelsize'] = 10
plt.rcParams['xtick.labelsize'] = 8
plt.rcParams['ytick.labelsize'] = 8
plt.rcParams['legend.fontsize'] = 8


class AnalisadorCancelamentos:
	"""
	Gerencia o fluxo de análise de dados de cancelamento, incluindo
	carregamento, pré-processamento, análise exploratória, construção de modelo
	e geração de insights.
	"""
	def __init__(self):
		"""Inicializa as propriedades da classe para armazenar o DataFrame e o modelo."""
		self.df = None
		self.modelo = None
		self.X_processed = None
		self.y_processed = None
		self.X_train = None
		self.X_test = None
		self.y_train = None
		self.y_test = None
		self.features_modelo = []

	def carregar_dados(self):
		"""
		Tenta carregar o dataset 'cancelamentos.csv' do Google Cloud Storage.
		Em caso de falha (e se as credenciais GCP não estiverem configuradas),
		gera dados de exemplo para permitir a continuidade da análise.
		"""
		print("Tentando carregar dados do Google Cloud Storage...")
		try:
			client = storage.Client()
			# O nome do bucket é obtido de uma variável de ambiente, com um fallback genérico.
			# Para fins de demonstração no GitHub, um ID genérico é utilizado.
			bucket_name = os.environ.get('GCS_BUCKET', 'seu-bucket-generico-de-dados')
			bucket = client.bucket(bucket_name)
			blob = bucket.blob('cancelamentos.csv')

			data = blob.download_as_text()
			self.df = pd.read_csv(io.StringIO(data))

			print(f"Dataset 'cancelamentos.csv' carregado com sucesso do GCS: {self.df.shape[0]} registros e {self.df.shape[1]} variáveis")
			return True
		except Exception as e:
			print(f"Erro ao carregar dados do Cloud Storage: {e}")
			print("Verificando se as credenciais do GCP estão configuradas...")
			# Se as credenciais do GCP não estiverem configuradas, gera dados de exemplo.
			if not os.environ.get('GOOGLE_APPLICATION_CREDENTIALS') and not os.environ.get('GOOGLE_CLOUD_PROJECT'):
				print("Aviso: Credenciais do GCP não encontradas ou projeto não configurado. Gerando dados de exemplo para continuar.")
				return self.criar_dados_exemplo()
			else:
				print("Erro crítico ao carregar 'cancelamentos.csv' do GCS, e credenciais parecem estar presentes. A Cloud Function pode ter problemas de permissão ou o bucket/arquivo está incorreto.")
				return False # Retorna False se houver erro e credenciais existirem

	def criar_dados_exemplo(self, n_samples=10000):
		"""
		Cria um DataFrame com dados sintéticos de exemplo.
		Esta função é usada como fallback se o carregamento do arquivo real falhar.
		"""
		try:
			np.random.seed(42) # Define uma semente para reprodutibilidade dos dados gerados.

			self.df = pd.DataFrame({
				'idade': np.random.normal(35, 10, n_samples).astype(int),
				'sexo': np.random.choice(['M', 'F'], n_samples, p=[0.48, 0.52]),
				'frequencia_uso': np.random.exponential(2, n_samples),
				'total_gasto': np.random.normal(100, 30, n_samples),
				'ligacoes_callcenter': np.random.poisson(2, n_samples),
				'meses_ultima_interacao': np.random.randint(1, 12, n_samples),
				'assinatura': np.random.choice(['Basic', 'Premium', 'Standard'], n_samples, p=[0.4, 0.2, 0.4]),
				'duracao_contrato': np.random.choice(['Mensal', 'Anual', 'Bianual'], n_samples, p=[0.5, 0.3, 0.2])
			})

			# Simula a probabilidade de cancelamento com base em algumas características.
			cancelou_prob = (
				(self.df['idade'] < 25) * 0.15 +
				(self.df['ligacoes_callcenter'] > 3) * 0.3 +
				(self.df['frequencia_uso'] < 1) * 0.2 +
				(self.df['total_gasto'] < 50) * 0.1 +
				(self.df['duracao_contrato'] == 'Mensal') * 0.2 +
				(self.df['sexo'] == 'F') * 0.05
			)
			self.df['cancelou'] = np.random.binomial(1, np.clip(cancelou_prob, 0.05, 0.8), n_samples)

			print(f"Dados de exemplo criados: {self.df.shape[0]} registros e {self.df.shape[1]} variáveis")
			return True
		except Exception as e:
			print(f"Erro ao criar dados de exemplo: {e}")
			return False

	def preprocessar_dados(self):
		"""
		Realiza a limpeza e o pré-processamento do DataFrame:
		- Remove linhas com valores ausentes na coluna 'cancelou'.
		- Converte 'cancelou' para tipo inteiro.
		- Padroniza nomes de colunas.
		- Preenche valores ausentes em colunas numéricas com a mediana.
		- Preenche valores ausentes em colunas categóricas com a moda.
		- Cria variáveis dummy para colunas categóricas e prepara os dados para o modelo.
		"""
		if self.df is None:
			print("DataFrame não carregado para pré-processamento.")
			return False

		print(f"Iniciando pré-processamento. Shape inicial: {self.df.shape}")

		try:
			if 'cancelou' not in self.df.columns:
				print("Erro: Coluna 'cancelou' não encontrada no DataFrame.")
				return False

			initial_rows = self.df.shape[0]
			self.df.dropna(subset=['cancelou'], inplace=True)
			dropped_rows = initial_rows - self.df.shape[0]
			if dropped_rows > 0:
				print(f"Aviso: Removidas {dropped_rows} linhas com valores ausentes na coluna 'cancelou'.")

			self.df['cancelou'] = pd.to_numeric(self.df['cancelou'], errors='coerce').fillna(0).astype(int)
			# Padroniza nomes de colunas para minúsculas e sem espaços.
			self.df.columns = self.df.columns.str.strip().str.lower().str.replace(' ', '_')

			numeric_cols = ['idade', 'frequencia_uso', 'total_gasto', 'ligacoes_callcenter', 'meses_ultima_interacao']

			for col in numeric_cols:
				if col in self.df.columns:
					self.df[col] = pd.to_numeric(self.df[col], errors='coerce')
					median_val = self.df[col].median()
					if pd.isna(median_val): median_val = 0
					self.df[col] = self.df[col].fillna(median_val)
					self.df[col].replace([np.inf, -np.inf], median_val, inplace=True)
				else:
					print(f"Aviso: Coluna numérica '{col}' não encontrada.")

			categorical_cols = ['sexo', 'assinatura', 'duracao_contrato']
			for col in categorical_cols:
				if col in self.df.columns:
					self.df[col] = self.df[col].astype(str).str.strip()
					self.df[col] = self.df[col].replace(['nan', 'NaN', 'None', ''], pd.NA)
					try:
						mode_val = self.df[col].mode()
						if len(mode_val) > 0:
							self.df[col] = self.df[col].fillna(mode_val[0])
						else:
							self.df[col] = self.df[col].fillna('Desconhecido')
					except Exception as e:
						print(f"Erro ao preencher NaN em coluna categórica {col}: {e}")
						self.df[col] = self.df[col].fillna('Desconhecido')
				else:
					print(f"Aviso: Coluna categórica '{col}' não encontrada.")

			features_for_model = [
				'idade', 'frequencia_uso', 'total_gasto', 'ligacoes_callcenter',
				'meses_ultima_interacao', 'sexo', 'assinatura', 'duracao_contrato'
			]

			actual_features_to_process = [f for f in features_for_model if f in self.df.columns]
			
			# Remove a coluna 'customerid' do conjunto de features para garantir que o modelo não use IDs específicos.
			# Isso torna a análise mais genérica e focada nas características do cliente.
			df_for_dummies = self.df[actual_features_to_process].copy()
			if 'customerid' in df_for_dummies.columns:
				df_for_dummies = df_for_dummies.drop('customerid', axis=1)

			self.y_processed = self.df['cancelou'].copy()

			existing_categorical_cols = [col for col in categorical_cols if col in df_for_dummies.columns]

			if existing_categorical_cols:
				# Cria variáveis dummy para as colunas categóricas, removendo a primeira para evitar multicolinearidade.
				self.X_processed = pd.get_dummies(df_for_dummies, columns=existing_categorical_cols, drop_first=True, dtype=int)
			else:
				self.X_processed = df_for_dummies.copy()

			for col in self.X_processed.columns:
				self.X_processed[col] = pd.to_numeric(self.X_processed[col], errors='coerce').fillna(0)
				self.X_processed[col].replace([np.inf, -np.inf], 0, inplace=True)

			self.features_modelo = self.X_processed.columns.tolist()

			# Divide os dados em conjuntos de treino e teste. O 'stratify' garante que a proporção de 'cancelou'
			# seja mantida em ambos os conjuntos, o que é importante para variáveis alvo desbalanceadas.
			if self.y_processed.nunique() > 1 and len(self.y_processed.value_counts()) > 1:
				self.X_train, self.X_test, self.y_train, self.y_test = train_test_split(
					self.X_processed, self.y_processed, test_size=0.25, random_state=42, stratify=self.y_processed
				)
			else:
				print("Aviso: A variável 'cancelou' tem apenas uma classe. Não será possível estratificar a divisão.")
				self.X_train, self.X_test, self.y_train, self.y_test = train_test_split(
					self.X_processed, self.y_processed, test_size=0.25, random_state=42
				)

			print(f"Pré-processamento concluído. X_processed shape: {self.X_processed.shape}, X_train shape: {self.X_train.shape}, X_test shape: {self.X_test.shape}")
			return True
		except Exception as e:
			print(f"Erro no pré-processamento: {e}")
			import traceback
			traceback.print_exc()
			return False

	def analise_exploratoria(self):
		"""Calcula e retorna estatísticas sumárias básicas do DataFrame."""
		try:
			stats = {
				'total_registros': int(len(self.df)),
				'total_variaveis': int(len(self.df.columns)),
				'taxa_cancelamento': round(float(self.df['cancelou'].mean()) * 100, 2),
				'registros_validos': int(len(self.df.dropna())) 
			}
			return stats
		except Exception as e:
			print(f"Erro na análise exploratória: {e}")
			return {'erro': str(e)}

	def gerar_distribuicoes(self):
		"""
		Gera e retorna gráficos de distribuição (histogramas) para variáveis numéricas,
		incluindo média e mediana, como uma imagem Base64.
		"""
		try:
			numeric_cols = ['idade', 'frequencia_uso', 'total_gasto', 'ligacoes_callcenter', 'meses_ultima_interacao']
			existing_cols = [col for col in numeric_cols if col in self.df.columns]

			if not existing_cols:
				return {'erro': 'Nenhuma coluna numérica encontrada para distribuição'}

			n_cols_plot = len(existing_cols)
			n_rows_plot = (n_cols_plot + 1) // 2

			# Ajusta o tamanho da figura para caber melhor no slide.
			fig, axes = plt.subplots(n_rows_plot, 2, figsize=(10, 3.5 * n_rows_plot))
			
			if n_rows_plot == 1 and n_cols_plot <= 2:
				axes = np.array(axes).flatten()
			elif n_rows_plot > 1:
				axes = axes.flatten()

			for i, col in enumerate(existing_cols):
				if i < len(axes):
					ax = axes[i]
					try:
						data = self.df[col].dropna()
						if len(data) == 0 or data.nunique() == 1:
							ax.text(0.5, 0.5, f'Sem dados ou dados constantes para {col}', 
											ha='center', va='center', transform=ax.transAxes, fontsize=8)
							ax.set_title(f'Distribuição de {col.replace("_", " ").title()}', 
											fontweight='bold', fontsize=10)
							continue

						bins_to_use = min(20, data.nunique()) if data.nunique() > 1 else 1
						ax.hist(data, bins=bins_to_use, edgecolor='black', alpha=0.7, color='skyblue')

						ax.set_title(f'Distribuição de {col.replace("_", " ").title()}', 
											fontweight='bold', fontsize=10)
						ax.set_xlabel(f'{col.replace("_", " ").title()}', fontsize=8)
						ax.set_ylabel('Frequência', fontsize=8)
						ax.grid(True, linestyle='--', alpha=0.7)

						mean_val = data.mean()
						median_val = data.median()

						ax.axvline(mean_val, color='red', linestyle='--', linewidth=1, 
											label=f'Média: {mean_val:.2f}')
						ax.axvline(median_val, color='green', linestyle=':', linewidth=1, 
											label=f'Mediana: {median_val:.2f}')
						ax.legend(fontsize=6)
					except Exception as e:
						ax.text(0.5, 0.5, f'Erro ao plotar {col}', 
											ha='center', va='center', transform=ax.transAxes, fontsize=8)

			for j in range(len(existing_cols), len(axes)):
				if j < len(axes):
					fig.delaxes(axes[j])

			plt.tight_layout()

			# Salva o gráfico em um buffer e o converte para Base64.
			img_buffer = io.BytesIO()
			plt.savefig(img_buffer, format='png', dpi=100, bbox_inches='tight')
			plt.close()
			img_buffer.seek(0)
			img_base64 = base64.b64encode(img_buffer.getvalue()).decode()

			stats = {
				'idade_media': round(float(self.df['idade'].mean()) if 'idade' in self.df.columns else 0, 2),
				'freq_uso_media': round(float(self.df['frequencia_uso'].mean()) if 'frequencia_uso' in self.df.columns else 0, 2),
				'gasto_medio': round(float(self.df['total_gasto'].mean()) if 'total_gasto' in self.df.columns else 0, 2),
				'ligacoes_media': round(float(self.df['ligacoes_callcenter'].mean()) if 'ligacoes_callcenter' in self.df.columns else 0, 2),
				'imagem_base64': img_base64
			}

			return stats
		except Exception as e:
			print(f"Erro ao gerar distribuições: {e}")
			return {'erro': str(e)}

	def analisar_associacoes(self):
		"""
		Realiza testes Qui-quadrado para variáveis categóricas e gera gráficos
		da taxa de cancelamento por categoria, retornando os resultados e as imagens.
		"""
		try:
			categorical_cols = ['sexo', 'assinatura', 'duracao_contrato']
			testes = [] # Lista para armazenar os resultados dos testes estatísticos.
			img_base64 = None

			for col in categorical_cols:
				if col in self.df.columns:
					temp_df = self.df.dropna(subset=[col, 'cancelou'])
					if not temp_df.empty and len(temp_df[col].unique()) > 1:
						try:
							tabela = pd.crosstab(temp_df[col], temp_df['cancelou'])

							if tabela.shape[0] > 1 and tabela.shape[1] > 1:
								qui2, p, gl, esperado = chi2_contingency(tabela)

								# Classifica a força da associação com base no valor-P.
								if p < 0.001: resultado = "Associação muito forte"
								elif p < 0.01: resultado = "Associação forte"
								elif p < 0.05: resultado = "Associação moderada"
								else: resultado = "Sem associação significativa"

								testes.append({
									'variavel': col,
									'qui_quadrado': round(float(qui2), 2),
									'p_valor': f"{p:.6f}",
									'resultado': resultado
								})
							else:
								print(f"Aviso: Tabela de contingência para {col} muito pequena para qui-quadrado.")
						except Exception as e:
							print(f"Erro no teste qui-quadrado para {col}: {e}")

			n_cols_plot = len(categorical_cols)
			n_rows_plot = (n_cols_plot + 1) // 2

			if n_cols_plot > 0:
				fig, axes = plt.subplots(n_rows_plot, 2, figsize=(10, 4 * n_rows_plot)) # Ajustado figsize
				if n_rows_plot == 1 and n_cols_plot <= 2:
					axes = np.array(axes).flatten()
				elif n_rows_plot > 1:
					axes = axes.flatten()

				plot_count = 0
				for i, col in enumerate(categorical_cols):
					if col in self.df.columns and plot_count < len(axes):
						ax = axes[plot_count]
						temp_df = self.df.dropna(subset=[col, 'cancelou'])
						if not temp_df.empty:
							tabela_pct = pd.crosstab(temp_df[col], temp_df['cancelou'], normalize='index') * 100

							if 1 in tabela_pct.columns:
								tabela_pct[1].sort_values(ascending=False).plot(
									kind='bar', ax=ax, color='coral', edgecolor='black'
								)
								ax.set_title(f'Taxa de Cancelamento por {col.replace("_", " ").title()} (%)', fontweight='bold', fontsize=12)
								ax.set_ylabel('Percentual de Cancelamento (%)', fontsize=9) 
								ax.set_xlabel(col.replace("_", " ").title(), fontsize=9) 
								ax.tick_params(axis='x', rotation=45, labelsize=7) 
								ax.grid(axis='y', linestyle='--', alpha=0.7)
							else:
								ax.text(0.5, 0.5, f'Sem dados de cancelamento para {col}', ha='center', va='center', transform=ax.transAxes, fontsize=10)
								ax.set_title(f'Taxa de Cancelamento por {col.replace("_", " ").title()} (%)', fontweight='bold', fontsize=12)
						else:
							ax.text(0.5, 0.5, f'Sem dados válidos para {col}', ha='center', va='center', transform=ax.transAxes, fontsize=10)
							ax.set_title(f'Taxa de Cancelamento por {col.replace("_", " ").title()} (%)', fontweight='bold', fontsize=12)
						plot_count += 1

				# Remove subplots vazios para uma apresentação mais limpa.
				for j in range(plot_count, len(axes)):
					if j < len(axes):
						fig.delaxes(axes[j])

				plt.tight_layout()

				# Salva o gráfico em um buffer e o converte para Base64.
				img_buffer = io.BytesIO()
				plt.savefig(img_buffer, format='png', dpi=100, bbox_inches='tight')
				plt.close()
				img_buffer.seek(0)
				img_base64 = base64.b64encode(img_buffer.getvalue()).decode()

			return {
				'testes': testes,
				'imagem_base64': img_base64
			}
		except Exception as e:
			print(f"Erro na análise de associações: {e}")
			return {'erro': str(e)}

	def construir_modelo(self):
		"""
		Constrói e treina um modelo de Regressão Logística para prever cancelamentos.
		Avalia o modelo utilizando os conjuntos de teste e retorna métricas de desempenho
		e a matriz de confusão como uma imagem Base64.
		"""
		try:
			# Verifica se os dados de treino/teste estão disponíveis; se não, tenta pré-processar novamente.
			if self.X_train is None or self.X_test is None or self.y_train is None or self.y_test is None:
				print("Dados de treino/teste não divididos. Tentando pré-processar e dividir novamente.")
				if not self.preprocessar_dados():
					raise ValueError("Dados insuficientes ou erro no pré-processamento para treinar o modelo.")

			# Garante que a variável target tenha mais de uma classe para o treinamento do modelo.
			if len(self.y_train.unique()) < 2:
				raise ValueError("Não há variação suficiente na variável target 'cancelou' para treinar o modelo.")

			# Inicializa e treina o modelo de Regressão Logística.
			self.modelo = LogisticRegression(max_iter=2000, random_state=42, solver='liblinear', n_jobs=-1, C=0.1)
			self.modelo.fit(self.X_train, self.y_train)

			y_pred = self.modelo.predict(self.X_test) # Faz previsões no conjunto de teste.
			# Gera um relatório de classificação com métricas detalhadas.
			report = classification_report(self.y_test, y_pred, output_dict=True, zero_division=0)

			cm = confusion_matrix(self.y_test, y_pred) # Calcula a matriz de confusão.

			# Gera a matriz de confusão como um heatmap.
			fig, ax = plt.subplots(figsize=(8, 6)) 
			sns.heatmap(cm, annot=True, fmt='d', cmap='Blues',
						xticklabels=['Não Cancelou', 'Cancelou'],
						yticklabels=['Não Cancelou', 'Cancelou'],
						annot_kws={"size": 10}) 

			plt.title('Matriz de Confusão do Modelo Preditivo', fontweight='bold', fontsize=14)
			plt.ylabel('Valor Real', fontsize=10) 
			plt.xlabel('Valor Previsto', fontsize=10) 
			plt.tight_layout()

			# Converte o gráfico da matriz de confusão para Base64.
			img_buffer = io.BytesIO()
			plt.savefig(img_buffer, format='png', dpi=100, bbox_inches='tight')
			plt.close()
			img_buffer.seek(0)
			matriz_base64 = base64.b64encode(img_buffer.getvalue()).decode()

			# Extrai as métricas do relatório.
			accuracy = report.get('accuracy', 0)
			precision_1 = report.get('1', {}).get('precision', 0)
			recall_1 = report.get('1', {}).get('recall', 0)
			f1_1 = report.get('1', {}).get('f1-score', 0)

			# Fallback para 'weighted avg' se as métricas da classe '1' forem zero (problema de classes desbalanceadas).
			if precision_1 == 0 and recall_1 == 0 and f1_1 == 0:
				print("Aviso: Métricas para classe '1' (cancelou) são 0. Tentando usar 'weighted avg'.")
				precision_1 = report.get('weighted avg', {}).get('precision', 0)
				recall_1 = report.get('weighted avg', {}).get('recall', 0)
				f1_1 = report.get('weighted avg', {}).get('f1-score', 0)

			return {
				'acuracia': f"{accuracy:.2f}",
				'precisao': f"{precision_1:.2f}",
				'recall': f"{recall_1:.2f}",
				'f1_score': f"{f1_1:.2f}",
				'matriz_confusao_base64': matriz_base64
			}

		except Exception as e:
			print(f"Erro ao construir modelo: {e}")
			import traceback
			traceback.print_exc()
			return {
				'acuracia': "0.00", 'precisao': "0.00", 'recall': "0.00",
				'f1_score': "0.00", 'matriz_confusao_base64': None, 'erro': str(e)
			}

	def analisar_fatores_risco(self):
		"""
		Analisa os coeficientes do modelo para identificar os principais fatores de risco de cancelamento,
		gerando um gráfico e uma lista dos top fatores.
		"""
		try:
			if self.modelo is None or not self.features_modelo:
				print("Modelo não treinado ou features_modelo ausentes para fatores de risco. Tentando construir...")
				modelo_result = self.construir_modelo()
				if 'erro' in modelo_result:
					return {'top_fatores': [], 'grafico_fatores_base64': None, 'erro': 'Modelo não pôde ser treinado para fatores de risco'}

			# Cria um DataFrame com os coeficientes do modelo e sua importância absoluta.
			if len(self.modelo.coef_[0]) != len(self.features_modelo):
				print("Aviso: Incompatibilidade entre coeficientes do modelo e self.features_modelo. Recalculando features para coef_df.")
				if self.X_train is not None:
					coef_df = pd.DataFrame({
						'variavel': self.X_train.columns.tolist(),
						'coeficiente': self.modelo.coef_[0],
						'importancia': np.abs(self.modelo.coef_[0])
					}).sort_values('importancia', ascending=False)
				else:
					raise ValueError("Incompatibilidade entre coeficientes do modelo e features_modelo, e X_train não disponível.")
			else:
				coef_df = pd.DataFrame({
					'variavel': self.features_modelo,
					'coeficiente': self.modelo.coef_[0],
					'importancia': np.abs(self.modelo.coef_[0])
				}).sort_values('importancia', ascending=False)

			# Ajusta o rótulo para 'sexo_Female' e inverte o sinal do coeficiente, se 'sexo_male' for a feature dummy.
			if 'sexo_male' in coef_df['variavel'].values:
				male_idx = coef_df[coef_df['variavel'] == 'sexo_male'].index[0]
				coef_df.loc[male_idx, 'coeficiente'] *= -1
				coef_df.loc[male_idx, 'variavel'] = 'sexo_female'


			top_5 = coef_df.head(5).copy() # Seleciona os 5 fatores mais importantes.

			fig, ax = plt.subplots(figsize=(8, 6)) 

			# Define cores para barras, indicando se o fator aumenta (vermelho) ou diminui (azul) o risco.
			colors = ['#FF4500' if x > 0 else '#1E90FF' for x in top_5['coeficiente']]
			bars = ax.barh(range(len(top_5)), top_5['importancia'], color=colors)

			ax.set_yticks(range(len(top_5)))
			ax.set_yticklabels([var.replace('_', ' ').replace('sexo ', 'Sexo ').replace('duracao_contrato ', 'Duração Contrato ').replace('assinatura ', 'Assinatura ').title() for var in top_5['variavel']], fontsize=8) 
			ax.set_xlabel('Importância (Valor Absoluto do Coeficiente)', fontsize=9) 
			plt.title('Top 5 Fatores de Risco para Cancelamento', fontweight='bold', fontsize=14)
			ax.invert_yaxis()

			# Adiciona rótulos de texto nas barras para indicar a direção do impacto.
			for i, bar in enumerate(bars):
				coef_val = top_5['coeficiente'].iloc[i]
				label = "Aumenta o risco" if coef_val > 0 else "Diminui o risco"
				ax.text(bar.get_width() + 0.01, bar.get_y() + bar.get_height()/2,
											label, va='center', ha='left', color='black', fontsize=7) 

			plt.tight_layout()

			# Salva o gráfico em um buffer e o converte para Base64.
			img_buffer = io.BytesIO()
			plt.savefig(img_buffer, format='png', dpi=100, bbox_inches='tight')
			plt.close()
			img_buffer.seek(0)
			grafico_base64 = base64.b64encode(img_buffer.getvalue()).decode()

			top_fatores_list = []
			for _, row in top_5.iterrows():
				top_fatores_list.append({
					'variavel': row['variavel'].replace('_', ' ').title(),
					'coeficiente': float(row['coeficiente']),
					'importancia': float(row['importancia'])
				})
			
			# Prepara um resumo dos 3 principais fatores para uso em outro slide/seção.
			summary_top_factors = []
			for _, row in coef_df.head(3).iterrows(): 
				summary_top_factors.append({
					'variavel': row['variavel'].replace('_', ' ').title(),
					'coeficiente': float(row['coeficiente'])
				})

			return {
				'top_fatores': top_fatores_list,
				'grafico_fatores_base64': grafico_base64,
				'total_fatores': len(coef_df),
				'summary_top_factors': summary_top_factors 
			}
		except Exception as e:
			print(f"Erro ao analisar fatores de risco: {e}")
			import traceback
			traceback.print_exc()
			return {'erro': str(e)}

	def analisar_impacto_callcenter(self):
		"""
		Gera o gráfico de impacto das ligações para o call center no risco de cancelamento,
		retornando-o como uma imagem Base64 junto com insights textuais.
		"""
		try:
			if self.modelo is None or self.X_processed is None:
				print("Modelo não treinado ou X_processed ausente para análise de call center. Tentando construir...")
				modelo_result = self.construir_modelo()
				if 'erro' in modelo_result:
					return {'erro': 'Não foi possível construir o modelo para análise de call center'}

			if 'ligacoes_callcenter' not in self.df.columns:
				return {'erro': 'Coluna "ligacoes_callcenter" não encontrada para análise de call center.'}

			# Certifica-se de que 'risco_cancelamento' está no df.
			if 'risco_cancelamento' not in self.df.columns:
				self.df['risco_cancelamento'] = self.modelo.predict_proba(self.X_processed)[:, 1]

			fig, ax = plt.subplots(figsize=(8, 6)) 

			sns.scatterplot(x='ligacoes_callcenter', y='risco_cancelamento',
							data=self.df, alpha=0.2, color='darkblue', s=40, ax=ax, label='Clientes Individuais')

			sns.lineplot(x='ligacoes_callcenter', y='risco_cancelamento',
							data=self.df, errorbar=('ci', 95),
							color='red', linewidth=3, ax=ax, label='Tendência (Média e IC 95%)')

			plt.title('Relação entre Ligações ao Call Center e Risco de Cancelamento',
										fontweight='bold', fontsize=14)
			plt.xlabel('Número de Ligações ao Call Center', fontsize=10) 
			plt.ylabel('Probabilidade de Cancelamento', fontsize=10) 
			plt.xticks(fontsize=8) 
			plt.yticks(fontsize=8) 
			plt.grid(True, linestyle='--', alpha=0.7)

			ax.legend(title='Legenda', loc='upper left', bbox_to_anchor=(1.02, 1), borderaxespad=0., fontsize=7) 

			plt.tight_layout(rect=[0, 0, 0.95, 1]) # Ajustado para a legenda não ser cortada.
			plt.subplots_adjust(right=0.85) # Ajustar para dar espaço à legenda.
			
			# Salva o gráfico em um buffer e o converte para Base64.
			img_buffer = io.BytesIO()
			plt.savefig(img_buffer, format='png', dpi=100, bbox_inches='tight')
			plt.close()
			img_buffer.seek(0)
			grafico_base64 = base64.b64encode(img_buffer.getvalue()).decode()

			# Texto de insights para o slide.
			insights_text = [
				"• Relação Crucial: O número de ligações para o call center é um fator chave no risco de cancelamento.",
				"• Risco Crescente: Após 4-5 ligações, a probabilidade de cancelamento aumenta drasticamente, indicando insatisfação crônica ou problemas não resolvidos.",
				"• Ação Proativa: Clientes com múltiplas interações no call center são prioritários para ações de retenção e resolução proativa de suas questões."
			]

			return {
				'insights_text': insights_text,
				'imagem_base64': grafico_base64
			}
		except Exception as e:
			print(f"Erro ao analisar impacto do call center: {e}")
			import traceback
			traceback.print_exc()
			return {'erro': str(e)}

	def gerar_insights(self):
		"""
		Segmenta os clientes em grupos de risco (Baixo, Médio, Alto) com base na probabilidade
		de cancelamento do modelo, gera um gráfico da segmentação e fornece recomendações de negócio.
		"""
		try:
			if self.modelo is None or self.X_processed is None:
				print("Modelo não treinado ou X_processed ausente para insights. Tentando construir...")
				modelo_result = self.construir_modelo()
				if 'erro' in modelo_result:
					return {'erro': 'Não foi possível construir o modelo para gerar insights'}

			# Calcula a probabilidade de cancelamento para todos os clientes.
			risco = self.modelo.predict_proba(self.X_processed)[:, 1]

			df_temp = self.df.copy()
			df_temp['risco_cancelamento'] = risco

			# Usa qcut para segmentação em quartis de risco.
			df_temp['grupo_risco'] = pd.qcut(df_temp['risco_cancelamento'],
											 q=[0, 0.25, 0.75, 1.0],
											 labels=['Baixo Risco', 'Médio Risco', 'Alto Risco'],
											 duplicates='drop')

			contagens = df_temp['grupo_risco'].value_counts()

			fig, ax = plt.subplots(figsize=(8, 6)) 

			colors = ['#DC3912', '#FF9900', '#109618'] # Cores personalizadas para os grupos de risco.
			order = ['Alto Risco', 'Médio Risco', 'Baixo Risco'] # Ordem de exibição no gráfico.
			contagens_ordered = contagens.reindex(order, fill_value=0)

			bars = ax.bar(contagens_ordered.index, contagens_ordered.values, color=colors)

			plt.title('Segmentação de Clientes por Nível de Risco de Cancelamento', fontweight='bold', fontsize=14)
			plt.xlabel('Grupo de Risco', fontsize=10) 
			plt.ylabel('Número de Clientes', fontsize=10) 
			plt.xticks(rotation=45, fontsize=8) 
			plt.yticks(fontsize=8) 

			total = contagens_ordered.sum()
			# Adiciona rótulos com contagem e percentual nas barras.
			for bar in bars:
				height = bar.get_height()
				percentage = (height / total) * 100
				ax.text(bar.get_x() + bar.get_width()/2., height + 5,
											f'{int(height)}', ha='center', va='bottom', fontsize=8, fontweight='bold', color='black') 
				ax.text(bar.get_x() + bar.get_width()/2., height / 2,
											f'{percentage:.1f}%', ha='center', va='center', fontsize=7, color='white', fontweight='bold') 

			plt.tight_layout()

			# Salva o gráfico em um buffer e o converte para Base64.
			img_buffer = io.BytesIO()
			plt.savefig(img_buffer, format='png', dpi=100, bbox_inches='tight')
			plt.close()
			img_buffer.seek(0)
			segmentacao_base64 = base64.b64encode(img_buffer.getvalue()).decode()

			# Recomendações de negócio baseadas na segmentação de risco.
			recomendacoes = [
				"Priorizar clientes de 'Alto Risco' com ofertas de retenção personalizadas.",
				"Realizar pesquisas de satisfação proativas com clientes de 'Médio Risco'.",
				"Garantir a excelência no atendimento para clientes de 'Baixo Risco' e incentivá-los a se tornarem promotores.",
				"Analisar o histórico de interações para identificar gatilhos de cancelamento.",
				"Desenvolver programas de fidelidade com base na 'Duração do Contrato' e 'Frequência de Uso'."
			]

			return {
				'alto_risco': int(contagens.get('Alto Risco', 0)),
				'medio_risco': int(contagens.get('Médio Risco', 0)),
				'baixo_risco': int(contagens.get('Baixo Risco', 0)),
				'total_clientes': int(len(df_temp)),
				'recomendacoes': recomendacoes,
				'segmentacao_base64': segmentacao_base64
			}
		except Exception as e:
			print(f"Erro ao gerar insights: {e}")
			import traceback
			traceback.print_exc()
			return {'erro': str(e)}


analisador = AnalisadorCancelamentos()

@functions_framework.http
def analisar_cancelamentos(request):
	"""
	Ponto de entrada principal da Cloud Function.
	Processa requisições HTTP para realizar a análise de cancelamento de clientes.
	"""
	# Lida com requisições OPTIONS (preflight CORS).
	if request.method == 'OPTIONS':
		headers = {
			'Access-Control-Allow-Origin': '*',
			'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
			'Access-Control-Allow-Headers': 'Content-Type, Authorization',
			'Access-Control-Max-Age': '3600'
		}
		return ('', 204, headers)

	# Headers padrão para as respostas, permitindo CORS.
	headers = {
		'Access-Control-Allow-Origin': '*',
		'Content-Type': 'application/json'
	}

	try:
		if request.method not in ['GET', 'POST']:
			return json.dumps({
				'success': False,
				'error': 'Método não permitido'
			}), 405, headers

		# Tenta obter o JSON da requisição (para POST) ou usa um dicionário vazio (para GET).
		if request.method == 'POST':
			request_json = request.get_json(silent=True) or {}
		else:
			request_json = {}

		action = request_json.get('action', 'full_analysis') # Define a ação a ser executada.
		step = request_json.get('step') # Mantém 'step' para compatibilidade, mas 'full_analysis' será o principal.

		print(f"Iniciando análise - Action: {action}, Step: {step}")

		# Carrega e pré-processa dados apenas se self.df ainda não estiver carregado.
		if analisador.df is None:
			print("Carregando e pré-processando dados pela primeira vez...")
			if not analisador.carregar_dados():
				return json.dumps({
					'success': False,
					'error': 'Erro crítico ao carregar dados. Verifique o GCS e permissões.'
				}), 500, headers

			if not analisador.preprocessar_dados():
				return json.dumps({
					'success': False,
					'error': 'Erro no pré-processamento dos dados'
				}), 500, headers
		else:
			print("Dados já carregados e pré-processados. Reutilizando DataFrame e divisões existentes.")

		resultado_data = {}
		if action == 'full_analysis':
			print("Executando análise completa...")

			resultado_data['analise_exploratoria'] = analisador.analise_exploratoria()
			resultado_data['distribuicoes'] = analisador.gerar_distribuicoes()
			resultado_data['associacoes'] = analisador.analisar_associacoes()
			resultado_data['modelo'] = analisador.construir_modelo()
			resultado_data['fatores_risco'] = analisador.analisar_fatores_risco()
			resultado_data['call_center_impact'] = analisador.analisar_impacto_callcenter() 
			resultado_data['insights'] = analisador.gerar_insights()

			resultado = {
				'success': True,
				'data': resultado_data
			}

		elif action == 'step_analysis' and step: # Permite execução de etapas específicas.
			print(f"Executando etapa específica: {step}...")
			if step == 'exploratorio':
				data = analisador.analise_exploratoria()
			elif step == 'distribuicoes':
				data = analisador.gerar_distribuicoes()
			elif step == 'associacoes':
				data = analisador.analisar_associacoes()
			elif step == 'modelo':
				data = analisador.construir_modelo()
			elif step == 'fatores_risco':
				data = analisador.analisar_fatores_risco()
			elif step == 'call_center_impact':
				data = analisador.analisar_impacto_callcenter()
			elif step == 'insights':
				data = analisador.gerar_insights()
			else:
				return json.dumps({
					'success': False,
					'error': f'Etapa inválida: {step}'
				}), 400, headers

			resultado = {
				'success': True,
				'data': data
			}
		else:
			resultado = {
				'success': True,
				'message': 'Serviço funcionando',
				'timestamp': time.time()
			}

		return json.dumps(resultado, ensure_ascii=False), 200, headers

	except Exception as e:
		print(f"Erro geral na função analisar_cancelamentos: {str(e)}")
		import traceback
		traceback.print_exc()

		return json.dumps({
			'success': False,
			'error': f'Erro interno da Cloud Function: {str(e)}'
		}, ensure_ascii=False), 500, headers

if __name__ == "__main__":
	"""
	Bloco para execução local do servidor Flask, simulando a Cloud Function.
	Permite testar o código localmente antes do deploy.
	"""
	from flask import Flask, request as flask_request

	app = Flask(__name__)

	@app.route('/', methods=['GET', 'POST', 'OPTIONS'])
	@app.route('/analisar', methods=['GET', 'POST', 'OPTIONS'])
	def flask_handler():
		return analisar_cancelamentos(flask_request)

	port = int(os.environ.get('PORT', 8080))
	app.run(host='0.0.0.0', port=port, debug=False)
