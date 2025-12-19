import pandas as pd
import random
from datetime import datetime, timedelta

# --- CONFIGURAÇÕES ---
def gerar_codigo_rmt(indice):
    return f"26-21-00-RMT-EG-0000-{indice:03d}"

materiais = [
    "Válvula Esfera 3/4 Pol Inox", "Tubo PVC Rígido 100mm", 
    "Cabo Elétrico 10mm Isolado", "Sensor de Pressão Digital",
    "Flange Aço Carbono ANSI 150", "Bomba Centrífuga 5CV"
]
fornecedores = ["TechSteel Ltda", "Engenharia S.A.", "Construtora Norte", "Global Pipes"]
status_possiveis = ["Entregue", "Em Trânsito", "Em Fabricação", "Atrasado"]

# --- GERAR BASE DE DADOS (RMs CRÍTICAS) ---
dados = []
for i in range(1, 51): # Cria 50 RMs
    data_futura = datetime.now() + timedelta(days=random.randint(-10, 60))
    dados.append({
        'RMT': gerar_codigo_rmt(i),
        'Descrição': random.choice(materiais),
        'Fornecedor': random.choice(fornecedores),
        'Status': random.choice(status_possiveis),
        'Previsão de entrega': data_futura.strftime("%d/%m/%Y"),
        'Revisão': str(random.randint(0, 5))
    })

df = pd.DataFrame(dados)
df.to_excel('RMs_Criticas_Realista.xlsx', index=False)
print("Arquivo 'RMs_Criticas_Realista.xlsx' criado com sucesso!")