import pandas as pd
import random
from datetime import datetime, timedelta

# --- CONFIGURAÇÕES PARA PARECER REAL ---
# Baseado nos seus prints da Elecnor/Deloitte
disciplinas = ["INSTRUMENTAÇÃO", "ELÉTRICA", "MECÂNICA", "TUBULAÇÃO", "CIVIL"]
descricoes = [
    "SUPPORTS AND MISCELLANEOUS - LOTE A",
    "VÁLVULAS ON-OFF ESFERA",
    "PAINÉIS DE ILUMINAÇÃO E FORÇA",
    "ESTRUTURA METÁLICA - PIPE RACK",
    "CABOS DE MÉDIA TENSÃO",
    "BOMBAS CENTRÍFUGAS - TAG 201",
    "INSTRUMENTOS DE MEDIÇÃO DE VAZÃO"
]
responsaveis = ["Vitor Dias", "Engenharia", "Suporte", "Coordenação"]
fornecedores = ["BRAFER", "METTA", "VALMET FLOW", "GH DO BRASIL", "TECHSTEEL", "SIEMENS"]
status_opcoes = ["Entregue", "Em Trânsito", "Em Fabricação", "PO emitido", "CANCELADO", "NOVO"]
setores = ["Diligenciamento", "Engenharia", "Compras"]

# Função para gerar datas realistas (ou vazio se for NOVO)
def gerar_data(status):
    if status in ["CANCELADO", "NOVO"]:
        return ""
    dias = random.randint(-30, 120) # De 1 mês atrás até 4 meses na frente
    return (datetime.now() + timedelta(days=dias)).strftime("%d/%m/%Y")

# Função para gerar código RMT igual ao print (Ex: 26-21-00-RMT-EG-0000-015)
def gerar_rmt(i):
    return f"26-21-00-RMT-EG-0000-{i:03d}"

# --- GERANDO A BASE DE DADOS ---
dados = []

# Vamos criar 50 linhas de dados
for i in range(1, 51):
    status_escolhido = random.choice(status_opcoes)
    
    # Lógica para a coluna "R" (Revisão) igual ao print (0, 1, 2...)
    revisao = str(random.randint(0, 3))
    
    dados.append({
        "Disciplina": random.choice(disciplinas),
        "RMT": gerar_rmt(i),  # A COLUNA CHAVE
        "Descrição": random.choice(descricoes),
        "R": revisao,         # A COLUNA DE REVISÃO (Pequena)
        "Setor": random.choice(setores),
        "Responsável": "Vitor Dias", # Baseado no print
        "Status": status_escolhido,
        "Fornecedor": random.choice(fornecedores),
        "Previsão de entrega": gerar_data(status_escolhido)
    })

# Criando DataFrame
df = pd.DataFrame(dados)

# Salvando
arquivo_saida = "RMs_Criticas_Oficial_Fake.xlsx"
df.to_excel(arquivo_saida, index=False)

print(f"✅ Planilha '{arquivo_saida}' gerada com sucesso!")
print("Estrutura das colunas:")
print(df.columns.tolist())