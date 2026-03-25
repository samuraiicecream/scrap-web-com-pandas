import pandas as pd

print("Carregando os arquivos...")
df_visitantes = pd.read_csv('visitantes_extraidos.csv')

df_matriculados = pd.read_excel('matriculados.xlsx')

print("Limpando e padronizando os e-mails")

df_visitantes['Email'] = df_visitantes['Email'].str.lower().str.strip()
df_visitantes['Nome'] = df_visitantes['Nome'].str.lower().str.strip()

df_matriculados['Email'] = df_matriculados['Email'].str.lower().str.strip()
df_matriculados['Nome'] = df_matriculados['Nome'].str.lower().str.strip()

print("Cruzando os dados...")

df_resultado = pd.merge(df_visitantes, df_matriculados, on='Email', how='inner')
df_resultado = pd.merge(df_visitantes, df_matriculados, on='Nome', how='inner')

print(f"\nCruzamento finalizado! Encontramos {len(df_resultado)} alunos que visitaram e se matricularam." )
print("\n-- PRIMEIRAS LINHAS DO RESULTADO ---")
print(df_resultado.head())

df_resultado.to_excel('resultado_final_dashboard.xlsx', index=False)
print("\nArquivo 'resultado_final_dashboard.xlsx' gerado com sucesso!")