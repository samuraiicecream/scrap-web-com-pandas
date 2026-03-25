import pandas as pd 

print("carregando arquivos...")

df_visitantes = pd.read_csv('visitantes_extraidos.csv')


df_matriculados = pd.read_excel('matriculados.xlsx')

print(f"--> Visitantes extraídos pelo robô: {len(df_visitantes)} linhas")
print(f"--> Alunos matriculados no Excel: {len(df_matriculados)}")

print("\n2. Executando a Super Limpeza...")




df_visitantes['Email'] = df_visitantes['Email'].astype(str).str.lower().str.replace(r'\s+', '', regex=True)
df_visitantes['Nome'] = df_visitantes['Nome'].astype(str).str.lower().str.replace(r'\s+', '', regex=True)

df_matriculados['Email'] = df_visitantes['Email'].astype(str).str.lower().str.replace(r'\s+', '', regex=True)
df_matriculados['Nome'] = df_visitantes['Nome'].astype(str).str.lower().str.replace(r'\s+', '', regex=True)
df_visitantes['Veio do Site'] = 'Sim'
df_visitantes_unicos = df_visitantes[['Email','Nome','Veio do Site']].drop_duplicates()
print("Cruzando as informações...")


df_resultado = pd.merge(df_visitantes_unicos, df_matriculados, on='Email', how='left')
df_resultado = pd.merge(df_visitantes_unicos, df_matriculados, on='Nome', how='left')

df_resultado['Veio do Site'] = df_resultado['Veio do Site'].fillna('Não')
total_sim = len(df_resultado[df_resultado['Veio do Site'] == 'Sim'])
total_nao = len(df_resultado[df_resultado['Veio do Site'] == 'Não'])

print(f"\n📊 RESUMO PARA O DASHBOARD:")
print(f"-> Matriculados que vieram do Site: {total_sim}")
print(f"-> Matriculados por outros meios: {total_nao}")

df_resultado.to_excel('base_final_dashboard.xlsx', index=False)
print("\n Sucesso! O arquivo 'base_final_dashboard.xlsx' está pronto!")