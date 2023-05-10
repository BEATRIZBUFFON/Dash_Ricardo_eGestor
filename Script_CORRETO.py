import pandas as pd

df1 = "spotter_etapa/etapa_janeiro.xlsx"
df1 = pd.read_excel(df1)

df2 = "spotter_etapa/etapa_fevereiro.xlsx"
df2 = pd.read_excel(df2)

df3 = "spotter_etapa/etapa_marco.xlsx"
df3 = pd.read_excel(df3)

df4 = "spotter_etapa/etapa_abril.xlsx"
df4 = pd.read_excel(df4)

df5 = "spotter_etapa/etapa_maio.xlsx"
df5 = pd.read_excel(df5)

def teste(df):
    df = df.drop([0, 1])
    df.columns = df.iloc[0]
    df = df[1:]
    df = df.reset_index(drop=True)
    df['TOTAL'] = df['TOTAL'].astype(int)
    df.rename(columns={df.columns[1]: "Pré-Vendedor"}, inplace=True)
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('$$Gu Gabriel Martins', 'Gabriel Martins')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('%$ Ana Paula Ramos', 'Ana Paula Ramos')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('%$ Henrique Lima', 'Henrique Lima')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@@Augusto Barreto', 'Augusto Barreto')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@@Beatriz Kowas', 'Beatriz Kowas')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@@Marcelo Machado', 'Marcelo Machado')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('%$ Milena Schuquel Machado', 'Milena Schuquel Machado')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@@Clarine Mantai', 'Clarine Mantai')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('$$Ca Daniel Escobar', 'Daniel Escobar')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('$$Ca Eduarda Fragoso', 'Eduarda Fragoso')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('$$Gu Felipe Fiorini', 'Felipe Fiorini')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('$$Li Renata Lopes', 'Renata Lopes')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('$$Va Patrícia Lima', 'Patrícia Lima')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('%$Li Tamires Cattani', 'Tamires Cattani')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@We Jennifer Dallanora', 'Jennifer Dallanora')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@We Matheus Mesquita Mesquita', 'Matheus Mesquita')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('$$Li Andressa Lorenzoni', 'Andressa Lorenzoni')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('$$Li Frederico Vargas', 'Frederico Vargas')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('$$Va Andressa Ely', 'Andressa Ely')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('$$Va Gabriele Motta', 'Gabriele Motta')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('%$ Pedro  Machado', 'Pedro Machado')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@@Kalleo Ethur', 'Kalleo Ethur')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@@Patrícia Melo', 'Patrícia Melo')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('%$ Jeniffer Soares', 'Jeniffer Soares')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@@Kauhan Cunha', 'Kauhan Cunha')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@@Brendon Silva', 'Brendon Silva') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@@Paola Lanes Viaro', 'Paola Lanes Viaro')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@@Gabriel Padoin', 'Gabriel Padoin')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('Sh Patricia Barcarolo', 'Patricia Barcarolo')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('Ma Lilian Palmeira', 'Lilian Palmeira')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@Yu Wellington Reis', 'Wellington Reis') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('Va Sharom Lopes', 'Sharom Lopes') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('Fi Marcelo Rangel', 'Marcelo Rangel') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@Yu Chaeni Franciele Dias', 'Chaeni Franciele Dias')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@@Andriele  Ferigolo', 'Andriele Ferigolo') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@Canais_Planilhas Wellington', 'Planilhas Wellington') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('#$Al Alisson  Leal', 'Alisson Leal') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('#$ Bruna Fioravante', 'Bruna Fioravante') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('Do Brendon Silva', 'Brendon Silva') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('Do Willian Martins', 'Willian Martins') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('SM Giovanni  Zanela', 'Giovanni Zanela') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@Rev Yuri Martins', 'Yuri Martins') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@Rev Leonardo Fagundes', 'Leonardo Fagundes') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('Alexander  Leite da Silva', 'Alexander Leite da Silva') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('Gi Eric Patrick', 'Eric Patrick') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('Natali  Saccol', 'Natali Saccol') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('%$ Eduarda Donadel', 'Eduarda Donadel') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('#$ Estéfani Campos', 'Estéfani Campos') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('Fernando  Ruoso', 'Fernando Ruoso')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('%$ Gianfranco Dalcin', 'Gianfranco Dalcin')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('Sh Ana Flávia Dutra', 'Ana Flávia Dutra')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('Ma Benhur Carvalho', 'Benhur Carvalho')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('Ma Rodrigo Silva ', 'Rodrigo Silva')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('Mi Nathalia  Dias ', 'Nathalia  Dias')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('Va Evelin Dutra ', 'Evelin Dutra')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('#$ Paula Fioravante', 'Paula Fioravante')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('#$ Almir Felipe', 'Almir Felipe')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('Gi Camila Wrasse', 'Camila Wrasse') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@Yu Amanda Keller', 'Amanda Keller') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('#$ Pedro Junger', 'Pedro Junger') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('#$ Luiz Guilherme Mello', 'Luiz Guilherme Mello') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('#$ Eduarda Puchale', 'Eduarda Puchale') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('Larissa  Staggemeier Soares', 'Larissa Staggemeier Soares') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('@Yu Arthur Nunes Nunes', 'Arthur Nunes') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('#$ Eduarda Puchale', 'Eduarda Puchale') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('Do Guilherme Augusto', 'Guilherme Augusto') 
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace('Mi Nathalia  Dias', 'Nathalia  Dias') 
    return df

df1t = teste(df1)
df2t = teste(df2)
df3t = teste(df3)
df4t = teste(df4)
df5t = teste(df5)

df1t['Dia '] = pd.to_datetime(df1t['Dia '], format='%d').apply(lambda x: x.replace(year=2023, month=1))
df1t = df1t.dropna(subset=['Dia ']) 
df1t['Dia '] = df1t['Dia '].dt.strftime('%d-%m-%Y')

df2t['Dia '] = pd.to_datetime(df2t['Dia '], format='%d').apply(lambda x: x.replace(year=2023, month=2))
df2t = df2t.dropna(subset=['Dia ']) 
df2t['Dia '] = df2t['Dia '].dt.strftime('%d-%m-%Y')

df3t['Dia '] = pd.to_datetime(df3t['Dia '], format='%d').apply(lambda x: x.replace(year=2023, month=3))
df3t = df3t.dropna(subset=['Dia ']) 
df3t['Dia '] = df3t['Dia '].dt.strftime('%d-%m-%Y')

df4t['Dia '] = pd.to_datetime(df4t['Dia '], format='%d').apply(lambda x: x.replace(year=2023, month=4))
df4t = df4t.dropna(subset=['Dia ']) 
df4t['Dia '] = df4t['Dia '].dt.strftime('%d-%m-%Y')

df5t['Dia '] = pd.to_datetime(df5t['Dia '], format='%d').apply(lambda x: x.replace(year=2023, month=5))
df5t = df5t.dropna(subset=['Dia ']) 
df5t['Dia '] = df5t['Dia '].dt.strftime('%d-%m-%Y')

df_pronto = pd.concat([df1t, df2t, df3t, df4t, df5t ], ignore_index=True)

vendedores_desejados = ['Gabriel Martins', 'Augusto Barreto',
                        'Eduarda Fragoso', 'Patrícia Lima',
                        'Matheus Mesquita', 'Andressa Ely',
                        'Patrícia Melo', 'Gabriel Padoin', 
                        'Alan Carvalho', 'Marcelo Rangel',
                        'Estéfani Campos', 'Camila Wrasse',
                        'Jaine Tonetto', 'Almir Felipe',
                        'Ana Paula Ramos', 'Beatriz Kowas',
                        'Clarine Mantai', 'Felipe Fiorini',
                        'Andressa Lorenzoni', 'Gabrile Motta', 
                        'Alisson Leal', 'Giovanni Zanela',
                        'Luiz Guilherme Mello', 'Henrique Lima',
                        'Marcelo Machado', 'Daniel Escobar',
                        'Renata Lopes', 'Jennifer Dallanora',
                        'Pedro Machado', 'Jeniffer Soares',
                        'Lilian Palmeira', 'Wellington Reis',
                        'Suelen Rossi', 'Yuri Martins',
                        'Vanessa Ortiz', 'Gianfranco Dalcin',
                        'Rodrigo Silva', 'Paula Fioravante',
                        'Naiane Bock', 'Pedro Junger',
                        'Kalleo Ethur', 'Fernando Penz']
df_pronto = df_pronto[df_pronto['Pré-Vendedor'].isin(vendedores_desejados)]

df_pronto.to_excel('Etapa2.xlsx', index=False)

df_pronto['Pré-Vendedor'].unique()
# TEMPERATURA

df11 = "spotter_temp/temperatura_janeiro.xlsx"
df11 = pd.read_excel(df11)

df22 = "spotter_temp/temperatura_fevereiro.xlsx"
df22 = pd.read_excel(df22)

df33 = "spotter_temp/temperatura_marco.xlsx"
df33 = pd.read_excel(df33)

df44 = "spotter_temp/temperatura_abril.xlsx"
df44 = pd.read_excel(df44)

df55 = "spotter_temp/temperatura_maio.xlsx"
df55 = pd.read_excel(df55)

def teste2(df):
    df = df.drop([0, 1])
    df.columns = df.iloc[0]
    df = df[1:]
    df = df.reset_index(drop=True)
    df['TOTAL'] = df['TOTAL'].astype(int)
    df.rename(columns={df.columns[1]: "Pré-Vendedor"}, inplace=True)
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Gu Gabriel Martins', 'Gabriel Martins')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '%$ Ana Paula Ramos', 'Ana Paula Ramos')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '%$ Henrique Lima', 'Henrique Lima')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Augusto Barreto', 'Augusto Barreto')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Beatriz Kowas', 'Beatriz Kowas')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Marcelo Machado', 'Marcelo Machado')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '%$ Milena Schuquel Machado', 'Milena Schuquel Machado')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Clarine Mantai', 'Clarine Mantai')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Ca Daniel Escobar', 'Daniel Escobar')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Ca Eduarda Fragoso', 'Eduarda Fragoso')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Gu Felipe Fiorini', 'Felipe Fiorini')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Li Renata Lopes', 'Renata Lopes')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Va Patrícia Lima', 'Patrícia Lima')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '%$Li Tamires Cattani', 'Tamires Cattani')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@We Jennifer Dallanora', 'Jennifer Dallanora')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@We Matheus Mesquita Mesquita', 'Matheus Mesquita')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Li Andressa Lorenzoni', 'Andressa Lorenzoni')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Li Frederico Vargas', 'Frederico Vargas')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Va Andressa Ely', 'Andressa Ely')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Va Gabriele Motta', 'Gabriele Motta')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '%$ Pedro  Machado', 'Pedro Machado')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Kalleo Ethur', 'Kalleo Ethur')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Patrícia Melo', 'Patrícia Melo')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '%$ Jeniffer Soares', 'Jeniffer Soares')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Kauhan Cunha', 'Kauhan Cunha')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Brendon Silva', 'Brendon Silva')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Paola Lanes Viaro', 'Paola Lanes Viaro')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Gabriel Padoin', 'Gabriel Padoin')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Sh Patricia Barcarolo', 'Patricia Barcarolo')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Ma Lilian Palmeira', 'Lilian Palmeira')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@Yu Wellington Reis', 'Wellington Reis')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Va Sharom Lopes', 'Sharom Lopes')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Fi Marcelo Rangel', 'Marcelo Rangel')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@Yu Chaeni Franciele Dias', 'Chaeni Franciele Dias')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Andriele  Ferigolo', 'Andriele Ferigolo')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@Canais_Planilhas Wellington', 'Planilhas Wellington')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '#$Al Alisson  Leal', 'Alisson Leal')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '#$ Bruna Fioravante', 'Bruna Fioravante')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Do Brendon Silva', 'Brendon Silva')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Do Willian Martins', 'Willian Martins')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'SM Giovanni  Zanela', 'Giovanni Zanela')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@Rev Yuri Martins', 'Yuri Martins')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@Rev Leonardo Fagundes', 'Leonardo Fagundes')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Alexander  Leite da Silva', 'Alexander Leite da Silva')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Gi Eric Patrick', 'Eric Patrick')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Natali  Saccol', 'Natali Saccol')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '%$ Eduarda Donadel', 'Eduarda Donadel')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '#$ Estéfani Campos', 'Estéfani Campos')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Fernando  Ruoso', 'Fernando Ruoso')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '%$ Gianfranco Dalcin', 'Gianfranco Dalcin')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Sh Ana Flávia Dutra', 'Ana Flávia Dutra')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Ma Benhur Carvalho', 'Benhur Carvalho')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Ma Rodrigo Silva ', 'Rodrigo Silva')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Mi Nathalia  Dias ', 'Nathalia  Dias')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Va Evelin Dutra ', 'Evelin Dutra')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '#$ Paula Fioravante', 'Paula Fioravante')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '#$ Almir Felipe', 'Almir Felipe')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Gi Camila Wrasse', 'Camila Wrasse')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@Yu Amanda Keller', 'Amanda Keller')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '#$ Pedro Junger', 'Pedro Junger')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '#$ Luiz Guilherme Mello', 'Luiz Guilherme Mello')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '#$ Eduarda Puchale', 'Eduarda Puchale')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Larissa  Staggemeier Soares', 'Larissa Staggemeier Soares')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@Yu Arthur Nunes Nunes', 'Arthur Nunes')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '#$ Eduarda Puchale', 'Eduarda Puchale')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Do Guilherme Augusto', 'Guilherme Augusto')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Mi Nathalia  Dias', 'Nathalia  Dias')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Pedro  Escobar (gerente)', 'Pedro Escobar (gerente)')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Ma Filipe Martins', 'Filipe Martins')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Li Raiane Prates', 'Raiane Prates')
    return df

df1p = teste(df11)
df2p = teste(df22)
df3p = teste(df33)
df4p = teste(df44)
df5p = teste(df55)

df1p['Dia '] = pd.to_datetime(df1p['Dia '], format='%d').apply(lambda x: x.replace(year=2023, month=1))
df1p = df1p.dropna(subset=['Dia ']) 
df1p['Dia '] = df1p['Dia '].dt.strftime('%d-%m-%Y')

df2p['Dia '] = pd.to_datetime(df2p['Dia '], format='%d').apply(lambda x: x.replace(year=2023, month=2))
df2p = df2p.dropna(subset=['Dia ']) 
df2p['Dia '] = df2p['Dia '].dt.strftime('%d-%m-%Y')

df3p['Dia '] = pd.to_datetime(df3p['Dia '], format='%d').apply(lambda x: x.replace(year=2023, month=3))
df3p = df3p.dropna(subset=['Dia ']) 
df3p['Dia '] = df3p['Dia '].dt.strftime('%d-%m-%Y')

df4p['Dia '] = pd.to_datetime(df4p['Dia '], format='%d').apply(lambda x: x.replace(year=2023, month=4))
df4p = df4p.dropna(subset=['Dia ']) 
df4p['Dia '] = df4p['Dia '].dt.strftime('%d-%m-%Y')

df5p['Dia '] = pd.to_datetime(df5p['Dia '], format='%d').apply(lambda x: x.replace(year=2023, month=5))
df5p = df5p.dropna(subset=['Dia ']) 
df5p['Dia '] = df5p['Dia '].dt.strftime('%d-%m-%Y')

df_pronto2 = pd.concat([df1p, df2p, df3p, df4p, df5p], ignore_index=True)
df_pronto2 = df_pronto2[df_pronto2['Pré-Vendedor'].isin(vendedores_desejados)]

df_pronto2.to_excel('Temperatura2.xlsx', index=False)



# OCORRÊNCIA DE REUNIÃO

df111 = "spotter_ocorrencia/ocorrencia_janeiro.xlsx"
df111 = pd.read_excel(df111)

df222 = "spotter_ocorrencia/ocorrencia_fevereiro.xlsx"
df222 = pd.read_excel(df222)

df333 = "spotter_ocorrencia/ocorrencia_marco.xlsx"
df333 = pd.read_excel(df333)

df444 = "spotter_ocorrencia/ocorrencia_abril.xlsx"
df444 = pd.read_excel(df444)

df555 = "spotter_ocorrencia/ocorrencia_maio.xlsx"
df555 = pd.read_excel(df555)


df111 = pd.read_excel('ocorrencia_janeiro.xlsx')
df222 = pd.read_excel('ocorrencia_fevereiro.xlsx')
df333 = pd.read_excel('ocorrencia_marco.xlsx') 
df444 = pd.read_excel('ocorrencia_abril.xlsx')
df555 = pd.read_excel('ocorrencia_maio.xlsx')

def teste3(df):
    df = df.drop([0, 1])
    df.columns = df.iloc[0]
    df = df[1:]
    df = df.reset_index(drop=True)
    df['TOTAL'] = df['TOTAL'].astype(int)
    df.rename(columns={df.columns[1]: "Pré-Vendedor"}, inplace=True)
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Gu Gabriel Martins', 'Gabriel Martins')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '%$ Ana Paula Ramos', 'Ana Paula Ramos')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '%$ Henrique Lima', 'Henrique Lima')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Augusto Barreto', 'Augusto Barreto')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Beatriz Kowas', 'Beatriz Kowas')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Marcelo Machado', 'Marcelo Machado')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '%$ Milena Schuquel Machado', 'Milena Schuquel Machado')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Clarine Mantai', 'Clarine Mantai')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Ca Daniel Escobar', 'Daniel Escobar')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Ca Eduarda Fragoso', 'Eduarda Fragoso')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Gu Felipe Fiorini', 'Felipe Fiorini')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Li Renata Lopes', 'Renata Lopes')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Va Patrícia Lima', 'Patrícia Lima')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '%$Li Tamires Cattani', 'Tamires Cattani')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@We Jennifer Dallanora', 'Jennifer Dallanora')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@We Matheus Mesquita Mesquita', 'Matheus Mesquita')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Li Andressa Lorenzoni', 'Andressa Lorenzoni')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Li Frederico Vargas', 'Frederico Vargas')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Va Andressa Ely', 'Andressa Ely')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '$$Va Gabriele Motta', 'Gabriele Motta')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '%$ Pedro  Machado', 'Pedro Machado')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Kalleo Ethur', 'Kalleo Ethur')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Patrícia Melo', 'Patrícia Melo')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '%$ Jeniffer Soares', 'Jeniffer Soares')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Kauhan Cunha', 'Kauhan Cunha')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Brendon Silva', 'Brendon Silva')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Paola Lanes Viaro', 'Paola Lanes Viaro')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Gabriel Padoin', 'Gabriel Padoin')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Sh Patricia Barcarolo', 'Patricia Barcarolo')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Ma Lilian Palmeira', 'Lilian Palmeira')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@Yu Wellington Reis', 'Wellington Reis')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Va Sharom Lopes', 'Sharom Lopes')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Fi Marcelo Rangel', 'Marcelo Rangel')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@Yu Chaeni Franciele Dias', 'Chaeni Franciele Dias')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@@Andriele  Ferigolo', 'Andriele Ferigolo')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@Canais_Planilhas Wellington', 'Planilhas Wellington')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '#$Al Alisson  Leal', 'Alisson Leal')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '#$ Bruna Fioravante', 'Bruna Fioravante')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Do Brendon Silva', 'Brendon Silva')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Do Willian Martins', 'Willian Martins')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'SM Giovanni  Zanela', 'Giovanni Zanela')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@Rev Yuri Martins', 'Yuri Martins')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@Rev Leonardo Fagundes', 'Leonardo Fagundes')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Alexander  Leite da Silva', 'Alexander Leite da Silva')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Gi Eric Patrick', 'Eric Patrick')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Natali  Saccol', 'Natali Saccol')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '%$ Eduarda Donadel', 'Eduarda Donadel')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '#$ Estéfani Campos', 'Estéfani Campos')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Fernando  Ruoso', 'Fernando Ruoso')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '%$ Gianfranco Dalcin', 'Gianfranco Dalcin')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Sh Ana Flávia Dutra', 'Ana Flávia Dutra')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Ma Benhur Carvalho', 'Benhur Carvalho')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Ma Rodrigo Silva ', 'Rodrigo Silva')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Mi Nathalia  Dias ', 'Nathalia  Dias')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Va Evelin Dutra ', 'Evelin Dutra')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '#$ Paula Fioravante', 'Paula Fioravante')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '#$ Almir Felipe', 'Almir Felipe')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Gi Camila Wrasse', 'Camila Wrasse')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@Yu Amanda Keller', 'Amanda Keller')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '#$ Pedro Junger', 'Pedro Junger')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '#$ Luiz Guilherme Mello', 'Luiz Guilherme Mello')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '#$ Eduarda Puchale', 'Eduarda Puchale')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Larissa  Staggemeier Soares', 'Larissa Staggemeier Soares')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '@Yu Arthur Nunes Nunes', 'Arthur Nunes')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        '#$ Eduarda Puchale', 'Eduarda Puchale')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Do Guilherme Augusto', 'Guilherme Augusto')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Mi Nathalia  Dias', 'Nathalia  Dias')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Pedro  Escobar (gerente)', 'Pedro Escobar (gerente)')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Ma Filipe Martins', 'Filipe Martins')
    df['Pré-Vendedor'] = df['Pré-Vendedor'].replace(
        'Li Raiane Prates', 'Raiane Prates')
    return df

df1o = teste(df111)
df2o = teste(df222)
df3o = teste(df333)
df4o = teste(df444)
df5o = teste(df555)

df1o['Dia '] = pd.to_datetime(df1o['Dia '], format='%d').apply(lambda x: x.replace(year=2023, month=1))
df1o = df1o.dropna(subset=['Dia ']) 
df1o['Dia '] = df1o['Dia '].dt.strftime('%d-%m-%Y')

df2o['Dia '] = pd.to_datetime(df2o['Dia '], format='%d').apply(lambda x: x.replace(year=2023, month=2))
df2o = df2o.dropna(subset=['Dia ']) 
df2o['Dia '] = df2o['Dia '].dt.strftime('%d-%m-%Y')

df3o['Dia '] = pd.to_datetime(df3o['Dia '], format='%d').apply(lambda x: x.replace(year=2023, month=3))
df3o = df3o.dropna(subset=['Dia ']) 
df3o['Dia '] = df3o['Dia '].dt.strftime('%d-%m-%Y')

df4o['Dia '] = pd.to_datetime(df4o['Dia '], format='%d').apply(lambda x: x.replace(year=2023, month=4))
df4o = df4o.dropna(subset=['Dia ']) 
df4o['Dia '] = df4o['Dia '].dt.strftime('%d-%m-%Y')

df5o['Dia '] = pd.to_datetime(df5o['Dia '], format='%d').apply(lambda x: x.replace(year=2023, month=5))
df5o = df5o.dropna(subset=['Dia ']) 
df5o['Dia '] = df5o['Dia '].dt.strftime('%d-%m-%Y')

df_pronto3 = pd.concat([df1o, df2o, df3o, df4o, df5o], ignore_index=True)
df_pronto3 = df_pronto3[df_pronto3['Pré-Vendedor'].isin(vendedores_desejados)]

df_pronto3.to_excel('Ocorrencia2.xlsx', index=False)
