import pandas as pd
import matplotlib.pyplot as plt
import io
import base64
from flask import Flask, render_template, request

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    file_path = r'C:\Users\Alan Milan\Documents\dashboardhtml\Base de Dados.xlsx'
    sheet_name = 'Sheet1'
    
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df_cleaned = df.drop_duplicates()

    df_cleaned = df_cleaned.apply(lambda col: col.map(lambda x: ''.join([c for c in str(x) if c.isalnum() or c.isspace()]) if isinstance(x, str) else x))

    # Função para criar o gráfico com Matplotlib
    def create_plot(plot_type):
        fig, ax = plt.subplots()
        if plot_type == 'scatter':
            operadores = df_cleaned['Operador'].astype('category').cat.codes
            scatter = ax.scatter(df_cleaned['Leads Recebidos'], df_cleaned['Vendas Realizadas'], c=operadores, cmap='tab10')
            ax.set_title('Dispersão entre Leads Recebidos e Vendas Realizadas')
            ax.set_xlabel('Leads Recebidos')
            ax.set_ylabel('Vendas Realizadas')
            plt.colorbar(scatter, ax=ax, label='Operador')
        elif plot_type == 'line':
            for operador in df_cleaned['Operador'].unique():
                subset = df_cleaned[df_cleaned['Operador'] == operador]
                ax.plot(subset['Mês'], subset['Vendas Realizadas'], label=operador)
            ax.set_title('Vendas Realizadas ao Longo dos Meses')
            ax.set_xlabel('Mês')
            ax.set_ylabel('Vendas Realizadas')
            ax.legend()
        elif plot_type == 'bar':
            df_cleaned.groupby('Operador')['Vendas Realizadas'].sum().plot(kind='bar', ax=ax)
            ax.set_title('Vendas Realizadas por Operador')
            ax.set_xlabel('Operador')
            ax.set_ylabel('Vendas Realizadas')
        elif plot_type == 'pie':
            df_cleaned.groupby('Operador')['Vendas Realizadas'].sum().plot(kind='pie', ax=ax, autopct='%1.1f%%', legend=False)
            ax.set_title('Participação de Vendas por Operador')

        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        img = base64.b64encode(buf.getvalue()).decode('utf-8')
        plt.close(fig)
        return img

    plot_type = request.form.get('chart-type', 'scatter')
    graph_img = create_plot(plot_type)

    return render_template('index.html', 
                           table=df_cleaned.to_html(classes='table table-striped'),
                           graph_img=graph_img)

@app.route('/relatorios')
def relatorios():
    file_path = r'C:\Users\Alan Milan\Documents\dashboardhtml\Base de Dados.xlsx'
    sheet_name = 'Sheet1'
    
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df_cleaned = df.drop_duplicates()
    df_cleaned = df_cleaned.apply(lambda col: col.map(lambda x: ''.join([c for c in str(x) if c.isalnum() or c.isspace()]) if isinstance(x, str) else x))

    total_vendas = df_cleaned['Vendas Realizadas'].sum()
    operador_max_vendas = df_cleaned.groupby('Operador')['Vendas Realizadas'].sum().idxmax()
    crescimento_mensal = df_cleaned.groupby('Mês')['Vendas Realizadas'].sum().pct_change().fillna(0).mean()

    summary = (
        f"Total de Vendas Realizadas: {total_vendas:.2f}\n"
        f"Operador com Maior Número de Vendas: {operador_max_vendas}\n"
        f"Crescimento Médio Mensal: {crescimento_mensal:.2%}"
    )

    return render_template('relatorios.html', summary=summary)

if __name__ == '__main__':
    app.run(debug=True)