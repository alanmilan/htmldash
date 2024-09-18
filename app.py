from flask import Flask, render_template, request
import pandas as pd
import io
import base64
import matplotlib.pyplot as plt
import matplotlib
# Configure o backend do matplotlib para não usar a interface gráfica
matplotlib.use('Agg')

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    # Carregar e processar os dados
    file_path = 'C:\\Users\\cwb.editor\\Documents\\idle\\dashboardhtml\\Base de Dados.xlsx'
    df = pd.read_excel(file_path, sheet_name='Sheet1')

    graph_img = None  # Inicializa a variável do gráfico

    if request.method == 'POST':
        chart_type = request.form.get('chart-type')
        unit = request.form.get('unit')
        operator = request.form.get('operator')
        month = request.form.get('month')
        filter_option = request.form.get('filter-options')

        # Aplicar filtros
        if unit:
            df = df[df['Unidade'] == unit]
        if operator:
            df = df[df['Operador'] == operator]
        if month:
            df = df[df['Mês'] == month]
        
        # Gerar gráfico com base no tipo selecionado
        fig, ax = plt.subplots()
        
        if chart_type == 'scatter':
            df.plot(kind='scatter', x='Vendas Realizadas', y='Leads Recebidos', ax=ax)
        elif chart_type == 'line':
            df.plot(kind='line', x='Mês', y='Vendas Realizadas', ax=ax)
        elif chart_type == 'bar':
            df.plot(kind='bar', x='Unidade', y='Vendas Realizadas', ax=ax)
        elif chart_type == 'pie':
            df.groupby('Unidade')['Vendas Realizadas'].sum().plot(kind='pie', ax=ax)

        # Adicionar gráfico baseado na opção de filtro
        if filter_option and filter_option != "Nenhum":
            filter_column = filter_option.split(':')[0]
            df[filter_column].plot(kind='bar', ax=ax, label=filter_column)
            ax.legend()

        plt.close(fig)

        # Salvar o gráfico em base64
        img = io.BytesIO()
        fig.savefig(img, format='png', bbox_inches='tight')
        img.seek(0)
        graph_img = base64.b64encode(img.getvalue()).decode('utf-8')

    # Gerar tabela HTML
    table = df.to_html(classes='data')

    return render_template('index.html', graph_img=graph_img, table=table)

if __name__ == '__main__':
    app.run(debug=True)
