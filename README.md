# Automação de Cadastro de Produtos

Este script Python foi desenvolvido para automatizar o preenchimento de dados no site de cadastro de produtos hospedado em https://cadastro-produtos-devaprender.netlify.app/index.html.

### Requisitos

- Python 3.x
- Bibliotecas:
- openpyxl
- pyperclip
- pyautogui

### Como Usar

- Certifique-se de ter o arquivo Excel chamado produtos_ficticios.xlsx no mesmo diretório do script.

- Abra o site https://cadastro-produtos-devaprender.netlify.app/index.html em um navegador.

- Execute o script:

- O script acessará a planilha "Produtos" do arquivo Excel, copiará os valores das células e os colará nos campos correspondentes no site. Utilize as coordenadas (x, y) do mouse para interagir com os elementos da página.

- O script preencherá os campos com os dados da planilha, clicará em botões específicos e aguardará alguns segundos (definidos por sleep) para garantir a conclusão de cada etapa.

- Lembre-se de ajustar as coordenadas e os tempos de espera conforme necessário, dependendo da interface gráfica do seu sistema e das características do site.

- Observação: Certifique-se de ter backups dos seus dados antes de executar o script, pois ele realiza ações automatizadas que podem modificar os dados no site.

### Referência

Este script foi desenvolvido com base em um tutorial do canal Dev Aprender, por Jhonatan de Souza.
