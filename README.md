# Print Google Forms
![logo](https://raw.githubusercontent.com/raylan-oliveira/impressao_google_forms/main/img/icone.ico)

Projeto, em python, que imprime dados (linhas) de uma planilha do google, esses dados são gerados por formulário do google. Os pdf's gerado ficam salvo na área de trabalho.

[**Download última versão!!**](https://github.com/raylan-oliveira/impressao_google_forms/releases/latest)
## Demonstração:
![Demon](https://github.com/raylan-oliveira/impressao_google_forms/main/img/demo.gif)

### Dependências
   ```sh
	pip install -r requeriments.txt
   ```
   
### Executar
   ```sh
	python formulario.py
   ```
	
### Compilar
   ```sh
	pip install pyinstaller
	pyinstaller --noconsole --icon="img\icone.ico" formulario.py	
   ```