from flask import Flask, render_template
from flask import request
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN ,MSO_ANCHOR
from datetime import datetime
from time import sleep
import shutil
from waitress import serve

app = Flask(__name__)

# if __name__ == '__main__':
#     port = int(os.getenv('PORT'), '5000')
#     app.run(host='0.0.0.0', port = port)

# app = Flask(__name__, template_folder='tempalte', static_folder='tempalte/assets')

@app.route("/")
def index():
    return render_template("principal.html")

@app.route("/convert")
def conversor():
    return render_template("conversor.html")

@app.route('/autenticar', methods=['GET'])
def autenticar():
    vsl = request.args.get('vsl');
    vsl = vsl.split('\n')
    x = []

    for item in vsl:
        st = item.replace('\r', '')
        x.append(st)

    try:
        #Abrindo PowerPoint
        apresentacao = Presentation()
        apresentacao.slide_width = Inches(16)
        apresentacao.slide_height = Inches(9)

        # Cada linah do arquivo.txt ira fazer esse bloco
        for line in x:

            #Criando um slide
            slide = apresentacao.slides.add_slide(apresentacao.slide_layouts[6])
            

            #Criando e centralizando a TextBox
            x = Inches(3)
            y = Inches(3.5)
            largura = Inches(10)
            altura = Inches(2)
            caixa_texto = slide.shapes.add_textbox(x, y, largura, altura)
            caixa_texto = slide.shapes.add_textbox(x, y, largura, altura)

            #Formatando o TextBox
            caixa_texto.text = line
            caixa_texto.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            caixa_texto.text_frame.paragraphs[0].font.size = Pt(40)
            caixa_texto.text_frame.paragraphs[0].font.name = "Arial"
            caixa_texto.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        #Salvando o arquivo pptx
        apresentacao.save("VSL.pptx")  
        #Alterando o arquivo par a pasta static
        shutil.move('VSL.pptx', 'static/VSL.pptx')
    except Exception as e:
        print(e)

    return render_template("resultado.html")



# if __name__ == "__main__":
#     # app.run(debug=True)
#     app.run()

serve(app, host="0.0.0.0", port=8080)






















############### Exemplo de login
# @app.route('/login', methods=['POST', 'GET'])
# def login():
#     error = None
#     if request.method == 'POST':
#         if valid_login(request.form['username'],
#                        request.form['password']):
#             return log_the_user_in(request.form['username'])
#         else:
#             error = 'Invalid username/password'
#     # the code below is executed if the request method
#     # was GET or the credentials were invalid
#     return render_template('login.html', error=error)