from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import yagmail

class EditarCertificado:
    def inicio(self):
        self.leituraDeDados()
        self.enviarEmail()

    # Planilha dos dados:
    def leituraDeDados(self):
        self.df = pd.read_excel('ListaDeAlunos.xlsx')

    # Enviar Email de acordo com os dados a planilha:
    def enviarEmail(self):
        for _, row in self.df.iterrows():
            nome = row['Nome']
            email = row['Email']

            imagemDoCertificado = self.criarCertificado(nome)
            self.conteudoDoEmail(nome, email, imagemDoCertificado)

    # Manipulando imagem do modelo:
    def criarCertificado(self, nome):
        # Abrir a imagem:
        imagem = Image.open('Modelo.png')
        draw = ImageDraw.Draw(imagem)
        fonte = ImageFont.truetype('Questrial.ttf', 39)

        texto = f"Certificamos que {nome}"

        # Centralizar texto de acordo com seu tamanho:
        caixaDoTexto = draw.textbbox((0, 0), texto, font=fonte)
        larguraDoTexto = caixaDoTexto[2] - caixaDoTexto[0]
        larguraDaImagem = imagem.size[0]
        x = (larguraDaImagem - larguraDoTexto) / 2

        # Definindo localização do texto e cor da fonte:
        draw.text((x, 450), texto, font=fonte, fill=(255, 255, 255))

        imagemDoCertificado = f'{nome}_CertificadoICM.png'
        imagem.save(imagemDoCertificado)
        return imagemDoCertificado

    def conteudoDoEmail(self, nome, email, imagemDoCertificado):
        # Email do usuário:
        usuario = yagmail.SMTP(user='SeuEmail', password='SuaSenha')

        # Conteúdo do email:
        titulo = 'Certificado de Participação - Introdução ao Arduino'
        conteudo = f'Olá {nome},\nAqui está o seu certificado de participação.\nAtenciosamente,\nEquipe ICM'

        # Anexar a imagem do certificado
        usuario.send(
            to=email,
            subject=titulo,
            contents=conteudo,
            attachments=imagemDoCertificado
        )

        print(f'Email enviado para {email} com sucesso!')


inicio = EditarCertificado()
inicio.inicio()
