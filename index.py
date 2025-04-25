import openpyxl
from PIL import Image, ImageDraw, ImageFont

teste_aviso = 'Este certificado possui valor apenas estético e não tem validade acadêmica.'

# exceção criada pois sem ela o arquivo gerava erro de permissão, impedindo o script de rodar
try:
    planilha_arquivo = openpyxl.load_workbook('Teste_emissao_certificado.xlsx')
except PermissionError as e:
    print(f"Erro: Permissão negada ao tentar abrir o arquivo. Detalhes: {e}")
    exit(1)

# aqui escolhe a planilha desejada, mesmo tendo só uma coloca, deu muita dor de cabeça sem hahahaha
try:
    planilha_alunos = planilha_arquivo['Respostas']
except KeyError as e:
    print(f"Erro: A planilha 'Respostas' não foi encontrada no direttorio. Detalhes: {e}")
    exit(1)

# Aqui começa a ler da segunda linha por qaue a primeira é o tidulo da coluna
#pra não pegar o nome da coluna
for indice, linha in enumerate(planilha_alunos.iter_rows(min_row=2)):

    emissao_certificado = linha[0].value
    nome_aluno = linha[1].value
    carga_horaria = linha[5].value
    professor = linha[6].value
    nome_curso = linha[7].value

    # busca a font, lembre-se de ficar no mesmo diretorio se não vai ter que mostrar
    fonte_nome = ImageFont.truetype('./tahomabd.ttf', 90)
    fonte_geral = ImageFont.truetype('./tahoma.ttf', 80)
    fonte_data = ImageFont.truetype('./tahoma.ttf', 55)

    # faz get no  certificado
    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)

    # posiçoes do texto esse foi complicado kkkk
    desenhar.text((1020, 827), nome_aluno, fill='black', font=fonte_nome)
    desenhar.text((1060, 950), nome_curso, fill='black', font=fonte_geral)
    desenhar.text((1435, 1065), professor, fill='black', font=fonte_geral)
    desenhar.text((1480, 1182), str(carga_horaria), fill='black', font=fonte_geral)
    desenhar.text((50, 700), teste_aviso, fill='red', font=fonte_nome)

    # convete a coluna (emissao_certificado) para str afim de pergar somente horas tbm 
    desenhar.text((2220, 1930), emissao_certificado.strftime('%d/%m/%Y'), fill='blue', font=fonte_data)

    # nome final do arquivo com indice pra não criar problemas com nomes repitidos
    image.save(f'./{indice} {nome_aluno} certificado.pdf')
    