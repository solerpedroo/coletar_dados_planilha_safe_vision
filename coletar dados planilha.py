import face_recognition
import cv2
import openpyxl
import os

# Caminho da pasta onde estão as fotos dos motoristas
fotos_path = r'C:\Users\Pichau\Desktop\VISAO_COMPUTACIONAL\Pessoas'
# Caminho do arquivo Excel
file_path = r'C:\Users\Pichau\Desktop\VISAO_COMPUTACIONAL\Dados coletados.xlsx'


# Função para obter as informações do motorista com base no ID da foto reconhecida
def obter_informacoes_motorista(id_motorista_com_extensao, caminho_arquivo_excel):
    # Abre a planilha
    wb = openpyxl.load_workbook(caminho_arquivo_excel)
    sheet = wb['Sheet1']

    # Cabeçalhos e colunas relevantes
    colunas = {
        'C': 'Nome',
        'D': 'Sobrenome',
        'E': 'Idade',
        'F': 'Medicamentos',
        'G': 'Doenças',
        'I': 'Marca',
        'J': 'Modelo',
        'K': 'Ano',
        'L': 'Cor',
        'M': 'Placa'
    }

    # Encontra o ID na coluna A e coleta as informações relevantes
    for row in sheet.iter_rows(min_row=3, values_only=True):  # Ignora a linha 1 e 2
        if row[
            0] == id_motorista_com_extensao:  # Verifica se o ID da coluna A (índice 0) é o ID do motorista com extensão
            print(f"Informações para o motorista '{id_motorista_com_extensao}':")
            for col, header in colunas.items():
                col_idx = openpyxl.utils.column_index_from_string(col) - 1  # Converte letra da coluna em índice
                print(f'{header}: {row[col_idx]}')  # Imprime o cabeçalho e o valor correspondente
            print('-' * 40)  # Separador para facilitar a leitura
            break
    else:
        print(f"ID '{id_motorista_com_extensao}' não encontrado na planilha.")


# Função para carregar fotos e codificar faces dos motoristas
def carregar_fotos_motoristas(fotos_path):
    motoristas_encodings = []
    motoristas_ids = []

    # Itera por todas as fotos na pasta
    for file_name in os.listdir(fotos_path):
        if file_name.endswith('.jpg') or file_name.endswith('.png'):
            path_foto = os.path.join(fotos_path, file_name)
            imagem = face_recognition.load_image_file(path_foto)
            encoding = face_recognition.face_encodings(imagem)

            if encoding:  # Verifica se foi possível codificar a face
                motoristas_encodings.append(encoding[0])
                motoristas_ids.append(file_name)  # Inclui o nome do arquivo completo (com extensão) como ID

    return motoristas_encodings, motoristas_ids


# Função principal para reconhecimento facial em tempo real com a câmera
def reconhecimento_facial_tempo_real(motoristas_encodings, motoristas_ids, caminho_arquivo_excel):
    video_capture = cv2.VideoCapture(0)  # Acessa a câmera

    while True:
        ret, frame = video_capture.read()
        rgb_frame = frame[:, :, ::-1]  # Converte o frame de BGR para RGB

        # Localiza as faces no frame atual
        faces_localizadas = face_recognition.face_locations(rgb_frame)
        faces_encodings = face_recognition.face_encodings(rgb_frame, faces_localizadas)

        # Compara cada face detectada com as faces dos motoristas
        for face_encoding in faces_encodings:
            matches = face_recognition.compare_faces(motoristas_encodings, face_encoding)
            if True in matches:
                # Pega o ID do motorista correspondente
                match_index = matches.index(True)
                id_motorista_com_extensao = motoristas_ids[match_index]  # Nome da foto com extensão .jpg

                # Obtém as informações do motorista
                obter_informacoes_motorista(id_motorista_com_extensao, caminho_arquivo_excel)

        # Exibe o feed da câmera em tempo real
        cv2.imshow('Reconhecimento Facial', frame)

        # Pressione 'q' para sair
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    video_capture.release()
    cv2.destroyAllWindows()


# Carrega as fotos e codificações dos motoristas
motoristas_encodings, motoristas_ids = carregar_fotos_motoristas(fotos_path)

# Inicia o reconhecimento facial em tempo real
reconhecimento_facial_tempo_real(motoristas_encodings, motoristas_ids, file_path)
