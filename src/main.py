import sys
import os
import base64
import tempfile
import zipfile
import xml.etree.ElementTree as ET
import re
import shutil
import time
from flask import Flask, request, render_template, send_file, jsonify, url_for
from io import BytesIO

# Configuração do Flask
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), 'templates')
app = Flask(__name__, template_folder=TEMPLATE_DIR)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # Limite de 50MB para upload
app.config['UPLOAD_FOLDER'] = (
    '/tmp/xml_processor_uploads'
    if os.environ.get('VERCEL')
    else os.path.join(tempfile.gettempdir(), 'xml_processor_uploads')
)
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'zip'}

# Garante que a pasta de uploads exista
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Funções auxiliares
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def extract_nf_from_xml(xml_content, nf_list):
    """Extrai o número da NF/CTe do conteúdo XML e verifica se termina com algum dos números da lista."""
    try:
        root = ET.fromstring(xml_content)
        
        # Tenta encontrar o número da NF em NFe
        nf_element = root.find(".//nNF")
        if nf_element is not None:
            nf_number = nf_element.text.strip()
            # Verifica se o número da NF termina com algum dos números da lista
            for nf in nf_list:
                if nf_number.endswith(nf):
                    return nf_number, nf
        
        # Tenta encontrar o número do CTe
        ct_element = root.find(".//nCT")
        if ct_element is not None:
            ct_number = ct_element.text.strip()
            # Verifica se o número do CTe termina com algum dos números da lista
            for nf in nf_list:
                if ct_number.endswith(nf):
                    return ct_number, nf
        
        return None, None
    except Exception as e:
        app.logger.error(f"Erro ao processar XML: {str(e)}")
        return None, None

def extract_nf_from_filename(filename, nf_list):
    """Extrai o número da NF do nome do arquivo usando regex e verifica se corresponde a algum da lista."""
    try:
        # Regex para extrair o número da NF do nome do arquivo
        # Padrão: ^[0-9]{25}([0-9]{9})[0-9]+-nfe\.xml$
        match = re.search(r'^[0-9]{25}([0-9]{9})[0-9]+-nfe\.xml$', filename)
        if match:
            extracted_number = match.group(1)
            # Verifica se o número extraído termina com algum dos números da lista
            for nf in nf_list:
                if extracted_number.endswith(nf):
                    return extracted_number, nf
        
        # Tenta outro padrão comum para CTe
        match = re.search(r'^[0-9]{25}([0-9]{9})[0-9]+-cte\.xml$', filename)
        if match:
            extracted_number = match.group(1)
            for nf in nf_list:
                if extracted_number.endswith(nf):
                    return extracted_number, nf
                    
        # Tenta um padrão mais genérico para qualquer número no nome do arquivo
        numbers = re.findall(r'\d+', filename)
        for number in numbers:
            for nf in nf_list:
                if number.endswith(nf):
                    return number, nf
        
        return None, None
    except Exception as e:
        app.logger.error(f"Erro ao extrair número do nome do arquivo: {str(e)}")
        return None, None

def read_excel_in_memory(excel_file):
    """Lê a coluna B de um arquivo Excel diretamente da memória, sem salvar em disco."""
    try:
        # Carrega o arquivo Excel diretamente da memória
        from openpyxl import load_workbook
        
        # Lê o conteúdo do arquivo na memória
        excel_data = excel_file.read()
        
        # Cria um objeto BytesIO para trabalhar com os dados na memória
        excel_io = BytesIO(excel_data)
        
        # Carrega o workbook a partir do BytesIO
        wb = load_workbook(filename=excel_io, read_only=True)
        ws = wb.active
        
        # Extrair valores da coluna B (índice 1), ignorando a primeira linha (cabeçalho)
        nf_list = []
        for row in list(ws.rows)[1:]:  # Pula a primeira linha
            if len(row) >= 2 and row[1].value:  # Coluna B (índice 1)
                nf_list.append(str(row[1].value).strip())
        
        # Fecha o workbook e libera a memória
        wb.close()
        excel_io.close()
        
        return nf_list
    except Exception as e:
        app.logger.error(f"Erro ao ler Excel na memória: {str(e)}")
        
        # Fallback: tentar ler como CSV
        try:
            import csv
            from io import StringIO
            
            # Reseta o ponteiro do arquivo
            excel_file.seek(0)
            
            # Lê o conteúdo como texto
            content = excel_file.read().decode('utf-8')
            
            # Cria um StringIO para trabalhar com o CSV
            csv_io = StringIO(content)
            
            # Lê o CSV
            nf_list = []
            reader = csv.reader(csv_io)
            next(reader)  # Pula o cabeçalho
            for row in reader:
                if len(row) >= 2:  # Garante que há pelo menos 2 colunas
                    nf_list.append(row[1].strip())
            
            return nf_list
        except Exception as csv_error:
            app.logger.error(f"Erro ao ler como CSV: {str(csv_error)}")
            raise ValueError("Não foi possível ler o arquivo Excel. Certifique-se de que é um arquivo .xlsx válido.")

def process_files(excel_data, zip_data):
    """Processa os arquivos Excel e ZIP diretamente da memória, sem salvar em disco."""
    try:
        # Cria diretórios temporários
        temp_dir = tempfile.mkdtemp(prefix="xml_processor_")
        extracted_dir = os.path.join(temp_dir, "extracted")
        selected_dir = os.path.join(temp_dir, "selected")
        os.makedirs(extracted_dir, exist_ok=True)
        os.makedirs(selected_dir, exist_ok=True)
        
        # Lê a planilha Excel diretamente da memória
        excel_io = BytesIO(excel_data)
        nf_list = read_excel_in_memory(excel_io)
        excel_io.close()  # Fecha o BytesIO para liberar memória
        
        if not nf_list:
            raise ValueError("Nenhum número de NF encontrado na coluna B da planilha")
        
        app.logger.info(f"NFs lidas da planilha ({len(nf_list)}): {nf_list}")
        
        # Extrai o ZIP diretamente da memória
        zip_io = BytesIO(zip_data)
        with zipfile.ZipFile(zip_io, 'r') as zip_ref:
            zip_ref.extractall(extracted_dir)
        zip_io.close()  # Fecha o BytesIO para liberar memória
        
        # Processa os XMLs recursivamente
        xml_count = 0
        matches = {}  # Para registrar quais NFs foram encontradas
        
        for root, _, files in os.walk(extracted_dir):
            for file in files:
                if file.lower().endswith('.xml'):
                    xml_path = os.path.join(root, file)
                    
                    # Primeiro tenta extrair o número do nome do arquivo
                    file_number, matched_nf = extract_nf_from_filename(file, nf_list)
                    
                    # Se não conseguiu pelo nome, tenta pelo conteúdo
                    if not file_number:
                        try:
                            with open(xml_path, 'r', encoding='utf-8') as xml_file:
                                xml_content = xml_file.read()
                                file_number, matched_nf = extract_nf_from_xml(xml_content, nf_list)
                        except Exception as e:
                            app.logger.error(f"Erro ao ler o arquivo {file}: {str(e)}")
                            continue
                    
                    # Se encontrou correspondência, copia o arquivo
                    if file_number and matched_nf:
                        # Copia o XML para a pasta de selecionados
                        shutil.copy2(xml_path, os.path.join(selected_dir, file))
                        app.logger.info(f"Copiado: {file} (NF completa: {file_number}, correspondeu com: {matched_nf})")
                        xml_count += 1
                        matches[matched_nf] = file_number
        
        # Registra quais NFs foram encontradas e quais não foram
        found_nfs = list(matches.keys())
        missing_nfs = [nf for nf in nf_list if nf not in found_nfs]
        
        if missing_nfs:
            app.logger.warning(f"NFs não encontradas: {missing_nfs}")
        
        if xml_count == 0:
            raise ValueError("Nenhum XML correspondente encontrado. Verifique se os números de NF na planilha correspondem aos números nos XMLs.")
        
        # Cria o ZIP final
        memory_file = BytesIO()
        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(selected_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    zipf.write(file_path, arcname=file)
        
        # Limpa os diretórios temporários
        try:
            shutil.rmtree(temp_dir)
        except Exception as e:
            app.logger.error(f"Erro ao limpar diretório temporário: {str(e)}")
            # Continua mesmo se não conseguir limpar
        
        # Prepara o arquivo para download
        memory_file.seek(0)
        return memory_file, xml_count, matches
    
    except Exception as e:
        # Limpa os diretórios temporários em caso de erro
        if 'temp_dir' in locals():
            try:
                shutil.rmtree(temp_dir)
            except:
                pass  # Ignora erros ao limpar
        raise e

# Rotas
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        # Verifica se os arquivos foram enviados
        if 'excel_file' not in request.files or 'zip_file' not in request.files:
            return jsonify({'error': 'Ambos os arquivos (Excel e ZIP) são obrigatórios'}), 400
        
        excel_file = request.files['excel_file']
        zip_file = request.files['zip_file']
        
        # Verifica se os arquivos têm nomes válidos
        if excel_file.filename == '' or zip_file.filename == '':
            return jsonify({'error': 'Nenhum arquivo selecionado'}), 400
        
        # Verifica se os arquivos têm extensões permitidas
        if not (allowed_file(excel_file.filename) and allowed_file(zip_file.filename)):
            return jsonify({'error': 'Tipo de arquivo não permitido. Use .xlsx para Excel e .zip para o arquivo ZIP'}), 400
        
        # Lê os arquivos diretamente na memória
        excel_data = excel_file.read()
        zip_data = zip_file.read()
        
        # Processa os arquivos diretamente da memória
        result_file, xml_count, matches = process_files(excel_data, zip_data)
        zip_bytes = result_file.getvalue()

        # Prepara mensagem detalhada
        match_details = []
        for nf, full_number in matches.items():
            match_details.append(f"NF {nf} → {full_number}")

        match_message = "<br>".join(match_details)

        response = {
            'success': True,
            'message': f'{xml_count} XMLs foram processados e compactados com sucesso!',
            'details': match_message,
            'download_filename': 'entrada.zip',
            # Base64 funciona no Vercel serverless (download na mesma resposta)
            'file_base64': base64.b64encode(zip_bytes).decode('ascii'),
        }

        # Fallback local: salva em disco para rota /download
        if not os.environ.get('VERCEL'):
            result_path = os.path.join(app.config['UPLOAD_FOLDER'], 'entrada.zip')
            max_attempts = 5
            for attempt in range(max_attempts):
                try:
                    with open(result_path, 'wb') as f:
                        f.write(zip_bytes)
                    break
                except PermissionError:
                    if attempt < max_attempts - 1:
                        time.sleep(0.5)
                    else:
                        result_path = os.path.join(
                            app.config['UPLOAD_FOLDER'],
                            f'entrada_{int(time.time())}.zip',
                        )
                        with open(result_path, 'wb') as f:
                            f.write(zip_bytes)
            response['download_url'] = url_for('download_result', timestamp=int(time.time()))

        return jsonify(response)
    
    except Exception as e:
        app.logger.error(f"Erro durante o processamento: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/download', methods=['GET'])
def download_result():
    # Usa um timestamp para evitar cache do navegador
    timestamp = request.args.get('timestamp', '')
    
    result_path = os.path.join(app.config['UPLOAD_FOLDER'], 'entrada.zip')
    if not os.path.exists(result_path):
        # Tenta encontrar um arquivo alternativo
        files = [f for f in os.listdir(app.config['UPLOAD_FOLDER']) if f.startswith('entrada_') and f.endswith('.zip')]
        if files:
            # Usa o arquivo mais recente
            result_path = os.path.join(app.config['UPLOAD_FOLDER'], sorted(files)[-1])
        else:
            return jsonify({'error': 'Arquivo não encontrado. Faça o upload dos arquivos primeiro.'}), 404
    
    # Tenta abrir o arquivo com várias tentativas
    max_attempts = 5
    for attempt in range(max_attempts):
        try:
            return send_file(result_path, as_attachment=True, download_name='entrada.zip')
        except PermissionError:
            if attempt < max_attempts - 1:
                time.sleep(0.5)  # Espera um pouco antes de tentar novamente
            else:
                return jsonify({'error': 'Não foi possível acessar o arquivo. Tente novamente em alguns segundos.'}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')
