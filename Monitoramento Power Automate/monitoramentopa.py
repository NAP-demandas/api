import msal
import requests
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from apscheduler.schedulers.blocking import BlockingScheduler
import os

# Configurações (protegendo dados sensíveis)
CONFIG = {
    'client_id': 'daa404e6-737d-42fa-9941-d007fc19463a',
    'authority': 'https://login.microsoftonline.com/c2c7d69b-25e1-4dd3-bb7c-cec85a3e1913',
    'scope': ['https://service.flow.microsoft.com/.default'],
    'username': 'automacaoengenharia@vittaresidencial.com.br',
    'password': 'Bild@2024',
    'environment_name': 'Default-c2c7d69b-25e1-4dd3-bb7c-cec85a3e1913',
    'smtp': {
        'server': 'smtp.office365.com',
        'port': 587,
        'sender': 'automacaoengenharia@vittaresidencial.com.br',
        'password': 'Bild@2024',
        'recipients': [
    'samuel.carvalho@bild.com.br',
    
]
    }
}

def get_token():
    """Obtém token de acesso com tratamento aprimorado de erros"""
    try:
        app = msal.PublicClientApplication(
            CONFIG['client_id'],
            authority=CONFIG['authority']
        )
        
        # Primeiro tenta obter token silenciosamente (se houver cache)
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(
                scopes=CONFIG['scope'],
                account=accounts[0]
            )
            if result and "access_token" in result:
                return result['access_token']
        
        # Se não conseguir silenciosamente, tenta com username/password
        result = app.acquire_token_by_username_password(
            username=CONFIG['username'],
            password=CONFIG['password'],
            scopes=CONFIG['scope']
        )
        
        if "access_token" in result:
            return result['access_token']
        else:
            error = result.get('error_description', 'Erro desconhecido')
            raise Exception(f"Falha na autenticação: {error}")
            
    except requests.exceptions.ConnectionError as e:
        raise Exception(f"Não foi possível conectar ao servidor de autenticação. Verifique sua conexão de rede e a URL do servidor. Detalhes: {str(e)}")
    except Exception as e:
        raise Exception(f"Erro ao obter token: {str(e)}")

def get_last_flow_run(token, flow_id):
    """Obtém a última execução de um fluxo com tratamento de erros"""
    try:
        headers = {'Authorization': f'Bearer {token}'}
        url = f"https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/{CONFIG['environment_name']}/flows/{flow_id}/runs?" \
              f"api-version=2016-11-01&$top=1&$orderby=properties/startTime desc"
        
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        runs = response.json().get('value', [])
        return runs[0] if runs else None
    except requests.exceptions.RequestException as e:
        print(f"⚠️ Atenção: Erro ao obter execução do fluxo {flow_id[:8]}... - {str(e)}")
        return None

def format_duration(start_time, end_time):
    """Formata a duração no padrão desejado (igual ao relatório correto)"""
    try:
        start = datetime.fromisoformat(start_time.replace('Z', '+00:00'))
        end = datetime.fromisoformat(end_time.replace('Z', '+00:00'))
        duration = end - start
        
        total_seconds = duration.total_seconds()
        
        # Ajuste para retornar exatamente como no relatório correto
        if total_seconds < 1:
            return f"{int(duration.microseconds/1000)} ms"
        elif total_seconds < 60:
            return f"{int(total_seconds)} segundos"
        elif total_seconds < 3600:
            minutes = int(total_seconds // 60)
            return f"{minutes} minutos"
        else:
            hours = int(total_seconds // 3600)
            minutes = int((total_seconds % 3600) // 60)
            return f"{hours}h {minutes}m"
    except:
        return "N/A"

def format_date(datetime_str):
    """Formata a data para dd/mm/aaaa (padrão brasileiro)"""
    try:
        dt = datetime.fromisoformat(datetime_str.replace('Z', '+00:00'))
        return dt.strftime('%d/%m/%Y')
    except:
        return datetime_str

def get_flows_report(token):
    """Gera o relatório de fluxos com tratamento robusto"""
    try:
        headers = {'Authorization': f'Bearer {token}'}
        flows_url = f"https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/{CONFIG['environment_name']}/flows?api-version=2016-11-01"
        
        response = requests.get(flows_url, headers=headers, timeout=30)
        response.raise_for_status()
        flows = response.json().get('value', [])
        
        report_data = []
        for flow in flows:
            try:
                flow_id = flow['name']
                last_run = get_last_flow_run(token, flow_id)
                
                if last_run:
                    start_time = last_run['properties']['startTime']
                    
                    # Mantenha os nomes das chaves consistentes
                    report_data.append({
                        'Nome do Fluxo': flow['properties']['displayName'],
                        'Data': format_date(start_time),  # Agora usando 'Data' consistentemente
                        'Última Execução': format_date(start_time),  # Mantido para compatibilidade
                        'Dias sem Executar': get_days_since_last_run(start_time),
                        'Duração': format_duration(start_time, last_run['properties']['endTime']),
                        'Status': last_run['properties']['status'],
                        'Erro': last_run['properties'].get('error', {}).get('message', '')
                    })
                    
            except Exception as e:
                print(f"⚠️ Erro ao processar fluxo {flow.get('name', '')[:8]}... - {str(e)}")
                continue
        
        return sorted(report_data, key=lambda x: x['Nome do Fluxo'].upper())
    
    except Exception as e:
        raise Exception(f"Erro ao gerar relatório: {str(e)}")

def get_days_since_last_run(last_run_date):
    """Calcula dias desde a última execução de forma robusta"""
    try:
        # Aceita tanto o formato ISO quanto o formato brasileiro já formatado
        if isinstance(last_run_date, str) and '/' in last_run_date:
            day, month, year = map(int, last_run_date.split('/'))
            last_run = datetime(year, month, day).date()
        else:
            last_run = datetime.fromisoformat(last_run_date.replace('Z', '+00:00')).date()
        
        today = datetime.now().date()
        days_diff = (today - last_run).days
        
        # Retorna 0 se for negativo ou a mesma data
        return max(0, days_diff)
    except:
        return "N/A"

def generate_email_body(report_data):
    """Gera o corpo do e-mail com visual profissional melhorado"""
    status_icons = {
        'Succeeded': '✅',
        'Failed': '❌',
        'Running': '🔄'
    }
    
    # Ordenar os fluxos por nome (A-Z)
    sorted_report_data = sorted(report_data, key=lambda x: x['Nome do Fluxo'].upper())
    
    html = f"""
    <html>
    <head>
        <style>
            body {{
                font-family: 'Segoe UI', Arial, sans-serif;
                margin: 20px;
                color: #333;
                line-height: 1.6;
            }}
            .header {{
                color: #2F5496;
                border-bottom: 2px solid #2F5496;
                padding-bottom: 10px;
                margin-bottom: 20px;
            }}
            .report-date {{
                color: #666;
                font-size: 0.9em;
                margin-bottom: 20px;
            }}
            table {{
                border-collapse: collapse;
                width: 100%;
                margin-top: 20px;
                box-shadow: 0 1px 3px rgba(0,0,0,0.1);
                font-size: 0.9em;
            }}
            th {{
                background-color: #2F5496;
                color: white;
                padding: 12px 15px;
                text-align: left;
                font-weight: 600;
            }}
            td {{
                padding: 10px 15px;
                border-bottom: 1px solid #e0e0e0;
                vertical-align: top;
            }}
            tr:hover {{
                background-color: #f8f8f8;
            }}
            .success {{
                color: #107C10;
            }}
            .failed {{
                color: #D83B01;
            }}
            .running {{
                color: #F2C811;
            }}
            .error-message {{
                color: #A4262C;
                font-size: 0.85em;
                padding: 8px 15px;
                background-color: #fde7e9;
                border-left: 3px solid #A4262C;
            }}
            .warning {{
                color: #FF8C00;
                font-weight: 500;
            }}
            .footer {{
                margin-top: 30px;
                font-size: 0.85em;
                color: #666;
                border-top: 1px solid #eee;
                padding-top: 15px;
            }}
            .category {{
                font-weight: 600;
                color: #2F5496;
                margin-top: 15px;
                margin-bottom: 5px;
            }}
        </style>
    </head>
    <body>
        <div class="header">
            <h2>Monitoramento de Fluxos Power Automate</h2>
            <div class="report-date">
                Relatório gerado em: {datetime.now().strftime('%d/%m/%Y às %H:%M')}<br>
                Ambiente: {CONFIG['environment_name']}
            </div>
        </div>
        
        <table>
            <thead>
                <tr>
                    <th style="width: 30%;">Nome do Fluxo</th>
                    <th style="width: 15%;">Última Execução</th>
                    <th style="width: 15%;">Dias sem Executar</th>
                    <th style="width: 15%;">Duração</th>
                    <th style="width: 25%;">Status</th>
                </tr>
            </thead>
            <tbody>
    """
    
    # Agrupar por categoria (prefixo antes do |)
    categories = {}
    for item in sorted_report_data:
        # Extrai a categoria do nome do fluxo (parte antes do |)
        if '|' in item['Nome do Fluxo']:
            category, name = [part.strip() for part in item['Nome do Fluxo'].split('|', 1)]
        else:
            category = 'Outros'
            name = item['Nome do Fluxo']
        
        if category not in categories:
            categories[category] = []
        categories[category].append({**item, 'Nome Real': name})
    
    # Ordenar categorias alfabeticamente
    sorted_categories = sorted(categories.keys())
    
    for category in sorted_categories:
        html += f"""
            <tr><td colspan="5" class="category">{category}</td></tr>
        """
        
        for item in categories[category]:
            status_class = item['Status'].lower()
            icon = status_icons.get(item['Status'], '')
            days_since_run = item['Dias sem Executar']
            
            # Aplicar classe warning se não executar há mais de 7 dias
            days_class = "warning" if isinstance(days_since_run, int) and days_since_run > 7 else ""
            
            html += f"""
                <tr>
                    <td>{item['Nome Real']}</td>
                    <td>{item['Última Execução']}</td>
                    <td class="{days_class}">{days_since_run}</td>
                    <td>{item['Duração']}</td>
                    <td class="{status_class}">{icon} {item['Status']}</td>
                </tr>
            """
            
            if item['Status'] == 'Failed' and item['Erro']:
                html += f"""
                    <tr>
                        <td colspan="5" class="error-message">
                            <strong>Erro:</strong> {item['Erro']}
                        </td>
                    </tr>
                """
    
    html += f"""
            </tbody>
        </table>
        
        <div class="footer">
            <p>Este é um relatório automático. Favor não responder este e-mail.</p>
            <p>Total de fluxos monitorados: {len(report_data)}</p>
        </div>
    </body>
    </html>
    """
    
    return html

def send_email(subject, body):
    """Envia o e-mail com tratamento de erros"""
    try:
        msg = MIMEMultipart()
        msg['From'] = CONFIG['smtp']['sender']
        msg['To'] = ", ".join(CONFIG['smtp']['recipients'])
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html'))
        
        with smtplib.SMTP(CONFIG['smtp']['server'], CONFIG['smtp']['port']) as server:
            server.starttls()
            server.login(CONFIG['smtp']['sender'], CONFIG['smtp']['password'])
            server.send_message(msg)
        
        print(f"✅ E-mail enviado com sucesso para {len(CONFIG['smtp']['recipients'])} destinatário(s)")
    except Exception as e:
        print(f"❌ Falha crítica ao enviar e-mail: {str(e)}")
        raise

def generate_and_send_report():
    """Função principal para gerar e enviar o relatório"""
    print(f"\n{datetime.now().strftime('%d/%m/%Y %H:%M:%S')} - Iniciando processo de monitoramento...")
    
    try:
        # 1. Autenticação
        print("🔑 Obtendo token de acesso...")
        token = get_token()
        
        # 2. Coletar dados
        print("📊 Coletando dados dos fluxos...")
        report_data = get_flows_report(token)
        
        if not report_data:
            raise Exception("Nenhum dado de fluxo foi encontrado para gerar o relatório")
        
        # Validação dos dados
        required_keys = ['Nome do Fluxo', 'Última Execução', 'Dias sem Executar', 'Duração', 'Status']
        for item in report_data:
            for key in required_keys:
                if key not in item:
                    raise Exception(f"Chave '{key}' não encontrada nos dados do relatório")
        
        # 3. Gerar relatório
        print("✍️ Gerando relatório...")
        email_body = generate_email_body(report_data)
        
        # 4. Enviar e-mail
        print("📨 Enviando e-mail...")
        send_email("Fluxos Power Automate | Monitoramento Diário", email_body)
        
        print("✅ Processo concluído com sucesso!")
    
    except Exception as e:
        error_msg = f"❌ Erro no processo de monitoramento: {str(e)}"
        print(error_msg)
        send_email("ERRO - Monitoramento Power Automate", error_msg)

if __name__ == "__main__":
    
    generate_and_send_report()  # Executa o script uma vez