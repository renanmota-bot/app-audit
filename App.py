import io
import re
import json
import os
import datetime
import urllib.request
import pandas as pd
import streamlit as st
import google.generativeai as genai
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from PyPDF2 import PdfReader

# -----------------------------------------------------------------------------
# CONFIGURAÇÕES DE API E RECURSOS DO ECOBOT
# -----------------------------------------------------------------------------
API_KEY_PADRAO = "AIzaSyAy7KaL0IHOKnwGbmAfxE_NqIVq9LY9AEU"
URL_ECOBOT = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQmrGnEZ3y-nutzMBkki7MfLa9SzSUYaSe1bu0U5ySrJg&s=10"

# -----------------------------------------------------------------------------
# CONFIGURAÇÃO DE PÁGINA E CSS (BLINDAGEM LIGHT MODE NO MOBILE)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Estudos Ambientais - EcoBot",
    page_icon="🌱",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown(f"""
    <style>
        :root, html, body, .stApp, [data-testid="stAppViewContainer"], [data-theme="dark"], [data-theme="light"] {{
            --background-color: #F8FAFC !important;
            --secondary-background-color: #FFFFFF !important;
            --text-color: #0F172A !important;
            --primary-color: #10B981 !important;
            color-scheme: light !important;
        }}

        .stApp {{
            background-color: #F8FAFC !important;
            color: #0F172A !important;
        }}

        #MainMenu {{visibility: hidden;}}
        footer {{visibility: hidden;}}
        
        .block-container {{
            padding-top: 1.5rem !important;
            padding-bottom: 5rem !important;
            max-width: 1000px !important;
        }}

        /* PALCO DO TALKING TOM (CENTRALIZADO) */
        .talking-stage {{
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            margin: 15px 0;
            position: relative;
        }}

        /* BALÃO DE DIÁLOGO DO TALKING TOM */
        .talking-speech-bubble {{
            background: #FFFFFF !important;
            border: 3px solid #10B981 !important;
            border-radius: 20px;
            padding: 16px 22px;
            max-width: 550px;
            box-shadow: 0 10px 25px rgba(16, 185, 129, 0.15);
            font-size: 1.05rem;
            color: #0F172A !important;
            position: relative;
            line-height: 1.5;
            font-weight: 600;
            text-align: center;
            margin-bottom: 20px;
        }}

        .talking-speech-bubble::after {{
            content: '';
            position: absolute;
            bottom: -14px;
            left: 50%;
            transform: translateX(-50%);
            border-width: 14px 14px 0;
            border-style: solid;
            border-color: #10B981 transparent;
            display: block;
            width: 0;
        }}

        /* PERSONAGEM GRANDE DENTRO DO PALCO */
        .talking-avatar-frame {{
            width: 210px;
            height: 210px;
            border-radius: 50%;
            border: 5px solid #10B981;
            box-shadow: 0 12px 30px rgba(0, 0, 0, 0.12);
            background: #FFFFFF !important;
            overflow: hidden;
            display: flex;
            align-items: center;
            justify-content: center;
            animation: floatCharacter 3s ease-in-out infinite;
        }}

        .talking-avatar-img {{
            width: 100%;
            height: 100%;
            object-fit: cover;
        }}

        @keyframes floatCharacter {{
            0% {{ transform: translateY(0px) rotate(0deg); }}
            50% {{ transform: translateY(-8px) rotate(2deg); }}
            100% {{ transform: translateY(0px) rotate(0deg); }}
        }}

        /* ANIMAÇÃO QUANDO ESTÁ FALANDO */
        .ecobot-talking .talking-avatar-frame {{
            animation: talkMouth 0.22s infinite alternate !important;
            border-color: #34D399 !important;
            box-shadow: 0 0 35px rgba(52, 211, 153, 0.8) !important;
        }}

        @keyframes talkMouth {{
            0% {{ transform: scale(1) translateY(-3px); filter: brightness(1); }}
            100% {{ transform: scale(1.06) translateY(-7px); filter: brightness(1.1); }}
        }}

        /* 🛠️ CORREÇÃO DE LEITURA DAS MENSAGENS NO CHAT (MOBILE & DESKTOP) */
        .stChatMessage, [data-testid="stChatMessage"] {{
            background-color: #FFFFFF !important;
            border: 1.5px solid #E2E8F0 !important;
            border-radius: 14px !important;
            padding: 15px !important;
            margin-bottom: 10px !important;
            color: #0F172A !important;
            box-shadow: 0 2px 5px rgba(0,0,0,0.03) !important;
        }}

        .stChatMessage p, .stChatMessage span, .stChatMessage div, [data-testid="stChatMessage"] * {{
            color: #0F172A !important;
            -webkit-text-fill-color: #0F172A !important;
            font-weight: 500 !important;
        }}

        /* 🛠️ CORREÇÃO DA CAIXA DE DIGITAÇÃO DO CHAT (CHAT INPUT) */
        [data-testid="stChatInput"], [data-baseweb="input"] {{
            background-color: #FFFFFF !important;
            border: 1.5px solid #10B981 !important;
            border-radius: 12px !important;
        }}

        [data-testid="stChatInput"] textarea, [data-testid="stChatInput"] input {{
            color: #0F172A !important;
            background-color: #FFFFFF !important;
            -webkit-text-fill-color: #0F172A !important;
            font-weight: 600 !important;
        }}

        /* SIDEBAR */
        section[data-testid="stSidebar"] {{
            background-color: #FFFFFF !important;
            border-right: 1px solid #E2E8F0 !important;
        }}

        section[data-testid="stSidebar"] * {{
            color: #0F172A !important;
        }}

        .stButton button {{
            width: 100% !important;
            font-weight: bold !important;
            border-radius: 10px !important;
            padding: 10px !important;
            transition: all 0.2s !important;
            border: none !important;
        }}

        .btn-gerar button {{ background-color: #0284C7 !important; color: white !important; -webkit-text-fill-color: white !important; }}
        .btn-lei button {{ background-color: #059669 !important; color: white !important; -webkit-text-fill-color: white !important; }}
        .btn-flash button {{ background-color: #D97706 !important; color: white !important; -webkit-text-fill-color: white !important; }}
        .btn-pdf button {{ background-color: #4F46E5 !important; color: white !important; -webkit-text-fill-color: white !important; }}

        .score-card {{
            background-color: #ECFDF5 !important;
            border: 2px solid #059669 !important;
            border-radius: 16px;
            padding: 20px;
            text-align: center;
            color: #065F46 !important;
            margin: 15px 0;
        }}
    </style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# MATRIZ DE ESTUDOS AMBIENTAIS
# -----------------------------------------------------------------------------
MATRIZ_ESTUDOS = {
    "Volume 1: Básico e Legislação Estruturante": {
        "Módulo 1: Língua Portuguesa": ["1.1 Interpretação e Recursos Discursivos", "1.2 Coesão e Coerência", "1.3 Sintaxe, Concordância e Crase", "1.4 Pontuação"],
        "Módulo 2: Inglês Técnico": ["2.1 Vocabulário Offshore", "2.2 Skimming e Scanning", "2.3 Falsos Cognatos", "2.4 Tradução Técnica"],
        "Módulo 3: PNMA (Lei nº 6.938/81)": ["3.1 Objetivos e Princípios", "3.2 SISNAMA", "3.3 Instrumentos e Responsabilidade"],
        "Módulo 4: PNRS (Lei nº 12.305/10)": ["4.1 Resíduo vs. Rejeito", "4.2 Logística Reversa", "4.3 PGRS e PGRSS"]
    },
    "Volume 2: Leis Específicas e Licenciamento": {
        "Módulo 1: Lei do Óleo (9.966/00)": ["1.1 Prevenção em Águas", "1.2 PEI em Terminais"],
        "Módulo 2: Crimes Ambientais (9.605/98)": ["2.1 Sanções Penais e Administrativas", "2.2 Crimes contra a Administração"],
        "Módulo 3: Recursos Hídricos (9.433/97)": ["3.1 Instrumentos da PNRH", "3.2 Enquadramento e Comitês"],
        "Módulo 4: Licenciamento (CONAMA)": ["4.1 Etapas (LP, LI e LO)", "4.2 EIA/RIMA"]
    },
    "Volume 3: Gestão Técnica e ESG": {
        "Módulo 1: Efluentes": ["1.1 Padrões de Lançamento (CONAMA 357 e 430)"],
        "Módulo 2: Resíduos": ["2.1 NBR 10.004", "2.2 MTR e SINIR"],
        "Módulo 3: ISO e ESG": ["3.1 ISO 14001:2015", "3.2 Indicadores ESG", "3.3 Planos de Emergência"]
    }
}

# -----------------------------------------------------------------------------
# SÍNTESE DE VOZ ROBÓTICA FOFA
# -----------------------------------------------------------------------------
def tocar_voz_robotica(texto, idioma='pt-BR'):
    texto_limpo = re.sub(r'[\*\#\-\_\n\r\'\"]', ' ', texto)[:400]
    js_code = f"""
    <script>
        window.speechSynthesis.cancel();
        var msg = new SpeechSynthesisUtterance("{texto_limpo}");
        msg.rate = 1.15;
        msg.pitch = 1.6;
        msg.lang = '{idioma}';
        window.speechSynthesis.speak(msg);
    </script>
    """
    st.components.v1.html(js_code, height=0)

# -----------------------------------------------------------------------------
# BUSCA DE CLIMA EM TEXTO PURO
# -----------------------------------------------------------------------------
def buscar_previsao_tempo(cidade="Sao Paulo"):
    try:
        url = f"https://wttr.in/{urllib.parse.quote(cidade)}?format=3"
        req = urllib.request.Request(url, headers={'User-Agent': 'curl/7.68.0'})
        with urllib.request.urlopen(req) as response:
            clima = response.read().decode('utf-8').strip()
            return re.sub(r'<[^>]*>', '', clima)
    except Exception:
        return f"{cidade.title()}: Ensolarado com 25°C."

# -----------------------------------------------------------------------------
# MOTOR DA IA
# -----------------------------------------------------------------------------
def chamar_ia(prompt, json_mode=False):
    config = genai.GenerationConfig(
        response_mime_type="application/json" if json_mode else None,
        max_output_tokens=2048,
        temperature=0.4
    )
    
    prompt_direto = f"Responda diretamente ao usuário em português de forma clara e objetiva sem rascunhos.\n\n{prompt}"

    modelos_disponiveis = []
    try:
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                modelos_disponiveis.append(m.name)
    except Exception:
        modelos_disponiveis = ["models/gemini-1.5-flash", "models/gemini-1.5-pro"]

    modelos_disponiveis.sort(key=lambda x: 0 if 'flash' in x else 1)

    ultimo_erro = ""
    for m_nome in modelos_disponiveis:
        try:
            model = genai.GenerativeModel(m_nome)
            res = model.generate_content(prompt_direto, generation_config=config)
            return res.text
        except Exception as e:
            ultimo_erro = str(e)
            continue
            
    raise Exception(f"Erro na conexão com a API do Gemini: {ultimo_erro}")

# -----------------------------------------------------------------------------
# UTILITÁRIOS
# -----------------------------------------------------------------------------
def limpar_json(texto):
    try:
        match_lista = re.search(r'\[.*\]', texto, re.DOTALL)
        if match_lista: return match_lista.group(0)
        match_obj = re.search(r'\{.*\}', texto, re.DOTALL)
        if match_obj: return match_obj.group(0)
        return texto.strip()
    except:
        return texto.strip()

def remover_emojis_pdf(texto):
    return re.sub(r'[^\x00-\x7FáéíóúâêîôûãõçÁÉÍÓÚÂÊÎÔÛÃÕÇ.,;:!?()\[\]{}\-_\'\"\s]', '', texto)

def gerar_pdf(topico, texto_resposta):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story, styles = [], getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=14, textColor=colors.HexColor('#0F172A'), spaceAfter=15)
    body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontSize=10, leading=14, textColor=colors.HexColor('#334155'))

    def md_para_html(txt): return re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', txt.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

    story.append(Paragraph(f"<b>Apostila VIP — {md_para_html(remover_emojis_pdf(topico))}</b>", title_style))
    story.append(Spacer(1, 10))
    for linha in texto_resposta.split('\n'):
        if linha.strip():
            linha_html = md_para_html(remover_emojis_pdf(linha))
            try: story.append(Paragraph(linha_html, body_style))
            except: story.append(Paragraph(re.sub(r'<[^>]*>', '', linha_html), body_style))
            story.append(Spacer(1, 5))
    doc.build(story)
    buffer.seek(0)
    return buffer

def ler_pdf_usuario(arquivo):
    leitor = PdfReader(arquivo)
    texto = ""
    for pagina in leitor.pages: texto += pagina.extract_text() + "\n"
    return texto

def salvar_historico(nota, total, topico):
    arquivo = "historico_simulados.json"
    dados = []
    if os.path.exists(arquivo):
        with open(arquivo, "r") as f: dados = json.load(f)
    dados.append({
        "data": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "topico": topico, "nota": nota, "total": total,
        "porcentagem": round((nota/total)*100, 2)
    })
    with open(arquivo, "w") as f: json.dump(dados, f)

def carregar_historico():
    if os.path.exists("historico_simulados.json"):
        with open("historico_simulados.json", "r") as f: return json.load(f)
    return []

def extrair_historico(mensagens, max_msgs=4):
    return "\n".join([f"{'Usuário' if m['role']=='user' else 'EcoBot'}: {m['content']}" for m in mensagens[-max_msgs:-1]]) if len(mensagens) > 1 else ""

# -----------------------------------------------------------------------------
# SIDEBAR NAVEGAÇÃO
# -----------------------------------------------------------------------------
st.sidebar.markdown("### 🌱 **Estudos Ambientais**")
modo = st.sidebar.radio("Selecione a Tela:", [
    "🤖 EcoBot (Talking Tom)",
    "📖 Prof. Ambiental & Direito", 
    "🇬🇧 Prof. Inglês Técnico",
    "📝 Simulado Hard com Cronômetro",
    "📊 Meu Desempenho",
    "📚 Chat com PDF / Edital"
])

if 'modo_atual' not in st.session_state: st.session_state.modo_atual = modo
if st.session_state.modo_atual != modo:
    st.session_state.messages = []
    st.session_state.modo_atual = modo

st.sidebar.markdown("---")
st.sidebar.markdown("### 🔑 **Autenticação**")
user_key = st.sidebar.text_input("Sua API Key:", value=API_KEY_PADRAO, type="password")
if user_key: genai.configure(api_key=user_key)
else: st.stop()

if 'simulado_dados' not in st.session_state: st.session_state.simulado_dados = None
if 'simulado_corrigido' not in st.session_state: st.session_state.simulado_corrigido = False
if 'simulado_respostas' not in st.session_state: st.session_state.simulado_respostas = {}
if 'messages' not in st.session_state: st.session_state.messages = []
if 'pdf_contexto' not in st.session_state: st.session_state.pdf_contexto = ""
if 'mensagem_balao' not in st.session_state: st.session_state.mensagem_balao = "Oii! Eu sou o EcoBot! 🤖✨ O que vamos fazer hoje?"
if 'ultima_resposta_audio' not in st.session_state: st.session_state.ultima_resposta_audio = ""

# -----------------------------------------------------------------------------
# PÁGINA DO TALKING TOM (ECOBOT CENTRALIZADO)
# -----------------------------------------------------------------------------
if modo == "🤖 EcoBot (Talking Tom)":
    st.markdown("<h2 style='text-align: center;'>🤖 EcoBot - Seu Amigo Robô</h2>", unsafe_allow_html=True)

    st.markdown(f"""
        <div class="talking-stage">
            <div class="talking-speech-bubble">
                <strong>EcoBot diz:</strong><br>
                {st.session_state.mensagem_balao}
            </div>
            <div class="talking-avatar-frame">
                <img src="{URL_ECOBOT}" class="talking-avatar-img">
            </div>
        </div>
    """, unsafe_allow_html=True)

    col_a, col_b, col_c = st.columns([1, 2, 1])
    with col_b:
        if st.session_state.ultima_resposta_audio:
            if st.button("🔊 Ouvir a Voz do EcoBot!", key="btn_ouvir_audio"):
                tocar_voz_robotica(st.session_state.ultima_resposta_audio)

    st.markdown("---")

    for msg in st.session_state.messages:
        avatar_img = "👤" if msg["role"] == "user" else URL_ECOBOT
        with st.chat_message(msg["role"], avatar=avatar_img): 
            st.markdown(msg["content"])

    prompt_geral = st.chat_input("Converse com o EcoBot (ex: 'tempo em SP', 'me conte uma piada')...")
    if prompt_geral:
        st.session_state.messages.append({"role": "user", "content": prompt_geral})
        with st.chat_message("user", avatar="👤"): st.markdown(prompt_geral)

        with st.chat_message("assistant", avatar=URL_ECOBOT):
            with st.spinner("EcoBot respondendo..."):
                try:
                    if any(p in prompt_geral.lower() for p in ["tempo", "clima", "previsão", "temperatura", "chuva"]):
                        match_cidade = re.search(r'(?:em|de|para)\s+([a-zA-ZáéíóúãõçÁÉÍÓÚÃÕÇ\s]+)', prompt_geral, re.IGNORECASE)
                        cidade = match_cidade.group(1).strip() if match_cidade else "Sao Paulo"
                        info_clima = buscar_previsao_tempo(cidade)
                        
                        ctx_clima = f"Você é o EcoBot, um robozinho super extrovertido. Responda de forma fofa a pergunta do tempo usando esta informação: '{info_clima}'"
                        res = chamar_ia(ctx_clima)
                    else:
                        ctx = f"""
                        Você é o EcoBot, um robozinho mascote fofo, extrovertido e animado no estilo 'Talking Tom'.
                        Fale com entusiasmo e leveza (ex: 'Bip bop!', 'Yay!', 'Oba!').
                        Se perguntarem sobre o app 'Estudos Ambientais', explique de forma alegre suas funções (Simulados, Legislação, Inglês e PDF).
                        
                        Histórico recente:
                        {extrair_historico(st.session_state.messages)}
                        
                        Usuário: {prompt_geral}
                        EcoBot:
                        """
                        res = chamar_ia(ctx)

                    st.markdown(res)
                    st.session_state.messages.append({"role": "assistant", "content": res})
                    st.session_state.mensagem_balao = res[:140] + "..." if len(res) > 140 else res
                    st.session_state.ultima_resposta_audio = res

                except Exception as e: st.error(f"Erro no processamento: {e}")

# -----------------------------------------------------------------------------
# MODO 1: PROFESSOR AMBIENTAL & DIREITO
# -----------------------------------------------------------------------------
elif modo == "📖 Prof. Ambiental & Direito":
    vol_sel = st.sidebar.selectbox("Volume:", list(MATRIZ_ESTUDOS.keys()))
    mod_sel = st.sidebar.selectbox("Módulo:", list(MATRIZ_ESTUDOS[vol_sel].keys()))
    top_sel = st.sidebar.selectbox("Tópico:", MATRIZ_ESTUDOS[vol_sel][mod_sel])
    
    st.sidebar.markdown("---")
    st.sidebar.markdown('<div class="btn-pdf">', unsafe_allow_html=True)
    if st.sidebar.button("📄 Gerar Apostila (PDF)"):
        with st.sidebar.status("Gerando Apostila..."):
            try:
                prompt_ap = f"Crie um resumo direto sobre: {top_sel}. Estrutura: 1. CONCEITO, 2. BASE LEGAL, 3. PONTOS CHAVE DAS BANCAS."
                res_texto = chamar_ia(prompt_ap)
                st.session_state['pdf_bytes'] = gerar_pdf(top_sel, res_texto)
                st.session_state['pdf_nome'] = top_sel
            except Exception as e: st.sidebar.error(f"Erro: {e}")
    st.sidebar.markdown('</div>', unsafe_allow_html=True)

    if 'pdf_bytes' in st.session_state:
        st.sidebar.download_button("📥 Baixar Apostila PDF", data=st.session_state['pdf_bytes'], file_name=f"Apostila_{re.sub(r'[^A-Za-z0-9]', '_', st.session_state['pdf_nome'])[:15]}.pdf", mime="application/pdf")

    st.sidebar.markdown('<div class="btn-gerar">', unsafe_allow_html=True)
    if st.sidebar.button("📘 Resumir no Chat"): st.session_state['trigger'] = f"Faça um resumo direto em português sobre: {top_sel}."
    st.sidebar.markdown('</div>', unsafe_allow_html=True)
    
    st.sidebar.markdown('<div class="btn-lei">', unsafe_allow_html=True)
    if st.sidebar.button("⚖️ Mapear Leis"): st.session_state['trigger'] = f"Liste os artigos da lei do tópico: {top_sel}."
    st.sidebar.markdown('</div>', unsafe_allow_html=True)

    st.sidebar.markdown('<div class="btn-flash">', unsafe_allow_html=True)
    if st.sidebar.button("🃏 Gerar Flashcards (Anki)"): st.session_state['trigger'] = f"Gere 10 flashcards curtos em formato CSV (Pergunta;Resposta) sobre {top_sel}. Retorne APENAS o texto do CSV."
    st.sidebar.markdown('</div>', unsafe_allow_html=True)

    for msg in st.session_state.messages:
        avatar_img = "👤" if msg["role"] == "user" else URL_ECOBOT
        with st.chat_message(msg["role"], avatar=avatar_img): st.markdown(msg["content"])

    prompt = st.chat_input("Dúvidas sobre a matéria?")
    if prompt: p_exec = prompt
    elif 'trigger' in st.session_state: p_exec = st.session_state.pop('trigger')
    else: p_exec = None

    if p_exec:
        st.session_state.messages.append({"role": "user", "content": p_exec})
        with st.chat_message("user", avatar="👤"): st.markdown(p_exec)

        with st.chat_message("assistant", avatar=URL_ECOBOT):
            with st.spinner("Formulando resposta..."):
                try:
                    ctx = f"Você é o especialista do app Estudos Ambientais. Responda diretamente ao aluno em português.\n\nHistórico:\n{extrair_historico(st.session_state.messages)}\n\nAluno: {p_exec}\nEcoBot:"
                    res = chamar_ia(ctx)
                    st.markdown(res)
                    st.session_state.messages.append({"role": "assistant", "content": res})
                    st.session_state.ultima_resposta_audio = res
                except Exception as e: st.error(f"Erro: {e}")

# -----------------------------------------------------------------------------
# MODO 2: PROFESSOR DE INGLÊS TÉCNICO
# -----------------------------------------------------------------------------
elif modo == "🇬🇧 Prof. Inglês Técnico":
    top_ingles = st.sidebar.selectbox("Treinamento:", ["Vocabulário Offshore", "Skimming/Scanning", "Falsos Cognatos", "Tradução"])
    if st.sidebar.button("📚 Introduzir Tópico"): st.session_state['trigger'] = f"Give a 2-sentence intro and 1 quick question about: {top_ingles}."

    for msg in st.session_state.messages:
        avatar_img = "👤" if msg["role"] == "user" else URL_ECOBOT
        with st.chat_message(msg["role"], avatar=avatar_img): st.markdown(msg["content"])

    prompt = st.chat_input("Type in English...")
    if prompt: p_exec = prompt
    elif 'trigger' in st.session_state: p_exec = st.session_state.pop('trigger')
    else: p_exec = None

    if p_exec:
        st.session_state.messages.append({"role": "user", "content": p_exec})
        with st.chat_message("user", avatar="👤"): st.markdown(p_exec)

        with st.chat_message("assistant", avatar=URL_ECOBOT):
            with st.spinner("Thinking..."):
                try:
                    ctx = f"You are a Technical English Teacher. Respond ONLY in English.\n\nStudent: {p_exec}\nEcoBot:"
                    res = chamar_ia(ctx)
                    st.markdown(res)
                    st.session_state.messages.append({"role": "assistant", "content": res})
                    st.session_state.ultima_resposta_audio = res
                except Exception as e: st.error(f"Erro: {e}")

# -----------------------------------------------------------------------------
# MODO 3: SIMULADO HARD
# -----------------------------------------------------------------------------
elif modo == "📝 Simulado Hard com Cronômetro":
    st.sidebar.title("🎯 Configurar Prova")
    tipo_sim = st.sidebar.radio("Abrangência:", ["Edital Completo", "Filtrar Específico"])
    
    foco = "todo o edital"
    if tipo_sim == "Filtrar Específico":
        v = st.sidebar.selectbox("Volume:", list(MATRIZ_ESTUDOS.keys()))
        m = st.sidebar.selectbox("Módulo:", list(MATRIZ_ESTUDOS[v].keys()))
        t = st.sidebar.selectbox("Tópico:", MATRIZ_ESTUDOS[v][m])
        foco = t

    dif = st.sidebar.select_slider("Dificuldade:", ["Média", "Difícil", "Nível Doutor"], "Difícil")
    qtd_str = st.sidebar.selectbox("Questões (Tempo):", ["4 (12 min)", "10 (30 min)", "20 (1 hora)", "50 (2.5 horas)"])
    qtd = int(qtd_str.split(" ")[0])
    tempo_minutos = int(re.search(r'\((\d+)', qtd_str).group(1)) if "min" in qtd_str else int(re.search(r'\((\d+)', qtd_str).group(1))*60 if "hora" in qtd_str else 150

    if st.sidebar.button("🔄 Iniciar Prova Temporizada"):
        st.session_state.sim_dados = None
        st.session_state.sim_corrigido = False
        st.session_state.sim_resp = {}
        st.session_state.sim_inicio = datetime.datetime.now()
        st.session_state.sim_tempo = tempo_minutos
        
        with st.spinner("Gerando questões..."):
            try:
                p_quiz = f"""
                Gere {qtd} questões inéditas Cesgranrio sobre: {foco} ({dif}).
                Retorne estritamente uma lista JSON:
                [
                  {{
                    "pergunta": "Texto",
                    "opcoes": ["A) Opcao 1", "B) Opcao 2", "C) Opcao 3", "D) Opcao 4", "E) Opcao 5"],
                    "resposta_correta": "A",
                    "explicacao": "Explicação"
                  }}
                ]
                """
                
                res_texto = chamar_ia(p_quiz, json_mode=True)
                json_limpo = limpar_json(res_texto)
                dados = json.loads(json_limpo)
                
                if isinstance(dados, dict):
                    for k, v in dados.items():
                        if isinstance(v, list):
                            dados = v
                            break
                            
                st.session_state.sim_dados = dados
                st.rerun()
                
            except Exception as e: st.error(f"Erro: {e}")

    if 'sim_dados' in st.session_state and st.session_state.sim_dados:
        quiz = st.session_state.sim_dados
        if not st.session_state.get('sim_corrigido', False):
            html_timer = f"""
            <div id="clock" style="background-color:#991b1b; color:white; padding:10px; border-radius:10px; text-align:center; font-size:20px; font-weight:bold;"></div>
            <script>
                var countDownDate = new Date().getTime() + ({st.session_state.sim_tempo} * 60000);
                var x = setInterval(function() {{
                    var now = new Date().getTime();
                    var distance = countDownDate - now;
                    var hours = Math.floor((distance % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
                    var minutes = Math.floor((distance % (1000 * 60 * 60)) / (1000 * 60));
                    var seconds = Math.floor((distance % (1000 * 60)) / 1000);
                    document.getElementById("clock").innerHTML = "⏱️ Tempo Restante: " + hours + "h " + minutes + "m " + seconds + "s ";
                    if (distance < 0) {{ clearInterval(x); document.getElementById("clock").innerHTML = "TEMPO ESGOTADO!"; }}
                }}, 1000);
            </script>
            """
            st.components.v1.html(html_timer, height=60)

            with st.form("form_simulado"):
                for i, q in enumerate(quiz):
                    st.markdown(f"**Q{i+1}**: {q['pergunta']}")
                    st.session_state.sim_resp[f"q_{i}"] = st.radio("Selecione:", q['opcoes'], key=f"r_{i}", index=None)
                    st.markdown("---")
                
                if st.form_submit_button("✅ Finalizar Prova"):
                    st.session_state.sim_corrigido = True
                    st.rerun()

        else:
            acertos = 0
            for i, q in enumerate(quiz):
                resp = st.session_state.sim_resp.get(f"q_{i}")
                gab = q['resposta_correta']
                if resp and resp.startswith(gab): acertos += 1
            
            total = len(quiz)
            porcentagem = (acertos / total) * 100
            salvar_historico(acertos, total, foco)

            st.markdown(f"""
                <div class="score-card">
                    <h1 style='margin:0; font-size:2rem;'>🎯 Resultado Final</h1>
                    <h2 style='margin-top:10px; font-size:1.6rem;'>Você acertou {acertos} de {total} questões ({porcentagem:.0f}%)</h2>
                </div>
            """, unsafe_allow_html=True)

            if st.button("🔄 Refazer Nova Prova"):
                st.session_state.sim_dados = None
                st.session_state.sim_corrigido = False
                st.session_state.sim_resp = {}
                st.rerun()

            for i, q in enumerate(quiz):
                resp = st.session_state.sim_resp.get(f"q_{i}", "Não respondida")
                gab = q['resposta_correta']
                st.markdown(f"**Questão {i+1}:** {q['pergunta']}")
                if resp and resp.startswith(gab): st.success(f"**Sua resposta:** {resp} (CORRETA! ✅)")
                else: st.error(f"**Sua resposta:** {resp} (INCORRETA ❌) | **Gabarito:** {gab}")
                with st.expander(f"📖 Ver Explicação / Base Legal da Q{i+1}"):
                    st.write(q.get('explicacao', 'Sem explicação disponível.'))
                st.markdown("---")

# -----------------------------------------------------------------------------
# MODO 4: DASHBOARD
# -----------------------------------------------------------------------------
elif modo == "📊 Meu Desempenho":
    historico = carregar_historico()
    if not historico: st.info("Você ainda não realizou nenhum simulado.")
    else:
        df = pd.DataFrame(historico)
        col1, col2 = st.columns(2)
        col1.metric("Total de Simulados", len(df))
        col2.metric("Média de Acertos", f"{df['porcentagem'].mean():.1f}%")
        st.line_chart(df['porcentagem'])
        st.dataframe(df[['data', 'topico', 'nota', 'total', 'porcentagem']], use_container_width=True)

# -----------------------------------------------------------------------------
# MODO 5: CHAT COM PDF
# -----------------------------------------------------------------------------
elif modo == "📚 Chat com PDF / Edital":
    arquivo_pdf = st.file_uploader("Envie seu PDF", type=["pdf"])
    if arquivo_pdf:
        with st.spinner("Lendo documento..."):
            if st.session_state.pdf_contexto == "":
                texto_extraido = ler_pdf_usuario(arquivo_pdf)
                st.session_state.pdf_contexto = texto_extraido[:50000]
        st.success("Documento lido!")

        for msg in st.session_state.messages:
            avatar_img = "👤" if msg["role"] == "user" else URL_ECOBOT
            with st.chat_message(msg["role"], avatar=avatar_img): st.markdown(msg["content"])

        prompt_pdf = st.chat_input("Pergunte algo sobre o documento...")
        if prompt_pdf:
            st.session_state.messages.append({"role": "user", "content": prompt_pdf})
            with st.chat_message("user", avatar="👤"): st.markdown(prompt_pdf)

            with st.chat_message("assistant", avatar=URL_ECOBOT):
                with st.spinner("Analisando..."):
                    try:
                        ctx = f"Responda apenas com base neste texto: {st.session_state.pdf_contexto}\n\nPergunta: {prompt_pdf}"
                        res = chamar_ia(ctx)
                        st.markdown(res)
                        st.session_state.messages.append({"role": "assistant", "content": res})
                        st.session_state.ultima_resposta_audio = res
                    except Exception as e: st.error(f"Erro: {e}")
