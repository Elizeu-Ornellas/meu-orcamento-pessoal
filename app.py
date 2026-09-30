from datetime import date, timedelta
import os
import urllib.parse
import dateutil.relativedelta
import pandas as pd
import plotly.express as px
import streamlit as st

NOME_ARQUIVO = 'Orcamento_Pessoal.xlsx'


# Carregar dados da planilha
def carregar_dados():
  if os.path.exists(NOME_ARQUIVO):
    try:
      df = pd.read_excel(NOME_ARQUIVO)
      if 'Data' in df.columns:
        df['Data'] = pd.to_datetime(df['Data']).dt.date
      if 'Vencimento' in df.columns:
        df['Vencimento'] = pd.to_datetime(df['Vencimento']).dt.date
      else:
        df['Vencimento'] = df['Data']
      if 'Status' not in df.columns:
        df['Status'] = 'Pendente'
      return df
    except Exception:
      pass
  return pd.DataFrame(
      columns=[
          'Data',
          'Vencimento',
          'Descrição',
          'Categoria',
          'Tipo',
          'Valor (R$)',
          'Status',
      ]
  )


# Salvar dados no Excel
def salvar_dados(df):
  df.to_excel(NOME_ARQUIVO, index=False)


st.set_page_config(
    page_title="Gestão de Gastos & Orçamento",
    page_icon="💳",
    layout="wide",
)
st.title("💳 Gestão de Gastos & Orçamento Pessoal")

if 'df_dados' not in st.session_state:
  st.session_state.df_dados = carregar_dados()

df = st.session_state.df_dados

CATEGORIAS_DESPESA = [
    "Alimentação",
    "Habitação/Contas",
    "Transporte",
    "Saúde/Academia",
    "Lazer/Variados",
    "Assinaturas/Serviços",
    "Outros",
]
CATEGORIAS_RECEITA = [
    "Salário",
    "Rendimento",
    "Vendas/Serviços",
    "Outras Receitas",
]

# --- MENU LATERAL ---
st.sidebar.header("⚙️️ Operações")
opcao_menu = st.sidebar.radio(
    "Escolha uma ação:",
    [
        "➕ Cadastrar Novo Lançamento",
        "📲 Enviar Alerta via WhatsApp",
        "✏️ Alterar Lançamento",
        "🗑️ Excluir Lançamento",
    ],
)

# ---------------------------------------------------------
# 1. CADASTRAR NOVO
# ---------------------------------------------------------
if opcao_menu == "➕ Cadastrar Novo Lançamento":
  st.sidebar.subheader("➕ Novo Registro")

  data_input = st.sidebar.date_input("Data do Lançamento", value=date.today())
  vencimento_input = st.sidebar.date_input(
      "Data de Vencimento", value=date.today() + timedelta(days=5)
  )
  descricao_input = st.sidebar.text_input(
      "Descrição", placeholder="ex: Conta de Energia, Cartão de Crédito"
  )
  tipo_input = st.sidebar.selectbox("Tipo", ["Despesa", "Receita"])
  cats = CATEGORIAS_DESPESA if tipo_input == "Despesa" else CATEGORIAS_RECEITA
  categoria_input = st.sidebar.selectbox("Categoria", cats)
  valor_input = st.sidebar.number_input(
      "Valor (R$)", min_value=0.01, value=100.00, format="%.2f"
  )

  st.sidebar.markdown("---")
  frequencia = st.sidebar.radio(
      "Frequência:",
      [
          "1️⃣ Único (Apenas este mês)",
          "🔢 Parcelado (Número fixo de vezes)",
          "♾️ Recorrente/Variável (Água, Energia, etc.)",
      ],
  )

  qtd_meses = 1
  if frequencia == "🔢 Parcelado (Número fixo de vezes)":
    qtd_meses = st.sidebar.number_input(
        "Quantidade de Parcelas:", min_value=2, max_value=120, value=12, step=1
    )
  elif frequencia == "♾️ Recorrente/Variável (Água, Energia, etc.)":
    qtd_meses = st.sidebar.number_input(
        "Projetar quantos meses no futuro?",
        min_value=6,
        max_value=60,
        value=24,
        step=6,
    )

  if st.sidebar.button("💾 Salvar Lançamento", use_container_width=True):
    if not descricao_input.strip():
      st.sidebar.error("Preencha a descrição do lançamento.")
    else:
      novos_registros = []
      num_ocorrencias = int(qtd_meses)

      for i in range(num_ocorrencias):
        data_parc = data_input + dateutil.relativedelta.relativedelta(months=i)
        venc_parc = vencimento_input + dateutil.relativedelta.relativedelta(
            months=i
        )

        if frequencia == "🔢 Parcelado (Número fixo de vezes)":
          desc_final = f"{descricao_input} ({i+1}/{num_ocorrencias})"
        else:
          desc_final = descricao_input

        novos_registros.append({
            "Data": data_parc,
            "Vencimento": venc_parc,
            "Descrição": desc_final,
            "Categoria": categoria_input,
            "Tipo": tipo_input,
            "Valor (R$)": valor_input,
            "Status": "Pendente",
        })

      df_novos = pd.DataFrame(novos_registros)
      st.session_state.df_dados = pd.concat(
          [st.session_state.df_dados, df_novos], ignore_index=True
      )
      salvar_dados(st.session_state.df_dados)
      st.sidebar.success("Lançamento(s) salvo(s) com sucesso!")
      st.rerun()

# ---------------------------------------------------------
# 2. ENVIAR ALERTAS WHATSAPP (COMPATÍVEL COM CELULAR/NUVEM)
# ---------------------------------------------------------
elif opcao_menu == "📲 Enviar Alerta via WhatsApp":
  st.sidebar.subheader("📲 Notificação de Vencimentos")
  numero_celular = st.sidebar.text_input(
      "Número do WhatsApp (com DDD):",
      placeholder="ex: 5567999999999",
      help="Apenas números: Código do país (55) + DDD + Número",
  )
  dias_antecedencia = st.sidebar.slider(
      "Avisar contas a vencer nos próximos (dias):", 1, 15, 5
  )

  hoje = date.today()
  limite_venc = hoje + timedelta(days=dias_antecedencia)

  # Filtra contas pendentes a vencer
  df_a_vencer = df[
      (df["Tipo"] == "Despesa")
      & (df["Status"] == "Pendente")
      & (df["Vencimento"] >= hoje)
      & (df["Vencimento"] <= limite_venc)
  ]

  if df_a_vencer.empty:
    st.sidebar.info("Nenhuma conta pendente a vencer no período definido!")
  else:
    num_limpo = (
        numero_celular.replace("+", "")
        .replace("-", "")
        .replace(" ", "")
        .strip()
    )

    if not num_limpo:
      st.sidebar.warning("Digite seu número de celular acima para gerar o link.")
    else:
      mensagem = (
          "🔔 *AVISO DE VENCIMENTO DE CONTAS*\n\nOlá! Segue a lista de contas"
          f" a vencer nos próximos {dias_antecedencia} dias:\n\n"
      )
      for _, row in df_a_vencer.iterrows():
        venc_str = (
            row["Vencimento"].strftime("%d/%m/%Y")
            if isinstance(row["Vencimento"], date)
            else str(row["Vencimento"])
        )
        mensagem += (
            f"📌 *{row['Descrição']}*\n   Vencimento: {venc_str}\n   Valor: R$"
            f" {row['Valor (R$)']:.2f}\n\n"
        )

      msg_url = urllib.parse.quote(mensagem)
      whatsapp_link = f"https://wa.me/{num_limpo}?text={msg_url}"

      st.sidebar.markdown(
          f'<a href="{whatsapp_link}" target="_blank" style="text-decoration:none;">'
          '<button style="width:100%; background-color:#25D366; color:white; border:none; padding:10px; border-radius:5px; font-weight:bold; cursor:pointer;">'
          "💬 Abrir WhatsApp e Enviar Aviso</button></a>",
          unsafe_allow_html=True,
      )

# ---------------------------------------------------------
# 3. ALTERAR LANÇAMENTO
# ---------------------------------------------------------
elif opcao_menu == "✏️ Alterar Lançamento":
  st.sidebar.subheader("✏️ Editar Registro")
  if df.empty:
    st.sidebar.info("Nenhum registro para alterar.")
  else:
    opcoes_registros = {
        f"ID {idx} | Venc: {row['Vencimento']} - {row['Descrição']} (R$"
        f" {row['Valor (R$)']:.2f}) [{row.get('Status', 'Pendente')}]": idx
        for idx, row in df.iterrows()
    }
    registro_sel = st.sidebar.selectbox(
        "Selecione o registro:", list(opcoes_registros.keys())
    )
    idx_sel = opcoes_registros[registro_sel]
    dados_reg = df.loc[idx_sel]

    data_edit = st.sidebar.date_input(
        "Data Lançamento", value=dados_reg["Data"]
    )
    venc_edit = st.sidebar.date_input(
        "Data Vencimento", value=dados_reg.get("Vencimento", dados_reg["Data"])
    )
    desc_edit = st.sidebar.text_input(
        "Descrição", value=str(dados_reg["Descrição"])
    )
    tipo_idx = 0 if dados_reg["Tipo"] == "Despesa" else 1
    tipo_edit = st.sidebar.selectbox(
        "Tipo", ["Despesa", "Receita"], index=tipo_idx
    )

    cats = CATEGORIAS_DESPESA if tipo_edit == "Despesa" else CATEGORIAS_RECEITA
    cat_idx = (
        cats.index(dados_reg["Categoria"])
        if dados_reg["Categoria"] in cats
        else 0
    )
    cat_edit = st.sidebar.selectbox("Categoria", cats, index=cat_idx)
    valor_edit = st.sidebar.number_input(
        "Valor (R$)",
        min_value=0.01,
        value=float(dados_reg["Valor (R$)"]),
        format="%.2f",
    )

    status_atual = dados_reg.get("Status", "Pendente")
    status_options = ["Pendente", "Pago", "Estimado"]
    status_idx = (
        status_options.index(status_atual)
        if status_atual in status_options
        else 0
    )
    status_edit = st.sidebar.selectbox(
        "Status do Pagamento", status_options, index=status_idx
    )

    if st.sidebar.button("💾 Salvar Alteração", use_container_width=True):
      st.session_state.df_dados.loc[idx_sel] = {
          "Data": data_edit,
          "Vencimento": venc_edit,
          "Descrição": desc_edit,
          "Categoria": cat_edit,
          "Tipo": tipo_edit,
          "Valor (R$)": valor_edit,
          "Status": status_edit,
      }
      salvar_dados(st.session_state.df_dados)
      st.sidebar.success("Registro alterado com sucesso!")
      st.rerun()

# ---------------------------------------------------------
# 4. EXCLUIR LANÇAMENTO
# ---------------------------------------------------------
elif opcao_menu == "🗑️️ Excluir Lançamento":
  st.sidebar.subheader("🗑️ Apagar Registro")
  if df.empty:
    st.sidebar.info("Nenhum registro para excluir.")
  else:
    opcoes_registros = {
        f"ID {idx} | Venc: {row['Vencimento']} - {row['Descrição']} (R$"
        f" {row['Valor (R$)']:.2f})": idx
        for idx, row in df.iterrows()
    }
    registro_sel = st.sidebar.selectbox(
        "Selecione o registro para apagar:", list(opcoes_registros.keys())
    )
    idx_sel = opcoes_registros[registro_sel]

    st.sidebar.warning(f"Confirma a exclusão do registro ID {idx_sel}?")
    if st.sidebar.button("🗑️ Confirmar Exclusão", use_container_width=True):
      st.session_state.df_dados = st.session_state.df_dados.drop(
          index=idx_sel
      ).reset_index(drop=True)
      salvar_dados(st.session_state.df_dados)
      st.sidebar.success("Registro excluído com sucesso!")
      st.rerun()

# --- PAINEL PRINCIPAL ---
if not df.empty:
  # Alertas de Vencimento Próximo
  hoje = date.today()
  limite_aviso = hoje + timedelta(days=5)

  df_pendentes_proximas = df[
      (df["Tipo"] == "Despesa")
      & (df["Status"] == "Pendente")
      & (df["Vencimento"] >= hoje)
      & (df["Vencimento"] <= limite_aviso)
  ]

  if not df_pendentes_proximas.empty:
    st.warning(
        f"⚠️ **ATENÇÃO:** Você possui {len(df_pendentes_proximas)} conta(s) a"
        " vencer nos próximos 5 dias!"
    )
    st.dataframe(
        df_pendentes_proximas[["Vencimento", "Descrição", "Valor (R$)", "Status"]],
        use_container_width=True,
    )

  st.divider()

  df_temp = df.copy()
  df_temp["Data_DT"] = pd.to_datetime(df_temp["Vencimento"])
  df_temp["Ano_Mes"] = df_temp["Data_DT"].dt.strftime("%Y-%m")

  meses_disponiveis = ["Todos"] + sorted(
      df_temp["Ano_Mes"].unique().tolist(), reverse=False
  )

  col_filtro1, col_filtro2 = st.columns([2, 1])
  with col_filtro1:
    mes_selecionado = st.selectbox(
        "📅 Filtrar por Mês de Vencimento (Ano-Mês):", meses_disponiveis
    )

  if mes_selecionado != "Todos":
    df_filtrado = df_temp[df_temp["Ano_Mes"] == mes_selecionado].drop(
        columns=["Data_DT", "Ano_Mes"]
    )
  else:
    df_filtrado = df.copy()

  with col_filtro2:
    st.write("")
    st.write("")
    csv_data = df_filtrado.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Baixar Dados (CSV)",
        data=csv_data,
        file_name=f'orcamento_{mes_selecionado}.csv',
        mime="text/csv",
    )

  total_receita = df_filtrado[df_filtrado["Tipo"] == "Receita"][
      "Valor (R$)"
  ].sum()
  total_despesa = df_filtrado[df_filtrado["Tipo"] == "Despesa"][
      "Valor (R$)"
  ].sum()
  saldo = total_receita - total_despesa

  c1, c2, c3 = st.columns(3)
  c1.metric("Total de Receitas", f"R$ {total_receita:,.2f}")
  c2.metric("Total de Despesas", f"R$ {total_despesa:,.2f}")
  c3.metric(
      "Saldo do Período",
      f"R$ {saldo:,.2f}",
      delta=f"R$ {saldo:,.2f}",
      delta_color="normal" if saldo >= 0 else "inverse",
  )

  st.divider()

  g1, g2 = st.columns(2)
  df_despesas = df_filtrado[df_filtrado["Tipo"] == "Despesa"]

  if not df_despesas.empty:
    with g1:
      fig_pizza = px.pie(
          df_despesas,
          values="Valor (R$)",
          names="Categoria",
          title="Gastos por Categoria",
      )
      st.plotly_chart(fig_pizza, use_container_width=True)

    with g2:
      fig_barras = px.bar(
          df_despesas,
          x="Descrição",
          y="Valor (R$)",
          color="Categoria",
          title="Gastos por Descrição",
      )
      st.plotly_chart(fig_barras, use_container_width=True)

  st.subheader("📋 Tabela de Lançamentos")
  df_exibicao = df_filtrado.copy()
  df_exibicao.index.name = "ID"
  st.dataframe(df_exibicao, use_container_width=True)

else:
  st.info(
      "Nenhum lançamento registrado ainda. Selecione **'➕ Cadastrar Novo"
      " Lançamento'** no menu lateral para começar!"
  )
