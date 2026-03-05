import flet as ft
import pandas as pd
import smtplib
from email.message import EmailMessage
import json
import os
import keyring

def main(page: ft.Page):
    # --- Configurações Visuais da Página ---
    page.title = "DispEmail"
    page.padding = 30
    page.theme = ft.Theme(use_material3=True, color_scheme_seed=ft.Colors.INDIGO)
    page.theme_mode = ft.ThemeMode.LIGHT
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER 
    
    # CORREÇÃO DE COR: Usando uma string segura ou Colors.SURFACE
    page.bgcolor = "#f0f2f5" 
    
    df_dados = None
    ARQUIVO_CONFIG = "config_email.json"
    NOME_APP = "DispEmail" 

    config_email = {
        "email": "",
        "senha": "",
        "servidor": "smtp.gmail.com",
        "porta": "587"
    }

    # --- Carregar configurações salvas ---
    if os.path.exists(ARQUIVO_CONFIG):
        try:
            with open(ARQUIVO_CONFIG, "r", encoding="utf-8") as f:
                dados_salvos = json.load(f)
                config_email["email"] = dados_salvos.get("email", "")
                config_email["servidor"] = dados_salvos.get("servidor", "smtp.gmail.com")
                config_email["porta"] = dados_salvos.get("porta", "587")
                
                if config_email["email"]:
                    senha_salva = keyring.get_password(NOME_APP, config_email["email"])
                    if senha_salva:
                        config_email["senha"] = senha_salva
        except Exception as ex:
            print(f"Erro ao carregar configs: {ex}")

    # --- Lógica de Seleção de Arquivo ---
    def on_file_result(e: ft.FilePickerResultEvent):
        nonlocal df_dados
        if e.files:
            file_path = e.files[0].path
            try:
                if file_path.endswith('.xlsx'):
                    df_dados = pd.read_excel(file_path)
                else:
                    df_dados = pd.read_csv(file_path)
                
                data_table.rows.clear()
                for _, row in df_dados.iterrows():
                    data_table.rows.append(
                        ft.DataRow(cells=[
                            ft.DataCell(ft.Text(str(row['Nº']))),
                            ft.DataCell(ft.Text(str(row['NOME']))),
                            ft.DataCell(ft.Text(str(row['EMAIL']))),
                            ft.DataCell(ft.Text(str(row['CONTATO']))),
                        ])
                    )
                # Reseta a barra de progresso ao carregar nova planilha
                progress_bar.value = 0
                progress_text.value = f"Aguardando disparo (0/{len(df_dados)})"
                page.update()
                
                page.snack_bar = ft.SnackBar(ft.Text("Planilha carregada com sucesso!"))
                page.snack_bar.open = True
                page.update()
            except Exception as ex:
                print(f"Erro ao ler arquivo: {ex}")

    file_picker = ft.FilePicker(on_result=on_file_result)
    page.overlay.append(file_picker)

    # --- Janela de Configuração ---
    inp_email = ft.TextField(label="Seu E-mail", value=config_email["email"], border_radius=8)
    inp_senha = ft.TextField(label="Senha de App", password=True, can_reveal_password=True, value=config_email["senha"], border_radius=8) 
    inp_servidor = ft.TextField(label="Servidor SMTP", value=config_email["servidor"], border_radius=8)
    inp_porta = ft.TextField(label="Porta SMTP", value=config_email["porta"], border_radius=8)

    def salvar_config(e):
        config_email["email"] = inp_email.value
        config_email["senha"] = inp_senha.value
        config_email["servidor"] = inp_servidor.value
        config_email["porta"] = inp_porta.value
        
        dados_salvos = {"email": config_email["email"], "servidor": config_email["servidor"], "porta": config_email["porta"]}
        with open(ARQUIVO_CONFIG, "w", encoding="utf-8") as f:
            json.dump(dados_salvos, f, indent=4)

        if config_email["email"] and config_email["senha"]:
            keyring.set_password(NOME_APP, config_email["email"], config_email["senha"])

        dlg_config.open = False
        page.snack_bar = ft.SnackBar(ft.Text("Configurações salvas!"))
        page.snack_bar.open = True
        page.update()

    dlg_config = ft.AlertDialog(
        title=ft.Text("Configurar Servidor", weight="bold"),
        content=ft.Column([inp_email, inp_senha, inp_servidor, inp_porta], tight=True),
        actions=[
            ft.TextButton("Cancelar", on_click=lambda _: setattr(dlg_config, "open", False) or page.update()),
            ft.FilledButton("Salvar", on_click=salvar_config), 
        ],
        shape=ft.RoundedRectangleBorder(radius=16) 
    )
    page.overlay.append(dlg_config)

    # --- AppBar ---
    page.appbar = ft.AppBar(
        leading=ft.Icon(ft.Icons.MARK_EMAIL_READ_ROUNDED, color=ft.Colors.INDIGO),
        title=ft.Text("DispEmail", weight="w800", color=ft.Colors.INDIGO),
        bgcolor=ft.Colors.SURFACE,
        actions=[
            ft.PopupMenuButton(
                items=[
                    ft.PopupMenuItem(text="Abrir Planilha", icon=ft.Icons.UPLOAD_FILE, on_click=lambda _: file_picker.pick_files()),
                    ft.PopupMenuItem(text="Configurar Servidor", icon=ft.Icons.SETTINGS, on_click=lambda _: setattr(dlg_config, "open", True) or page.update()),
                ]
            ),
        ]
    )

    # --- Lógica de Envio com Barra de Progresso ---
    def disparar_emails(e):
        if df_dados is None or df_dados.empty:
            page.snack_bar = ft.SnackBar(ft.Text("Abra uma planilha primeiro!"), bgcolor=ft.Colors.ERROR)
            page.snack_bar.open = True
            page.update()
            return

        try:
            btn_enviar.disabled = True
            total = len(df_dados)
            page.update()

            porta = int(config_email["porta"])
            with smtplib.SMTP(config_email["servidor"], porta) as smtp:
                smtp.starttls()
                smtp.login(config_email["email"], config_email["senha"])
                
                for i, row in df_dados.iterrows():
                    # Atualiza a barra de progresso (valor entre 0.0 e 1.0)
                    progresso_atual = (i + 1) / total
                    progress_bar.value = progresso_atual
                    progress_text.value = f"Enviando: {i+1} de {total}..."
                    page.update()

                    msg = EmailMessage()
                    msg['Subject'] = assunto_input.value if assunto_input.value else "Sem Assunto"
                    msg['From'] = config_email["email"]
                    msg['To'] = str(row['EMAIL'])
                    
                    corpo = msg_input.value.replace("{NOME}", str(row['NOME'])).replace("{CONTATO}", str(row['CONTATO']))
                    msg.set_content(corpo)
                    
                    smtp.send_message(msg)
            
            progress_text.value = f"Concluído! {total} e-mails enviados."
            page.snack_bar = ft.SnackBar(ft.Text("Envio finalizado com sucesso!"), bgcolor=ft.Colors.GREEN)
            page.snack_bar.open = True
            
        except Exception as ex:
            page.snack_bar = ft.SnackBar(ft.Text(f"Erro: {ex}"), bgcolor=ft.Colors.ERROR)
            page.snack_bar.open = True
        finally:
            btn_enviar.disabled = False
            page.update()

    # --- Componentes da Interface ---
    assunto_input = ft.TextField(label="Assunto", border_radius=12, filled=True)
    msg_input = ft.TextField(label="Mensagem", multiline=True, min_lines=6, border_radius=12, filled=True)
    
    # Barra de Progresso
    progress_bar = ft.ProgressBar(value=0, color=ft.Colors.INDIGO, bgcolor=ft.Colors.OUTLINE_VARIANT)
    progress_text = ft.Text("Aguardando disparo...", size=12, color=ft.Colors.ON_SURFACE_VARIANT)

    data_table = ft.DataTable(
        columns=[ft.DataColumn(ft.Text("Nº")), ft.DataColumn(ft.Text("Nome")), ft.DataColumn(ft.Text("E-mail")), ft.DataColumn(ft.Text("Contato"))],
        rows=[]
    )

    table_container = ft.Container(
        content=ft.Column([data_table], scroll=ft.ScrollMode.ALWAYS),
        height=200, bgcolor=ft.Colors.SURFACE, border_radius=12, padding=10, border=ft.border.all(1, "#dee2e6")
    )

    btn_enviar = ft.FilledButton("Disparar para todos", icon=ft.Icons.SEND_ROUNDED, on_click=disparar_emails)

    # Layout
    main_content = ft.Column(
        width=800,
        controls=[
            ft.Text("Lista de Envio", size=18, weight="bold"),
            table_container,
            ft.Divider(height=20, color=ft.Colors.TRANSPARENT),
            progress_text,
            progress_bar,
            ft.Divider(height=10, color=ft.Colors.TRANSPARENT),
            assunto_input,
            msg_input,
            ft.Row([btn_enviar], alignment=ft.MainAxisAlignment.END)
        ]
    )

    page.add(main_content)

ft.app(target=main)
