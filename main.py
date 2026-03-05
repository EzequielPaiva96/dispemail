import flet as ft
import pandas as pd
import smtplib
from email.message import EmailMessage

def main(page: ft.Page):
    page.title = "AutoMail - Disparador de E-mails"
    page.padding = 30
    page.theme_mode = ft.ThemeMode.LIGHT
    
    df_dados = None

    # --- Configurações de E-mail ---
    EMAIL_REMETENTE = "ecomaes@cesjo.org.br"
    SENHA_APP = "zger wkat mqwu xebw" # AVISO: Mude essa senha, pois ela foi exposta publicamente
    SERVIDOR_SMTP = "smtp.gmail.com" 
    PORTA_SMTP = 587

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
                page.update()
            except Exception as ex:
                print(f"Erro ao ler arquivo: {ex}")

    # --- Lógica de Envio de E-mail ---
    def disparar_emails(e):
        if df_dados is None or df_dados.empty:
            page.snack_bar = ft.SnackBar(ft.Text("Por favor, carregue uma planilha primeiro!"))
            page.snack_bar.open = True
            page.update()
            return
            
        texto_base = msg_input.value
        assunto_texto = assunto_input.value # Pega o texto digitado no campo de assunto
        
        # Verifica se preencheu a mensagem
        if not texto_base:
            page.snack_bar = ft.SnackBar(ft.Text("Por favor, digite uma mensagem para o e-mail!"))
            page.snack_bar.open = True
            page.update()
            return

        try:
            # Mostra um feedback visual de carregamento
            e.control.text = "Enviando..."
            e.control.disabled = True
            page.update()

            # Conexão SMTP (Porta 587 requer STARTTLS ANTES do login)
            with smtplib.SMTP(SERVIDOR_SMTP, PORTA_SMTP) as smtp:
                smtp.starttls() # CORREÇÃO: Isso OBRIGATORIAMENTE vem antes do login
                smtp.login(EMAIL_REMETENTE, SENHA_APP)
                
                for _, row in df_dados.iterrows():
                    nome_dest = str(row['NOME'])
                    email_dest = str(row['EMAIL'])
                    contato_dest = str(row['CONTATO'])
                    
                    # Personaliza a mensagem substituindo as tags
                    msg_personalizada = texto_base.replace("{N}", nome_dest).replace("{C}", contato_dest)
                    
                    # Monta o e-mail
                    msg = EmailMessage()
                    # Usa o assunto digitado ou um padrão caso tenha deixado em branco
                    msg['Subject'] = assunto_texto if assunto_texto else "Sem Assunto" 
                    msg['From'] = EMAIL_REMETENTE
                    msg['To'] = email_dest
                    msg.set_content(msg_personalizada)
                    
                    # Envia a mensagem
                    smtp.send_message(msg)
            
            # Sucesso
            page.snack_bar = ft.SnackBar(ft.Text("Todos os e-mails foram enviados com sucesso!", color="green"))
            page.snack_bar.open = True
            
        except Exception as ex:
            # Erro
            page.snack_bar = ft.SnackBar(ft.Text(f"Erro ao enviar: {ex}", color="red"))
            page.snack_bar.open = True
            print(ex)
            
        finally:
            # Restaura o botão
            e.control.text = "Disparar para todos"
            e.control.disabled = False
            page.update()

    # --- Interface ---
    file_picker = ft.FilePicker(on_result=on_file_result)
    page.overlay.append(file_picker)

    # NOVO: Campo para o Assunto
    assunto_input = ft.TextField(
        label="Assunto do E-mail",
        hint_text="Ex: Atualização de Cadastro",
        prefix_icon=ft.Icons.SUBJECT
    )

    msg_input = ft.TextField(
        label="Corpo do E-mail", 
        multiline=True, 
        min_lines=5,
        hint_text="Use as tags {N} para nome personalizado...", 
        helper_text="Dica: Use as tags {N} para nome personalizado"
    )

    data_table = ft.DataTable(
        columns=[
            ft.DataColumn(ft.Text("Nº")),
            ft.DataColumn(ft.Text("Nome")),
            ft.DataColumn(ft.Text("E-mail")),
            ft.DataColumn(ft.Text("Contato")),
        ],
        rows=[]
    )

    # Layout
    page.add(
        ft.Row([
            ft.Icon(ft.Icons.EMAIL_OUTLINED, color="blue", size=30),
            ft.Text("AutoMail Personalizado", size=25, weight="bold"),
        ]),
        ft.ElevatedButton(
            "Selecionar Planilha", 
            icon=ft.Icons.UPLOAD_FILE, 
            on_click=lambda _: file_picker.pick_files()
        ),
        ft.Column([data_table], scroll=ft.ScrollMode.ALWAYS, height=200),
        assunto_input, # Adicionado o campo de Assunto na tela
        msg_input,
        ft.ElevatedButton(
            "Disparar para todos", 
            icon=ft.Icons.SEND_ROUNDED, 
            bgcolor="blue",
            color="white",
            on_click=disparar_emails 
        )
    )

ft.app(target=main)
