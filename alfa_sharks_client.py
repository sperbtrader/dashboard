import requests
import base64
import time

def atualizar_via_api():
    try:
        # Sua configuração
        token = "seu_token_github"
        repo_owner = "sperbtrader"
        repo_name = "dashboard"
        file_path = "planilhafluxoWINFUT.xlsx"
        
        # Ler e codificar o arquivo
        with open(arquivo_local, "rb") as f:
            content = base64.b64encode(f.read()).decode("utf-8")
        
        # Fazer requisição para API do GitHub
        url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/contents/{file_path}"
        headers = {"Authorization": f"token {token}"}
        
        response = requests.put(url, headers=headers, json={
            "message": f"Update automático {time.strftime('%Y-%m-%d %H:%M:%S')}",
            "content": content,
            "sha": get_current_sha()  # Você precisa obter o SHA atual do arquivo
        })
        
        print(f"Planilha atualizada: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    except Exception as e:
        print(f"Erro: {e}")
