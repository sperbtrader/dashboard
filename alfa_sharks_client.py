import requests
import base64
import time

def get_current_sha(token, repo_owner, repo_name, file_path):
    """Obtém o SHA atual do arquivo do GitHub para que a atualização seja válida."""
    try:
        url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/contents/{file_path}"
        headers = {"Authorization": f"token {token}"}
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.json()["sha"]
        else:
            print(f"Erro ao obter o SHA: {response.status_code}, {response.text}")
            return None
    except Exception as e:
        print(f"Erro ao obter o SHA: {e}")
        return None

def atualizar_via_api():
    try:
        # Sua configuração
        token = "seu_token_github"
        repo_owner = "sperbtrader"
        repo_name = "dashboard"
        file_path = "planilhafluxoWINFUT.xlsx"
        
        # Obter o SHA atual do arquivo
        current_sha = get_current_sha(token, repo_owner, repo_name, file_path)
        if not current_sha:
            print("Não foi possível obter o SHA atual. A atualização não pode ser feita.")
            return

        # Ler e codificar o arquivo
        with open(file_path, "rb") as f:
            content = base64.b64encode(f.read()).decode("utf-8")
        
        # Fazer requisição para API do GitHub
        url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/contents/{file_path}"
        headers = {"Authorization": f"token {token}"}
        
        response = requests.put(url, headers=headers, json={
            "message": f"Update automático {time.strftime('%Y-%m-%d %H:%M:%S')}",
            "content": content,
            "sha": current_sha
        })
        
        if response.status_code == 200:
            print(f"Planilha atualizada: {time.strftime('%Y-%m-%d %H:%M:%S')}")
        else:
            print(f"Erro na atualização: {response.status_code}, {response.text}")
            
    except FileNotFoundError:
        print(f"Erro: O arquivo '{file_path}' não foi encontrado.")
    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    # É necessário que o arquivo planilhafluxoWINFUT.xlsx esteja no mesmo diretório
    atualizar_via_api()
