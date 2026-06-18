import os
import re
import sys
import subprocess
import requests
from dotenv import load_dotenv

# Configuração de caminhos
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(BASE_DIR, '.env')

# Carregar variáveis do .env
load_dotenv(ENV_PATH)

def update_env(key, value):
    """Atualiza de forma segura uma chave no arquivo .env sem afetar outras chaves."""
    if not os.path.exists(ENV_PATH):
        with open(ENV_PATH, 'w', encoding='utf-8') as f:
            f.write(f"{key}={value}\n")
        return

    with open(ENV_PATH, 'r', encoding='utf-8') as f:
        content = f.read()

    # Busca a linha da chave correspondente
    pattern = re.compile(rf'^\s*{key}\s*=.*$', re.MULTILINE)
    if pattern.search(content):
        new_content = pattern.sub(f"{key}={value}", content)
    else:
        new_content = content.rstrip() + f"\n{key}={value}\n"

    with open(ENV_PATH, 'w', encoding='utf-8') as f:
        f.write(new_content)

def get_credentials():
    """Recupera as credenciais do .env ou solicita ao usuário de forma interativa."""
    username = os.getenv("GITHUB_USERNAME")
    token = os.getenv("GITHUB_TOKEN")

    updated = False
    if not username:
        print("\n=== Configuração de Credenciais do GitHub ===")
        username = input("Digite seu nome de usuário do GitHub: ").strip()
        if not username:
            print("[ERROR] Nome de usuário é obrigatório.")
            sys.exit(1)
        update_env("GITHUB_USERNAME", username)
        updated = True

    if not token:
        token = input("Digite seu Personal Access Token (PAT) do GitHub: ").strip()
        if not token:
            print("[ERROR] Token de acesso é obrigatório.")
            sys.exit(1)
        update_env("GITHUB_TOKEN", token)
        updated = True

    if updated:
        print("[OK] Credenciais salvas no arquivo .env para futuras execuções.")
        # Recarregar as variáveis para garantir que o script as veja atualizadas
        load_dotenv(ENV_PATH)
        username = os.getenv("GITHUB_USERNAME")
        token = os.getenv("GITHUB_TOKEN")

    return username, token

def parse_tasks():
    """Lê o arquivo task.md e extrai a lista de tasks com seu status."""
    task_file = os.path.join(BASE_DIR, 'task.md')
    if not os.path.exists(task_file):
        print(f"[WARNING] Arquivo task.md não encontrado no root do projeto.")
        return []

    tasks = []
    # Padrão: - [x] Descrição da task ou - [ ] Descrição da task
    pattern = re.compile(r'^\s*-\s*\[([ xX/])\]\s*(.+)$')
    with open(task_file, 'r', encoding='utf-8') as f:
        for line in f:
            m = pattern.match(line)
            if m:
                status_char = m.group(1).lower()
                completed = status_char == 'x'
                title = m.group(2).strip()
                tasks.append({'title': title, 'completed': completed})
    return tasks

def check_or_create_repo(username, token, repo_name):
    """Verifica se o repositório existe no GitHub. Se não, cria-o como público."""
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json"
    }

    # 1. Verifica se já existe
    repo_url = f"https://api.github.com/repos/{username}/{repo_name}"
    res = requests.get(repo_url, headers=headers)

    if res.status_code == 200:
        print(f"[REPO] Repositorio '{repo_name}' ja existe no GitHub.")
        return True
    elif res.status_code == 404:
        # 2. Cria o repositório como público
        print(f"[START] Criando repositorio publico '{repo_name}' no GitHub...")
        create_url = "https://api.github.com/user/repos"
        payload = {
            "name": repo_name,
            "private": False,
            "description": "Control-T | Compliance & Routing Engine v3.0 - Monitoramento de Licitacoes"
        }
        res_create = requests.post(create_url, json=payload, headers=headers)
        if res_create.status_code == 201:
            print(f"[OK] Repositorio publico '{repo_name}' criado com sucesso!")
            return True
        else:
            print(f"[ERROR] Erro ao criar repositorio: {res_create.status_code}")
            print(res_create.text)
            return False
    else:
        print(f"[ERROR] Erro ao conectar com o GitHub API: {res.status_code}")
        print(res.text)
        return False

def push_codebase(username, token, repo_name):
    """Configura o Git local e envia o código para o GitHub."""
    print("\nPreparing to push code to GitHub...")
    
    # Adicionar todos os arquivos modificados (incluindo README.md, .env.example, task.md)
    subprocess.run(["git", "add", "."], capture_output=True, shell=True)
    subprocess.run(["git", "commit", "-m", "chore: sync project files, tasks, and readme"], capture_output=True, shell=True)

    # Remover remote anterior para garantir que não há conflitos
    subprocess.run(["git", "remote", "remove", "origin"], capture_output=True, shell=True)

    # Adicionar o novo remote contendo as credenciais para evitar prompts bloqueantes no push
    remote_url = f"https://{username}:{token}@github.com/{username}/{repo_name}.git"
    
    res_add = subprocess.run(["git", "remote", "add", "origin", remote_url], capture_output=True, shell=True)
    if res_add.returncode != 0:
        print("[ERROR] Erro ao adicionar o remote origin.")
        return False

    print("Enviando codigo para o branch 'master' no GitHub...")
    res_push = subprocess.run(["git", "push", "-u", "origin", "master", "--force"], capture_output=True, text=True, shell=True)
    
    if res_push.returncode == 0:
        print("[OK] Codigo enviado com sucesso!")
        return True
    else:
        print("[ERROR] Erro ao enviar o codigo para o GitHub.")
        print(res_push.stderr)
        return False

def get_existing_issues(username, token, repo_name):
    """Recupera todas as issues (abertas e fechadas) existentes no repositório."""
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json"
    }

    existing = {}
    page = 1
    while True:
        url = f"https://api.github.com/repos/{username}/{repo_name}/issues?state=all&per_page=100&page={page}"
        res = requests.get(url, headers=headers)
        if res.status_code != 200:
            break
        data = res.json()
        if not data:
            break
        for issue in data:
            # Ignorar Pull Requests que a API do GitHub retorna junto com issues
            if 'pull_request' not in issue:
                existing[issue['title']] = issue
        page += 1

    return existing

def sync_tasks_to_issues(username, token, repo_name, tasks):
    """Cria ou atualiza as issues no GitHub correspondentes a cada task de task.md."""
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json"
    }

    print(f"\nSincronizando {len(tasks)} tasks com as GitHub Issues...")
    existing = get_existing_issues(username, token, repo_name)

    for i, task in enumerate(tasks, start=1):
        title = task['title']
        completed = task['completed']
        status_label = "[CONCLUIDA]" if completed else "[PENDENTE]"
        
        print(f"[{i}/{len(tasks)}] Processando: {status_label} {title}")

        if title in existing:
            issue = existing[title]
            number = issue['number']
            state = issue['state']
            
            # Se a task está concluída no arquivo mas aberta no GitHub, fecha
            if completed and state == 'open':
                print(f"  -> Fechando issue #{number} no GitHub...")
                patch_url = f"https://api.github.com/repos/{username}/{repo_name}/issues/{number}"
                requests.patch(patch_url, json={"state": "closed"}, headers=headers)
            # Se a task está pendente no arquivo mas fechada no GitHub, reabre
            elif not completed and state == 'closed':
                print(f"  -> Reabrindo issue #{number} no GitHub...")
                patch_url = f"https://api.github.com/repos/{username}/{repo_name}/issues/{number}"
                requests.patch(patch_url, json={"state": "open"}, headers=headers)
            else:
                print(f"  -> OK (ja sincronizada)")
        else:
            # Criar a issue
            print(f"  -> Criando nova issue...")
            post_url = f"https://api.github.com/repos/{username}/{repo_name}/issues"
            payload = {
                "title": title,
                "body": f"Task importada do arquivo `task.md` do projeto Control-T.",
                "labels": ["task"]
            }
            res = requests.post(post_url, json=payload, headers=headers)
            if res.status_code == 201:
                new_issue = res.json()
                number = new_issue['number']
                if completed:
                    # Se for concluída, fecha ela imediatamente
                    print(f"  -> Fechando issue #{number} criada...")
                    patch_url = f"https://api.github.com/repos/{username}/{repo_name}/issues/{number}"
                    requests.patch(patch_url, json={"state": "closed"}, headers=headers)
            else:
                print(f"  [ERROR] Erro ao criar issue: {res.status_code}")
                print(res.text)

    print("\n[SUCCESS] Sincronizacao de tasks concluida com sucesso!")
    print(f"[LINK] Link do repositorio: https://github.com/{username}/{repo_name}")

def main():
    repo_name = "thunderbird-control-t"
    
    # 1. Obter credenciais
    username, token = get_credentials()

    # 2. Carregar as tasks do task.md
    tasks = parse_tasks()
    if not tasks:
        print("[ERROR] Nenhuma task encontrada em task.md. Encerrando.")
        return

    # 3. Validar/Criar repositório remoto
    if not check_or_create_repo(username, token, repo_name):
        return

    # 4. Enviar código para o repositório
    if not push_codebase(username, token, repo_name):
        return

    # 5. Sincronizar as tasks com as Issues do GitHub
    sync_tasks_to_issues(username, token, repo_name, tasks)

if __name__ == "__main__":
    main()
