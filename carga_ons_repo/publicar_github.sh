#!/bin/bash
# publicar_github.sh
# Cria o repositório no GitHub e faz o push do código.
# Requer: git e GitHub CLI (gh) instalados.
# Instale o gh em: https://cli.github.com/

set -e

REPO_NAME="carga-ons"
DESCRICAO="Download automático da carga efetiva de energia elétrica (ONS/SDRO)"

echo "=== Inicializando repositório local ==="
git init
git add .
git commit -m "feat: script de download da carga efetiva ONS"

echo ""
echo "=== Criando repositório no GitHub ==="
# Troque --public por --private se preferir repositório privado
gh repo create "$REPO_NAME" \
    --description "$DESCRICAO" \
    --public \
    --source=. \
    --remote=origin \
    --push

echo ""
echo "✅ Repositório publicado com sucesso!"
gh repo view --web
