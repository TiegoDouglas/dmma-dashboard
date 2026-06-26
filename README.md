# Publicação no GitHub Pages

Este projeto contém a página `dashboard.html`, `styles.css` e `script.js`.

## Pré-requisitos

1. Conta no GitHub.
2. Git instalado no Windows.
3. Acesso ao terminal PowerShell.

## Instalar Git no Windows

- Acesse: https://git-scm.com/download/win
- Baixe e instale usando as opções padrão.
- Após instalar, abra um novo PowerShell e execute:
  ```powershell
  git --version
  ```

## Passo a passo para publicar no GitHub Pages

### 1. Preparar o repositório local

Abra o PowerShell na pasta do projeto e execute:

```powershell
cd "c:\Users\t.douglas.medeiros\Projeto Teste 1"
git init
git add .
git commit -m "Publicação GitHub Pages"
git branch -M main
```

### 2. Criar repositório no GitHub

Opção manual:

1. Acesse https://github.com
2. Clique em **New repository**
3. Nomeie como, por exemplo, `dmma-dashboard`
4. Escolha **Public**
5. Clique em **Create repository**

### 3. Conectar o repositório local ao GitHub

Substitua `SEU_USUARIO` e `REPO` pelos valores reais:

```powershell
git remote add origin https://github.com/SEU_USUARIO/REPO.git
git push -u origin main
```

### 4. Configurar GitHub Pages

1. No GitHub, acesse o repositório.
2. Vá em **Settings** → **Pages**.
3. Em **Build and deployment**, escolha **Deploy from a branch**.
4. Em **Branch**, selecione `main` e `/ (root)`.
5. Clique em **Save**.
6. Aguarde alguns minutos.

A URL de publicação será:

```
https://SEU_USUARIO.github.io/REPO/
```

### 5. Tornar a dashboard a página inicial

Se quiser que o dashboard abra direto, renomeie `dashboard.html` para `index.html` antes de enviar.

Ou crie um `index.html` simples com redirecionamento:

```html
<!doctype html>
<html>
  <head>
    <meta http-equiv="refresh" content="0; url=dashboard.html">
    <title>Redirecionando...</title>
  </head>
  <body>Redirecionando para dashboard...</body>
</html>
```

### 6. Verificar o site

- Abra a URL do GitHub Pages no navegador.
- Se não carregar, verifique o cabeçalho `Pages` em **Settings** e confirme que `index.html` está na raiz.

## Observações

- HTTPS é habilitado automaticamente pelo GitHub Pages.
- Se quiser domínio customizado, adicione o domínio em **Settings > Pages** e configure os registros DNS no seu provedor.
