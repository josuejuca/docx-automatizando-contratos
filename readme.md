# API - Gerador de Contratos

### Ubuntu - Debian  

### ✅ 1. Dependências Python (instalar via `pip`)

No seu ambiente virtual (ou global):

```bash
pip install -r req_linux.txt
```

---

### ✅ 2. Dependência do sistema: **LibreOffice**

Essa parte é obrigatória para que o `.docx` seja convertido em `.pdf` via linha de comando.

#### Instalar no Ubuntu/Debian:

```bash
sudo apt update
sudo apt install libreoffice
sudo apt install language-pack-pt
```

---

### ✅ 3. Verifique se está instalado corretamente:

Rode no terminal:

```bash
libreoffice --headless --convert-to pdf --version
```

Se aparecer a versão do LibreOffice, está tudo certo ✅

---

### ✅ 4. Como rodar sua API:

```bash
nohup uvicorn app_linux:app --host 0.0.0.0 --port 8000 > uvicorn.log 2>&1 &
```

(ou substitua `app_linux` pelo nome real, sem `.py`)

---

### ✅ 5. Teste a conversão manual (opcional)

Para garantir que o LibreOffice funciona no servidor, teste diretamente:

```bash
libreoffice --headless --convert-to pdf seu-modelo.docx --outdir .
```

Se gerar o `.pdf`, está 100% funcional.

---

### ✅ 6. No painel DNS da `imogo.com.br`

Adicione um **registro A**:

- **Tipo**: A  
- **Nome**: `docx`  
- **Valor**: IP público do seu servidor (ex: `189.2.33.10`)  
- **TTL**: 5 min ou automático

---

### ✅ 7. Configurar NGINX (proxy reverso)

Crie o arquivo `/etc/nginx/sites-available/docx.imogo.com.br`:

```nginx
server {
    listen 80;
    server_name docx.imogo.com.br;

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```

Ative o site:

```bash
sudo ln -s /etc/nginx/sites-available/docx.imogo.com.br /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl reload nginx
```

---

### 🔒 (Opcional) 8. Habilitar HTTPS com Let's Encrypt

Se quiser HTTPS gratuito:

```bash
sudo apt install certbot python3-certbot-nginx
sudo certbot --nginx -d docx.imogo.com.br
```

---

### ✅ Resultado:

Agora, ao acessar docx.imogo.com.br, você estará acessando sua API FastAPI rodando na porta `8000`.

`Esse é um exemplo usando o dominio da imogo.com.br, se você quiser usar (e deve) o seu dominio faça todos os passos trocando apenas o docx.imogo.com.br por <seu_dominio>`

### Utils

```bash
fc-list | grep -i nunito
```
Use esse comando para verificar se a fonte nunito foi instalada 

```bash
$ fc-list | grep -i nunito
/root/.local/share/fonts/imogo-nunito/Nunito-Bold.ttf: Nunito:style=Bold
/root/.local/share/fonts/imogo-nunito/Nunito-Regular.ttf: Nunito:style=Regular
```


