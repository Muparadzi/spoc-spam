#!/bin/bash

# Create directory for uploads
UPLOAD_DIR="/home/dev/upload"
sudo mkdir -p "$UPLOAD_DIR"
sudo chmod 755 "$UPLOAD_DIR"
sudo chown -R $USER:$USER "$UPLOAD_DIR"

# Create index.html
INDEX_HTML_CONTENT='<!DOCTYPE html>
<html>
<head>
    <title>File Upload</title>
</head>
<body>
    <h1>Upload .pst File</h1>
    <form action="/upload" method="post" enctype="multipart/form-data">
        <input type="file" name="pstFile" accept=".pst" required>
        <br><br>
        <input type="submit" value="Upload">
        <button type="button" onclick="alert(`You clicked the additional button`);">Click Me</button>
    </form>
    <br>
    <a href="/control"><button>Go to Control Page</button></a>
</body>
</html>
'
sudo mkdir -p /var/www/html/uploads/template
echo "$INDEX_HTML_CONTENT" | sudo tee /var/www/html/uploads/template/index.html > /dev/null

# Create app.py content
APP_PY_CONTENT='''from flask import Flask, request, render_template
import os
import time

app = Flask(__name__, template_folder="/var/www/html/uploads/template")

UPLOAD_FOLDER = "/home/dev/upload"
ALLOWED_EXTENSIONS = {"pst"}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload_file():
    if "pstFile" not in request.files:
        return "No file part"
    
    file = request.files["pstFile"]
    
    if file.filename == "":
        return "No selected file"
    
    if file and allowed_file(file.filename):
        start_time = time.time()  # Start timing
        filename = file.filename
        file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file.save(file_path)
        end_time = time.time()  # End timing
        
        upload_time = end_time - start_time
        
        return f"File uploaded successfully. Upload time: {upload_time:.2f} seconds"
    else:
        return "Invalid file type"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
'''

# Create app.py
echo "$APP_PY_CONTENT" | sudo tee /var/www/html/uploads/app.py > /dev/null

# Create Nginx configuration file
NGINX_CONF_CONTENT='server {
    listen 80;
    server_name cyber.lan;

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
    location /control {
        proxy_pass http://127.0.0.1:5000/control;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}'
echo "$NGINX_CONF_CONTENT" | sudo tee /etc/nginx/conf.d/upload.conf > /dev/null

# Restart Nginx
sudo service nginx restart

# Create systemd service unit for the Flask app
SERVICE_CONTENT='''[Unit]
Description=Upload App
After=network.target

[Service]
User=$USER
WorkingDirectory=/var/www/html/uploads
ExecStart=/usr/bin/python3 app.py
Restart=always

[Install]
WantedBy=multi-user.target
'''
echo "$SERVICE_CONTENT" | sudo tee /etc/systemd/system/upload_app.service > /dev/null

# Reload systemd and start the service
sudo systemctl daemon-reload
sudo systemctl start upload_app.service
sudo systemctl enable upload_app.service

