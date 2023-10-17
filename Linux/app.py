from flask import Flask, request, render_template
import os
import time

app = Flask(__name__, template_folder="/var/www/html/uploads/template")

UPLOAD_FOLDER = "/home/dev/upload"
ALLOWED_EXTENSIONS = {"pst"}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

def log_uploaded_filename(filename):
    with open("/home/dev/upload/uploaded_log.txt", "w") as log_file:
        log_file.write(filename + "\n")


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/")
def index():
    return render_template("index.html")

@app.route('/control')
def control():
    return render_template('control.html')

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
        log_uploaded_filename(filename)
        end_time = time.time()  # End timing
        
        upload_time = end_time - start_time
        
        return f"File uploaded successfully. Upload time: {upload_time:.2f} seconds"
    else:
        return "Invalid file type"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
