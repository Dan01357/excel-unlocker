from flask import Flask, request, jsonify
from flask_cors import CORS
import msoffcrypto
import io
import base64

app = Flask(__name__)
CORS(app) # Allows your Google Web App to talk to this tool

@app.route('/decrypt', methods=['POST'])
def decrypt_file():
    data = request.json
    password = data.get('password')
    file_b64 = data.get('file_b64')

    if not password or not file_b64:
        return jsonify({"error": "Missing data", "success": False}), 400

    try:
        file_data = base64.b64decode(file_b64)
        encrypted_file = io.BytesIO(file_data)
        decrypted_file = io.BytesIO()

        # Unlock the Excel file
        office_file = msoffcrypto.OfficeFile(encrypted_file)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted_file)

        # Return the unlocked file back to the website
        decrypted_b64 = base64.b64encode(decrypted_file.getvalue()).decode('utf-8')
        return jsonify({"success": True, "decrypted_b64": decrypted_b64})

    except Exception as e:
        return jsonify({"error": "Incorrect password or invalid file", "success": False}), 400

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
