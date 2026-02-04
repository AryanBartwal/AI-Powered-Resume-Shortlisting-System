from flask import Flask
from flask_cors import CORS
import os
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
CORS(app)

# Import routes after app creation
from routes.parse_routes import parse_routes
from routes.ranking_routes import ranking_routes

app.config['MAX_CONTENT_LENGTH'] = int(os.getenv('MAX_FILE_SIZE', 10485760))
app.config['UPLOAD_FOLDER'] = os.getenv('TEMP_UPLOAD_DIR', './temp_uploads')

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

app.register_blueprint(parse_routes, url_prefix='/api/parse')
app.register_blueprint(ranking_routes, url_prefix='/api/ranking')

@app.route('/health', methods=['GET'])
def health_check():
    return {'status': 'OK', 'message': 'ML Pipeline Service is running'}, 200

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5001))
    host = os.getenv('HOST', '0.0.0.0')
    print(f'✅ ML Pipeline Service starting on http://{host}:{port}')
    app.run(host=host, port=port, debug=True)
