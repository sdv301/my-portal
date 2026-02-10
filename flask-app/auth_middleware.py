from functools import wraps
from flask import request, jsonify
import datetime
import jwt
import os

JWT_SECRET = os.getenv('JWT_SECRET', 'your-secret-key')

def token_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        token = None
        
        if 'Authorization' in request.headers:
            token = request.headers['Authorization'].split(" ")[1]
        
        if not token:
            return jsonify({'message': 'Токен отсутствует!'}), 401
            
        try:
            data = jwt.decode(token, JWT_SECRET, algorithms=["HS256"])
            current_user = data['user']
        except:
            return jsonify({'message': 'Токен невалидный!'}), 401
            
        return f(current_user, *args, **kwargs)
    
    return decorated

# Эндпоинт для получения токена
@app.route('/api/login', methods=['POST'])
def login():
    auth = request.json
    if not auth or not auth.get('username') or not auth.get('password'):
        return jsonify({'message': 'Некорректные данные'}), 400
    
    # Проверка учетных данных (в реальности - проверка в БД)
    if auth['username'] == 'admin' and auth['password'] == 'admin123':
        token = jwt.encode({
            'user': auth['username'],
            'exp': datetime.datetime.utcnow() + datetime.timedelta(hours=24)
        }, JWT_SECRET)
        
        return jsonify({'token': token}), 200
    
    return jsonify({'message': 'Неверные учетные данные'}), 401