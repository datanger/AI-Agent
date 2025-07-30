#!/usr/bin/env python3
"""
Gemini Proxy Server for Continue
提供 gemini-cli 的 HTTP API 接口
"""

import os
import json
import logging
from flask import Flask, request, jsonify, Response, stream_template
from flask_cors import CORS
from gemini_process import GeminiProcess

# 配置日志 - 写入到 console.log 文件
log_dir = os.path.dirname(os.path.abspath(__file__))
log_file = os.path.join(log_dir, 'console.log')

# 创建日志格式
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# 配置文件处理器
file_handler = logging.FileHandler(log_file, encoding='utf-8')
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(formatter)

# 配置控制台处理器
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_handler.setFormatter(formatter)

# 配置根日志器
logging.basicConfig(
    level=logging.INFO,
    handlers=[file_handler, console_handler],
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

# 默认配置
DEFAULT_PROVIDER = os.environ.get('GEMINI_PROVIDER', 'deepseek')
DEFAULT_MODEL = os.environ.get('GEMINI_MODEL', 'deepseek-chat')
DEFAULT_API_KEY = os.environ.get('GEMINI_API_KEY', None)

# 环境变量默认值
DEFAULT_OLLAMA_BASE_URL = os.environ.get('OLLAMA_BASE_URL', 'http://127.0.0.1:11434')
DEFAULT_LOCAL_BASE_URL = os.environ.get('LOCAL_BASE_URL', 'http://127.0.0.1:8080')
DEFAULT_DEEPSEEK_API_BASE = os.environ.get('DEEPSEEK_API_BASE', 'https://api.deepseek.com')
DEFAULT_OPENAI_API_BASE = os.environ.get('OPENAI_API_BASE', 'https://api.openai.com')

# 初始化 Gemini 进程
gemini = GeminiProcess(
    provider=DEFAULT_PROVIDER,
    model=DEFAULT_MODEL,
    api_key=DEFAULT_API_KEY
)

@app.route('/ask', methods=['POST'])
def ask():
    """处理单次问答请求"""
    try:
        data = request.get_json()
        prompt = data.get('prompt', '')
        
        if not prompt:
            return jsonify({'error': 'No prompt provided'}), 400
        
        logger.info(f"Received prompt: {prompt[:50]}...")
        
        # 发送 prompt 并获取响应
        response = gemini.send_prompt(prompt)
        
        return jsonify({
            'response': response,
            'provider': gemini.provider,
            'model': gemini.model
        })
        
    except Exception as e:
        logger.error(f"Error processing request: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/ask/stream', methods=['POST'])
def ask_stream():
    """处理流式问答请求"""
    try:
        data = request.get_json()
        prompt = data.get('prompt', '')
        
        if not prompt:
            return jsonify({'error': 'No prompt provided'}), 400
        
        logger.info(f"Received streaming prompt: {prompt[:50]}...")
        
        def generate():
            try:
                # 使用流式发送
                for chunk in gemini.send_prompt_stream(prompt):
                    yield f"data: {json.dumps({'chunk': chunk})}\n\n"
                
                # 发送完成信号
                yield f"data: {json.dumps({'done': True})}\n\n"
                
            except Exception as e:
                logger.error(f"Error in streaming: {e}")
                yield f"data: {json.dumps({'error': str(e)})}\n\n"
        
        return Response(generate(), mimetype='text/plain')
        
    except Exception as e:
        logger.error(f"Error processing streaming request: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/config', methods=['GET'])
def get_config():
    """获取当前配置"""
    return jsonify({
        'provider': gemini.provider,
        'model': gemini.model,
        'environment_variables': {
            'GEMINI_PROVIDER': os.environ.get('GEMINI_PROVIDER', DEFAULT_PROVIDER),
            'GEMINI_MODEL': os.environ.get('GEMINI_MODEL', DEFAULT_MODEL),
            'OLLAMA_BASE_URL': os.environ.get('OLLAMA_BASE_URL', DEFAULT_OLLAMA_BASE_URL),
            'LOCAL_BASE_URL': os.environ.get('LOCAL_BASE_URL', DEFAULT_LOCAL_BASE_URL),
            'DEEPSEEK_API_BASE': os.environ.get('DEEPSEEK_API_BASE', DEFAULT_DEEPSEEK_API_BASE),
            'OPENAI_API_BASE': os.environ.get('OPENAI_API_BASE', DEFAULT_OPENAI_API_BASE)
        }
    })

@app.route('/config', methods=['POST'])
def update_config():
    """更新配置"""
    try:
        data = request.get_json()
        provider = data.get('provider')
        model = data.get('model')
        api_key = data.get('api_key')
        
        # 更新环境变量
        if provider:
            os.environ['GEMINI_PROVIDER'] = provider
        if model:
            os.environ['GEMINI_MODEL'] = model
        if api_key:
            os.environ['GEMINI_API_KEY'] = api_key
        
        # 更新 Gemini 进程配置
        gemini.update_config(provider, model, api_key)
        
        return jsonify({
            'message': 'Configuration updated successfully',
            'provider': gemini.provider,
            'model': gemini.model
        })
        
    except Exception as e:
        logger.error(f"Error updating config: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/health', methods=['GET'])
def health_check():
    """健康检查端点"""
    is_alive = gemini.is_process_alive()
    status = "healthy" if is_alive else "unhealthy"
    logger.info(f"Health check: {status}, process_alive: {is_alive}")
    return jsonify({
        'status': status, 
        'process_alive': is_alive,
        'provider': gemini.provider,
        'model': gemini.model,
        'environment_variables': {
            'OLLAMA_BASE_URL': os.environ.get('OLLAMA_BASE_URL', DEFAULT_OLLAMA_BASE_URL),
            'LOCAL_BASE_URL': os.environ.get('LOCAL_BASE_URL', DEFAULT_LOCAL_BASE_URL),
            'DEEPSEEK_API_BASE': os.environ.get('DEEPSEEK_API_BASE', DEFAULT_DEEPSEEK_API_BASE),
            'OPENAI_API_BASE': os.environ.get('OPENAI_API_BASE', DEFAULT_OPENAI_API_BASE)
        }
    }), 200 if is_alive else 503

@app.route('/', methods=['GET'])
def root():
    """根端点"""
    return jsonify({
        'service': 'gemini-proxy',
        'version': '1.4.0',
        'status': 'running',
        'endpoints': {
            'health': '/health',
            'ask': '/ask',
            'ask_stream': '/ask/stream',
            'config': '/config'
        }
    })

if __name__ == '__main__':
    logger.info(f"Starting gemini-proxy server v1.4.0 on port 5001")
    logger.info(f"Default config: provider={DEFAULT_PROVIDER}, model={DEFAULT_MODEL}")
    logger.info("Features: multi-turn conversation, agent functionality, tool calls, streaming, configurable provider")
    logger.info("Supported providers: gemini, openai, ollama, local, deepseek")
    logger.info(f"Log file: {log_file}")

    app.run(host='127.0.0.1', port=5001, debug=False)
