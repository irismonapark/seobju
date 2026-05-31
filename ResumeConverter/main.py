"""Vercel Flask entrypoint (pyproject: tool.vercel.entrypoint = main:app)."""
import os
import traceback


def _fallback_app(exc: BaseException):
    from flask import Flask, jsonify

    application = Flask(__name__)
    detail = traceback.format_exc()

    @application.get('/health')
    def health_error():
        return jsonify({
            'status': 'error',
            'service': 'resume-converter',
            'error': str(exc),
            'detail': detail[-3000:],
        }), 500

    @application.get('/')
    def index_error():
        return health_error()

    return application


try:
    from app import app
except Exception as boot_exc:
    app = _fallback_app(boot_exc)
    os.environ['BOOT_ERROR'] = '1'
