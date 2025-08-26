from flask import Blueprint, request, jsonify, send_file
from werkzeug.utils import secure_filename
import os
import pandas as pd
import openpyxl
from datetime import datetime, timedelta
import zipfile
import tempfile
import shutil
from docx import Document
from docx.shared import Inches
import re
from src.utils.letter_generator_utils import (
    process_reporting_data, 
    generate_letter_document, 
    generate_appendix_document,
    format_amount_in_words
)

letter_bp = Blueprint('letter', __name__)

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'uploads')
GENERATED_FOLDER = os.path.join(os.path.dirname(__file__), 'generated_letters')

# Создаем папки если их нет
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls'}

def clean_contractor_name(name):
    """Удаляет первые 10 цифр из названия контрагента"""
    if isinstance(name, str):
        # Удаляем первые 10 цифр если они есть в начале
        cleaned = re.sub(r'^\d{10}', '', name).strip()
        return cleaned
    return name

def calculate_penalty(amount, days_overdue):
    """Расчет пени с учетом сложного процента"""
    if days_overdue <= 0:
        return 0
    
    penalty = 0
    remaining_days = days_overdue
    current_amount = amount
    
    # Первые 14 дней - 0.1% в день
    if remaining_days > 0:
        days_first_period = min(remaining_days, 14)
        for day in range(days_first_period):
            daily_penalty = current_amount * 0.001  # 0.1%
            penalty += daily_penalty
            current_amount += daily_penalty
        remaining_days -= days_first_period
    
    # После 14 дней - 0.5% в день
    if remaining_days > 0:
        for day in range(remaining_days):
            daily_penalty = current_amount * 0.005  # 0.5%
            penalty += daily_penalty
            current_amount += daily_penalty
    
    return penalty

def number_to_words_russian(number):
    """Преобразование числа в слова на русском языке"""
    # Упрощенная версия для демонстрации
    # В реальном проекте следует использовать библиотеку num2words
    
    units = ['', 'один', 'два', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять']
    teens = ['десять', 'одиннадцать', 'двенадцать', 'тринадцать', 'четырнадцать', 
             'пятнадцать', 'шестнадцать', 'семнадцать', 'восемнадцать', 'девятнадцать']
    tens = ['', '', 'двадцать', 'тридцать', 'сорок', 'пятьдесят', 'шестьдесят', 'семьдесят', 'восемьдесят', 'девяносто']
    hundreds = ['', 'сто', 'двести', 'триста', 'четыреста', 'пятьсот', 'шестьсот', 'семьсот', 'восемьсот', 'девятьсот']
    
    if number == 0:
        return 'ноль'
    
    # Для простоты возвращаем строковое представление
    # В реальном проекте нужна полная реализация
    return str(number)

def format_amount_in_words(amount):
    """Форматирование суммы прописью"""
    rubles = int(amount)
    kopecks = int((amount - rubles) * 100)
    
    rubles_words = number_to_words_russian(rubles)
    
    return f"{rubles_words} рублей {kopecks:02d} копеек"

@letter_bp.route('/upload', methods=['POST'])
def upload_files():
    """Загрузка Excel файлов"""
    try:
        if 'reporting_file' not in request.files or 'sed_file' not in request.files:
            return jsonify({'error': 'Необходимо загрузить оба файла: отчетность и СЭД'}), 400
        
        reporting_file = request.files['reporting_file']
        sed_file = request.files['sed_file']
        
        if reporting_file.filename == '' or sed_file.filename == '':
            return jsonify({'error': 'Файлы не выбраны'}), 400
        
        if not (allowed_file(reporting_file.filename) and allowed_file(sed_file.filename)):
            return jsonify({'error': 'Разрешены только Excel файлы (.xlsx, .xls)'}), 400
        
        # Сохраняем файлы
        reporting_filename = secure_filename('reporting.xlsx')
        sed_filename = secure_filename('sed.xlsx')
        
        reporting_path = os.path.join(UPLOAD_FOLDER, reporting_filename)
        sed_path = os.path.join(UPLOAD_FOLDER, sed_filename)
        
        reporting_file.save(reporting_path)
        sed_file.save(sed_path)
        
        return jsonify({
            'message': 'Файлы успешно загружены',
            'reporting_file': reporting_filename,
            'sed_file': sed_filename
        })
        
    except Exception as e:
        return jsonify({'error': f'Ошибка при загрузке файлов: {str(e)}'}), 500

@letter_bp.route('/process', methods=['POST'])
def process_files():
    """Обработка файлов и генерация писем"""
    try:
        # Проверяем наличие файлов
        reporting_path = os.path.join(UPLOAD_FOLDER, 'reporting.xlsx')
        sed_path = os.path.join(UPLOAD_FOLDER, 'sed.xlsx')
        
        if not (os.path.exists(reporting_path) and os.path.exists(sed_path)):
            return jsonify({'error': 'Файлы не найдены. Загрузите файлы сначала.'}), 400
        
        # Обрабатываем данные
        letters_data = process_reporting_data(reporting_path, sed_path)
        
        # Очищаем папку с сгенерированными письмами
        if os.path.exists(GENERATED_FOLDER):
            shutil.rmtree(GENERATED_FOLDER)
        os.makedirs(GENERATED_FOLDER, exist_ok=True)
        
        # Генерируем письма и приложения
        generated_files = []
        for i, letter_data in enumerate(letters_data):
            try:
                # Генерируем основное письмо
                letter_filename = f"letter_{i+1}_{letter_data['contractor_short_name']}_{letter_data['order_number']}.docx"
                letter_path = os.path.join(GENERATED_FOLDER, letter_filename)
                generate_letter_document(letter_data, letter_path)
                
                # Генерируем приложение
                appendix_filename = f"appendix_{i+1}_{letter_data['contractor_short_name']}_{letter_data['order_number']}.docx"
                appendix_path = os.path.join(GENERATED_FOLDER, appendix_filename)
                generate_appendix_document(letter_data, appendix_path)
                
                generated_files.extend([letter_filename, appendix_filename])
                
            except Exception as e:
                print(f"Ошибка генерации письма {i+1}: {str(e)}")
                continue
        
        return jsonify({
            'message': f'Обработано и сгенерировано {len(letters_data)} писем',
            'letters_count': len(letters_data),
            'files_generated': generated_files,
            'letters_data': letters_data
        })
        
    except Exception as e:
        return jsonify({'error': f'Ошибка при обработке файлов: {str(e)}'}), 500

@letter_bp.route('/download/<filename>', methods=['GET'])
def download_single_file(filename):
    """Скачивание одного файла"""
    try:
        file_path = os.path.join(GENERATED_FOLDER, filename)
        if not os.path.exists(file_path):
            return jsonify({'error': 'Файл не найден'}), 404
        
        return send_file(file_path, as_attachment=True, download_name=filename)
        
    except Exception as e:
        return jsonify({'error': f'Ошибка при скачивании файла: {str(e)}'}), 500

@letter_bp.route('/download_all', methods=['GET'])
def download_all_letters():
    """Скачивание всех писем в ZIP архиве"""
    try:
        if not os.path.exists(GENERATED_FOLDER) or not os.listdir(GENERATED_FOLDER):
            return jsonify({'error': 'Нет сгенерированных файлов для скачивания'}), 404
        
        # Создаем временную папку для ZIP архива
        temp_dir = tempfile.mkdtemp()
        zip_path = os.path.join(temp_dir, 'all_letters.zip')
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Добавляем все сгенерированные файлы в архив
            for filename in os.listdir(GENERATED_FOLDER):
                if filename.endswith('.docx'):
                    file_path = os.path.join(GENERATED_FOLDER, filename)
                    zipf.write(file_path, filename)
        
        return send_file(zip_path, as_attachment=True, download_name='all_letters.zip')
        
    except Exception as e:
        return jsonify({'error': f'Ошибка при создании архива: {str(e)}'}), 500

@letter_bp.route('/status', methods=['GET'])
def get_status():
    """Получение статуса системы"""
    try:
        reporting_exists = os.path.exists(os.path.join(UPLOAD_FOLDER, 'reporting.xlsx'))
        sed_exists = os.path.exists(os.path.join(UPLOAD_FOLDER, 'sed.xlsx'))
        
        generated_count = len([f for f in os.listdir(GENERATED_FOLDER) if f.endswith('.docx')])
        
        return jsonify({
            'reporting_file_uploaded': reporting_exists,
            'sed_file_uploaded': sed_exists,
            'generated_letters_count': generated_count
        })
        
    except Exception as e:
        return jsonify({'error': f'Ошибка при получении статуса: {str(e)}'}), 500

