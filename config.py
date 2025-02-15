import configparser
import os
import sys
import shutil
from pathlib import Path
import logging
from datetime import datetime

def get_mac_dirs():
    """Mac用の各種ディレクトリパスを取得"""
    app_name = '勤怠表自動作成ツール'
    home = Path.home()
    return {
        'config': home / 'Library' / 'Application Support' / app_name,
        'documents': home / 'Documents' / app_name,
        'templates': home / 'Documents' / app_name / 'templates'
    }

def create_default_config(config_path):
    """デフォルトのconfig.iniを作成"""
    config = configparser.ConfigParser(interpolation=None)
    
    config['DEFAULT'] = {
        'employee_name': '',
        'template_path': str(get_mac_dirs()['templates'] / '勤怠表雛形_2025年版.xlsx')
    }
    
    config['PATHS'] = {
        'input_dir': str(get_mac_dirs()['documents'] / 'input'),
        'output_dir': str(get_mac_dirs()['documents'] / 'output')
    }
    
    config['CSV'] = {
        'encoding': 'utf-8',
        'date_format': '%Y-%m-%d'
    }
    
    with open(config_path, 'w', encoding='utf-8') as f:
        config.write(f)

def get_config_path():
    """プラットフォームに応じたconfig.iniのパスを取得"""
    if sys.platform == 'darwin':
        return get_mac_dirs()['config'] / 'config.ini'
    else:
        if getattr(sys, 'frozen', False):
            base_path = Path(sys.executable).parent
        else:
            base_path = Path(__file__).parent
        return base_path / 'config.ini'

def initialize_mac_environment():
    """Mac用の初期環境をセットアップ"""
    if sys.platform != 'darwin':
        return

    dirs = get_mac_dirs()
    
    # 必要なディレクトリを作成
    for dir_path in dirs.values():
        dir_path.mkdir(parents=True, exist_ok=True)
    
    # input/outputディレクトリを作成
    docs_dir = dirs['documents']
    (docs_dir / 'input').mkdir(exist_ok=True)
    (docs_dir / 'output').mkdir(exist_ok=True)
    
    # テンプレートファイルをコピー
    if getattr(sys, 'frozen', False):
        bundle_templates = Path(sys._MEIPASS) / 'templates'
        if bundle_templates.exists():
            for template in bundle_templates.glob('*.xlsx'):
                shutil.copy2(template, dirs['templates'])
    
    # 初期設定ファイルの作成
    config_path = dirs['config'] / 'config.ini'
    if not config_path.exists():
        create_default_config(config_path)

def get_default_employee_name():
    """デフォルトの従業員名を取得"""
    config = configparser.ConfigParser(interpolation=None)
    config.read(get_config_path(), encoding='utf-8')
    return config.get('DEFAULT', 'employee_name', fallback='')

def get_template_path():
    """テンプレートファイルのパスを取得"""
    config = configparser.ConfigParser(interpolation=None)
    config.read(get_config_path(), encoding='utf-8')
    return config.get('DEFAULT', 'template_path', fallback='')

def get_input_dir():
    """入力ディレクトリのパスを取得"""
    config = configparser.ConfigParser(interpolation=None)
    config.read(get_config_path(), encoding='utf-8')
    return config.get('PATHS', 'input_dir', fallback='input')

def get_output_dir() -> str:
    """出力ディレクトリの絶対パスを取得"""
    output_dir = get_app_path() / 'output'
    return str(output_dir.resolve())  # 絶対パスを確実に取得

def get_csv_encoding():
    """CSVファイルのエンコーディングを取得"""
    config = configparser.ConfigParser(interpolation=None)
    config.read(get_config_path(), encoding='utf-8')
    return config.get('CSV', 'encoding', fallback='utf-8')

def get_date_format():
    """日付フォーマットを取得"""
    config = configparser.ConfigParser(interpolation=None)
    config.read(get_config_path(), encoding='utf-8')
    return config.get('CSV', 'date_format', fallback='%Y-%m-%d')

def update_config(employee_name, template_path):
    """設定を更新"""
    config = configparser.ConfigParser(interpolation=None)
    config.read(get_config_path(), encoding='utf-8')
    
    if 'DEFAULT' not in config:
        config['DEFAULT'] = {}
    
    config['DEFAULT']['employee_name'] = employee_name
    config['DEFAULT']['template_path'] = template_path
    
    with open(get_config_path(), 'w', encoding='utf-8') as f:
        config.write(f)

def ensure_directories():
    """必要なディレクトリを作成"""
    dirs = ['input', 'output', 'templates']
    for dir_name in dirs:
        os.makedirs(dir_name, exist_ok=True)

# アプリケーション起動時に実行
initialize_mac_environment()

if sys.platform == 'darwin':  # Mac OS の場合
    from tkinter import ttk
    style = ttk.Style()
    style.theme_use('aqua')  # Mac スタイルを適用

def get_app_path() -> Path:
    """アプリケーションの作業ディレクトリのベースパスを取得"""
    # ユーザーのホームディレクトリ配下に作業ディレクトリを作成
    work_dir = Path.home() / 'Documents' / '勤怠表自動作成ツール'
    work_dir.mkdir(parents=True, exist_ok=True)
    logging.info(f"Application work directory: {work_dir}")
    return work_dir

def get_logs_dir() -> str:
    """ログディレクトリのパスを取得"""
    logs_dir = Path.home() / 'Library' / 'Logs' / '勤怠表自動作成ツール'
    logs_dir.mkdir(parents=True, exist_ok=True)
    return str(logs_dir.resolve())

def setup_logging():
    """ログの設定"""
    logs_dir = Path(get_logs_dir())
    log_file = logs_dir / f"app_{datetime.now().strftime('%Y%m%d')}.log"
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    logging.info("Logging initialized")

def ensure_config():
    """設定ファイルの存在を確認し、必要に応じて作成する"""
    try:
        app_path = get_app_path()
        config_path = app_path / 'config.ini'
        
        if not config_path.exists():
            print(f"Creating new config file at: {config_path}")
            config = configparser.ConfigParser()
            
            config['DEFAULT'] = {
                'employee_name': '小島知将',
                'template_path': str(app_path / 'templates/勤怠表雛形_2025年版.xlsx')
            }
            
            config['PATHS'] = {
                'input_dir': str(app_path / 'input'),
                'output_dir': str(app_path / 'output')
            }
            
            config['CSV'] = {
                'encoding': 'utf-8',
                'date_format': '%Y-%m-%d'
            }
            
            with open(config_path, 'w', encoding='utf-8') as configfile:
                config.write(configfile)
        
        return config_path
    except Exception as e:
        print(f"Error ensuring config: {e}")
        raise

def get_config():
    """設定ファイルを読み込む"""
    try:
        config_path = ensure_config()
        config = configparser.ConfigParser()
        config.read(config_path, encoding='utf-8')
        return config
    except Exception as e:
        print(f"Error reading config: {e}")
        raise

def get_input_dir() -> str:
    """入力ディレクトリのパスを取得"""
    input_dir = get_app_path() / 'input'
    input_dir.mkdir(parents=True, exist_ok=True)
    logging.info(f"Input directory path: {input_dir}")
    return str(input_dir.resolve())

def get_output_dir() -> str:
    """出力ディレクトリのパスを取得"""
    output_dir = get_app_path() / 'output'
    output_dir.mkdir(parents=True, exist_ok=True)
    logging.info(f"Output directory path: {output_dir}")
    return str(output_dir.resolve())

def get_template_path() -> str:
    """テンプレートファイルのパスを取得"""
    template_dir = get_app_path() / 'templates'
    template_dir.mkdir(parents=True, exist_ok=True)
    template_path = template_dir / '勤怠表雛形_2025年版.xlsx'
    
    # テンプレートファイルが存在しない場合、アプリケーションバンドルから取得
    if not template_path.exists() and getattr(sys, 'frozen', False):
        bundle_template = Path(sys._MEIPASS) / 'templates' / '勤怠表雛形_2025年版.xlsx'
        if bundle_template.exists():
            shutil.copy2(bundle_template, template_path)
    
    logging.info(f"Template path: {template_path}")
    return str(template_path.resolve())

def get_default_employee_name():
    """デフォルトの従業員名を取得"""
    config = get_config()
    return config['DEFAULT']['employee_name']

def update_config(employee_name, template_path):
    """設定を更新する"""
    try:
        config = get_config()
        config['DEFAULT']['employee_name'] = employee_name
        config['DEFAULT']['template_path'] = template_path
        
        config_path = get_app_path() / 'config.ini'
        with open(config_path, 'w', encoding='utf-8') as configfile:
            config.write(configfile)
    except Exception as e:
        print(f"Error updating config: {e}")
        raise